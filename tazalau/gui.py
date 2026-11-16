"""Окно программы.

Вся работа идёт в отдельном потоке и общается с окном только через очередь сообщений —
окно не «зависает», а кнопка «Стоп» срабатывает сразу после текущего ИИН.
"""
from __future__ import annotations

import queue
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl.utils import get_column_letter

from . import APP_NAME, __version__
from .config import (LINK_HEADER, RESULT_HEADER, SCREENSHOT_DIR_NAME, Settings, load_prefs,
                     update_prefs)
from .excel_io import ExcelError
from .registries import DEFAULT_KEYS, GROUP_BANKRUPTS, GROUP_DEBTORS, REGISTRIES
from .runner import (Analysis, RunOptions, Runner, Summary, analyze, estimate_minutes,
                     nothing_to_do, plan_from_tasks)
from .stamp import open_in_system

POLL_MS = 100
INTRO = ("Программа ищет каждый ИИН из файла Excel в открытых реестрах сайта tazalau.qoldau.kz, "
         "сохраняет снимок страницы с результатом и записывает ссылку на снимок в столбец "
         f"«{LINK_HEADER}». Вход на сайт и ЭЦП не требуются.")


class App:
    def __init__(self, root: tk.Tk, settings: Settings, browser_factory=None):
        self.root = root
        self.settings = settings
        self._browser_factory = browser_factory      # подмена браузера в тестах
        self.queue: "queue.Queue[tuple]" = queue.Queue()
        self.stop_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.analysis: Analysis | None = None
        self._analysis_token = 0
        self._out_dir_auto = True
        self._copy_auto = False
        self._closing = False
        self._started_at = 0.0

        self.var_xlsx = tk.StringVar()
        self.var_out = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_info = tk.StringVar(value="Выберите файл Excel.")
        self.var_status = tk.StringVar(value="Готово к работе.")
        self.var_recheck = tk.BooleanVar(value=False)
        self.var_result = tk.BooleanVar(value=True)
        self.var_visible = tk.BooleanVar(value=False)
        self.var_copy = tk.BooleanVar(value=False)
        self.var_registry = {r.key: tk.BooleanVar(value=r.key in DEFAULT_KEYS) for r in REGISTRIES}

        self._build()
        self._load_prefs()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(POLL_MS, self._poll)

    # ------------------------------------------------------------------ окно
    def _build(self) -> None:
        root = self.root
        root.title(f"{APP_NAME} {__version__} — проверка ИИН по реестрам Tazalau")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main = ttk.Frame(root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(7, weight=1)

        self.lbl_intro = ttk.Label(main, text=INTRO, justify="left", wraplength=800)
        self.lbl_intro.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        # 1. файл
        box = ttk.LabelFrame(main, text=" 1. Файл Excel со столбцом «ИИН/БИН» ", padding=8)
        box.grid(row=1, column=0, sticky="ew", pady=4)
        box.columnconfigure(1, weight=1)
        ttk.Entry(box, textvariable=self.var_xlsx, state="readonly").grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        self.btn_xlsx = ttk.Button(box, text="Выбрать файл…", command=self._choose_xlsx)
        self.btn_xlsx.grid(row=0, column=2)
        ttk.Label(box, text="Лист:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.cmb_sheet = ttk.Combobox(box, textvariable=self.var_sheet, state="disabled", width=40)
        self.cmb_sheet.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        self.cmb_sheet.bind("<<ComboboxSelected>>", lambda _event: self._start_analysis())
        self.lbl_info = ttk.Label(box, textvariable=self.var_info, justify="left", wraplength=780)
        self.lbl_info.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        # 2. папка
        box = ttk.LabelFrame(main, text=" 2. Папка для снимков ", padding=8)
        box.grid(row=2, column=0, sticky="ew", pady=4)
        box.columnconfigure(0, weight=1)
        ttk.Entry(box, textvariable=self.var_out, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=(0, 8))
        self.btn_out = ttk.Button(box, text="Выбрать папку…", command=self._choose_out)
        self.btn_out.grid(row=0, column=1)

        # 3. реестры
        box = ttk.LabelFrame(main, text=" 3. В каких реестрах искать ", padding=8)
        box.grid(row=3, column=0, sticky="ew", pady=4)
        box.columnconfigure(0, weight=1)
        box.columnconfigure(1, weight=1)
        self.registry_checks: list[ttk.Checkbutton] = []
        for column, group in enumerate((GROUP_BANKRUPTS, GROUP_DEBTORS)):
            frame = ttk.LabelFrame(box, text=f" {group} ", padding=6)
            frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0))
            for registry in (r for r in REGISTRIES if r.group == group):
                check = ttk.Checkbutton(frame, text=registry.title, command=self._refresh_info,
                                        variable=self.var_registry[registry.key])
                check.pack(anchor="w")
                self.registry_checks.append(check)

        # параметры
        box = ttk.LabelFrame(main, text=" Параметры ", padding=8)
        box.grid(row=4, column=0, sticky="ew", pady=4)
        box.columnconfigure(0, weight=1)
        box.columnconfigure(1, weight=1)
        self.option_checks = [
            ttk.Checkbutton(box, variable=self.var_result, command=self._refresh_info,
                            text=f"Писать итог (найден / не найден) в столбец «{RESULT_HEADER}»"),
            ttk.Checkbutton(box, variable=self.var_recheck, command=self._refresh_info,
                            text="Проверить заново и те строки, где ссылка уже есть"),
            ttk.Checkbutton(box, variable=self.var_visible, command=self._refresh_info,
                            text="Только видимые строки (учитывать фильтр Excel)"),
            ttk.Checkbutton(box, variable=self.var_copy, command=self._on_copy_toggled,
                            text="Не менять исходный файл — записать результат в копию"),
        ]
        for index, check in enumerate(self.option_checks):
            check.grid(row=index // 2, column=index % 2, sticky="w", pady=1)

        # кнопки
        bar = ttk.Frame(main)
        bar.grid(row=5, column=0, sticky="ew", pady=(8, 4))
        bar.columnconfigure(2, weight=1)
        self.btn_start = ttk.Button(bar, text="▶  Старт", command=self._start, width=16)
        self.btn_start.grid(row=0, column=0, padx=(0, 6))
        self.btn_stop = ttk.Button(bar, text="■  Стоп", command=self._stop, width=12, state="disabled")
        self.btn_stop.grid(row=0, column=1)
        self.btn_open_out = ttk.Button(bar, text="Открыть папку со снимками", command=self._open_out)
        self.btn_open_out.grid(row=0, column=3, padx=(0, 6))
        self.btn_open_xlsx = ttk.Button(bar, text="Открыть файл Excel", command=self._open_xlsx)
        self.btn_open_xlsx.grid(row=0, column=4)

        # ход работы
        self.progress = ttk.Progressbar(main, mode="determinate", maximum=1, value=0)
        self.progress.grid(row=6, column=0, sticky="ew", pady=(4, 2))
        log_frame = ttk.Frame(main)
        log_frame.grid(row=7, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        ttk.Label(log_frame, textvariable=self.var_status).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.txt_log = tk.Text(log_frame, height=6, wrap="word", state="disabled",
                               borderwidth=1, relief="solid")
        self.txt_log.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=scroll.set)
        self.txt_log.tag_configure("error", foreground="#b00020")
        self.txt_log.tag_configure("warn", foreground="#9a6700")
        self.txt_log.tag_configure("found", foreground="#0a7d32")
        self.txt_log.tag_configure("time", foreground="#808080")

        main.bind("<Configure>", self._on_resize)
        self._fit_to_screen()

    def _fit_to_screen(self, screen_w: int | None = None, screen_h: int | None = None) -> None:
        """Начальный размер окна. Всё, кроме журнала, занимает ровно столько, сколько ему нужно;
        журнал получает остальное и первым уступает место, если сведений о файле стало больше.
        Окно не выше экрана: на ноутбуке 1366×768 оно тоже должно помещаться целиком."""
        root = self.root
        screen_w = screen_w or root.winfo_screenwidth()
        screen_h = screen_h or root.winfo_screenheight()
        room_h = max(480, screen_h - 110)                 # заголовок окна и панель задач
        room_w = max(640, screen_w - 40)
        current = self.var_info.get()
        self.var_info.set("Сведения о файле\nСведения о файле")      # обычный объём — две строки
        root.update_idletasks()
        if root.winfo_reqheight() + 60 > room_h:          # маленький экран: вводный текст уступает
            self.lbl_intro.grid_remove()                  # место журналу
            root.update_idletasks()
        need_w, need_h = root.winfo_reqwidth(), root.winfo_reqheight()
        self.var_info.set(current)
        width = min(max(need_w, 860), room_w)
        height = min(need_h + 120, room_h)
        root.minsize(min(need_w, room_w), min(need_h, room_h))
        root.geometry(f"{width}x{height}")

    def _on_resize(self, event) -> None:
        width = max(300, event.width - 40)
        self.lbl_intro.configure(wraplength=width)
        self.lbl_info.configure(wraplength=width - 20)

    # ------------------------------------------------------------------ диалоги (подменяются в тестах)
    def ask(self, title: str, text: str) -> bool:
        return messagebox.askyesno(title, text, parent=self.root)

    def info(self, title: str, text: str) -> None:
        messagebox.showinfo(title, text, parent=self.root)

    def warn(self, title: str, text: str) -> None:
        messagebox.showwarning(title, text, parent=self.root)

    def error(self, title: str, text: str) -> None:
        messagebox.showerror(title, text, parent=self.root)

    # ------------------------------------------------------------------ настройки между запусками
    def _load_prefs(self) -> None:
        prefs = load_prefs()
        keys = prefs.get("registries")
        if isinstance(keys, list) and any(k in self.var_registry for k in keys):
            for key, var in self.var_registry.items():
                var.set(key in keys)
        if isinstance(prefs.get("write_result"), bool):
            self.var_result.set(prefs["write_result"])
        out_dir = prefs.get("out_dir")
        if isinstance(out_dir, str) and out_dir and not prefs.get("out_dir_auto", True):
            self.var_out.set(out_dir)
            self._out_dir_auto = False
        xlsx = prefs.get("xlsx")
        if isinstance(xlsx, str) and xlsx and Path(xlsx).is_file():
            self._set_xlsx(Path(xlsx))

    def _save_prefs(self) -> None:
        update_prefs({
            "xlsx": self.var_xlsx.get(),
            "out_dir": self.var_out.get(),
            "out_dir_auto": self._out_dir_auto,
            "registries": self._selected_registries(),
            "write_result": self.var_result.get(),
        })

    # ------------------------------------------------------------------ выбор файла и папки
    def _choose_xlsx(self) -> None:
        current = self.var_xlsx.get()
        path = filedialog.askopenfilename(
            parent=self.root, title="Файл Excel со столбцом «ИИН/БИН»",
            initialdir=str(Path(current).parent) if current else None,
            filetypes=[("Книга Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")])
        if path:
            self._set_xlsx(Path(path))

    def _set_xlsx(self, path: Path) -> None:
        self.var_xlsx.set(str(path))
        if self._out_dir_auto:
            self.var_out.set(str(path.parent / SCREENSHOT_DIR_NAME))
        self.var_sheet.set("")
        self._copy_auto = False
        self.var_copy.set(False)
        self._start_analysis()

    def _choose_out(self) -> None:
        current = self.var_out.get()
        path = filedialog.askdirectory(parent=self.root, title="Папка для снимков",
                                       initialdir=current if current and Path(current).exists() else None)
        if path:
            self.var_out.set(str(Path(path)))
            self._out_dir_auto = False

    def _on_copy_toggled(self) -> None:
        self._copy_auto = False            # пользователь решил сам — больше не переключаем
        self._refresh_info()

    # ------------------------------------------------------------------ чтение файла
    def _start_analysis(self) -> None:
        path = self.var_xlsx.get()
        if not path:
            return
        self._analysis_token += 1
        token = self._analysis_token
        sheet = self.var_sheet.get() or None
        self.analysis = None
        self.var_info.set("Читаю файл…")

        def work() -> None:
            try:
                self.queue.put(("analysis", token, analyze(Path(path), sheet, self.settings)))
            except ExcelError as exc:
                self.queue.put(("analysis_error", token, str(exc)))
            except Exception as exc:
                self.queue.put(("analysis_error", token,
                                f"Не удалось прочитать файл: {type(exc).__name__}: {exc}"))

        threading.Thread(target=work, daemon=True).start()

    def _on_analysis(self, result: Analysis) -> None:
        self.analysis = result
        self.cmb_sheet.configure(values=result.sheets,
                                 state="readonly" if self.worker is None else "disabled")
        self.var_sheet.set(result.layout.sheet)
        if result.risks and not self.var_copy.get() and not self._is_working_copy():
            self.var_copy.set(True)
            self._copy_auto = True
        self._refresh_info()

    def _is_working_copy(self) -> bool:
        """Файл уже был записан программой как копия: всё, что могло потеряться, уже потеряно,
        и предлагать «копию копии» незачем."""
        return "_tazalau_" in Path(self.var_xlsx.get()).stem

    def _options(self) -> RunOptions:
        return RunOptions(
            xlsx=Path(self.var_xlsx.get()), out_dir=Path(self.var_out.get()),
            sheet=self.var_sheet.get() or None, registries=tuple(self._selected_registries()),
            recheck=self.var_recheck.get(), write_result=self.var_result.get(),
            only_visible=self.var_visible.get(), save_copy=self.var_copy.get())

    def _selected_registries(self) -> list[str]:
        return [r.key for r in REGISTRIES if self.var_registry[r.key].get()]

    def _refresh_info(self) -> None:
        result = self.analysis
        if result is None:
            return
        plan = plan_from_tasks(result.layout, result.tasks, result.risks, self._options())
        layout = result.layout
        lines = [f"Столбец «ИИН/БИН»: {get_column_letter(layout.iin_col)} (заголовки в строке "
                 f"{layout.header_row}). Строк с ИИН: {plan.rows_with_value}.",
                 f"К проверке: {plan.rows_to_check} строк, разных ИИН — {len(plan.groups)}."]
        extra = []
        if plan.skipped_linked:
            extra.append(f"уже со ссылкой — {plan.skipped_linked}")
        if plan.skipped_hidden:
            extra.append(f"скрыто фильтром — {plan.skipped_hidden}")
        if plan.invalid:
            extra.append(f"с неверным ИИН — {len(plan.invalid)}")
        if extra:
            lines[-1] += " Пропускаются: " + ", ".join(extra) + "."
        if result.risks and not self._is_working_copy():
            lines.append("⚠ В файле есть: " + ", ".join(result.risks) + ". Программа не умеет сохранять "
                         "их без потерь, поэтому результат лучше записать в копию (галочка ниже).")
        if result.locked:
            lines.append("⚠ Файл сейчас открыт в Excel — закройте его перед запуском.")
        self.var_info.set("\n".join(lines))

    # ------------------------------------------------------------------ запуск и остановка
    def _start(self) -> None:
        if self.worker is not None:
            return
        if not self.var_xlsx.get():
            self.warn("Нет файла", "Сначала выберите файл Excel.")
            return
        if self.analysis is None:
            self.warn("Файл не прочитан", self.var_info.get() or "Файл ещё читается, подождите.")
            return
        if not self.var_out.get():
            self.warn("Нет папки", "Выберите папку для снимков.")
            return
        options = self._options()
        if not options.registries:
            self.warn("Не выбран реестр", "Отметьте хотя бы один реестр.")
            return
        plan = plan_from_tasks(self.analysis.layout, self.analysis.tasks, self.analysis.risks, options)
        if not plan.groups and not plan.invalid:
            hint = ""
            if plan.skipped_linked:
                hint = ("\n\nЧтобы проверить заново, отметьте «Проверить заново и те строки, "
                        "где ссылка уже есть».")
            self.info("Нечего проверять", nothing_to_do(plan) + hint)
            return
        minutes = estimate_minutes(len(plan.groups), len(options.registries), self.settings)
        duration = "меньше минуты" if minutes < 1 else "около " + _duration(round(minutes) * 60)
        where = ("в КОПИЮ файла (исходный не изменится)" if options.save_copy
                 else "в выбранный файл (перед этим будет сделана резервная копия)")
        question = (f"Будет проверено ИИН: {len(plan.groups)} (строк: {plan.rows_to_check}).\n"
                    f"Реестров: {len(options.registries)}. Примерное время: {duration}.\n\n"
                    f"Ссылки будут записаны {where}.\n\n"
                    "Файл Excel во время проверки должен быть закрыт.\n\nНачать?")
        if not self.ask("Начать проверку?", question):
            return

        self._save_prefs()
        self.stop_event.clear()
        self._set_running(True)
        self.progress.configure(maximum=max(1, len(plan.groups)), value=0)
        self._started_at = time.monotonic()
        self.var_status.set("Проверка началась…")

        def work() -> None:
            try:
                runner = Runner(
                    options, self.settings, stop_event=self.stop_event,
                    browser_factory=self._browser_factory,
                    on_log=lambda level, text: self.queue.put(("log", level, text)),
                    on_progress=lambda done, total, s: self.queue.put(
                        ("progress", done, total, s.found, s.not_found, s.errors)))
                summary = runner.run()
            except BaseException as exc:             # окно обязано узнать, что работа кончилась
                summary = Summary(fatal=f"Непредвиденная ошибка: {type(exc).__name__}: {exc}")
            self.queue.put(("done", summary))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _stop(self) -> None:
        if self.worker is None:
            return
        self.stop_event.set()
        self.btn_stop.configure(state="disabled")
        self.var_status.set("Останавливаю: закончу текущий ИИН и сохраню файл…")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        for widget in (self.btn_start, self.btn_xlsx, self.btn_out,
                       *self.registry_checks, *self.option_checks):
            widget.configure(state=state)
        self.cmb_sheet.configure(state="disabled" if running or self.analysis is None else "readonly")
        self.btn_stop.configure(state="normal" if running else "disabled")

    def _on_done(self, summary: Summary) -> None:
        self.worker = None
        self._set_running(False)
        text = summary.text()
        if summary.elapsed_s >= 1 and not summary.fatal:
            text += f"\nВремя работы: {_duration(summary.elapsed_s)}"
        self.var_status.set(text.splitlines()[0])
        if summary.workbook and Path(self.var_xlsx.get()) != summary.workbook and not summary.fatal:
            # Результат ушёл в копию — дальше работаем с ней, иначе повторный запуск начнёт с нуля.
            self.var_xlsx.set(str(summary.workbook))
            self._copy_auto = False
            self.var_copy.set(False)
            self._log("info", f"Дальше используется файл с результатом: {summary.workbook.name}")
        if self._closing:
            self.root.destroy()
            return
        if summary.fatal:
            self.error("Проверка не выполнена", text)
        elif summary.errors or summary.aborted or summary.notes or summary.invalid:
            self.warn("Проверка завершена с замечаниями", text)
        else:
            self.info("Готово", text)
        if self.var_xlsx.get():
            self._start_analysis()

    # ------------------------------------------------------------------ сообщения от рабочего потока
    def _poll(self) -> None:
        """Разбор сообщений рабочего потока. Ошибка в обработке одного сообщения не должна
        останавливать сам разбор — иначе окно навсегда застынет в состоянии «идёт проверка»."""
        try:
            while True:
                message = self.queue.get_nowait()
                try:
                    self._handle(message)
                except tk.TclError:
                    raise
                except Exception as exc:
                    self._log("error", f"Внутренняя ошибка окна: {type(exc).__name__}: {exc}")
        except queue.Empty:
            pass
        except tk.TclError:
            return                                   # окно уже закрыто
        if not self._closing or self.worker is not None:
            self.root.after(POLL_MS, self._poll)

    def _handle(self, message: tuple) -> None:
        kind = message[0]
        if kind == "log":
            self._log(message[1], message[2])
        elif kind == "progress":
            _, done, total, found, not_found, errors = message
            self.progress.configure(maximum=max(1, total), value=done)
            status = (f"Проверено {done} из {total}  ·  найдено {found}  ·  "
                      f"не найдено {not_found}  ·  ошибок {errors}")
            if 0 < done < total and not self.stop_event.is_set():
                left = (time.monotonic() - self._started_at) / done * (total - done)
                status += f"  ·  осталось ≈ {_duration(left)}"
            if not self.stop_event.is_set():
                self.var_status.set(status)
        elif kind == "analysis" and message[1] == self._analysis_token:
            self._on_analysis(message[2])
        elif kind == "analysis_error" and message[1] == self._analysis_token:
            self.analysis = None
            self.cmb_sheet.configure(values=[], state="disabled")
            self.var_info.set(message[2])
        elif kind == "done":
            self._on_done(message[1])

    def _log(self, level: str, text: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", time.strftime("%H:%M:%S  "), "time")
        self.txt_log.insert("end", text + "\n", level if level in ("error", "warn", "found") else "")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ------------------------------------------------------------------ прочее
    def _open_out(self) -> None:
        path = Path(self.var_out.get()) if self.var_out.get() else None
        if path is None or not path.is_dir():
            self.info("Папки ещё нет", "Папка со снимками появится после первой проверки.")
            return
        self._open(path)

    def _open_xlsx(self) -> None:
        if not self.var_xlsx.get():
            return
        if self.worker is not None:
            self.warn("Идёт проверка", "Не открывайте файл в Excel до конца проверки: "
                      "программа не сможет сохранить в него результат.")
            return
        self._open(Path(self.var_xlsx.get()))

    def _open(self, path: Path) -> None:
        try:
            open_in_system(path)
        except OSError as exc:
            self.error("Не удалось открыть", f"{path}\n{exc}")

    def _on_close(self) -> None:
        if self.worker is None:
            self.root.destroy()
            return
        if self.ask("Идёт проверка", "Проверка ещё не закончена.\n\nОстановить её, сохранить "
                    "уже сделанное и закрыть программу?"):
            self._closing = True
            self._stop()


def _duration(seconds: float) -> str:
    seconds = int(round(seconds))
    if seconds < 60:
        return f"{seconds} с"
    minutes, rest = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes} мин {rest:02d} с"
    hours, minutes = divmod(minutes, 60)
    return f"{hours} ч {minutes:02d} мин"


def _enable_dpi_awareness() -> None:
    """Без этого окно в Windows при масштабе 125–150 % выглядит размытым."""
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def run_gui(settings: Settings) -> int:
    _enable_dpi_awareness()
    root = tk.Tk()
    App(root, settings)
    root.mainloop()
    return 0


def show_fatal(text: str) -> None:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(APP_NAME, text, parent=root)
    root.destroy()
