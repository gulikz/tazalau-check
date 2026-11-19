# Soft-collection segmentation: auto loans, DPD 1-30

`soft_collection_segmentation.ipynb` builds a PD model for delinquent auto loans and assigns each contract a treatment segment: hard rules first, then a risk x capacity x engagement matrix. Every contract in today's DPD 1-30 stock is scored on its current day in arrears. It comes out with its segment, today's treatment, a challenger arm and its compliance dates.

## Files
| file | what |
|---|---|
| `soft_collection_segmentation.ipynb` | the notebook, saved without outputs |
| `demo_data.py` | synthetic generator with the real schema of the 13 tables (fake people, including partial payments / re-ageing); `make_demo_tables()` |
| `sql/hes_pkb_monitoring_agg.sql` | Vertica aggregate of `dds.hes_pkb_monitoring_event` (long format, as of `:snapshot_date`). **The IIN and event-timestamp column names are placeholders** (`*__PLACEHOLDER`) |
| `out/`, `out_demo/` | run outputs (real / DEMO). Git ignores both folders and `data/` because they contain IINs |

## Method in brief
1. **Episodes and checkpoints.**
   - `ovrd` gives a monthly maximum DPD per contract. Each row implies an overdue start, `start = act_date - (dpd - 1)`.
   - A new episode needs a **cure**: its start must lie more than `EPISODE_GAP_DAYS` (5) after the latest overdue day seen so far. A partial payment that moves the oldest unpaid instalment forward while the client stays overdue (**re-ageing**) is merged into the running episode and flagged `reaged`.
   - Each episode is scored on the checkpoints `CHECKPOINT_DAYS` (days 1, 8, 15, 22, 29 in arrears) **while it is still overdue**. An episode that survives k days without a cure is much riskier: in the DEMO the bad rate goes 22% (day 1), 33% (day 8), 36% (day 15), 40% (day 22) and 49% (day 29).
   - `bad` = the contract reaches 90+ DPD within 6 months after the checkpoint. Secondary targets: `self_cure30` and `roll31_60d`.
2. **Features.**
   - There are 13 feature blocks, `fb_*`: contract, DPD history, client profile, internal and external debits, enforcement, calls, app logins, exclusions, bankruptcy, applications, rejections and PKB. Each block uses only events **before** the checkpoint date. The DPD history uses only *previous* episodes, and `dh_days_in_arrears` carries the checkpoint.
   - Profile features come from **calendar month-end snapshots** only, with the 3-month trend taken exactly three month-ends earlier. `ver_df` switched to daily snapshots in late 2026, so this keeps training and the stock on the same cadence. The stock's rules read the latest daily snapshot.
   - An **as-of self-test** rebuilds the data as it looked on two entry dates and two later checkpoint dates. It checks that every block gives identical features.
   - Sources that start late (geo, vnpt, vneshpt, PKB) are used for rules only (`MODEL_BLOCKS='auto'`).
   - Values derived from `contr.end_date` are kept out of the model: the term, the remaining term, the instalment proxy and the capacity ratios built on them. `end_date` is an extraction-time value and moves on restructuring, so these feed only the segmentation axes.
3. **Model.**
   - LightGBM, with a fallback to sklearn `HistGradientBoosting`, benchmarked against a logistic regression. One model covers every checkpoint, with days in arrears as a feature.
   - Time split **by episode**: OOT = the episodes of the last 3 labelled entry months. Validation is grouped by contract.
   - Train PDs are **out-of-fold**. The report gives the Gini per checkpoint, and the feature PSI from train to OOT and from train to the **current stock**. For the stock comparison, train is re-weighted to the stock's days-in-arrears mix.
4. **Grades.** G1-G5 come from the train PD quantiles of new entries (40/65/85/95%), with fixed cutpoints for every checkpoint. An episode still overdue on day 15 therefore usually sits one or two grades higher than it did on day 1. The monotonicity check uses a significance test.
5. **Segments.**
   - **Hard rules**, first match wins: R01 deceased, R02 bankruptcy, R03 bankruptcy attempt, R06 abroad/skip, R07 FPD fraud check, R08 gambling, R09 collateral. The actions of every other rule hit are kept in `additional_actions`.
   - **Modifiers** R04 (social exclusion lists) and R05 (open ИР3/ИР4 enforcement orders accepted within 2 years) do not replace the risk segment, because risk inside them still spans G1-G5. They add "no auto-debit retry / early restructuring" actions on top.
   - Model segments: S1 self-cure, S2 low-risk remind, S3 standard, S4 hardship, S5 high-risk contactable, S6 high-risk unreachable.
   - The **backtest** runs at every checkpoint, with Wilson confidence intervals, z-tests between neighbouring segments, a rule sanity table (warns when a high-touch rule is less risky than S2), rule x grade tables and an axis check.
   - Segments reached only through PKB events (R08, R09, the PKB part of R02) are marked `unvalidated`, because there is no labelled PKB history yet.
   - Treatments live in `SEGMENT_TREATMENT` and follow the Kazakhstan guardrails written into the notebook (section 9.0): the 10-day notice, the art. 61 change-of-terms window, call and visit caps (`contact_caps` column), and the debit limit of 50% / 2xПМ. Protected social and pension inflows are not counted as capacity.

## How to run
0. **PKB first (real data).** Replace the placeholders in `sql/hes_pkb_monitoring_agg.sql` and run it. Save the result as `DATA_DIR/hes_pkb_agg.parquet`, or keep it as a `hes_pkb_agg` DataFrame in the kernel. Alternatively, set `CONFIG_OVERRIDES = {'RUN_PKB_SQL': True}` in your Vertica kernel, and section 3 runs the SQL through your `query()` helper before loading. Without the aggregate, R08, R09, the PKB part of R02 and the «Стоп-Кредит» flag are **off**. The run prints a NOTICE when this happens.
1. **Your kernel (way a).** Load the tables under their usual names (`contr`, `ovrd`, `ver_df_all_2024_2026`, `bankrupt`, `ocrm_call`, `vneshpt_ot`, `vnpt`, `excl`, `vneshpt`, `geo`, `dm`, `dm_rej`, optionally `hes_pkb_agg`). Then run:
   ```python
   CONFIG_OVERRIDES = {'MODE': 'real'}          # a missing / misnamed table then fails loudly
   %run -i /full/path/to/credit_segments/soft_collection_segmentation.ipynb
   ```
   - Your DataFrames are read, never modified.
   - The run defines about 260 public globals with generic names (`X`, `S`, `OB`, `CONFIG`, `log`, `psi`, `col`, `MODEL`, `fig`, `ax`, ...), plus helpers whose names start with `_`. They **overwrite your own variables of the same name**, so rename anything you need first.
   - pandas display options and matplotlib settings are restored at the end of the run.
   - Relative `DATA_DIR` / `OUT_DIR` paths are resolved against the notebook's folder, so outputs land in `credit_segments/out/` and not in the kernel's working directory.
2. **Parquet (way b).** Set `CONFIG['DATA_DIR']` and the file names in `CONFIG['FILES']`; globs are allowed, e.g. `dm_*_rejected.parquet`. Only the needed columns and the population's keys are read through a pyarrow push-down filter. Keys are normalised *before* filtering, so CHAR-padded keys, `.0` float artefacts and lost leading zeros still match. Rows that match only after normalisation are reported.
3. **DEMO (way c).** `MODE='demo'` uses `demo_data.py`. `MODE='auto'` falls back to DEMO **only when no table is found at all**; a partial set (say `ovrd` loaded as `ovrd_zp`) stops with the list of found and missing tables. DEMO results go to `out_demo/` and carry `run_mode = 'demo'`. A DEMO run takes about 30 s with 8,000 synthetic contracts.
4. **Headless.** The kernel inherits `SEG_*` environment variables, which override `CONFIG`: `SEG_MODE`, `SEG_DROP_TABLES` (comma list), `SEG_OUT_DIR`, `SEG_DATA_DIR`, `SEG_DEMO_N_CONTRACTS`, `SEG_USE_LIGHTGBM`, `SEG_STRICT_BLOCKS` (1 = a failing block raises) and `SEG_N_JOBS`. For example:
   ```bash
   SEG_DROP_TABLES=geo,vnpt python -c "import nbformat; from nbclient import NotebookClient; \
   nb = nbformat.read('soft_collection_segmentation.ipynb', 4); NotebookClient(nb, timeout=1200, kernel_name='python3').execute(); \
   nbformat.write(nb, 'executed.ipynb')"
   ```

Only `contr` and `ovrd` are required. A missing, empty or broken optional table switches its block off with a warning. The rules that depend on it cannot fire, and block failures are listed in `run_summary.json` (`block_failures`). On a shared server, set `N_JOBS` to the cores that are actually free; the default is half the CPUs.

## Outputs (`out/`; DEMO: `out_demo/`)
- `segments_current_stock.parquet` / `.xlsx`: one row per stock contract. The workbook has a `legend` sheet that explains every column. Main columns:
  - identity and scoring: `run_mode`, `days_in_arrears`, `pd` (current day), `pd_at_entry`, `risk_grade`;
  - segment: `segment`, `segment_reason`, `modifiers`, `additional_actions`, `validation`;
  - treatment: `treatment`, `treatment_now`, `operator_due_now`, `contact_caps`, `arm` (champion / challenger / none);
  - flags: `capacity`, `engaged`, `debitable_funds`, `no_credit_offer`, `legal_review_flag`, `exclusion_review_flag`, `third_party_contact_flag`;
  - compliance dates: `notify_due_date`, `notify_overdue`, `restructuring_window_end`;
  - `top_drivers` (TreeSHAP, in business terms).
- `segments_history.parquet`: labelled observations at every checkpoint, with PD, grade, segment, cluster and the outcomes.
- `segment_report.xlsx`, sheets:
  - data and features: `data_quality`, `block_coverage`, `block_failures`, `load_log`, `feature_importance`, `psi_dropped`, `psi_stock`, `pkb_rule_map`;
  - model and grades: `model_metrics`, `metrics_by_checkpoint`, `grades`, `grades_by_checkpoint`;
  - segments: `segments_backtest`, `backtest_by_checkpoint`, `segment_tests`, `rule_sanity`, `rule_x_grade`, `segments_current`, `treatments`, `clusters`;
  - history and monitoring: `monthly`, `checkpoints`, `dealers`, `monitoring`.
- `model/`: `lgbm_model.txt`, `model_bundle.pkl` (model, LR benchmark, features, category levels, cutpoints, checkpoints, config), and `features.json`, `config.json`, `cutpoints.json`.
- `charts/*.png` (5 charts) and `run_summary.json`, which includes the Gini per checkpoint, train-to-stock drift, rule warnings, the as-of test result and block failures.

## Main knobs (`CONFIG`, or `CONFIG_OVERRIDES` from your kernel)
| knob | default | meaning |
|---|---|---|
| `MODE` | `auto` | `auto` = globals -> parquet -> DEMO only if nothing is found; `real`; `demo` |
| `CHECKPOINT_DAYS` | 1, 8, 15, 22, 29 | days in arrears at which still-overdue episodes are scored. `[1]` = entry-only model; `[6]` = "score on day 6" (this replaces the old `OBS_OFFSET_DAYS`) |
| `SOFT_DPD_MIN/MAX` | 1 / 30 | soft-collection window of the current stock |
| `BAD_DPD`, `HORIZON_MONTHS` | 90 / 6 | target definition |
| `EPISODE_GAP_DAYS` | 5 | a new episode needs this many days without overdue after the previous one |
| `OOT_MONTHS`, `VALID_SHARE` | 3 / 0.2 | time split (by episode entry month) |
| `GRADE_QUANTILES` | 0.40, 0.65, 0.85, 0.95 | PD cutpoints G1..G5 (train PD of new entries) |
| `MODEL_BLOCKS` | `auto` | a block feeds the model only if its source starts before the train P10 `obs_date` |
| `MODEL_EXCLUDE` | snapshot age, every `end_date`-based value | time artefacts and leakage risks |
| `PROFILE_MONTH_END_ONLY` | `auto` | model features from month-end `ver_df` snapshots only |
| `DM_ALLOW_COLUMNS` / `DM_DENY_PATTERN` | None / lifecycle words | dm leakage guard. Auto mode prints the selected columns for sign-off; set the reviewed list on real data |
| `PSI_DROP_THRESHOLD` | 0.25 | unstable features are dropped (PSI train -> OOT); train -> stock drift is warned about |
| `CAPACITY_RATIO_MIN`, `DEPO_COVER_MIN`, `CONTACT_RATE_ENGAGED` | 1.5 / 1.0 / 0.5 | segment axes |
| `INSTALMENT_RATE` | 0.18 | annual rate of the annuity instalment proxy (used with `MONTH_LOAN_PAY_AMOUNT`) |
| `FPD_MIN_DAYS` | 8 | R07: 1st/2nd instalment unpaid for at least this many days |
| `ENFORCEMENT_OPERTYPES`, `ENFORCEMENT_MAX_AGE_DAYS` | ИР3, ИР4 / 730 | R05 modifier definition |
| `ABROAD_NOTE_DAYS` | 90 | R06: CRM "abroad" notes older than this are ignored |
| `R02_PKB_COURT` | False | True: the PKB "court" status (possibly at another creditor) also stops collection. False: `legal_review_flag` only |
| `CHALLENGER_SHARE` | 0.15 | champion/challenger routing by contract hash (segments with a challenger treatment) |
| `N_JOBS` | half the CPUs | LightGBM / KMeans threads |
| `COMPLIANCE` | memo values | notice / restructuring days, call and visit caps, ПМ, MRP, debit limits |

## Data caveats
- **Survivorship.** `contr` must include closed, written-off and sold contracts, or the history under-represents the worst outcomes. The notebook counts `ovrd` contracts that look like portfolio contracts but are missing from `contr`, and warns about them. `close_date`, `status_name`, `restructing_fl` and `end_date` are extraction-time values. They and everything derived from them (term, instalment proxy, `capacity_ratio`, `depo_cover`, `debitable_cover`) never enter the model.
- **`ovrd` holds only monthly maxima.** Cures and re-entries inside one month are invisible, and the episode end is the date of the month's maximum. A cure followed by a new overdue within `EPISODE_GAP_DAYS` is treated as one episode. For the last, partial month, the stock keeps only rows still overdue on the cut (`STOCK_ONLY_AT_CUT`).
- **Static extracts.** `contr` and `dm` are current snapshots, so the as-of test cannot catch lifecycle fields in them. Keep `dm` to application-time fields (`DM_ALLOW_COLUMNS`).
- **`excl`** is a snapshot list. History uses only entries whose `updated_date` is before `obs_date`; entries without `updated_date` count only for today's stock.
- **Treatment contamination.** Outcomes reflect today's SC 4-30 calling. The backtest is *risk under the current treatment*, not uplift; use the challenger arm to measure uplift.
- **Short histories.** `geo` starts 2025-09, `vnpt` 2026-01, `vneshpt` 2026-03 and the PKB events 2026-09-07. They are overlay/rules only until they have 6+ months of labelled history; until then the PKB-driven segments are `unvalidated`.
- **Demo numbers are not real.** The synthetic data has a planted signal and inflated delinquency rates, so its Gini and bad rates say nothing about the real portfolio. Some of the 77 `attr_*` filler columns in the synthetic `dm` are pure noise and can still appear among the top features.
- The regulatory constants come from a web-search memo (2026). Items marked **[verify]** must be confirmed by Legal on adilet.zan.kz.
