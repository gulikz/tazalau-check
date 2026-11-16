"""ИИН: контрольная цифра и приведение значений ячеек Excel."""
import pytest

from tazalau.iin import check_digit, checksum_ok, normalize_iin

from conftest import IIN_A, IIN_B, IIN_C


def test_check_digit_first_pass():
    # 9·1+9·2+1·3+3·4+9·5+9·6+9·7+0·8+0·9+0·10+1·11 = 215; 215 mod 11 = 6
    assert check_digit("99139990001") == 6


def test_check_digit_second_pass_and_unused_number():
    # Первый проход даёт остаток 10 -> применяется второй набор весов.
    assert check_digit("00000000010") == 1
    # Оба прохода дают 10 -> номера с таким началом не существует.
    assert check_digit("00000000281") is None
    assert not any(checksum_ok("00000000281" + digit) for digit in "0123456789")


@pytest.mark.parametrize("value", ["000000000000", IIN_A, IIN_B, IIN_C])
def test_checksum_ok(value):
    assert checksum_ok(value)


@pytest.mark.parametrize("value", ["000000000001", "12345678901", "1234567890123", "abcdefghijkl", ""])
def test_checksum_rejects(value):
    assert not checksum_ok(value)


def test_checksum_catches_most_single_digit_typos_but_not_the_11th_digit():
    """Что именно гарантирует официальный алгоритм — чтобы не обещать лишнего.

    Вес 11-й цифры равен 11 (остаток 0 по модулю 11), поэтому ошибка в ней не обнаруживается.
    В остальных позициях замена одной цифры ловится почти всегда (редкое исключение — когда
    после замены срабатывает второй проход и остаток случайно совпадает).
    """
    caught = total = 0
    for iin in (IIN_A, IIN_B, IIN_C, "000000000000"):
        for position in range(12):
            if position == 10:
                continue                          # известный пробел алгоритма, проверен ниже
            for digit in "0123456789":
                if digit != iin[position]:
                    total += 1
                    caught += not checksum_ok(iin[:position] + digit + iin[position + 1:])
    assert caught / total > 0.97
    unnoticed = [d for d in "0123456789"
                 if d != IIN_A[10] and checksum_ok(IIN_A[:10] + d + IIN_A[11:])]
    assert len(unnoticed) == 9                    # ни одна замена 11-й цифры не замечена


def test_adjacent_digit_swap_is_caught():
    for iin in (IIN_A, IIN_B, IIN_C):
        for i in range(9):                        # пары до 10-й цифры включительно
            if iin[i] != iin[i + 1]:
                assert not checksum_ok(iin[:i] + iin[i + 1] + iin[i] + iin[i + 2:]), (iin, i)


@pytest.mark.parametrize("value", [None, "", "   ", "\u00a0", "'"])
def test_empty_cells(value):
    parsed = normalize_iin(value)
    assert parsed.empty and not parsed.ok and parsed.error is None


@pytest.mark.parametrize("value", [
    IIN_A, f"  {IIN_A}  ", f"{IIN_A[:6]} {IIN_A[6:]}", f"'{IIN_A}", f"{IIN_A}\n",
    f"{IIN_A[:3]}\u00a0{IIN_A[3:]}", f"{IIN_A}.0", f"{IIN_A},00", int(IIN_A), float(IIN_A),
])
def test_accepted_forms(value):
    parsed = normalize_iin(value)
    assert parsed.ok and parsed.iin == IIN_A and not parsed.padded


def test_number_with_lost_leading_zeros_is_restored():
    two = normalize_iin(int(IIN_B))            # 10 цифр
    one = normalize_iin(float(IIN_C))          # 11 цифр
    assert (two.iin, two.padded) == (IIN_B, True)
    assert (one.iin, one.padded) == (IIN_C, True)
    assert normalize_iin(0).ok is False        # одиночный ноль — не «двенадцать нулей»


def test_short_text_is_a_typo_not_lost_zeros():
    parsed = normalize_iin(IIN_C[1:])          # те же 11 цифр, но текстом
    assert not parsed.ok and "11 цифр" in parsed.error


@pytest.mark.parametrize("value,fragment", [
    (IIN_A[:11] + "O", "кроме цифр"),          # буква O вместо нуля
    ("8,11E+11", "кроме цифр"),
    ("１２３４５６７８９０１２", "кроме цифр"),      # полноширинные цифры
    ("1234567890123", "13 цифр"),
    (float(IIN_A) + 0.5, "дробное"),
    (-5, "отрицательное"),
    (True, "логическое"),
    (float("nan"), "не число"),
    (12345, "5 цифр"),
])
def test_rejected_with_reason(value, fragment):
    parsed = normalize_iin(value)
    assert not parsed.ok and not parsed.empty and fragment in parsed.error


def test_bad_checksum_blocked_by_default_and_allowed_on_request():
    bad = IIN_A[:11] + str((int(IIN_A[11]) + 1) % 10)
    assert "контрольная цифра" in normalize_iin(bad).error
    assert normalize_iin(bad, allow_bad_checksum=True).iin == bad


def test_padded_number_must_still_pass_checksum():
    parsed = normalize_iin(int(IIN_B) + 1)     # 10 цифр, но после дополнения сумма не сходится
    assert not parsed.ok and "меньше 12 цифр" in parsed.error
    assert not normalize_iin(int(IIN_B) + 1, allow_bad_checksum=True).ok
