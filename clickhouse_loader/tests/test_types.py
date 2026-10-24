"""Определение типов, совместимость с типами таблицы и строгая конвертация значений."""
import pyarrow as pa
import pytest

import ch_loader as L
from conftest import profile_of


# ─────────────────────────── выбор типа для новой таблицы ───────────────────────────

@pytest.mark.parametrize("values, expected", [
    (["1", "2", "255"], "UInt8"),                       # минимальная разрядность
    (["-1", "100"], "Int8"),
    (["1", "70000"], "UInt32"),
    (["9223372036854775808"], "UInt64"),                # больше Int64, но влезает в UInt64
    (["123456789012345678901234567890"], "Int128"),     # больше 64 бит
    (["00123", "00456"], "String"),                     # ведущие нули — это код, не число
    (["true", "FALSE"], "Bool"),
    (["1.5", "2.25"], "Decimal(9, 2)"),                 # точный тип вместо Float
    (["1,5", "22,75"], "Decimal(9, 2)"),                # десятичная запятая
    (["1e10", "2.5"], "Float64"),                       # экспонента → Float64
    (["2024-01-31", "2023-12-01"], "Date"),
    (["31.01.2024", "1.2.2023"], "Date"),               # ДД.ММ.ГГГГ, в т.ч. без ведущих нулей
    (["1950-05-01"], "Date32"),                         # раньше 1970 — нужен Date32
    (["2024-01-01 10:00:00", "2024-01-02"], "DateTime"),  # смесь даты и даты-времени
    (["2024-01-01 10:00:00.5"], "DateTime64(3)"),
    (["2024-01-01T10:00:00.123456"], "DateTime64(6)"),
    (["2024-01-01T10:00:00Z"], "DateTime('UTC')"),      # Z — момент в UTC
    (["0f8fad5b-d9cb-469f-a165-70867728950e"], "UUID"),
    (["Москва", "Казань"], "String"),
])
def test_propose_type(values, expected):
    """Тип подбирается по всем значениям так, чтобы ничего не потерять."""
    assert profile_of(values).propose().ch == expected


def test_invalid_date_stays_string():
    """Несуществующая дата (30 февраля) не превращается молча в 2 марта — колонка остаётся текстом."""
    assert profile_of(["2023-02-28", "2023-02-30"]).propose().ch == "String"


def test_date_out_of_clickhouse_range_stays_string():
    """Даты вне диапазона ClickHouse (1100 год) хранятся текстом, а не обрезаются."""
    assert profile_of(["1100-05-10"]).propose().ch == "String"


def test_nulls_make_nullable_only_for_typed_columns():
    """Пустые значения дают Nullable у чисел; у текста пустая строка остаётся пустой строкой."""
    assert profile_of(["1", "", "NULL", "3"]).propose().ch == "Nullable(UInt8)"
    assert profile_of(["a", "", "b"]).propose().ch == "String"


def test_one_bad_value_late_in_column_turns_it_into_string():
    """Тип определяется по ВСЕМ строкам: одно нечисловое значение в конце делает колонку текстом
    и запоминается, где оно встретилось."""
    p = L.ColumnProfile("x")
    p.update(pa.array([str(i) for i in range(1000)]), 0)
    p.update(pa.array(["5", "abc"]), 1000)
    assert p.propose().ch == "String"
    assert p.killed_by["int"] == (1001, "abc")


def test_low_cardinality_for_repeated_strings():
    """Строки с малым числом уникальных значений получают LowCardinality(String)."""
    p = profile_of(["Москва", "СПб", "Казань"] * 1000)
    assert p.propose().ch == "LowCardinality(String)"


def test_ambiguous_day_month_is_reported():
    """Если все дни ≤ 12, формат ДД/ММ vs ММ/ДД неоднозначен — об этом предупреждаем."""
    p = profile_of(["01/02/2024", "03/04/2024"])
    assert any("ДД/ММ" in a for a in p.ambiguities())


def test_ambiguous_thousands_comma_is_reported():
    """'1,234' может быть разделителем тысяч — предупреждаем, а не молча делаем 1.234."""
    p = profile_of(["1,234", "5,678"])
    assert any("тысяч" in a for a in p.ambiguities())


# ─────────────────────────── совместимость с существующей таблицей ───────────────────────────

@pytest.mark.parametrize("values, ch_type, ok", [
    (["10", "300"], "UInt8", False),            # переполнение (ClickHouse сам бы молча обрезал)
    (["10", "200"], "UInt8", True),
    (["-1"], "UInt32", False),
    (["1.25"], "Decimal(9, 1)", False),         # было бы округление
    (["1.25"], "Decimal(9, 2)", True),
    (["123456789.5"], "Decimal(9, 1)", False),  # целая часть не влезает
    (["2024-01-01 10:00:00"], "Date", False),   # время было бы отброшено
    (["2024-01-01 10:00:00.123"], "DateTime64(3)", True),
    (["2024-01-01 10:00:00.1234"], "DateTime64(3)", False),
    (["0", "1"], "Bool", True),
    (["abc"], "Int64", False),
    (["abcdef"], "FixedString(3)", False),
    (["anything"], "String", True),
])
def test_compat_with_target_type(values, ch_type, ok):
    """Данные проверяются на запись в тип колонки таблицы без потерь."""
    assert profile_of(values).compat(L.parse_ch_type(ch_type))[0] is ok


@pytest.mark.parametrize("text, kind, extra", [
    ("LowCardinality(Nullable(String))", "string", {"nullable": True, "lowcard": True}),
    ("Nullable(UInt16)", "int", {"bits": 16, "signed": False, "nullable": True}),
    ("Decimal(18, 4)", "decimal", {"precision": 18, "scale": 4}),
    ("Decimal64(3)", "decimal", {"precision": 18, "scale": 3}),
    ("DateTime64(6, 'Europe/Moscow')", "datetime", {"precision": 6, "tz": "Europe/Moscow"}),
    ("DateTime('UTC')", "datetime", {"precision": 0, "tz": "UTC"}),
    ("FixedString(8)", "fixedstring", {"length": 8}),
    ("IPv4", "raw", {}),
])
def test_parse_ch_type(text, kind, extra):
    """Разбор типов ClickHouse из system.columns."""
    s = L.parse_ch_type(text)
    assert s.kind == kind
    for k, v in extra.items():
        assert getattr(s, k) == v


# ─────────────────────────────────── конвертация ───────────────────────────────────

def convert(values, ch_type, profile_values=None):
    prof = profile_of(profile_values or values)
    spec = L.parse_ch_type(ch_type)
    return L.make_converter(prof, spec)(pa.array(values, pa.string()), 0).to_pylist()


def test_convert_canonical_forms():
    """Значения приводятся к виду, который ClickHouse разбирает однозначно."""
    assert convert(["+5", " 7 "], "Int32") == ["5", "7"]
    assert convert(["1,5", "2"], "Decimal(9, 2)") == ["1.5", "2"]
    assert convert(["31.1.2024"], "Date") == ["2024-01-31"]
    assert convert(["2024-1-5 7:05"], "DateTime") == ["2024-01-05 07:05:00"]
    assert convert(["2024-01-05 07:05:09.1"], "DateTime64(3)") == ["2024-01-05 07:05:09.100"]
    assert convert(["TRUE"], "Bool") == ["true"]
    assert convert(["1", ""], "Nullable(Int8)") == ["1", None]


def test_convert_rejects_overflow():
    """Значение вне диапазона типа останавливает загрузку с номером строки."""
    with pytest.raises(L.ConvertError, match="строка данных 2"):
        convert(["1", "300"], "UInt8", profile_values=["1"])


def test_convert_rejects_null_in_non_nullable():
    """Пустое значение в не-Nullable колонке — ошибка, если не разрешено значение по умолчанию."""
    with pytest.raises(L.ConvertError, match="NULL"):
        convert(["1", ""], "Int8")


def test_convert_fill_default_when_allowed():
    """С явным разрешением пустые значения заменяются значением по умолчанию типа."""
    prof = profile_of(["1", ""])
    spec = L.parse_ch_type("Int8")
    spec.fill_default = True
    assert L.make_converter(prof, spec)(pa.array(["1", ""]), 0).to_pylist() == ["1", "0"]


def test_convert_string_keeps_raw_value():
    """Текст передаётся как есть, включая пробелы по краям и 'NULL' как слово."""
    assert convert([" a ", "NULL"], "String") == [" a ", "NULL"]


def test_to_python_for_native_protocol():
    """Для TCP-протокола каноничные строки превращаются в значения Python нужных типов."""
    import datetime as dt
    import decimal
    spec = L.parse_ch_type("Nullable(Decimal(9, 2))")
    assert L.to_python(pa.array(["1.50", None]), spec) == [decimal.Decimal("1.50"), None]
    spec = L.parse_ch_type("DateTime('UTC')")
    v = L.to_python(pa.array(["2024-01-01 10:00:00Z"]), spec)[0]
    assert v == dt.datetime(2024, 1, 1, 10, tzinfo=dt.timezone.utc)
