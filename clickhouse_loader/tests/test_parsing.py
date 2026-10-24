"""Чтение файлов: кодировка, разделитель, заголовок, «вечные кавычки», битые строки, Excel, Parquet."""
import datetime as dt
import decimal

import pyarrow as pa
import pytest

import ch_loader as L
from conftest import csv_source, read_all


# ─────────────────────────────────── кодировка ───────────────────────────────────

def test_encoding_utf8_checked_over_whole_file(write_file):
    """UTF-8 подтверждается проверкой всего файла, а не только начала."""
    path = write_file("a.csv", "имя;возраст\nАня;30\n")
    assert L.detect_encoding(path)[0] == "utf-8"


def test_encoding_bom(write_file):
    """BOM распознаётся сразу (файлы из Excel «CSV UTF-8»)."""
    path = write_file("a.csv", b"\xef\xbb\xbfa,b\n1,2\n")
    assert L.detect_encoding(path)[0] == "utf-8-sig"


def test_encoding_cp1251(write_file):
    """Русский текст в Windows-1251 определяется как cp1251."""
    text = "Город;Население\n" + "Москва;13000000\nСанкт-Петербург;5600000\nНовосибирск;1600000\n" * 50
    path = write_file("a.csv", text, encoding="cp1251")
    assert L.detect_encoding(path)[0] == "cp1251"


def test_utf8_error_far_from_start_is_detected(write_file):
    """Испорченный байт далеко от начала файла: файл не считается UTF-8."""
    good = ("a,b\n" + "x,y\n" * 1_000_000).encode()
    path = write_file("a.csv", good + "z,я\n".encode("cp1251"))
    assert L.detect_encoding(path)[0] != "utf-8"


# ─────────────────────────────────── разделитель ───────────────────────────────────

@pytest.mark.parametrize("text, delim, ncols", [
    ("a;b;c\n1;2;3\n4;5;6\n", ";", 3),
    ("a\tb\n1\t2\n", "\t", 2),
    ("a|b\n1|2\n", "|", 2),
    ('a,b\n"1,5",2\n"3,5",4\n', ",", 2),        # запятая внутри кавычек не путает
    ("a;b\n1,5;2\n3,5;4\n", ";", 2),            # десятичная запятая при разделителе ;
])
def test_sniff_delimiter(text, delim, ncols):
    """Разделитель выбирается по согласованности числа полей в строках."""
    dia, n, _ = L.sniff_dialect(text, ".csv")
    assert (dia.delimiter, n) == (delim, ncols)


def test_sniff_line_wrapped():
    """Каждая строка целиком в кавычках ("a,b,c") — распознаётся и разворачивается."""
    text = '"a,b,c"\n' + "".join(f'"{i},""x{i}"",{i}"\n' for i in range(20))
    dia, n, _ = L.sniff_dialect(text, ".csv")
    assert dia.line_wrapped and n == 3


def test_guess_header():
    """Заголовок определяется по тому, что первая строка не подходит под типы остальных."""
    assert L.guess_header([["id", "сумма"], ["1", "2.5"], ["2", "3"]]) is True
    assert L.guess_header([["1", "2.5"], ["2", "3"], ["3", "4"]]) is False   # все строки — данные
    assert L.guess_header([["имя", "город"], ["Аня", "Тверь"], ["Олег", "Сочи"]]) is True


def test_clean_names():
    """Пустые и повторяющиеся имена колонок исправляются."""
    assert L.clean_names(["id", "", "id", " a  b "]) == ["id", "col_2", "id_2", "a b"]


# ─────────────────────────────── «вечные кавычки» ───────────────────────────────

@pytest.mark.parametrize("raw, expected", [
    ('"ООО ""Ромашка"""', 'ООО "Ромашка"'),     # обёртка + удвоенные внутренние
    ('"""x"""', "x"),                           # несколько слоёв
    ("'abc'", "abc"),
    ('ООО "Ромашка"', 'ООО "Ромашка"'),         # кавычки внутри текста не трогаем
    ('"a" и "b"', '"a" и "b"'),                 # начинается и кончается кавычкой, но не обёртка
    ('"', '"'),
])
def test_unwrap_quotes(raw, expected):
    """Снимаются только лишние кавычки, целиком обрамляющие значение."""
    out, changed = L.unwrap_quotes(pa.array([raw]))
    assert out[0].as_py() == expected
    assert changed[0].as_py() is (raw != expected)


# ─────────────────────────────── битые строки CSV ───────────────────────────────

@pytest.mark.parametrize("line, expected", [
    ('7,"ООО "Ромашка"",7', ["7", 'ООО "Ромашка"', "7"]),
    ('1,"Кавычка "внутри" текста",2', ["1", 'Кавычка "внутри" текста', "2"]),
    ('1,"a""b",2', ["1", 'a"b', "2"]),          # корректное экранирование не ломается
    ('1,"abc""",2', ["1", 'abc"', "2"]),
    ("1,,3", ["1", "", "3"]),
    ("1,2,", ["1", "2", ""]),
    ('1,"не закрыта', None),
])
def test_tolerant_split(line, expected):
    """«Терпимый» разбор: кавычка закрывает поле только перед разделителем или концом строки."""
    assert L.tolerant_split(line, ",") == expected


def test_fast_path_refuses_malformed_quotes(write_file):
    """Быстрый парсер (pyarrow) склеил бы '"ООО "Ромашка""' молча, поэтому при некорректных
    кавычках он отказывается и передаёт файл строгому построчному разбору."""
    path = write_file("a.csv", 'id,name,qty\n1,"ООО "Ромашка"",7\n')
    src = csv_source(path)
    with pytest.raises(L._NeedRepair) as e:
        read_all(src)
    assert e.value.fast


def test_broken_quotes_repaired_end_to_end(write_file):
    """Полный анализ файла с битыми кавычками: строки исправлены, значения точные, всё учтено."""
    body = "".join(f'{i},"Имя {i}",{i}\n' for i in range(1, 6))
    path = write_file("a.csv", "id,name,qty\n" + body + '6,"ООО "Ромашка"",6\n' + '7,"мульти\nстрока",7\n')
    src = csv_source(path)
    an = L.analyze(src, unwrap=False)
    assert an.rows == 7
    assert src.issues.repaired_count == 1
    rows = read_all(src)
    assert rows[5] == ["6", 'ООО "Ромашка"', "6"]
    assert rows[6] == ["7", "мульти\nстрока", "7"]         # законный перевод строки сохранён


def test_trailing_empty_fields_dropped(write_file):
    """Лишние ПУСТЫЕ поля в конце строки (типичный экспорт с ';' в конце) отбрасываются без потерь."""
    path = write_file("a.csv", "a;b\n1;2;\n3;4;\n")
    src = csv_source(path, ";", ncols=2)
    L.analyze(src, unwrap=False)
    assert read_all(src) == [["1", "2"], ["3", "4"]]


def test_unparseable_row_is_reported_not_dropped(write_file):
    """Строку, которую нельзя разобрать, не теряем молча: она попадает в список отклонённых."""
    path = write_file("a.csv", "a,b\n1,2\n3,4,5\n6,7\n")
    src = csv_source(path)
    an = L.analyze(src, unwrap=False)
    assert an.rows == 2
    assert src.issues.rejected_count == 1
    assert src.issues.rejected[0][0] == 3      # номер строки в файле


def test_fast_and_strict_paths_give_identical_rows(write_file):
    """Быстрый (pyarrow) и строгий (csv) разборы дают одинаковый результат на корректном файле."""
    lines = ['id,text,n\n'] + [f'{i},"текст, с ""кавычками"" {i}",{i}\r\n' for i in range(5000)]
    lines.append('5000,"многострочное\nзначение",1\n')
    path = write_file("a.csv", "".join(lines))
    fast = csv_source(path)
    strict = csv_source(path)
    strict.fast = False
    assert read_all(fast) == read_all(strict)


def test_cp1251_file_read_correctly(write_file):
    """Файл в cp1251 читается с правильными русскими буквами."""
    path = write_file("a.csv", "город;n\nМосква;1\n", encoding="cp1251")
    src = csv_source(path, ";", encoding="cp1251")
    assert read_all(src) == [["Москва", "1"]]


# ─────────────────────────────────── Excel и Parquet ───────────────────────────────────

@pytest.mark.parametrize("value, expected", [
    (None, ""),
    (True, "true"),
    (5.0, "5"),                                  # целое число из Excel без ".0"
    (4985.010000000001, "4985.01"),              # 15 значащих цифр — точность Excel
    (dt.datetime(2024, 1, 2), "2024-01-02"),     # дата без времени
    (dt.datetime(2024, 1, 2, 3, 4, 5), "2024-01-02 03:04:05"),
    ("текст", "текст"),
])
def test_excel_cell(value, expected):
    """Значения ячеек Excel переводятся в текст без потери смысла."""
    assert L.excel_cell(value) == expected


def test_excel_source(tmp_path):
    """Лист Excel читается целиком, пустые строки в начале пропускаются."""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([])
    ws.append(["Дата", "Сумма"])
    ws.append([dt.date(2024, 1, 1), 10.5])
    path = str(tmp_path / "a.xlsx")
    wb.save(path)
    src = L.ExcelSource(path, ws.title, True, 2)
    assert read_all(src) == [["2024-01-01", "10.5"]]


def test_parquet_types_preserved(tmp_path):
    """Типы Parquet переносятся напрямую, значения — без искажений."""
    import pyarrow.parquet as pq
    t = pa.table({
        "u64": pa.array([2**63], pa.uint64()),
        "d": pa.array([decimal.Decimal("12.50")], pa.decimal128(10, 2)),
        "ts": pa.array([dt.datetime(2024, 1, 1, 10, 0, 0, 500000)], pa.timestamp("us")),
        "tsz": pa.array([dt.datetime(2024, 1, 1, tzinfo=dt.timezone.utc)], pa.timestamp("ms", tz="UTC")),
    })
    path = str(tmp_path / "a.parquet")
    pq.write_table(t, path)
    src = L.ParquetSource(path)
    assert [src.fixed_types[n].ch for n in src.names] == [
        "UInt64", "Decimal(10, 2)", "DateTime64(6)", "DateTime64(3, 'UTC')"]
    assert read_all(src) == [["9223372036854775808", "12.50", "2024-01-01 10:00:00.500000",
                              "2024-01-01 00:00:00.000"]]


# ─────────────────────────────────── адрес сервера ───────────────────────────────────

@pytest.mark.parametrize("raw, host, port, forced", [
    ("localhost", "localhost", None, None),
    ("ch.example.com:9440", "ch.example.com", 9440, None),
    ("https://ch.example.com:8443/", "ch.example.com", 8443, ("http", True)),
    ("tcp://10.0.0.1", "10.0.0.1", None, ("native", False)),
])
def test_parse_address(raw, host, port, forced):
    """Адрес принимается в любом привычном виде: хост, хост:порт, URL."""
    assert L.parse_address(raw) == (host, port, forced)


def test_candidate_endpoints_order():
    """Без порта перебираются стандартные порты; по известному порту сначала пробуется его протокол."""
    assert L.candidate_endpoints("h", None, None)[0] == ("http", False, 8123)
    assert L.candidate_endpoints("h", 9440, None)[0] == ("native", True, 9440)
    assert L.candidate_endpoints("x.clickhouse.cloud", None, None)[0] == ("http", True, 8443)
