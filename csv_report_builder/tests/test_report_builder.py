"""Тесты сборщика отчётов: чтение CSV, роли колонок, порядок иерархии, значения вместо формул."""
import csv
import os
import sys

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import report_builder as R  # noqa: E402

HEAD = ["Уровень", "Наименование", "GROUP_NAME", "OTD_NAME", "Из них",
        "Случаи 2026", "Сумма 2026, млн руб", "Случаи 2025", "Сумма 2025, млн руб"]


def write_csv(path, header, rows, encoding="cp1251", sep=";"):
    with open(path, "w", encoding=encoding, newline="") as f:
        w = csv.writer(f, delimiter=sep)
        w.writerow(header)
        w.writerows(rows)


def build(src, **cfg_extra):
    cfg = R.deep_merge(R.DEFAULTS, cfg_extra)
    out = str(src) + ".xlsx"
    R.build(str(src), cfg, out)
    return load_workbook(out)


@pytest.fixture
def hierarchy_dir(tmp_path):
    """Корректная структура: дети под родителем, «из них» — сразу под своей строкой, соседи перемешаны."""
    rows = [[0, "Всего", "", "", "", 600, 60.0, 500, 50.0]]
    for g, s in (("Терапия", 10.0), ("Хирургия", 40.0), ("Офтальмология", 10.0)):
        rows.append([1, g, g, "", "", 100, s, 100, s / 2])
        rows.append([1, "из них операции", g, "", "операции", 10, 1.0, 10, 1.0])
        for o, so in ((0, 1.0), (1, 5.0), (2, 3.0)):
            rows.append([2, f"{g} {o}", g, f"{g} {o}", "", 10, so, 10, so])
    write_csv(tmp_path / "rep_1_5.csv", HEAD, rows)
    return tmp_path


def test_to_numbers_comma_spaces_percent():
    nums, all_pct = R.to_numbers(pd.Series(["1 234,50", "-10", "1,234.5"]))
    assert list(nums) == [1234.5, -10.0, 1234.5]
    assert not all_pct
    nums, all_pct = R.to_numbers(pd.Series(["12,5%", "3%"]))
    assert list(nums) == [0.125, 0.03] and all_pct


def test_fix_header_restores_delta():
    assert R.fix_header("? Случаи") == "Δ Случаи"
    assert R.fix_header("?% Сумма") == "Δ% Сумма"


def test_roles_leading_zero_codes_dates_and_metrics(tmp_path):
    write_csv(tmp_path / "t.csv", ["Регион", "Код", "Дата", "Выручка 2026", "Выручка 2025", "Доля"],
              [["Москва", "01", "01.10.2026", "10,5", "5", "12,5%"],
               ["Казань", "02", "02.10.2026", "3", "0", "1%"]], encoding="utf-8", sep=",")
    rep = R.prepare(str(tmp_path / "t.csv"), "t.csv", R.file_config(R.DEFAULTS, "t.csv"), "")
    by = {c.name: c for c in rep.columns}
    assert by["Код"].role == "dimension" and by["Код"].values == ["01", "02"]
    assert by["Дата"].kind == "date"
    assert by["Выручка 2026"].kind == "money" and by["Доля"].kind == "pct"
    assert by["Δ Выручка"].values == [5.5, 3.0]
    assert by["Δ% Выручка"].values == [1.1, None]          # при 2025 = 0 — пусто


def test_hierarchy_sorted_inside_groups_subrows_stay(hierarchy_dir):
    ws = build(hierarchy_dir).worksheets[1]
    names = [r[1] for r in ws.iter_rows(min_row=8, values_only=True) if r[0] is not None]
    assert names == ["Всего",
                     "Хирургия", "из них операции", "Хирургия 1", "Хирургия 2", "Хирургия 0",
                     "Терапия", "из них операции", "Терапия 1", "Терапия 2", "Терапия 0",
                     "Офтальмология", "из них операции", "Офтальмология 1", "Офтальмология 2", "Офтальмология 0"]


def test_values_only_and_formulas_sheet(hierarchy_dir):
    wb = build(hierarchy_dir)
    for ws in wb.worksheets:
        if ws.title == R.FORMULAS_SHEET:
            continue
        assert not any(c.data_type == "f" for row in ws.iter_rows() for c in row), ws.title
        assert len(ws.conditional_formatting) == 0
    fs = wb[R.FORMULAS_SHEET]
    assert fs.sheet_state == "hidden"
    texts = [r[3] for r in fs.iter_rows(min_row=6, values_only=True)]
    assert any(t.startswith("=IF(N(") for t in texts)
    assert all(fs.cell(6 + i, 4).data_type == "s" for i in range(len(texts)))


def test_live_mode_writes_formulas(hierarchy_dir):
    wb = build(hierarchy_dir, table={"formulas": "live"})
    ws = wb.worksheets[1]
    assert any(c.data_type == "f" for row in ws.iter_rows() for c in row)


def test_flat_totals_are_values(tmp_path):
    write_csv(tmp_path / "s.csv", ["Товар", "Кол-во 2026", "Кол-во 2025"], [["a", 5, 2], ["b", 1, 4]])
    ws = build(tmp_path).worksheets[1]
    # шапка (5–6) -> номера колонок (7) -> «Итого» (8, на нём автофильтр: строку заголовков фильтра Excel
    # не фильтрует и не сортирует) -> данные (9…)
    assert ws.cell(5, 1).value == "Товар" and ws.cell(7, 1).value == 1
    total = [ws.cell(8, i).value for i in range(1, 6)]
    assert total == ["Итого", 6, 6, 0, 0.0]
    assert ws.auto_filter.ref.startswith("A8:")
    assert ws.cell(9, 1).value == "a" and ws.freeze_panes.endswith("9")


def test_report_number_and_skip(tmp_path):
    assert R.report_number("rep_app_stac_1_1_202610021617") == ("1.1", "rep app stac")
    assert R.report_number("19.1 ЖК") == ("19.1", "ЖК")
    assert R.report_number("продажи") == ("", "продажи")
    write_csv(tmp_path / "a.csv", ["x", "n"], [["a", 1]])
    write_csv(tmp_path / "черновик.csv", ["x", "n"], [["a", 1]])
    wb = build(tmp_path, files={"*черновик*": {"skip": True}})
    assert len([w for w in wb.worksheets if w.title not in (R.TOC_SHEET, R.HELP_SHEET, R.FORMULAS_SHEET)]) == 1


def fill_of(cell):
    return cell.fill.fgColor.rgb[-6:] if cell.fill and cell.fill.fill_type == "solid" else None


def test_palette_b_default_and_a_switch(hierarchy_dir):
    ws = build(hierarchy_dir).worksheets[1]
    assert fill_of(ws.cell(8, 2)) == "C9DCEC"                 # «Всего» — уровень 0 палитры B
    assert fill_of(ws.cell(9, 2)) == "DCE8F3"                 # уровень 1
    assert fill_of(ws.cell(5, 2)) == "E8ECF0"                 # шапка
    assert ws.cell(2, 2).font.color.rgb[-6:] == "1B3A5C"      # заголовок — глубокий синий
    detail = next(r for r in range(8, 30) if ws.cell(r, 2).value == "Хирургия 1")
    assert fill_of(ws.cell(detail, 4)) == "ECEFF2"            # ключевая колонка «Сумма 2026»
    ws = build(hierarchy_dir, style={"palette": "A"}).worksheets[1]
    assert fill_of(ws.cell(8, 2)) == "F4B183"
    ws = build(hierarchy_dir, style={"colors": {"levels": {0: "112233"}}}).worksheets[1]
    assert fill_of(ws.cell(8, 2)) == "112233"


@pytest.fixture
def ambiguous_dir(tmp_path):
    rows = [[f"ГКБ {i}", f"45{i}", 100 + i, 90 + i, f"{1.5 + i / 10:.1f}".replace(".", ","), 1000 + 7 * i]
            for i in range(6)]
    write_csv(tmp_path / "mo.csv", ["MO_NAME", "OKATO_REG", "Случаи 2026", "Случаи 2025", "Балл", "Показатель"], rows)
    return tmp_path


def answers(*items):
    it = iter(items)
    return R.Asker(inp=lambda prompt: next(it), out=lambda *a: None)


def test_plan_groups_files_by_structure(ambiguous_dir):
    for k in range(3):                                         # ещё 3 файла той же структуры и 1 другой
        write_csv(ambiguous_dir / f"mo_{k}.csv", ["MO_NAME", "OKATO_REG", "Случаи 2026", "Случаи 2025", "Балл",
                                                  "Показатель"], [["ГКБ", "450", 1, 2, "1,5", 7]])
    write_csv(ambiguous_dir / "other.csv", ["Товар", "Цена"], [["a", "1,5"]])
    cfg = R.deep_merge(R.DEFAULTS, {})
    plan = R.plan_questions(str(ambiguous_dir), cfg, {})
    assert len(plan["reports"]) == 5 and len(plan["structures"]) == 2 and len(plan["todo"]) == 2
    assert "positive_is" in plan["book"] and "period" in plan["book"]


def test_interview_saves_answers_and_stops_asking(ambiguous_dir):
    cfg = R.deep_merge(R.DEFAULTS, {})
    # колонки структуры: 1 MO_NAME, 2 OKATO_REG, 3 Случаи 2026, 4 Случаи 2025, 5 Δ Случаи, 6 Δ% Случаи,
    #                    7 Балл, 8 Показатель
    raw, changed = R.interview(str(ambiguous_dir), cfg, {}, answers(
        "Анализ МО", "", "9м", "2",          # заголовок, подзаголовок, период, рост — хорошо
        "2",                                 # 5-й вопрос: колонки — проверить в консоли
        "7,8", "3", "7",                     # вид: Балл -> другое число, Показатель -> разрез
        "6",                                 # подсвечивать только Δ%
        "",                                  # главная метрика — по умолчанию
        "2",                                 # строка «Итого» — нет
        "МО", "",                            # подписи техн. имён: MO_NAME -> МО, OKATO_REG — как есть
        ""))                                 # имена листов — как в списке
    assert changed
    assert raw["book"] == {"title": "Анализ МО", "subtitle": "", "period": "9м"}
    assert raw["coloring"]["positive_is"] == "good"
    assert raw["columns"]["Балл"] == {"role": "metric", "kind": "number"}
    assert raw["columns"]["Показатель"] == {"role": "dimension"}
    assert raw["columns"]["MO_NAME"]["label"] == "МО"
    assert raw["files"]["mo.csv"] == {"coloring": {"columns": ["Δ% Случаи"]}, "table": {"totals": "off"},
                                      "sheet": "mo"}
    cfg = R.deep_merge(R.DEFAULTS, raw)
    assert R.plan_questions(str(ambiguous_dir), cfg, raw)["count"] == 0     # второй раз — без вопросов
    ws = build(ambiguous_dir, **raw).worksheets[1]
    assert ws.cell(5, 1).value == "МО" and ws.cell(7, 1).value != "Итого"


def test_kind_change_applies_to_all_years(ambiguous_dir):
    cfg = R.deep_merge(R.DEFAULTS, {})
    raw, _ = R.interview(str(ambiguous_dir), cfg, {}, answers("", "", "", "", "2", "3", "2", "!"))
    assert raw["columns"]["Случаи 2026"]["kind"] == raw["columns"]["Случаи 2025"]["kind"] == "money"


def test_bad_numbers_reasked(ambiguous_dir):
    cfg = R.deep_merge(R.DEFAULTS, {})
    said = []
    it = iter(["", "", "", "", "2", "99", "abc", "", "!"])
    asker = R.Asker(inp=lambda p: next(it), out=lambda *a: said.append(" ".join(map(str, a))))
    R.interview(str(ambiguous_dir), cfg, {}, asker)
    assert sum("не понял" in x for x in said) == 2


def test_bang_accepts_all_defaults(ambiguous_dir):
    cfg = R.deep_merge(R.DEFAULTS, {})
    raw, _ = R.interview(str(ambiguous_dir), cfg, {}, answers("!"))
    assert raw["book"]["title"] == ambiguous_dir.name
    assert raw["files"]["mo.csv"] == {"sheet": "mo"}
    assert "Показатель" not in raw["columns"]            # без ответа вид не записывается — решает автоматика


def test_column_group_colours_and_no_sum_for_numbers(ambiguous_dir):
    wb = build(ambiguous_dir, columns={"Балл": {"kind": "number"}, "Показатель": {"role": "dimension"}})
    ws = wb.worksheets[1]
    heads = {}
    for c in range(1, ws.max_column + 1):       # первое вхождение подписи («Случаи» есть и в 2026, и в Δ)
        heads.setdefault(ws.cell(6, c).value or ws.cell(5, c).value,
                         fill_of(ws.cell(6, c)) or fill_of(ws.cell(5, c)))
    assert heads["Случаи"] == "DCE8F3" and heads["Балл"] == "D6ECEC"
    total = {ws.cell(5, c).value: ws.cell(8, c).value for c in range(1, ws.max_column + 1)}
    assert total["Балл"] is None                              # баллы в «Итого» не складываются


# ---------------------------------------------------------------- крайние случаи CSV
def write_raw(path, text, encoding="utf-8"):
    with open(path, "w", encoding=encoding, newline="") as f:
        f.write(text)


def test_read_ragged_dups_blank_lines_trailing_delims(tmp_path):
    p = tmp_path / "x.csv"
    write_raw(p, "Регион;Сумма;Сумма;\n\nМосква;1;2;\nКазань;3;4;5;лишнее\n\n")
    df, enc, sep, notes = R.read_csv_any(str(p))
    assert sep == ";" and list(df.columns) == ["Регион", "Сумма", "Сумма (2)", "Колонка 4", "Колонка 5"]
    assert df.values.tolist() == [["Москва", "1", "2", "", ""], ["Казань", "3", "4", "5", "лишнее"]]
    assert any("лишними полями" in n for n in notes) and any("Сумма (2)" in n for n in notes)


def test_read_utf16_quotes_newlines_pipe(tmp_path):
    write_raw(tmp_path / "u.csv", "Регион\tСумма\nМосква\t1,5\n", encoding="utf-16")
    df, enc, sep, _ = R.read_csv_any(str(tmp_path / "u.csv"))
    assert enc == "utf-16" and sep == "\t" and df.iloc[0, 1] == "1,5"
    write_raw(tmp_path / "q.csv", 'Клиент,Комментарий,Сумма\n"ООО ""Р""","a\nb; c",1\n')
    df, _, sep, _ = R.read_csv_any(str(tmp_path / "q.csv"))
    assert sep == "," and df.iloc[0].tolist() == ['ООО "Р"', "a\nb; c", "1"]


def test_read_rejects_binary_and_empty(tmp_path):
    (tmp_path / "b.csv").write_bytes(bytes(range(256)))
    with pytest.raises(ValueError, match="не текстовый"):
        R.read_csv_any(str(tmp_path / "b.csv"))
    write_raw(tmp_path / "e.csv", "\ufeff\n\n")
    with pytest.raises(ValueError, match="пустой"):
        R.read_csv_any(str(tmp_path / "e.csv"))


def test_numbers_brackets_currency_inf():
    nums, _ = R.to_numbers(pd.Series(["(150)", "1 000 ₽", "$5", "inf", "1.5E+3"]))
    assert nums.iloc[:3].tolist() == [-150.0, 1000.0, 5.0] and pd.isna(nums.iloc[3]) and nums.iloc[4] == 1500


def test_dates_strict_no_day_month_swap():
    d, fmt = R.parse_dates(pd.Series(["2026-02-01", "2026-13-01", "2025-12-31"]))
    assert fmt == "%Y-%m-%d" and d.iloc[0].month == 2 and pd.isna(d.iloc[1])
    d, fmt = R.parse_dates(pd.Series(["01.02.2026", "31.12.2025 23:59"]))
    assert d.iloc[0].month == 2 and d.iloc[1].hour == 23


def test_bad_values_do_not_break_book(tmp_path):
    write_raw(tmp_path / "a.csv", "Имя;Значение;Сумма\n=1+1;@x;10\nА\x01Б;" + "Я" * 40000 + ";20\n")
    write_raw(tmp_path / "b.csv", "")
    (tmp_path / "c.csv").write_bytes(b"\x00\x01\x02")
    wb = build(tmp_path)
    ws = wb.worksheets[1]
    rows = [r for r in ws.iter_rows(min_row=8, values_only=False) if r[0].value is not None]
    assert rows[0][0].value == "=1+1" and rows[0][0].data_type == "s"         # не формула
    assert rows[1][0].value == "АБ" and len(rows[1][1].value) == R.EXCEL_MAX_TEXT
    toc = [r for r in wb[R.TOC_SHEET].iter_rows(values_only=True)]
    assert any(r[1] == "b.csv" for r in toc) and any(r[1] == "c.csv" for r in toc)   # пропущенные — в содержании


def test_hierarchy_without_name_column_and_level_gaps(tmp_path):
    write_raw(tmp_path / "h.csv", "Уровень;MO;Сумма 2026;Сумма 2025\n0;;10;9\n2;ГКБ 2;1;1\n1;ГКБ 1;6;5\n")
    ws = build(tmp_path).worksheets[1]
    names = [r[1] for r in ws.iter_rows(min_row=8, values_only=True) if r[0] is not None]
    assert names == ["Всего", "ГКБ 1", "ГКБ 2"]


def test_many_files_and_no_files(tmp_path):
    for k in range(40):
        write_raw(tmp_path / f"r_{k}.csv", "Регион;Сумма\nМосква;1\n")
    wb = build(tmp_path)
    assert len(wb.sheetnames) == 40 + 3                 # + «Содержание», «Справка по колонкам», скрытые «Формулы»
    empty = tmp_path / "empty"
    empty.mkdir()
    with pytest.raises(SystemExit):
        build(empty)


# ---------------------------------------------------------------- разнородные папки и широкие файлы
def test_single_column_decimal_vs_two_integer_columns(tmp_path):
    write_raw(tmp_path / "a.csv", "Оплата, руб\n55723,69\n46968,05\n98045,98\n")
    df, _, sep, _ = R.read_csv_any(str(tmp_path / "a.csv"))
    assert sep == R.SINGLE_COLUMN and df.columns.tolist() == ["Оплата, руб"] and df.iloc[0, 0] == "55723,69"
    write_raw(tmp_path / "b.csv", "Случаи 2026,Случаи 2025\n0,397\n310,427\n57,494\n")
    df, _, sep, _ = R.read_csv_any(str(tmp_path / "b.csv"))
    assert sep == "," and df.shape == (3, 2)
    df, _, sep, _ = R.read_csv_any(str(tmp_path / "a.csv"), sep=",")       # ручная настройка files.sep
    assert df.shape == (3, 2)


def test_blank_rows_skipped_with_note(tmp_path):
    write_raw(tmp_path / "x.csv", 'Отделение\n""\nХирургия\n""\n')
    df, _, _, notes = R.read_csv_any(str(tmp_path / "x.csv"))
    assert df.iloc[:, 0].tolist() == ["Хирургия"] and any("2 полностью пустых" in n for n in notes)


def test_totals_label_when_first_column_is_metric(tmp_path):
    write_raw(tmp_path / "x.csv", "Значение;Комментарий\n1;\n2;а\n")
    ws = build(tmp_path).worksheets[1]
    assert [ws.cell(7, c).value for c in (1, 2)] == [3, "Итого"]
    write_raw(tmp_path / "x.csv", "Случаи;Сумма\n1;2\n")               # одни метрики — подпись в колонке «№»
    ws = build(tmp_path).worksheets[1]
    assert ws.cell(7, 1).value == "Итого"


def mixed_folder(tmp_path):
    write_raw(tmp_path / "ref.csv", "Код;Название\n001;Альфа\n002;Бета\n")                       # без метрик
    write_raw(tmp_path / "sum.csv", "Регион;Случаи 2026;Случаи 2025\nМосква;5;4\n")              # уверенно
    head = ["Регион"] + [f"Показатель {i}" for i in range(60)] + [f"Сумма {i}" for i in range(10)]
    rows = [["Р1"] + [str(i * 7 + j) for j in range(60)] + ["1,5"] * 10 for i in range(6)]
    write_csv(tmp_path / "wide.csv", head, rows, encoding="utf-8")                              # 60 сомнительных


def test_review_only_doubtful_structures(tmp_path):
    mixed_folder(tmp_path)
    cfg = R.deep_merge(R.DEFAULTS, {})
    said = []
    it = iter(["", "", "", "",              # книга: заголовок, подзаголовок, период, рост
               "2",                         # 5-й вопрос: колонки — проверить в консоли
               "",                          # какие структуры проверять — только с сомнениями
               "",                          # wide.csv: показать все колонки — нет
               "*", "3",                    # все сомнительные -> «другое число»
               "!"])
    asker = R.Asker(inp=lambda p: next(it, ""), out=lambda *a: said.append(" ".join(map(str, a))))
    raw, _ = R.interview(str(tmp_path), cfg, {}, asker)
    text = "\n".join(said)
    assert "— Структуры файлов: 3 —" in text and "только с сомнениями (1 из 3)" in text
    assert "Структура 1 из 1: 1 файл(ов) — wide.csv" in text
    assert "колонок 71:" in text                                   # широкий файл — сводка, не 71 строка
    assert raw["columns"]["Показатель 0"] == {"role": "metric", "kind": "number"}     # «*» -> «другое число»
    assert raw["columns"]["Показатель 59"]["kind"] == "number"
    assert set(raw["files"]) == {"ref.csv", "sum.csv", "wide.csv"}  # имена листов — у всех


def test_random_folders_keep_every_row(tmp_path):
    rnd = __import__("random").Random(5)
    gens = [lambda: rnd.choice(["Москва", "", "a;b", 'ООО "Р"']), lambda: f"{rnd.randint(0, 999):05d}",
            lambda: str(rnd.randint(-5, 900)), lambda: f"{rnd.uniform(-9, 9999):.2f}".replace(".", ","),
            lambda: f"{rnd.randint(1, 28):02d}.03.2026", lambda: rnd.choice(["1", "н/д"]), lambda: ""]
    expected = {}
    for k in range(30):
        ncol = rnd.choice([1, 2, 3, 5, 9, 20, 45, 120])
        kinds = [rnd.randrange(len(gens)) for _ in range(ncol)]
        rows = [[gens[g]() for g in kinds] for _ in range(rnd.choice([1, 3, 25]))]
        rows = [r for r in rows if any(r)] or [["x"] + [""] * (ncol - 1)]
        write_csv(tmp_path / f"f{k}.csv", [f"К{i}" for i in range(ncol)], rows,
                  encoding=rnd.choice(["utf-8", "cp1251"]), sep=rnd.choice([";", "\t"]))
        expected[f"f{k}.csv"] = len(rows)
    wb = build(tmp_path)
    toc = {r[8]: r[1] for r in wb[R.TOC_SHEET].iter_rows(values_only=True) if r and len(r) > 8 and r[8]}
    assert set(toc) >= set(expected)
    for f, n in expected.items():
        ws = wb[toc[f]]
        assert ws.max_row >= 6 + n                                 # шапка + данные; ничего не потеряно


def test_arrow_strings_use_python_regex():
    """С pyarrow pandas хранит строки в Arrow и гоняет регулярки через RE2 — раньше это роняло чтение всех
    файлов («invalid escape sequence: \\u»)."""
    pytest.importorskip("pyarrow")
    s = pd.Series(["1 234,5", "(150)", "12,5%"], dtype="string[pyarrow]")
    nums, _ = R.to_numbers(s)
    assert nums.tolist()[:2] == [1234.5, -150.0]
    c = R.analyze_column("Сумма", pd.Series(["1 000", "2"], dtype="string[pyarrow]"), {})
    assert c.is_metric and c.values == [1000.0, 2.0]


def test_cli_asks_at_most_five_questions(tmp_path):
    mixed_folder(tmp_path)
    prompts = []

    def inp(p):
        prompts.append(p)
        return ""
    raw, changed = R.interview(str(tmp_path), R.deep_merge(R.DEFAULTS, {}), {}, R.Asker(inp=inp, out=lambda *a: None))
    assert changed and len(prompts) <= 5
    assert set(raw["files"]) == {"ref.csv", "sum.csv", "wide.csv"}          # листы назначены автоматически
    cfg = R.deep_merge(R.DEFAULTS, raw)
    assert R.plan_questions(str(tmp_path), cfg, raw)["count"] == 0          # второй запуск — без вопросов


def test_hierarchy_from_level_sorted_csv(tmp_path):
    """Выгрузка идёт по уровням (все группы, потом все отделения, потом все «из них») — родитель ищется
    по разрезам, «из них операции» встаёт под своё отделение."""
    head = ["Уровень", "Наименование", "Большая группа", "Группа отделений", "Из них", "Сумма 2026", "Сумма 2025"]
    rows = [[0, "Всего", "", "", "", 100, 90],
            [1, "Основные", "Основные", "", "", 50, 40], [1, "Центры", "Центры", "", "", 40, 30],
            [2, "ЖК", "Основные", "ЖК", "", 10, 9], [2, "КДО", "Основные", "КДО", "", 30, 20],
            [2, "МГГЦ", "Центры", "МГГЦ", "", 5, 5], [2, "Эндоскопия", "Центры", "Эндоскопия", "", 35, 25],
            [3, "из них операции", "Основные", "ЖК", "из них операции", 2, 1],
            [3, "из них операции", "Основные", "КДО", "из них операции", 4, 3]]
    write_csv(tmp_path / "r.csv", head, rows)
    ws = build(tmp_path).worksheets[1]
    got = [(r[1], r[2]) for r in ws.iter_rows(min_row=8, values_only=True) if r[0] is not None]
    assert got == [("Всего", 100), ("Основные", 50), ("КДО", 30), ("из них операции", 4), ("ЖК", 10),
                   ("из них операции", 2), ("Центры", 40), ("Эндоскопия", 35), ("МГГЦ", 5)]
    ws = build(tmp_path, table={"hierarchy_sort": "off"}).worksheets[1]       # без сортировки — порядок CSV,
    got = [r[1] for r in ws.iter_rows(min_row=8, values_only=True) if r[0] is not None]   # но под родителями
    assert got == ["Всего", "Основные", "ЖК", "из них операции", "КДО", "из них операции", "Центры", "МГГЦ",
                   "Эндоскопия"]


def test_default_sheet_names_drop_common_part():
    files = ["rep_app_stac_1_202610290117.csv", "rep_app_stac_1_1_202610290117.csv", "ref.csv"]
    assert R.default_sheet_name(files[0], files) == "Отчёт 1"
    assert R.default_sheet_name(files[1], files) == "Отчёт 1.1"
    assert R.default_sheet_name("ref.csv", files) == "ref"
    assert R.default_sheet_name("rep_x_2.csv", ["rep_x_2.csv", "sales_3.csv"]) == "2 rep x"


def test_code_columns_are_not_summed(tmp_path):
    """Справочник: ORD (1, 2, 0…) — код, не метрика: без «Итого 42» и с кодом 0, а не «-»."""
    write_raw(tmp_path / "ord.csv", "ORD;CASE_TYPE;ORD_NAME\n1;HOSP;плановая\n2;HOSP;экстренная\n0;AMB;самотек\n")
    ws = build(tmp_path).worksheets[1]
    assert [ws.cell(r, 1).value for r in (5, 6, 7, 8, 9)] == ["ORD", 1, 1, 2, 0]   # шапка, номера, данные
    assert isinstance(ws.cell(9, 1).value, int) and "\\-" not in ws.cell(9, 1).number_format   # 0, а не «-»
    assert R.is_code_name("ORD") and R.is_code_name("CASE_TYPE") and not R.is_code_name("Order amount 2026")


def test_streaming_writes_the_same_book(tmp_path):
    """Потоковая запись (большие книги) даёт ту же книгу, что обычная: значения, стили, объединения, закрепление,
    фильтр, ширины, примечания, группировка, ссылки."""
    write_csv(tmp_path / "rep_1_5.csv", HEAD, [[0, "Всего", "", "", "", 600, 60.0, 500, 50.0],
                                               [1, "Хирургия", "Хирургия", "", "", 100, 40.0, 90, 30.0],
                                               [1, "из них операции", "Хирургия", "", "операции", 10, 1.0, 10, 1.0],
                                               [2, "Хир 1", "Хирургия", "Хир 1", "", 50, 20.0, 40, 10.0]])
    write_raw(tmp_path / "sum.csv", "Регион;Случаи 2026;Случаи 2025;Код\nМосква;5;4;001\nКазань;3;0;002\n")
    books = {}
    for stream in (False, True):
        out = str(tmp_path / f"s{stream}.xlsx")
        R.build(str(tmp_path), R.deep_merge(R.DEFAULTS, {"book": {"stream": stream}}), out)
        books[stream] = load_workbook(out)
    a, b = books[False], books[True]
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        x, y = a[name], b[name]
        assert (x.freeze_panes, x.auto_filter.ref, x.sheet_state) == (y.freeze_panes, y.auto_filter.ref, y.sheet_state)
        assert set(map(str, x.merged_cells.ranges)) == set(map(str, y.merged_cells.ranges))
        for r in range(1, x.max_row + 1):
            assert (x.row_dimensions[r].outline_level or 0) == (y.row_dimensions[r].outline_level or 0)
            for c in range(1, x.max_column + 1):
                p, q = x.cell(r, c), y.cell(r, c)
                assert (p.value, p.number_format, p.font.b, p.fill.fgColor.rgb, p.border.left.style,
                        bool(p.comment), bool(p.hyperlink)) == \
                       (q.value, q.number_format, q.font.b, q.fill.fgColor.rgb, q.border.left.style,
                        bool(q.comment), bool(q.hyperlink)), (name, p.coordinate)
