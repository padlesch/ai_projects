"""Отчёт-мейкер (веб): анализ папки -> правки (по одной и группой) -> предпросмотр -> сборка; HTTP-сервер."""
import json
import os
import sys
import threading
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

import pytest
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import report_app as A  # noqa: E402
import report_builder as R  # noqa: E402
from test_report_builder import HEAD, mixed_folder, write_csv, write_raw  # noqa: E402


@pytest.fixture
def folder(tmp_path):
    d = tmp_path / "reps"
    d.mkdir()
    mixed_folder(d)
    write_csv(d / "rep_1_5.csv", HEAD, [[0, "Всего", "", "", "", 600, 60.0, 500, 50.0],
                                        [1, "Хирургия", "Хирургия", "", "", 100, 40.0, 90, 30.0],
                                        [1, "Терапия", "Терапия", "", "", 50, 50.0, 40, 30.0]])
    (d / "bin.csv").write_bytes(bytes(range(256)))
    return d


def page_state(folder):
    """Состояние, как его держит страница после анализа (с исходными значениями для «сбросить»)."""
    a = A.analyze(str(folder))
    a["by_col"] = {c["name"]: c for c in a["columns"]}
    a["by_file"] = {f["file"]: f for f in a["files"]}
    return a


def payload(a, **book):
    b = dict(a["book"], **book)
    return {"folder": a["folder"], "out": a["folder"] + "_out.xlsx", "book": b,
            "columns": a["columns"], "files": a["files"]}


def test_analyze_columns_across_folder(folder):
    a = page_state(folder)
    assert a["n_files"] == 5 and a["n_structures"] == 4 and [b["file"] for b in a["broken"]] == ["bin.csv"]
    assert a["by_col"]["Регион"]["n_files"] == 2                      # sum.csv и wide.csv — одна строка таблицы
    assert a["by_col"]["Уровень"]["kind"] == "level" and a["by_col"]["Наименование"]["kind"] == "name"
    assert a["by_col"]["Δ Случаи"]["formula"] and a["by_col"]["Δ Случаи"]["highlight"]
    assert a["columns"][0]["doubt"]                                  # с пометками — первыми
    f = a["by_file"]["rep_1_5.csv"]
    assert f["hierarchy"] and f["key_auto"] == "Сумма 2026, млн руб" and not f["totals_applicable"]
    assert a["by_file"]["ref.csv"]["sheet"] == "ref"


def test_bulk_edit_save_build_and_reset(folder):
    a = page_state(folder)
    for i in range(60):                                              # «группой»: 60 колонок -> другое число
        c = a["by_col"][f"Показатель {i}"]
        c.update(kind="number", set_kind=True)
    a["by_col"]["Регион"].update(label="Субъект РФ", set_label=True)
    a["by_col"]["Δ% Случаи"].update(highlight=False, set_hl=True)
    s = a["by_file"]["sum.csv"]
    s.update(sheet="Сводная", set_sheet=True, totals=False, subtitle="только Москва", description="Свод")
    a["by_file"]["ref.csv"]["skip"] = True
    res = A.run_build(payload(a, title="АПП", period="9м", positive_is="good"))
    assert res["ok"], res["log"]
    raw = R.load_raw(str(folder / "report.yaml"))
    assert raw["columns"]["Показатель 7"] == {"role": "metric", "kind": "number"}
    assert raw["columns"]["Регион"] == {"label": "Субъект РФ"} and raw["columns"]["Δ% Случаи"] == {"colorize": False}
    assert raw["files"]["sum.csv"] == {"sheet": "Сводная", "description": "Свод", "subtitle": "только Москва",
                                       "table": {"totals": "off"}}
    assert raw["files"]["ref.csv"] == {"skip": True}
    assert "wide.csv" not in raw["files"]                           # не трогали — в конфиге ничего
    wb = load_workbook(res["out"])
    assert "ref" not in wb.sheetnames and "Сводная" in wb.sheetnames
    ws = wb["Сводная"]
    assert ws.cell(3, 1).value == "только Москва" and ws.cell(5, 1).value == "Субъект РФ"
    assert ws.cell(7, 1).value != "Итого"
    # повторный анализ видит сохранённое; «сбросить к авто» убирает правила
    b = page_state(folder)
    assert b["by_col"]["Показатель 7"]["kind"] == "number" and b["by_col"]["Показатель 7"]["set_kind"]
    assert b["by_file"]["ref.csv"]["skip"] and b["by_file"]["sum.csv"]["sheet"] == "Сводная"
    for i in range(60):
        b["by_col"][f"Показатель {i}"].update(set_kind=False)
    b["by_file"]["ref.csv"]["skip"] = False
    assert A.run_build(payload(b))["ok"]
    raw = R.load_raw(str(folder / "report.yaml"))
    assert "Показатель 7" not in raw["columns"] and "ref.csv" not in raw["files"]


def test_manual_level_flag_turns_on_hierarchy(tmp_path):
    write_raw(tmp_path / "x.csv", "Глуб;Группа;Сумма 2026;Сумма 2025\n0;Всего;10;9\n1;А;4;3\n1;Б;6;6\n")
    a = page_state(tmp_path)
    assert not a["by_file"]["x.csv"]["hierarchy"]
    a["by_col"]["Глуб"].update(kind="level", set_kind=True)
    A.run_build(payload(a))
    b = page_state(tmp_path)
    assert b["by_file"]["x.csv"]["hierarchy"] and b["by_col"]["Группа"]["kind"] == "name"


@pytest.mark.parametrize("name", ["Ур", "Level", "lvl", "Depth", "Уровень иерархии", "level_1"])
def test_level_synonyms(tmp_path, name):
    write_raw(tmp_path / "x.csv", f"{name};Наименование;Сумма\n0;Всего;10\n1;А;10\n")
    assert page_state(tmp_path)["by_file"]["x.csv"]["hierarchy"]


def test_main_metric_does_not_reorder_flat_rows(tmp_path):
    write_raw(tmp_path / "x.csv", "Регион;Случаи;Сумма\nА;1;30\nБ;2;10\nВ;3;20\n")
    out = str(tmp_path) + ".xlsx"
    R.build(str(tmp_path), R.deep_merge(R.DEFAULTS, {"files": {"x.csv": {"table": {"key": "Случаи"}}}}), out)
    ws = load_workbook(out).worksheets[1]
    assert [ws.cell(r, 1).value for r in (8, 9, 10)] == ["А", "Б", "В"]


def test_preview_uses_unsaved_settings(folder):
    a = page_state(folder)
    a["by_col"]["Регион"].update(label="Субъект", set_label=True)
    p = A.preview(dict(payload(a), file="sum.csv"))
    texts = [c["v"] for row in p["rows"] for c in row]
    assert "Субъект" in texts and "Итого" in texts and p["total"] == 1
    assert not (folder / "report.yaml").exists()                    # предпросмотр ничего не сохраняет
    h = A.preview(dict(payload(a), file="rep_1_5.csv"))
    first = [c for c in h["rows"][3] if c["v"]]                      # строка «Всего»: заливка уровня 0, жирный
    assert first[1]["v"] == "Всего" and first[1]["b"] and first[1]["bg"] == "#C9DCEC"


def test_fmt_value():
    assert A.fmt_value(1234.5, R.FMT["money"]) == "1 234,50"
    assert A.fmt_value(0.125, R.FMT["delta_pct"]) == "+12,5%"
    assert A.fmt_value(3.0, R.FMT_PCT_CAP) == "> 200%" and A.fmt_value(0, R.FMT["count"]) == "-"


def test_http_server_and_origin_check(folder):
    server = ThreadingHTTPServer(("127.0.0.1", 0), A.Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{server.server_address[1]}"
    try:
        page = urllib.request.urlopen(base + "/").read().decode("utf-8")
        assert "Отчёт-мейкер" in page
        req = urllib.request.Request(base + "/api/analyze", data=json.dumps({"folder": str(folder)}).encode(),
                                     headers={"Content-Type": "application/json"})
        assert len(json.loads(urllib.request.urlopen(req).read())["files"]) == 4
        evil = urllib.request.Request(base + "/api/analyze", data=b"{}", headers={"Origin": "http://evil.example"})
        with pytest.raises(urllib.error.HTTPError) as e:
            urllib.request.urlopen(evil)
        assert e.value.code == 403
        bad = urllib.request.Request(base + "/api/analyze", data=json.dumps({"folder": "/нет/такой"}).encode())
        with pytest.raises(urllib.error.HTTPError) as e:
            urllib.request.urlopen(bad)
        assert e.value.code == 400 and "Нет папки" in json.loads(e.value.read())["error"]
    finally:
        server.shutdown()


def test_pick_folder_cyrillic_path_and_errors(monkeypatch):
    """На Windows ответ окна выбора идёт в кодировке консоли (cp1251) — путь с кириллицей раньше терялся,
    и кнопка «Выбрать…» молча ничего не делала."""
    monkeypatch.setenv("PYTHONIOENCODING", "cp1251")
    path = "C:\\\\Users\\\\comp\\\\обучение\\\\апп стац\\\\reps"
    ok = A.pick_folder(f'import json; print(json.dumps({{"folder": "{path}"}}))')
    assert ok == {"folder": "C:\\Users\\comp\\обучение\\апп стац\\reps", "error": ""}
    assert "no display" in A.pick_folder('raise RuntimeError("no display")')["error"]
    assert A.pick_folder('import json; print(json.dumps({"folder": ""}))') == {"folder": "", "error": ""}


# ---------------------------------------------------------------- правки для части файлов, шаблоны, прогресс
def test_column_settings_for_some_files_only(tmp_path):
    for name in ("rep_19_1_202610290117.csv", "rep_19_2_202610290117.csv", "rep_20_1_202610290117.csv"):
        write_raw(tmp_path / name, "Регион;Показатель\nМосква;5\nКазань;7\n")
    a = page_state(tmp_path)
    a["by_col"]["Регион"].update(label="Субъект", set_label=True)                 # для всех файлов
    for f in a["files"]:
        if f["group"] == "19":                                                     # только отчёты 19.*
            f["colrules"]["Показатель"] = {"kind": "money", "label": "Показатель", "highlight": False,
                                           "set_kind": True, "set_label": False, "set_hl": False}
    res = A.run_build(payload(a))
    assert res["ok"], res["log"]
    raw = R.load_raw(str(tmp_path / "report.yaml"))
    assert raw["columns"] == {"Регион": {"label": "Субъект"}}
    assert raw["files"]["rep_19_1_202610290117.csv"]["columns"] == {"Показатель": {"role": "metric", "kind": "money"}}
    assert "rep_20_1_202610290117.csv" not in raw["files"]
    wb = load_workbook(res["out"])
    w19, w20 = wb["Отчёт 19.1"], wb["Отчёт 20.1"]
    assert w19.cell(5, 1).value == w20.cell(5, 1).value == "Субъект"               # общая подпись — везде
    assert w19.cell(9, 2).number_format == R.FMT["money"] and w20.cell(9, 2).number_format != R.FMT["money"]
    b = page_state(tmp_path)                                                       # повторный анализ видит своё
    f19 = next(f for f in b["files"] if f["file"].startswith("rep_19_1"))
    assert f19["colrules"]["Показатель"]["kind"] == "money" and f19["colrules"]["Показатель"]["set_kind"]


def test_template_save_and_apply_to_next_period(tmp_path, monkeypatch):
    monkeypatch.setenv("REPORT_MAKER_HOME", str(tmp_path / "home"))
    oct_ = tmp_path / "oct"
    nov = tmp_path / "nov"
    oct_.mkdir()
    nov.mkdir()
    write_raw(oct_ / "rep_1_1_202610290117.csv", "Регион;Сумма\nМосква;5\n")
    write_raw(nov / "rep_1_1_202611290900.csv", "Регион;Сумма\nКазань;9\n")       # другой день выгрузки
    a = page_state(oct_)
    a["by_col"]["Регион"].update(label="Субъект", set_label=True)
    a["by_file"]["rep_1_1_202610290117.csv"].update(sheet="Свод", set_sheet=True, description="Свод по регионам")
    saved = A.save_template(dict(payload(a, title="АПП"), name="АПП / стационары"))
    assert saved["name"] == "АПП _ стационары" and saved["templates"] == ["АПП _ стационары"]
    cfg = A.load_template("АПП _ стационары")
    assert list(cfg["files"]) == [r"re:^rep_1_1_\d{8,}\.csv$"] and "output" not in cfg.get("book", {})
    b = A.analyze(str(nov), template="АПП _ стационары")                           # применить к ноябрю
    f = b["files"][0]
    assert b["book"]["title"] == "АПП" and f["sheet"] == "Свод" and f["description"] == "Свод по регионам"
    assert next(c for c in b["columns"] if c["name"] == "Регион")["label"] == "Субъект"
    res = A.run_build(dict(payload(b), template="АПП _ стационары"))
    assert res["ok"] and load_workbook(res["out"])["Свод"].cell(5, 1).value == "Субъект"
    assert (nov / "report.yaml").exists()                                           # сохранено в папку ноября
    with pytest.raises(ValueError):
        A.load_template("нет такого")


def test_background_build_reports_progress(folder):
    import time
    a = page_state(folder)
    job = A.start_build(payload(a))["job"]
    for _ in range(200):
        st = A.job_status(job)
        if not st["running"]:
            break
        time.sleep(0.05)
    assert st["result"]["ok"] and st["total"] == 5 and st["done"] == 5


def test_sum_check_and_outline_in_hierarchy(tmp_path):
    head = ["Уровень", "Наименование", "Группа", "Сумма 2026", "Пациенты 2026"]
    write_csv(tmp_path / "h.csv", head, [[0, "Всего", "", "10,00", 10], [1, "А", "А", "4,00", 6],
                                         [1, "Б", "Б", "5,00", 6]])                 # 4 + 5 ≠ 10
    out = str(tmp_path) + ".xlsx"
    R.build(str(tmp_path), R.deep_merge(R.DEFAULTS, {}), out)
    wb = load_workbook(out)
    ws = wb.worksheets[1]
    notes = [r[9] for r in wb[R.TOC_SHEET].iter_rows(values_only=True) if r and len(r) > 9 and r[9]]
    assert any("сверка сумм: 1 расхожд." in n and "Сумма 2026" in n for n in notes)  # пациенты не сверяются
    assert ws.cell(7, 2).value == "Всего" and ws.cell(7, 3).font.color.rgb.endswith("9E3423")   # выделена
    assert not ws.cell(7, 4).font.color.rgb.endswith("9E3423")                    # пациенты не сверяются
    assert [ws.row_dimensions[r].outline_level for r in (7, 8, 9)] == [0, 1, 1]
    assert ws.sheet_properties.outlinePr.summaryBelow is False


def test_picker_command_in_frozen_exe(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    assert A.picker_command("code") == [sys.executable, "--pick-folder"]
    monkeypatch.setattr(sys, "frozen", False)
    assert A.picker_command("code")[1:] == ["-c", "code"]


def test_template_mask_does_not_leak_between_reports():
    """Маска отчёта 1 не должна подходить к отчёту 1.1 (иначе его имя листа протекло бы в 1.1)."""
    m1 = A.generalize("rep_app_stac_1_202610290117.csv")
    assert R.pattern_match(m1, "rep_app_stac_1_202611300800.csv")
    assert not R.pattern_match(m1, "rep_app_stac_1_1_202611300800.csv")
    assert A.generalize("ref.csv") == "ref.csv"
