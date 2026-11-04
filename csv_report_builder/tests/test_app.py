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
