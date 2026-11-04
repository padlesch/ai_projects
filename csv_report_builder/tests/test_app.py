"""Веб-приложение: анализ -> настройки -> сборка (без браузера) и HTTP-сервер."""
import json
import os
import sys
import threading
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


def test_analyze_groups_types_and_reports_broken(folder):
    a = A.analyze(str(folder))
    assert a["n_files"] == 5 and len(a["types"]) == 4 and [b["file"] for b in a["broken"]] == ["bin.csv"]
    assert a["types"][0]["doubt"]                                   # с сомнениями — первыми
    wide = next(t for t in a["types"] if t["files"][0]["file"] == "wide.csv")
    assert len(wide["columns"]) == 71 and any(c["doubt"] for c in wide["columns"])
    hier = next(t for t in a["types"] if t["hierarchy"])
    assert hier["columns"][0]["fixed"] == "наименование" and not hier["totals_applicable"]
    assert [c["name"] for c in hier["columns"] if c["main"]] == ["Сумма 2026, млн руб"]


def test_settings_roundtrip_and_build(folder):
    a = A.analyze(str(folder))
    for t in a["types"]:
        for c in t["columns"]:
            c["kind_orig"], c["label_orig"] = c["kind"], c["label"]
    s = next(t for t in a["types"] if t["files"][0]["file"] == "sum.csv")
    s["columns"][0]["label"] = "Субъект РФ"
    s["totals"] = False
    for c in s["columns"]:
        c["highlight"] = c["name"] == "Δ% Случаи"
    s["files"][0]["sheet"] = "Сводная"
    w = next(t for t in a["types"] if t["files"][0]["file"] == "wide.csv")
    w["columns"][1]["kind"] = "number"                              # «Показатель 0» -> другое число
    payload = {"folder": a["folder"], "out": str(folder) + "_out.xlsx", "types": a["types"],
               "book": {"title": "АПП", "subtitle": "ЗЛ", "period": "9м", "positive_is": "good", "palette": "B"}}
    res = A.run_build(payload)
    assert res["ok"], res["log"]
    raw = R.load_raw(str(folder / "report.yaml"))
    assert raw["book"]["title"] == "АПП" and raw["coloring"]["positive_is"] == "good"
    assert raw["columns"]["Регион"]["label"] == "Субъект РФ"
    assert raw["columns"]["Показатель 0"] == {"role": "metric", "kind": "number"}
    assert "Показатель 1" not in raw["columns"]                      # не трогали — решает автоматика
    assert raw["files"]["sum.csv"]["coloring"] == {"columns": ["Δ% Случаи"]}
    assert raw["files"]["sum.csv"]["table"]["totals"] == "off"
    assert "&id" not in (folder / "report.yaml").read_text(encoding="utf-8")   # без YAML-ссылок
    ws = load_workbook(res["out"])["Сводная"]
    assert ws.cell(5, 1).value == "Субъект РФ" and ws.cell(7, 1).value != "Итого"
    again = A.analyze(str(folder))                                   # повторный анализ видит сохранённое
    s2 = next(t for t in again["types"] if t["files"][0]["file"] == "sum.csv")
    assert s2["files"][0]["sheet"] == "Сводная" and not s2["totals"]
    assert [c["name"] for c in s2["columns"] if c["highlight"]] == ["Δ% Случаи"]


def test_main_metric_does_not_reorder_flat_rows(tmp_path):
    write_raw(tmp_path / "x.csv", "Регион;Случаи;Сумма\nА;1;30\nБ;2;10\nВ;3;20\n")
    ws = load_workbook_built(tmp_path, {"files": {"x.csv": {"table": {"key": "Случаи"}}}})
    assert [ws.cell(r, 1).value for r in (8, 9, 10)] == ["А", "Б", "В"]


def load_workbook_built(folder, extra):
    out = str(folder) + ".xlsx"
    R.build(str(folder), R.deep_merge(R.DEFAULTS, extra), out)
    return load_workbook(out).worksheets[1]


def test_http_server_and_origin_check(folder):
    server = ThreadingHTTPServer(("127.0.0.1", 0), A.Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{server.server_address[1]}"
    try:
        page = urllib.request.urlopen(base + "/").read().decode("utf-8")
        assert "Отчёты из CSV" in page and "%KINDS%" not in page
        req = urllib.request.Request(base + "/api/analyze", data=json.dumps({"folder": str(folder)}).encode(),
                                     headers={"Content-Type": "application/json"})
        assert len(json.loads(urllib.request.urlopen(req).read())["types"]) == 4
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
