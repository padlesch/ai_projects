"""
Веб-приложение: Excel-отчёты из папки CSV. Локально, без интернета: сервер слушает только 127.0.0.1,
CSV никуда не отправляются. Движок — report_builder.py (тот же, что в консоли).

  python report_app.py                  -> откроет браузер на http://127.0.0.1:8765
  python report_app.py --port 9000 --no-browser

Шаги на странице: папка -> 5 настроек книги -> типы отчётов (однотипные файлы — одна карточка; всё уже
распознано, правится только то, что не нравится) -> «Собрать Excel». Настройки сохраняются в report.yaml
в папке с CSV (прежний — report.yaml.bak), консольный запуск их тоже подхватывает.
"""
import argparse
import contextlib
import io
import json
import os
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import report_builder as R  # noqa: E402

UI_KINDS = [k for k, _ in R.KIND_CHOICES]
KIND_LABELS = {"count": "штуки", "money": "деньги", "number": "другое число", "pct": "процент / доля",
               "delta": "Δ изменение", "delta_pct": "Δ% изменение", "dimension": "разрез / текст",
               "skip": "не выводить"}
FIXED_ROLES = {"level": "уровень иерархии", "name": "наименование", "subrow": "«из них»"}
STATE = {"last_out": None, "last_folder": None}


# ---------------------------------------------------------------- анализ и сборка (без HTTP — тестируются)
def config_path_for(folder):
    return next((os.path.join(folder, n) for n in R.CONFIG_NAMES if os.path.exists(os.path.join(folder, n))), None)


def ui_kind(c):
    if c.is_metric:
        return "pct" if c.kind == "pct_points" else c.kind
    return "dimension"


def analyze(folder):
    """Папка -> всё, что нужно странице: настройки книги и типы отчётов с распознанными колонками."""
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise ValueError(f"Нет папки {folder}")
    cpath = config_path_for(folder)
    cfg, raw = R.load_config(cpath), R.load_raw(cpath)
    plan = R.plan_questions(folder, cfg, raw)
    if not plan["reports"] and not plan["broken"]:
        raise ValueError("В папке нет CSV-файлов")
    files_raw = raw.get("files") or {}
    all_files = [f for f, _ in plan["reports"]]
    used = set()
    types = []
    for items in plan["structures"]:
        f0, rep = items[0]
        fc0 = R.file_config(cfg, f0)
        hl = {id(c) for c in R.colored_columns(rep, fc0["coloring"])} \
            if str(fc0["coloring"].get("mode", "static")).lower() != "off" else set()
        main, _ = R.main_metric(rep.metrics, rep.table)
        unsure, tech = R.doubts(rep, raw.get("columns") or {})
        cols = []
        for c in rep.columns:
            fixed = FIXED_ROLES.get(c.role) or ("служебная" if c.role == "service" and not c.formula else None)
            cols.append({
                "name": c.name, "label": c.label, "kind": ui_kind(c), "fixed": fixed if c.role in FIXED_ROLES else None,
                "editable": not c.formula and c.role not in FIXED_ROLES,
                "metric": c.is_metric, "formula": bool(c.formula), "sumable": c.is_metric and c.kind in
                ("count", "money", "number"), "highlight": id(c) in hl, "main": c is main,
                "sample": [str(v) for v in c.sample[:3]],
                "flags": [x for x in R.column_flags(c) if x != "добавлена скриптом"] +
                         (["техническое имя — подпишите"] if c in tech else []),
                "doubt": c in unsure or c in tech,
            })
        files = []
        for f, r in items:
            entry = files_raw.get(f) or {}
            sheet = R.sheet_title(entry.get("sheet") or R.file_config(cfg, f).get("sheet") or
                                  R.default_sheet_name(f, all_files), used)
            files.append({"file": f, "sheet": sheet, "description": entry.get("description", ""),
                          "rows": r.n_rows, "notes": r.notes})
        types.append({
            "id": len(types), "hierarchy": rep.hierarchy, "columns": cols, "files": files,
            "n_source": len(rep.source_columns), "doubt": any(c["doubt"] for c in cols),
            "totals_applicable": not rep.hierarchy and any(c.additive for c in rep.metrics),
            "totals": str(fc0["table"].get("totals", "auto")).lower() not in ("off", "false", "0"),
        })
    types.sort(key=lambda t: (not t["doubt"], -len(t["files"])))
    book = raw.get("book") or {}
    return {
        "folder": folder, "config": cpath,
        "out": (book.get("output") or folder.rstrip("\\/") + ".xlsx"),
        "book": {"title": book.get("title") or os.path.basename(folder.rstrip("\\/")),
                 "subtitle": book.get("subtitle") or "", "period": book.get("period") or "",
                 "positive_is": (raw.get("coloring") or {}).get("positive_is", "bad"),
                 "palette": str((raw.get("style") or {}).get("palette", "B")).upper()},
        "has_years": any(c.year for _, r in plan["reports"] for c in r.metrics),
        "types": types, "broken": [{"file": f, "reason": e} for f, e in plan["broken"]],
        "n_files": len(plan["reports"]) + len(plan["broken"]),
    }


def settings_to_config(folder, payload, raw):
    """Ответы страницы -> report.yaml. Вид и подпись колонки пишутся, только если их поменяли (иначе решает
    автоматика); подсветка, главная метрика, «Итого», имя листа — явно для каждого файла: что на экране, то и в файле."""
    raw = json.loads(json.dumps(raw or {}))
    book = raw.setdefault("book", {})
    b = payload.get("book") or {}
    for key in ("title", "subtitle", "period"):
        book[key] = str(b.get(key, "")).strip()
    raw.setdefault("coloring", {})["positive_is"] = "good" if b.get("positive_is") == "good" else "bad"
    raw.setdefault("style", {})["palette"] = "A" if str(b.get("palette")).upper() == "A" else "B"
    columns = raw.setdefault("columns", {})
    files = raw.setdefault("files", {})
    for t in payload.get("types") or []:
        cols = t.get("columns") or []
        for c in cols:
            if c.get("editable") and c.get("kind") != c.get("kind_orig") and c.get("kind") in UI_KINDS:
                rule = columns.setdefault(c["name"], {})
                for key in ("kind", "role"):
                    rule.pop(key, None)
                if c["kind"] in ("dimension", "skip"):
                    rule["role"] = c["kind"]
                else:
                    rule.update({"role": "metric", "kind": c["kind"]})
            label = str(c.get("label", "")).strip()
            if label and label != c.get("label_orig"):
                columns.setdefault(c["name"], {})["label"] = label
        hl = [c["name"] for c in cols if c.get("highlight") and c.get("kind") != "skip"]
        main = next((c["name"] for c in cols if c.get("main")), None)
        for f in t.get("files") or []:
            entry = files.setdefault(f["file"], {})
            entry["sheet"] = str(f.get("sheet") or "").strip()[:31] or entry.get("sheet")
            desc = str(f.get("description") or "").strip()
            if desc:
                entry["description"] = desc
            else:
                entry.pop("description", None)
            entry["coloring"] = {"columns": hl} if hl else {"mode": "off"}
            table = entry.setdefault("table", {})
            if main:
                table["key"] = main
            if t.get("totals_applicable"):
                table["totals"] = "auto" if t.get("totals") else "off"
            if not table:
                entry.pop("table")
    return raw


def run_build(payload):
    """Сохранить настройки и собрать книгу. -> {ok, out, log}."""
    folder = os.path.abspath(payload["folder"])
    cpath = config_path_for(folder)
    raw = settings_to_config(folder, payload, R.load_raw(cpath))
    log = io.StringIO()
    with contextlib.redirect_stdout(log):
        cpath = R.save_answers(raw, cpath or os.path.join(folder, "report.yaml"))
    cfg = R.load_config(cpath)
    out = os.path.abspath(str(payload.get("out") or "").strip() or folder.rstrip("\\/") + ".xlsx")
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"
    if os.path.exists(out):
        try:
            os.rename(out, out)            # открыт в Excel -> на Windows ошибка
        except OSError:
            return {"ok": False, "out": out, "log": log.getvalue() +
                    f"Файл {out} открыт в Excel — закройте его или укажите другое имя.\n"}
    ok = True
    with contextlib.redirect_stdout(log):
        try:
            R.build(folder, cfg, out)
        except SystemExit as err:          # движок сообщает о фатальных ситуациях через sys.exit
            ok = False
            print(err)
        except Exception as err:           # чтобы страница показала причину, а не «упала»
            ok = False
            print(f"Ошибка: {type(err).__name__}: {err}")
    ok = ok and os.path.exists(out)
    if ok:
        STATE["last_out"], STATE["last_folder"] = out, folder
    return {"ok": ok, "out": out, "log": log.getvalue()}


def pick_folder():
    """Системный диалог выбора папки (tkinter в отдельном процессе — не мешает серверу)."""
    code = ("import tkinter as t, tkinter.filedialog as d; r=t.Tk(); r.withdraw(); r.attributes('-topmost', 1); "
            "print(d.askdirectory(title='Папка с CSV') or '')")
    try:
        res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=600,
                             encoding="utf-8")
        return res.stdout.strip()
    except Exception:
        return ""


def open_path(path):
    if sys.platform.startswith("win"):
        os.startfile(path)                 # noqa: S606 — локальное приложение, открывает только свою книгу
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):          # без шума в консоли
        pass

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            body = PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_json({"error": "нет такой страницы"}, 404)

    def do_POST(self):
        # только запросы со своей страницы: другой сайт в браузере не сможет дёрнуть локальный сервер
        origin = self.headers.get("Origin")
        if origin and origin != f"http://{self.headers.get('Host')}":
            return self.send_json({"error": "чужой источник запроса"}, 403)
        try:
            payload = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
            if self.path == "/api/analyze":
                return self.send_json(analyze(payload.get("folder", "")))
            if self.path == "/api/build":
                return self.send_json(run_build(payload))
            if self.path == "/api/pick":
                return self.send_json({"folder": pick_folder()})
            if self.path == "/api/open":
                target = STATE["last_out"] if payload.get("what") == "book" else STATE["last_folder"]
                if not target:
                    return self.send_json({"error": "книга ещё не собрана"}, 400)
                open_path(target)
                return self.send_json({"ok": True})
            return self.send_json({"error": "нет такого действия"}, 404)
        except ValueError as err:
            return self.send_json({"error": str(err)}, 400)
        except Exception as err:
            return self.send_json({"error": f"{type(err).__name__}: {err}"}, 500)


def serve(port=8765, browser=True):
    for p in range(port, port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            break
        except OSError:
            continue
    else:
        sys.exit(f"Порты {port}–{port + 19} заняты")
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"Отчёты из CSV: {url}  (остановить — Ctrl+C или закрыть это окно)")
    if browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


# ---------------------------------------------------------------- страница
PAGE = r"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Отчёты из CSV</title>
<style>
:root{--ink:#1F1F1F;--title:#1B3A5C;--sub:#2D6A73;--muted:#555F69;--line:#D3D9DF;--frame:#9AA5B1;--head:#E8ECF0;
--plate:#F5F7F9;--count:#DCE8F3;--money:#DCEAD3;--number:#D6ECEC;--delta:#FBF3E2;--warn:#7A5600;--warnbg:#FBEFD0;
--bad:#9E3423;--badbg:#F6DAD0;--ok:#2F6A30;--okbg:#E1EEDC}
*{box-sizing:border-box}
body{margin:0;background:#fff;color:var(--ink);font:15px/1.45 "Segoe UI",Roboto,Arial,sans-serif}
main{max-width:1200px;margin:0 auto;padding:24px 16px 120px}
h1{margin:0 0 4px;font:700 28px/1.2 Arial,sans-serif;color:var(--title)}
.lead{margin:0 0 24px;color:var(--muted)}
h2{margin:0 0 12px;font:700 19px/1.3 Arial,sans-serif;color:var(--title)}
section.card{border:1px solid var(--line);border-top:3px solid var(--title);background:#fff;padding:18px 20px;margin:0 0 18px}
label{display:block;font-weight:600;font-size:13px;color:var(--muted);margin:0 0 4px}
input[type=text]{width:100%;padding:8px 10px;border:1px solid var(--frame);font:inherit;color:var(--ink);background:#fff}
input[type=text]:focus,select:focus,button:focus-visible{outline:2px solid var(--sub);outline-offset:1px}
select{padding:6px 8px;border:1px solid var(--frame);font:inherit;background:#fff;max-width:100%}
.row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
.row>.grow{flex:1 1 320px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}
button{font:600 15px/1 "Segoe UI",Arial,sans-serif;padding:10px 16px;min-height:40px;border:1px solid var(--title);
background:#fff;color:var(--title);cursor:pointer}
button.primary{background:var(--title);color:#fff}
button:disabled{opacity:.5;cursor:default}
.seg{display:flex;flex-wrap:wrap;gap:0;border:1px solid var(--frame);width:max-content;max-width:100%}
.seg label{margin:0;padding:8px 12px;font-weight:500;color:var(--ink);cursor:pointer;border-right:1px solid var(--line)}
.seg label:last-child{border-right:0}
.seg input{margin:0 6px 0 0}
.summary{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 14px}
.pill{padding:4px 10px;background:var(--plate);border:1px solid var(--line);font-size:13px}
.pill.warn{background:var(--warnbg);color:var(--warn);border-color:#E8D49A}
.pill.bad{background:var(--badbg);color:var(--bad);border-color:#E9B7A6}
details.type{border:1px solid var(--line);margin:0 0 12px;background:#fff}
details.type>summary{list-style:none;cursor:pointer;padding:12px 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;background:var(--plate)}
details.type>summary::-webkit-details-marker{display:none}
details.type>summary::before{content:"▸";color:var(--title)}
details.type[open]>summary::before{content:"▾"}
.tname{font-weight:700;color:var(--title)}
.tmeta{color:var(--muted);font-size:13px}
.tbody{padding:12px 16px 16px}
.tools{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin:0 0 10px;font-size:14px}
.tools label{display:inline;font-weight:500;color:var(--ink);font-size:14px}
.scroll{overflow-x:auto;border:1px solid var(--frame)}
table{border-collapse:collapse;width:100%;font-size:14px}
th{background:var(--head);color:var(--title);text-align:left;padding:7px 8px;border-bottom:1px solid var(--frame);white-space:nowrap}
td{padding:5px 8px;border-top:1px solid var(--line);vertical-align:middle}
td.num{color:#66707A;width:36px;text-align:right}
td.kind{border-left:6px solid transparent}
td.k-count{border-left-color:var(--count)}td.k-money{border-left-color:var(--money)}
td.k-number,td.k-pct{border-left-color:var(--number)}td.k-delta,td.k-delta_pct{border-left-color:var(--delta)}
td.k-dimension{border-left-color:var(--head)}td.k-skip{border-left-color:#fff}
tr.doubt td{background:#FFFBF0}
.flag{color:var(--warn);font-size:13px}
.sample{color:var(--muted);font-size:13px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.label-in{min-width:140px}
.files td input{min-width:160px}
.note{font-size:12px;color:var(--bad)}
.hint{font-size:13px;color:var(--muted);margin:8px 0 0}
.bar{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid var(--frame);padding:12px 16px}
.bar .in{max-width:1200px;margin:0 auto;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.status{color:var(--muted);font-size:14px;flex:1 1 200px}
pre.log{white-space:pre-wrap;background:var(--plate);border:1px solid var(--line);padding:12px;max-height:340px;overflow:auto;font:13px/1.4 Consolas,monospace;margin:12px 0 0}
.err{background:var(--badbg);color:var(--bad);padding:10px 12px;margin:12px 0 0;border:1px solid #E9B7A6}
.okmsg{background:var(--okbg);color:var(--ok);padding:10px 12px;margin:12px 0 0;border:1px solid #BFD9B5}
.hidden{display:none}
@media (max-width:640px){h1{font-size:23px}.sample{max-width:120px}}
</style>
</head>
<body>
<main>
<h1>Отчёты из CSV</h1>
<p class="lead">Папка с выгрузками → одна Excel-книга в едином оформлении. Всё распознаётся автоматически — поправьте только то, что не нравится.</p>

<section class="card" aria-labelledby="s1">
<h2 id="s1">1. Папка с CSV</h2>
<div class="row">
<div class="grow"><label for="folder">Путь к папке</label><input type="text" id="folder" placeholder="C:\Users\...\reps" autocomplete="off"></div>
<button type="button" id="pick">Выбрать…</button>
<button type="button" class="primary" id="analyze">Анализировать</button>
</div>
<div id="err1" class="err hidden" role="alert"></div>
</section>

<div id="after" class="hidden">
<section class="card" aria-labelledby="s2">
<h2 id="s2">2. Книга</h2>
<div class="grid">
<div><label for="title">Заголовок (вверху каждого листа)</label><input type="text" id="title"></div>
<div><label for="subtitle">Подзаголовок — отбор / источник</label><input type="text" id="subtitle" placeholder="ЗЛ, ДЗМ, взрослые"></div>
<div id="periodBox"><label for="period">Период к годам в шапке</label><input type="text" id="period" placeholder="9м → «2026 9м»"></div>
</div>
<div class="grid" style="margin-top:14px">
<div><label id="lgrow">Рост показателей — это</label>
<div class="seg" role="radiogroup" aria-labelledby="lgrow">
<label><input type="radio" name="grow" value="bad">плохо: рост красный</label>
<label><input type="radio" name="grow" value="good">хорошо: рост зелёный</label></div></div>
<div><label id="lpal">Оформление</label>
<div class="seg" role="radiogroup" aria-labelledby="lpal">
<label><input type="radio" name="pal" value="B">палитра B (основная)</label>
<label><input type="radio" name="pal" value="A">A (старый формат)</label></div></div>
</div>
</section>

<section class="card" aria-labelledby="s3">
<h2 id="s3">3. Типы отчётов</h2>
<div class="summary" id="summary"></div>
<p class="hint" style="margin:0 0 12px">Однотипные файлы (одинаковый набор колонок) — одна карточка. Карточки с сомнениями открыты и стоят первыми. Цвет полоски у вида — цвет шапки колонки в Excel. Вид и подпись колонки применяются ко всем файлам, где есть колонка с таким именем.</p>
<div id="types"></div>
<div id="broken"></div>
</section>
</div>

<div id="result"></div>
</main>

<div class="bar"><div class="in">
<div class="grow" style="flex:1 1 320px"><label for="out">Сохранить книгу как</label><input type="text" id="out" disabled></div>
<button type="button" class="primary" id="build" disabled>Собрать Excel</button>
<span class="status" id="status" aria-live="polite">Укажите папку и нажмите «Анализировать».</span>
</div></div>

<script>
const $ = s => document.querySelector(s);
const KINDS = %KINDS%;
let data = null;

async function api(path, body) {
  const r = await fetch(path, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body || {})});
  const j = await r.json();
  if (!r.ok) throw new Error(j.error || ("Ошибка " + r.status));
  return j;
}
function esc(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c])); }
function setStatus(t) { $("#status").textContent = t; }

$("#pick").onclick = async () => {
  setStatus("Открываю окно выбора папки…");
  try { const j = await api("/api/pick"); if (j.folder) { $("#folder").value = j.folder; analyze(); } else setStatus("Папка не выбрана."); }
  catch (e) { setStatus("Не удалось открыть окно — вставьте путь вручную."); }
};
$("#analyze").onclick = () => analyze();
$("#folder").addEventListener("keydown", e => { if (e.key === "Enter") analyze(); });

async function analyze() {
  const folder = $("#folder").value.trim().replace(/^"|"$/g, "");
  $("#err1").classList.add("hidden"); $("#result").innerHTML = "";
  if (!folder) { $("#err1").textContent = "Укажите путь к папке с CSV."; $("#err1").classList.remove("hidden"); return; }
  setStatus("Читаю файлы…"); $("#analyze").disabled = true;
  try {
    data = await api("/api/analyze", {folder});
    render();
    setStatus("Проверьте карточки с сомнениями и нажмите «Собрать Excel».");
  } catch (e) {
    $("#err1").textContent = e.message; $("#err1").classList.remove("hidden"); setStatus("");
  } finally { $("#analyze").disabled = false; }
}

function render() {
  $("#after").classList.remove("hidden");
  $("#title").value = data.book.title; $("#subtitle").value = data.book.subtitle; $("#period").value = data.book.period;
  $("#periodBox").classList.toggle("hidden", !data.has_years && !data.book.period);
  document.querySelector(`input[name=grow][value=${data.book.positive_is}]`).checked = true;
  document.querySelector(`input[name=pal][value=${data.book.palette}]`).checked = true;
  $("#out").value = data.out; $("#out").disabled = false; $("#build").disabled = false;
  const nd = data.types.filter(t => t.doubt).length;
  const nf = data.types.reduce((a, t) => a + t.files.length, 0);
  $("#summary").innerHTML = `<span class="pill">CSV: ${data.n_files}</span><span class="pill">прочитано: ${nf}</span>` +
    `<span class="pill">типов отчётов: ${data.types.length}</span>` +
    (nd ? `<span class="pill warn">с сомнениями: ${nd}</span>` : `<span class="pill">сомнений нет</span>`) +
    (data.broken.length ? `<span class="pill bad">не прочитано: ${data.broken.length}</span>` : "");
  data.types.forEach(t => t.columns.forEach(c => { c.kind_orig = c.kind; c.label_orig = c.label; }));
  $("#types").innerHTML = data.types.map(typeCard).join("");
  $("#broken").innerHTML = data.broken.length ? `<details class="type"><summary><span class="tname">Не прочитаны (${data.broken.length})</span></summary>
    <div class="tbody"><div class="scroll"><table><thead><tr><th>Файл</th><th>Причина</th></tr></thead><tbody>` +
    data.broken.map(b => `<tr><td>${esc(b.file)}</td><td>${esc(b.reason)}</td></tr>`).join("") + `</tbody></table></div></div></details>` : "";
  bind();
}

function typeCard(t, i) {
  const metrics = t.columns.filter(c => c.metric && !c.formula).length;
  const name = t.files.length === 1 ? t.files[0].file : `${t.files[0].file} и ещё ${t.files.length - 1}`;
  const wide = t.columns.length > 30;
  return `<details class="type" data-t="${i}" ${t.doubt ? "open" : ""}>
  <summary><span class="tname">Тип ${i + 1}: ${t.files.length} файл(ов)</span>
  <span class="tmeta">${esc(name)} · ${t.hierarchy ? "иерархия · " : ""}колонок ${t.n_source} · метрик ${metrics}</span>
  ${t.doubt ? '<span class="pill warn">есть сомнения</span>' : '<span class="pill">распознано уверенно</span>'}</summary>
  <div class="tbody">
  <div class="tools">
   ${t.totals_applicable ? `<label><input type="checkbox" data-totals ${t.totals ? "checked" : ""}> строка «Итого» под шапкой</label>` : ""}
   ${wide ? `<label><input type="checkbox" data-onlydoubt checked> показывать только сомнительные и метрики</label>` : ""}
  </div>
  <div class="scroll"><table><thead><tr><th>№</th><th>Колонка в CSV</th><th>Подпись в шапке</th><th>Вид</th><th>Пример</th>
  <th title="подсветка шкалой «рост / снижение»">Подсветка</th><th title="сортировка в иерархии и выделенная колонка">Главная</th><th>Пометки</th></tr></thead>
  <tbody>${t.columns.map((c, j) => colRow(t, c, j, wide)).join("")}</tbody></table></div>
  <h3 style="font-size:15px;margin:16px 0 6px;color:var(--title)">Листы</h3>
  <div class="scroll"><table class="files"><thead><tr><th>Файл</th><th>Имя листа (до 31 симв.)</th><th>Описание под заголовком</th><th>Строк</th><th>Замечания</th></tr></thead><tbody>
  ${t.files.map((f, k) => `<tr><td>${esc(f.file)}</td>
   <td><input type="text" maxlength="31" data-sheet="${k}" value="${esc(f.sheet)}" aria-label="Имя листа для ${esc(f.file)}"></td>
   <td><input type="text" data-desc="${k}" value="${esc(f.description)}" aria-label="Описание для ${esc(f.file)}"></td>
   <td>${f.rows}</td><td class="note">${esc((f.notes || []).join("; "))}</td></tr>`).join("")}</tbody></table></div>
  </div></details>`;
}

function colRow(t, c, j, wide) {
  const hide = wide && !c.doubt && !c.metric ? ' data-quiet="1" class="hidden"' : (c.doubt ? ' class="doubt"' : "");
  const kind = c.fixed ? `<span>${esc(c.fixed)}</span>` : c.formula ? `<span>${esc(KINDS[c.kind] || c.kind)} (добавлена)</span>`
    : `<select data-kind="${j}" aria-label="Вид колонки ${esc(c.name)}">${Object.entries(KINDS).map(([k, v]) =>
      `<option value="${k}" ${k === c.kind ? "selected" : ""}>${esc(v)}</option>`).join("")}</select>`;
  return `<tr${hide}><td class="num">${j + 1}</td><td>${esc(c.name)}</td>
  <td>${c.formula ? esc(c.label) : `<input type="text" class="label-in" data-label="${j}" value="${esc(c.label)}" aria-label="Подпись для ${esc(c.name)}">`}</td>
  <td class="kind k-${c.kind}" data-kcell="${j}">${kind}</td>
  <td class="sample" title="${esc(c.sample.join(", "))}">${esc(c.sample.join(", "))}</td>
  <td>${c.metric || !c.fixed ? `<input type="checkbox" data-hl="${j}" ${c.highlight ? "checked" : ""} ${c.metric ? "" : "disabled"} aria-label="Подсвечивать ${esc(c.name)}">` : ""}</td>
  <td>${c.sumable || (c.editable && !c.fixed) ? `<input type="radio" name="main${t.id}" data-main="${j}" ${c.main ? "checked" : ""} ${c.sumable ? "" : "disabled"} aria-label="Главная метрика ${esc(c.name)}">` : ""}</td>
  <td class="flag">${esc(c.flags.join("; "))}</td></tr>`;
}

function bind() {
  document.querySelectorAll("details.type[data-t]").forEach(card => {
    const t = data.types[+card.dataset.t];
    card.querySelectorAll("[data-kind]").forEach(sel => sel.onchange = () => {
      const c = t.columns[+sel.dataset.kind]; c.kind = sel.value;
      const metric = !["dimension", "skip"].includes(c.kind);
      c.metric = metric; c.sumable = ["count", "money", "number"].includes(c.kind);
      const row = sel.closest("tr");
      row.querySelector("[data-kcell]").className = "kind k-" + c.kind;
      const hl = row.querySelector("[data-hl]"); if (hl) { hl.disabled = !metric; if (!metric) { hl.checked = false; c.highlight = false; } }
      if (["delta", "delta_pct"].includes(c.kind) && hl) { hl.checked = true; c.highlight = true; }
      const mn = row.querySelector("[data-main]"); if (mn) { mn.disabled = !c.sumable; if (!c.sumable && mn.checked) { mn.checked = false; c.main = false; } }
    });
    card.querySelectorAll("[data-label]").forEach(inp => inp.oninput = () => { t.columns[+inp.dataset.label].label = inp.value; });
    card.querySelectorAll("[data-hl]").forEach(cb => cb.onchange = () => { t.columns[+cb.dataset.hl].highlight = cb.checked; });
    card.querySelectorAll("[data-main]").forEach(rb => rb.onchange = () => { t.columns.forEach((c, j) => c.main = j === +rb.dataset.main); });
    const tot = card.querySelector("[data-totals]"); if (tot) tot.onchange = () => { t.totals = tot.checked; };
    card.querySelectorAll("[data-sheet]").forEach(inp => inp.oninput = () => { t.files[+inp.dataset.sheet].sheet = inp.value; });
    card.querySelectorAll("[data-desc]").forEach(inp => inp.oninput = () => { t.files[+inp.dataset.desc].description = inp.value; });
    const only = card.querySelector("[data-onlydoubt]");
    if (only) only.onchange = () => card.querySelectorAll("tr[data-quiet]").forEach(r => r.classList.toggle("hidden", only.checked));
  });
}

$("#build").onclick = async () => {
  const sheets = data.types.flatMap(t => t.files.map(f => f.sheet.trim().toLowerCase()));
  const dup = sheets.find((s, i) => s && sheets.indexOf(s) !== i);
  if (dup) { setStatus(`Имя листа «${dup}» повторяется — исправьте, иначе к одному добавится «(2)».`); }
  const payload = {folder: data.folder, out: $("#out").value.trim(),
    book: {title: $("#title").value, subtitle: $("#subtitle").value, period: $("#period").value,
           positive_is: document.querySelector("input[name=grow]:checked").value,
           palette: document.querySelector("input[name=pal]:checked").value},
    types: data.types};
  $("#build").disabled = true; setStatus("Собираю книгу… на больших файлах это может занять минуты.");
  try {
    const j = await api("/api/build", payload);
    $("#result").innerHTML = `<section class="card"><h2>Результат</h2>
      ${j.ok ? `<div class="okmsg">Готово: ${esc(j.out)}</div>
      <div class="row" style="margin-top:12px"><button type="button" class="primary" id="openBook">Открыть книгу</button>
      <button type="button" id="openDir">Открыть папку</button></div>` : `<div class="err">Книга не собрана — причина в журнале ниже.</div>`}
      <pre class="log">${esc(j.log)}</pre></section>`;
    if (j.ok) {
      $("#openBook").onclick = () => api("/api/open", {what: "book"}).catch(e => setStatus(e.message));
      $("#openDir").onclick = () => api("/api/open", {what: "folder"}).catch(e => setStatus(e.message));
    }
    setStatus(j.ok ? "Книга собрана. Настройки сохранены в report.yaml в папке с CSV." : "Не получилось — см. журнал.");
    $("#result").scrollIntoView({behavior: "smooth"});
  } catch (e) { setStatus(e.message); }
  finally { $("#build").disabled = false; }
};
</script>
</body>
</html>
""".replace("%KINDS%", json.dumps(KIND_LABELS, ensure_ascii=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Веб-приложение: Excel-отчёты из папки CSV (локально)")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true", help="не открывать браузер автоматически")
    a = ap.parse_args()
    serve(a.port, not a.no_browser)
