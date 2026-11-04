"""
Отчёт-мейкер: Excel-книга из папки CSV любой структуры. Локально, без интернета: сервер слушает только
127.0.0.1, CSV никуда не отправляются. Движок — report_builder.py (тот же, что в консоли).

  python report_app.py                  -> откроет браузер на http://127.0.0.1:8770
  python report_app.py --port 9000 --no-browser

Вкладки: «Книга» (5 настроек) · «Колонки» (все колонки папки одной таблицей: вид, подпись, подсветка —
по одной или группой: фильтр / поиск / галочки -> «Применить к выбранным») · «Файлы» (имя листа, описание,
подзаголовок, «Итого», главная метрика, пропустить — тоже группой). Предпросмотр листа — справа.
Всё, что не трогали, решает автоматика; изменённое помечено и сбрасывается к автоматике.
Настройки сохраняются в report.yaml в папке с CSV (прежний — report.yaml.bak); консоль читает тот же файл.
"""
import argparse
import contextlib
import datetime as dt
import io
import json
import os
import re
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import report_builder as R  # noqa: E402

KIND_LABELS = {"count": "штуки", "money": "деньги", "number": "другое число", "pct": "процент / доля",
               "delta": "Δ изменение", "delta_pct": "Δ% изменение", "dimension": "разрез / текст",
               "level": "уровень иерархии", "name": "наименование", "subrow": "«из них»", "skip": "не выводить"}
VERSION = "1.4 · 04.11.2026"               # видна в шапке страницы и в консоли — чтобы не гадать, какая запущена
try:                                       # build_info.py пишет сборка exe на GitHub: коммит, из которого собран exe
    from build_info import BUILD           # noqa: E402
    VERSION += f" · сборка {BUILD}"
except ImportError:
    pass
STATE = {"last_out": None, "last_folder": None}
PREVIEW_ROWS, PREVIEW_COLS = 40, 30


# ---------------------------------------------------------------- модель страницы
def config_path_for(folder):
    return next((os.path.join(folder, n) for n in R.CONFIG_NAMES if os.path.exists(os.path.join(folder, n))), None)


def ui_kind(c):
    if getattr(c, "was_level", False):
        return "level"
    if c.role in ("level", "name", "subrow"):
        return c.role
    if c.is_metric:
        return "pct" if c.kind == "pct_points" else c.kind
    return "dimension"


def group_of(file_name):
    """Группа файла для быстрого выбора: номер отчёта до точки («19» для 19.3) или «без номера»."""
    no, _ = R.report_number(os.path.splitext(file_name)[0])
    return no.split(".")[0] if no else "без номера"


def own_file_rules(raw, f):
    """Настройки файла из конфига: все подходящие ключи files (точное имя и маски, например из шаблона)."""
    own = {}
    for pattern, rule in (raw.get("files") or {}).items():
        if R.pattern_match(pattern, f) or R.pattern_match(pattern, os.path.splitext(f)[0]):
            own = R.deep_merge(own, rule or {})
    return own


def col_state(c, hl, rule):
    return {"kind": ui_kind(c), "label": c.label, "highlight": id(c) in hl,
            "set_kind": "kind" in rule or "role" in rule, "set_label": "label" in rule,
            "set_hl": "colorize" in rule}


def analyze(folder, template=None):
    """Папка -> книга, все колонки папки (по имени, с числом файлов) и все файлы.
    template — взять настройки из шаблона вместо report.yaml папки (сохранятся в папку при сборке)."""
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise ValueError(f"Нет папки {folder}")
    cpath = config_path_for(folder)
    raw = load_template(template) if template else R.load_raw(cpath)
    cfg = R.deep_merge(R.DEFAULTS, raw)
    plan = R.plan_questions(folder, cfg, raw)
    if not plan["reports"] and not plan["broken"]:
        raise ValueError("В папке нет CSV-файлов")
    rules = raw.get("columns") or {}
    all_files = [f for f, _ in plan["reports"]]
    struct_id = {}
    columns, files = {}, []
    used = set()
    for f, rep in plan["reports"]:
        fc = R.file_config(cfg, f)
        own = own_file_rules(raw, f)
        own_cols = own.get("columns") or {}
        # подсветка — свойство колонки (правило colorize или «все Δ и Δ%»); старые списки по файлу не учитываются
        hl = {id(c) for c in R.colored_columns(rep, dict(cfg["coloring"], columns=None))}
        unsure, tech = R.doubts(rep, rules)
        main, _ = R.main_metric(rep.metrics, rep.table)
        file_cols = {}
        for c in rep.columns:
            overridden = c.name in own_cols
            st = col_state(c, hl, own_cols.get(c.name) or {})
            file_cols[c.name] = st
            item = columns.get(c.name)
            if item is None or (item["from_override"] and not overridden):
                g = col_state(c, hl, R.rules_for(rules, c.name))
                base = item or {"name": c.name, "formula": bool(c.formula), "files": [], "flags": [],
                                "sample": [str(v) for v in c.sample[:3]], "kinds_seen": {}}
                base.update(g, from_override=overridden)
                item = columns[c.name] = base
            item["files"].append(f)
            k = ui_kind(c)
            item["kinds_seen"][k] = item["kinds_seen"].get(k, 0) + 1
            flags = [x for x in R.column_flags(c) if x != "добавлена скриптом"]
            if c in tech:
                flags.append("техническое имя — подпишите")
            for x in flags:
                if x not in item["flags"]:
                    item["flags"].append(x)
        key = (rep.hierarchy, tuple(c.name for c in rep.source_columns))
        sheet_auto = R.default_sheet_name(f, all_files)
        sheet = R.sheet_title(own.get("sheet") or sheet_auto, used)
        files.append({
            "file": f, "sheet": sheet, "set_sheet": bool(own.get("sheet")), "sheet_auto": sheet_auto,
            "description": own.get("description", ""), "subtitle": own.get("subtitle", ""),
            "rows": rep.n_rows, "cols": len(rep.source_columns), "hierarchy": rep.hierarchy,
            "structure": struct_id.setdefault(key, len(struct_id) + 1), "group": group_of(f),
            "totals_applicable": not rep.hierarchy and any(c.additive for c in rep.metrics),
            "totals": str(fc["table"].get("totals", "auto")).lower() not in ("off", "false", "0"),
            "key": (rep.table.get("key") or "") if isinstance(rep.table.get("key"), str) else "",
            "key_auto": main.name if main is not None else "",
            "key_options": [c.name for c in rep.metrics if c.kind in ("count", "money", "number")],
            "skip": False, "notes": rep.notes, "doubt": bool(unsure or tech),
            "colstate": file_cols,               # как колонки выглядят в этом файле (с учётом его настроек)
            "colrules": {n: dict(s) for n, s in file_cols.items()
                         if s["set_kind"] or s["set_label"] or s["set_hl"]},   # настройки только этого файла
        })
    for f in sorted(x for x in os.listdir(folder) if x.lower().endswith(".csv")):   # пропущенные в конфиге
        if own_file_rules(raw, f).get("skip") and f not in all_files:
            files.append({"file": f, "sheet": "", "set_sheet": False, "sheet_auto": "", "description": "",
                          "subtitle": "", "rows": None, "cols": None, "hierarchy": False, "structure": 0,
                          "group": group_of(f), "totals_applicable": False, "totals": True, "key": "",
                          "key_auto": "", "key_options": [], "skip": True, "notes": [], "doubt": False,
                          "colstate": {}, "colrules": {}})
    cols = []
    for item in columns.values():
        seen = item.pop("kinds_seen")
        item.pop("from_override", None)
        if len(seen) > 1:      # в разных файлах распознано по-разному — показать и предложить выровнять
            item["flags"].append("в разных файлах: " + ", ".join(f"{KIND_LABELS.get(k, k)} — {n}"
                                                                    for k, n in seen.items()))
        item["doubt"] = bool(item["flags"])
        item["n_files"] = len(item["files"])
        cols.append(item)
    cols.sort(key=lambda c: (not c["doubt"], c["formula"], -c["n_files"], c["name"].lower()))
    book = raw.get("book") or {}
    return {
        "folder": folder, "config": cpath, "template": template or "",
        "out": book.get("output") or folder.rstrip("\\/") + ".xlsx",
        "book": {"title": book.get("title") or os.path.basename(folder.rstrip("\\/")),
                 "subtitle": book.get("subtitle") or "", "period": book.get("period") or "",
                 "positive_is": (raw.get("coloring") or {}).get("positive_is", "bad"),
                 "palette": str((raw.get("style") or {}).get("palette", "B")).upper()},
        "has_years": any(c.year for _, r in plan["reports"] for c in r.metrics),
        "columns": cols, "files": files, "kinds": KIND_LABELS, "templates": list_templates(),
        "broken": [{"file": f, "reason": e} for f, e in plan["broken"]],
        "n_files": len(plan["reports"]) + len(plan["broken"]),
        "n_structures": len(struct_id),
    }


def apply_column_state(rules, c):
    """Правило колонки по флагам set_*: True — записать явно, False — убрать (решает автоматика / общее)."""
    rule = rules.setdefault(c["name"], {})
    if not c.get("formula"):
        rule.pop("kind", None)
        rule.pop("role", None)
        if c.get("set_kind") and c.get("kind") in KIND_LABELS:
            if c["kind"] in R.ROLE_KINDS:
                rule["role"] = c["kind"]
            else:
                rule.update({"role": "metric", "kind": c["kind"]})
        rule.pop("label", None)
        label = str(c.get("label") or "").strip()
        if c.get("set_label") and label:
            rule["label"] = label
    rule.pop("colorize", None)
    if c.get("set_hl"):
        rule["colorize"] = bool(c.get("highlight"))
    if not rule:
        rules.pop(c["name"])


def settings_to_config(payload, raw):
    """Состояние страницы -> report.yaml. Флаги set_kind / set_label / set_hl: True — записать явно,
    False — убрать из конфига. Колонки — для всех файлов (columns) и для отдельных (files.X.columns).
    Прочие ключи правил (формат, ширина…) не трогаются."""
    raw = json.loads(json.dumps(raw or {}))
    book = raw.setdefault("book", {})
    b = payload.get("book") or {}
    for key in ("title", "subtitle", "period"):
        book[key] = str(b.get(key, "")).strip()
    raw.setdefault("coloring", {})["positive_is"] = "good" if b.get("positive_is") == "good" else "bad"
    raw.setdefault("style", {})["palette"] = "A" if str(b.get("palette")).upper() == "A" else "B"
    rules = raw.setdefault("columns", {})
    for c in payload.get("columns") or []:
        apply_column_state(rules, c)
    files = raw.setdefault("files", {})
    for f in payload.get("files") or []:
        entry = files.setdefault(f["file"], {})
        if f.get("skip"):
            entry["skip"] = True
            continue
        entry.pop("skip", None)
        sheet = str(f.get("sheet") or "").strip()[:31]
        if f.get("set_sheet") and sheet:
            entry["sheet"] = sheet
        else:
            entry.pop("sheet", None)
        for key in ("description", "subtitle"):
            v = str(f.get(key) or "").strip()
            if v:
                entry[key] = v
            else:
                entry.pop(key, None)
        table = entry.setdefault("table", {})
        if f.get("key"):
            table["key"] = f["key"]
        else:
            table.pop("key", None)
        if f.get("totals_applicable"):
            if f.get("totals"):
                table.pop("totals", None)
            else:
                table["totals"] = "off"
        if not table:
            entry.pop("table")
        entry.pop("coloring", None)        # подсветка — правилом колонки; старые списки по файлу не мешают
        file_rules = entry.setdefault("columns", {})
        for name, st in (f.get("colrules") or {}).items():
            apply_column_state(file_rules, dict(st, name=name))
        if not file_rules:
            entry.pop("columns")
        if not entry:
            files.pop(f["file"])
    return raw


def config_from_payload(payload):
    folder = os.path.abspath(payload["folder"])
    cpath = config_path_for(folder)
    base = load_template(payload["template"]) if payload.get("template") else R.load_raw(cpath)
    return folder, cpath, settings_to_config(payload, base)


# ---------------------------------------------------------------- шаблоны настроек
def app_home():
    """Папка приложения для шаблонов: %APPDATA%\\ReportMaker (Windows) или ~/.report_maker."""
    base = os.environ.get("REPORT_MAKER_HOME") or (
        os.path.join(os.environ["APPDATA"], "ReportMaker") if os.environ.get("APPDATA")
        else os.path.join(os.path.expanduser("~"), ".report_maker"))
    return os.path.join(base, "templates")


def safe_name(name):
    name = re.sub(r'[\\/:*?"<>|]+', "_", str(name or "").strip()).strip(". ")
    if not name:
        raise ValueError("Укажите название шаблона")
    return name[:80]


def list_templates():
    d = app_home()
    if not os.path.isdir(d):
        return []
    return sorted(os.path.splitext(f)[0] for f in os.listdir(d) if f.lower().endswith(".json"))


def load_template(name):
    path = os.path.join(app_home(), safe_name(name) + ".json")
    if not os.path.exists(path):
        raise ValueError(f"Нет шаблона «{name}»")
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("config") or {}


def generalize(file_name):
    """rep_app_stac_1_1_202610290117.csv -> re:^rep_app_stac_1_1_\\d{8,}\\.csv$ — шаблон подходит к выгрузке
    другого дня. Точная маска, а не «rep_app_stac_1_*.csv»: та подошла бы и к отчёту 1.1."""
    m = re.match(r"^(.*)_\d{8,}\.csv$", file_name, flags=re.I)
    return f"re:^{re.escape(m.group(1))}_\\d{{8,}}\\.csv$" if m else file_name


def save_template(payload):
    """Текущие настройки страницы -> шаблон (имена файлов обобщены, путь книги не сохраняется)."""
    name = safe_name(payload.get("name"))
    _, _, raw = config_from_payload(payload)
    files = {}
    for key, rule in (raw.get("files") or {}).items():
        files[generalize(key)] = R.deep_merge(files.get(generalize(key)) or {}, rule or {})
    raw["files"] = files
    (raw.get("book") or {}).pop("output", None)
    os.makedirs(app_home(), exist_ok=True)
    path = os.path.join(app_home(), name + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"name": name, "saved": dt.datetime.now().isoformat(timespec="seconds"), "config": raw},
                  f, ensure_ascii=False, indent=2)
    return {"ok": True, "name": name, "path": path, "templates": list_templates()}


# ---------------------------------------------------------------- сборка (с прогрессом)
class ProgressLog(io.StringIO):
    """Журнал сборки, который на лету считает обработанные файлы ([ok] / [пропуск])."""
    def __init__(self, job):
        super().__init__()
        self.job = job

    def write(self, text):
        for line in text.splitlines():
            if line.startswith(("[ok]", "[пропуск]")):
                self.job["done"] += 1
                self.job["current"] = line.split("<- ")[-1].strip() if "<- " in line else line.split(":")[0][10:]
        return super().write(text)


def run_build(payload, job=None):
    """Сохранить настройки и собрать книгу. -> {ok, out, log}."""
    folder, cpath, raw = config_from_payload(payload)
    job = job if job is not None else {"done": 0, "total": 0, "current": ""}
    log = ProgressLog(job)
    with contextlib.redirect_stdout(log):
        cpath = R.save_answers(raw, cpath or os.path.join(folder, "report.yaml"))
    cfg = R.load_config(cpath)
    job["total"] = sum(1 for f in os.listdir(folder) if f.lower().endswith(".csv"))
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


JOBS = {}


def start_build(payload):
    """Сборка в фоне: страница опрашивает /api/job и показывает прогресс."""
    job_id = f"{len(JOBS) + 1}"
    job = JOBS[job_id] = {"done": 0, "total": 0, "current": "", "running": True, "result": None}

    def work():
        try:
            job["result"] = run_build(payload, job)
        except Exception as err:
            job["result"] = {"ok": False, "out": "", "log": f"Ошибка: {type(err).__name__}: {err}"}
        job["running"] = False
    threading.Thread(target=work, daemon=True).start()
    return {"job": job_id}


def job_status(job_id):
    job = JOBS.get(str(job_id))
    if job is None:
        raise ValueError("Нет такой сборки")
    return job


# ---------------------------------------------------------------- предпросмотр листа
def fmt_value(v, number_format):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d.%m.%Y %H:%M" if isinstance(v, dt.datetime) and (v.hour or v.minute) else "%d.%m.%Y")
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return str(v)
    f = number_format or "General"
    if '"> 200%"' in f:
        return "> 200%"
    if "%" in f and '"%"' not in f:
        s = f"{v * 100:.1f}%".replace(".", ",")
        return ("+" + s) if (f.startswith("\\+") and v > 0) else s
    if '"%"' in f:
        return f"{v:.1f}%".replace(".", ",")
    if v == 0 and ";\\-" in f:
        return "-"
    decimals = 2 if "0.00" in f else (0 if "#,##0" in f or f == "0" else (2 if isinstance(v, float) and not
                                                                           float(v).is_integer() else 0))
    return f"{v:,.{decimals}f}".replace(",", " ").replace(".", ",")


def preview(payload):
    """Лист выбранного файла с текущими (несохранёнными) настройками: первые строки, как в Excel."""
    folder, cpath, raw = config_from_payload(payload)
    f = payload.get("file")
    if not f or not os.path.exists(os.path.join(folder, f)):
        raise ValueError("Нет такого файла")
    cfg = R.deep_merge(R.DEFAULTS, raw)
    R.apply_palette(cfg["style"])
    st = R.Styler(cfg["style"])
    fc = R.file_config(cfg, f)
    rep = R.prepare(os.path.join(folder, f), f, fc, cfg["book"].get("period") or "")
    total_rows = rep.n_rows
    if rep.n_rows > PREVIEW_ROWS:
        rep.n_rows = PREVIEW_ROWS
        for c in rep.columns:
            c.values = c.values[:PREVIEW_ROWS]
    no, rest = R.report_number(os.path.splitext(f)[0])
    rep.title = fc.get("title") or cfg["book"].get("title") or os.path.basename(folder)
    rep.description = fc.get("description") or (f"Отчёт {no}. {rest}" if no else rest)
    wb = Workbook()
    ws = wb.active
    with contextlib.redirect_stdout(io.StringIO()):
        R.write_sheet(ws, rep, fc, cfg, st)
    ncols = min(ws.max_column, PREVIEW_COLS)
    merged = {}
    covered = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row >= 5:
            merged[(rng.min_row, rng.min_col)] = (rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        covered.add((r, c))
    rows = []
    for r in range(5, ws.max_row + 1):
        cells = []
        for c in range(1, ncols + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fill_type == "solid" else None
            span = merged.get((r, c), (1, 1))
            cells.append({"v": fmt_value(cell.value, cell.number_format),
                          "bg": ("#" + fill[-6:]) if isinstance(fill, str) and fill[-6:] != "000000" else None,
                          "b": bool(cell.font and cell.font.bold), "i": bool(cell.font and cell.font.italic),
                          "fg": ("#" + cell.font.color.rgb[-6:]) if cell.font and cell.font.color is not None
                          and isinstance(cell.font.color.rgb, str) else None,
                          "ind": int(cell.alignment.indent or 0) if cell.alignment else 0,
                          "al": (cell.alignment.horizontal if cell.alignment else None) or "left",
                          "rs": span[0], "cs": min(span[1], ncols - c + 1)})
        rows.append(cells)
    return {"file": f, "sheet": fc.get("sheet") or R.default_sheet_name(f, [f]),
            "title": rep.title, "description": rep.description, "rows": rows, "shown": rep.n_rows,
            "total": total_rows, "cols": ws.max_column, "cols_shown": ncols, "hierarchy": rep.hierarchy,
            "notes": rep.notes}


# ---------------------------------------------------------------- система
PICK_CODE = """
import json, sys
try:
    import tkinter as tk
    from tkinter import filedialog
except Exception as err:
    print(json.dumps({"error": "в этом Python нет модуля tkinter (окно выбора папки): " + str(err)}))
    sys.exit()
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)      # поверх браузера
root.lift()
root.focus_force()
path = filedialog.askdirectory(parent=root, title="Папка с CSV", mustexist=True)
root.destroy()
print(json.dumps({"folder": path or ""}))   # ASCII-экранирование: кириллица не зависит от кодировки консоли
"""


def picker_command(code):
    if getattr(sys, "frozen", False):          # собранный exe: «python -c» нет — само приложение с ключом
        return [sys.executable, "--pick-folder"]
    return [sys.executable, "-c", code]


def pick_folder(code=PICK_CODE):
    """Системное окно выбора папки (tkinter в отдельном процессе — не мешает серверу).
    Ответ дочернего процесса — JSON в ASCII: на Windows канал идёт в кодировке консоли (cp1251), и путь
    с кириллицей («…\\обучение\\апп стац…») иначе не читается."""
    try:
        res = subprocess.run(picker_command(code), capture_output=True, timeout=600)
    except subprocess.TimeoutExpired:
        return {"folder": "", "error": "окно выбора папки не ответило за 10 минут"}
    except Exception as err:
        return {"folder": "", "error": f"не удалось открыть окно выбора папки: {err}"}
    out = res.stdout.decode("ascii", errors="replace").strip().splitlines()
    try:
        data = json.loads(out[-1]) if out else {}
    except ValueError:
        data = {}
    if not data:
        tail = res.stderr.decode("utf-8", errors="replace").strip().splitlines()[-1:] or ["нет ответа"]
        return {"folder": "", "error": f"окно выбора папки не открылось: {tail[0]}"}
    return {"folder": data.get("folder", ""), "error": data.get("error", "")}


def open_path(path):
    if sys.platform.startswith("win"):
        os.startfile(path)                 # noqa: S606 — локальное приложение, открывает только свою книгу
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])


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
            self.send_header("Cache-Control", "no-store")      # браузер не должен показывать старую версию из кэша
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/api/version":
            self.send_json({"app": "report-maker", "version": VERSION})
        else:
            self.send_json({"error": "нет такой страницы"}, 404)

    def do_POST(self):
        # только запросы со своей страницы: другой сайт в браузере не сможет дёрнуть локальный сервер
        origin = self.headers.get("Origin")
        if origin and origin != f"http://{self.headers.get('Host')}":
            return self.send_json({"error": "чужой источник запроса"}, 403)
        try:
            payload = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
            routes = {"/api/analyze": lambda: analyze(payload.get("folder", ""), payload.get("template") or None),
                      "/api/build": lambda: start_build(payload),
                      "/api/job": lambda: job_status(payload.get("job")),
                      "/api/preview": lambda: preview(payload),
                      "/api/template/save": lambda: save_template(payload),
                      "/api/templates": lambda: {"templates": list_templates()},
                      "/api/pick": pick_folder}
            if self.path in routes:
                return self.send_json(routes[self.path]())
            if self.path == "/api/shutdown":       # новый запуск просит прежний освободить порт
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return self.send_json({"ok": True})
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


DEFAULT_PORT = 8770      # старые версии (до 1.4) жили на 8765 — новая на них не натыкается
OLD_PORT = 8765


def probe(port):
    """Кто слушает порт: {"version": …} — наше приложение (версия есть с 1.4), "old" — наше старое
    (без версии), None — свободен или чужая программа."""
    import urllib.request
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/version", timeout=1.5) as r:
            data = json.loads(r.read())
            return data if data.get("app") == "report-maker" else None
    except urllib.error.HTTPError as err:          # старая версия: страницы /api/version нет, но это мы
        try:
            body = err.read().decode("utf-8", "replace")
        except Exception:
            body = ""
        return "old" if "нет такой страницы" in body else None
    except Exception:
        return None


def take_over(port):
    """Порт занят нашим же прежним запуском (1.4+) — попросить его закрыться и подождать."""
    import time
    import urllib.request
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/api/shutdown", data=b"{}",
                                     headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=2).read()
    except Exception:
        return False
    for _ in range(30):
        time.sleep(0.2)
        try:
            ThreadingHTTPServer(("127.0.0.1", port), Handler).server_close()
            return True
        except OSError:
            continue
    return False


def serve(port=DEFAULT_PORT, browser=True):
    server = None
    for p in range(port, port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            break
        except OSError:
            who = probe(p)
            if isinstance(who, dict) and take_over(p):
                print(f"Прежний запуск (версия {who.get('version')}) на порту {p} закрыт — открываю новую версию.")
                server = ThreadingHTTPServer(("127.0.0.1", p), Handler)
                break
    if server is None:
        sys.exit(f"Порты {port}–{port + 19} заняты")
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"Отчёт-мейкер, версия {VERSION}: {url}  (остановить — Ctrl+C или закрыть это окно)")
    if probe(OLD_PORT) == "old":
        print(f"ВНИМАНИЕ: на порту {OLD_PORT} ещё работает СТАРАЯ версия Отчёт-мейкера (без шаблонов). Закройте её "
              f"окно консоли или в Диспетчере задач завершите «ReportMaker.exe» / «python.exe». Работайте во вкладке "
              f"{url} — в шапке страницы должна быть строка «версия {VERSION}».")
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
<title>Отчёт-мейкер</title>
<style>
:root{--ink:#1F1F1F;--title:#1B3A5C;--sub:#2D6A73;--muted:#555F69;--line:#D3D9DF;--frame:#9AA5B1;--head:#E8ECF0;
--plate:#F5F7F9;--count:#DCE8F3;--money:#DCEAD3;--number:#D6ECEC;--delta:#FBF3E2;--role:#C9DCEC;
--warn:#7A5600;--warnbg:#FBEFD0;--bad:#9E3423;--badbg:#F6DAD0;--ok:#2F6A30;--okbg:#E1EEDC;--sel:#EAF1F8}
*{box-sizing:border-box}
html,body{margin:0;background:#fff;color:var(--ink);font:14px/1.4 "Segoe UI",Roboto,Arial,sans-serif}
button,input,select{font:inherit;color:inherit}
header.top{border-bottom:1px solid var(--frame);background:#fff;position:sticky;top:0;z-index:5}
.top .in{max-width:1500px;margin:0 auto;padding:10px 16px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.brand{font:700 18px Arial,sans-serif;color:var(--title);margin-right:8px;white-space:nowrap}
.top input[type=text]{flex:1 1 360px;min-width:0}
input[type=text]{padding:6px 9px;border:1px solid var(--frame);background:#fff;width:100%}
select{padding:5px 6px;border:1px solid var(--frame);background:#fff;max-width:100%}
input:focus,select:focus,button:focus-visible{outline:2px solid var(--sub);outline-offset:1px}
button{font-weight:600;padding:7px 13px;min-height:34px;border:1px solid var(--title);background:#fff;color:var(--title);cursor:pointer}
button.primary{background:var(--title);color:#fff}
button.link{border:0;background:none;padding:2px 4px;min-height:0;color:var(--title);text-decoration:underline;font-weight:500}
button:disabled{opacity:.45;cursor:default}
.wrap{max-width:1500px;margin:0 auto;padding:12px 16px 110px}
.tabs{display:flex;gap:0;border-bottom:1px solid var(--frame);margin:0 0 12px;overflow-x:auto}
.tabs button{border:0;border-bottom:3px solid transparent;background:none;color:var(--muted);padding:9px 16px;white-space:nowrap}
.tabs button[aria-selected=true]{color:var(--title);border-bottom-color:var(--title)}
.badge{display:inline-block;min-width:18px;padding:0 5px;margin-left:5px;background:var(--warnbg);color:var(--warn);font-size:12px;border-radius:9px;text-align:center}
.pills{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 10px}
.pill{padding:2px 9px;background:var(--plate);border:1px solid var(--line);font-size:12.5px}
.pill.warn{background:var(--warnbg);color:var(--warn);border-color:#E8D49A}
.pill.bad{background:var(--badbg);color:var(--bad);border-color:#E9B7A6}
.layout{display:grid;grid-template-columns:minmax(0,1fr);gap:14px}
.layout.with-preview{grid-template-columns:minmax(0,1fr) minmax(0,46%)}
.panel{min-width:0}
.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 8px}
.toolbar input[type=text]{width:220px}
.toolbar label{display:flex;gap:5px;align-items:center;color:var(--muted)}
.bulk{display:none;flex-wrap:wrap;gap:8px;align-items:center;padding:8px 10px;margin:0 0 8px;background:var(--sel);border:1px solid #BCD0E4}
.bulk.on{display:flex}
.bulk b{color:var(--title)}
.bulk .sep{width:1px;align-self:stretch;background:#BCD0E4}
.scroll{overflow:auto;border:1px solid var(--frame);max-height:calc(100vh - 290px)}
table.grid{border-collapse:collapse;width:100%}
.grid th{position:sticky;top:0;background:var(--head);color:var(--title);text-align:left;padding:6px 7px;border-bottom:1px solid var(--frame);white-space:nowrap;font-size:13px;z-index:1}
.grid td{padding:3px 7px;border-top:1px solid var(--line);height:32px;vertical-align:middle}
.grid tr.picked td{background:var(--sel)}
.grid tr.doubt td.nm{box-shadow:inset 3px 0 0 #E2B23C}
.grid td.cb,.grid th.cb{width:30px;text-align:center}
.grid td.num{color:var(--muted);text-align:right;width:52px}
.grid td.kind{border-left:6px solid transparent;min-width:150px}
.k-count{border-left-color:var(--count)!important}.k-money{border-left-color:var(--money)!important}
.k-number,.k-pct{border-left-color:var(--number)!important}.k-delta,.k-delta_pct{border-left-color:var(--delta)!important}
.k-dimension,.k-skip{border-left-color:var(--head)!important}.k-level,.k-name,.k-subrow{border-left-color:var(--role)!important}
.grid input[type=text]{padding:3px 6px;min-width:120px}
.grid select{padding:2px 4px}
.set{font-weight:700}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--sub);margin-left:4px;vertical-align:middle}
.flag{color:var(--warn);font-size:12.5px}
.sample{color:var(--muted);font-size:12.5px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.note{color:var(--bad);font-size:12.5px}
.form{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px 18px;max-width:1100px}
.form .wide{grid-column:span 2}
@media (max-width:640px){.form .wide{grid-column:auto}}
.form label.f{display:block;font-weight:600;font-size:12.5px;color:var(--muted);margin:0 0 3px}
.seg{display:flex;flex-wrap:wrap;gap:6px}
.seg label{padding:6px 11px;cursor:pointer;border:1px solid var(--frame);display:flex;gap:6px;align-items:center;white-space:nowrap}
.seg label:has(input:checked){border-color:var(--title);background:var(--sel)}
.prev{border:1px solid var(--frame);min-width:0;display:flex;flex-direction:column;max-height:calc(100vh - 200px)}
.prev .ph{display:flex;gap:8px;align-items:center;padding:7px 10px;background:var(--head);border-bottom:1px solid var(--frame)}
.prev .ph b{color:var(--title);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.prev .pb{overflow:auto;padding:10px;flex:1}
.xl-title{font:700 15px Arial,sans-serif;color:var(--title)}
.xl-sub{font:600 13px Arial,sans-serif;color:var(--sub)}
.xl-desc{font-size:12.5px;color:var(--muted);margin:0 0 6px}
table.xl{border-collapse:collapse;font:12.5px "Arial Narrow",Arial,sans-serif}
.xl td{border:1px solid #D3D9DF;padding:2px 6px;white-space:nowrap;max-width:260px;overflow:hidden;text-overflow:ellipsis}
.xl-more{color:var(--muted);font-size:12.5px;margin:6px 0 0}
.bar{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid var(--frame);z-index:6}
.bar .in{max-width:1500px;margin:0 auto;padding:9px 16px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.bar .out{flex:1 1 340px;display:flex;gap:6px;align-items:center}
.bar .out span{color:var(--muted);white-space:nowrap}
.status{color:var(--muted);flex:1 1 220px}
progress{width:180px;height:14px;accent-color:var(--title)}
#c-scope{font-size:12.5px;color:var(--sub)}
.msg{padding:8px 11px;margin:8px 0;border:1px solid}
.msg.err{background:var(--badbg);color:var(--bad);border-color:#E9B7A6}
.msg.ok{background:var(--okbg);color:var(--ok);border-color:#BFD9B5}
pre.log{white-space:pre-wrap;background:var(--plate);border:1px solid var(--line);padding:10px;max-height:300px;overflow:auto;font:12.5px/1.4 Consolas,monospace}
.empty{color:var(--muted);padding:40px 0;text-align:center}
.hidden{display:none!important}
@media (max-width:1000px){.layout.with-preview{grid-template-columns:minmax(0,1fr)}.scroll{max-height:none}}
@media (max-width:640px){.brand{width:100%}.toolbar input[type=text]{width:100%}}
</style>
</head>
<body>
<header class="top"><div class="in">
<span class="brand">Отчёт-мейкер <small style="font:400 11px Arial;color:var(--muted)">версия %VERSION%</small></span>
<input type="text" id="folder" placeholder="Папка с CSV, например C:\Users\...\reps" aria-label="Папка с CSV" autocomplete="off">
<button type="button" id="pick">Выбрать…</button>
<button type="button" class="primary" id="analyze">Анализировать</button>
<select id="tpl" aria-label="Шаблон настроек" title="Шаблон настроек: применить к этой папке" disabled><option value="">настройки папки</option></select>
<button type="button" id="tpl-save" disabled title="Сохранить текущие настройки как шаблон для других папок / периодов">Сохранить как шаблон</button>
</div></header>

<div class="wrap">
<div id="err" class="msg err hidden" role="alert"></div>
<div id="start" class="empty">Укажите папку с CSV-выгрузками и нажмите «Анализировать».<br>Файлы могут быть любой структуры: иерархия включается, если есть колонка уровня.</div>

<div id="app" class="hidden">
<div class="pills" id="summary"></div>
<div class="tabs" role="tablist">
<button type="button" role="tab" id="t-book" aria-selected="false" data-tab="book">Книга</button>
<button type="button" role="tab" id="t-cols" aria-selected="true" data-tab="cols">Колонки<span class="badge hidden" id="b-cols"></span></button>
<button type="button" role="tab" id="t-files" aria-selected="false" data-tab="files">Файлы<span class="badge hidden" id="b-files"></span></button>
</div>

<section id="p-book" class="hidden" role="tabpanel" aria-labelledby="t-book">
<div class="form">
<div><label class="f" for="title">Заголовок (вверху каждого листа)</label><input type="text" id="title"></div>
<div><label class="f" for="subtitle">Подзаголовок — отбор / источник</label><input type="text" id="subtitle" placeholder="ЗЛ, ДЗМ, взрослые"></div>
<div><label class="f" for="period">Период к годам в шапке</label><input type="text" id="period" placeholder="9м → «2026 9м»"></div>
<div class="wide"><span class="f" style="display:block;font-weight:600;font-size:12.5px;color:var(--muted);margin:0 0 3px" id="lgrow">Рост показателей — это</span>
<div class="seg" role="radiogroup" aria-labelledby="lgrow"><label title="затраты, объёмы оплаты"><input type="radio" name="grow" value="bad">плохо (красный)</label><label title="выручка, выполнение плана"><input type="radio" name="grow" value="good">хорошо (зелёный)</label></div></div>
<div class="wide"><span class="f" style="display:block;font-weight:600;font-size:12.5px;color:var(--muted);margin:0 0 3px" id="lpal">Оформление</span>
<div class="seg" role="radiogroup" aria-labelledby="lpal"><label title="стандартный шаблон"><input type="radio" name="pal" value="B">B — основная</label><label title="старый формат АПП"><input type="radio" name="pal" value="A">A — старая</label></div></div>
</div>
</section>

<div class="layout" id="layout">
<section id="p-cols" class="panel" role="tabpanel" aria-labelledby="t-cols">
<div class="toolbar">
<input type="text" id="c-q" placeholder="Поиск: Сумма, Δ, MO_…" aria-label="Поиск колонок">
<select id="c-f" aria-label="Фильтр колонок">
<option value="all">все колонки</option><option value="doubt">с пометками</option><option value="metric">метрики</option>
<option value="dim">разрезы и роли</option><option value="set">изменённые вручную</option><option value="added">добавленные Δ / Δ%</option></select>
<label title="Где действуют правки: для всех файлов или только для группы / одного файла (поверх общих)">Правки для
<select id="c-file" aria-label="Где действуют правки колонок"><option value="">всех файлов</option></select></label>
<span id="c-scope"></span>
<span class="status" id="c-count"></span>
</div>
<div class="bulk" id="c-bulk" role="region" aria-label="Действия с выбранными колонками">
<b id="c-n"></b><span class="sep"></span>
<label>Вид <select id="cb-kind"><option value="">— не менять —</option></select></label>
<label>Подсветка <select id="cb-hl"><option value="">— не менять —</option><option value="1">вкл</option><option value="0">выкл</option></select></label>
<label>Подпись <input type="text" id="cb-label" style="width:150px" placeholder="не менять"></label>
<button type="button" class="primary" id="cb-apply">Применить</button>
<button type="button" id="cb-reset" title="Убрать ручные настройки — решит автоматика">Сбросить к авто</button>
<button type="button" class="link" id="cb-clear">снять выбор</button>
</div>
<div class="scroll"><table class="grid" id="c-table"><thead><tr>
<th class="cb"><input type="checkbox" id="c-all" aria-label="Выбрать все видимые"></th><th>Колонка в CSV</th><th>Вид</th><th>Подпись в шапке</th>
<th title="шкала «рост / снижение»">Подсв.</th><th>Файлов</th><th>Пример</th><th>Пометки</th></tr></thead><tbody></tbody></table></div>
</section>

<section id="p-files" class="panel hidden" role="tabpanel" aria-labelledby="t-files">
<div class="toolbar">
<input type="text" id="f-q" placeholder="Поиск по файлу или листу" aria-label="Поиск файлов">
<select id="f-g" aria-label="Группа файлов"><option value="">все группы</option></select>
<select id="f-s" aria-label="Структура"><option value="">все структуры</option></select>
<select id="f-f" aria-label="Фильтр файлов"><option value="all">все файлы</option><option value="doubt">с пометками</option><option value="notes">с замечаниями</option><option value="skip">пропущенные</option></select>
<span class="status" id="f-count"></span>
</div>
<div class="bulk" id="f-bulk" role="region" aria-label="Действия с выбранными файлами">
<b id="f-n"></b><span class="sep"></span>
<label>Имя листа <select id="fb-sheet"><option value="">— не менять —</option><option value="auto">автоматически</option><option value="no">{номер}</option><option value="otch">Отчёт {номер}</option><option value="norest">{номер} {имя файла}</option><option value="rest">{имя файла}</option></select></label>
<label>Подзаголовок <input type="text" id="fb-sub" style="width:170px" placeholder="не менять"></label>
<label>«Итого» <select id="fb-tot"><option value="">— не менять —</option><option value="1">да</option><option value="0">нет</option></select></label>
<label>Главная метрика <select id="fb-key"><option value="">— не менять —</option><option value="@auto">автоматически</option></select></label>
<label>В книгу <select id="fb-skip"><option value="">— не менять —</option><option value="0">включить</option><option value="1">пропустить</option></select></label>
<button type="button" class="primary" id="fb-apply">Применить</button>
<button type="button" class="link" id="fb-clear">снять выбор</button>
</div>
<div class="scroll"><table class="grid" id="f-table"><thead><tr>
<th class="cb"><input type="checkbox" id="f-all" aria-label="Выбрать все видимые"></th><th>Файл</th><th>Имя листа</th><th>Описание под заголовком</th>
<th>Подзаголовок</th><th>Итого</th><th>Главная метрика</th><th>Строк×кол.</th><th>В книгу</th><th></th></tr></thead><tbody></tbody></table></div>
<div id="broken"></div>
</section>

<aside class="prev hidden" id="prev" aria-label="Предпросмотр листа">
<div class="ph"><b id="pv-name">Предпросмотр</b><button type="button" id="pv-refresh">Обновить</button><button type="button" id="pv-close" aria-label="Закрыть предпросмотр">✕</button></div>
<div class="pb" id="pv-body"></div>
</aside>
</div>

<div id="result"></div>
</div>
</div>

<div class="bar"><div class="in">
<div class="out"><span>Книга:</span><input type="text" id="out" disabled aria-label="Куда сохранить книгу"></div>
<button type="button" class="primary" id="build" disabled>Собрать Excel</button>
<progress id="prog" class="hidden" max="1" value="0" aria-label="Прогресс сборки"></progress>
<span class="status" id="status" aria-live="polite"></span>
</div></div>

<script>
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
let D = null, tab = "cols", pvFile = null;
const pickC = new Set(), pickF = new Set();
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const status = t => $("#status").textContent = t || "";
async function api(path, body) {
  const r = await fetch(path, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body || {})});
  const j = await r.json(); if (!r.ok) throw new Error(j.error || ("Ошибка " + r.status)); return j;
}
function showErr(m) { $("#err").textContent = m; $("#err").classList.toggle("hidden", !m); }

// ---------- папка
const PASTE_HINT = " Вставьте путь вручную: в Проводнике щёлкните по адресной строке папки и скопируйте путь (или Shift + правая кнопка по папке → «Копировать как путь»).";
$("#pick").onclick = async () => { status("Открываю окно выбора папки… (если не видно — проверьте панель задач)"); showErr("");
  $("#pick").disabled = true;
  try { const j = await api("/api/pick");
    if (j.error) { showErr("Кнопка «Выбрать…»: " + j.error + "." + PASTE_HINT); status(""); }
    else if (j.folder) { $("#folder").value = j.folder; analyze(); } else status("Папка не выбрана."); }
  catch (e) { showErr("Кнопка «Выбрать…»: " + e.message + "." + PASTE_HINT); status(""); }
  finally { $("#pick").disabled = false; } };
$("#analyze").onclick = () => analyze();
$("#folder").addEventListener("keydown", e => { if (e.key === "Enter") analyze(); });
async function analyze(template) {
  const folder = $("#folder").value.trim().replace(/^"|"$/g, ""); showErr(""); $("#result").innerHTML = "";
  if (!folder) return showErr("Укажите путь к папке с CSV.");
  status(template ? `Применяю шаблон «${template}»…` : "Читаю файлы…"); $("#analyze").disabled = true;
  try { D = await api("/api/analyze", {folder, template: template || ""}); pickC.clear(); pickF.clear(); init();
        status(template ? `Шаблон «${template}» применён — проверьте и нажмите «Собрать Excel» (настройки сохранятся в папку).`
                        : "Поправьте, что нужно, и нажмите «Собрать Excel». Всё остальное решит автоматика."); }
  catch (e) { showErr(e.message); status(""); } finally { $("#analyze").disabled = false; }
}
function fillTemplates(list, current) {
  $("#tpl").innerHTML = `<option value="">настройки папки</option>` + list.map(t => `<option ${t === current ? "selected" : ""}>${esc(t)}</option>`).join("");
  $("#tpl").disabled = false; $("#tpl-save").disabled = false; }
$("#tpl").onchange = () => analyze($("#tpl").value);
$("#tpl-save").onclick = async () => {
  const name = prompt("Название шаблона (например «АПП стационары»):", D.template || ""); if (!name) return;
  try { const j = await api("/api/template/save", payload({name})); fillTemplates(j.templates, D.template);
        status(`Шаблон «${j.name}» сохранён: ${j.path}. Его можно выбрать для другой папки или следующего периода.`); }
  catch (e) { showErr(e.message); } };

function init() {
  $("#start").classList.add("hidden"); $("#app").classList.remove("hidden");
  const b = D.book; $("#title").value = b.title; $("#subtitle").value = b.subtitle; $("#period").value = b.period;
  document.querySelector(`input[name=grow][value=${b.positive_is}]`).checked = true;
  document.querySelector(`input[name=pal][value=${b.palette}]`).checked = true;
  $("#out").value = D.out; $("#out").disabled = false; $("#build").disabled = false;
  D.columns.forEach(c => { c.kind0 = c.kind; c.label0 = c.label; c.hl0 = c.highlight; });
  const kopts = Object.entries(D.kinds).map(([k, v]) => `<option value="${k}">${esc(v)}</option>`).join("");
  $("#cb-kind").innerHTML = `<option value="">— не менять —</option>` + kopts;
  const live = D.files.filter(f => !f.skip), gs = {};
  live.forEach(f => gs[f.group] = (gs[f.group] || 0) + 1);
  $("#c-file").innerHTML = `<option value="">всех файлов</option>` +
    `<optgroup label="группы отчётов">${Object.entries(gs).filter(([, n]) => n > 1).map(([g, n]) => `<option value="g:${esc(g)}">${g === "без номера" ? g : "отчёты " + esc(g) + ".*"} (${n} ф.)</option>`).join("")}</optgroup>` +
    `<optgroup label="один файл">${live.map(f => `<option value="f:${esc(f.file)}">${esc(f.file)}</option>`).join("")}</optgroup>`;
  fillTemplates(D.templates, D.template);
  const groups = [...new Set(D.files.map(f => f.group))];
  $("#f-g").innerHTML = `<option value="">все группы</option>` + groups.map(g => `<option value="${esc(g)}">${g === "без номера" ? g : "отчёты " + esc(g) + ".*"}</option>`).join("");
  $("#f-s").innerHTML = `<option value="">все структуры</option>` + [...new Set(D.files.map(f => f.structure))].filter(Boolean)
    .map(s => `<option value="${s}">структура ${s} (${D.files.filter(f => f.structure === s).length} ф.)</option>`).join("");
  const nd = D.columns.filter(c => c.doubt).length, nfd = D.files.filter(f => f.doubt).length;
  $("#summary").innerHTML = `<span class="pill">CSV: ${D.n_files}</span><span class="pill">структур: ${D.n_structures}</span>` +
    `<span class="pill">колонок: ${D.columns.length}</span>` + (nd ? `<span class="pill warn">колонок с пометками: ${nd}</span>` : "") +
    (D.broken.length ? `<span class="pill bad">не прочитано: ${D.broken.length}</span>` : "");
  $("#b-cols").textContent = nd; $("#b-cols").classList.toggle("hidden", !nd);
  $("#b-files").textContent = nfd; $("#b-files").classList.toggle("hidden", !nfd);
  $("#c-f").value = nd ? "doubt" : "all";
  renderCols(); renderFiles(); renderBroken(); setTab(tab);
}

// ---------- вкладки
$$(".tabs button").forEach(b => b.onclick = () => setTab(b.dataset.tab));
function setTab(t) { tab = t;
  $$(".tabs button").forEach(b => b.setAttribute("aria-selected", b.dataset.tab === t));
  $("#p-book").classList.toggle("hidden", t !== "book"); $("#p-cols").classList.toggle("hidden", t !== "cols");
  $("#p-files").classList.toggle("hidden", t !== "files"); $("#layout").classList.toggle("hidden", t === "book"); }

// ---------- колонки
const isMetric = k => !["dimension", "skip", "level", "name", "subrow"].includes(k);
// пометки, на которые уже ответили вручную, не показываются: вид задан — «угадан по цифрам» снят, подпись — «тех. имя»
function liveFlags(c, v) { return c.flags.filter(f => f.startsWith("техническое имя") ? !v.set_label : !v.set_kind); }
// область правки: все файлы (общие правила) или группа / один файл (правила этих файлов поверх общих)
function scopeFiles() { const v = $("#c-file").value; if (!v) return null;
  return v.startsWith("g:") ? D.files.filter(f => !f.skip && f.group === v.slice(2)) : D.files.filter(f => f.file === v.slice(2)); }
function view(c) { const fs = scopeFiles(); if (!fs) return c; const f = fs.find(f => f.colrules[c.name]); return f ? f.colrules[c.name] : c; }
function own(c) { const fs = scopeFiles(); return fs ? fs.some(f => f.colrules[c.name]) : (c.set_kind || c.set_label || c.set_hl); }
function edit(c, fn) { const fs = scopeFiles(); if (!fs) return fn(c);
  fs.filter(f => f.colstate[c.name]).forEach(f => fn(f.colrules[c.name] ||= {kind: c.kind, label: c.label, highlight: c.highlight, set_kind: false, set_label: false, set_hl: false})); }
function visibleCols() {
  const q = $("#c-q").value.trim().toLowerCase(), f = $("#c-f").value, fs = scopeFiles();
  return D.columns.filter(c => { const v = view(c);
    return (!q || c.name.toLowerCase().includes(q) || (v.label || "").toLowerCase().includes(q))
    && (!fs || fs.some(x => x.colstate[c.name]))
    && (f === "all" || (f === "doubt" && c.doubt) || (f === "metric" && isMetric(v.kind)) || (f === "dim" && !isMetric(v.kind))
        || (f === "set" && own(c)) || (f === "added" && c.formula)); });
}
["#c-q", "#c-f", "#c-file"].forEach(s => $(s).addEventListener("input", renderCols));
function renderCols() {
  const rows = visibleCols(), fs = scopeFiles();
  $("#c-count").textContent = `показано ${rows.length} из ${D.columns.length}`;
  $("#c-scope").textContent = fs ? `правки — только для ${fs.length} файл(ов), поверх общих; точка — своя настройка` : "";
  $("#c-table tbody").innerHTML = rows.map(c => {
    const i = D.columns.indexOf(c), v = view(c), mark = fs ? fs.some(f => f.colrules[c.name]) : false;
    const kind = c.formula ? `<span>${esc(D.kinds[v.kind] || v.kind)} (добавлена)</span>` :
      `<select data-ck="${i}" aria-label="Вид ${esc(c.name)}" class="${v.set_kind ? "set" : ""}">${Object.entries(D.kinds).map(([k, t]) => `<option value="${k}" ${k === v.kind ? "selected" : ""}>${esc(t)}</option>`).join("")}</select>${v.set_kind || mark ? `<span class="dot" title="${mark ? "своя настройка этих файлов" : "задано вручную"}"></span>` : ""}`;
    return `<tr class="${pickC.has(i) ? "picked" : ""} ${c.doubt ? "doubt" : ""}">
    <td class="cb"><input type="checkbox" data-cp="${i}" ${pickC.has(i) ? "checked" : ""} aria-label="Выбрать ${esc(c.name)}"></td>
    <td class="nm">${esc(c.name)}</td><td class="kind k-${v.kind}">${kind}</td>
    <td><input type="text" data-cl="${i}" value="${esc(v.label)}" class="${v.set_label ? "set" : ""}" aria-label="Подпись ${esc(c.name)}"></td>
    <td class="cb"><input type="checkbox" data-ch="${i}" ${v.highlight ? "checked" : ""} ${isMetric(v.kind) ? "" : "disabled"} aria-label="Подсветка ${esc(c.name)}"></td>
    <td class="num">${c.n_files}</td><td class="sample" title="${esc(c.sample.join(", "))}">${esc(c.sample.join(", "))}</td>
    <td class="flag">${esc(liveFlags(c, v).join("; "))}</td></tr>`; }).join("");
  $("#c-all").checked = rows.length > 0 && rows.every(c => pickC.has(D.columns.indexOf(c)));
  bulkState();
}
$("#c-table").addEventListener("change", e => { const t = e.target;
  if (t.dataset.cp !== undefined) { const i = +t.dataset.cp; t.checked ? pickC.add(i) : pickC.delete(i); t.closest("tr").classList.toggle("picked", t.checked); bulkState(); return; }
  if (t.dataset.ck !== undefined) { edit(D.columns[+t.dataset.ck], o => setKind(o, t.value)); renderCols(); schedulePreview(); }
  if (t.dataset.ch !== undefined) { edit(D.columns[+t.dataset.ch], o => { o.highlight = t.checked; o.set_hl = true; }); schedulePreview(); }
});
$("#c-table").addEventListener("input", e => { const t = e.target;
  if (t.dataset.cl !== undefined) { edit(D.columns[+t.dataset.cl], o => { o.label = t.value; o.set_label = true; }); t.classList.add("set"); schedulePreview(); } });
function setKind(c, k) { c.kind = k; c.set_kind = true;
  if (!isMetric(k)) { c.highlight = false; c.set_hl = c.set_hl || c.hl0; }
  else if (["delta", "delta_pct"].includes(k) && !c.highlight) { c.highlight = true; c.set_hl = true; } }
$("#c-all").onchange = () => { visibleCols().forEach(c => { const i = D.columns.indexOf(c); $("#c-all").checked ? pickC.add(i) : pickC.delete(i); }); renderCols(); };
function bulkState() { $("#c-bulk").classList.toggle("on", pickC.size > 0); $("#c-n").textContent = `Выбрано: ${pickC.size}`;
  $("#f-bulk").classList.toggle("on", pickF.size > 0); $("#f-n").textContent = `Выбрано: ${pickF.size}`; }
$("#cb-clear").onclick = () => { pickC.clear(); renderCols(); };
$("#cb-apply").onclick = () => {
  const k = $("#cb-kind").value, hl = $("#cb-hl").value, lb = $("#cb-label").value.trim(); let n = 0;
  pickC.forEach(i => { const c = D.columns[i]; n++;
    edit(c, o => {
      if (k && !c.formula) setKind(o, k);
      if (hl !== "" && isMetric(o.kind)) { o.highlight = hl === "1"; o.set_hl = true; }
      if (lb && !c.formula) { o.label = lb; o.set_label = true; } }); });
  $("#cb-kind").value = ""; $("#cb-hl").value = ""; $("#cb-label").value = "";
  status(`Применено к ${n} колонкам.`); renderCols(); schedulePreview(); };
$("#cb-reset").onclick = () => { const fs = scopeFiles();
  pickC.forEach(i => { const c = D.columns[i];
    if (fs) fs.forEach(f => delete f.colrules[c.name]);
    else { c.kind = c.kind0; c.label = c.label0; c.highlight = c.hl0; c.set_kind = false; c.set_label = false; c.set_hl = false; } });
  status(fs ? "У выбранных колонок в этих файлах — снова общие настройки." : "Выбранные колонки — снова на автоматике (после «Собрать» распознаются заново).");
  renderCols(); schedulePreview(); };

// ---------- файлы
function visibleFiles() {
  const q = $("#f-q").value.trim().toLowerCase(), g = $("#f-g").value, s = $("#f-s").value, f = $("#f-f").value;
  return D.files.filter(x => (!q || x.file.toLowerCase().includes(q) || (x.sheet || "").toLowerCase().includes(q))
    && (!g || x.group === g) && (!s || String(x.structure) === s)
    && (f === "all" || (f === "doubt" && x.doubt) || (f === "notes" && x.notes.length) || (f === "skip" && x.skip)));
}
["#f-q", "#f-g", "#f-s", "#f-f"].forEach(s => $(s).addEventListener("input", renderFiles));
function renderFiles() {
  const rows = visibleFiles();
  $("#f-count").textContent = `показано ${rows.length} из ${D.files.length}`;
  $("#f-table tbody").innerHTML = rows.map(x => { const i = D.files.indexOf(x);
    const keyOpts = `<option value="">авто${x.key_auto ? ": " + esc(x.key_auto) : ""}</option>` + x.key_options.map(k => `<option ${k === x.key ? "selected" : ""}>${esc(k)}</option>`).join("");
    return `<tr class="${pickF.has(i) ? "picked" : ""} ${x.doubt ? "doubt" : ""}">
    <td class="cb"><input type="checkbox" data-fp="${i}" ${pickF.has(i) ? "checked" : ""} aria-label="Выбрать ${esc(x.file)}"></td>
    <td class="nm" title="${esc(x.notes.join("; "))}">${esc(x.file)}${x.hierarchy ? ' <span class="pill">иерархия</span>' : ""}${x.notes.length ? '<br><span class="note">' + esc(x.notes.join("; ")) + "</span>" : ""}</td>
    <td><input type="text" maxlength="31" data-fs="${i}" value="${esc(x.sheet)}" class="${x.set_sheet ? "set" : ""}" ${x.skip ? "disabled" : ""} aria-label="Имя листа ${esc(x.file)}"></td>
    <td><input type="text" data-fd="${i}" value="${esc(x.description)}" ${x.skip ? "disabled" : ""} aria-label="Описание ${esc(x.file)}"></td>
    <td><input type="text" data-fu="${i}" value="${esc(x.subtitle)}" placeholder="как у книги" ${x.skip ? "disabled" : ""} aria-label="Подзаголовок ${esc(x.file)}"></td>
    <td class="cb">${x.totals_applicable ? `<input type="checkbox" data-ft="${i}" ${x.totals ? "checked" : ""} aria-label="Итого ${esc(x.file)}">` : ""}</td>
    <td>${x.key_options.length ? `<select data-fk="${i}" aria-label="Главная метрика ${esc(x.file)}">${keyOpts}</select>` : ""}</td>
    <td class="num">${x.rows ?? "—"}×${x.cols ?? "—"}</td>
    <td class="cb"><input type="checkbox" data-fx="${i}" ${x.skip ? "" : "checked"} aria-label="Включить ${esc(x.file)} в книгу"></td>
    <td><button type="button" class="link" data-fv="${i}" ${x.skip ? "disabled" : ""}>просмотр</button></td></tr>`; }).join("");
  $("#f-all").checked = rows.length > 0 && rows.every(x => pickF.has(D.files.indexOf(x)));
  const keys = new Set(); pickF.forEach(i => D.files[i].key_options.forEach(k => keys.add(k)));
  $("#fb-key").innerHTML = `<option value="">— не менять —</option><option value="@auto">автоматически</option>` + [...keys].map(k => `<option>${esc(k)}</option>`).join("");
  bulkState();
}
$("#f-table").addEventListener("change", e => { const t = e.target, d = t.dataset;
  if (d.fp !== undefined) { const i = +d.fp; t.checked ? pickF.add(i) : pickF.delete(i); renderFiles(); return; }
  if (d.ft !== undefined) { D.files[+d.ft].totals = t.checked; previewIf(+d.ft); }
  if (d.fk !== undefined) { D.files[+d.fk].key = t.value; previewIf(+d.fk); }
  if (d.fx !== undefined) { D.files[+d.fx].skip = !t.checked; renderFiles(); }
});
$("#f-table").addEventListener("input", e => { const t = e.target, d = t.dataset;
  if (d.fs !== undefined) { const x = D.files[+d.fs]; x.sheet = t.value; x.set_sheet = true; t.classList.add("set"); }
  if (d.fd !== undefined) { D.files[+d.fd].description = t.value; previewIf(+d.fd); }
  if (d.fu !== undefined) { D.files[+d.fu].subtitle = t.value; } });
$("#f-table").addEventListener("click", e => { const d = e.target.dataset; if (d.fv !== undefined) openPreview(D.files[+d.fv].file); });
$("#f-all").onchange = () => { visibleFiles().forEach(x => { const i = D.files.indexOf(x); $("#f-all").checked ? pickF.add(i) : pickF.delete(i); }); renderFiles(); };
$("#fb-clear").onclick = () => { pickF.clear(); renderFiles(); };
function sheetFrom(x, how) { const m = x.file.replace(/\.csv$/i, "").replace(/_\d{8,}$/, "");
  const no = (x.sheet_auto.match(/^(?:Отчёт )?(\d+(?:\.\d+)*)/) || [])[1] || "";
  const rest = x.sheet_auto.replace(/^(?:Отчёт )?\d+(?:\.\d+)*\s*/, "") || m.replace(/_/g, " ");
  return ({auto: x.sheet_auto, no: no || rest, otch: no ? "Отчёт " + no : rest, norest: (no + " " + rest).trim(), rest})[how].slice(0, 31); }
$("#fb-apply").onclick = () => {
  const sh = $("#fb-sheet").value, sub = $("#fb-sub").value.trim(), tot = $("#fb-tot").value, key = $("#fb-key").value, sk = $("#fb-skip").value; let n = 0;
  pickF.forEach(i => { const x = D.files[i]; n++;
    if (sh) { x.sheet = sheetFrom(x, sh); x.set_sheet = sh !== "auto"; }
    if (sub) x.subtitle = sub;
    if (tot !== "" && x.totals_applicable) x.totals = tot === "1";
    if (key === "@auto") x.key = ""; else if (key && x.key_options.includes(key)) x.key = key;
    if (sk !== "") x.skip = sk === "1"; });
  ["#fb-sheet", "#fb-tot", "#fb-key", "#fb-skip"].forEach(s => $(s).value = ""); $("#fb-sub").value = "";
  status(`Применено к ${n} файлам.`); renderFiles(); schedulePreview(); };
function renderBroken() { $("#broken").innerHTML = D.broken.length ? `<p class="note" style="margin:10px 0 4px">Не прочитаны (${D.broken.length}):</p>` +
  D.broken.map(b => `<div class="note">${esc(b.file)} — ${esc(b.reason)}</div>`).join("") : ""; }

// ---------- предпросмотр
function payload(extra) { return Object.assign({folder: D.folder, out: $("#out").value.trim(),
  book: {title: $("#title").value, subtitle: $("#subtitle").value, period: $("#period").value,
         positive_is: document.querySelector("input[name=grow]:checked").value, palette: document.querySelector("input[name=pal]:checked").value},
  columns: D.columns, files: D.files, template: D.template || ""}, extra || {}); }
function openPreview(file) { pvFile = file; $("#prev").classList.remove("hidden"); $("#layout").classList.add("with-preview"); loadPreview(); }
$("#pv-close").onclick = () => { pvFile = null; $("#prev").classList.add("hidden"); $("#layout").classList.remove("with-preview"); };
$("#pv-refresh").onclick = () => loadPreview();
let pvTimer = null;
function schedulePreview() { if (!pvFile) return; clearTimeout(pvTimer); pvTimer = setTimeout(loadPreview, 500); }
function previewIf(i) { if (pvFile === D.files[i].file) schedulePreview(); }
$("#c-file").addEventListener("change", () => { const v = $("#c-file").value; pickC.clear(); renderCols();
  if (v.startsWith("f:")) openPreview(v.slice(2)); else if (v.startsWith("g:")) { const fs = scopeFiles(); if (fs.length) openPreview(fs[0].file); } });
async function loadPreview() {
  if (!pvFile) return; $("#pv-name").textContent = "Предпросмотр: " + pvFile;
  try { const p = await api("/api/preview", payload({file: pvFile}));
    $("#pv-body").innerHTML = `<div class="xl-title">${esc(p.title)}</div>${$("#subtitle").value ? `<div class="xl-sub">${esc($("#subtitle").value)}</div>` : ""}
    <div class="xl-desc">${esc(p.description)}</div><table class="xl">${p.rows.map(r => `<tr>${r.map(c =>
      `<td${c.rs > 1 ? ` rowspan="${c.rs}"` : ""}${c.cs > 1 ? ` colspan="${c.cs}"` : ""} style="${c.bg ? "background:" + c.bg + ";" : ""}${c.b ? "font-weight:700;" : ""}${c.i ? "font-style:italic;" : ""}${c.fg ? "color:" + c.fg + ";" : ""}text-align:${c.al === "center" ? "center" : c.al === "right" ? "right" : "left"};${c.ind ? "padding-left:" + (6 + c.ind * 12) + "px;" : ""}">${esc(c.v)}</td>`).join("")}</tr>`).join("")}</table>
    <p class="xl-more">Показано строк: ${p.shown} из ${p.total}${p.cols > p.cols_shown ? `, колонок: ${p.cols_shown} из ${p.cols}` : ""}. Раскраска на превью — по показанным строкам.${p.notes.length ? "<br>" + esc(p.notes.join("; ")) : ""}</p>`; }
  catch (e) { $("#pv-body").innerHTML = `<div class="msg err">${esc(e.message)}</div>`; }
}

// ---------- сборка
const sleep = ms => new Promise(r => setTimeout(r, ms));
$("#build").onclick = async () => {
  $("#build").disabled = true; status("Собираю книгу…"); const pr = $("#prog"); pr.value = 0; pr.classList.remove("hidden");
  try { const {job} = await api("/api/build", payload()); let st;
    while (true) { st = await api("/api/job", {job});
      if (st.total) { pr.max = st.total; pr.value = Math.min(st.done, st.total); }
      if (!st.running) break;
      status(st.total ? `Собираю: ${Math.min(st.done, st.total)} из ${st.total}${st.current ? " — " + st.current : ""}` : "Сохраняю настройки…");
      await sleep(400); }
    const j = st.result;
    $("#result").innerHTML = `<h2 style="font:700 16px Arial;color:var(--title);margin:16px 0 6px">Результат</h2>${j.ok ? `<div class="msg ok">Готово: ${esc(j.out)}</div>
      <div style="display:flex;gap:8px"><button type="button" class="primary" id="openBook">Открыть книгу</button><button type="button" id="openDir">Открыть папку</button></div>`
      : `<div class="msg err">Книга не собрана — причина в журнале.</div>`}<pre class="log">${esc(j.log)}</pre>`;
    if (j.ok) { $("#openBook").onclick = () => api("/api/open", {what: "book"}).catch(e => status(e.message));
                $("#openDir").onclick = () => api("/api/open", {what: "folder"}).catch(e => status(e.message)); }
    status(j.ok ? "Книга собрана. Настройки сохранены в report.yaml в папке с CSV." : "Не получилось — см. журнал.");
    $("#result").scrollIntoView({block: "start"}); }
  catch (e) { status(e.message); } finally { $("#build").disabled = false; pr.classList.add("hidden"); }
};
</script>
</body>
</html>
""".replace("%VERSION%", VERSION)


if __name__ == "__main__":
    R.safe_console()
    if "--pick-folder" in sys.argv:            # окно выбора папки в отдельном процессе (нужно для exe)
        exec(PICK_CODE)
        sys.exit()
    if "--version" in sys.argv:
        print(f"Отчёт-мейкер, версия {VERSION}")
        sys.exit()
    ap = argparse.ArgumentParser(description="Отчёт-мейкер: Excel-книга из папки CSV любой структуры (локально)")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--no-browser", action="store_true", help="не открывать браузер автоматически")
    a = ap.parse_args()
    serve(a.port, not a.no_browser)
