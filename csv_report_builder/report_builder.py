"""
Универсальная сборка Excel-отчёта из папки CSV: каждый CSV -> отдельный лист в едином оформлении
(палитра B «стандартный шаблон» с канваса «Палитра отчётов»; палитра A — старый формат АПП) + лист «Содержание»
+ лист «Справка по колонкам».

Что делает сам, без настроек:
  • читает CSV: кодировка (UTF-8 / Windows-1251), разделитель (; , TAB |), десятичная запятая, «12,5%»,
    «\\N»/NULL — пусто; «? » в начале заголовка (Δ, потерянная при выгрузке в cp1251) -> «Δ »;
  • определяет роль каждой колонки:
      level      — «Уровень» (0 = Всего, 1, 2 …) -> иерархический отчёт: структура — как в CSV (родитель, под ним дети;
                   «из них …» — сразу под своей строкой), сортируются только дети внутри групп; отступы,
                   заливка итоговых уровней, жирный «Всего» и 1-й уровень (как в отчётах АПП);
      name       — «Наименование» (в иерархическом отчёте — главная колонка);
      dimension  — разрезы/текст/коды/даты (коды с ведущими нулями остаются текстом);
      metric     — числовые колонки; вид: count (целые), money (дробные / «сумма», «руб»), pct («%», «доля»),
                   delta («Δ», «дельта», «изм»), delta_pct («Δ%»);
  • колонки «X 2026» / «X 2025» -> шапка в 2 строки по годам; если Δ в CSV нет — добавляет Δ и Δ%;
    Δ% больше 200% показывает как «> 200%»;
  • в ячейках — только значения и заливки (файл не тормозит, ничего не «плывёт»); формулы (Δ, Δ%, Итого) и
    правила раскраски записаны текстом: в примечании к заголовку колонки и на скрытом листе «Формулы» —
    скопировать формулу в первую ячейку диапазона и протянуть вниз (table.formulas: live — формулы прямо в ячейках);
  • раскраска Δ и Δ% мягкой цветовой шкалой (рост — красный, снижение — зелёный; можно наоборот),
    по умолчанию — заливкой ячеек (без условного форматирования), можно — условным форматированием;
  • плоские таблицы: строка «Итого» сразу под шапкой, вне автофильтра (формула SUBTOTAL по видимым строкам — на листе «Формулы»), автофильтр;
  • ширина колонок по содержимому, закреплённая шапка, печать по ширине, «Страница N из M», ссылки
    «← Содержание» на каждом листе.

Уточняющие вопросы: при запуске в консоли скрипт анализирует папку, группирует файлы по структуре (одинаковый
набор колонок) и по каждой новой структуре показывает таблицу колонок с распознанным видом: поправить вид,
выбрать, что подсвечивать, главную метрику, строку «Итого», подписи технических имён; затем имена листов.
Enter — вариант в скобках, «!» — принять все остальные. Ответы сохраняются в report.yaml (прежний — в .bak),
при следующем запуске спрашивается только новое. --ask — спросить всё заново, --no-ask — без вопросов.

Шапка колонок окрашена по виду (палитра B): штуки — голубой, деньги — зелёный, прочие числа и проценты —
бирюзовый, Δ / Δ% — песочный, разрезы — нейтральный серо-голубой.

Всё, что нужно вписать руками (названия, описания, подписи колонок, какие колонки метрики, что красить),
задаётся в конфиге report.yaml в папке с CSV. Шаблон с найденными файлами и колонками:
  python report_builder.py "D:\\папка_с_csv" --init       -> создаст report.yaml (существующий не трогает)
  python report_builder.py "D:\\папка_с_csv" --explain    -> покажет, как распознаны колонки, без сборки
  python report_builder.py "D:\\папка_с_csv"              -> соберёт <папка>.xlsx рядом с папкой
  python report_builder.py "D:\\папка" --out "D:\\итог.xlsx" --config "D:\\общий_конфиг.yaml"

Нужны: pip install pandas openpyxl pyyaml   (pyyaml — только для конфига .yaml; .json работает без него)
Цифры из CSV не пересчитываются: в ячейки попадают значения CSV; добавленные Δ / Δ% / Итого посчитаны скриптом
и записаны числами, их формулы — на листе «Формулы» и в примечаниях к заголовкам.
"""
import argparse
import copy
import csv
import datetime as dt
import fnmatch
import json
import math
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import coordinate_to_tuple, get_column_letter
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.hyperlink import Hyperlink

try:
    import yaml
except ImportError:
    yaml = None

CONFIG_NAMES = ("report.yaml", "report.yml", "report.json")
TOC_SHEET, HELP_SHEET, FORMULAS_SHEET = "Содержание", "Справка по колонкам", "Формулы"
EXCEL_MAX_ROWS, EXCEL_MAX_COLS, EXCEL_MAX_TEXT = 1_048_576, 16_384, 32_767
BIG_SHEET_CELLS = 1_500_000              # больше — предупреждение о времени (и памяти без потоковой записи)
STREAM_BYTES = 2 * 1024 * 1024           # CSV в сумме больше (~25 тыс. строк) — потоковая запись
                                         # (book.stream: auto | true | false); книга та же, память — в разы меньше
try:
    csv.field_size_limit(sys.maxsize)
except OverflowError:          # Windows: long — 32 бита
    csv.field_size_limit(2 ** 31 - 1)

# ---------------------------------------------------------------- настройки по умолчанию (переопределяются конфигом)
DEFAULTS = {
    "book": {
        "title": None,                 # заголовок книги; по умолчанию — имя папки
        "subtitle": "",                # красная строка под заголовком (на всех листах)
        "period": "",                  # подпись периода к годам в шапке: «2026 9м»
        "notes": [],                   # пояснения на листе «Содержание» (список строк)
        "output": None,                # путь к .xlsx; по умолчанию <папка>.xlsx рядом с папкой
        "columns_sheet": True,         # лист «Справка по колонкам»
        "formulas_sheet": "hidden",    # лист «Формулы» (формулы и правила раскраски текстом): hidden | visible | off
        "header_notes": True,          # примечание с формулой / правилом раскраски на заголовке колонки
    },
    "style": {
        "font": "Arial Narrow", "size": 14,
        "palette": "B",                # B — стандартный шаблон (основной) | A — старый формат АПП
        "colors": {},                  # переопределение цветов палитры: {title: "1B3A5C", levels: {0: "C9DCEC"}}
        "view": "page_break",          # page_break (страничный, как в АПП) | normal
        "zoom": 70,
        "paper": "auto",               # auto (A3, если таблица широкая) | A3 | A4
        "orientation": "landscape",
    },
    "table": {
        "compare": "auto",             # auto — добавить Δ / Δ% для пар «X 2026 / X 2025»; off — не добавлять
        "formulas": "values",          # values — в ячейках числа, формулы — текстом на листе «Формулы» | live
        "totals": "auto",              # плоские таблицы: auto/on — строка «Итого» под шапкой; off — нет
        "autofilter": True,            # плоские таблицы
        "sort": None,                  # «-Сумма 2026» — по убыванию; «Наименование» — по возрастанию
        "key": None,                   # главная метрика: выделенная колонка, сортировка детей в иерархии
        "hierarchy_sort": "auto",      # иерархия: auto — дети по убыванию главной метрики; off — порядок CSV
        "check": "money",              # иерархия: сверка «родитель = сумма строк ниже» — money (деньги) | all | off
        "outline": True,               # иерархия: группировка строк Excel (+/−) по уровням
        "row_numbers": "auto",         # колонка «№»: auto — только в иерархических отчётах
        "column_numbers": True,        # строка с номерами колонок под шапкой
        "freeze_columns": "auto",      # сколько колонок закрепить слева; auto — колонки до первой метрики (≤3)
        "max_width": 45, "name_max_width": 90,
    },
    "coloring": {
        "mode": "static",              # static — цвет заливкой (рекомендуется) | conditional — условное форм. | off
        "positive_is": "bad",          # bad — рост красный, снижение зелёный | good — наоборот
        "columns": None,               # список колонок для раскраски; None — все Δ и Δ%
        "gamma": 2.5,                  # насыщенность = доля^gamma: средние значения еле заметны
        "pct_cap": 2.0,                # Δ% больше 200% -> «> 200%», край шкалы Δ%
    },
    "columns": {},                     # правила по именам колонок (для всех файлов): см. --init
    "files": {},                       # настройки отдельных файлов по маске имени: см. --init
}

# ---------------------------------------------------------------- палитра (из build_app_excel12.py)
# Палитры. B — «стандартный шаблон (основной)» с канваса «Палитра отчётов»: сине-серые уровни, тёплые цвета —
# только акценты, без фиолетовых. A — старый формат АПП (build_app_excel12.py). Любой цвет переопределяется
# в конфиге: style: {palette: B, colors: {title: "1B3A5C", levels: {0: "C9DCEC"}}}
PALETTES = {
    "B": {
        "text": "1F1F1F", "title": "1B3A5C", "subtitle": "2D6A73", "muted": "555F69", "service": "555F69",
        "numbers": "66707A", "subrow": "555F69", "link": "1B3A5C", "alert": "9E3423",
        "head_fill": "E8ECF0", "head_text": "1B3A5C", "head_delta_text": "1B3A5C",
        "head_kind_fills": {"count": "DCE8F3",         # штуки — голубой
                            "money": "DCEAD3",         # деньги — шалфей
                            "number": "D6ECEC", "pct": "D6ECEC", "pct_points": "D6ECEC",   # прочие — бирюзовый
                            "delta": "FBF3E2", "delta_pct": "FBF3E2"},                     # Δ — песок светлый
        "group_fill_year": "E8ECF0", "group_fill_delta": "FBF3E2",
        "levels": {0: "C9DCEC", 1: "DCE8F3", 2: "EAF1F8", 3: "EEF4EA"}, "level_deep": "F5F7F9",
        "total": "C9DCEC", "key_column": "ECEFF2",
        "head_border": "9AA5B1", "data_border": "D3D9DF", "outer_border": "9AA5B1",
        "red": "F6DAD0", "green": "E1EEDC",           # пары «Отклонение» / «Норма, рост» из блока 4
        "title_font": "Arial",                        # заголовки — Arial, таблицы — Arial Narrow
    },
    "A": {
        "text": "000000", "title": "002060", "subtitle": "C00000", "muted": "808080", "service": "595959",
        "numbers": "BFBFBF", "subrow": "7F7F7F", "link": "0563C1", "alert": "C00000",
        "head_fill": "F2F2F2", "head_text": "000000", "head_delta_text": "808080",
        "head_kind_fills": {"count": "DEEBF7", "money": "E2F0D9", "pct": "FFF2CC", "delta": "FFF2CC",
                            "delta_pct": "FFF2CC"},
        "group_fill_year": "F2F2F2", "group_fill_delta": "FFF2CC",
        "levels": {0: "F4B183", 1: "FBE5D6", 2: "FFF2CC", 3: "DDEBF7", 4: "EDEDED"}, "level_deep": "EDEDED",
        "total": "F4B183", "key_column": None,
        "head_border": "BFBFBF", "data_border": "D9D9D9", "outer_border": None,
        "red": "FCE0E0", "green": "C8E8D0",
        "title_font": None,
    },
}
CF_WHITE = "FFFFFF"
P = {}                                   # активная палитра (apply_palette)
HEAD_BORDER = DATA_BORDER = None


def apply_palette(style):
    """Палитра из style.palette + переопределения style.colors; границы — из её цветов."""
    global P, HEAD_BORDER, DATA_BORDER
    name = str(style.get("palette", "B")).upper()
    if name not in PALETTES:
        sys.exit(f"Нет палитры «{name}» — есть: {', '.join(PALETTES)}")
    P = deep_merge(PALETTES[name], style.get("colors") or {})
    P["levels"] = {int(k): v for k, v in P["levels"].items()}
    head = Side(style="thin", color=P["head_border"])
    data = Side(style="thin", color=P["data_border"])
    HEAD_BORDER = Border(left=head, right=head, top=head, bottom=head)
    DATA_BORDER = Border(left=data, right=data, top=data, bottom=data)


def level_fill(level):
    return P["levels"].get(level, P["level_deep"])


def head_fill(kind):
    return P["head_kind_fills"].get(kind, P["head_fill"])

FMT = {
    "count": r"#,##0;\-#,##0;\-",
    "money": r"#,##0.00;\-#,##0.00;\-",
    "number": r"#,##0.##;\-#,##0.##;\-",
    "pct": r"0.0%",
    "pct_points": r'0.0"%"',
    "delta_pct": r"\+0.0%;\-0.0%;0.0%",
    "date": "DD.MM.YYYY",
    "datetime": "DD.MM.YYYY HH:MM",
    "text": "@",
}
FMT_PCT_CAP = '"> 200%"'
CHAR_W, HEAD_CHAR_W = 1.3, 1.65   # ширина символа (данные / жирная шапка) в единицах Excel при размере 14

# ---------------------------------------------------------------- распознавание по имени колонки
RX = {
    # уровень иерархии: «Уровень», «Ур», «Ур.», «Level», «Lvl», «Lev», «Depth», «Глубина», «Уровень иерархии»,
    # level_1 / lvl_x … (любую колонку можно отметить уровнем вручную: columns: {X: {role: level}})
    "level": re.compile(r"^(уровень( иерархии)?|ур\.?|level|lvl|lev|depth|глубина|hier(archy)?[_ ]?level|"
                        r"(level|lvl)[_ ]\w+)$", re.I),
    "name": re.compile(r"^(наименование|название|name|наим\.?)$", re.I),
    "subrow": re.compile(r"^из них$", re.I),
    "id": re.compile(r"(^|[_\s])(id|code|cod|код|кода|инн|inn|огрн|ogrn|кпп|kpp|снилс|snils|окпо|okpo|окато|okato|"
                     r"октмо|oktmo|gid|pat_id|полис|polis|enp|номер|num|nom|ord|order|sort|rank|type|flag|"
                     r"порядок|тип|признак|категория|кат|статус|status|класс|class)"
                     r"($|[_\s])|^(№|n|no|nn|пп|№ п/п)$", re.I),
    "period": re.compile(r"^(год|year|месяц|month|квартал|quarter|период|period|дата|date)\b", re.I),
    "delta_pct": re.compile(r"^(Δ|δ)\s*%|^(дельта|delta|изм\w*)\s*%", re.I),
    "delta": re.compile(r"^(Δ|δ|дельта|delta|diff|разница|изм)", re.I),
    "pct": re.compile(r"%|процент|доля|pct|percent|share|удельн", re.I),
    "money": re.compile(r"сумм|руб|стоим|цена|выручк|затрат|тариф|оплат|amount|price|cost|revenue|₽", re.I),
    "count": re.compile(r"случа|пациент|кол-?во|количеств|\bшт\b|услуг|посещ|визит|обращ|госпит|койк|человек|"
                        r"count|qty|quantity|число", re.I),
}
YEAR_RE = re.compile(r"^(?P<pre>.*?)[\s_,]*(?P<year>(?:19|20)\d\d)(?!\d)(?P<post>.*)$")
SUBROW_RE = re.compile(r"^\s*из них\b", re.I)
TOTAL_RE = re.compile(r"^\s*(всего|итого|total)\b", re.I)
INT_TEXT = re.compile(r"^-?(0|[1-9]\d{0,14})$")          # «0» — число; «01» (ведущий ноль) — текст
DATE_LIKE = re.compile(r"^\d{1,4}[.\-/]\d{1,2}[.\-/]\d{1,4}([ T]\d{1,2}:\d{2}(:\d{2})?)?$")
NULLS = {"", "\\N", "NULL", "null", "None", "nan", "NaN", "#N/A"}


# ---------------------------------------------------------------- конфиг
def deep_merge(base, extra):
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else copy.deepcopy(v)
    return out


def load_config(path):
    if not path:
        return copy.deepcopy(DEFAULTS)
    with open(path, encoding="utf-8") as f:
        if path.lower().endswith(".json"):
            data = json.load(f)
        else:
            if yaml is None:
                sys.exit("Для конфига .yaml нужен pyyaml: pip install pyyaml (или используйте report.json)")
            data = yaml.safe_load(f) or {}
    return deep_merge(DEFAULTS, data)


def pattern_match(pattern, name):
    """Ключ правила: точное имя (без учёта регистра), маска (* ? [..]) или «re:регулярка»."""
    pattern = str(pattern)
    if pattern.startswith("re:"):
        return re.search(pattern[3:], name, re.I) is not None
    if any(ch in pattern for ch in "*?["):
        return fnmatch.fnmatch(name.lower(), pattern.lower())
    return pattern.lower() == name.lower()


def rules_for(rules, name):
    """Все подходящие правила по порядку записи сливаются; точное имя — сильнее масок."""
    merged = {}
    for pattern, rule in (rules or {}).items():
        if pattern_match(pattern, name) and any(ch in str(pattern) for ch in "*?[") | str(pattern).startswith("re:"):
            merged.update(rule or {})
    for pattern, rule in (rules or {}).items():
        if not str(pattern).startswith("re:") and not any(ch in str(pattern) for ch in "*?[") \
                and pattern_match(pattern, name):
            merged.update(rule or {})
    return merged


def file_config(cfg, file_name):
    """Настройки файла: общие + все подходящие маски из files (поздние перекрывают ранние)."""
    fc = {"table": cfg["table"], "coloring": cfg["coloring"], "columns": dict(cfg["columns"])}
    own = {}
    for pattern, rule in (cfg.get("files") or {}).items():
        if pattern_match(pattern, file_name) or pattern_match(pattern, os.path.splitext(file_name)[0]):
            own = deep_merge(own, rule or {})
    fc["table"] = deep_merge(cfg["table"], own.pop("table", {}))
    fc["coloring"] = deep_merge(cfg["coloring"], own.pop("coloring", {}))
    for name, rule in (own.pop("columns", {}) or {}).items():    # правило файла — поверх общего, по ключам
        fc["columns"][name] = dict(fc["columns"].get(name) or {}, **(rule or {}))
    fc.update(own)          # sheet, title, subtitle, description, order, skip
    return fc


# ---------------------------------------------------------------- чтение CSV
def decode_bytes(raw):
    """-> (текст, кодировка). UTF-16 / UTF-8 (с BOM и без) / Windows-1251; двоичный файл — ошибка."""
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "utf-16"
    if b"\x00" in raw[:4096]:
        raise ValueError("не текстовый файл (двоичные данные)")
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError("не удалось определить кодировку")


SINGLE_COLUMN = "\x1f"                    # разделитель, которого нет в тексте: весь файл — одна колонка
DECIMAL_LINE = re.compile(r"^\s*\"?-?[\d\s\u00a0]+,\d+\s*%?\"?\s*$")


def sep_name(sep):
    return {SINGLE_COLUMN: "нет (одна колонка)", "\t": "TAB"}.get(sep, f"«{sep}»")


def detect_sep(text):
    """Разделитель — тот, что даёт одинаковое число полей (>1) в большинстве первых строк; кавычки учитываются."""
    lines = [ln for ln in text.splitlines()[:200] if ln.strip()][:60]
    best, best_score = ";", (-1, 0, 0)
    for sep in (";", "\t", ",", "|"):
        counts = [len(r) for r in csv.reader(lines, delimiter=sep)]
        if not counts:
            continue
        mode = max(set(counts), key=counts.count)
        score = (int(mode > 1), counts.count(mode) / len(counts), mode)
        if score > best_score:
            best, best_score = sep, score
    # «55723,69» в каждой строке — это десятичная запятая в файле из одной колонки, а не два поля
    # но «0,397» / «310,427» при шапке «Случаи 2026,Случаи 2025» — две колонки целых: запятая в шапке без
    # пробела (имена колонок) и разная длина «дробной части» — признаки разделителя
    if best == "," and len(lines) > 1:
        decimal = [ln for ln in lines[1:] if DECIMAL_LINE.match(ln)]
        head = lines[0]
        head_is_text = "," not in head or re.search(r",\s", head) and not re.search(r",\S", head)
        tails = {len(ln.strip().strip('"').rstrip("%").split(",")[-1]) for ln in decimal}
        if (len(decimal) >= 0.9 * (len(lines) - 1) and not any(ch in head for ch in ";\t|")
                and head_is_text and len(tails) == 1):
            return SINGLE_COLUMN
    return best


def read_csv_any(path, sep=None):
    """-> (DataFrame строк, кодировка, разделитель, замечания). Все значения — строки, пустые — ''.
    Разбор сам (модуль csv): кавычки и переносы внутри поля, пустые строки пропускаются, короткие строки
    дополняются, лишние поля — в колонки «Колонка N» (с замечанием), пустые/повторяющиеся имена колонок —
    уникальные, пустая колонка от «;» в конце строки — убирается."""
    with open(path, "rb") as fh:
        raw = fh.read()
    text, encoding = decode_bytes(raw)
    text = text.lstrip("\ufeff")
    if not text.strip():
        raise ValueError("пустой файл")
    sep = {"tab": "\t", "none": SINGLE_COLUMN}.get(str(sep).lower(), sep) if sep else detect_sep(text)
    rows = list(csv.reader(_lines(text), delimiter=sep))
    filled = [r for r in rows if any(v.strip() for v in r)]
    if not filled:
        raise ValueError("пустой файл")
    header = filled[0]
    start = next(i for i, r in enumerate(rows) if r is header or r == header) + 1
    body = [r for r in rows[start:] if any(v.strip() for v in r)]
    notes = []
    blank = sum(1 for r in rows[start:] if r and not any(v.strip() for v in r))
    if blank:
        notes.append(f"{blank} полностью пустых строк(и) пропущено")
    width = max([len(header)] + [len(r) for r in body])
    extra = [r for r in body if len(r) > len(header) and any(v.strip() for v in r[len(header):])]
    if extra:
        notes.append(f"{len(extra)} строк(и) с лишними полями — добавлены колонки «Колонка N»")
    else:
        width = len(header)
    limit = EXCEL_MAX_COLS - 50          # запас под добавленные Δ / Δ% и служебные колонки
    if width > limit:
        notes.append(f"колонок {width} — больше, чем помещается на лист Excel; оставлены первые {limit}")
        width = limit
    header = (header + [""] * (width - len(header)))[:width]
    body = [(r + [""] * (width - len(r)))[:width] for r in body]
    names, seen = [], {}
    for i, h in enumerate(header):
        h = fix_header(h) or f"Колонка {i + 1}"
        base = h
        while h.lower() in seen:
            seen[base.lower()] += 1
            h = f"{base} ({seen[base.lower()]})"
        seen.setdefault(h.lower(), 1)
        names.append(h)
    renamed = [n for n in names if re.search(r" \(\d+\)$", n) and n not in header]
    if renamed:
        notes.append("повторяющиеся имена колонок переименованы: " + ", ".join(renamed))
    # dtype object, не «str»: при установленном pyarrow pandas хранит строки в Arrow и гоняет регулярки через RE2,
    # который не понимает часть шаблонов; с object всегда работает обычный модуль re
    df = pd.DataFrame(body, columns=names, dtype=object)
    df = df.apply(lambda s: s.str.strip())
    df = df.mask(df.isin(NULLS), "")
    bad = df.apply(lambda s: s.str.contains(ILLEGAL_CHARACTERS_RE)).to_numpy().sum() if len(df) else 0
    if bad:
        df = df.apply(lambda s: s.str.replace(ILLEGAL_CHARACTERS_RE, "", regex=True))
        notes.append(f"{bad} знач. с управляющими символами — символы удалены")
    long_ = df.apply(lambda s: s.str.len() > EXCEL_MAX_TEXT).to_numpy().sum() if len(df) else 0
    if long_:
        df = df.apply(lambda s: s.where(s.str.len() <= EXCEL_MAX_TEXT, s.str.slice(0, EXCEL_MAX_TEXT - 1) + "…"))
        notes.append(f"{long_} знач. длиннее {EXCEL_MAX_TEXT} символов (лимит Excel) — обрезаны")
    # колонка без имени и без значений (лишний «;» в конце строк) — не нужна
    for i, h in enumerate(header):
        if not h.strip() and len(df) and (df[names[i]] == "").all():
            df = df.drop(columns=[names[i]])
    return df, encoding, sep, notes


def _lines(text):
    """Строки для csv.reader с сохранением переводов строк внутри кавычек."""
    import io
    return io.StringIO(text, newline="")


def fix_header(name):
    name = ILLEGAL_CHARACTERS_RE.sub("", str(name)).strip().strip('"')[:255]
    if name.startswith("?%") or name.startswith("? %"):
        return "Δ%" + name.split("%", 1)[1]
    if name.startswith("? "):
        return "Δ " + name[2:]
    return name


def to_numbers(series):
    """Строки -> float (NaN, где не число). Пробелы-разделители, десятичная запятая, «12,5%» (-> 0.125)."""
    series = series.astype(object)                                 # см. read_csv_any: регулярки — через re
    s = series.str.replace("[\\s\u00a0\u202f]", "", regex=True)       # пробелы, в т.ч. неразрывные
    s = s.str.replace(r"(?i)(₽|руб\.?|р\.|\$|€|usd|eur|rub)$", "", regex=True).str.replace(r"^[$€₽]", "", regex=True)
    neg = s.str.match(r"^\(.*\)$")                                 # (150) — бухгалтерский минус
    s = s.where(~neg, "-" + s.str.slice(1, -1))
    pct = s.str.endswith("%")
    s = s.str.rstrip("%")
    both = s.str.contains(",", regex=False) & s.str.contains(".", regex=False)
    s = s.where(~both, s.str.replace(",", "", regex=False))        # 1,234.56 -> 1234.56
    s = s.str.replace(",", ".", regex=False)                        # 1234,56 -> 1234.56
    num = pd.to_numeric(s, errors="coerce")
    num = num.where(~num.isin([math.inf, -math.inf]))                # inf в Excel не записать — не число
    return num.where(~pct, num / 100), bool(pct[series != ""].all()) if (series != "").any() else False


# ---------------------------------------------------------------- колонки
class Col:
    def __init__(self, name):
        self.name = name
        self.label = name          # подпись в шапке (нижняя строка)
        self.group = None          # верхняя строка шапки
        self.role = "dimension"    # level | name | dimension | metric | service | subrow | skip
        self.kind = "text"         # text | date | datetime | count | money | number | pct | pct_points | delta | delta_pct
        self.fmt = None
        self.values = []           # значения для записи (None — пусто)
        self.formula = None        # ("diff"|"ratio", Col a, Col b) — добавленные колонки
        self.year = None
        self.base = None           # «Случаи» для «Случаи 2026»
        self.width = None
        self.hidden = False
        self.colorize = None
        self.positive_is = None
        self.description = ""
        self.cap = False           # Δ% > cap -> «> 200%»
        self.hint = None           # формула для справки у колонок из CSV (в ячейках — значения CSV)
        self.kind_source = "name"  # config | name | values — по значениям = не уверен, спросить
        self.maybe_code = False    # целые без подсказки в имени, все разные — возможно, код, а не метрика
        self.sample = []           # примеры значений для вопросов
        self.notes = []            # замечания по колонке (в отчёт по файлу)
        self.numeric_share = 0.0   # доля чисел среди значений (для подсказки «может, метрика?»)
        self.odd_value = ""        # пример не-числа в почти числовой колонке
        self.nosum = False         # Δ от «другого числа» — не суммировать

    @property
    def is_metric(self):
        return self.role == "metric"

    @property
    def additive(self):
        # «другое число» (балл, коэффициент) и его Δ в «Итого» не суммируются
        return self.is_metric and self.kind in ("count", "money", "delta") and not self.nosum


DATE_FORMATS = ["%d.%m.%Y", "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%d.%m.%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y", "%d-%m-%Y", "%Y.%m.%d"]


def parse_dates(series):
    """Строгий разбор: для колонки выбирается формат (или пара «дата» + «дата время»), который распознаёт больше
    всего значений. Никакого угадывания «день/месяц»: 2026-02-01 — 1 февраля, 2026-13-01 — не дата."""
    filled = series.where(series != "")
    best, best_n, best_fmt = None, -1, None
    for fmt in DATE_FORMATS:
        parsed = pd.to_datetime(filled, format=fmt, errors="coerce")
        if " " in fmt or "T" in fmt:     # «дата время» вместе с «дата» без времени
            short = pd.to_datetime(filled, format=fmt.split(" ")[0].split("T")[0], errors="coerce")
            parsed = parsed.fillna(short)
        n = int(parsed.notna().sum())
        if n > best_n:
            best, best_n, best_fmt = parsed, n, fmt
    return best, (best_fmt if best_n > 0 else None)


def is_code_name(name):
    """Имя кода / номера / типа (ORD, CASE_TYPE, Код МО, Порядок…), если в нём нет признака метрики
    («Order amount», «Сумма по типу» — метрики)."""
    return bool(RX["id"].search(name)) and not RX["money"].search(name) and not RX["count"].search(name)


def norm(text):
    return re.sub(r"[^0-9a-zа-яё]", "", str(text).lower())


def analyze_column(name, series, rule):
    series = series.astype(object)          # регулярки — через re, не RE2 (pyarrow), см. read_csv_any
    col = Col(name)
    role, kind = rule.get("role"), rule.get("kind")
    nonempty = series[series != ""]
    if role is None:
        if RX["level"].match(name):
            role = "level"
        elif RX["name"].match(name):
            role = "name"
        elif RX["subrow"].match(name):
            role = "subrow"
    nums, all_pct_strings = to_numbers(series) if len(series) else (pd.Series(dtype=float), False)
    numeric_share = nums[series != ""].notna().mean() if len(nonempty) else 0.0
    leading_zero = nonempty.str.match(r"^0\d").any() if len(nonempty) else False
    if role is None:
        if len(nonempty) and numeric_share >= 0.98 and not leading_zero \
                and not is_code_name(name) and not RX["period"].match(name):
            role = "metric"
        else:
            role = "dimension"
    col.role = role

    col.sample = list(dict.fromkeys(nonempty.head(200)))[:3]
    col.numeric_share = float(numeric_share)
    odd = nonempty[nums[series != ""].isna()] if len(nonempty) else nonempty
    col.odd_value = odd.iloc[0] if len(odd) else ""
    if role == "metric":
        if kind is None:
            fractional = bool(((nums.dropna() % 1) != 0).any())
            if RX["delta_pct"].search(name):
                kind = "delta_pct"
            elif RX["delta"].match(name):
                kind = "delta"
            elif all_pct_strings or RX["pct"].search(name):
                kind = "pct" if (all_pct_strings or nums.abs().max() <= 1.5) else "pct_points"
            elif RX["money"].search(name):
                kind = "money"
            elif RX["count"].search(name):
                kind = "count"
            else:                       # по имени не понять — решаем по значениям и потом уточняем
                kind = "money" if fractional else "count"
                col.kind_source = "values"
                vals = nums.dropna()
                col.maybe_code = (not fractional and len(vals) >= 5 and vals.is_unique and (vals >= 0).all())
        else:
            col.kind_source = "config"
        col.kind = kind
        col.values = [None if pd.isna(v) else (int(v) if kind == "count" and float(v).is_integer() else float(v))
                      for v in nums]
        if kind == "delta":   # формат Δ — как у исходной метрики: целые или с копейками
            col.fmt = FMT["money"] if bool(((nums.dropna() % 1) != 0).any()) or RX["money"].search(name) else FMT["count"]
        col.cap = kind == "delta_pct"
    elif role in ("dimension", "service") and kind in (None, "date", "datetime") and len(nonempty) \
            and nonempty.str.match(DATE_LIKE).mean() >= 0.98:
        dates, fmt = parse_dates(series)
        ok = dates[series != ""].notna()
        if fmt and ok.mean() >= 0.9:      # дата — значением; не распознанные (30.02) — остаются текстом
            has_time = bool((dates.dropna().dt.time != dt.time(0)).any())
            col.kind = "datetime" if has_time else "date"
            col.values = [v if v == "" or v is None else (d.to_pydatetime() if not pd.isna(d) else v)
                          for v, d in zip(series, dates)]
            col.values = [None if v == "" else v for v in col.values]
            if not ok.all():
                col.notes.append(f"«{name}»: {int((~ok).sum())} знач. не похожи на дату "
                                 f"(например «{series[series != ''][~ok].iloc[0]}») — оставлены текстом")
    if not col.values:
        if role == "level" and len(nonempty) and nums[series != ""].notna().all() \
                and bool((nums.dropna() % 1 == 0).all()):
            col.kind = "count"
            col.values = [None if pd.isna(v) else int(v) for v in nums]
        else:
            if role == "level":           # «Уровень» не из целых чисел — обычный разрез
                if rule.get("role") == "level":
                    col.notes.append(f"«{name}» отмечена уровнем, но значения не целые числа — оставлена разрезом")
                col.role = role = "dimension"
            col.kind = kind or "text"
            # коды из цифр без ведущих нулей — числом (Excel не ругается на «число как текст»); kind: text — как есть
            keep_text = kind == "text"
            col.values = [None if v == "" else (int(v) if INT_TEXT.match(v) and not keep_text else v)
                          for v in series]

    col.fmt = rule.get("format") or col.fmt or FMT.get(col.kind)
    col.label = rule.get("label", name)
    col.group = rule.get("group")
    col.width = rule.get("width")
    col.hidden = bool(rule.get("hidden", False))
    col.colorize = rule.get("colorize")
    col.positive_is = rule.get("positive_is")
    col.description = rule.get("description", "")
    if rule.get("role") == "skip":
        col.role = "skip"
    return col


def apply_year_groups(cols, period, compare):
    """«X 2026» / «X 2025» -> группы шапки по годам; при compare=auto и отсутствии Δ — добавить Δ и Δ%."""
    metrics = [c for c in cols if c.is_metric]
    for c in metrics:
        if c.kind in ("delta", "delta_pct"):
            continue
        m = YEAR_RE.match(c.name)
        if m:
            c.year = int(m.group("year"))
            pre, post = m.group("pre").strip(" _,"), m.group("post").strip(" _")
            c.base = (pre + post if post.startswith(",") else f"{pre} {post}").strip(" _,")
    years = sorted({c.year for c in metrics if c.year}, reverse=True)
    if len(years) < 2:
        for c in metrics:
            c.year = None
        return cols
    y1, y0 = years[0], years[1]
    suffix = f" {period}" if period else ""
    by_base = {}
    for c in metrics:
        if c.year:
            by_base.setdefault(norm(c.base), {})[c.year] = c
            if c.group is None:
                c.group = f"{c.year}{suffix}"
            if c.label == c.name:
                c.label = c.base or c.name
    has_delta = any(c.kind == "delta" for c in metrics)
    pct_names = {norm(re.sub(RX["delta_pct"], "", c.name)) for c in metrics if c.kind == "delta_pct"}
    added_delta, added_pct = [], []
    for key, pair in by_base.items():
        if y1 not in pair or y0 not in pair or pair[y1].kind not in ("count", "money", "number"):
            continue
        a, b = pair[y1], pair[y0]
        if compare == "auto" and not has_delta:
            d = Col(f"Δ {a.base}")
            d.role, d.kind, d.formula, d.fmt = "metric", "delta", ("diff", a, b), a.fmt
            d.nosum = a.kind == "number"
            d.values = [None if (x is None and y is None) else round((x or 0) - (y or 0), 6)
                        for x, y in zip(a.values, b.values)]
            d.description = f"добавлено: {a.name} − {b.name}"
            added_delta.append(d)
        if compare == "auto" and key not in pct_names:
            p = Col(f"Δ% {a.base}")
            p.role, p.kind, p.formula, p.fmt, p.cap = "metric", "delta_pct", ("ratio", a, b), FMT["delta_pct"], True
            p.values = [None if not y else (x or 0) / y - 1 for x, y in zip(a.values, b.values)]
            p.description = f"добавлено: {a.name} / {b.name} − 1 (при {b.name} = 0 — пусто)"
            added_pct.append(p)
    for c in metrics:   # Δ из CSV: если нашлась пара по названию — формула для ручной правки (в ячейках — значения CSV)
        if c.kind == "delta":
            pair = by_base.get(norm(re.sub(RX["delta"], "", c.name, count=1)))
            if pair and y1 in pair and y0 in pair:
                c.hint = ("diff", pair[y1], pair[y0])
    cols = cols + added_delta + added_pct
    for c in cols:
        if c.is_metric and c.kind == "delta" and c.group is None:
            c.group = f"Δ {y1}-{y0}"
            c.label = re.sub(r"^(Δ|δ|дельта|delta|diff|разница|изм\w*)\s*", "", c.label, flags=re.I) or c.label
        elif c.is_metric and c.kind == "delta_pct" and c.group is None:
            c.group = f"Δ% {y1}/{y0}"
            c.label = re.sub(RX["delta_pct"], "", c.label).strip() or c.label
    # порядок метрик: годы по убыванию, Δ, Δ%, прочие метрики; остальные колонки — на своих местах
    rank = {y: i for i, y in enumerate(years)}

    def metric_order(c):
        if c.year:
            return (0, rank[c.year])
        return {"delta": (1, 0), "delta_pct": (2, 0)}.get(c.kind, (3, 0))
    metric_cols = sorted([c for c in cols if c.is_metric], key=metric_order)
    it = iter(metric_cols)
    first = next(i for i, c in enumerate(cols) if c.is_metric)
    rest = [c for c in cols if not c.is_metric]
    return rest[:first] + list(it) + rest[first:]


# ---------------------------------------------------------------- подготовка отчёта
class Report:
    pass


def prepare(path, file_name, fc, period):
    df, encoding, sep, notes = read_csv_any(path, fc.get("sep"))
    if sep == SINGLE_COLUMN and not fc.get("sep"):
        notes.append("запятая принята за десятичную, файл — одна колонка (если это разделитель: files: "
                     f"{{\"{file_name}\": {{sep: \",\"}}}} в report.yaml)")
    if df.empty:
        raise ValueError("нет данных (только шапка или пустые строки)")
    rules = fc["columns"]
    table = fc["table"]
    cols = []
    for name in df.columns:
        rule = rules_for(rules, name)
        c = analyze_column(name, df[name], rule)
        if c.role != "skip":
            cols.append(c)
    source_columns = list(cols)
    no_metrics = not any(c.is_metric for c in cols)
    level = next((c for c in cols if c.role == "level"), None)
    if level is not None and any(v is None for v in level.values):
        level.role = "dimension"
        level = None
    rep = Report()
    rep.file, rep.encoding, rep.sep = file_name, encoding, sep
    rep.source_columns = source_columns
    rep.notes = notes + [n for c in cols for n in c.notes]
    if no_metrics:
        rep.notes.append("числовых колонок нет — таблица без итогов и раскраски")
    rep.hierarchy = level is not None
    rep.n_rows = len(df)
    if rep.n_rows > EXCEL_MAX_ROWS - 20:
        print(f"   [!] {file_name}: {rep.n_rows} строк — больше лимита Excel, лишние строки обрезаны")
        rep.n_rows = EXCEL_MAX_ROWS - 20
        for c in cols:
            c.values = c.values[:rep.n_rows]
    cols = apply_year_groups(cols, period, table["compare"])

    rows = list(range(rep.n_rows))
    if rep.hierarchy:
        issues = []
        rows = hierarchy_order(cols, level, table, issues)
    elif table.get("sort"):
        rows = flat_order(cols, table["sort"], file_name)
    for c in cols:
        c.values = [c.values[i] for i in rows]
    rep.check_cells = set()
    if rep.hierarchy and issues:
        pos = {i: k for k, i in enumerate(rows)}
        rep.check_cells = {(pos[i], id(c)) for i, c, _, _ in issues}
        names = next((c for c in cols if c.role == "name"), None)
        examples = []
        for i, c, parent, total in issues[:3]:
            who = names.values[pos[i]] if names is not None else f"строка {pos[i] + 1}"
            num = lambda v: f"{v:,.2f}".replace(",", " ").replace(".", ",")   # noqa: E731
            examples.append(f"«{who}», {c.name}: {num(parent)} ≠ {num(total)}")
        rep.check_notes = [f"сверка сумм: {len(issues)} расхожд. «строка ≠ сумма строк ниже» (ячейки выделены): "
                           + "; ".join(examples) + (" …" if len(issues) > 3 else "")]
    else:
        rep.check_notes = []
    rep.notes += rep.check_notes

    # раскладка колонок
    if rep.hierarchy:
        name = next((c for c in cols if c.role == "name"), None) or next(
            (c for c in cols if c.role == "dimension" and c.kind == "text"), None)
        if name is None:
            raise ValueError("иерархический отчёт без колонки «Наименование»")
        name.role = "name"
        # пустое наименование: уровень 0 — «Всего», иначе — значение самого глубокого заполненного разреза
        dims = [c for c in cols if c.role == "dimension" and c is not name]
        for i, v in enumerate(name.values):
            if v in (None, ""):
                filled = [c.values[i] for c in dims if c.values[i] not in (None, "")]
                name.values[i] = "Всего" if level.values[i] == 0 else (filled[-1] if filled else None)
        keys = [c for c in cols if c.role in ("dimension", "subrow") and c is not name]
        name.label = rules_for(rules, name.name).get("label") or " / ".join(level_labels(level, keys, rules)) or name.name
        for c in keys:
            c.role = "service" if c.role == "dimension" else c.role
        level.role = "service"
        level.was_level = True          # для приложения: показать как «уровень», а не «служебная»
        metrics = [c for c in cols if c.is_metric]
        service = [level] + [c for c in cols if c.role in ("service", "subrow") and c is not level]
        rep.columns = [name] + metrics + service
        rep.level = level
        rep.subrow = next((c for c in cols if c.role == "subrow"), None)
        rep.name_col = name
    else:
        rep.columns = [c for c in cols if c.role != "level"] + [c for c in cols if c.role == "level"]
        rep.level, rep.subrow = None, None
        rep.name_col = next((c for c in rep.columns if not c.is_metric), None)
    rep.metrics = [c for c in rep.columns if c.is_metric]
    rep.table = table
    rn = table.get("row_numbers", "auto")
    rep.row_numbers = rep.hierarchy if rn == "auto" else bool(rn)
    if rn == "auto" and not rep.hierarchy and rep.columns and all(c.is_metric for c in rep.columns):
        rep.row_numbers = True    # одни метрики — колонка «№» нужна, чтобы было где подписать «Итого»
    return rep


def level_labels(level, keys, rules):
    """Подписи уровней: колонка разреза относится к уровню, на котором она впервые заполнена."""
    by_level = {}
    for c in keys:
        if c.role == "subrow":
            continue
        lv = [l for l, v in zip(level.values, c.values) if v not in (None, "")]
        if lv:
            by_level.setdefault(min(lv), []).append(c)
    labels = []
    for _, group in sorted(by_level.items()):
        names = []
        for c in group:
            label = rules_for(rules, c.name).get("level_label") or c.label
            if label not in names:
                names.append(label)
        labels.append(", ".join(names))
    return labels


def main_metric(cols, table):
    """Главная метрика: table.key (выделенная колонка и сортировка детей в иерархии; порядок строк плоской
    таблицы не меняет), иначе table.sort, иначе — деньги (или первая метрика) самого свежего года."""
    sort = table.get("key") or table.get("sort")
    if sort:
        name = sort.lstrip("-+")
        found = next((c for c in cols if c.name.lower() == name.lower()), None)
        if found:
            return found, not sort.startswith("-")
    metrics = [c for c in cols if c.is_metric and c.kind in ("count", "money", "number")]
    if not metrics:
        return None, False
    newest = max((c.year or 0) for c in metrics)
    pool = [c for c in metrics if (c.year or 0) == newest]
    return next((c for c in pool if c.kind == "money"), pool[0]), False


def subrow_flags(cols, n):
    """Подстроки «из них …»: заполнена колонка «Из них» или наименование начинается с «из них»."""
    sub = next((c for c in cols if c.role == "subrow"), None)
    name = next((c for c in cols if c.role == "name"), None)
    return [bool(sub is not None and sub.values[i] not in (None, ""))
            or bool(name is not None and isinstance(name.values[i], str) and SUBROW_RE.match(name.values[i]))
            for i in range(n)]


def hierarchy_order(cols, level, table, issues=None):
    """Порядок строк иерархического отчёта. Порядок строк в CSV не важен (выгрузка часто идёт по уровням:
    сначала все группы 1-го уровня, потом все 2-го…):
      • родитель строки уровня N — строка уровня N−1 с теми же значениями разрезов; разрез относится к уровню,
        на котором он впервые заполнен (MO_NAME — уровень 1, GROUP_NAME — уровень 2 …);
      • подстрока «из них …» — сразу под строкой с теми же разрезами того же уровня или уровнем выше
        (ЖК -> «из них операции»);
      • если по разрезам родителя не найти — ближайшая строка выше с меньшим уровнем (как в CSV).
    Дети каждой группы — по главной метрике по убыванию (при равенстве — та же метрика прошлого года, затем
    порядок CSV); hierarchy_sort: off — дети в порядке CSV, но под своими родителями."""
    n = len(level.values)
    do_sort = str(table.get("hierarchy_sort", "auto")).lower() != "off"
    metric, ascending = main_metric(cols, table)
    second = None
    if metric is not None and metric.year:
        second = next((c for c in cols if c.is_metric and c.base == metric.base and c.year and c.year < metric.year), None)
    is_sub = subrow_flags(cols, n)
    keys = [c for c in cols if c.role == "dimension"]
    col_level = {}
    for c in keys:
        lv = [l for l, v, sub in zip(level.values, c.values, is_sub) if v not in (None, "") and not sub]
        col_level[id(c)] = min(lv) if lv else 10 ** 6

    def path(i, lvl):
        return tuple("" if c.values[i] is None else str(c.values[i]) for c in keys if col_level[id(c)] <= lvl)

    nodes = [{"i": i, "level": level.values[i], "children": [], "subs": []} for i in range(n)]
    by_path = {}
    for node in nodes:
        if not is_sub[node["i"]]:
            by_path.setdefault((node["level"], path(node["i"], node["level"])), node)
    root = {"i": None, "level": -1, "children": [], "subs": []}
    stack, last = [root], root            # запасной путь — по порядку строк CSV
    for node in nodes:
        i, lvl = node["i"], node["level"]
        if is_sub[i]:   # владелец — строка того же уровня с теми же разрезами, иначе уровнем выше (оба формата)
            owner = (by_path.get((lvl, path(i, lvl))) or by_path.get((lvl - 1, path(i, lvl - 1)))) if keys else None
            (owner or last)["subs"].append(i)
            continue
        while len(stack) > 1 and stack[-1]["level"] >= lvl:
            stack.pop()
        parent = root if lvl <= 0 else (by_path.get((lvl - 1, path(i, lvl - 1))) if keys else None)
        if parent is None or parent is node:
            parent = stack[-1]
        parent["children"].append(node)
        stack.append(node)
        last = node

    def val(c, i):
        v = c.values[i] if c is not None else None
        return v if isinstance(v, (int, float)) else 0.0

    sign = 1 if ascending else -1

    def key(node):
        i = node["i"]
        return (sign * val(metric, i), sign * val(second, i), i) if do_sort else (i,)

    out = list(root["subs"])

    def walk(node):
        out.append(node["i"])
        out.extend(node["subs"])
        for child in sorted(node["children"], key=key):
            walk(child)
    for child in sorted(root["children"], key=key):
        walk(child)
    if issues is not None:
        check_sums(nodes, cols, table, issues)
    return out


def check_sums(nodes, cols, table, issues):
    """Сверка иерархии: значение строки = сумма строк уровнем ниже (подстроки «из них» не считаются).
    По умолчанию только деньги: штуки в отчётах часто уникальные (пациенты), родитель законно ≠ сумме."""
    mode = str(table.get("check", "money")).lower()
    if mode in ("off", "false", "0", "none"):
        return
    kinds = ("money",) if mode == "money" else ("money", "count")
    metrics = [c for c in cols if c.is_metric and c.kind in kinds and not c.formula and not c.nosum]
    for node in nodes:
        if not node["children"]:
            continue
        for c in metrics:
            parent = c.values[node["i"]]
            if not isinstance(parent, (int, float)):
                continue
            total = sum(v for v in (c.values[ch["i"]] for ch in node["children"]) if isinstance(v, (int, float)))
            if abs(parent - total) > max(0.011, abs(parent) * 1e-6):
                issues.append((node["i"], c, parent, total))


def flat_order(cols, sort, file_name):
    name = sort.lstrip("-+")
    col = next((c for c in cols if c.name.lower() == name.lower()), None)
    if col is None:
        print(f"   [!] {file_name}: нет колонки для сортировки «{name}» — порядок CSV")
        return list(range(len(cols[0].values)))
    desc = sort.startswith("-")

    def key(i):
        v = col.values[i]
        return (v is None, (-v if desc and isinstance(v, (int, float)) else v) if v is not None else 0)
    if not col.is_metric:
        idx = sorted(range(len(col.values)), key=lambda i: ("" if col.values[i] is None else str(col.values[i])),
                     reverse=desc)
        return idx
    return sorted(range(len(col.values)), key=key)


# ---------------------------------------------------------------- оформление
class Styler:
    def __init__(self, style):
        self.s = style
        self._fonts, self._fills, self._align = {}, {}, {}
        self._styles, self._borders = {}, {}

    def stamp(self, cell, font, align, border, fill=None, fmt="General"):
        """Стиль ячейки целиком из кэша: openpyxl не пересчитывает хеши шрифта/заливки/рамки на каждой ячейке
        (на больших листах — в десятки раз быстрее)."""
        key = (id(font), id(align), id(border), id(fill), fmt)
        tpl = self._styles.get(key)
        if tpl is None:
            cell.font, cell.alignment, cell.border = font, align, border
            if fill is not None:
                cell.fill = fill
            cell.number_format = fmt
            self._styles[key] = copy.copy(cell._style)
        else:
            cell._style = copy.copy(tpl)

    def border(self, left=False, right=False, bottom=False):
        """Рамка ячейки данных: внутренние линии + внешняя рамка таблицы на краях (если в палитре есть)."""
        key = (left, right, bottom)
        if key not in self._borders:
            if not P["outer_border"] or not (left or right or bottom):
                self._borders[key] = DATA_BORDER
            else:
                outer = Side(style="medium", color=P["outer_border"])
                b = DATA_BORDER
                self._borders[key] = Border(left=outer if left else b.left, right=outer if right else b.right,
                                            top=b.top, bottom=outer if bottom else b.bottom)
        return self._borders[key]

    def font(self, size=None, bold=False, color=None, italic=False, underline=None, name=None):
        color = color or P["text"]
        key = (size, bold, color, italic, underline, name)
        if key not in self._fonts:
            self._fonts[key] = Font(name=name or self.s["font"], size=size or self.s["size"], bold=bold, color=color,
                                    italic=italic, underline=underline)
        return self._fonts[key]

    def fill(self, color):
        if color not in self._fills:
            self._fills[color] = PatternFill("solid", fgColor=color)
        return self._fills[color]

    def align(self, h="center", indent=0, wrap=False):
        key = (h, indent, wrap)
        if key not in self._align:
            self._align[key] = Alignment(horizontal=h, vertical="center", indent=indent, wrap_text=wrap)
        return self._align[key]

    def head(self, ws, row, col, value=None, fill=None, color=None, size=None):
        cell = put(ws, row, col, value) if value is not None else ws.cell(row, col)
        cell.font = self.font(size, True, color or P["head_text"])
        cell.fill = self.fill(fill or P["head_fill"])
        cell.alignment = self.align("center", wrap=True)
        cell.border = HEAD_BORDER
        return cell


def put(ws, row, col, value):
    """Значение в ячейку: текст, начинающийся с «=», остаётся текстом (не формулой) — CSV не исполняется в Excel."""
    cell = ws.cell(row, col)
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"
        cell.quotePrefix = True
    return cell


def link(cell, sheet, st, size=11, text=None):
    if text is not None:
        cell.value = text
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1")
    cell.font = st.font(size, False, P["link"], underline="single")


def mix(color, share):
    rgb = [int(color[k:k + 2], 16) for k in (0, 2, 4)]
    return "".join(f"{round(255 + (x - 255) * share):02X}" for x in rgb)


def colors_for(col, coloring):
    positive = (col.positive_is or coloring["positive_is"]).lower()
    return (P["green"], P["red"]) if positive == "good" else (P["red"], P["green"])   # (рост, снижение)


def static_colors(values, col, coloring):
    """Мягкая шкала блока: 0 — белый, края — min/max блока (для Δ% — не выше pct_cap)."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return [None] * len(values)
    lo, hi = min(min(nums), 0.0), max(max(nums), 0.0)
    if col.kind in ("delta_pct", "pct"):
        hi = min(hi, coloring["pct_cap"])
    grow, fall = colors_for(col, coloring)
    gamma = coloring["gamma"]
    out = []
    for v in values:
        if not isinstance(v, (int, float)) or v == 0:
            out.append(None)
        elif v > 0 and hi > 0:
            out.append(mix(grow, min(v / hi, 1.0) ** gamma))
        elif v < 0 and lo < 0:
            out.append(mix(fall, min(v / lo, 1.0) ** gamma))
        else:
            out.append(None)
    return [None if c == CF_WHITE else c for c in out]


def colored_columns(rep, coloring):
    wanted = coloring.get("columns")
    out = []
    for c in rep.metrics:
        if c.colorize is not None:
            ok = bool(c.colorize)
        elif wanted:
            ok = any(pattern_match(p, c.name) for p in wanted)
        else:
            ok = c.kind in ("delta", "delta_pct")
        if ok:
            out.append(c)
    return out


class Grid:
    """Ячейки листа. Обычный режим — прямо в лист openpyxl. Потоковый (write-only, для больших книг) — ячейки
    копятся по строкам и уходят в файл по порядку (flush), после чего память освобождается; вернуться к уже
    записанной строке нельзя, поэтому всё, что выше данных, готовится заранее."""

    def __init__(self, ws):
        self.ws = ws
        self.stream = isinstance(ws, WriteOnlyWorksheet)
        self.rows, self.written = {}, 0

    def cell(self, row, column, value=None):
        if not self.stream:
            return self.ws.cell(row, column) if value is None else self.ws.cell(row, column, value)
        if row <= self.written:
            raise RuntimeError(f"строка {row} уже записана в потоковом режиме")
        cells = self.rows.setdefault(row, {})
        cell = cells.get(column)
        if cell is None:
            cell = cells[column] = WriteOnlyCell(self.ws)
            cell.row, cell.column = row, column          # координата — для гиперссылок
        if value is not None:
            cell.value = value
        return cell

    def __getitem__(self, coord):                        # g["A1"]
        return self.cell(*coordinate_to_tuple(coord))

    def __setitem__(self, coord, value):                 # g["A1"] = …
        self.cell(*coordinate_to_tuple(coord)).value = value

    def merge(self, r1, c1, r2, c2):
        if self.stream:
            self.ws.merged_cells.add(f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}")
        else:
            self.ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    def flush(self, upto=None):
        """Записать строки до upto включительно (по умолчанию — все накопленные)."""
        if not self.stream:
            return
        upto = max(self.rows, default=self.written) if upto is None else upto
        for r in range(self.written + 1, upto + 1):
            cells = self.rows.pop(r, {})
            self.ws.append([cells.get(c) for c in range(1, max(cells, default=0) + 1)])
        self.written = max(self.written, upto)

    # то, что у листа одинаково в обоих режимах (freeze_panes, sheet_view…), — прямо с листа: и чтение, и запись
    _own = ("ws", "stream", "rows", "written")

    def __getattr__(self, name):
        return getattr(self.ws, name)

    def __setattr__(self, name, value):
        if name in Grid._own:
            object.__setattr__(self, name, value)
        else:
            setattr(self.ws, name, value)


def write_sheet(ws, rep, fc, cfg, st):
    g = Grid(ws)
    style, table, coloring = cfg["style"], fc["table"], fc["coloring"]
    size = style["size"]
    cols = rep.columns
    offset = 1 if rep.row_numbers else 0
    ncols = len(cols) + offset
    col_idx = {id(c): offset + 1 + j for j, c in enumerate(cols)}
    letter = {id(c): get_column_letter(col_idx[id(c)]) for c in cols}
    n = rep.n_rows

    # --- строки 1–4: ссылка, заголовок, подзаголовок, описание
    link(g.cell(1, 1), TOC_SHEET, st, 10, "← Содержание")
    title_col = 1 + offset
    g.cell(2, title_col, fc.get("title") or rep.title).font = st.font(16, True, P["title"], name=P["title_font"])
    subtitle = fc.get("subtitle", cfg["book"].get("subtitle"))
    if subtitle:
        g.cell(3, title_col, subtitle).font = st.font(size, True, P["subtitle"], name=P["title_font"])
    g.cell(4, title_col, rep.description).font = st.font(size, False, P["muted"], name=P["title_font"])
    for r in (2, 3, 4):
        g.cell(r, title_col).alignment = st.align("left")

    totals_on = not rep.hierarchy and str(table.get("totals", "auto")).lower() not in ("off", "false", "0") \
        and any(c.additive for c in rep.metrics)
    first_col_text = rep.name_col.values if rep.name_col is not None else [None] * n
    is_total = [isinstance(v, str) and bool(TOTAL_RE.match(v)) for v in first_col_text] if not rep.hierarchy else [False] * n
    if str(table.get("totals", "auto")).lower() == "auto" and any(is_total):
        totals_on = False
    # строки: шапка -> номера колонок -> «Итого» -> данные. Автофильтр встаёт на «Итого» (строку заголовков
    # фильтра Excel не фильтрует и не сортирует) — итог всегда на месте, фильтруются только данные
    two_level = any(c.group for c in cols)
    head1, head2 = 5, 5 + (1 if two_level else 0)
    numbers_row = head2 + 1 if table.get("column_numbers", True) else None
    total_row = (numbers_row or head2) + 1 if totals_on else None
    first_data = (total_row or numbers_row or head2) + 1
    last_data = first_data + n - 1

    # --- шапка
    if rep.row_numbers:
        for r in range(head1, head2 + 1):
            st.head(g, r, 1, size=10)
        g.cell(head1, 1).value = "№"
        if two_level:
            g.merge(head1, 1, head2, 1)
    j = 0
    while j < len(cols):
        c = cols[j]
        ci = col_idx[id(c)]
        kind = "service" if c.role in ("service", "subrow") else (c.kind if c.is_metric else "dimension")
        fill = head_fill(kind)
        color = P["muted"] if c.role in ("service", "subrow") else P["head_text"]
        hsize = 11 if c.role in ("service", "subrow") else (size if len(str(c.label)) <= 60 else 12)
        if two_level and c.group:
            k = j
            while k + 1 < len(cols) and cols[k + 1].group == c.group:
                k += 1
            gfill = P["group_fill_delta"] if str(c.group).startswith("Δ") else P["group_fill_year"]
            for m in range(j, k + 1):
                st.head(g, head1, col_idx[id(cols[m])], fill=gfill)
                mk = cols[m].kind
                st.head(g, head2, col_idx[id(cols[m])], cols[m].label, head_fill(mk),
                        P["head_delta_text"] if mk in ("delta", "delta_pct") else P["head_text"])
            g.cell(head1, ci).value = c.group
            if k > j:
                g.merge(head1, ci, head1, col_idx[id(cols[k])])
            j = k + 1
            continue
        for r in range(head1, head2 + 1):
            st.head(g, r, ci, fill=fill, color=color, size=hsize)
        put(g, head1, ci, c.label)
        if two_level and c.role not in ("service", "subrow"):   # служебные — без объединения, скрываются по одной
            g.merge(head1, ci, head2, ci)
        elif two_level:
            g.cell(head1, ci).value = "служебная"
            g.cell(head1, ci).font = st.font(9, True, P["muted"])
            put(g, head2, ci, c.label)
        j += 1
    if numbers_row:
        for ci in range(1, ncols + 1):
            cell = g.cell(numbers_row, ci, "+" if (ci == 1 and rep.row_numbers) else ci - offset)
            cell.font = st.font(10, False, P["numbers"])
            cell.alignment = st.align("center")
            cell.border = DATA_BORDER

    # --- подготовка данных
    levels = rep.level.values if rep.hierarchy else [None] * n
    is_sub = subrow_flags(cols, n) if rep.hierarchy else [False] * n
    if rep.hierarchy:
        next_main, nxt = [None] * n, None
        for i in range(n - 1, -1, -1):
            next_main[i] = nxt
            if not is_sub[i]:
                nxt = i
        is_leaf = [is_sub[i] or next_main[i] is None or levels[next_main[i]] <= levels[i] for i in range(n)]
    else:
        is_leaf = [not t for t in is_total]
    sub_font = st.font(size - 2, False, P["subrow"], italic=True)
    service_font = st.font(11, False, P["service"])
    no_font = st.font(10, False, P["numbers"])
    key_metric = main_metric(rep.metrics, table)[0] if P["key_column"] else None
    key_fill = st.fill(P["key_column"]) if P["key_column"] else None
    pct_cap = coloring["pct_cap"]
    live = str(table.get("formulas", "values")).lower() == "live"
    tech = []          # (колонка, ячейки, формула/правило, пояснение) -> лист «Формулы» и примечания к заголовкам
    head_cell = {id(c): g.cell(head2 if (two_level and c.group) else head1, col_idx[id(c)]) for c in cols}

    # раскраска Δ / Δ% (заливкой) — заранее, чтобы стиль каждой ячейки записать один раз
    mode = str(coloring.get("mode", "static")).lower()
    targets = colored_columns(rep, coloring) if mode != "off" else []
    blocks = color_blocks(rep, levels, is_leaf, is_sub, n)
    color_at, n_colored = {}, 0
    if mode != "conditional":
        for c in targets:
            per_row = color_at.setdefault(id(c), {})
            for block in blocks:
                for i, color in zip(block, static_colors([c.values[i] for i in block], c, coloring)):
                    if color:
                        per_row[i] = st.fill(color)
                        n_colored += 1

    # --- строка «Итого» под номерами колонок (до данных: в потоковом режиме назад не вернуться)
    if total_row:
        label_col = col_idx[id(cols[0])]
        if rep.row_numbers:
            label_col = 1
        totals = {}
        for c in cols:
            if c.additive and not c.formula:
                totals[id(c)] = round(sum(v for v in c.values if isinstance(v, (int, float))), 6)
        for c in cols:
            if c.formula:
                op, a, b = c.formula
                x, y = totals.get(id(a)), totals.get(id(b))
                totals[id(c)] = (None if x is None or y is None else round(x - y, 6) if op == "diff"
                                 else (x / y - 1 if y else None))
        for c in cols:
            ci = col_idx[id(c)]
            cell = g.cell(total_row, ci)
            L = letter[id(c)]
            if c.additive and not c.formula:
                text = f"=SUBTOTAL(109,{L}{first_data}:{L}{last_data})"
                cell.value = text if live else totals[id(c)]
                tech.append((c, f"{L}{total_row}", text, "строка «Итого»: сумма по видимым (после фильтра) строкам"))
            elif c.formula:
                text = formula_text(c.formula, total_row, letter)
                cell.value = text if live else totals[id(c)]
                tech.append((c, f"{L}{total_row}", text, "строка «Итого»"))
            over_cap = c.cap and not live and isinstance(totals.get(id(c)), (int, float)) and totals[id(c)] > pct_cap
            cell.number_format = FMT_PCT_CAP if over_cap else (c.fmt or "General")
            cell.font = st.font(None, True)
            cell.fill = st.fill(P["total"])
            cell.border = DATA_BORDER
            cell.alignment = st.align("center")
        if not rep.row_numbers:   # подпись — в первую свободную ячейку (первая колонка может быть метрикой)
            label_col = next((col_idx[id(c)] for c in cols if g.cell(total_row, col_idx[id(c)]).value is None),
                             label_col)
        cell = g.cell(total_row, label_col)
        if cell.value is None:
            cell.value = "Итого (по видимым строкам)" if live else "Итого"
            cell.alignment = st.align("left")
        if rep.row_numbers:
            cell.fill, cell.font, cell.border = st.fill(P["total"]), st.font(10, True), DATA_BORDER
        g.row_dimensions[total_row].height = 18.75 * size / 14

    # --- формулы колонок: добавленные Δ / Δ% и подсказки для Δ из CSV; правила раскраски
    for c in cols:
        spec = c.formula or c.hint
        if spec and n:
            L = letter[id(c)]
            note = ("добавлено скриптом" if c.formula else "в ячейках — значения из CSV; формула — для проверки/правки")
            if c.cap:
                note += f"; формат «> 200%» для значений больше {pct_cap:.0%}: {FMT_PCT_CAP}"
            tech.append((c, f"{L}{first_data}:{L}{last_data}", formula_text(spec, first_data, letter),
                         note + f"; вставить в {L}{first_data} и протянуть вниз"))
    for c in targets:
        grow, fall = colors_for(c, coloring)
        cells, n_ranges = ranges_text(letter[id(c)], first_data, blocks)
        rule = (f"Цветовая шкала (3 цвета): минимум — #{fall}, число 0 — #FFFFFF, максимум — #{grow}"
                + (f"; для Δ% максимум не выше {coloring['pct_cap']:.0%}" if c.kind in ("pct", "delta_pct") else ""))
        tech.append((c, cells, rule,
                     f"раскраска: {n_ranges} диапазон(ов), шкала отдельно для каждого"
                     + ("" if mode == "conditional" else
                        f"; в файле цвет записан заливкой, насыщенность = доля^{coloring['gamma']} "
                        "(средние значения бледнее, чем в шкале Excel)")))
        if mode == "conditional":
            for block in blocks:
                for a, b in contiguous(block):
                    ref = f"{letter[id(c)]}{first_data + a}:{letter[id(c)]}{first_data + b}"
                    g.conditional_formatting.add(ref, ColorScaleRule(
                        start_type="min", start_color=fall, mid_type="num", mid_value=0, mid_color=CF_WHITE,
                        end_type="max", end_color=grow))
                    n_colored += b - a + 1
    rep.tech = tech
    if cfg["book"].get("header_notes", True):
        by_col = {}
        for c, cells, text, note in tech:
            by_col.setdefault(id(c), []).append(f"{cells}: {text}\n({note})")
        for c in cols:
            if id(c) in by_col:
                body = "\n\n".join(by_col[id(c)])
                head_cell[id(c)].comment = Comment(body, "report_builder", width=420,
                                                   height=min(400, 60 + 18 * body.count("\n") + len(body) // 3))

    # --- внешняя рамка шапки (палитра B: внешняя #9AA5B1, внутренние #D3D9DF); у данных — в цикле ниже
    if P["outer_border"] and n:
        frame(g, head1, first_data - 1, 1, ncols, P["outer_border"], bottom_edge=False)

    # --- ширины (в потоковом режиме — до первой строки)
    widths = {}
    scale = size / 14
    for c in cols:
        ci = col_idx[id(c)]
        if c.width:
            widths[ci] = c.width
            continue
        sample = c.values[:3000]
        if c.is_metric:
            longest = max((len(shown(v, c)) for v in sample if isinstance(v, (int, float))), default=4)
            if total_row and c.additive:   # строка «Итого» — сумма длиннее любого значения
                longest = max(longest, len(shown(sum(abs(v) for v in c.values if isinstance(v, (int, float))), c)) + 1)
        elif c.kind in ("date", "datetime"):
            longest = 10 if c.kind == "date" else 16
        else:
            longest = max((len(str(v)) for v in sample if v is not None), default=4)
        if c.role == "name":
            longest = max((len(str(v)) * 1.0 + 2.5 * (lv or 0) for v, lv in zip(sample, levels[:3000])), default=40)
            top = table.get("name_max_width", 90)
            w = min(top, max(40, longest * CHAR_W * scale + 2))
        else:
            label_word = max((len(w) for w in str(c.label).split()), default=4)
            top = table.get("max_width", 45) if c.role != "service" else 24
            w = min(top, max(8, longest * CHAR_W * scale + 2.5, label_word * HEAD_CHAR_W * scale + 2))
        widths[ci] = round(w, 1)
    if rep.row_numbers:
        widths[1] = max(5, len(str(n)) * 1.2 + 2)
    for ci, w in widths.items():
        g.column_dimensions[get_column_letter(ci)].width = w
    for c in cols:
        if c.hidden:
            g.column_dimensions[letter[id(c)]].hidden = True
    # высота шапки — по самому длинному переносу
    lines = max((wrapped_lines(c.label, (widths[col_idx[id(c)]] - 1) / (HEAD_CHAR_W * scale)) for c in cols), default=1)
    lines = min(lines, 6)
    if two_level:
        g.row_dimensions[head1].height = 21 * scale
        g.row_dimensions[head2].height = max(21, 18 * lines) * scale
    else:
        g.row_dimensions[head1].height = max(21, 18 * lines) * scale
    if numbers_row:
        g.row_dimensions[numbers_row].height = 13.5

    # --- вид, закрепление, фильтр, печать (в потоковом режиме — до первой строки)
    g.sheet_view.showGridLines = False
    if style.get("view") == "page_break":
        g.sheet_view.view = "pageBreakPreview"
        g.sheet_view.zoomScaleSheetLayoutView = style["zoom"]
    g.sheet_view.zoomScale = style["zoom"]
    g.sheet_view.zoomScaleNormal = style["zoom"]
    fcols = table.get("freeze_columns", "auto")
    if fcols == "auto":
        lead = 0
        for c in cols:
            if c.is_metric or c.role in ("service", "subrow"):
                break
            lead += 1
        fcols = lead if 0 < lead <= 3 and sum(widths[col_idx[id(c)]] for c in cols[:lead]) <= 100 else (1 if lead else 0)
    g.freeze_panes = f"{get_column_letter(offset + 1 + int(fcols))}{first_data}"
    filter_row = total_row or numbers_row or head2
    if not rep.hierarchy and table.get("autofilter", True) and n:
        g.auto_filter.ref = f"A{filter_row}:{get_column_letter(ncols)}{last_data}"
    total_width = sum(w for ci, w in widths.items() if not g.column_dimensions[get_column_letter(ci)].hidden)
    paper = style.get("paper", "auto")
    print_setup(ws, style, f"{head1}:{filter_row}", paper == "A3" or (paper == "auto" and total_width > 180))
    g.sheet_format.defaultRowHeight = 18.75 * size / 14      # высота строк данных — по умолчанию листа
    g.sheet_format.customHeight = True
    check_cells = getattr(rep, "check_cells", set())
    alert_font = st.font(None, True, P["alert"])
    outline = rep.hierarchy and bool(table.get("outline", True))
    if outline:                       # группировка строк: кнопки +/− над группой (итог группы — сверху)
        g.sheet_properties.outlinePr.summaryBelow = False
    g.flush(first_data - 1)           # шапка и «Итого» готовы — в файл

    # --- данные: строка за строкой (в потоковом режиме каждая строка сразу уходит в файл)
    align_center, align_left = st.align("center"), st.align("left")
    first_ci, last_ci = 1, ncols
    for i in range(n):
        r = first_data + i
        lvl = levels[i]
        last = i == n - 1
        if outline and lvl:
            g.row_dimensions[r].outline_level = min(int(lvl), 7)
        if rep.hierarchy:
            fill = level_fill(lvl) if (not is_leaf[i] or lvl == 0) else None
            bold = lvl <= 1
        else:
            fill = P["total"] if is_total[i] else None
            bold = is_total[i]
        sub = is_sub[i]
        fill_obj = st.fill(fill) if fill else None
        if rep.row_numbers:
            st.stamp(g.cell(r, 1, i + 1), no_font, align_center, st.border(left=True, bottom=last))
        for c in cols:
            ci = col_idx[id(c)]
            value = formula_text(c.formula, r, letter) if (c.formula and live) else c.values[i]
            cell = put(g, r, ci, value) if not (c.formula and live) else g.cell(r, ci, value)
            if c.role in ("service", "subrow"):
                font, align = service_font, align_left
            elif c.role == "name":
                font, align = (sub_font if sub else st.font(None, bold)), st.align("left", indent=min(lvl or 0, 15))
            elif c.is_metric:
                font, align = (sub_font if sub else st.font(None, bold)), align_center
            else:
                font, align = st.font(None, bold), (align_left if c.kind == "text" else align_center)
            fmt = c.fmt or "General"
            if c.cap and isinstance(c.values[i], (int, float)) and c.values[i] > pct_cap:
                fmt = FMT_PCT_CAP
            fill = fill_obj
            if fill is None:
                fill = color_at.get(id(c), {}).get(i) or (key_fill if (c is key_metric and not sub) else None)
            border = st.border(left=ci == first_ci, right=ci == last_ci, bottom=last)
            if (i, id(c)) in check_cells:          # сверка: строка ≠ сумме строк ниже
                font = alert_font
            st.stamp(cell, font, align, border, fill, fmt)
        g.flush(r)
    g.flush()
    return n_colored, len(blocks)


def frame(ws, top, bottom, left, right, color, bottom_edge=True):
    side = Side(style="medium", color=color)
    for r in range(top, bottom + 1):
        for ci in (left, right):
            cell = ws.cell(r, ci)
            b = cell.border
            cell.border = Border(left=side if ci == left else b.left, right=side if ci == right else b.right,
                                 top=b.top, bottom=b.bottom)
    for ci in range(left, right + 1):
        for r in (top, bottom) if bottom_edge else (top,):
            cell = ws.cell(r, ci)
            b = cell.border
            cell.border = Border(left=b.left, right=b.right, top=side if r == top else b.top,
                                 bottom=side if (r == bottom and bottom_edge) else b.bottom)


def shown(v, c):
    """Как число будет выглядеть в ячейке — для ширины колонки."""
    if c.kind in ("pct", "delta_pct"):
        return "> 200%" if c.cap and v > 2 else f"+{v:.1%}"
    if c.kind == "pct_points":
        return f"{v:.1f}%"
    if c.kind == "money" or (c.fmt and "0.00" in c.fmt):
        return f"{v:,.2f}"
    return f"{v:,.0f}"


def wrapped_lines(text, per_line):
    """Сколько строк займёт подпись при переносе по словам."""
    per_line = max(int(per_line), 1)
    lines, cur = 1, 0
    for word in str(text).split():
        need = len(word) if cur == 0 else cur + 1 + len(word)
        if need <= per_line:
            cur = need
        else:
            lines += 1 if cur else 0
            cur = len(word)
    return lines


def formula_text(spec, r, letter):
    op, a, b = spec
    la, lb = letter[id(a)], letter[id(b)]
    return f"={la}{r}-{lb}{r}" if op == "diff" else f'=IF(N({lb}{r})=0,"",{la}{r}/{lb}{r}-1)'


def ranges_text(col_letter, first, block_list, limit=40):
    parts = [f"{col_letter}{first + a}" if a == b else f"{col_letter}{first + a}:{col_letter}{first + b}"
             for block in block_list for a, b in contiguous(block)]
    return "; ".join(parts[:limit]) + (f"; … ещё {len(parts) - limit}" if len(parts) > limit else ""), len(parts)


def color_blocks(rep, levels, is_leaf, is_sub, n):
    """Иерархия: подряд идущие строки-детали одного уровня — отдельная шкала; подстроки «из них» — своя шкала.
    Плоская таблица: все строки, кроме «Всего/Итого», — одна шкала."""
    if not rep.hierarchy:
        return [[i for i in range(n) if is_leaf[i]]]
    blocks, current = [], []
    for i in range(n + 1):
        if i < n and is_sub[i]:
            continue
        detail = i < n and is_leaf[i] and levels[i] > 0
        if current and (not detail or levels[i] != levels[current[0]]):
            blocks.append(current)
            current = []
        if detail:
            current.append(i)
    subs = [i for i in range(n) if is_sub[i]]
    if subs:
        blocks.append(subs)
    return [b for b in blocks if b]


def contiguous(block):
    out, start, prev = [], None, None
    for i in block:
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            out.append((start, prev))
            start = prev = i
    if start is not None:
        out.append((start, prev))
    return out


# ---------------------------------------------------------------- книга
def natural_key(text):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", text)]


def report_number(stem):
    """-> (номер, имя без номера): «rep_x_1_1_202610021617» -> («1.1», «rep x»); «19.1 ЖК» -> («19.1», «ЖК»);
    без номера -> ('', имя). Хвост-штамп даты выгрузки DBeaver (_202610021617) отбрасывается."""
    stem = re.sub(r"_\d{8,}$", "", stem)
    m = re.match(r"^(\d+(?:[._]\d+)*)(?=\D|$)", stem)
    if m:
        return m.group(1).replace("_", "."), humanize(stem[m.end():]).strip(" .-")
    m = re.search(r"_(\d{1,3})(?:_(\d{1,2}))?$", stem)
    if m:
        return m.group(1) + (f".{m.group(2)}" if m.group(2) else ""), humanize(stem[:m.start()])
    return "", humanize(stem)


def default_sheet_name(file_name, folder_files):
    """Имя листа по умолчанию: «номер + имя файла»; если имя без номера одинаковое у всех файлов папки
    с номером (rep_app_stac_1_…, rep_app_stac_1_1_…) — оно ничего не различает: «Отчёт 1.1»."""
    no, rest = report_number(os.path.splitext(file_name)[0])
    if not no:
        return rest
    numbered = [report_number(os.path.splitext(f)[0]) for f in folder_files]
    rests = {r for n, r in numbered if n}            # сравниваются только файлы с номером отчёта
    return f"Отчёт {no}" if sum(1 for n, _ in numbered if n) > 1 and rests == {rest} else f"{no} {rest}"


def sheet_title(name, used):
    base = re.sub(r"[\[\]:*?/\\]", " ", str(name)).strip().strip("'")[:31] or "Отчёт"
    title, k = base, 2
    while title.lower() in used:
        suffix = f" ({k})"
        title, k = base[:31 - len(suffix)] + suffix, k + 1
    used.add(title.lower())
    return title


def humanize(stem):
    stem = re.sub(r"_\d{8,}$", "", stem)
    return re.sub(r"[_]+", " ", stem).strip()


def discover(src_dir, cfg):
    files = [f for f in os.listdir(src_dir) if f.lower().endswith(".csv")]
    out = []
    for f in files:
        fc = file_config(cfg, f)
        if fc.get("skip"):
            print(f"[пропуск] {f}: skip в конфиге")
            continue
        stem = os.path.splitext(f)[0]
        no, _ = report_number(stem)
        order = fc.get("order")
        if order is not None:      # order: «1.4.1» — сразу после отчёта 1.4; «0» — в начало
            no = str(order)
        key = ((0, tuple(int(x) for x in re.findall(r"\d+", no)), natural_key(f)) if no
               else (1, (), natural_key(f)))
        out.append((key, f, fc))
    return [(f, fc) for _, f, fc in sorted(out, key=lambda x: x[0])]


def structure_text(rep):
    if rep.hierarchy:
        return "Всего → " + " → ".join(p.strip() for p in str(rep.name_col.label).split("/"))
    dims = [c.label for c in rep.columns if not c.is_metric and c.role not in ("service", "subrow", "level")]
    return ", ".join(dims) or "—"


def key_total(rep, table):
    metric, _ = main_metric(rep.metrics, table)
    if metric is None:
        return None, None
    if rep.hierarchy:
        zero = [v for v, lv in zip(metric.values, rep.level.values) if lv == 0]
        return metric, (zero[0] if zero else None)
    name_vals = rep.name_col.values if rep.name_col is not None else [None] * rep.n_rows
    totals = [v for v, nv in zip(metric.values, name_vals) if isinstance(nv, str) and TOTAL_RE.match(nv)]
    if totals:
        return metric, totals[0]
    return metric, sum(v for v in metric.values if isinstance(v, (int, float)))


def build(src_dir, cfg, out_path, explain=False):
    apply_palette(cfg["style"])
    st = Styler(cfg["style"])
    book = cfg["book"]
    files = discover(src_dir, cfg)
    if not files:
        sys.exit(f"В папке {src_dir} нет CSV-файлов")
    all_files = [f for f, _ in files]
    period = book.get("period") or ""
    # потоковая запись (write-only): строки сразу уходят в файл — память не растёт с размером данных;
    # auto — если CSV в сумме больше STREAM_BYTES
    stream = book.get("stream", "auto")
    if str(stream).lower() == "auto":
        stream = sum(os.path.getsize(os.path.join(src_dir, f)) for f in all_files) > STREAM_BYTES
    stream = bool(stream) and str(stream).lower() not in ("false", "off", "0")
    wb = Workbook(write_only=stream)
    if stream:
        toc = wb.create_sheet(TOC_SHEET)
        print("Потоковая запись: большие файлы — строки пишутся сразу в файл (меньше памяти).")
    else:
        toc = wb.active
        toc.title = TOC_SHEET
    used = {TOC_SHEET.lower(), HELP_SHEET.lower(), FORMULAS_SHEET.lower()}
    entries, skipped, help_rows = [], [], {}

    for f, fc in files:
        path = os.path.join(src_dir, f)
        stem = os.path.splitext(f)[0]
        try:
            rep = prepare(path, f, fc, period)
        except Exception as err:   # плохой файл не должен ронять сборку остальных
            print(f"[пропуск] {f}: {err}")
            skipped.append((f, str(err)))
            continue
        no, rest = report_number(stem)
        rep.title = fc.get("title") or book.get("title") or os.path.basename(src_dir.rstrip("\\/"))
        rep.description = fc.get("description") or (f"Отчёт {no}. {rest}" if no else rest)
        if fc.get("description") and no and not fc.get("description").startswith("Отчёт"):
            rep.description = f"Отчёт {no}. {fc['description']}"
        roles = ", ".join(f"{c.name}={c.role}{'/' + c.kind if c.is_metric else ''}{'(+)' if c.formula else ''}"
                          for c in rep.columns)
        if explain:
            print(f"[{'иерархия' if rep.hierarchy else 'таблица '}] {f}  ({rep.encoding}, разделитель {sep_name(rep.sep)}, строк {rep.n_rows})")
            print(f"     {roles}")
            continue
        name = sheet_title(fc.get("sheet") or default_sheet_name(f, all_files), used)
        cells = rep.n_rows * len(rep.columns)
        if cells > BIG_SHEET_CELLS:   # замер: 200 тыс. строк × 18 колонок ≈ 1,5 мин (потоком — ~0,4 ГБ памяти)
            print(f"   [!] {f}: большой лист ({rep.n_rows:,} строк × {len(rep.columns)} колонок) — сборка займёт "
                  f"~{max(1, round(cells / 2_500_000 * 1.5))} мин"
                  + ("" if stream else f" и ~{cells / 2_000_000:.1f} ГБ памяти (потоковая запись: book.stream: true)")
                  + "; для отчёта лучше агрегировать данные".replace(",", " "))
        ws = wb.create_sheet(name)
        try:
            n_colored, n_blocks = write_sheet(ws, rep, fc, cfg, st)
        except Exception as err:   # сбой записи листа — лист убираем, остальные файлы собираем
            wb.remove(ws)
            used.discard(name.lower())
            reason = f"ошибка при записи листа: {type(err).__name__}: {err}"
            print(f"[пропуск] {f}: {reason}")
            skipped.append((f, reason))
            continue
        print(f"[ok] {name:31} строк={rep.n_rows:7}  метрик={len(rep.metrics):3}  окрашено={n_colored:7}  <- {f}")
        for note in rep.notes:
            print(f"     ! {note}")
        entries.append((name, rep))
        for c in rep.columns:
            h = help_rows.setdefault(c.name, {"label": c.label, "role": c.role, "kind": c.kind, "fmt": c.fmt,
                                              "description": c.description, "sheets": []})
            h["sheets"].append(name)
    if explain:
        return
    if not entries:
        sys.exit("Ни один CSV не подошёл — книга не сохранена")

    write_toc(toc, entries, skipped, cfg, src_dir, st)
    if book.get("columns_sheet", True):
        write_help(wb.create_sheet(HELP_SHEET), help_rows, st, cfg["style"])
    mode = str(book.get("formulas_sheet", "hidden")).lower()
    if mode != "off" and any(rep.tech for _, rep in entries):
        fws = wb.create_sheet(FORMULAS_SHEET)
        write_formulas(fws, entries, st, cfg["style"])
        if mode == "hidden":
            fws.sheet_state = "hidden"
    wb.save(out_path)
    print(f"Готово: {out_path}")


def write_toc(ws, entries, skipped, cfg, src_dir, st):
    ws = Grid(ws)                     # обычный или потоковый лист — одинаково
    book = cfg["book"]
    ws.sheet_view.showGridLines = False
    title = book.get("title") or os.path.basename(src_dir.rstrip("\\/"))
    ws["A1"] = title
    ws["A1"].font = st.font(16, True, P["title"], name=P["title_font"])
    row = 2
    if book.get("subtitle"):
        ws.cell(row, 1, book["subtitle"]).font = st.font(12, True, P["subtitle"], name=P["title_font"])
        row += 1
    ws.cell(row, 1, f"Источник: {src_dir}. Собрано {dt.datetime.now():%d.%m.%Y %H:%M}.").font = st.font(11, False, P["muted"])
    row += 1
    if str(book.get("formulas_sheet", "hidden")).lower() != "off":
        hidden = str(book.get("formulas_sheet", "hidden")).lower() == "hidden"
        ws.cell(row, 1, "В ячейках — значения (без формул и условного форматирования). Формулы Δ / Δ% / Итого и правила "
                        "раскраски — в примечаниях к заголовкам колонок и на листе «Формулы»"
                        + (" (скрыт: ПКМ по ярлыку листа → Показать)." if hidden else ".")).font = st.font(11, False, P["muted"])
        row += 1
    for note in book.get("notes") or []:
        ws.cell(row, 1, note).font = st.font(11, False, P["muted"])
        row += 1
    row += 1
    heads = ["№", "Лист", "Описание", "Структура", "Метрики", "Строк", "Ключевой итог", "Значение", "Файл", "Замечания"]
    for ci, h in enumerate(heads, start=1):
        st.head(ws, row, ci, h, size=11)
    head_row = row
    for k, (name, rep) in enumerate(entries, start=1):
        row += 1
        ws.cell(row, 1, k)
        link(ws.cell(row, 2), name, st, 11, name)
        put(ws, row, 3, rep.description)
        put(ws, row, 4, structure_text(rep))
        names = [c.name for c in rep.metrics if not c.formula]
        text = ", ".join(names)
        ws.cell(row, 5, (text if len(text) <= 120 else text[:117].rsplit(",", 1)[0] + f" … всего {len(names)}") or "—")
        ws.cell(row, 6, rep.n_rows).number_format = "#,##0"
        metric, value = key_total(rep, rep.table)
        ws.cell(row, 7, metric.name if metric else None)
        if value is not None:
            ws.cell(row, 8, value).number_format = metric.fmt or FMT["money"]
        put(ws, row, 9, rep.file)
        put(ws, row, 10, "; ".join(rep.notes) or None)
        for ci in range(1, len(heads) + 1):
            cell = ws.cell(row, ci)
            cell.border = DATA_BORDER
            if ci != 2:
                cell.font = st.font(11)
            cell.alignment = st.align("center" if ci in (1, 6, 8) else "left", wrap=ci in (3, 4, 5, 10))
            if ci == 10 and cell.value:
                cell.font = st.font(11, False, P["alert"])
    if skipped:
        row += 2
        ws.cell(row, 1, "Пропущенные файлы").font = st.font(11, True, P["alert"])
        for f, reason in skipped:
            row += 1
            ws.cell(row, 2, f).font = st.font(11)
            ws.cell(row, 3, reason).font = st.font(11, False, P["muted"])
    for col, width in zip("ABCDEFGHIJ", (5, 32, 60, 45, 45, 9, 24, 16, 40, 50)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = f"A{head_row + 1}"
    print_setup(ws, cfg["style"], f"{head_row}:{head_row}")
    ws.flush()


def print_setup(ws, style, title_rows=None, a3=False):
    if title_rows:
        ws.print_title_rows = title_rows
    ws.page_setup.paperSize = Worksheet.PAPERSIZE_A3 if a3 else Worksheet.PAPERSIZE_A4   # константы класса: у потокового листа их нет
    ws.page_setup.orientation = style.get("orientation", "landscape")
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = 0.12
    ws.page_margins.bottom, ws.page_margins.header, ws.page_margins.footer = 0.4, 0.1, 0.15
    ws.oddFooter.center.text = "Страница &P из &N"
    ws.oddFooter.center.size = 10
    ws.oddFooter.center.font = f"{style['font']},Regular"
    ws.oddFooter.center.color = P["muted"]


def write_formulas(ws, entries, st, style):
    """Формулы и правила раскраски текстом: скопировать из колонки «Формула / правило» в первую ячейку диапазона
    и протянуть вниз (ссылки относительные)."""
    ws = Grid(ws)                     # обычный или потоковый лист — одинаково
    ws.sheet_view.showGridLines = False
    link(ws.cell(1, 1), TOC_SHEET, st, 10, "← Содержание")
    ws.cell(2, 1, "Формулы и правила раскраски").font = st.font(16, True, P["title"], name=P["title_font"])
    ws.cell(3, 1, "В листах отчёта — только значения. Формулу из колонки D вставьте в первую ячейку диапазона и "
                  "протяните вниз; правило раскраски — Главная → Условное форматирование → Цветовые шкалы → "
                  "Другие правила.").font = st.font(11, False, P["muted"])
    heads = ["Лист", "Колонка", "Ячейки", "Формула / правило", "Пояснение"]
    for ci, h in enumerate(heads, start=1):
        st.head(ws, 5, ci, h, size=11)
    r = 5
    for name, rep in entries:
        for c, cells, text, note in rep.tech:
            r += 1
            link(ws.cell(r, 1), name, st, 11, name)
            put(ws, r, 2, c.name)
            ws.cell(r, 3, cells)
            cell = ws.cell(r, 4, text)
            cell.data_type, cell.quotePrefix = "s", True     # текст «=…», Excel его не вычисляет
            ws.cell(r, 5, note)
            for ci in range(1, 6):
                cell = ws.cell(r, ci)
                cell.border = DATA_BORDER
                if ci != 1:
                    cell.font = st.font(11, False, P["title"] if ci == 4 else P["text"])
                cell.alignment = st.align("left", wrap=ci in (3, 4, 5))
    for col, width in zip("ABCDE", (30, 30, 30, 60, 60)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:E{r}"
    print_setup(ws, style, "5:5")
    ws.flush()


def write_help(ws, rows, st, style):
    ws = Grid(ws)                     # обычный или потоковый лист — одинаково
    ws.sheet_view.showGridLines = False
    link(ws.cell(1, 1), TOC_SHEET, st, 10, "← Содержание")
    ws.cell(2, 1, "Справка по колонкам").font = st.font(16, True, P["title"], name=P["title_font"])
    roles = {"metric": "метрика", "dimension": "разрез", "name": "наименование", "service": "служебная",
             "level": "уровень", "subrow": "подстрока «из них»"}
    heads = ["Колонка CSV", "Подпись в шапке", "Роль", "Вид", "Описание", "Листы"]
    for ci, h in enumerate(heads, start=1):
        st.head(ws, 4, ci, h, size=11)
    r = 4
    for name, h in rows.items():
        r += 1
        values = [name, h["label"], roles.get(h["role"], h["role"]), h["kind"], h["description"], ", ".join(h["sheets"])]
        for ci, v in enumerate(values, start=1):
            cell = put(ws, r, ci, v)
            cell.font, cell.border = st.font(11), DATA_BORDER
            cell.alignment = st.align("left", wrap=ci in (5, 6))
    for col, width in zip("ABCDEF", (34, 30, 18, 12, 50, 60)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:F{r}"
    print_setup(ws, style, "4:4")
    ws.flush()


# ---------------------------------------------------------------- уточняющие вопросы
WIDE_COLUMNS = 30                        # шире — в вопросах сводка вместо полной таблицы колонок
TECH_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9]*(?:_[A-Za-z0-9]+)+$|^[A-Z][A-Z0-9]{2,}$")
class Asker:
    """Вопрос в консоли: Enter — вариант в скобках, «!» — принять все остальные варианты по умолчанию."""

    def __init__(self, inp=input, out=print):
        self.inp, self.out, self.auto = inp, out, False

    def _read(self, prompt):
        if self.auto:
            return ""
        try:
            answer = self.inp(prompt).strip()
        except EOFError:
            self.auto = True
            return ""
        if answer == "!":
            self.auto = True
            return ""
        return answer

    def text(self, question, default=""):
        answer = self._read(f"{question} [{default}]: ")
        return answer if answer else default

    def choice(self, question, options, default):
        self.out(question)
        keys = [k for k, _ in options]
        for i, (_, label) in enumerate(options, start=1):
            self.out(f"   {i}. {label}{'   <- сейчас' if keys[i - 1] == default else ''}")
        while True:
            answer = self._read(f"   номер [{keys.index(default) + 1}]: ")
            if not answer:
                return default
            if answer.isdigit() and 1 <= int(answer) <= len(options):
                return keys[int(answer) - 1]
            self.out(f"   нужен номер от 1 до {len(options)}")


def load_raw(path):
    """Конфиг как он записан в файле (без значений по умолчанию) — в него дописываются ответы."""
    if not path or not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        if path.lower().endswith(".json"):
            return json.load(f) or {}
        if yaml is None:
            sys.exit("Для конфига .yaml нужен pyyaml: pip install pyyaml (или используйте report.json)")
        return yaml.safe_load(f) or {}


KIND_CHOICES = [
    ("count", "штуки (случаи, пациенты, количество) — голубая шапка, суммируется"),
    ("money", "деньги — зелёная шапка, суммируется"),
    ("number", "другое число (балл, коэффициент, среднее) — бирюзовая, не суммируется"),
    ("pct", "процент / доля — бирюзовая, не суммируется"),
    ("delta", "Δ — изменение (разница), подсвечивается"),
    ("delta_pct", "Δ% — изменение в процентах, подсвечивается"),
    ("dimension", "разрез / код / текст — не метрика"),
    ("level", "уровень иерархии (0 — Всего, 1, 2 …)"),
    ("name", "наименование строки (в иерархии)"),
    ("subrow", "«из них» — подстрока под своей строкой"),
    ("skip", "не выводить в отчёт"),
]
ROLE_KINDS = ("dimension", "skip", "level", "name", "subrow")   # не метрики — пишутся как role
KIND_RU = {"count": "штуки", "money": "деньги", "number": "другое число", "pct": "процент", "pct_points": "процент",
           "delta": "Δ", "delta_pct": "Δ%", "date": "дата", "datetime": "дата и время", "text": "текст"}
ROLE_RU = {"level": "уровень иерархии", "name": "наименование", "subrow": "«из них»", "service": "служебная",
           "dimension": "разрез"}


def describe(c):
    if c.is_metric:
        return KIND_RU.get(c.kind, c.kind)
    if c.role == "dimension" and c.kind in ("date", "datetime"):
        return KIND_RU[c.kind]
    return ROLE_RU.get(c.role, c.role)


def column_flags(c, columns_rules=None):
    flags = []
    if c.formula:
        flags.append("добавлена скриптом")
    if c.kind_source == "values":
        flags.append("вид угадан по цифрам — проверьте")
    if c.maybe_code:
        flags.append("все значения разные — похоже на код")
    if not c.is_metric and 0.5 <= c.numeric_share < 0.98:
        flags.append(f"{c.numeric_share:.0%} чисел, есть текст: «{c.odd_value}» — может, метрика?")
    return flags


def doubts(rep, columns_rules):
    """Колонки, о которых скрипт не уверен (без «добавлена скриптом»), и технические имена без подписи."""
    unsure = [c for c in rep.columns if [f for f in column_flags(c) if f != "добавлена скриптом"]]
    tech = [c for c in rep.columns if TECH_NAME.match(c.name) and not rules_for(columns_rules, c.name).get("label")]
    return unsure, tech


def parse_numbers(text, upper):
    """«2,5 7-9» -> [2, 5, 7, 8, 9]; номера вне 1..upper — ошибка."""
    out = []
    for part in re.split(r"[,\s;]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
        nums = list(range(int(m.group(1)), int(m.group(2)) + 1)) if m else ([int(part)] if part.isdigit() else None)
        if nums is None or any(not 1 <= x <= upper for x in nums):
            raise ValueError(part)
        out += [x for x in nums if x not in out]
    return out


def ask_numbers(asker, question, upper, default_text):
    while True:
        answer = asker._read(f"{question} [{default_text}]: ")
        if not answer:
            return None
        if answer.strip() == "0":
            return []
        try:
            return parse_numbers(answer, upper)
        except ValueError as err:
            asker.out(f"   не понял «{err}»: номера от 1 до {upper} через запятую или диапазон 2-5; 0 — ничего")


def structure_key(rep):
    return (rep.hierarchy, tuple(c.name for c in rep.source_columns))


def plan_questions(src_dir, cfg, raw, force=False):
    """Анализ папки: файлы -> структуры (одинаковый набор колонок). Спрашивается только о структурах, у которых
    есть файлы без настроек в конфиге (force / --ask — обо всех), и о параметрах книги, которых нет в конфиге."""
    if force:            # распознать колонки заново, без сохранённых ответов про вид
        cfg = dict(cfg, columns={k: {kk: vv for kk, vv in (v or {}).items() if kk not in ("kind", "role")}
                                 for k, v in (cfg.get("columns") or {}).items()})
    files_raw = {} if force else (raw.get("files") or {})
    reports, broken = [], []
    for f, fc in discover(src_dir, cfg):
        try:
            reports.append((f, prepare(os.path.join(src_dir, f), f, fc, cfg["book"].get("period") or "")))
        except Exception as err:   # о битых файлах только сообщаем
            broken.append((f, str(err)))
    structures = {}
    for f, r in reports:
        structures.setdefault(structure_key(r), []).append((f, r))
    book_raw = {} if force else (raw.get("book") or {})
    has_years = any(c.year for _, r in reports for c in r.metrics)
    has_delta = any(c.kind in ("delta", "delta_pct") for _, r in reports for c in r.metrics)
    book = [k for k in ("title", "subtitle") + (("period",) if has_years else ()) if k not in book_raw]
    if (has_delta or has_years) and (force or "positive_is" not in (raw.get("coloring") or {})):
        book.append("positive_is")

    def known(f):
        return any(pattern_match(p, f) or pattern_match(p, os.path.splitext(f)[0]) for p in files_raw)
    todo = [items for items in structures.values() if not all(known(f) for f, _ in items)]
    return {"reports": reports, "broken": broken, "structures": list(structures.values()), "todo": todo,
            "book": book, "count": len(book) + len(todo)}


def interview(src_dir, cfg, raw, asker, force=False):
    """Вопросы аналитику. Возвращает (конфиг с ответами, были ли вопросы)."""
    plan = plan_questions(src_dir, cfg, raw, force)
    out = asker.out
    raw = copy.deepcopy(raw)
    n_files = len(plan["reports"])
    out(f"\nПапка: {src_dir}\nCSV: {n_files + len(plan['broken'])}, прочитано {n_files}, "
        f"структур (одинаковый набор колонок): {len(plan['structures'])}")
    for f, err in plan["broken"]:
        out(f"  • {f}: не прочитан — {err}")
    if not plan["count"]:
        out("Уточнять нечего — всё есть в конфиге (спросить заново: --ask).")
        return raw, False
    out("Enter — оставить вариант в скобках, «!» — принять все остальные варианты.\n")

    book = raw.setdefault("book", {})
    folder = os.path.basename(src_dir.rstrip("\\/"))
    if plan["book"]:
        out("— Отчёт целиком —")
    if "title" in plan["book"]:
        book["title"] = asker.text("Заголовок отчёта (вверху каждого листа)", book.get("title") or folder)
    if "subtitle" in plan["book"]:
        book["subtitle"] = asker.text("Подзаголовок — отбор / источник (например «ЗЛ, ДЗМ, взрослые»)",
                                      book.get("subtitle") or "")
    if "period" in plan["book"]:
        book["period"] = asker.text("Период к годам в шапке (например «9м» -> «2026 9м»)", book.get("period") or "")
    if "positive_is" in plan["book"]:
        raw.setdefault("coloring", {})["positive_is"] = asker.choice(
            "Рост показателей в этих отчётах — это:",
            [("bad", "плохо — рост красный, снижение зелёное (затраты, объёмы оплаты)"),
             ("good", "хорошо — рост зелёный, снижение красное (выручка, выполнение плана)")],
            (raw.get("coloring") or {}).get("positive_is", "bad"))

    columns = raw.setdefault("columns", {})
    files = raw.setdefault("files", {})
    todo = plan["todo"]
    # вопрос 5 (последний): колонки — автоматически или проверить в консоли; подробно — в веб-приложении
    if todo:
        n_doubt = sum(1 for items in todo if any(doubts(items[0][1], columns)))
        how = asker.choice(
            f"Колонки и листы ({len(todo)} тип(ов) отчётов, с сомнениями: {n_doubt}):",
            [("auto", "распознать автоматически (подробная настройка — веб-приложение report_app.py)"),
             ("review", "проверить здесь, в консоли")], "auto")
        if how == "auto":
            auto_sheet_names([it for items in todo for it in items], files)
            return raw, True
    if len(todo) > 1:
        out(f"\n— Структуры файлов: {len(todo)} —")
        for k, items in enumerate(todo, start=1):
            r = items[0][1]
            unsure, tech = doubts(r, columns)
            n_metric = sum(1 for c in r.columns if c.is_metric and not c.formula)
            what = (f"сомнения: {len(unsure)} кол." if unsure else "") + \
                   (f"{'; ' if unsure else ''}тех. имена: {len(tech)}" if tech else "")
            out(f"  {k:>3}. {len(items)} файл(ов), колонок {len(r.source_columns)}, метрик {n_metric}"
                f"{', иерархия' if r.hierarchy else ''} — {what or 'распознано уверенно'}  ({items[0][0]}"
                f"{' …' if len(items) > 1 else ''})")
        n_doubt = sum(1 for items in todo if any(doubts(items[0][1], columns)))
        mode = asker.choice("Какие структуры проверить вручную?",
                            [("doubt", f"только с сомнениями ({n_doubt} из {len(todo)})"),
                             ("all", "все"), ("none", "ни одной — принять, как распознано")], "doubt")
        if mode == "doubt":
            todo = [items for items in todo if any(doubts(items[0][1], columns))]
        elif mode == "none":
            todo = []
    for k, items in enumerate(todo, start=1):
        ask_structure(asker, k, len(todo), items, columns, files)
    todo = {f for items in plan["todo"] for f, _ in items}
    ask_sheet_names(asker, [(f, r) for f, r in plan["reports"] if f in todo], files)
    return raw, True


def ask_structure(asker, k, total, items, columns, files):
    out = asker.out
    f0, rep = items[0]
    names = [f for f, _ in items]
    shown_files = ", ".join(names[:4]) + (f" и ещё {len(names) - 4}" if len(names) > 4 else "")
    out(f"\n— Структура {k} из {total}: {len(items)} файл(ов) — {shown_files}")
    out(f"  {'иерархия (колонка «Уровень»)' if rep.hierarchy else 'плоская таблица'}, строк в первом файле: {rep.n_rows}")
    cols = rep.columns

    def line(i, c):
        sample = ", ".join(map(str, c.sample[:2]))
        shown = c.name if c.label in (c.name, c.base) or c.formula else f"{c.name} -> «{c.label}»"
        out(f"  {i:>3}. {shown[:38]:<38} {describe(c):<16} {sample[:26]:<26} {'; '.join(column_flags(c))}")

    if len(cols) <= WIDE_COLUMNS:
        for i, c in enumerate(cols, start=1):
            line(i, c)
    else:                      # широкий файл: сводка по видам и только сомнительные колонки
        counts = {}
        for c in cols:
            counts[describe(c)] = counts.get(describe(c), 0) + 1
        out(f"  колонок {len(cols)}: " + ", ".join(f"{k} — {v}" for k, v in counts.items()))
        unsure, _ = doubts(rep, columns)
        if unsure:
            out("  сомнительные колонки (номера — сквозные):")
            for c in unsure:
                line(cols.index(c) + 1, c)
        if asker.choice("Показать все колонки?", [("no", "нет"), ("yes", "да")], "no") == "yes":
            for i, c in enumerate(cols, start=1):
                line(i, c)

    # 1. вид колонок
    editable = [c for c in cols if not c.formula and c.role not in ("level", "name", "subrow")]
    unsure = [c for c in doubts(rep, columns)[0] if c in editable]
    hint = ", * — все сомнительные одним видом" if len(unsure) > 1 else ""
    while True:
        answer = asker._read(f"Изменить вид колонок? Номера через запятую (2,5 или 3-6){hint}, Enter — всё верно "
                             "[всё верно]: ").strip()
        if answer == "*" and hint:
            kind = asker.choice(f"Вид для всех сомнительных ({len(unsure)}):", KIND_CHOICES, unsure[0].kind)
            set_kind(columns, unsure, kind)
            chosen = []
            break
        try:
            chosen = parse_numbers(answer, len(cols)) if answer else []
            break
        except ValueError as err:
            out(f"   не понял «{err}»: номера от 1 до {len(cols)} через запятую или диапазон 2-5")
    for i in chosen:
        c = cols[i - 1]
        if c not in editable:
            out(f"   «{c.label}»: {'добавлена скриптом — меняется через исходные колонки' if c.formula else 'служебная роль, не меняется'}")
            continue
        current = "dimension" if not c.is_metric else ("pct" if c.kind == "pct_points" else c.kind)
        siblings = [x for x in editable if x is c or (c.year and x.year and x.base == c.base)]
        also = [x.name for x in siblings if x is not c]
        kind = asker.choice(f"«{c.name}» (например: {', '.join(map(str, c.sample))})"
                            + (f" — то же для: {', '.join(also)}" if also else "") + ":", KIND_CHOICES, current)
        set_kind(columns, siblings, kind)       # «Сумма 2026» и «Сумма 2025» — одного вида

    # 2. подсветка
    per_file = {}
    metric_idx = [i for i, c in enumerate(cols, start=1) if c.is_metric]
    default_hl = [i for i, c in enumerate(cols, start=1) if c.is_metric and c.kind in ("delta", "delta_pct")]
    if metric_idx:
        default_text = (("Δ и Δ%: " + ",".join(map(str, default_hl))) if len(default_hl) <= 8
                        else f"все Δ и Δ% ({len(default_hl)} шт.)") if default_hl else "ничего"
        hl = ask_numbers(asker, "Что подсвечивать шкалой «рост / снижение»? Номера, 0 — ничего", len(cols), default_text)
        if hl is not None and hl != default_hl:
            if hl:
                per_file["coloring"] = {"columns": [cols[i - 1].name for i in hl]}
            else:
                per_file["coloring"] = {"mode": "off"}

        # 3. главная метрика
        main, _ = main_metric(rep.metrics, rep.table)
        if main is not None:
            default_i = cols.index(main) + 1
            what = "сортировка детей внутри групп и выделенная колонка" if rep.hierarchy else "выделенная колонка"
            while True:
                ans = asker._read(f"Главная метрика ({what}) — номер [{default_i}. {main.name}]: ")
                if not ans:
                    break
                if ans.isdigit() and int(ans) in metric_idx:
                    if int(ans) != default_i:
                        per_file.setdefault("table", {})["key"] = cols[int(ans) - 1].name
                    break
                out(f"   нужен номер метрики: {', '.join(map(str, metric_idx))}")

        # 4. строка «Итого»
        if not rep.hierarchy and any(c.additive for c in rep.metrics):
            totals = asker.choice("Строка «Итого» под шапкой:", [("on", "да"), ("off", "нет")], "on")
            if totals == "off":
                per_file.setdefault("table", {})["totals"] = "off"

    # 5. подписи технических имён
    tech = [c for c in cols if TECH_NAME.match(c.name) and not rules_for(columns, c.name).get("label")]
    if tech:
        out("Подписи для технических имён (Enter — оставить как есть):")
    for c in tech:
        columns.setdefault(c.name, {})["label"] = asker.text(
            f"  {c.name} (например: {', '.join(map(str, c.sample[:3]))})", c.name)
    for f, _ in items:
        files[f] = deep_merge(files.get(f) or {}, per_file)


def set_kind(columns, cols, kind):
    for x in cols:
        rule = columns.setdefault(x.name, {})
        for key in ("kind", "role"):
            rule.pop(key, None)
        if kind in ROLE_KINDS:
            rule["role"] = kind
        else:
            rule.update({"role": "metric", "kind": kind})


def auto_sheet_names(items, files):
    """Имена листов без вопросов: из номера отчёта и имени файла (до 31 символа, без повторов)."""
    used = {(e or {}).get("sheet", "").lower() for f, e in files.items() if f not in dict(items)}
    for f, _ in items:
        no, rest = report_number(os.path.splitext(f)[0])
        entry = files.setdefault(f, {})
        entry["sheet"] = sheet_title(entry.get("sheet") or default_sheet_name(f, [x for x, _ in items]), used)


def ask_sheet_names(asker, items, files):
    if not items:
        return
    out = asker.out
    proposed = {}
    used = set()
    for f, _ in items:
        no, rest = report_number(os.path.splitext(f)[0])
        proposed[f] = sheet_title((files.get(f) or {}).get("sheet") or default_sheet_name(f, [x for x, _ in items]),
                                  used)
    out("\n— Имена листов —")
    for f, name in proposed.items():
        out(f"  {name:<31}  <- {f}")
    how = asker.choice("Имена листов:", [("auto", "как в списке"), ("manual", "задать каждому (и описание)")], "auto")
    for f, _ in items:
        entry = files.setdefault(f, {})
        if how == "manual":
            entry["sheet"] = asker.text(f"  {f}: имя листа (до 31 символа)", proposed[f])[:31]
            description = asker.text("     описание под заголовком (Enter — пропустить)", entry.get("description") or "")
            if description:
                entry["description"] = description
        else:
            entry["sheet"] = proposed[f]


if yaml is not None:
    class _PlainDumper(yaml.SafeDumper):
        """Без ссылок &id001 / *id001: одинаковые списки у разных файлов пишутся полностью — удобно править руками."""
        def ignore_aliases(self, data):
            return True


def save_answers(raw, path):
    """Записывает конфиг с ответами; прежний файл — в .bak (комментарии в нём сохраняются)."""
    if os.path.exists(path):
        os.replace(path, path + ".bak")
        print(f"Прежний конфиг сохранён: {path}.bak")
    if path.lower().endswith(".json") or yaml is None:
        path = os.path.splitext(path)[0] + ".json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)
    else:
        with open(path, "w", encoding="utf-8") as f:
            f.write("# Настройки отчёта: ответы на вопросы report_builder.py. Править можно вручную;\n"
                    "# все ключи — в report.example.yaml. Повторить вопросы: --ask\n")
            yaml.dump(raw, f, Dumper=_PlainDumper, allow_unicode=True, sort_keys=False, default_flow_style=False)
    print(f"Ответы сохранены: {path}")
    return path


# ---------------------------------------------------------------- шаблон конфига
def yq(text):
    return json.dumps(str(text), ensure_ascii=False)


def write_template(src_dir, cfg, path):
    lines = [
        "# Настройки сборки отчёта. Все поля необязательны: удалите лишнее — сработает автоопределение.",
        "# Маски имён: точное имя, * и ? (glob) или re:регулярка. Регистр не важен.",
        "",
        "book:",
        f"  title: {yq(os.path.basename(src_dir.rstrip(chr(92) + '/')))}     # заголовок книги и листов",
        '  subtitle: ""            # красная строка под заголовком (отбор, источник): «ЗЛ, ДЗМ, без МРВСМ»',
        '  period: ""              # подпись к годам в шапке: «9м» -> «2026 9м»',
        "  notes:                  # пояснения на листе «Содержание»",
        '    - "Суммы — млн руб."',
        "  # output: D:\\\\отчёты\\\\итог.xlsx",
        "",
        "style:",
        "  palette: B             # B — стандартный шаблон (основной) | A — старый формат АПП",
        "  # colors: {title: \"1B3A5C\"}   # переопределить отдельные цвета палитры",
        "  font: Arial Narrow",
        "  size: 14",
        "  view: page_break        # page_break | normal",
        "  zoom: 70",
        "  paper: auto             # auto | A3 | A4",
        "",
        "table:",
        "  compare: auto           # auto — добавить Δ и Δ% для пар «X 2026 / X 2025» | off",
        "  formulas: values        # values — в ячейках числа, формулы текстом на листе «Формулы» | live",
        "  totals: auto            # плоские таблицы: строка «Итого» под шапкой | off",
        "  autofilter: true",
        "  # sort: \"-Сумма 2026\"   # минус — по убыванию",
        "  hierarchy_sort: auto    # иерархия (колонка «Уровень»): дети по убыванию главной метрики | off",
        "",
        "coloring:",
        "  mode: static            # static — заливкой (быстро, не «плывёт») | conditional — условное форм. | off",
        "  positive_is: bad        # bad — рост красный (затраты) | good — рост зелёный (выручка)",
        "  # columns: [\"Δ Случаи\", \"Δ% *\"]   # что красить; по умолчанию — все Δ и Δ%",
        "  pct_cap: 2.0            # Δ% > 200% -> «> 200%»",
        "",
        "# Правила для колонок во всех файлах. Ключи: role (metric | dimension | name | level | service | skip),",
        "# kind (count | money | number | pct | pct_points | delta | delta_pct | text | date), label, group,",
        "# format (формат Excel), width, hidden, colorize (true/false), positive_is, level_label, description.",
        "columns:",
        '  "re:^(id|gid|pat_id)$": {role: dimension}',
        '  # "Сумма*": {kind: money, description: "Сумма оплаты, млн руб."}',
        '  # "OTD_NAME": {label: "Отделение", level_label: "Отделение"}',
        "",
        "# Настройки файлов (по маске имени). Ключи: sheet (имя листа ≤31), title, subtitle, description,",
        "# order (место листа: «1.4.1» — сразу после 1.4), skip, а также table / coloring / columns — только для этого файла.",
        "files:",
    ]
    for f, fc in discover(src_dir, cfg):
        stem = os.path.splitext(f)[0]
        lines.append(f"  {yq(f)}:")
        lines.append(f"    sheet: {yq(humanize(stem)[:31])}")
        lines.append('    description: ""')
        try:
            rep = prepare(os.path.join(src_dir, f), f, fc, cfg["book"].get("period") or "")
            kind = "иерархия по «Уровень»" if rep.hierarchy else "плоская таблица"
            lines.append(f"    # {kind}; строк {rep.n_rows}; {rep.encoding}, разделитель {sep_name(rep.sep)}")
            lines.append("    # метрики: " + (", ".join(f"{c.name} [{c.kind}{', добавлено' if c.formula else ''}]"
                                                for c in rep.metrics) or "нет"))
            dims = [c.name for c in rep.columns if not c.is_metric]
            lines.append("    # разрезы: " + (", ".join(dims) or "нет"))
        except Exception as err:
            lines.append(f"    # не прочитан: {err}")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print(f"Шаблон настроек: {path}")


# ---------------------------------------------------------------- запуск
def safe_console():
    """Вывод в консоль без падений: при перенаправлении (в файл, в другую программу, планировщик) Windows отдаёт
    кодировку cp1252/cp1251, где нет части символов, и print падал с UnicodeEncodeError."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream.isatty():
                stream.reconfigure(errors="replace")
            else:
                stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError, OSError):
            pass


def main():
    safe_console()
    p = argparse.ArgumentParser(description="Excel-отчёт в едином оформлении из всех CSV в папке")
    p.add_argument("src", help="папка с CSV")
    p.add_argument("--out", help="итоговый .xlsx (по умолчанию <папка>.xlsx рядом с папкой)")
    p.add_argument("--config", help="конфиг .yaml/.json (по умолчанию report.yaml в папке, если есть)")
    p.add_argument("--init", action="store_true", help="создать шаблон report.yaml по найденным файлам")
    p.add_argument("--explain", action="store_true", help="показать распознанные колонки, без сборки")
    p.add_argument("--ask", action="store_true", help="задать уточняющие вопросы, даже если всё уже в конфиге")
    p.add_argument("--no-ask", action="store_true", help="не задавать вопросов (для запуска по расписанию)")
    args = p.parse_args()
    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        sys.exit(f"Нет папки {src}")
    config_path = args.config or next((os.path.join(src, n) for n in CONFIG_NAMES
                                       if os.path.exists(os.path.join(src, n))), None)
    cfg = load_config(config_path)
    if config_path:
        print(f"Конфиг: {config_path}")
    if args.init:
        target = os.path.join(src, "report.yaml")
        if os.path.exists(target):
            target = os.path.join(src, "report.template.yaml")
        write_template(src, cfg, target)
        return
    # уточнения: по умолчанию — в консоли, если есть что спросить; ответы -> report.yaml
    if not args.no_ask and not args.explain and (args.ask or sys.stdin.isatty()):
        answers, changed = interview(src, cfg, load_raw(config_path), Asker(), force=args.ask)
        if changed:
            config_path = save_answers(answers, config_path or os.path.join(src, "report.yaml"))
            cfg = load_config(config_path)
    out = args.out or cfg["book"].get("output") or src.rstrip("\\/") + ".xlsx"
    if os.path.exists(out) and not args.explain:
        try:
            os.rename(out, out)   # открыт в Excel -> на Windows ошибка
        except OSError:
            sys.exit(f"Файл {out} открыт в Excel — закройте его или укажите другой --out")
    build(src, cfg, out, explain=args.explain)


if __name__ == "__main__":
    main()
