#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ch_loader.py — пошаговый загрузчик файлов (CSV / TSV / TXT, Excel, Parquet) в ClickHouse.

=============================== КРАТКАЯ ИНСТРУКЦИЯ ===============================
1. Python 3.9+ и зависимости:
       pip install clickhouse-connect clickhouse-driver pyarrow charset-normalizer python-calamine
   (clickhouse-driver нужен только для TCP-портов 9000/9440; если python-calamine не ставится —
    pip install openpyxl, Excel будет читаться медленнее.)

2. Что подготовить перед запуском:
   • адрес сервера: хост, «хост:порт» или URL (https://host:8443). Порт и протокол
     (HTTP / HTTPS / TCP / TCP+TLS) можно не указывать — скрипт определит сам;
   • логин, пароль и базу данных;
   • файл: .csv .tsv .txt (можно .gz), .xlsx / .xlsm, .parquet;
   • имя таблицы — новой или уже существующей;
   • для НОВОЙ таблицы решите, по каким колонкам вы чаще всего фильтруете данные:
     это ORDER BY (ключ сортировки), после создания таблицы его не поменять.

3. Запуск:  python ch_loader.py   — дальше отвечайте на вопросы. Enter = значение в [скобках].

Как обеспечивается корректность:
   • кодировка, разделитель, заголовок и типы определяются автоматически, типы — по ВСЕМУ файлу;
     всё найденное показывается и подтверждается вами;
   • данные сначала грузятся во временную таблицу, сверяется число строк, и только потом
     попадают в целевую: новая таблица появляется атомарно (RENAME), в существующую
     данные добавляются одним INSERT ... SELECT;
   • существующие данные не удаляются и не перезаписываются без явного подтверждения;
     при замене старые данные остаются в резервной таблице;
   • любое значение, которое не помещается в тип без потерь (переполнение, лишние знаки
     после запятой, несуществующая дата), останавливает загрузку до записи в целевую таблицу;
   • Ctrl+C в любой момент — временная таблица удаляется, целевая не меняется.
==================================================================================
"""
from __future__ import annotations

import codecs
import csv
import datetime as dt
import decimal
import getpass
import gzip
import io
import os
import re
import socket
import sys
import time
import uuid
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Callable, Iterator, Optional

try:
    import pyarrow as pa
    import pyarrow.compute as pc
    import pyarrow.csv as pacsv
except ImportError:  # pragma: no cover
    sys.exit("Не установлен pyarrow:  pip install pyarrow")

APP = "ch_loader"
SAMPLE_BYTES = 2 * 1024 * 1024        # объём начала файла для автоопределения
BATCH_BYTES = 16 * 1024 * 1024        # целевой объём одной пачки (до конвертации)
MIN_BATCH_ROWS, MAX_BATCH_ROWS = 10_000, 500_000  # insert-batch-size: 10K+ строк на INSERT
LOWCARD_MAX = 10_000                  # schema-types-lowcardinality: < 10K уникальных значений
NULL_TOKENS = ["", "NULL", "null", "Null", "\\N", "None", "N/A", "n/a", "#N/A"]
NULL_SET = pa.array(NULL_TOKENS, pa.string())
csv.field_size_limit(2**31 - 1)


# ═══════════════════════════════════ ВВОД / ВЫВОД ═══════════════════════════════════

class Abort(Exception):
    """Пользователь отменил операцию или продолжение небезопасно."""


def say(msg: str = "") -> None:
    print(msg, flush=True)


def step(title: str) -> None:
    say("\n" + "─" * 78 + f"\n▶ {title}\n" + "─" * 78)


def warn(msg: str) -> None:
    say(f"⚠  {msg}")


def ask(prompt: str, default: Optional[str] = None, secret: bool = False) -> str:
    suffix = f" [{default}]" if default not in (None, "") else ""
    while True:
        try:
            raw = getpass.getpass(f"{prompt}{suffix}: ") if secret else input(f"{prompt}{suffix}: ")
        except EOFError:
            raise Abort("ввод прерван")
        raw = raw.strip()
        if raw:
            return raw
        if default is not None:
            return default
        say("  (нужно ввести значение)")


def ask_yes(prompt: str, default: bool) -> bool:
    hint = "Д/н" if default else "д/Н"
    while True:
        a = ask(f"{prompt} ({hint})", "").lower()
        if a == "":
            return default
        if a in ("д", "да", "y", "yes", "l"):
            return True
        if a in ("н", "нет", "n", "no", "т"):
            return False
        say("  ответьте «д» или «н»")


def choose(prompt: str, options: list[str], default: int = 1) -> int:
    """Возвращает индекс (с 0) выбранного варианта."""
    for i, o in enumerate(options, 1):
        say(f"  {i}) {o}")
    while True:
        a = ask(prompt, str(default))
        if a.isdigit() and 1 <= int(a) <= len(options):
            return int(a) - 1
        say(f"  введите число от 1 до {len(options)}")


def short(v, n: int = 40) -> str:
    s = repr(v) if not isinstance(v, str) else v
    s = s.replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
    return s if len(s) <= n else s[: n - 1] + "…"


def print_table(headers: list[str], rows: list[list[str]], max_w: int = 34) -> None:
    rows = [[short(c, max_w) for c in r] for r in rows]
    w = [max([len(h)] + [len(r[i]) for r in rows]) for i, h in enumerate(headers)]
    say("  " + " │ ".join(h.ljust(w[i]) for i, h in enumerate(headers)))
    say("  " + "─┼─".join("─" * x for x in w))
    for r in rows:
        say("  " + " │ ".join(c.ljust(w[i]) for i, c in enumerate(r)))


class Progress:
    def __init__(self, label: str, total_bytes: int = 0):
        self.label, self.total, self.t0, self.last = label, total_bytes, time.time(), 0.0

    def update(self, rows: int, done_bytes: int = 0, force: bool = False) -> None:
        now = time.time()
        if not force and now - self.last < 0.5:
            return
        self.last = now
        el = max(now - self.t0, 1e-6)
        pct = f" {100 * done_bytes / self.total:5.1f}%" if self.total and done_bytes else ""
        sys.stderr.write(f"\r  {self.label}:{pct} {rows:,} строк, {rows / el:,.0f} строк/с, {el:,.0f} с   ")
        sys.stderr.flush()

    def done(self, rows: int) -> None:
        self.update(rows, self.total, force=True)
        sys.stderr.write("\n")


def qi(name: str) -> str:
    """Экранирование идентификатора ClickHouse."""
    return "`" + name.replace("\\", "\\\\").replace("`", "\\`") + "`"


def ql(value: str) -> str:
    """Экранирование строкового литерала ClickHouse."""
    return "'" + value.replace("\\", "\\\\").replace("'", "\\'") + "'"


# ════════════════════════════════ ПОДКЛЮЧЕНИЕ К СЕРВЕРУ ═══════════════════════════════

class AuthFailed(Exception):
    pass


def _is_auth_error(e: Exception) -> bool:
    code = getattr(e, "code", None)
    text = str(e)
    return code in (516, 192, 194) or any(
        f"Code: {c}" in text for c in (516, 192, 194)) or "Authentication failed" in text


def _is_network_error(e: Exception) -> bool:
    if isinstance(e, (ConnectionError, TimeoutError, socket.timeout, OSError)):
        return True
    name = type(e).__name__
    text = str(e)
    if name in ("OperationalError", "NetworkError", "SocketTimeoutError"):
        return True
    return any(s in text for s in ("Connection reset", "Connection aborted", "timed out",
                                   "RemoteDisconnected", "Broken pipe", "Max retries exceeded"))


class Backend:
    protocol = ""
    secure = False
    host = ""
    port = 0

    def query(self, sql: str, settings: Optional[dict] = None) -> list[tuple]:
        raise NotImplementedError

    def command(self, sql: str, settings: Optional[dict] = None) -> int:
        """Выполняет запрос, возвращает written_rows (если сервер его сообщил, иначе -1)."""
        raise NotImplementedError

    def insert(self, table: str, columns: list[str], specs: list["TypeSpec"],
               data: list[pa.Array], settings: dict) -> None:
        raise NotImplementedError

    def describe(self) -> str:
        name = {("http", False): "HTTP", ("http", True): "HTTPS",
                ("native", False): "TCP (native)", ("native", True): "TCP+TLS (native)"}
        return f"{name[(self.protocol, self.secure)]} {self.host}:{self.port}"

    def scalar(self, sql: str, settings: Optional[dict] = None):
        rows = self.query(sql, settings)
        return rows[0][0] if rows else None


INSERT_SETTINGS_BASE = {
    "async_insert": 0,                      # синхронная вставка: пачки уже крупные
    "insert_deduplicate": 1,
    "input_format_allow_errors_num": 0,
    "input_format_allow_errors_ratio": 0,
    "format_csv_allow_single_quotes": 0,
    "input_format_null_as_default": 0,
    "input_format_csv_empty_as_default": 0,
}


class HttpBackend(Backend):
    protocol = "http"

    def __init__(self, host, port, user, password, database, secure, verify, timeout, probe=False):
        import clickhouse_connect
        self.host, self.port, self.secure = host, port, secure
        self.client = clickhouse_connect.get_client(
            host=host, port=port, username=user, password=password, database=database,
            secure=secure, verify=verify, connect_timeout=timeout,
            send_receive_timeout=timeout + 2 if probe else 3600, compress=False,
            autogenerate_session_id=False)
        self.client.query("SELECT 1")
        try:
            import lz4.frame
            self._compress, self._encoding = lz4.frame.compress, "lz4"
        except ImportError:
            self._compress = lambda b: gzip.compress(b, compresslevel=1)
            self._encoding = "gzip"

    def query(self, sql, settings=None):
        return list(self.client.query(sql, settings=settings).result_rows)

    def command(self, sql, settings=None):
        r = self.client.command(sql, settings=settings)
        return int(r.written_rows) if hasattr(r, "written_rows") else -1

    def insert(self, table, columns, specs, data, settings):
        # Bool в кавычках ClickHouse в CSV не принимает — отправляем 1/0 без кавычек
        data = [pc.cast(pc.equal(a, "true"), pa.int8()) if s.kind == "bool" else a for a, s in zip(data, specs)]
        tbl = pa.Table.from_arrays(data, names=[f"c{i}" for i in range(len(data))])
        buf = io.BytesIO()
        pacsv.write_csv(tbl, buf, pacsv.WriteOptions(
            include_header=False, null_string="\\N", quoting_style="needed"))
        payload = self._compress(buf.getvalue())
        self.client.raw_insert(table, columns, payload, settings=settings, fmt="CSV",
                               compression=self._encoding)


class NativeBackend(Backend):
    protocol = "native"

    def __init__(self, host, port, user, password, database, secure, verify, timeout, probe=False):
        try:
            from clickhouse_driver import Client
        except ImportError:
            raise RuntimeError("для TCP-протокола нужен пакет clickhouse-driver: pip install clickhouse-driver")
        self.host, self.port, self.secure = host, port, secure
        self.client = Client(host=host, port=port, user=user, password=password, database=database,
                             secure=secure, verify=verify, connect_timeout=timeout,
                             send_receive_timeout=timeout + 2 if probe else 3600,
                             sync_request_timeout=timeout)
        self.client.execute("SELECT 1")

    def query(self, sql, settings=None):
        return self.client.execute(sql, settings=settings)

    def command(self, sql, settings=None):
        self.client.execute(sql, settings=settings)
        prog = getattr(self.client.last_query, "progress", None)
        return int(getattr(prog, "written_rows", -1)) if prog is not None else -1

    def insert(self, table, columns, specs, data, settings):
        cols = [to_python(arr, spec) for arr, spec in zip(data, specs)]
        sql = f"INSERT INTO {table} ({', '.join(qi(c) for c in columns)}) VALUES"
        self.client.execute(sql, cols, columnar=True, settings=settings)


def parse_address(raw: str) -> tuple[str, Optional[int], Optional[tuple[str, bool]]]:
    raw = raw.strip().rstrip("/")
    forced = None
    m = re.match(r"^(\w+)://(.*)$", raw)
    if m:
        scheme, raw = m.group(1).lower(), m.group(2)
        forced = {"http": ("http", False), "https": ("http", True), "tcp": ("native", False),
                  "native": ("native", False), "clickhouse": ("native", False),
                  "tls": ("native", True), "clickhouses": ("native", True)}.get(scheme)
        if forced is None:
            raise ValueError(f"неизвестная схема {scheme}://")
    raw = raw.split("/")[0].split("@")[-1]
    port = None
    m = re.match(r"^\[(.+)\](?::(\d+))?$", raw) or re.match(r"^([^:]+)(?::(\d+))?$", raw)
    if not m:
        raise ValueError("не удалось разобрать адрес")
    host = m.group(1)
    if m.group(2):
        port = int(m.group(2))
    return host, port, forced


def candidate_endpoints(host, port, forced) -> list[tuple[str, bool, int]]:
    defaults = {("http", False): 8123, ("http", True): 8443, ("native", False): 9000, ("native", True): 9440}
    if forced:
        return [(forced[0], forced[1], port or defaults[forced])]
    if port:
        order = [("http", False), ("http", True), ("native", False), ("native", True)]
        hint = {8123: ("http", False), 8443: ("http", True), 443: ("http", True), 80: ("http", False),
                9000: ("native", False), 9440: ("native", True)}.get(port)
        if hint:
            order.remove(hint)
            order.insert(0, hint)
        return [(p, s, port) for p, s in order]
    order = [("http", False), ("http", True), ("native", False), ("native", True)]
    if host.endswith(".clickhouse.cloud"):
        order = [("http", True), ("native", True)]
    return [(p, s, defaults[(p, s)]) for p, s in order]


def port_open(host: str, port: int, timeout: float = 3.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def connect_wizard() -> Backend:
    step("Шаг 1. Подключение к ClickHouse")
    while True:
        try:
            host, port, forced = parse_address(ask("Адрес сервера (хост, хост:порт или URL)", "localhost"))
        except ValueError as e:
            warn(str(e))
            continue
        user = ask("Пользователь", "default")
        password = ask("Пароль (ввод скрыт, Enter — пустой)", "", secret=True)
        database = ask("База данных", "default")
        verify = True
        backend = None
        tried = []
        for proto, secure, p in candidate_endpoints(host, port, forced):
            if not port_open(host, p):
                tried.append(f"{p}: порт закрыт")
                continue
            cls = HttpBackend if proto == "http" else NativeBackend
            label = f"{'HTTPS' if proto == 'http' and secure else proto.upper()}{'+TLS' if proto == 'native' and secure else ''}:{p}"
            try:
                say(f"  пробую {label} …")
                cls(host, p, user, password, database, secure, verify, 8, probe=True)
                backend = cls(host, p, user, password, database, secure, verify, 8)
                break
            except Exception as e:  # noqa: BLE001 — классифицируем ниже
                if _is_auth_error(e):
                    warn(f"сервер ответил по {label}, но отклонил логин/пароль: {short(str(e), 200)}")
                    backend = "auth"
                    break
                if "CERTIFICATE_VERIFY_FAILED" in str(e) and verify:
                    warn(f"{label}: сертификат сервера не прошёл проверку ({short(str(e), 120)})")
                    if ask_yes("Подключиться без проверки сертификата (только если доверяете сети)?", False):
                        verify = False
                        try:
                            backend = cls(host, p, user, password, database, secure, verify, 8)
                            break
                        except Exception as e2:  # noqa: BLE001
                            e = e2
                tried.append(f"{label}: {short(str(e), 120)}")
        if isinstance(backend, Backend):
            ver = backend.scalar("SELECT version()")
            tz = backend.scalar("SELECT timezone()")
            say(f"✓ Подключено: {backend.describe()}, ClickHouse {ver}, часовой пояс сервера {tz}")
            backend.server_tz = tz
            backend.database = database
            return backend
        if backend != "auth":
            warn("не удалось подключиться:\n    " + "\n    ".join(tried or ["нет кандидатов"]))
        if not ask_yes("Попробовать снова?", True):
            raise Abort("нет подключения")


# ═══════════════════════════════════════ ТИПЫ ═══════════════════════════════════════

INT_BOUNDS = {(b, s): ((-(2 ** (b - 1)), 2 ** (b - 1) - 1) if s else (0, 2 ** b - 1))
              for b in (8, 16, 32, 64, 128, 256) for s in (True, False)}
DATE_RANGE = {"Date": (dt.datetime(1970, 1, 1), dt.datetime(2149, 6, 6)),
              "Date32": (dt.datetime(1900, 1, 1), dt.datetime(2299, 12, 31))}
DATETIME_RANGE = (dt.datetime(1970, 1, 1), dt.datetime(2106, 2, 7, 6, 28, 15))
DATETIME64_RANGE = (dt.datetime(1900, 1, 1), dt.datetime(2299, 12, 31, 23, 59, 59))
TYPE_DEFAULT = {"int": "0", "float": "0", "decimal": "0", "bool": "false", "date": "1970-01-01",
                "datetime": "1970-01-01 00:00:00", "uuid": "00000000-0000-0000-0000-000000000000"}


@dataclass
class TypeSpec:
    kind: str                  # int float decimal bool date datetime uuid string fixedstring raw
    base: str                  # тип без Nullable/LowCardinality, как в ClickHouse
    nullable: bool = False
    lowcard: bool = False
    bits: int = 0
    signed: bool = True
    precision: int = 0         # Decimal P или точность DateTime64
    scale: int = 0
    tz: Optional[str] = None
    length: int = 0            # FixedString(N)
    fill_default: bool = False  # пустые значения → значение по умолчанию (для не-Nullable)

    @property
    def ch(self) -> str:
        t = self.base
        if self.nullable:
            t = f"Nullable({t})"
        if self.lowcard:
            t = f"LowCardinality({t})"
        return t


def parse_ch_type(t: str) -> TypeSpec:
    t = t.strip()
    lowcard = nullable = False
    m = re.fullmatch(r"LowCardinality\((.*)\)", t)
    if m:
        lowcard, t = True, m.group(1).strip()
    m = re.fullmatch(r"Nullable\((.*)\)", t)
    if m:
        nullable, t = True, m.group(1).strip()
    s = TypeSpec("raw", t, nullable, lowcard)
    if m := re.fullmatch(r"(U?)Int(8|16|32|64|128|256)", t):
        s.kind, s.signed, s.bits = "int", m.group(1) == "", int(m.group(2))
    elif re.fullmatch(r"Float(32|64)|BFloat16", t):
        s.kind = "float"
    elif m := re.fullmatch(r"Decimal\(\s*(\d+)\s*,\s*(\d+)\s*\)", t):
        s.kind, s.precision, s.scale = "decimal", int(m.group(1)), int(m.group(2))
    elif m := re.fullmatch(r"Decimal(32|64|128|256)\(\s*(\d+)\s*\)", t):
        s.kind, s.scale = "decimal", int(m.group(2))
        s.precision = {"32": 9, "64": 18, "128": 38, "256": 76}[m.group(1)]
    elif t in ("Bool", "Boolean"):
        s.kind = "bool"
    elif t in ("Date", "Date32"):
        s.kind = "date"
    elif m := re.fullmatch(r"DateTime(?:\(\s*'([^']*)'\s*\))?", t):
        s.kind, s.precision, s.tz = "datetime", 0, m.group(1)
    elif m := re.fullmatch(r"DateTime64\(\s*(\d)\s*(?:,\s*'([^']*)'\s*)?\)", t):
        s.kind, s.precision, s.tz = "datetime", int(m.group(1)), m.group(2)
    elif t == "UUID":
        s.kind = "uuid"
    elif t == "String":
        s.kind = "string"
    elif m := re.fullmatch(r"FixedString\((\d+)\)", t):
        s.kind, s.length = "fixedstring", int(m.group(1))
    return s


def smallest_int(lo: int, hi: int) -> str:
    for bits in (8, 16, 32, 64, 128):
        if lo >= 0 and hi <= INT_BOUNDS[(bits, False)][1]:
            return f"UInt{bits}"
        a, b = INT_BOUNDS[(bits, True)]
        if a <= lo and hi <= b:
            return f"Int{bits}"
    return "Int256"


# ═══════════════════════ РАСПОЗНАВАНИЕ ЗНАЧЕНИЙ (векторно, Arrow) ═══════════════════════

INT_RE = r"^[+-]?(?:0|[1-9][0-9]*)$"
BOOL_RE = r"^(?i:true|false)$"
UUID_RE = r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"


def dec_re(sep: str) -> str:
    return rf"^[+-]?(?:0|[1-9][0-9]*)(?:{re.escape(sep)}[0-9]+)?$"


def dec_parts_re(sep: str) -> str:
    return rf"^[+-]?(?P<i>[0-9]+)(?:{re.escape(sep)}(?P<f>[0-9]+))?$"


def float_re(sep: str) -> str:
    s = re.escape(sep)
    return (rf"^(?:[+-]?(?:(?:0|[1-9][0-9]*)(?:{s}[0-9]*)?|{s}[0-9]+)(?:[eE][+-]?[0-9]+)?"
            r"|(?i:[+-]?(?:nan|inf|infinity)))$")


_D, _M, _Y = r"(?P<d>[0-9]{1,2})", r"(?P<m>[0-9]{1,2})", r"(?P<y>[0-9]{4})"
DATE_PATTERNS = {   # ключ → (описание, регулярка даты)
    "ymd-": ("ГГГГ-ММ-ДД", rf"{_Y}-{_M}-{_D}"),
    "dmy.": ("ДД.ММ.ГГГГ", rf"{_D}\.{_M}\.{_Y}"),
    "dmy/": ("ДД/ММ/ГГГГ", rf"{_D}/{_M}/{_Y}"),
    "mdy/": ("ММ/ДД/ГГГГ", rf"{_M}/{_D}/{_Y}"),
    "ymd/": ("ГГГГ/ММ/ДД", rf"{_Y}/{_M}/{_D}"),
    "ymd.": ("ГГГГ.ММ.ДД", rf"{_Y}\.{_M}\.{_D}"),
    "dmy-": ("ДД-ММ-ГГГГ", rf"{_D}-{_M}-{_Y}"),
}
TIME_RE = r"(?:[ T](?P<H>[0-9]{1,2}):(?P<M>[0-9]{2})(?::(?P<S>[0-9]{2})(?:[.,](?P<f>[0-9]{1,9}))?)?(?P<z>Z)?)?"
DT_RE = {k: "^" + v[1] + TIME_RE + "$" for k, v in DATE_PATTERNS.items()}


def n_true(mask) -> int:
    return int(pc.sum(pc.cast(pc.fill_null(mask, False), pa.int64())).as_py() or 0)


def first_false(mask) -> int:
    """Индекс первого False (null считается True); -1 если таких нет."""
    m = pc.fill_null(mask, True)
    idx = pc.index(m, False).as_py()
    return idx


def amax(arr) -> int:
    v = pc.max(arr).as_py()
    return 0 if v is None else v


@dataclass
class DtParse:
    date: pa.Array        # 'YYYY-MM-DD'
    hms: pa.Array         # 'HH:MM:SS'
    frac: pa.Array        # дробная часть секунд ('' если нет)
    has_time: pa.Array
    has_z: pa.Array
    valid: pa.Array       # True — корректная дата/время (null для пустых)
    ts: pa.Array          # timestamp[s]


def parse_dt(t: pa.Array, pattern: str) -> DtParse:
    """t — строки (null = пусто), все уже совпали с DT_RE[pattern]."""
    ex = pc.extract_regex(t, DT_RE[pattern])
    f = {n: ex.field(n) for n in ("y", "m", "d", "H", "M", "S", "f", "z")}
    date = pc.binary_join_element_wise(f["y"], pc.utf8_lpad(f["m"], 2, "0"), pc.utf8_lpad(f["d"], 2, "0"), "-")
    has_time = pc.not_equal(f["H"], "")
    hh = pc.if_else(has_time, pc.utf8_lpad(f["H"], 2, "0"), "00")
    mm = pc.if_else(has_time, f["M"], "00")
    ss = pc.if_else(pc.equal(f["S"], ""), "00", f["S"])
    hms = pc.binary_join_element_wise(hh, mm, ss, ":")
    full = pc.binary_join_element_wise(date, hms, " ")
    full = pc.if_else(pc.is_null(t), pa.scalar(None, pa.string()), full)
    ts = pc.strptime(full, format="%Y-%m-%d %H:%M:%S", unit="s", error_is_null=True)
    back = pc.strftime(ts, format="%Y-%m-%d %H:%M:%S")
    valid = pc.if_else(pc.is_null(t), pa.scalar(None, pa.bool_()), pc.fill_null(pc.equal(back, full), False))
    return DtParse(date, hms, f["f"], has_time, pc.equal(f["z"], "Z"), valid, ts)


# ═══════════════════════════════════ ПРОФИЛЬ КОЛОНКИ ═══════════════════════════════════

CANDIDATES = ["bool", "int", "dec.", "dec,", "float.", "float,"] + [f"dt:{k}" for k in DATE_PATTERNS] + ["uuid"]


@dataclass
class ColumnProfile:
    name: str
    rows: int = 0
    nulls: int = 0                      # пустые / NULL-подобные значения
    alive: list = field(default_factory=lambda: list(CANDIDATES))
    killed_by: dict = field(default_factory=dict)  # кандидат → (номер строки, значение)
    int_min: Optional[int] = None
    int_max: Optional[int] = None
    int_big: bool = False               # не влезло в Int64/UInt64
    dec: dict = field(default_factory=lambda: {".": [0, 0, 0, 0], ",": [0, 0, 0, 0]})  # i, f, has_f, f==3
    float_special: bool = False         # экспонента / nan / inf
    dt: dict = field(default_factory=dict)  # pattern → {time, frac, z, min, max}
    max_bytes: int = 0
    distinct: Optional[set] = field(default_factory=set)
    examples: list = field(default_factory=list)
    fixed: Optional[TypeSpec] = None    # тип задан источником (Parquet)
    note: str = ""

    def update(self, arr: pa.Array, offset: int) -> None:
        n = len(arr)
        self.rows += n
        lens = pc.binary_length(arr)
        self.max_bytes = max(self.max_bytes, amax(lens))
        if len(self.examples) < 3:
            for v in pc.filter(arr, pc.greater(lens, 0)).slice(0, 3).to_pylist():
                if len(self.examples) < 3 and v not in self.examples:
                    self.examples.append(v)
        if self.distinct is not None:
            self.distinct.update(pc.unique(arr).to_pylist())
            if len(self.distinct) > LOWCARD_MAX:
                self.distinct = None
        trimmed = pc.utf8_trim_whitespace(arr)
        isnull = pc.is_in(trimmed, value_set=NULL_SET)
        self.nulls += n_true(isnull)
        t = pc.if_else(isnull, pa.scalar(None, pa.string()), trimmed)
        if t.null_count == n:
            return
        for cand in list(self.alive):
            bad = self._check(cand, t)
            if bad is not None:
                self.alive.remove(cand)
                self.killed_by[cand] = (offset + bad, t[bad].as_py())

    def _check(self, cand: str, t: pa.Array) -> Optional[int]:
        """Возвращает индекс первого неподходящего значения или None."""
        if cand == "bool":
            i = first_false(pc.match_substring_regex(t, BOOL_RE))
            return None if i < 0 else i
        if cand == "uuid":
            i = first_false(pc.match_substring_regex(t, UUID_RE))
            return None if i < 0 else i
        if cand == "int":
            i = first_false(pc.match_substring_regex(t, INT_RE))
            if i >= 0:
                return i
            try:
                v = pc.cast(t, pa.int64())
            except pa.ArrowInvalid:
                try:
                    v = pc.cast(t, pa.uint64())
                except pa.ArrowInvalid:
                    digits = pc.utf8_length(pc.replace_substring_regex(t, r"^[+-]", ""))
                    if amax(digits) > 38:
                        return first_false(pc.less_equal(digits, 38))
                    self.int_big = True
                    return None
            mm = pc.min_max(v)
            lo, hi = mm["min"].as_py(), mm["max"].as_py()
            self.int_min = lo if self.int_min is None else min(self.int_min, lo)
            self.int_max = hi if self.int_max is None else max(self.int_max, hi)
            return None
        if cand.startswith("dec"):
            sep = cand[3]
            i = first_false(pc.match_substring_regex(t, dec_re(sep)))
            if i >= 0:
                return i
            ex = pc.extract_regex(t, dec_parts_re(sep))
            li, lf = pc.utf8_length(ex.field("i")), pc.utf8_length(ex.field("f"))
            st = self.dec[sep]
            st[0], st[1] = max(st[0], amax(li)), max(st[1], amax(lf))
            st[2] += n_true(pc.greater(lf, 0))
            st[3] += n_true(pc.equal(lf, 3))
            return None
        if cand.startswith("float"):
            sep = cand[5]
            i = first_false(pc.match_substring_regex(t, float_re(sep)))
            if i >= 0:
                return i
            if n_true(pc.match_substring_regex(t, r"[eEnN]")):
                self.float_special = True
            return None
        if cand.startswith("dt:"):
            p = cand[3:]
            i = first_false(pc.match_substring_regex(t, DT_RE[p]))
            if i >= 0:
                return i
            r = parse_dt(t, p)
            i = first_false(r.valid)
            if i >= 0:
                return i
            st = self.dt.setdefault(p, {"time": False, "frac": 0, "z": 0, "nz": 0, "min": None, "max": None})
            nonnull = len(t) - t.null_count
            st["time"] = st["time"] or n_true(r.has_time) > 0
            st["frac"] = max(st["frac"], amax(pc.utf8_length(r.frac)))
            st["z"] += n_true(r.has_z)
            st["nz"] += nonnull - n_true(r.has_z)
            mm = pc.min_max(r.ts)
            lo, hi = mm["min"].as_py(), mm["max"].as_py()
            st["min"] = lo if st["min"] is None else min(st["min"], lo)
            st["max"] = hi if st["max"] is None else max(st["max"], hi)
            if st["z"] and st["nz"]:   # смесь значений с 'Z' и без — неоднозначно
                return first_false(r.has_z) if n_true(r.has_z) else 0
            return None
        return None

    @property
    def nonnull(self) -> int:
        return self.rows - self.nulls

    def dt_pattern(self) -> Optional[str]:
        for c in self.alive:
            if c.startswith("dt:"):
                return c[3:]
        return None

    def ambiguities(self) -> list[str]:
        out = []
        if "dt:dmy/" in self.alive and "dt:mdy/" in self.alive:
            out.append("дата может быть как ДД/ММ/ГГГГ, так и ММ/ДД/ГГГГ (все дни ≤ 12) — выбран ДД/ММ")
        if "int" not in self.alive and "dec." not in self.alive and "dec," in self.alive:
            st = self.dec[","]
            if st[2] and st[2] == st[3]:
                out.append("запятая везде отделяет ровно 3 цифры: это может быть разделитель тысяч, а не дробная часть")
        return out

    def propose(self) -> TypeSpec:
        """Тип без потерь по всему файлу (schema-types-native-types, minimize-bitwidth)."""
        if self.fixed is not None:
            s = TypeSpec(**self.fixed.__dict__)
            s.nullable = s.nullable or (self.nulls > 0 and s.kind not in ("string", "fixedstring", "raw"))
            return s
        if self.nonnull == 0:
            return TypeSpec("string", "String")
        a = self.alive
        s: Optional[TypeSpec] = None
        if "bool" in a:
            s = TypeSpec("bool", "Bool")
        elif "int" in a:
            if self.int_big:
                s = TypeSpec("int", "Int128", bits=128)
            else:
                base = smallest_int(self.int_min, self.int_max)
                s = parse_ch_type(base)
        else:
            for sep in (".", ","):
                if f"dec{sep}" in a:
                    i, f = self.dec[sep][0], self.dec[sep][1]
                    need = i + f
                    if need <= 38:
                        p = 9 if need <= 9 else 18 if need <= 18 else 38
                        s = TypeSpec("decimal", f"Decimal({p}, {f})", precision=p, scale=f)
                    break
            if s is None and self.float_special and ("float." in a or "float," in a):
                s = TypeSpec("float", "Float64")
            if s is None and (p := self.dt_pattern()):
                st = self.dt[p]
                tz = "UTC" if st["z"] else None
                if not st["time"]:
                    for base in ("Date", "Date32"):
                        lo, hi = DATE_RANGE[base]
                        if lo <= st["min"] and st["max"] <= hi:
                            s = TypeSpec("date", base)
                            break
                else:
                    if st["frac"] == 0 and DATETIME_RANGE[0] <= st["min"] and st["max"] <= DATETIME_RANGE[1]:
                        s = TypeSpec("datetime", f"DateTime('{tz}')" if tz else "DateTime", tz=tz)
                    elif DATETIME64_RANGE[0] <= st["min"] and st["max"] <= DATETIME64_RANGE[1]:
                        prec = 3 if st["frac"] <= 3 else 6 if st["frac"] <= 6 else 9
                        base = f"DateTime64({prec}, '{tz}')" if tz else f"DateTime64({prec})"
                        s = TypeSpec("datetime", base, precision=prec, tz=tz)
            if s is None and "uuid" in a:
                s = TypeSpec("uuid", "UUID")
        if s is None:
            s = TypeSpec("string", "String")
            if self.distinct is not None and self.rows >= 1000 and len(self.distinct) <= self.rows // 2:
                s.lowcard = True
            return s
        s.nullable = self.nulls > 0
        return s

    def compat(self, s: TypeSpec) -> tuple[bool, str]:
        """Можно ли без потерь записать данные колонки в тип s. (ok, причина/предупреждение)."""
        if s.kind == "string":
            return True, ""
        if s.kind == "fixedstring":
            return (self.max_bytes <= s.length,
                    "" if self.max_bytes <= s.length else f"есть значения длиной {self.max_bytes} байт > {s.length}")
        if s.kind == "raw":
            return True, f"тип {s.base} проверит сам сервер при загрузке во временную таблицу"
        if self.nonnull == 0:
            return True, ""
        a = self.alive
        why = lambda c: (f"строка {self.killed_by[c][0] + 1}: {short(self.killed_by[c][1])!r}"
                         if c in self.killed_by else "")
        if s.kind == "int":
            if "int" not in a:
                return False, f"не целое число ({why('int')})"
            if self.int_big and s.bits <= 64:
                return False, "есть числа больше 64 бит"
            lo, hi = INT_BOUNDS[(s.bits, s.signed)]
            if not self.int_big and not (lo <= self.int_min and self.int_max <= hi):
                return False, f"диапазон данных [{self.int_min}; {self.int_max}] не помещается в {s.base}"
            return True, ""
        if s.kind == "float":
            if any(c in a for c in ("int", "float.", "float,")):
                return True, ""
            return False, f"не число ({why('float.')})"
        if s.kind == "decimal":
            if "int" in a and not self.int_big:
                digits = max(len(str(abs(self.int_min))), len(str(abs(self.int_max))))
                ok = digits <= s.precision - s.scale
                return ok, "" if ok else f"целая часть до {digits} знаков > {s.precision - s.scale}"
            for sep in (".", ","):
                if f"dec{sep}" in a:
                    i, f = self.dec[sep][0], self.dec[sep][1]
                    if f > s.scale:
                        return False, f"до {f} знаков после запятой > {s.scale} — было бы округление"
                    if i > s.precision - s.scale:
                        return False, f"целая часть до {i} знаков > {s.precision - s.scale}"
                    return True, ""
            return False, f"не десятичное число ({why('dec.')})"
        if s.kind == "bool":
            if "bool" in a or ("int" in a and self.int_min is not None and self.int_min >= 0 and self.int_max <= 1):
                return True, ""
            return False, f"не true/false ({why('bool')})"
        if s.kind in ("date", "datetime"):
            p = self.dt_pattern()
            if not p:
                return False, f"не дата ({why('dt:ymd-')})"
            st = self.dt[p]
            if s.kind == "date":
                if st["time"]:
                    return False, "в значениях есть время — тип Date его отбросит"
                lo, hi = DATE_RANGE[s.base]
            else:
                if st["frac"] > s.precision:
                    return False, f"до {st['frac']} знаков долей секунды > точности {s.precision}"
                lo, hi = DATETIME_RANGE if s.base.startswith("DateTime(") or s.base == "DateTime" else DATETIME64_RANGE
            if not (lo <= st["min"] and st["max"] <= hi):
                return False, f"даты от {st['min']} до {st['max']} вне диапазона {s.base}"
            return True, ""
        if s.kind == "uuid":
            return ("uuid" in a), ("" if "uuid" in a else f"не UUID ({why('uuid')})")
        return True, ""


# ═══════════════════════════════════ КОНВЕРТАЦИЯ ═══════════════════════════════════

class ConvertError(Exception):
    def __init__(self, column: str, row: int, value, reason: str):
        super().__init__(f"колонка «{column}», строка данных {row + 1}: значение {short(value)!r} — {reason}")


def make_converter(prof: ColumnProfile, spec: TypeSpec) -> Callable[[pa.Array, int], pa.Array]:
    """Возвращает функцию: сырые строки → каноничные строки для ClickHouse (null = NULL).
    Каждая функция заново и строго проверяет каждое значение."""
    null_str = pa.scalar(None, pa.string())
    kind = spec.kind
    sep = "."
    if kind == "decimal" or kind == "float":
        sep = "," if ("dec," in prof.alive or "float," in prof.alive) and not (
            "dec." in prof.alive or "float." in prof.alive or "int" in prof.alive) else "."
    pattern = prof.dt_pattern()
    bool_from_int = kind == "bool" and "bool" not in prof.alive

    def fail(mask, src, offset, reason):
        i = first_false(mask)
        if i >= 0:
            raise ConvertError(prof.name, offset + i, src[i].as_py(), reason)

    def conv(arr: pa.Array, offset: int) -> pa.Array:
        if kind in ("string", "fixedstring"):
            if kind == "fixedstring":
                fail(pc.less_equal(pc.binary_length(arr), spec.length), arr, offset,
                     f"длиннее FixedString({spec.length})")
            if spec.nullable:
                return pc.if_else(pc.equal(arr, "\\N"), null_str, arr)
            return arr
        trimmed = pc.utf8_trim_whitespace(arr)
        isnull = pc.is_in(trimmed, value_set=NULL_SET)
        t = pc.if_else(isnull, null_str, trimmed)
        if kind == "raw":
            out = t
        elif kind == "int":
            fail(pc.match_substring_regex(t, INT_RE), arr, offset, "не целое число")
            out = pc.replace_substring_regex(t, r"^\+", "")
            lo, hi = INT_BOUNDS[(spec.bits, spec.signed)]
            if spec.bits <= 64:
                try:
                    v = pc.cast(out, pa.int64() if spec.signed else pa.uint64())
                except pa.ArrowInvalid:
                    v = None
                if v is None or (len(v) - v.null_count and (pc.min(v).as_py() < lo or pc.max(v).as_py() > hi)):
                    bad = [i for i, x in enumerate(out.to_pylist()) if x is not None and not lo <= int(x) <= hi]
                    raise ConvertError(prof.name, offset + bad[0], arr[bad[0]].as_py(), f"вне диапазона {spec.base}")
            else:
                for i, x in enumerate(out.to_pylist()):
                    if x is not None and len(x) > 18 and not lo <= int(x) <= hi:
                        raise ConvertError(prof.name, offset + i, arr[i].as_py(), f"вне диапазона {spec.base}")
        elif kind == "float":
            fail(pc.match_substring_regex(t, float_re(sep)), arr, offset, "не число")
            norm = pc.replace_substring(t, ",", ".") if sep == "," else t
            out = pc.cast(pc.cast(norm, pa.float64()), pa.string())
        elif kind == "decimal":
            fail(pc.match_substring_regex(t, dec_re(sep)), arr, offset, "не десятичное число")
            ex = pc.extract_regex(t, dec_parts_re(sep))
            fail(pc.less_equal(pc.utf8_length(ex.field("f")), spec.scale), arr, offset,
                 f"больше {spec.scale} знаков после запятой — было бы округление")
            fail(pc.less_equal(pc.utf8_length(pc.utf8_ltrim(ex.field("i"), "0")), spec.precision - spec.scale),
                 arr, offset, f"не помещается в {spec.base}")
            out = pc.replace_substring_regex(t, r"^\+", "")
            if sep == ",":
                out = pc.replace_substring(out, ",", ".")
        elif kind == "bool":
            if bool_from_int:
                fail(pc.is_in(t, value_set=pa.array(["0", "1"])), arr, offset, "не 0/1")
                out = pc.if_else(pc.equal(t, "1"), "true", "false")
            else:
                fail(pc.match_substring_regex(t, BOOL_RE), arr, offset, "не true/false")
                out = pc.utf8_lower(t)
        elif kind in ("date", "datetime"):
            if pattern is None:
                raise ConvertError(prof.name, offset, arr[0].as_py(), "колонка не распознана как дата")
            fail(pc.match_substring_regex(t, DT_RE[pattern]), arr, offset,
                 f"не дата в формате {DATE_PATTERNS[pattern][0]}")
            r = parse_dt(t, pattern)
            fail(r.valid, arr, offset, "несуществующая дата/время")
            if kind == "date":
                fail(pc.invert(r.has_time), arr, offset, "есть время, а тип Date")
                lo, hi = DATE_RANGE[spec.base]
                out = r.date
            else:
                lo, hi = DATETIME_RANGE if spec.precision == 0 and spec.base.startswith("DateTime") \
                    and not spec.base.startswith("DateTime64") else DATETIME64_RANGE
                fail(pc.less_equal(pc.utf8_length(r.frac), spec.precision), arr, offset,
                     f"долей секунды больше, чем точность {spec.precision}")
                out = pc.binary_join_element_wise(r.date, r.hms, " ")
                if spec.precision:
                    out = pc.binary_join_element_wise(out, pc.utf8_rpad(r.frac, spec.precision, "0"), ".")
                out = pc.if_else(r.has_z, pc.binary_join_element_wise(out, "Z", ""), out)
            fail(pc.and_(pc.greater_equal(r.ts, pa.scalar(lo, pa.timestamp("s"))),
                         pc.less_equal(r.ts, pa.scalar(hi, pa.timestamp("s")))),
                 arr, offset, f"вне диапазона {spec.base}")
        elif kind == "uuid":
            fail(pc.match_substring_regex(t, UUID_RE), arr, offset, "не UUID")
            out = pc.utf8_lower(t)
        else:
            out = t
        out = pc.if_else(isnull, null_str, out)
        if out.null_count and not spec.nullable:
            if spec.fill_default and kind in TYPE_DEFAULT:
                out = pc.fill_null(out, TYPE_DEFAULT[kind])
            else:
                i = first_false(pc.is_valid(out))
                raise ConvertError(prof.name, offset + i, arr[i].as_py(),
                                   f"пустое значение, а тип {spec.ch} не допускает NULL")
        return out

    return conv


def to_python(arr: pa.Array, spec: TypeSpec) -> list:
    """Каноничные строки → значения Python для clickhouse-driver (TCP)."""
    k = spec.kind
    if k == "int":
        if spec.bits <= 64:
            return pc.cast(arr, pa.int64() if spec.signed else pa.uint64()).to_pylist()
        return [None if x is None else int(x) for x in arr.to_pylist()]
    if k == "float":
        return pc.cast(arr, pa.float64()).to_pylist()
    if k == "decimal":
        return [None if x is None else decimal.Decimal(x) for x in arr.to_pylist()]
    if k == "bool":
        return [None if x is None else x == "true" for x in arr.to_pylist()]
    if k == "date":
        return pc.cast(pc.strptime(arr, format="%Y-%m-%d", unit="s"), pa.date32()).to_pylist()
    if k == "datetime":
        has_z = pc.fill_null(pc.ends_with(arr, "Z"), False)
        ts = pc.cast(pc.replace_substring(arr, "Z", ""), pa.timestamp("us")).to_pylist()
        z = has_z.to_pylist()
        return [None if v is None else (v.replace(tzinfo=dt.timezone.utc) if zz else v) for v, zz in zip(ts, z)]
    if k == "uuid":
        return [None if x is None else uuid.UUID(x) for x in arr.to_pylist()]
    return arr.to_pylist()


# ═════════════════════════════════════ ИСТОЧНИКИ ═════════════════════════════════════

QUOTE_WRAP = [  # «вечные кавычки»: значение целиком обёрнуто в кавычки
    (r'^"(?:[^"]|"")*"$', r'^"(.*)"$', '""', '"'),
    (r"^'[^']*'$", r"^'(.*)'$", None, None),
]


def unwrap_quotes(arr: pa.Array) -> tuple[pa.Array, pa.Array]:
    """Снимает лишние внешние кавычки (повторно, до 5 слоёв). Возвращает (новые значения, маска изменённых)."""
    changed = pa.array([False] * len(arr))
    for _ in range(5):
        any_hit = False
        for full, inner, dbl, single in QUOTE_WRAP:
            hit = pc.and_(pc.match_substring_regex(arr, full), pc.greater_equal(pc.utf8_length(arr), 2))
            if not n_true(hit):
                continue
            any_hit = True
            stripped = pc.utf8_slice_codeunits(arr, 1, -1)
            if dbl:
                stripped = pc.replace_substring(stripped, dbl, single)
            arr = pc.if_else(hit, stripped, arr)
            changed = pc.or_(changed, hit)
        if not any_hit:
            break
    return arr, changed


@dataclass
class Dialect:
    delimiter: str
    quoting: int = csv.QUOTE_MINIMAL
    line_wrapped: bool = False      # каждая строка целиком в кавычках: "a;b;c"

    def describe(self) -> str:
        names = {",": "запятая", ";": "точка с запятой", "\t": "табуляция", "|": "вертикальная черта"}
        s = names.get(self.delimiter, repr(self.delimiter))
        if self.quoting == csv.QUOTE_NONE:
            s += ", без кавычек"
        if self.line_wrapped:
            s += ", каждая строка обёрнута в кавычки"
        return s

    def reader_kwargs(self, strict: bool = False) -> dict:
        # strict=True: кавычка, после которой не идёт разделитель, — ошибка, а не «молчаливая» склейка
        return dict(delimiter=self.delimiter, quotechar='"', doublequote=True,
                    quoting=self.quoting, skipinitialspace=False, strict=strict)


def open_binary(path: str):
    return gzip.open(path, "rb") if path.lower().endswith(".gz") else open(path, "rb")


def detect_encoding(path: str) -> tuple[str, str]:
    """(кодировка, пояснение). UTF-8 проверяется по всему файлу."""
    with open_binary(path) as f:
        head = f.read(SAMPLE_BYTES)
    if head.startswith(codecs.BOM_UTF8):
        return "utf-8-sig", "BOM UTF-8"
    if head.startswith(codecs.BOM_UTF16_LE) or head.startswith(codecs.BOM_UTF16_BE):
        return "utf-16", "BOM UTF-16"
    dec = codecs.getincrementaldecoder("utf-8")("strict")
    try:
        with open_binary(path) as f:
            while chunk := f.read(8 * 1024 * 1024):
                dec.decode(chunk)
            dec.decode(b"", final=True)
        return "utf-8", "весь файл — корректный UTF-8"
    except UnicodeDecodeError:
        pass
    try:
        from charset_normalizer import from_bytes
        # mac_cyrillic / iso8859_5 почти неотличимы от cp1251 и на практике встречаются редко —
        # не предлагаем их, чтобы файл из Windows не прочитался с искажёнными буквами
        best = from_bytes(head, cp_isolation=["cp1251", "koi8_r", "cp866", "cp1252", "utf_16"]).best()
        if best is not None:
            enc = best.encoding.replace("_", "-")
            return enc, "угадана по содержимому — проверьте текст в примере ниже"
    except ImportError:
        pass
    return "cp1251", "не UTF-8; по умолчанию Windows-1251"


def sniff_dialect(text: str, ext: str) -> tuple[Dialect, int, float]:
    """Подбирает разделитель по согласованности числа полей. (диалект, полей, доля согласованных строк)."""
    text = text[: text.rfind("\n") + 1] if "\n" in text else text
    delims = ["\t", ",", ";", "|"] if ext in (".tsv", ".tab") else [",", ";", "\t", "|"]

    def score(sample: str, d: Dialect) -> tuple[int, float]:
        try:
            rows = [r for r in csv.reader(io.StringIO(sample), **d.reader_kwargs()) if r]
        except csv.Error:
            return 1, 0.0
        counts = [len(r) for r in rows]
        if not counts:
            return 1, 0.0
        mode = max(set(counts), key=lambda c: (counts.count(c), c))
        return mode, counts.count(mode) / len(counts)

    def best(variants: list[Dialect], sample: str):
        res = []
        for prio, dia in enumerate(variants):
            n, cons = score(sample, dia)
            if n > 1:   # согласованность важнее всего, затем больше колонок, затем привычный порядок
                res.append(((round(cons, 2), n, -prio), dia, n, cons))
        return max(res, key=lambda r: r[0])[1:] if res else None

    # 1) обычный CSV с кавычками
    found = best([Dialect(d) for d in delims], text)
    if found and found[2] >= 0.9:
        return found
    # 2) каждая строка целиком обёрнута в кавычки: "a;b;c"
    lines = [ln for ln in text.splitlines() if ln.strip()][:500]
    if lines and all(len(ln) >= 2 and ln[0] == '"' and ln[-1] == '"' for ln in lines):
        unwrapped = "\n".join(ln[1:-1].replace('""', '"') for ln in lines) + "\n"
        wrapped = best([Dialect(d, csv.QUOTE_MINIMAL, True) for d in delims], unwrapped)
        if wrapped and wrapped[2] >= 0.9:
            return wrapped
    # 3) кавычки не используются как ограничители (одиночные " внутри полей)
    plain = best([Dialect(d, csv.QUOTE_NONE) for d in delims], text)
    if plain and (not found or plain[2] > found[2] + 0.05):
        return plain
    return found or (Dialect(delims[0]), 1, 1.0)


def tolerant_split(text: str, delim: str) -> Optional[list[str]]:
    """Разбор строки с «кривыми» кавычками: кавычка закрывает поле, только если за ней
    идёт разделитель или конец строки. Возвращает None, если поле не закрыто."""
    text = text.rstrip("\r\n")
    fields, i, n = [], 0, len(text)
    while True:
        if i < n and text[i] == '"':
            j, buf = i + 1, []
            while True:
                k = text.find('"', j)
                if k == -1:
                    return None
                nxt = text[k + 1] if k + 1 < n else ""
                if nxt == "" or nxt == delim:
                    buf.append(text[j:k])
                    i = k + 1
                    break
                if nxt == '"':
                    after = text[k + 2] if k + 2 < n else ""
                    buf.append(text[j:k] + '"')
                    if after == "" or after == delim:   # "" перед разделителем: кавычка + закрытие
                        i = k + 2
                        break
                    j = k + 2
                    continue
                buf.append(text[j:k + 1])               # одиночная кавычка внутри — литерал
                j = k + 1
            fields.append("".join(buf))
            if i >= n:
                break
            i += 1
            if i == n:
                fields.append("")
                break
        else:
            k = text.find(delim, i)
            if k == -1:
                fields.append(text[i:])
                break
            fields.append(text[i:k])
            i = k + 1
            if i == n:
                fields.append("")
                break
    return fields


class _Pushback:
    """Итератор строк с возвратом — для повторного разбора «кривых» записей.
    Хранит физические номера строк файла."""

    def __init__(self, it):
        self.it, self.back, self.taken, self.num = it, deque(), [], 0

    def __iter__(self):
        return self

    def __next__(self):
        if self.back:
            item = self.back.popleft()
        else:
            self.num += 1
            item = (next(self.it), self.num)
        self.taken.append(item)
        return item[0]

    def push_back(self, items) -> None:
        self.back.extendleft(reversed(items))


@dataclass
class RowIssues:
    repaired: list = field(default_factory=list)   # (номер строки файла, исходный текст, поля)
    rejected: list = field(default_factory=list)   # (номер строки файла, исходный текст, причина)
    repaired_count: int = 0
    rejected_count: int = 0
    trimmed_trailing: int = 0


class Source:
    """Общий интерфейс: names, iter_batches() → (список pa.StringArray по колонкам, кол-во строк)."""
    kind = ""
    names: list[str] = []
    fixed_types: dict = {}
    total_bytes = 0

    def iter_batches(self, rows_per_batch: int) -> Iterator[tuple[list[pa.Array], int, int]]:
        raise NotImplementedError


class CsvSource(Source):
    kind = "csv"

    def __init__(self, path: str, encoding: str, dialect: Dialect, has_header: bool, ncols: int):
        self.path, self.encoding, self.dialect, self.has_header = path, encoding, dialect, has_header
        self.ncols = ncols
        self.repair = False
        # быстрый путь (pyarrow) — только для «чистого» RFC-CSV; иначе строгий разбор модулем csv
        self.fast = (dialect.quoting == csv.QUOTE_MINIMAL and not dialect.line_wrapped
                     and not encoding.lower().replace("_", "-").startswith("utf-16"))
        self.issues = RowIssues()
        self.total_bytes = os.path.getsize(path)
        self.names = []

    def _open_text(self):
        return io.TextIOWrapper(open_binary(self.path), encoding=self.encoding, errors="strict", newline="")

    def _lines(self, f):
        if not self.dialect.line_wrapped:
            return f
        def gen():
            for ln in f:
                body = ln.rstrip("\r\n")
                if len(body) >= 2 and body[0] == '"' and body[-1] == '"':
                    yield body[1:-1].replace('""', '"') + "\n"
                else:
                    yield ln
        return gen()

    def _trim(self, row: Optional[list[str]]) -> Optional[list[str]]:
        """Лишние ПУСТЫЕ поля в конце строки (частая ошибка экспорта) отбрасываются без потерь."""
        n = self.ncols
        if row is not None and len(row) > n and not any(x.strip() for x in row[n:]):
            return row[:n]
        return row

    def _record_re(self):
        d = re.escape(self.dialect.delimiter.encode())
        field_ = rb'(?:"(?:[^"]|"")*+"|[^"' + d + rb'\r\n][^' + d + rb'\r\n]*+|)'
        rec = field_ + rb'(?:' + d + field_ + rb')*+(?:\r\n|\n|\r)'
        return re.compile(rb'(?:' + rec + rb')*+', re.S)

    def _iter_fast(self, rows_per_batch):
        """Быстрый путь. Каждый кусок файла сначала проверяется строгой регуляркой RFC 4180
        (кавычка закрывает поле только перед разделителем/концом строки). Любое отклонение
        или не то число полей → _NeedRepair(fast=True) и переход на строгий разбор модулем csv,
        поэтому «молчаливой» склейки полей здесь не бывает."""
        records = self._record_re()
        names = [f"c{i}" for i in range(self.ncols)]
        conv = pacsv.ConvertOptions(column_types={c: pa.string() for c in names},
                                    strings_can_be_null=False, quoted_strings_can_be_null=False)
        parse = pacsv.ParseOptions(delimiter=self.dialect.delimiter, quote_char='"', double_quote=True,
                                   escape_char=False, newlines_in_values=True, ignore_empty_lines=True,
                                   invalid_row_handler=lambda row: "error")
        enc = self.encoding.lower().replace("_", "-")
        read_enc = "utf8" if enc in ("utf-8", "utf8", "utf-8-sig", "ascii") else self.encoding
        chunk_size = max(BATCH_BYTES, rows_per_batch * 64)
        buf, first, done, pos, lineno = b"", True, False, 0, 1
        with open_binary(self.path) as f:
            while not done:
                chunk = f.read(chunk_size)
                done = not chunk
                if first and chunk.startswith(codecs.BOM_UTF8):
                    chunk = chunk[3:]
                data = buf + chunk
                if done and data and not data.endswith((b"\n", b"\r")):
                    data += b"\n"
                end = records.match(data).end()
                if end == 0 and data:
                    if done or len(data) > 4 * chunk_size:
                        raise _NeedRepair(lineno, fast=True)
                    buf = data
                    continue
                good, buf = data[:end], data[end:]
                if done and buf:
                    raise _NeedRepair(lineno + good.count(b"\n"), fast=True)
                if not good:
                    continue
                skip = 1 if first and self.has_header else 0
                first = False
                try:
                    tbl = pacsv.read_csv(io.BytesIO(good), read_options=pacsv.ReadOptions(
                        column_names=names, skip_rows=skip, encoding=read_enc, block_size=1 << 24),
                        parse_options=parse, convert_options=conv)
                except (pa.ArrowInvalid, UnicodeDecodeError):
                    raise _NeedRepair(lineno, fast=True)
                lineno += good.count(b"\n")
                pos += len(good)
                for off in range(0, tbl.num_rows, rows_per_batch):
                    part = tbl.slice(off, rows_per_batch)
                    yield [part.column(i).combine_chunks() for i in range(self.ncols)], part.num_rows, pos

    def iter_batches(self, rows_per_batch):
        """Пачки по колонкам. self.repair=True — медленный режим с исправлением «кривых» строк."""
        if self.fast and not self.repair:
            yield from self._iter_fast(rows_per_batch)
            return
        self.issues = RowIssues()
        n = self.ncols
        with self._open_text() as f:
            raw = f.buffer
            lines = self._lines(f)
            src = _Pushback(iter(lines)) if self.repair else None
            reader = csv.reader(src if self.repair else lines, **self.dialect.reader_kwargs(strict=True))
            first = True
            batch: list[list[str]] = []
            while True:
                try:
                    row = next(reader)
                    broken = False
                except StopIteration:
                    break
                except csv.Error:
                    if not self.repair:
                        raise _NeedRepair(reader.line_num)
                    row, broken = None, True
                taken = []
                if self.repair:
                    taken, src.taken = src.taken, []
                lineno = taken[0][1] if taken else reader.line_num
                if first:
                    first = False
                    if self.has_header:
                        continue
                if not broken and not row:
                    continue                       # пустая строка
                if broken or len(row) != n:
                    fixed = None if broken else self._trim(row)
                    if fixed is not None and len(fixed) == n:
                        self.issues.trimmed_trailing += 1
                        row = fixed
                    elif not self.repair:
                        raise _NeedRepair(lineno)
                    else:
                        row = self._repair(row, taken, src, lineno)
                        if row is None:
                            continue
                batch.append(row)
                if len(batch) >= rows_per_batch:
                    yield self._to_columns(batch), len(batch), self._pos(raw)
                    batch = []
            if batch:
                yield self._to_columns(batch), len(batch), self._pos(raw)

    def _repair(self, row, taken, src, lineno) -> Optional[list[str]]:
        n, delim = self.ncols, self.dialect.delimiter
        texts = [t for t, _ in taken]
        # 1) первая физическая строка сама по себе, с «терпимым» разбором кавычек
        cand = self._trim(tolerant_split(texts[0], delim)) if texts else None
        if cand is not None and len(cand) == n:
            if len(taken) > 1:
                src.push_back(taken[1:])            # остальные строки разберём заново
            self._note_repair(lineno, texts[0], cand)
            return cand
        # 2) запись на несколько строк (перевод строки внутри кавычек) с «кривыми» кавычками
        if len(texts) > 1:
            cand = self._trim(tolerant_split("".join(texts), delim))
            if cand is not None and len(cand) == n:
                self._note_repair(lineno, "".join(texts), cand)
                return cand
        self.issues.rejected_count += 1
        if len(self.issues.rejected) < 10_000:
            why = "некорректные кавычки" if row is None else f"полей {len(row)} вместо {n}"
            self.issues.rejected.append((lineno, texts[0] if texts else "", why))
        if len(taken) > 1:
            src.push_back(taken[1:])
        return None

    def _note_repair(self, lineno, text, fields) -> None:
        self.issues.repaired_count += 1
        if len(self.issues.repaired) < 20:
            self.issues.repaired.append((lineno, text, fields))

    def _pos(self, raw) -> int:
        try:
            return raw.tell() if not self.path.lower().endswith(".gz") else 0
        except (OSError, ValueError):
            return 0

    def _to_columns(self, rows: list[list[str]]) -> list[pa.Array]:
        return [pa.array(col, pa.string()) for col in zip(*rows)]

    def read_header(self) -> list[str]:
        with self._open_text() as f:
            for row in csv.reader(self._lines(f), **self.dialect.reader_kwargs()):
                if row:
                    return row
        return []


class _NeedRepair(Exception):
    def __init__(self, lineno, fast=False):
        self.lineno, self.fast = lineno, fast


def excel_cell(v) -> str:
    """Значение ячейки Excel → строка без потери информации."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 2 ** 53:
            return str(int(v))
        # Excel хранит и показывает 15 значащих цифр: 4985.010000000001 → 4985.01
        return f"{v:.15g}"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0) and v.tzinfo is None:
            return v.date().isoformat()
        return v.isoformat(sep=" ")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, dt.timedelta):
        return str(v)
    return str(v)


class ExcelSource(Source):
    kind = "excel"

    def __init__(self, path: str, sheet: str, has_header: bool, ncols: int):
        self.path, self.sheet, self.has_header, self.ncols = path, sheet, has_header, ncols
        self.issues = RowIssues()
        self.names = []

    @staticmethod
    def sheets(path: str) -> list[str]:
        try:
            from python_calamine import CalamineWorkbook
            return list(CalamineWorkbook.from_path(path).sheet_names)
        except ImportError:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            names = wb.sheetnames
            wb.close()
            return names

    def raw_rows(self) -> Iterator[list[str]]:
        try:
            from python_calamine import CalamineWorkbook
            sh = CalamineWorkbook.from_path(self.path).get_sheet_by_name(self.sheet)
            it = sh.iter_rows()
        except ImportError:
            import openpyxl
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
            it = wb[self.sheet].iter_rows(values_only=True)
        started = False
        for r in it:
            cells = [excel_cell(v) for v in r]
            if not started:
                if not any(c != "" for c in cells):
                    continue          # пустые строки в начале листа
                started = True
            yield cells

    def iter_batches(self, rows_per_batch):
        self.issues = RowIssues()
        batch, first = [], True
        for cells in self.raw_rows():
            if first:
                first = False
                if self.has_header:
                    continue
            if not any(c != "" for c in cells):
                continue
            if len(cells) < self.ncols:
                cells += [""] * (self.ncols - len(cells))
            elif len(cells) > self.ncols:
                if any(c != "" for c in cells[self.ncols:]):
                    self.issues.rejected_count += 1
                    self.issues.rejected.append((0, "\t".join(cells), "данные правее последней колонки заголовка"))
                    continue
                cells = cells[: self.ncols]
            batch.append(cells)
            if len(batch) >= rows_per_batch:
                yield [pa.array(c, pa.string()) for c in zip(*batch)], len(batch), 0
                batch = []
        if batch:
            yield [pa.array(c, pa.string()) for c in zip(*batch)], len(batch), 0


def arrow_to_spec(t: pa.DataType) -> Optional[TypeSpec]:
    if pa.types.is_boolean(t):
        return TypeSpec("bool", "Bool")
    if pa.types.is_integer(t):
        b, s = t.bit_width, pa.types.is_signed_integer(t)
        return TypeSpec("int", f"{'' if s else 'U'}Int{b}", bits=b, signed=s)
    if pa.types.is_floating(t):
        return TypeSpec("float", "Float32" if t.bit_width == 32 else "Float64")
    if pa.types.is_decimal(t):
        if t.precision > 76:
            return None
        return TypeSpec("decimal", f"Decimal({t.precision}, {t.scale})", precision=t.precision, scale=t.scale)
    if pa.types.is_date(t):
        return TypeSpec("date", "Date32")
    if pa.types.is_timestamp(t):
        p = {"s": 0, "ms": 3, "us": 6, "ns": 9}[t.unit]
        tz = "UTC" if t.tz else None
        base = f"DateTime64({p}, 'UTC')" if tz else f"DateTime64({p})"
        return TypeSpec("datetime", base, precision=p, tz=tz)
    if pa.types.is_string(t) or pa.types.is_large_string(t):
        return TypeSpec("string", "String")
    return None


def arrow_col_to_text(col: pa.Array) -> pa.Array:
    t = col.type
    if pa.types.is_dictionary(t):
        col = col.dictionary_decode()
        t = col.type
    if pa.types.is_timestamp(t):
        if t.tz:
            col = col.cast(pa.timestamp(t.unit))       # момент в UTC без пояса
        s = pc.strftime(col, format="%Y-%m-%d %H:%M:%S") if t.unit == "s" else pc.cast(col, pa.string())
        return pc.fill_null(s, "")
    if pa.types.is_date(t):
        return pc.fill_null(pc.strftime(col.cast(pa.date32()), format="%Y-%m-%d"), "")
    if pa.types.is_binary(t) or pa.types.is_large_binary(t):
        try:
            return pc.fill_null(pc.cast(col, pa.string()), "")
        except pa.ArrowInvalid:
            return pa.array(["" if v is None else v.hex() for v in col.to_pylist()], pa.string())
    if pa.types.is_nested(t):
        import json
        return pa.array(["" if v is None else json.dumps(v, ensure_ascii=False, default=str)
                         for v in col.to_pylist()], pa.string())
    try:
        return pc.fill_null(pc.cast(col, pa.string()), "")
    except (pa.ArrowInvalid, pa.ArrowNotImplementedError):
        return pa.array(["" if v is None else str(v) for v in col.to_pylist()], pa.string())


class ParquetSource(Source):
    kind = "parquet"

    def __init__(self, path: str):
        import pyarrow.parquet as pq
        self.path = path
        self.pf = pq.ParquetFile(path)
        schema = self.pf.schema_arrow
        self.names = list(schema.names)
        self.ncols = len(self.names)
        self.fixed_types = {}
        for name, f in zip(schema.names, schema):
            t = f.type.value_type if pa.types.is_dictionary(f.type) else f.type
            self.fixed_types[name] = arrow_to_spec(t) or TypeSpec("string", "String")
        self.issues = RowIssues()
        self.total_rows = self.pf.metadata.num_rows

    def iter_batches(self, rows_per_batch):
        for rb in self.pf.iter_batches(batch_size=rows_per_batch):
            cols = [arrow_col_to_text(rb.column(i)) for i in range(rb.num_columns)]
            # null в Parquet → пустая строка; для типизированных колонок это NULL
            yield cols, rb.num_rows, 0


# ═════════════════════════════ ВЫБОР ФАЙЛА И ФОРМАТА (шаг 2) ═════════════════════════════

def clean_names(raw: list[str]) -> list[str]:
    out, seen = [], {}
    for i, n in enumerate(raw, 1):
        n = re.sub(r"\s+", " ", str(n)).strip() or f"col_{i}"
        base, k = n, 1
        while n.lower() in seen:
            k += 1
            n = f"{base}_{k}"
        seen[n.lower()] = 1
        out.append(n)
    return out


def guess_header(rows: list[list[str]]) -> bool:
    """Есть ли заголовок: значение первой строки не подходит под тип колонки в остальных строках."""
    if len(rows) < 2:
        return True
    head, body = rows[0], rows[1:300]
    n = len(head)
    if any(h.strip() == "" for h in head) or len({h.lower() for h in head}) < n:
        return False
    typed_cols = 0
    for j in range(n):
        p = ColumnProfile("x")
        p.update(pa.array([r[j] if j < len(r) else "" for r in body], pa.string()), 0)
        spec = p.propose()
        if spec.kind in ("string", "fixedstring", "raw"):
            continue
        typed_cols += 1
        q = ColumnProfile("h")
        q.update(pa.array([head[j]], pa.string()), 0)
        if q.propose().kind == "string" and not q.compat(spec)[0]:
            return True       # над числами/датами стоит текст — это заголовок
    if typed_cols:
        return False
    # все колонки текстовые: заголовок, если его значения не встречаются в данных
    return all(head[j] not in {r[j] for r in body if j < len(r)} for j in range(n))


def file_wizard():
    step("Шаг 2. Файл")
    while True:
        path = ask("Путь к файлу").strip().strip('"').strip("'")
        path = os.path.expanduser(path)
        if os.path.isfile(path):
            break
        warn(f"файл не найден: {path}")
    low = path.lower()
    ext = os.path.splitext(low[:-3] if low.endswith(".gz") else low)[1]
    size = os.path.getsize(path)
    say(f"  размер: {size / 1024 / 1024:,.1f} МБ")

    if ext == ".parquet":
        src = ParquetSource(path)
        say(f"✓ Parquet: {src.ncols} колонок, {src.total_rows:,} строк (типы берутся из схемы файла)")
        src.names = clean_names(src.names)
        src.fixed_types = dict(zip(src.names, src.fixed_types.values()))
        return src, path

    if ext in (".xlsx", ".xlsm", ".xlsb", ".xls", ".ods"):
        sheets = ExcelSource.sheets(path)
        sheet = sheets[0]
        if len(sheets) > 1:
            say("  листы в книге:")
            sheet = sheets[choose("Какой лист загрузить", sheets)]
        probe = ExcelSource(path, sheet, False, 0)
        sample = []
        for r in probe.raw_rows():
            sample.append(r)
            if len(sample) >= 200:
                break
        if not sample:
            raise Abort("лист пустой")
        ncols = max(len(r) for r in sample)
        while ncols > 1 and all(len(r) < ncols or r[ncols - 1] == "" for r in sample):
            ncols -= 1
        sample = [(r + [""] * ncols)[:ncols] for r in sample]
        say(f"✓ Excel, лист «{sheet}», колонок: {ncols}")
        preview(sample)
        has_header = ask_yes("Первая строка — заголовок?", guess_header(sample))
        src = ExcelSource(path, sheet, has_header, ncols)
        src.names = clean_names(sample[0] if has_header else [f"col_{i}" for i in range(1, ncols + 1)])
        return src, path

    # текстовые форматы
    enc, why = detect_encoding(path)
    while True:
        say(f"✓ Кодировка: {enc} ({why})")
        try:
            with io.TextIOWrapper(open_binary(path), encoding=enc, errors="strict", newline="") as f:
                text = f.read(SAMPLE_BYTES)
        except (UnicodeDecodeError, LookupError) as e:
            warn(f"не удалось прочитать в {enc}: {e}")
            enc, why = ask("Укажите кодировку (utf-8, cp1251, koi8-r, cp866, utf-16)"), "указана вручную"
            continue
        dia, ncols, cons = sniff_dialect(text, ext)
        say(f"✓ Разделитель: {dia.describe()}; колонок: {ncols}; согласованность строк в начале файла {cons:.0%}")
        sample = []
        lines = text.splitlines(keepends=True)
        if dia.line_wrapped:
            lines = [ln.rstrip("\r\n")[1:-1].replace('""', '"') + "\n" for ln in lines if ln.strip()]
        for r in csv.reader(lines[:300], **dia.reader_kwargs()):
            if r:
                sample.append((r + [""] * ncols)[:ncols])
        preview(sample)
        if ask_yes("Текст и колонки выглядят правильно?", True):
            break
        a = choose("Что поправить", ["кодировку", "разделитель", "и то и другое"])
        if a in (0, 2):
            enc, why = ask("Кодировка (utf-8, cp1251, koi8-r, cp866, utf-16, latin-1)"), "указана вручную"
        if a in (1, 2):
            d = ask("Разделитель (символ; для табуляции введите \\t)", ";")
            d = "\t" if d in ("\\t", "tab") else d[0]
            with io.TextIOWrapper(open_binary(path), encoding=enc, errors="replace", newline="") as f:
                text = f.read(SAMPLE_BYTES)
            dia = Dialect(d)
            ncols = sniff_counts(text, dia)
            sample = [(r + [""] * ncols)[:ncols] for r in csv.reader(text.splitlines(True)[:300], **dia.reader_kwargs()) if r]
            preview(sample)
            if ask_yes("Теперь правильно?", True):
                break
    has_header = ask_yes("Первая строка — заголовок?", guess_header(sample))
    src = CsvSource(path, enc, dia, has_header, ncols)
    head = src.read_header() if has_header else []
    src.names = clean_names((head + [""] * ncols)[:ncols] if has_header else [f"col_{i}" for i in range(1, ncols + 1)])
    return src, path


def sniff_counts(text: str, dia: Dialect) -> int:
    counts = [len(r) for r in csv.reader(io.StringIO(text), **dia.reader_kwargs()) if r][:-1] or [1]
    return max(set(counts), key=counts.count)


def preview(sample: list[list[str]], n: int = 5) -> None:
    if not sample:
        return
    k = min(len(sample[0]), 8)
    say("  первые строки (до 8 колонок):")
    print_table([f"#{i + 1}" for i in range(k)], [r[:k] for r in sample[:n]], 22)


# ═════════════════════════════════ АНАЛИЗ ФАЙЛА (шаг 3) ═════════════════════════════════

@dataclass
class Analysis:
    profiles: list[ColumnProfile]
    rows: int
    quote_changes: list[int]
    quote_examples: list[list[tuple[str, str]]]
    unwrap: bool


WORKERS = ThreadPoolExecutor(max_workers=min(8, os.cpu_count() or 2))


def prefetch(gen: Iterator, depth: int = 2) -> Iterator:
    """Читает следующую пачку в фоновом потоке, пока текущая обрабатывается."""
    import queue
    import threading
    q: queue.Queue = queue.Queue(maxsize=depth)
    stop = threading.Event()
    end = object()

    def run():
        try:
            for item in gen:
                while not stop.is_set():
                    try:
                        q.put(item, timeout=0.2)
                        break
                    except queue.Full:
                        continue
                if stop.is_set():
                    return
            q.put(end)
        except BaseException as e:  # noqa: BLE001 — передаём в основной поток
            q.put(e)

    t = threading.Thread(target=run, daemon=True)
    t.start()
    try:
        while True:
            item = q.get()
            if item is end:
                return
            if isinstance(item, BaseException):
                raise item
            yield item
    finally:
        stop.set()
        t.join(timeout=5)


def rows_per_batch(src: Source) -> int:
    if isinstance(src, CsvSource):
        with open_binary(src.path) as f:
            head = f.read(SAMPLE_BYTES)
        lines = max(head.count(b"\n"), 1)
        avg = max(len(head) / lines, 1)
        return int(min(MAX_BATCH_ROWS, max(MIN_BATCH_ROWS, BATCH_BYTES / avg)))
    return 100_000


def analyze(src: Source, unwrap: bool) -> Analysis:
    profiles = [ColumnProfile(n) for n in src.names]
    for p in profiles:
        if src.fixed_types.get(p.name):
            p.fixed = src.fixed_types[p.name]
    qchg = [0] * len(profiles)
    qex: list[list] = [[] for _ in profiles]
    rpb = rows_per_batch(src)
    while True:
        prog = Progress("анализ", getattr(src, "total_bytes", 0))
        rows = 0
        try:
            for cols, n, pos in prefetch(src.iter_batches(rpb)):
                def one(j, arr=None, base=rows):
                    arr = cols[j]
                    new, changed = unwrap_quotes(arr)
                    c = n_true(changed)
                    if c:
                        qchg[j] += c
                        if len(qex[j]) < 3:
                            idx = pc.indices_nonzero(changed).to_pylist()[:3]
                            qex[j] += [(arr[i].as_py(), new[i].as_py()) for i in idx][: 3 - len(qex[j])]
                        if unwrap:
                            arr = new
                    profiles[j].update(arr, base)
                list(WORKERS.map(one, range(len(cols))))   # колонки обрабатываются параллельно
                rows += n
                prog.update(rows, pos)
            prog.done(rows)
            break
        except _NeedRepair as e:
            sys.stderr.write("\n")
            if e.fast:
                say(f"  около строки {e.lineno} формат отличается от строгого CSV — перехожу на построчный разбор")
                src.fast = False
            else:
                warn(f"строка файла {e.lineno}: некорректные кавычки или не то число полей — "
                     "включаю режим исправления (медленнее)")
                src.repair = True
            profiles = [ColumnProfile(n) for n in src.names]
            qchg, qex = [0] * len(profiles), [[] for _ in profiles]
    return Analysis(profiles, rows, qchg, qex, unwrap)


# ═══════════════════════════════════ СХЕМА (шаг 4–5) ═══════════════════════════════════

def show_schema(profiles, specs, include, extra_cols=None):
    rows = []
    for i, (p, s) in enumerate(zip(profiles, specs), 1):
        if not include[i - 1]:
            rows.append([str(i), p.name, "— пропуск —", "", "", ""])
            continue
        note = "; ".join(p.ambiguities())
        if s.kind == "string" and p.dt_pattern() and p.nonnull:
            st = p.dt[p.dt_pattern()]
            note = note or f"даты {st['min']:%Y-%m-%d}…{st['max']:%Y-%m-%d} вне диапазона ClickHouse — текстом"
        elif s.kind == "string" and p.killed_by and p.nonnull:
            # почему не число/дата — первое «мешающее» значение
            # подсказка нужна, только если начало колонки было похоже на число/дату
            c = max((c for c in ("int", "dec.", "dec,", "dt:ymd-", "dt:dmy.") if c in p.killed_by),
                    key=lambda c: p.killed_by[c][0], default=None)
            if c and p.killed_by[c][0] > 0:
                note = note or f"текст из-за {short(p.killed_by[c][1], 18)!r} (стр. {p.killed_by[c][0] + 1})"
        rows.append([str(i), p.name, s.ch, f"{p.nulls:,}", short(p.examples[0] if p.examples else "", 20), note])
    print_table(["#", "колонка", "тип ClickHouse", "пустых", "пример", "примечание"], rows, 40)


def edit_schema(profiles: list[ColumnProfile], specs: list[TypeSpec], include: list[bool]) -> None:
    say("\n  Команды правки (или Enter — принять):\n"
        "    3=Float64        сменить тип колонки 3 (любой тип ClickHouse)\n"
        "    3:новое_имя      переименовать колонку 3\n"
        "    -3 / +3          исключить / вернуть колонку 3")
    while True:
        show_schema(profiles, specs, include)
        cmd = ask("Правка", "").strip()
        if not cmd:
            return
        try:
            if m := re.fullmatch(r"([+-])(\d+)", cmd):
                include[int(m.group(2)) - 1] = m.group(1) == "+"
            elif m := re.fullmatch(r"(\d+)\s*=\s*(.+)", cmd):
                i, spec = int(m.group(1)) - 1, parse_ch_type(m.group(2))
                p = profiles[i]
                ok, why = p.compat(spec)
                if not ok:
                    warn(f"нельзя: {why}")
                    continue
                if why:
                    warn(why)
                if p.nulls and not spec.nullable and spec.kind not in ("string", "fixedstring", "raw"):
                    warn(f"в колонке {p.nulls:,} пустых значений; тип без Nullable — они станут "
                         f"значением по умолчанию ({TYPE_DEFAULT.get(spec.kind, '?')})")
                    if not ask_yes("Согласны?", False):
                        continue
                    spec.fill_default = True
                specs[i] = spec
            elif m := re.fullmatch(r"(\d+)\s*:\s*(.+)", cmd):
                i, name = int(m.group(1)) - 1, m.group(2).strip()
                if name.lower() in {p.name.lower() for j, p in enumerate(profiles) if j != i}:
                    warn("такое имя уже есть")
                    continue
                profiles[i].name = name
            else:
                warn("не понял команду")
        except (IndexError, ValueError):
            warn("нет такой колонки")


def suggest_order_by(profiles, specs, include) -> list[str]:
    """schema-pk-cardinality-order: низкая кардинальность → высокая; без Nullable."""
    lc, dates = [], []
    for p, s, inc in zip(profiles, specs, include):
        if not inc or s.nullable:
            continue
        if s.lowcard and p.distinct is not None:
            lc.append((len(p.distinct), p.name))
        elif s.kind in ("date", "datetime"):
            dates.append(p.name)
    return [n for _, n in sorted(lc)[:2]] + dates[:1]


# ═══════════════════════════════════ ЗАГРУЗКА (шаг 7) ═══════════════════════════════════

def stage_name(target: str, run_id: str) -> str:
    return f"{target[:150]}__{APP}_{run_id}"


def load_into_stage(backend: Backend, src: Source, an: Analysis, table_fq: str, columns: list[str],
                    convs: list[Callable], specs: list[TypeSpec], col_idx: list[int], run_id: str) -> int:
    """Пачки → временная таблица. Конвертация и отправка идут параллельно (2 потока)."""
    rpb = rows_per_batch(src)
    prog = Progress("загрузка", getattr(src, "total_bytes", 0))
    sent = 0
    offset = 0
    pending: deque = deque()

    def send(data, token):
        settings = dict(INSERT_SETTINGS_BASE, insert_deduplication_token=token,
                        date_time_input_format="best_effort" if any(
                            s.kind == "datetime" for s in specs) else "basic")
        for attempt in range(4):
            try:
                backend.insert(table_fq, columns, specs, data, settings)
                return
            except Exception as e:  # noqa: BLE001
                if attempt < 3 and _is_network_error(e):
                    time.sleep(2 ** (attempt + 1))   # повтор безопасен: тот же токен дедупликации
                    continue
                raise

    with ThreadPoolExecutor(max_workers=1) as pool:
        try:
            for batch_no, (cols, n, pos) in enumerate(prefetch(src.iter_batches(rpb))):
                def one(out_i, base=offset):
                    arr = cols[col_idx[out_i]]
                    if an.unwrap:
                        arr = unwrap_quotes(arr)[0]
                    return convs[out_i](arr, base)
                data = list(WORKERS.map(one, range(len(col_idx))))
                offset += n
                pending.append(pool.submit(send, data, f"{run_id}-{batch_no}"))
                while len(pending) >= 2:
                    pending.popleft().result()
                sent += n
                prog.update(sent, pos)
            while pending:
                pending.popleft().result()
        except BaseException:
            for f in pending:
                f.cancel()
            raise
    prog.done(sent)
    return sent


def drop_quietly(backend: Backend, table_fq: str) -> None:
    try:
        backend.command(f"DROP TABLE IF EXISTS {table_fq}")
    except Exception as e:  # noqa: BLE001
        warn(f"не удалось удалить временную таблицу {table_fq}: {e}. Удалите её вручную.")


def save_rejects(path: str, issues: RowIssues) -> Optional[str]:
    if not issues.rejected:
        return None
    for out in (path + ".rejected.txt", os.path.join(os.getcwd(), os.path.basename(path) + ".rejected.txt")):
        try:
            with open(out, "w", encoding="utf-8") as f:
                f.write("# строка_файла\tпричина\tисходный_текст\n")
                for ln, raw, why in issues.rejected:
                    f.write(f"{ln}\t{why}\t{raw.rstrip()}\n")
            return out
        except OSError:
            continue
    return None


# ═══════════════════════════════════════ МАСТЕР ═══════════════════════════════════════

def main() -> int:
    say(__doc__.split("Как обеспечивается")[0])
    backend = connect_wizard()
    src, path = file_wizard()

    # --- «вечные кавычки» по началу файла ---
    step("Шаг 3. Анализ всего файла (кодировка, кавычки, типы)")
    an = analyze(src, unwrap=False)
    total_q = sum(an.quote_changes)
    if total_q:
        say(f"\n  Найдены значения, целиком обёрнутые в лишние кавычки («вечные кавычки»): {total_q:,} шт.")
        for p, c, ex in zip(an.profiles, an.quote_changes, an.quote_examples):
            if c:
                say(f"    «{p.name}»: {c:,} шт., напр. " + ", ".join(f"{short(a, 25)!r} → {short(b, 25)!r}" for a, b in ex))
        if ask_yes("Снять лишние кавычки?", True):
            an = analyze(src, unwrap=True)
    issues = src.issues
    if issues.trimmed_trailing:
        say(f"  У {issues.trimmed_trailing:,} строк в конце были лишние пустые поля — отброшены (данных в них нет).")
    if issues.repaired_count:
        warn(f"Исправлено строк с некорректными кавычками: {issues.repaired_count:,}. Примеры:")
        for ln, raw, fields in issues.repaired[:5]:
            say(f"    строка {ln}: {short(raw.rstrip(chr(13) + chr(10)), 70)}\n      → {[short(x, 30) for x in fields]}")
        if not ask_yes("Принять такое исправление?", True):
            raise Abort("исправление строк отклонено")
    if issues.rejected_count:
        out = save_rejects(path, issues)
        warn(f"Не удалось разобрать строк: {issues.rejected_count:,}" + (f" — сохранены в {out}" if out else ""))
        for ln, raw, why in issues.rejected[:3]:
            say(f"    строка {ln} ({why}): {short(raw, 80)!r}")
        if not ask_yes("Загрузить файл БЕЗ этих строк?", False):
            raise Abort("в файле есть нераспознанные строки")
    say(f"✓ Строк данных: {an.rows:,}, колонок: {len(an.profiles)}")
    if an.rows == 0:
        raise Abort("в файле нет строк данных")

    # пустые колонки без имени (хвостовые разделители) — не загружаем
    include = [not (p.nonnull == 0 and re.fullmatch(r"col_\d+", p.name) and src.kind != "parquet"
                    and getattr(src, "has_header", True)) for p in an.profiles]
    if not all(include):
        say("  Колонки без заголовка и без данных исключены: " +
            ", ".join(p.name for p, inc in zip(an.profiles, include) if not inc))

    # --- куда грузим ---
    step("Шаг 4. Куда загрузить")
    dbs = [r[0] for r in backend.query("SELECT name FROM system.databases ORDER BY name")]
    db = ask("База данных", backend.database)
    if db not in dbs:
        if not ask_yes(f"Базы «{db}» нет. Создать?", False):
            raise Abort("база не выбрана")
        backend.command(f"CREATE DATABASE IF NOT EXISTS {qi(db)}")
    default_tbl = re.sub(r"\W+", "_", os.path.splitext(os.path.basename(path.removesuffix(".gz")))[0]).strip("_").lower() or "data"
    table = ask("Таблица", default_tbl)
    t_fq = f"{qi(db)}.{qi(table)}"
    exists = bool(backend.scalar(
        f"SELECT count() FROM system.tables WHERE database = {ql(db)} AND name = {ql(table)}"))
    run_id = time.strftime("%Y%m%d_%H%M%S")
    stage = stage_name(table, run_id)
    s_fq = f"{qi(db)}.{qi(stage)}"

    if not exists:
        plan = plan_new_table(an, include, db, table, s_fq)
    else:
        plan = plan_existing_table(backend, an, include, db, table)
        plan["stage_ddl"] = (f"CREATE TABLE {s_fq} (" + ", ".join(
            f"{qi(c)} {s.ch}" for c, s in zip(plan["columns"], plan["specs"])) +
            ") ENGINE = MergeTree ORDER BY tuple()")

    # --- итог перед загрузкой ---
    step("Шаг 6. Подтверждение")
    say(f"  Файл:     {path}\n  Строк:    {an.rows:,}\n  Сервер:   {backend.describe()}\n"
        f"  Таблица:  {db}.{table} ({'новая' if not exists else 'существующая, режим: ' + plan['mode_text']})\n"
        f"  Колонок:  {len(plan['columns'])}")
    if not ask_yes("Начать загрузку?", True):
        raise Abort("отменено пользователем")

    # --- загрузка ---
    step("Шаг 7. Загрузка")
    stage_created = False
    try:
        ddl = plan["stage_ddl"]
        try:
            backend.command(ddl + " SETTINGS non_replicated_deduplication_window = 10000")
        except Exception as e:  # noqa: BLE001 — напр. в ClickHouse Cloud настройка не нужна
            if "non_replicated_deduplication_window" not in str(e) and "SETTING" not in str(e).upper():
                raise
            backend.command(ddl)
        stage_created = True
        sent = load_into_stage(backend, src, an, s_fq, plan["columns"], plan["convs"], plan["specs"],
                               plan["col_idx"], run_id)
        in_stage = int(backend.scalar(f"SELECT count() FROM {s_fq}"))
        if in_stage != sent:
            raise RuntimeError(f"во временной таблице {in_stage:,} строк, а отправлено {sent:,} — загрузка остановлена")
        if sent != an.rows:
            raise RuntimeError(f"при загрузке прочитано {sent:,} строк, при анализе было {an.rows:,} — файл изменился?")
        say(f"✓ Временная таблица: {in_stage:,} строк — совпадает с файлом")

        if not exists:
            backend.command(f"RENAME TABLE {s_fq} TO {t_fq}")
            stage_created = False
            try:
                backend.command(f"ALTER TABLE {t_fq} RESET SETTING non_replicated_deduplication_window")
            except Exception:  # noqa: BLE001
                pass
            final = int(backend.scalar(f"SELECT count() FROM {t_fq}"))
            say(f"✓ Таблица {db}.{table} создана, строк: {final:,}")
        else:
            finalize_existing(backend, plan, db, table, t_fq, s_fq, in_stage, run_id)
    except BaseException:
        if stage_created:
            drop_quietly(backend, s_fq)
        raise
    if stage_created:
        drop_quietly(backend, s_fq)
    say("\nГотово.")
    return 0


def plan_new_table(an: Analysis, include: list[bool], db: str, table: str, s_fq: str) -> dict:
    step("Шаг 5. Схема новой таблицы")
    specs = [p.propose() for p in an.profiles]
    amb = [(p, a) for p in an.profiles for a in p.ambiguities()]
    for p, a in amb:
        warn(f"«{p.name}»: {a}")
    say("  Типы подобраны по всему файлу так, чтобы ничего не потерять. Nullable — только там, где есть пустые значения.")
    edit_schema(an.profiles, specs, include)
    for p, a in amb:
        if "тысяч" in a and specs[an.profiles.index(p)].kind == "decimal":
            if not ask_yes(f"«{p.name}»: считать запятую десятичным разделителем (иначе колонка будет String)?", False):
                specs[an.profiles.index(p)] = TypeSpec("string", "String")
    idx = [i for i, inc in enumerate(include) if inc]
    names = [an.profiles[i].name for i in idx]
    sp = [specs[i] for i in idx]

    say("\n  ORDER BY — главный индекс таблицы, после создания не меняется.\n"
        "  Перечислите колонки, по которым чаще всего фильтруете, от редко меняющихся\n"
        "  (категории, коды) к часто меняющимся (дата, id). Пусто = без сортировки (tuple()).")
    sug = suggest_order_by([an.profiles[i] for i in idx], sp, [True] * len(idx))
    while True:
        raw = ask("ORDER BY (имена через запятую)", ", ".join(sug))
        keys = [k.strip() for k in raw.split(",") if k.strip()] if raw.strip() not in ("", "tuple()") else []
        bad = [k for k in keys if k not in names]
        if bad:
            warn(f"нет таких колонок: {bad}")
            continue
        nul = [k for k in keys if sp[names.index(k)].nullable]
        if nul:
            warn(f"колонки с Nullable в ключе сортировки нежелательны: {nul}. Уберите их или смените тип.")
            continue
        break
    say("  PARTITION BY нужен в основном для удаления старых данных по периодам; обычно не нужен.")
    part = ask("PARTITION BY (например toYYYYMM(дата); Enter — без партиций)", "")
    order = "(" + ", ".join(qi(k) for k in keys) + ")" if keys else "tuple()"
    ddl = ("CREATE TABLE {T} (\n    " + ",\n    ".join(f"{qi(n)} {s.ch}" for n, s in zip(names, sp)) +
           "\n) ENGINE = MergeTree\n" + (f"PARTITION BY {part}\n" if part else "") + f"ORDER BY {order}")
    say("\n" + ddl.replace("{T}", f"{qi(db)}.{qi(table)}"))
    if not ask_yes("Создать таблицу с такой структурой?", True):
        raise Abort("структура не подтверждена")
    convs = [make_converter(an.profiles[i], specs[i]) for i in idx]
    return {"columns": names, "specs": sp, "convs": convs, "col_idx": idx,
            "stage_ddl": ddl.replace("{T}", s_fq), "mode_text": "новая"}


def plan_existing_table(backend: Backend, an: Analysis, include: list[bool], db: str, table: str) -> dict:
    step("Шаг 5. Сверка с существующей таблицей")
    engine = backend.scalar(f"SELECT engine FROM system.tables WHERE database = {ql(db)} AND name = {ql(table)}")
    tcols = backend.query(
        f"SELECT name, type, default_kind, default_expression FROM system.columns "
        f"WHERE database = {ql(db)} AND table = {ql(table)} ORDER BY position")
    nrows = backend.scalar(f"SELECT count() FROM {qi(db)}.{qi(table)}")
    say(f"  Таблица {db}.{table}: движок {engine}, строк сейчас: {nrows:,}")
    if engine in ("View", "MaterializedView", "Dictionary", "LiveView", "WindowView"):
        raise Abort(f"в {engine} загружать данные нельзя")
    insertable = {n: (t, k, e) for n, t, k, e in tcols if k not in ("MATERIALIZED", "ALIAS")}
    lower = {n.lower(): n for n in insertable}

    columns, specs, convs, col_idx, problems = [], [], [], [], []
    for i, (p, inc) in enumerate(zip(an.profiles, include)):
        if not inc:
            continue
        tname = p.name if p.name in insertable else lower.get(p.name.lower())
        if tname is None:
            problems.append(("missing", i, None, "колонки нет в таблице"))
            continue
        spec = parse_ch_type(insertable[tname][0])
        ok, why = p.compat(spec)
        row = [p.name, tname, spec.ch, "ok" if ok else "КОНФЛИКТ", why]
        if not ok:
            problems.append(("type", i, tname, why))
        elif p.nulls and not spec.nullable and spec.kind not in ("string", "fixedstring", "raw"):
            problems.append(("nulls", i, tname, f"{p.nulls:,} пустых значений, а тип {spec.ch} не допускает NULL"))
        columns.append(tname)
        specs.append(spec)
        col_idx.append(i)
        say(f"  {row[0]!s:>24} → {row[1]!s:<24} {row[2]:<28} {row[3]} {row[4]}")
    missing_in_file = [n for n in insertable if n not in columns]

    for kind, i, tname, why in problems:
        p = an.profiles[i]
        if kind == "missing":
            warn(f"Колонки «{p.name}» нет в таблице.")
            a = choose("Что делать", ["не загружать эту колонку", f"добавить колонку в таблицу (ALTER TABLE ADD COLUMN, тип {p.propose().ch})", "отменить загрузку"])
            if a == 2:
                raise Abort("отменено")
            if a == 1:
                spec = p.propose()
                backend.command(f"ALTER TABLE {qi(db)}.{qi(table)} ADD COLUMN IF NOT EXISTS {qi(p.name)} {spec.ch}")
                say(f"  ✓ колонка {p.name} {spec.ch} добавлена")
                pos = sorted(col_idx + [i]).index(i)
                columns.insert(pos, p.name)
                specs.insert(pos, spec)
                col_idx.insert(pos, i)
        elif kind == "type":
            warn(f"«{p.name}» → {tname}: {why}")
            a = choose("Что делать", ["не загружать эту колонку", "отменить загрузку"], 2)
            if a == 1:
                raise Abort(f"конфликт типов в колонке {p.name}")
            k = col_idx.index(i)
            for lst in (columns, specs, col_idx):
                lst.pop(k)
        elif kind == "nulls":
            k = col_idx.index(i)
            warn(f"«{p.name}»: {why}")
            a = choose("Что делать", [f"записать вместо пустых значение по умолчанию ({TYPE_DEFAULT.get(specs[k].kind)})",
                                      "не загружать эту колонку", "отменить загрузку"], 3)
            if a == 2:
                raise Abort("пустые значения в не-Nullable колонке")
            if a == 0:
                specs[k].fill_default = True
            else:
                for lst in (columns, specs, col_idx):
                    lst.pop(k)
    if not columns:
        raise Abort("нет ни одной колонки для загрузки")
    missing_in_file = [n for n in insertable if n not in columns]
    if missing_in_file:
        say("  Колонки таблицы, которых нет в файле (получат значения по умолчанию):")
        for n in missing_in_file:
            t, k, e = insertable[n]
            say(f"    {n} {t}" + (f" {k} {e}" if k else " — значение по умолчанию типа (0 / пусто)"))
        if not ask_yes("Продолжить?", True):
            raise Abort("отменено")

    mode, mode_text = "append", "добавление"
    if nrows:
        a = choose(f"В таблице уже {nrows:,} строк. Как загрузить", [
            "добавить строки к существующим",
            "заменить данные таблицы (старые останутся в резервной таблице)",
            "отменить"])
        if a == 2:
            raise Abort("отменено")
        if a == 1:
            confirm = ask(f"Для замены введите имя таблицы «{table}»")
            if confirm != table:
                raise Abort("подтверждение замены не совпало")
            mode, mode_text = "replace", "замена (с резервной копией)"
    convs = [make_converter(an.profiles[i], s) for i, s in zip(col_idx, specs)]
    return {"columns": columns, "specs": specs, "convs": convs, "col_idx": col_idx,
            "mode": mode, "mode_text": mode_text, "nrows": nrows}


def finalize_existing(backend: Backend, plan: dict, db: str, table: str, t_fq: str, s_fq: str,
                      in_stage: int, run_id: str) -> None:
    cols = ", ".join(qi(c) for c in plan["columns"])
    if plan["mode"] == "append" and plan["nrows"]:
        # «опасный конфликт»: не загружался ли этот файл уже раньше?
        say("  Проверяю, нет ли этих строк в таблице уже (выборка 1000 строк)…")
        h = f"cityHash64(tuple({cols}))"
        try:
            sample = backend.scalar(f"SELECT uniqExact(h) FROM (SELECT {h} AS h FROM {s_fq} LIMIT 1000)")
            found = backend.scalar(
                f"SELECT uniqExact(h) FROM (SELECT {h} AS h FROM {t_fq} WHERE h IN "
                f"(SELECT {h} FROM {s_fq} LIMIT 1000))", settings={"max_execution_time": 120})
            if found:
                warn(f"{found} из {sample} проверенных строк файла УЖЕ ЕСТЬ в таблице — возможно, файл загружался раньше.")
                if not ask_yes("Всё равно добавить (будут дубликаты)?", False):
                    raise Abort("возможные дубликаты")
            else:
                say("  ✓ совпадений не найдено")
        except Abort:
            raise
        except Exception as e:  # noqa: BLE001
            warn(f"проверку дубликатов выполнить не удалось ({short(str(e), 120)})")
            if not ask_yes("Продолжить без проверки?", True):
                raise Abort("отменено")

    settings = {"insert_deduplication_token": f"{run_id}-final", "max_execution_time": 0,
                "max_insert_threads": 4}
    if plan["mode"] == "append":
        written = backend.command(f"INSERT INTO {t_fq} ({cols}) SELECT {cols} FROM {s_fq}", settings=settings)
        report_written(written, in_stage)
        say(f"✓ Добавлено в {db}.{table}: {in_stage:,} строк. Всего в таблице: "
            f"{backend.scalar(f'SELECT count() FROM {t_fq}'):,}")
        return
    # замена: новая таблица той же структуры → EXCHANGE (атомарно) → старая становится резервной
    new_fq = f"{qi(db)}.{qi(stage_name(table, run_id) + '_new')}"
    backup = f"{table[:150]}__backup_{run_id}"
    try:
        backend.command(f"CREATE TABLE {new_fq} AS {t_fq}")
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"не удалось создать копию структуры таблицы для замены: {e}. Таблица не изменена.")
    try:
        written = backend.command(f"INSERT INTO {new_fq} ({cols}) SELECT {cols} FROM {s_fq}", settings=settings)
        report_written(written, in_stage)
        cnt = backend.scalar(f"SELECT count() FROM {new_fq}")
        if cnt != in_stage:
            raise RuntimeError(f"в новой таблице {cnt:,} строк вместо {in_stage:,}")
        backend.command(f"EXCHANGE TABLES {t_fq} AND {new_fq}")
    except BaseException:
        drop_quietly(backend, new_fq)
        raise
    backend.command(f"RENAME TABLE {new_fq} TO {qi(db)}.{qi(backup)}")
    say(f"✓ Данные {db}.{table} заменены: {in_stage:,} строк.\n"
        f"  Старые данные сохранены в {db}.{backup} — удалите её, когда убедитесь, что всё в порядке:\n"
        f"    DROP TABLE {qi(db)}.{qi(backup)}")


def report_written(written: int, expected: int) -> None:
    if written < 0:
        return
    if written < expected:
        raise RuntimeError(f"сервер записал {written:,} строк из {expected:,}")
    if written > expected:
        say(f"  (сервер сообщил о {written:,} записанных строках — включая материализованные представления)")


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Abort as e:
        say(f"\n✗ Остановлено: {e}. Целевая таблица не изменена.")
        sys.exit(1)
    except KeyboardInterrupt:
        say("\n✗ Прервано (Ctrl+C). Временная таблица удалена, целевая не изменена.")
        sys.exit(130)
    except ConvertError as e:
        say(f"\n✗ Ошибка данных: {e}\n  Целевая таблица не изменена (временная удалена).\n"
            "  Если значение корректно — запустите снова и задайте этой колонке тип String.")
        sys.exit(2)
    except Exception as e:  # noqa: BLE001
        say(f"\n✗ Ошибка: {type(e).__name__}: {e}\n  Целевая таблица не изменена (временная удалена).")
        sys.exit(3)
