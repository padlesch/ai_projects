"""Интеграционные тесты с настоящим ClickHouse.

Запуск:  CH_TEST_HOST=localhost pytest tests/test_integration.py
Необязательно: CH_TEST_HTTP_PORT (8123), CH_TEST_TCP_PORT (9000), CH_TEST_USER, CH_TEST_PASSWORD.
Тесты создают и удаляют таблицы в базе default с префиксом ch_loader_test_.
"""
import os
import time

import pytest

import ch_loader as L
from conftest import csv_source

HOST = os.environ.get("CH_TEST_HOST")
pytestmark = pytest.mark.skipif(not HOST, reason="не задан CH_TEST_HOST — нет тестового ClickHouse")

USER = os.environ.get("CH_TEST_USER", "default")
PASSWORD = os.environ.get("CH_TEST_PASSWORD", "")


def make_backend(protocol):
    if protocol == "http":
        port = int(os.environ.get("CH_TEST_HTTP_PORT", 8123))
        return L.HttpBackend(HOST, port, USER, PASSWORD, "default", False, True, 10)
    pytest.importorskip("clickhouse_driver")
    port = int(os.environ.get("CH_TEST_TCP_PORT", 9000))
    return L.NativeBackend(HOST, port, USER, PASSWORD, "default", False, True, 10)


@pytest.fixture(params=["http", "native"])
def backend(request):
    return make_backend(request.param)


@pytest.fixture
def table(backend):
    name = f"ch_loader_test_{backend.protocol}_{int(time.time() * 1000)}"
    yield name
    backend.command(f"DROP TABLE IF EXISTS default.{L.qi(name)}")


def load_file(backend, path, table, delimiter=","):
    """Анализ + загрузка во временную таблицу по той же схеме, что и мастер."""
    src = csv_source(path, delimiter)
    an = L.analyze(src, unwrap=True)
    specs = [p.propose() for p in an.profiles]
    names = [p.name for p in an.profiles]
    fq = f"default.{L.qi(table)}"
    backend.command(f"CREATE TABLE {fq} (" + ", ".join(f"{L.qi(n)} {s.ch}" for n, s in zip(names, specs)) +
                    ") ENGINE = MergeTree ORDER BY tuple() SETTINGS non_replicated_deduplication_window = 100")
    convs = [L.make_converter(p, s) for p, s in zip(an.profiles, specs)]
    sent = L.load_into_stage(backend, src, an, fq, names, convs, specs, list(range(len(names))), "t")
    return an, sent, fq


def test_round_trip_all_types(backend, table, write_file):
    """Файл со всеми основными типами загружается, и значения в ClickHouse совпадают с исходными."""
    path = write_file("a.csv", (
        "id,code,name,amount,day,ts,flag,opt\n"
        '1,00123,"""ООО """"Ромашка""""""",1.25,31.01.2024,2024-01-31 10:00:00.5,true,\n'
        '2,00456,"Текст с, запятой",-7.5,01.02.2024,2024-02-01 00:00:00,false,42\n'
    ))
    an, sent, fq = load_file(backend, path, table)
    assert sent == 2 == backend.scalar(f"SELECT count() FROM {fq}")
    rows = backend.query(f"SELECT toString(id), code, name, toString(amount), toString(day), "
                         f"toString(ts), toString(flag), ifNull(toString(opt), 'NULL') FROM {fq} ORDER BY id")
    assert [list(r) for r in rows] == [
        ["1", "00123", 'ООО "Ромашка"', "1.25", "2024-01-31", "2024-01-31 10:00:00.500", "true", "NULL"],
        ["2", "00456", "Текст с, запятой", "-7.5", "2024-02-01", "2024-02-01 00:00:00.000", "false", "42"],
    ]
    assert backend.scalar(f"SELECT countIf(opt IS NULL) FROM {fq}") == 1


def test_retry_with_same_token_does_not_duplicate(backend, table, write_file):
    """Повтор пачки с тем же токеном дедупликации (как при обрыве сети) не создаёт дублей."""
    path = write_file("a.csv", "id,v\n1,a\n2,b\n")
    an, sent, fq = load_file(backend, path, table)
    src = csv_source(path)
    specs = [p.propose() for p in an.profiles]
    convs = [L.make_converter(p, s) for p, s in zip(an.profiles, specs)]
    L.load_into_stage(backend, src, an, fq, ["id", "v"], convs, specs, [0, 1], "t")   # тот же run_id
    assert backend.scalar(f"SELECT count() FROM {fq}") == 2


def test_value_outside_type_stops_load_before_writing(backend, table, write_file):
    """Значение, не влезающее в тип (999 в UInt8), останавливает загрузку до отправки пачки:
    в таблицу не попадает ни одной строки."""
    good = write_file("a.csv", "id\n1\n2\n")
    an = L.analyze(csv_source(good), unwrap=False)
    spec = L.parse_ch_type("UInt8")
    fq = f"default.{L.qi(table)}"
    backend.command(f"CREATE TABLE {fq} (id UInt8) ENGINE = MergeTree ORDER BY tuple()")
    conv = L.make_converter(an.profiles[0], spec)
    bad = csv_source(write_file("b.csv", "id\n1\n999\n"))
    with pytest.raises(L.ConvertError, match="вне диапазона"):
        L.load_into_stage(backend, bad, an, fq, ["id"], [conv], [spec], [0], "x")
    assert backend.scalar(f"SELECT count() FROM {fq}") == 0
