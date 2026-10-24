"""Общие фикстуры тестов загрузчика.

Модульные тесты работают без ClickHouse. Интеграционные (test_integration.py)
запускаются, только если задана переменная окружения CH_TEST_HOST
(например CH_TEST_HOST=localhost, порт HTTP 8123 по умолчанию).
"""
import os
import sys

import pyarrow as pa
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ch_loader as L  # noqa: E402


@pytest.fixture
def write_file(tmp_path):
    """Создаёт файл с заданным содержимым (str или bytes) и возвращает путь."""

    def _write(name, content, encoding="utf-8"):
        path = tmp_path / name
        if isinstance(content, bytes):
            path.write_bytes(content)
        else:
            path.write_bytes(content.encode(encoding))
        return str(path)

    return _write


def profile_of(values, name="col"):
    """Профиль колонки по списку строковых значений."""
    p = L.ColumnProfile(name)
    p.update(pa.array(values, pa.string()), 0)
    return p


def csv_source(path, delimiter=",", has_header=True, ncols=None, encoding="utf-8"):
    """CsvSource с автоопределением числа колонок по заголовку."""
    dia = L.Dialect(delimiter)
    if ncols is None:
        with open(path, encoding=encoding) as f:
            ncols = len(f.readline().rstrip("\r\n").split(delimiter))
    src = L.CsvSource(path, encoding, dia, has_header, ncols)
    src.names = L.clean_names(src.read_header()) if has_header else [f"c{i}" for i in range(ncols)]
    return src


def read_all(src, rows_per_batch=10_000):
    """Все строки источника в виде списка списков (для сравнения)."""
    out = []
    for cols, n, _ in src.iter_batches(rows_per_batch):
        out.extend(map(list, zip(*[c.to_pylist() for c in cols])))
    return out
