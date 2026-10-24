"""Ручная проверка: сводные, записанные руками так, как их показывает Excel,
проходят через макрос (разбор + вывод на лист в эмуляции Excel), и готовый
отчёт печатается таблицей — чтобы посмотреть глазами.

python3 manual.py
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import run_out  # noqa: сборка модуля с заглушкой меток — та же, что в run_out
from runbasic import Basic

RS, US = chr(30), chr(31)
REG, SUB, GRAND = 0, 1, 2
CI, CS, CG = 1, 2, 3     # тип ячейки: элемент, итог, общий итог

# ячейка: (текст, тип, номер поля, отступ); пустая — None
def row(lt, cells, vals):
    return lt, cells, vals

N = None
CASES = {
 'СТРУКТУРА, итоги вверху (по умолчанию в Excel)': dict(
    caps=['Регион', 'Город', 'Месяц'], cols=['Янв', 'Фев', 'Общий итог'],
    rows=[
     row(SUB, [('Север', CS, 1, 0), N, N], [30, 40, 70]),
     row(SUB, [N, ('Москва', CS, 2, 0), N], [10, 20, 30]),
     row(REG, [N, N, ('Неделя 1', CI, 3, 0)], [4, 8, 12]),
     row(REG, [N, N, ('Неделя 2', CI, 3, 0)], [6, 12, 18]),
     row(SUB, [N, ('Тверь', CS, 2, 0), N], [20, 20, 40]),
     row(REG, [N, N, ('Неделя 1', CI, 3, 0)], [20, 20, 40]),
     row(SUB, [('Юг', CS, 1, 0), N, N], [5, 5, 10]),
     row(SUB, [N, ('Москва', CS, 2, 0), N], [5, 5, 10]),   # тёзка из другой группы
     row(REG, [N, N, ('Неделя 1', CI, 3, 0)], [5, 5, 10]),
     row(GRAND, [('Общий итог', CG, 0, 0), N, N], [35, 45, 80]),
    ]),
 'СТРУКТУРА, итоги внизу': dict(
    caps=['Регион', 'Город'], cols=['Янв', 'Фев', 'Общий итог'],
    rows=[
     row(REG, [('Север', CI, 1, 0), N], []),
     row(REG, [N, ('Москва', CI, 2, 0)], [10, 20, 30]),
     row(REG, [N, ('Тверь', CI, 2, 0)], [20, 20, 40]),
     row(SUB, [('Север Итог', CS, 1, 0), N], [30, 40, 70]),
     row(REG, [('Юг', CI, 1, 0), N], []),
     row(REG, [N, ('Сочи', CI, 2, 0)], [5, 5, 10]),
     row(SUB, [('Юг Итог', CS, 1, 0), N], [5, 5, 10]),
     row(GRAND, [('Общий итог', CG, 0, 0), N], [35, 45, 80]),
    ]),
 'ТАБЛИЧНАЯ, итоги внизу, подписи не повторяются': dict(
    caps=['Регион', 'Город'], cols=['Янв', 'Фев', 'Общий итог'],
    rows=[
     row(REG, [('Север', CI, 1, 0), ('Москва', CI, 2, 0)], [10, 20, 30]),
     row(REG, [N, ('Тверь', CI, 2, 0)], [20, 20, 40]),
     row(SUB, [('Север Итог', CS, 1, 0), N], [30, 40, 70]),
     row(REG, [('Юг', CI, 1, 0), ('Сочи', CI, 2, 0)], [5, 5, 10]),
     row(SUB, [('Юг Итог', CS, 1, 0), N], [5, 5, 10]),
     row(GRAND, [('Общий итог', CG, 0, 0), N], [35, 45, 80]),
    ]),
 'ТАБЛИЧНАЯ без итогов, подписи повторяются': dict(
    caps=['Регион', 'Город'], cols=['Янв', 'Фев', 'Общий итог'],
    rows=[
     row(REG, [('Север', CI, 1, 0), ('Москва', CI, 2, 0)], [10, 20, 30]),
     row(REG, [('Север', CI, 1, 0), ('Тверь', CI, 2, 0)], [20, 20, 40]),
     row(REG, [('Юг', CI, 1, 0), ('Сочи', CI, 2, 0)], [5, 5, 10]),
     row(GRAND, [('Общий итог', CG, 0, 0), N], [35, 45, 80]),
    ]),
 'СЖАТАЯ, итоги вверху': dict(
    caps=['Названия строк'], cols=['Янв', 'Фев', 'Общий итог'],
    rows=[
     row(SUB, [('Север', CS, 1, 0)], [30, 40, 70]),
     row(REG, [('Москва', CI, 2, 1)], [10, 20, 30]),
     row(REG, [('Тверь', CI, 2, 1)], [20, 20, 40]),
     row(SUB, [('Юг', CS, 1, 0)], [5, 5, 10]),
     row(REG, [('Сочи', CI, 2, 1)], [5, 5, 10]),
     row(GRAND, [('Общий итог', CG, 0, 0)], [35, 45, 80]),
    ]),
}


def encode(c):
    lines = [f"{len(c['rows'])} {len(c['caps'])} {len(c['cols'])}"]
    for lt, cells, vals in c['rows']:
        parts = [str(lt)]
        for x in cells:
            x = x or ('', -1, 0, 0)
            parts.append(US.join(map(str, x)))
        vv = list(vals) + [''] * (len(c['cols']) - len(vals))
        parts += [str(v) for v in vv]
        lines.append(RS.join(parts))
    lines.append(RS.join(c['cols']))
    lines.append(RS.join(c['caps']))
    return '\n'.join(lines)


def show_src(c):
    print('  сводная на листе:')
    for lt, cells, vals in c['rows']:
        lab = ' | '.join(f'{(x[0] if x else ""):12}' for x in cells)
        print(f'    {lab} || ' + '  '.join(f'{v:>4}' for v in vals))


def show_rep(out):
    print('  отчёт (* жирный, · отступ, [цвет заливки]):')
    for ln in out.split('\n'):
        if not ln:
            continue
        cells = [x.split(US) for x in ln.split(RS) if x]
        txt = []
        for t, col in cells:
            fill = '' if col in ('FFFFFFFF', 'FFFFFF') else f'[{col[-6:]}]'
            txt.append(f'{t + fill:24}')
        print('    ' + '|'.join(txt))


if __name__ == "__main__":
    src = run_out.build_src() + open(os.path.join(HERE, 'harness_manual.bas'), encoding='utf-8').read()
    b = Basic(2330)
    try:
        b.load(src)
        for name, c in CASES.items():
            print('=' * 100)
            print(name)
            show_src(c)
            show_rep(b.call('Построить', encode(c), 1))
    finally:
        b.close()
