"""Проверка на настоящей сводной из файла пользователя.

python3 real.py путь/к/книге.xlsm

Из файла берётся то, что макрос читает с листа сводной (подписи строк с
отступами, значения, шапка столбцов), прогоняется через макрос (ядро +
вывод на лист в эмуляции Excel) и печатается рядом с исходной сводной.
Два режима: сводная сообщает типы ячеек; сводная не сообщает ничего.
"""
import os, re, sys, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import openpyxl
import run_out
import excel_merge
from runbasic import Basic

RS, US, GS = chr(30), chr(31), chr(29)
PALETTE = 1


def ref(a):
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', a)
    col = lambda s: sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(s)))
    return col(m[1]), int(m[2]), col(m[3]), int(m[4])


def load(path):
    z = zipfile.ZipFile(path)
    pt = z.read('xl/pivotTables/pivotTable1.xml').decode('utf-8')
    loc = re.search(r'<location ref="([^"]+)" firstHeaderRow="(\d+)" firstDataRow="(\d+)" firstDataCol="(\d+)"', pt)
    c1, r1, c2, r2 = ref(loc[1])
    fdr, fdc = int(loc[3]), int(loc[4])
    meas = re.findall(r'<dataField name="([^"]+)"', pt)
    def lines(tag):
        blk = re.search(rf'<{tag}.*?</{tag}>', pt, re.S)
        out = []
        for m in re.finditer(r'<i(\s[^>]*)?>(.*?)</i>|<i(\s[^>]*)?/>', blk[0] if blk else ''):
            attrs = (m[1] or m[3] or '')
            r = int((re.search(r'\br="(\d+)"', attrs) or [0, 0])[1])
            t = (re.search(r'\bt="(\w+)"', attrs) or [0, ''])[1]
            depth = r + len(re.findall(r'<x\b', m[2] or ''))
            out.append((depth, t))
        return out
    rows_items = lines('rowItems')
    col_items = lines('colItems')
    wb = openpyxl.load_workbook(path)
    ws = [w for w in wb.worksheets if w._pivots][0]
    return pt, (c1, r1, c2, r2, fdr, fdc), meas, (rows_items, col_items), ws


def encode(path, with_types):
    pt, (c1, r1, c2, r2, fdr, fdc), meas, (rows_items, col_items), ws = load(path)
    LTMAP = {'': 0, 'default': 1, 'sum': 1, 'grand': 2, 'blank': 3}
    nlbl = fdc
    dr, dc = r1 + fdr, c1 + fdc          # первая строка и столбец значений
    nR, nV = r2 - dr + 1, c2 - dc + 1
    caps = [str(ws.cell(dr - 1, c1 + k).value or f'Уровень {k + 1}') for k in range(nlbl)]
    служ = '|названия столбцов|column labels|значения|values|' + '|'.join(
        str(ws.cell(r, c).value).lower() for r in range(r1, dr) for c in range(dc, c2 + 1)
        if ws.cell(r, c).value and r == r1) + '|'
    head = [str(nR), str(nlbl), str(nV), str(dr - r1), str(len(meas))] + meas + [служ] + caps
    rows = []
    for i in range(nR):
        r = dr + i
        dep, tt = rows_items[i]
        nxt = rows_items[i + 1][0] if i + 1 < len(rows_items) else 0
        top = tt == '' and nxt > dep          # родитель при итогах вверху — линия промитога
        lt = (LTMAP.get(tt, 0) if tt else (1 if top else 0)) if with_types else -1
        parts = [str(lt)]
        for k in range(nlbl):
            cell = ws.cell(r, c1 + k)
            v = cell.value
            t = '' if v is None else str(v)
            ind = int(cell.alignment.indent or 0)
            ct = ((3 if tt == 'grand' else 2 if (top or tt) else 1) if (with_types and t) else -1)
            parts.append(US.join([t, str(ct), str(ind)]))
        for j in range(nV):
            v = ws.cell(r, dc + j).value
            parts.append('' if v is None else repr(float(v)))
        rows.append(RS.join(parts))
    hdr = []
    for r in range(r1, dr):
        parts = []
        for j in range(nV):
            v = ws.cell(r, dc + j).value
            t = '' if v is None else str(v)
            ctt = col_items[j][1] if j < len(col_items) else ''
            if with_types:
                ct = 8 if (r == r1 and t == 'Значения') else 5 if r == r1 else 4 if t in meas else \
                    3 if ctt == 'grand' else 2 if ctt else 1
                if not t:
                    ct = -1
            else:
                ct = -1
            parts.append(US.join([t, str(ct)]))
        hdr.append(RS.join(parts))
    clt = RS.join([str(LTMAP.get(col_items[j][1], 0)) if with_types else '-1' for j in range(nV)])
    return GS.join([RS.join(head), '\n'.join(rows), '\n'.join(hdr), clt, str(PALETTE)]), ws, (c1, r1, c2, r2, dr, dc)


def show_src(ws, box):
    c1, r1, c2, r2, dr, dc = box
    print('  сводная на листе:')
    for r in range(r1, r2 + 1):
        cells = []
        for c in range(c1, c2 + 1):
            v = ws.cell(r, c).value
            s = '' if v is None else (f'{v:.0f}' if isinstance(v, float) else str(v))
            cells.append(f'{s[:13]:13}')
        print('    ' + '|'.join(cells))


def show_colors(out, names):
    """Заливки строк: номер, подпись, ключевая колонка (итог), обычная колонка."""
    rows = [l for l in out.split('\n') if l]
    for ln in rows:
        cells = [x.split(US) for x in ln.split(RS) if x]
        bgr = lambda h: h[-6:][4:6] + h[-6:][2:4] + h[-6:][0:2] if len(h) >= 6 else h
        col = [('—' if c[1].endswith('FFFFFF') else '#' + bgr(c[1])) for c in cells]
        lab = next((c[0] for c in cells[1:4] if c[0].strip('*<')), '')
        print(f'    {lab[:24]:24} №:{col[0]:8} подпись:{col[1]:8} итог-колонка:{col[4]:8} обычная:{col[7] if len(col) > 7 else "":8}')


def show_rep(out):
    if out.startswith('ОШИБКА'):
        print(out); return
    print('  отчёт (* жирный, · отступ, < объединена с ячейкой левее/выше):')
    for ln in out.split('\n'):
        if not ln:
            continue
        cells = [x.split(US)[0] for x in ln.split(RS) if x]
        print('    ' + '|'.join(f'{c[:13]:13}' for c in cells))


if __name__ == '__main__':
    path = sys.argv[1]
    src = run_out.build_src() + open(os.path.join(HERE, 'harness_real.bas'), encoding='utf-8').read()
    b = Basic(2360)
    try:
        b.load(src)
        outs = []
        for wt in (True, False):
            data, ws, box = encode(path, wt)
            if wt:
                show_src(ws, box)
            out = b.call('ИзФайла', data)
            outs.append(out)
            print('=' * 60, 'сводная сообщает типы ячеек' if wt else 'сводная НЕ сообщает ничего')
            show_rep(out)
            lost, g = excel_merge.lost_labels(b.call('ПланТекстомР'))
            print('  шапка, как её покажет Excel после объединения ячеек:')
            for row in g:
                print('    ' + '|'.join(f'{c[:13]:13}' for c in row))
            print('  подписи, стёртые объединением:', lost if lost else 'нет')
            if lost:
                outs.append('потеря')
        print('режимы совпали' if outs[0] == outs[1] else 'РЕЖИМЫ РАЗЛИЧАЮТСЯ')
    finally:
        b.close()
