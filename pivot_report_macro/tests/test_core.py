"""Автотесты ядра макроса: структура строк и шапки в любых макетах сводной.

Запуск:  python3 test_core.py [число случайных кейсов]
"""
import itertools, os, random, re, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '..', 'src', 'modОтчетИзСводнойV3.bas')
sys.path.insert(0, HERE)
from runbasic import Basic

RS, US, GS = chr(30), chr(31), chr(29)
LT_REG, LT_SUB, LT_GRAND, LT_BLANK = 0, 1, 2, 3
CT_ITEM, CT_SUB, CT_GRAND = 1, 2, 3


class Node:
    def __init__(self, name, children=()):
        self.name, self.children = name, list(children)


NAMES = ['Север', 'Юг', 'a', 'b', 'Итоговый', '2024', '1-2', 'Total Sales', 'x', 'y', '007', 'Б']


def gen_tree(rnd, depth, maxleaves):
    budget = [maxleaves]

    def mk(d):
        if d == depth:
            return []
        k = rnd.randint(1, 3)
        names = rnd.sample(NAMES, k)          # братья различны, кузены — часто тёзки
        out = []
        for nm in names:
            if budget[0] <= 0 and out:
                break
            ch = mk(d + 1)
            if d + 1 == depth:
                budget[0] -= 1
            out.append(Node(nm, ch))
        return out
    return Node('ROOT', mk(0))


def render(root, depth, layout, subt, repeat, grand, info, blanks):
    """Сводная на листе: список строк + эталон (путь узла каждой строки)."""
    ncol = 1 if layout == 'compact' else depth
    rows = []   # (lt, hasval, cells[(txt,ct,lev,ind)], kind, path)

    def cells_empty():
        return [['', -1, 0, 0] for _ in range(ncol)]

    def put(cells, d, txt, ct):
        if layout == 'compact':
            cells[0] = [txt, ct, d, d - 1]
        else:
            cells[d - 1] = [txt, ct, d, 0]

    def walk(node, path, shown):
        d = len(path)
        leaf = not node.children
        own_row = leaf or layout != 'tabular'
        if own_row:
            c = cells_empty()
            if layout != 'compact':
                for a in range(1, d):
                    if repeat or not shown[a - 1]:
                        put(c, a, path[a - 1], CT_ITEM)
                        shown[a - 1] = True
            # Итоги вверху: строка родителя — это и есть итог, Excel отдаёт её
            # как линию промитога, подпись — просто имя без «Итог»
            top_sub = (not leaf) and subt == 'top'
            put(c, d, node.name, CT_SUB if top_sub else CT_ITEM)
            hv = leaf or subt == 'top'
            kind = 'node' if (leaf or subt != 'bottom') else 'header'
            rows.append([LT_SUB if top_sub else LT_REG, hv, c, kind, tuple(path)])
            for a in range(len(shown)):
                if a >= d - 1:
                    shown[a] = True
        for ch in node.children:
            sh = [True] * (d) + [False] * (depth - d)
            if layout == 'tabular':
                # родители видны только на первой строке группы
                sh = shown[:d] + [False] * (depth - d)
            walk(ch, path + [ch.name], sh)
            shown = [True] * depth
        if node.children and subt == 'bottom':
            c = cells_empty()
            put(c, d, node.name + ' Total', CT_SUB)
            if layout == 'tabular' and repeat:
                for a in range(1, d):
                    put(c, a, path[a - 1], CT_ITEM)
            rows.append([LT_SUB, True, c, 'node', tuple(path)])
        if d == 1 and blanks:
            rows.append([LT_BLANK, False, cells_empty(), 'blank', ()])

    for ch in root.children:
        walk(ch, [ch.name], [False] * depth)
    if grand:
        c = cells_empty()
        c[0] = ['Grand Total', CT_GRAND, 0, 0]
        rows.append([LT_GRAND, True, c, 'grand', ()])

    # что знает «Excel» о ячейках
    for r in rows:
        if info == 'geom':        # ничего: только текст и отступ
            r[0] = -1
            for cl in r[2]:
                cl[1], cl[2] = -1, 0
        elif info == 'lt':        # только типы линий оси
            for cl in r[2]:
                cl[1], cl[2] = -1, 0
        elif info == 'field':     # типы и поля ячеек, без оси
            r[0] = -1
        elif info == 'partial':   # поле известно, кроме ячеек итогов
            r[0] = -1
            for cl in r[2]:
                if cl[1] == CT_SUB:
                    cl[2] = 0
        for cl in r[2]:
            if not cl[0]:
                cl[1], cl[2] = -1, 0
    return rows, ncol


def expected(root, layout, subt, grand):
    seq = []
    if grand:
        seq.append(('grand', ()))

    def walk(node, path):
        leaf = not node.children
        if leaf or layout != 'tabular' or subt == 'bottom':
            seq.append(('node', tuple(path)))
        for ch in node.children:
            walk(ch, path + [ch.name])
    for ch in root.children:
        walk(ch, [ch.name])
    return seq


def encode(rows, ncol):
    lines = [f'{len(rows)} {ncol}']
    for lt, hv, cells, kind, path in rows:
        parts = [str(lt), '1' if hv else '0']
        for t, ct, lev, ind in cells:
            parts.append(US.join([t, str(ct), str(lev), str(ind)]))
        lines.append(RS.join(parts))
    return '\n'.join(lines)


def check(case, rows, ncol, out, root):
    depth_, layout, subt, repeat, grand, info, blanks = case
    exp = expected(root, layout, subt, grand)
    got_rows = [l for l in out.split('\n') if l]
    got = []
    errs = []
    for ln in got_rows:
        f = ln.split(RS)
        src = int(f[0]) - 1
        r = rows[src]
        got.append((r[3] if r[3] != 'header' else 'HEADER', r[4]))
    if got != exp:
        errs.append(f'порядок:\n  ждали {exp}\n  вышло {got}')
        return errs
    owned = {p for k, p in exp if k == 'node'}
    prev = None
    for ln, (kind, path) in zip(got_rows, exp):
        f = ln.split(RS)
        dep, typ, olv, cls = int(f[1]), int(f[2]), int(f[3]), f[4]
        labels = [x.split(US) for x in f[5:]]
        if kind == 'grand':
            if cls != 'G':
                errs.append(f'общий итог без класса G: {ln!r}')
            prev = None
            continue
        d = len(path)
        if dep != d:
            errs.append(f'{path}: глубина {dep}, ждали {d}')
        has_child = any(p[:d] == path and len(p) > d for p in owned) or \
            (layout == 'tabular' and d < depth_)
        if (cls == 'S') != has_child:
            errs.append(f'{path}: класс {cls!r}, детей {has_child}')
        anc_owned = sum(1 for a in range(1, d) if path[:a] in owned)
        if olv != min(1 + anc_owned, 8):
            errs.append(f'{path}: уровень группировки {olv}, ждали {1 + anc_owned}')
        col = 0 if layout == 'compact' else d - 1
        txt, ind = labels[col]
        exp_ind = d - 1 if layout == 'compact' else 0
        if txt != path[-1] or int(ind) != exp_ind:
            errs.append(f'{path}: подпись {labels[col]}, ждали {path[-1]!r} отступ {exp_ind}')
        # подписи предков без своей строки — на первой строке группы
        for a in range(1, d):
            if path[:a] in owned:
                continue
            first = prev is None or prev[:a] != path[:a]
            t = labels[a - 1][0]
            if first and t != path[a - 1]:
                errs.append(f'{path}: нет подписи предка {a}: {labels}')
            if not first and t:
                errs.append(f'{path}: лишняя подпись предка {a}: {labels}')
        prev = path
    return errs


def build_module():
    subprocess.check_call([sys.executable, os.path.join(HERE, 'build_harness.py'), SRC,
                           os.path.join(HERE, 'harness.bas'), os.path.join(HERE, '_harness_full.bas')])
    return open(os.path.join(HERE, '_harness_full.bas'), encoding='utf-8').read()


def rows_cases(n_random, rnd):
    cases = []
    layouts = ['compact', 'outline', 'tabular']
    for depth in range(1, 11):
        for layout, subt, repeat, grand, info, blanks in itertools.product(
                layouts, ['none', 'top', 'bottom'], [False, True], [False, True],
                ['field', 'geom', 'lt', 'partial'], [False, True]):
            if layout == 'tabular' and subt == 'top':
                continue        # Excel так не умеет
            if layout == 'compact' and repeat:
                continue
            cases.append((depth, layout, subt, repeat, grand, info, blanks))
    rnd.shuffle(cases)
    return cases[:n_random] if n_random else cases


def run_rows(b, n, seed=1):
    rnd = random.Random(seed)
    cases = rows_cases(n, rnd)
    batch, metas, fails = [], [], 0
    for case in cases:
        depth = case[0]
        root = gen_tree(rnd, depth, 40 if depth < 6 else 14)
        rows, ncol = render(root, *case)
        batch.append(encode(rows, ncol))
        metas.append((case, rows, ncol, root))
    outs = []
    for i in range(0, len(batch), 40):
        outs += b.call('ТестСтроки', GS.join(batch[i:i + 40])).split(GS)[:len(batch[i:i + 40])]
    for (case, rows, ncol, root), out in zip(metas, outs):
        errs = check(case, rows, ncol, out, root)
        if errs:
            fails += 1
            if fails <= 5:
                print('FAIL', case)
                for e in errs[:4]:
                    print('   ', e)
    print(f'строки: {len(cases) - fails}/{len(cases)} кейсов прошли')
    return fails


# ---------- шапка столбцов ----------
def col_case(lines, mx, measlev, meas):
    head = f'{len(lines)} {mx} {measlev} {len(meas)} ' + ' '.join(meas)
    return '\n'.join([head] + [RS.join([str(t), df] + items) for t, df, items in lines])


def run_cols(b):
    fails = 0
    tests = []
    # 1) регион > год, итоги регионов справа (как отдаёт Excel), общий итог в конце;
    #    одноимённые годы у разных регионов; итог отдаёт только сам элемент
    lines = [(0, '', ['Север', '2023']), (0, '', ['Север', '2024']), (1, '', ['Север']),
             (0, '', ['Юг', '2023']), (0, '', ['Юг', '2024']), (1, '', ['Юг']),
             (2, '', [])]
    tests.append((col_case(lines, 2, 0, ['Сумма']),
                  [7, 3, 1, 2, 6, 4, 5]))
    # 2) OLAP: итог второго уровня отдаёт [x, x] — ищем ближайшую группу, не первую
    lines = [(0, '', ['A', 'x', 'p']), (0, '', ['A', 'x', 'q']), (1, '', ['x', 'x']),
             (0, '', ['B', 'x', 'p']), (0, '', ['B', 'x', 'q']), (1, '', ['x', 'x']),
             (1, '', ['A']), (1, '', ['B'])]
    tests.append((col_case(lines, 3, 0, ['Сумма']),
                  [7, 3, 1, 2, 8, 6, 4, 5]))
    # 3) две меры под годом: мера выносится наверх, итоги — по мере
    lines = [(0, 'S1', ['2023', 'S1']), (0, 'S2', ['2023', 'S2']),
             (0, 'S1', ['2024', 'S1']), (0, 'S2', ['2024', 'S2']),
             (2, '', []), (2, '', [])]
    tests.append((col_case(lines, 2, 2, ['S1', 'S2']),
                  [5, 1, 3, 6, 2, 4]))
    out = b.call('ТестСтолбцы', GS.join(t for t, _ in tests)).split(GS)
    for n, ((_, exp), o) in enumerate(zip(tests, out), 1):
        got = [int(l.split(RS)[0]) for l in o.split('\n') if l]
        if got != exp:
            fails += 1
            print(f'FAIL столбцы {n}: ждали {exp}, вышло {got}')
            print('   ', [l.split(RS) for l in o.split('\n') if l])
    print(f'столбцы: {len(tests) - fails}/{len(tests)} кейсов прошли')
    return fails


if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    b = Basic()
    try:
        b.load(build_module())
        f = run_rows(b, n) + run_cols(b)
    finally:
        b.close()
    sys.exit(1 if f else 0)
