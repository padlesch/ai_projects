"""10 уровней строк через весь путь: разбор + вывод на лист + раскраска.

Для каждой сгенерированной сводной (сжатая / структура / табличная, итоги
сверху / снизу) проверяется каждая строка отчёта: порядок как в сводной,
заливка — цвет своего уровня в палитре A, жирные — только уровни.
python3 deep.py
"""
import itertools, os, random, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import run_out, manual
from test_core import Node, gen_tree, expected
from runbasic import Basic

RS, US = chr(30), chr(31)
# палитра A: уровни 1..3 из дизайна, глубже — светлее третьего (Оттенок 0.5)
A = {1: 'FBE2D5', 2: 'FFEBCD', 3: 'E9F3E3', 4: 'F4F9F1'}


def build_case(root, depth, layout, subt):
    """Сводная «как на листе» в формате manual.py (поле = уровень)."""
    rows = []
    ncol = 1 if layout == 'compact' else depth

    def cells(d, txt, ct):
        c = [None] * ncol
        if layout == 'compact':
            c[0] = (txt, ct, d, d - 1)
        else:
            c[d - 1] = (txt, ct, d, 0)
        return c

    def walk(node, path, first_in_parent_chain):
        d = len(path)
        leaf = not node.children
        if leaf or layout != 'tabular':
            top = (not leaf) and subt == 'top'
            c = cells(d, node.name, 2 if top else 1)
            if layout == 'tabular':
                for a in range(1, d):
                    if first_in_parent_chain[a - 1]:
                        c[a - 1] = (path[a - 1], 1, a, 0)
                        first_in_parent_chain[a - 1] = False
            v = [] if (not leaf and subt != 'top') else [d * 10 + 1, d * 10 + 2]
            rows.append((1 if top else 0, c, v, tuple(path)))
        for ch in node.children:
            fl = first_in_parent_chain[:d] + [True] * (depth - d)
            walk(ch, path + [ch.name], fl)
            first_in_parent_chain = [False] * depth
        if node.children and subt == 'bottom':
            rows.append((1, cells(d, node.name + ' Итог', 2), [d * 10 + 5, d * 10 + 6], tuple(path)))

    for ch in root.children:
        walk(ch, [ch.name], [True] * depth)
    case = dict(caps=[f'Поле {k + 1}' for k in range(ncol)], cols=['Янв', 'Фев'],
                rows=[(lt, c, v) for lt, c, v, _ in rows])
    return case, rows


def main():
    rnd = random.Random(7)
    src = run_out.build_src() + open(os.path.join(HERE, 'harness_manual.bas'), encoding='utf-8').read()
    b = Basic(2380)
    fails = total = 0
    sample = None
    try:
        b.load(src)
        for layout, subt, depth in itertools.product(['compact', 'outline', 'tabular'], ['top', 'bottom', 'none'],
                                                     [3, 5, 7, 10]):
            if layout == 'tabular' and subt == 'top':
                continue
            root = gen_tree(rnd, depth, 12)
            case, rows = build_case(root, depth, layout, subt)
            out = b.call('Построить', manual.encode(case), 3)
            total += 1
            if not out or out.startswith('ОШИБКА'):
                fails += 1
                print('FAIL', layout, subt, depth, out)
                continue
            lines = [l for l in out.split('\n') if l][2:]          # без шапки и нумерации
            exp = [p for k, p in expected(root, layout, subt if subt != 'none' else 'none', False)]
            got = []
            errs = []
            for ln in lines:
                cells = [x.split(US) for x in ln.split(RS) if x]
                lab = [c[0].lstrip('*').lstrip('·') for c in cells[1:1 + (1 if layout == 'compact' else depth)]]
                name = next((t for t in reversed(lab) if t), '')
                fill = cells[1][1][-6:]
                fill = fill[4:6] + fill[2:4] + fill[0:2]        # BGR -> RGB
                bold = cells[1][0].startswith('*') or any(c[0].startswith('*') for c in cells[1:1 + len(lab)])
                got.append((name, fill, bold))
            if len(got) != len(exp):
                errs.append(f'строк {len(got)}, ждали {len(exp)}')
            for (name, fill, bold), path in zip(got, exp):
                d = len(path)
                want = 'FFFFFF' if d == depth else A[min(d, 4)]
                if name != path[-1]:
                    errs.append(f'порядок: «{name}» вместо «{path[-1]}»')
                    break
                if fill != want:
                    errs.append(f'«{name}» уровень {d}: заливка #{fill}, ждали #{want}')
                if bold != (d < depth) and name:
                    errs.append(f'«{name}» уровень {d}: жирный={bold}')
            if errs:
                fails += 1
                print('FAIL', layout, subt, depth, errs[:3])
            if layout == 'outline' and subt == 'top' and depth == 10:
                sample = got[:22]
        print(f'10 уровней, весь путь до листа: {total - fails}/{total} сводных без ошибок')
        if sample:
            print('пример: структура, итоги вверху, 10 уровней — первые строки отчёта (палитра A):')
            for name, fill, bold in sample:
                print(f'    {"*" if bold else " "} {name:14} #{fill}')
    finally:
        b.close()
    return fails


if __name__ == '__main__':
    sys.exit(1 if main() else 0)
