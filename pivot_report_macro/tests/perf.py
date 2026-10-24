"""Скорость ядра на большой сводной (табличный макет, итоги внизу)."""
import random, sys, time, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from runbasic import Basic
from test_core import Node, render, encode, build_module, GS

def big(depth, fan):
    def mk(d):
        if d == depth: return []
        return [Node(f'n{d}_{i}', mk(d + 1)) for i in range(fan[d])]
    return Node('ROOT', mk(0))

b = Basic(2302)
try:
    b.load(build_module())
    for depth, fan, layout in [(4, [10, 10, 10, 5], 'tabular'), (4, [10, 10, 10, 20], 'tabular')]:
        root = big(depth, fan)
        rows, ncol = render(root, depth, layout, 'bottom', False, True, 'field', False)
        t = time.time()
        out = b.call('ТестСтроки', encode(rows, ncol))
        n = len([l for l in out.split(GS)[0].split('\n') if l])
        print(f'{layout:8} глубина {depth}: {len(rows)} строк сводной -> {n} строк отчёта за {time.time() - t:.1f} с')
finally:
    b.close()
