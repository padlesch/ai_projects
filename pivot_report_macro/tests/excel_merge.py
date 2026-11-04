"""Как Excel объединит шапку по плану макроса.

Excel при объединении диапазона, задевающего уже объединённые ячейки,
расширяет его до общего прямоугольника и оставляет только значение
верхней левой клетки — подписи внутри стираются. Здесь это смоделировано,
чтобы проверять план, не полагаясь на эмуляцию LibreOffice (она так не делает).
"""
RS = chr(30)


def parse(plan):
    lines = [l for l in plan.split('\n') if l]
    nlev, nc = map(int, lines[0].split(RS))
    grid = [lines[1 + k].split(RS) for k in range(nlev)]
    rects = [tuple(map(int, l.split(RS))) for l in lines[1 + nlev:]]
    return nlev, nc, grid, rects


def apply_excel(nlev, nc, grid, rects):
    g = [row[:] for row in grid]
    merged = []                      # список прямоугольников (k1,c1,k2,c2)
    for r in rects:
        k1, c1, k2, c2 = r
        changed = True
        while changed:               # расширяем, пока задевает существующие
            changed = False
            for m in merged[:]:
                if not (m[2] < k1 or m[0] > k2 or m[3] < c1 or m[1] > c2):
                    k1, c1, k2, c2 = min(k1, m[0]), min(c1, m[1]), max(k2, m[2]), max(c2, m[3])
                    merged.remove(m)
                    changed = True
        merged.append((k1, c1, k2, c2))
    for k1, c1, k2, c2 in merged:
        top = g[k1 - 1][c1 - 1]
        for k in range(k1, k2 + 1):
            for c in range(c1, c2 + 1):
                g[k - 1][c - 1] = ''
        g[k1 - 1][c1 - 1] = top
    return g, merged


def lost_labels(plan):
    """Подписи, которые Excel сотрёт при объединении по этому плану:
    в итоговом блоке, куда попала ячейка, виден другой текст."""
    nlev, nc, grid, rects = parse(plan)
    g, merged = apply_excel(nlev, nc, grid, rects)
    lost = []
    for k in range(1, nlev + 1):
        for c in range(1, nc + 1):
            t = grid[k - 1][c - 1]
            if not t:
                continue
            box = next(((m[0], m[1]) for m in merged if m[0] <= k <= m[2] and m[1] <= c <= m[3]), (k, c))
            if g[box[0] - 1][box[1] - 1] != t:
                lost.append((k, c, t))
    return lost, g
