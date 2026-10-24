"""Прогон вывода отчёта и перекраски в эмуляции Excel (LibreOffice VBA).

python3 run_out.py          — результат прогона
С ключом --trace перед каждой строкой модуля ставится метка шага,
и при ошибке видно, на какой строке исходника она случилась.
"""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from runbasic import Basic

STUB = '''Private Sub Пометить(rng As Range, ByVal роль As String)
    Dim i As Long, j As Long, k As String
    On Error Resume Next
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            k = (rng.Row + i - 1) & "," & (rng.Column + j - 1)
            ТестМетки.Remove k
            ТестМетки.Add роль, k
        Next j
    Next i
End Sub

Private Function РольЯчейки(cel As Range) As String
    On Error Resume Next
    РольЯчейки = ТестМетки.Item(cel.Row & "," & cel.Column)
End Function
'''


def build_src(trace=False):
    """Модуль для эмуляции: VBASupport, метки ролей — в коллекции
    (стили ячеек эмуляция LibreOffice не создаёт), остальной код как есть."""
    src = open(os.path.join(HERE, '..', 'src', 'modОтчетИзСводнойV3.bas'), encoding='utf-8').read()
    lines = src.split('\n')
    lines[0] = 'Option VBASupport 1'
    if trace:
        out, in_proc, prev_cont = [], False, False
        for n, l in enumerate(lines, 1):
            st = l.strip()
            if re.match(r'(Public |Private )?(Sub|Function) ', st):
                in_proc = True
            ok = (in_proc and not prev_cont and st and not st.startswith("'")
                  and not re.match(r'(Case\b|End (Sub|Function)|(Public |Private )?(Sub|Function) )', st)
                  and not re.match(r'^[\wА-Яа-яЁё]+:$', st))
            if ok:
                out.append(f'ШАГ = {n}')
            out.append(l)
            prev_cont = st.endswith(' _')
        lines = out
        lines.insert(2, 'Public ШАГ As Long')
    src = '\n'.join(lines)
    for name in ('Private Function РольЯчейки(', 'Private Sub Пометить('):
        a = src.index(name)
        end = 'End Sub' if 'Sub' in name else 'End Function'
        b = src.index(end, a) + len(end)
        src = src[:a] + (STUB if 'Пометить' in name else '') + src[b:]
    return src.replace('Option VBASupport 1', 'Option VBASupport 1\nPrivate ТестМетки As New Collection', 1)


if __name__ == '__main__':
    trace = '--trace' in sys.argv
    src = build_src(trace) + open(os.path.join(HERE, 'harness_out.bas'), encoding='utf-8').read()
    if trace:
        src = src.replace('(строка " & Erl & ")"', '(строка исходника " & ШАГ & ")"')
    open(os.path.join(HERE, '_out_full.bas'), 'w', encoding='utf-8').write(src)
    b = Basic(2311)
    try:
        b.load(src)
        out = b.call('ТестВывод')
        print(out)
        # BGR-hex, как отдаёт Interior.Color
        exp = [
            'ПОСТРОЕН меток=19',
            '[Регион|6F4E2F|H|0|b][Общий итог|6F4E2F|H|0|b][2023|6F4E2F|H|0|b]',   # итог по столбцам — слева
            '[1|FFFFFFFF|RG|0|][Общий итог|DAD1C9||0|b][84.00|CBBFB4||0|b]',       # общий итог — первой строкой
            '[2|FFFFFFFF|RS1|0|][Север|E9E4E0||0|b][42.00|DBD3CB||0|b]',          # итог Севера поднят, заголовок поглощён
            '[3|FFFFFFFF|RD|0|][Москва|FFFFFFFF||1|][21.00|EBE6E2||0|b][20|FFFFFFFF||0|][100.0%',
            '[5|FFFFFFFF|RS1|0|][Юг|E9E4E0||0|b][73.50',
            'ТЕКУЩАЯ=1', 'ПЕРЕКРАШЕН в 4',
            '[№|F0ECE8|H|0|b]',                                                    # B: шапка #E8ECF0
            '[Общий итог|ECDCC9||0|b][84.00|ECDCC9||0|b]',                         # B: общий итог #C9DCEC
            '[Север|F3E8DC||0|b]',                                                 # B: уровень 1 #DCE8F3
            '[21.00|F2EFEC||0|b]',                                                 # B: ключевая колонка #ECEFF2
            'ТЕКУЩАЯ=4',
        ]
        bad = [e for e in exp if e not in out]
        for e in bad:
            print('НЕ НАЙДЕНО:', e)
        print('вывод и перекраска:', 'OK' if not bad else f'{len(bad)} расхождений')
        sys.exit(1 if bad else 0)
    finally:
        b.close()
