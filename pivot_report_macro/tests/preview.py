"""Образцы палитр: один и тот же отчёт (общий итог + 4 уровня) во всех
палитрах, экспорт в PDF и PNG — чтобы посмотреть глазами.
python3 preview.py папка_для_картинок
"""
import os, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import run_out, manual
from runbasic import Basic
from com.sun.star.beans import PropertyValue

REG, SUB, GRAND = 0, 1, 2
N = None


def r(lt, d, name, vals, ct=None):
    c = [N] * 4
    c[d - 1] = (name, ct or (2 if lt == SUB else 1), d, 0)
    return (lt, c, vals)


CASE = dict(caps=['Регион', 'Город', 'Отделение', 'Услуга'], cols=['Янв', 'Фев', 'Мар', 'Общий итог'],
            rows=[
    (GRAND, [('Общий итог', 3, 0, 0), N, N, N], [1250, 1310, 1402, 3962]),
    r(SUB, 1, 'Северный округ', [700, 720, 790, 2210]),
    r(SUB, 2, 'Москва', [520, 530, 600, 1650]),
    r(SUB, 3, 'Поликлиника № 1', [300, 310, 350, 960]),
    r(REG, 4, 'Консультации', [180, 185, 210, 575]),
    r(REG, 4, 'Диагностика', [120, 125, -140, 105]),
    r(SUB, 3, 'Стационар', [220, 220, 250, 690]),
    r(REG, 4, 'Операции', [220, 220, 250, 690]),
    r(SUB, 2, 'Тверь', [180, 190, 190, 560]),
    r(SUB, 3, 'Поликлиника № 7', [180, 190, 190, 560]),
    r(REG, 4, 'Консультации', [180, 190, 190, 560]),
    r(SUB, 1, 'Южный округ', [550, 590, 612, 1752]),
    r(SUB, 2, 'Сочи', [550, 590, 612, 1752]),
    r(SUB, 3, 'Центр компетенций', [550, 590, 612, 1752]),
    r(REG, 4, 'Диагностика', [300, 330, 352, 982]),
    r(REG, 4, 'Вакцинация', [250, 260, 260, 770]),
])

if __name__ == '__main__':
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    src = run_out.build_src() + open(os.path.join(HERE, 'harness_manual.bas'), encoding='utf-8').read()
    b = Basic(2390)
    try:
        b.load(src)
        for pal in range(1, 8):
            res = b.call('Построить', manual.encode(CASE), pal)
            if res.startswith('ОШИБКА'):
                print(pal, res)
        # убрать пустой исходный лист и экспортировать
        sheets = b.doc.Sheets
        for name in list(sheets.ElementNames):
            sh = sheets.getByName(name)
            if sh.getCellByPosition(1, 0).String == '':
                sheets.removeByName(name)
        pdf = os.path.join(out_dir, 'палитры.pdf')
        pv = PropertyValue(); pv.Name = 'FilterName'; pv.Value = 'calc_pdf_Export'
        b.doc.storeToURL('file://' + pdf, (pv,))
    finally:
        b.close()
    subprocess.check_call(['pdftoppm', '-png', '-r', '80', pdf, os.path.join(out_dir, 'п')])
    print(sorted(os.listdir(out_dir)))
