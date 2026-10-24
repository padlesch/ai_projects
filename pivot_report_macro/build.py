"""Собирает модуль для импорта в Excel: кодировка Windows-1251, переводы строк CRLF.

python3 build.py  ->  dist/modОтчетИзСводнойV3.bas
"""
import os
here = os.path.dirname(os.path.abspath(__file__))
src = os.path.join(here, 'src', 'modОтчетИзСводнойV3.bas')
dst = os.path.join(here, 'dist', 'modОтчетИзСводнойV3.bas')
os.makedirs(os.path.dirname(dst), exist_ok=True)
text = open(src, encoding='utf-8').read().replace('\r\n', '\n').replace('\n', '\r\n')
open(dst, 'wb').write(text.encode('cp1251'))
print(dst)
