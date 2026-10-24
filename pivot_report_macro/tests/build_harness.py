import re, sys
src = open(sys.argv[1], encoding='utf-8').read()
decl = re.search(r"'#ДЕКЛ-НАЧАЛО(.*?)'#ДЕКЛ-КОНЕЦ", src, re.S).group(1)
core = re.search(r"'#ЯДРО-НАЧАЛО(.*?)'#ЯДРО-КОНЕЦ", src, re.S).group(1)
h = open(sys.argv[2], encoding='utf-8').read()
open(sys.argv[3], 'w', encoding='utf-8').write("Option VBASupport 1\nOption Explicit\n" + decl + core + h)
