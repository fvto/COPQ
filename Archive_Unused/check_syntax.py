import sys
for f in ("Interactive_Dashboard.py", "process_ftt.py"):
    src = open(f, encoding="utf-8").read()
    try:
        compile(src, f, "exec")
        print(f, "OK")
    except SyntaxError as e:
        print(f, "ERROR line", e.lineno, "offset", e.offset)
        lines = src.split("\n")
        bad = lines[e.lineno - 1]
        print("  text:", repr(bad))
        if e.offset:
            ch = bad[e.offset - 1]
            print("  char:", repr(ch), hex(ord(ch)))
