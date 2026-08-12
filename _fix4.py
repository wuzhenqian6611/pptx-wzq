import ast
from pathlib import Path
BASE = Path(r"C:/Users/wuzhe/WorkBuddy/2026-08-02-14-20-14/pptx-kb/src/pptx_kb")
FILES = ["cli_img.py","cli_formula.py","cli_text.py","cli_caption.py","cli_author.py",
         "cli_bind.py","cli_paser.py","gen_html.py","build_deck.py"]
for f in FILES:
    p = BASE / f
    raw = p.read_bytes().decode("utf-8").replace("\r\n", "\n")
    if "banner" in raw.split("from pptx_kb.cli_common import", 1)[1][:200]:
        print(f"已含 banner: {f}")
        continue
    # 在 cli_common import 块内追加 banner, banner_end（在最后一个导入名前补）
    import re
    m = re.search(r"from pptx_kb.cli_common import \(([^)]*)\)", raw, re.S)
    if m:
        body = m.group(1)
        new_body = body.rstrip() + "\n    banner, banner_end,"
        raw = raw.replace(m.group(0), "from pptx_kb.cli_common import (" + new_body + ")", 1)
    elif "from pptx_kb.cli_common import " in raw and "banner" not in raw.split("cli_common import ")[1][:100]:
        raw = raw.replace("from pptx_kb.cli_common import ",
                          "from pptx_kb.cli_common import banner, banner_end, ", 1)
    p.write_text(raw, encoding="utf-8", newline="\n")
    ast.parse(raw)
    print(f"补 import: {f}")
print("完成")
