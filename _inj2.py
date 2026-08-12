import re, ast
from pathlib import Path
BASE = Path(r"C:\Users\wuzhe\WorkBuddy\2026-08-02-14-20-14\pptx-kb\src\pptx_kb")
NAMES = {"cli_img": "pptx-img", "cli_formula": "pptx-formula",
         "cli_text": "pptx-text", "cli_caption": "pptx-caption",
         "cli_author": "pptx-author", "cli_bind": "pptx-bind",
         "cli_paser": "pptx-paser", "gen_html": "pptx-html",
         "build_deck": "pptx-deck"}
for stem, name in NAMES.items():
    p = BASE / f"{stem}.py"
    t = p.read_text(encoding="utf-8")
    # 1) 移除 main() 包装里的 banner（避免重复/错位）
    t = re.sub(r'    banner\("[^"]+"\)\n', '', t)
    t = re.sub(r'    banner_end\("[^"]+"\)\n', '', t)
    # 2) 主函数：cli_img/formula 用 main(argv=None)，其余用 _main(argv=None)
    main_sig = ("def main(argv=None) -> int:\n" if stem in ("cli_img", "cli_formula")
                else "def _main(argv=None) -> int:\n")
    if main_sig not in t:
        print(f"[跳过] {stem} 未找到 {main_sig.strip()}")
        continue
    if f'banner("{name}")' in t:
        print(f"已存在: {stem}")
        continue
    t = t.replace(main_sig, main_sig + f'    banner("{name}")\n', 1)
    # 3) return EXIT_OK 前插 banner_end
    idx = t.rfind("    return EXIT_OK")
    if idx != -1:
        t = t[:idx] + f'    banner_end("{name}")\n' + t[idx:]
    p.write_text(t, encoding="utf-8")
    ast.parse(t)
    print(f"完成: {stem}")
print("全部完成")
