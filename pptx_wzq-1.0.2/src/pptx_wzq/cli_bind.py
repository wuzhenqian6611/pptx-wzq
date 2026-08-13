"""cli_bind.py — pptx-bind 薄壳：把教材文案与对应页图片关系绑定为 JSON。

输入：textbook.md（每页文案，## 第 N 页）、images 目录（文件名 slide_NN_
     编码页码）、captions.md（图片 AI 解读，### IMGxxxx — slide_NN_）。
输出：<名>_binding.json，按页组织：

    {
      "stem": "xxx",
      "pages": [
        {"page": 4,
         "text": "第 4 页文案全文…",
         "images": [{"file": "slide_04_pic_05.png",
                     "caption": "1) 图片类型…" }],
         "has_image": true}
      ]
    }

用法：
    pptx-bind <产物目录> [-o binding.json]
              [--textbook a.md] [--images-dir dir] [--captions c.md]
              [--json] [--version]

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        print_json, quiet_stdout, banner, banner_end)

VERSION = "pptx-bind 1.0.0 (方案B薄壳)"


def _split_pages(content: str) -> dict:
    pages = {}
    cur = None
    for line in content.splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", line.strip())
