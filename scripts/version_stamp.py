# -*- coding: utf-8 -*-
"""面板代码版本戳（唯一的算法定义处）。

为什么要单独一个文件：版本号是"重启时该不该替换旧进程"的唯一判据。web_panel 与
launch_panel 是两个进程，各自算一遍；一旦两边口径不同，就永远"指纹不一致"或反过来
——实测踩过：改了 insight_scan.py 却因为版本戳只含 web_panel.py+模板而没被替换，
新代码根本没生效（就是交接文档里"改了没生效"那类坑的模块版）。

覆盖范围 = scripts/*.py + scripts/templates/*.html（面板进程会 import 的一切本地代码），
但**跳过 `_*` 脚本**：它们是一次性验收/打包工具，面板不 import，打包时也不进包——
不跳过的话"磁盘指纹"与"包内指纹"永远对不上，就没法用它判断成品包是否过期。

`root` 可指向成品包的 scripts 目录，用同一套算法算出可比的指纹。
"""
import hashlib
import os

SCRIPTS = os.path.dirname(os.path.abspath(__file__))


def panel_version(sha=None, length=12, root=None):
    """面板与启动器共用；任何一处口径变了，另一处自动跟上。"""
    h = hashlib.md5() if sha is None else hashlib.new(sha)
    base = root or SCRIPTS
    files = []
    for pat, sub in (("*.py", ""), ("*.html", "templates")):
        d = os.path.join(base, sub) if sub else base
        try:
            files += [os.path.join(d, f) for f in os.listdir(d)
                      if f.endswith(pat.split("*")[1]) and not f.startswith("_")]
        except OSError:
            pass
    for p in sorted(files):
        try:
            h.update(os.path.basename(p).encode("utf-8"))
            with open(p, "rb") as f:
                h.update(f.read())
        except OSError:
            continue
    return h.hexdigest()[:length]


if __name__ == "__main__":
    print(panel_version())
