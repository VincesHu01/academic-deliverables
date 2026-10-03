#!/usr/bin/env python3
"""按统一命名与目录规范初始化一次交付工作区。

用法:
  python3 scaffold.py --dir <输出根目录> --items "1,2,3" [--code-dir <代码目录>]
  python3 scaffold.py --dir ./outputs --prefix 题 --items "一,二,三"

只创建目录与命名清单，不生成内容。
"""
import argparse, os, json, sys


def main():
    ap = argparse.ArgumentParser(description="初始化交付物工作区（目录 + 规范命名清单）")
    ap.add_argument("--dir", required=True, help="输出根目录")
    ap.add_argument("--items", default="", help="题号/章节号，逗号分隔，如 '1,2,3'")
    ap.add_argument("--prefix", default="", help="文件名前缀，默认无")
    ap.add_argument("--code-dir", default=None, help="代码汇总目录（不传则在 --dir 下建 代码）")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    root = os.path.abspath(os.path.expanduser(a.dir))
    os.makedirs(root, exist_ok=True)
    code_dir = os.path.abspath(os.path.expanduser(a.code_dir)) if a.code_dir else os.path.join(root, "代码")
    os.makedirs(code_dir, exist_ok=True)

    items = [s.strip() for s in a.items.split(",") if s.strip()]
    plan = []
    for it in items:
        base = f"{a.prefix}{it}+原始输出"
        plan.append({"item": it, "primary": os.path.join(root, f"{base}.docx"),
                     "extra_pattern": os.path.join(root, f"{base}+{{序号}}.docx")})

    manifest = {
        "output_dir": root,
        "code_dir": code_dir,
        "naming_rule": "{题号}+原始输出[+序号]",
        "plan": plan,
        "checklist": [
            "同一个题的原始输出与图表禁止跨页",
            "图注表注统一：图在上注在下，表在上注在上",
            "截图不截取软件标识栏与时间戳",
            "代码必须有标准化注释，且可直接粘贴运行",
            "所有源码汇总为单一文件放入代码目录",
        ],
    }
    with open(os.path.join(root, "交付清单.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    if a.json:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    else:
        print(f"\n工作区已就绪: {root}")
        print(f"代码目录:     {code_dir}")
        for p in plan:
            print(f"  · 第 {p['item']} 题 -> {os.path.basename(p['primary'])}")
        print("\n命名规则: 题号+原始输出；多个产物则 题号+原始输出+序号")
        print("清单已写入:", os.path.join(root, "交付清单.json"))


if __name__ == "__main__":
    main()
