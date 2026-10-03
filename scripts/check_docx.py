#!/usr/bin/env python3
"""DOCX 交付物自检 —— 按 Word 交付规格逐条检查。

字体: 宋体小四正文 / 黑体三号四号小四标题 / 楷体五号说明 / Times New Roman 斜体公式
行距: 1.5 倍   颜色: 仅黑色   公式: 必须是 OMML 对象
用法: python3 check_docx.py <file.docx> [--json]
退出码: 0 = 通过(可含警告)  1 = 有失败项
"""
import sys, os, json, zipfile, re
from collections import Counter

try:
    import docx
    from docx.shared import Pt
except ImportError:
    sys.exit("需要 python-docx: pip install python-docx")

ALLOWED_BODY = {"宋体", "SimSun", "Songti SC", "Times New Roman"}
ALLOWED_HEAD = {"黑体", "SimHei", "Heiti SC"}
ALLOWED_NOTE = {"楷体", "KaiTi", "STKaiti"}
ALLOWED_CODE = {"Consolas", "Courier New", "等线", "Monaco"}
TARGET_LINE_SPACING = 1.5
MATH_HINT = re.compile(r"[∑∏√∫≈≠≤≥±βγδθλμσφψω]|[A-Za-z]_\{?\d|<[A-Za-z]|\^[2-9]|\\(?:sqrt|frac|sum|alpha|beta|theta)\b")
BAN_WORDS = ["综上所述", "总而言之", "具有重要的意义", "发挥着重要作用",
             "在当今时代背景下", "值得关注的是", "不难发现", "赋能", "抓手", "闭环"]


def run_font(run):
    try:
        return run.font.name
    except Exception:
        return None


def style_font(par):
    """run 未显式指定字体时，回退到段落样式/文档默认字体。"""
    try:
        st = par.style
        while st is not None:
            f = st.font
            if f is not None and f.name:
                return f.name
            st = getattr(st, "base_style", None)
    except Exception:
        pass
    return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    path = args[0]
    if not os.path.exists(path):
        sys.exit(f"文件不存在: {path}")

    d = docx.Document(path)
    fails, warns, infos = [], [], []

    # ---------- 字体 / 字号 / 颜色 ----------
    fonts = Counter()
    sizes = Counter()
    colors = []
    for p in d.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            f = run_font(r) or style_font(p)
            if f:
                fonts[f] += 1
            if r.font.size:
                sizes[round(r.font.size.pt, 1)] += 1
            try:
                if r.font.color and r.font.color.rgb is not None:
                    c = str(r.font.color.rgb).upper()
                    if c not in ("000000", "FF000000"):
                        colors.append((c, r.text[:20]))
            except Exception:
                pass
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        f = run_font(r)
                        if f:
                            fonts[f] += 1
                        try:
                            if r.font.color and r.font.color.rgb is not None:
                                c = str(r.font.color.rgb).upper()
                                if c not in ("000000", "FF000000"):
                                    colors.append((c, r.text[:20]))
                        except Exception:
                            pass

    infos.append(f"检出字体: {dict(fonts.most_common(8))}")
    allowed = ALLOWED_BODY | ALLOWED_HEAD | ALLOWED_NOTE | ALLOWED_CODE
    bad_fonts = [f for f in fonts if f and f not in allowed]
    if bad_fonts:
        fails.append(f"[字体] 非白名单字体: {bad_fonts}（仅允许宋体/黑体/楷体/Times New Roman/等宽代码字体）")

    if colors:
        fails.append(f"[颜色] 检出 {len(colors)} 处非黑色文字，违反'全文只有黑色': {colors[:5]}")

    # ---------- 行距 ----------
    ls_counter = Counter()
    for p in d.paragraphs:
        if not p.text.strip():
            continue
        ls = p.paragraph_format.line_spacing
        ls_counter[str(ls)] += 1
    infos.append(f"行距分布: {dict(ls_counter.most_common(5))}")
    explicit = {k: v for k, v in ls_counter.items() if k != "None"}
    if explicit:
        good = sum(v for k, v in explicit.items() if k.startswith("1.5"))
        if good < sum(explicit.values()) * 0.8:
            fails.append(f"[行距] 显式行距未统一为 1.5 倍: {explicit}")
    else:
        warns.append("[行距] 未显式设置行距（继承样式），请确认样式本身为 1.5 倍")

    # ---------- 页眉页脚页码 ----------
    issues = []
    for i, s in enumerate(d.sections, 1):
        try:
            hp = s.header.paragraphs + [p for t in s.header.tables for r in t.rows
                                        for c in r.cells for p in c.paragraphs]
            fp = s.footer.paragraphs + [p for t in s.footer.tables for r in t.rows
                                        for c in r.cells for p in c.paragraphs]
            if any(x.text.strip() for x in hp):
                issues.append(f"第{i}节有页眉")
            if any(x.text.strip() for x in fp):
                issues.append(f"第{i}节有页脚")
        except Exception:
            pass
    if issues:
        fails.append(f"[页眉页脚] {issues} —— 当前规范不使用页眉、页脚、页码")

    # ---------- 公式是否为 OMML ----------
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    omml_count = xml.count("<m:oMath") + xml.count("<m:oMathPara")
    fake = []
    for p in d.paragraphs:
        t = p.text
        if t and MATH_HINT.search(t):
            # 该段若含 OMML，python-docx 的 p.text 通常仍会取到文本；用 XML 粗判
            if omml_count == 0:
                fake.append(t[:40])
    infos.append(f"OMML 公式对象数: {omml_count}")
    if fake:
        fails.append(
            f"[公式] 检出疑似公式但全文无 OMML 对象（必须以 Equation 对象呈现，不能纯文本）: {fake[:5]}")

    # ---------- 表格 ----------
    infos.append(f"表格数: {len(d.tables)}")

    # ---------- AI 味 ----------
    blob = "\n".join(p.text for p in d.paragraphs)
    hits = [b for b in BAN_WORDS if b in blob]
    if hits:
        warns.append(f"[AI味] 命中高频套话: {hits}")

    # ---------- 字数 ----------
    zh = len(re.findall(r"[\u4e00-\u9fff]", blob))
    infos.append(f"中文字符数(段落): {zh}")

    report = {"file": os.path.abspath(path), "info": infos,
              "failed": fails, "warned": warns, "pass": not fails}
    if "--json" in sys.argv:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"\n=== DOCX 自检: {os.path.basename(path)} ===")
        for i in infos:
            print(f"  · {i}")
        if fails:
            print("\n  [FAIL]")
            for f in fails:
                print(f"   ✗ {f}")
        if warns:
            print("\n  [WARN]")
            for w in warns:
                print(f"   ! {w}")
        if not fails and not warns:
            print("\n  ✓ 全部通过")
        print()
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
