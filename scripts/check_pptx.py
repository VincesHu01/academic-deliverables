#!/usr/bin/env python3
"""PPTX 交付物自检 —— 按 PPT 交付规格逐条检查。

用法: python3 check_pptx.py <file.pptx> [--json]
退出码: 0 = 通过(可含警告)  1 = 有失败项  2 = 用法错误
"""
import sys, os, json, colorsys
from collections import Counter

try:
    from pptx import Presentation
    from pptx.util import Pt
except ImportError:
    sys.exit("需要 python-pptx: pip install python-pptx")

MIN_BODY_PT = 14.0      # 正文下限
MIN_ANY_PT = 10.0       # 任何文字的绝对下限
MAX_HUES = 2            # 允许的主色相数量(含黑白灰之外的)
MAX_SLIDES = 15         # 默认页数上限
MAX_CHARS_PER_SLIDE = 420
ALLOWED_FONTS = {"微软雅黑", "Microsoft YaHei", "楷体", "KaiTi",
                 "STKaiti", "楷体_GB2312", "等线"}


def run_fonts(run):
    """取一个 run 声明的字体族(latin + eastAsia + cs)，去重返回集合。"""
    names = set()
    try:
        if run.font.name:
            names.add(run.font.name)
    except Exception:
        pass
    try:
        rPr = run._r.get_or_add_rPr()
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            v = rPr.get(attr)
            if v:
                names.add(v)
    except Exception:
        pass
    return names



def norm_hex(v):
    if v is None:
        return None
    try:
        if hasattr(v, "rgb") and v.rgb is not None:
            return str(v.rgb)
        s = str(v)
        return s if s and s != "None" else None
    except Exception:
        return None


def classify(hexs):
    """把十六进制颜色分成 'neutral' 或 色相桶(每30度一桶)。"""
    if not hexs:
        return None
    s = hexs.lstrip("#")
    if len(s) != 6:
        return None
    try:
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    except ValueError:
        return None
    h, l, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if sat < 0.12 or l < 0.06 or l > 0.96:
        return "neutral"
    return f"hue{int(h * 360) // 30 * 30}"


def walk_shapes(shapes, out):
    for sh in shapes:
        if sh.shape_type == 6:  # GROUP
            try:
                walk_shapes(sh.shapes, out)
            except Exception:
                pass
            continue
        out.append(sh)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    path = args[0]
    if not os.path.exists(path):
        sys.exit(f"文件不存在: {path}")

    prs = Presentation(path)
    fails, warns, infos = [], [], []

    shapes_by_slide = []
    for i, slide in enumerate(prs.slides, 1):
        acc = []
        walk_shapes(slide.shapes, acc)
        shapes_by_slide.append((i, acc))

    # ---- 页数
    n = len(prs.slides)
    infos.append(f"页数: {n}")
    if n > MAX_SLIDES:
        warns.append(f"[页数] {n} 页超过默认上限 {MAX_SLIDES} 页，请确认是否需要压缩")

    # ---- 配色
    hues = Counter()
    color_detail = Counter()
    fonts = Counter()
    tiny = []
    slide_text_len = []
    editable_text_boxes = 0

    for idx, shapes in shapes_by_slide:
        total = 0
        for sh in shapes:
            # 填充色
            try:
                if sh.has_text_frame is False and getattr(sh, "fill", None) is not None:
                    f = sh.fill
                    if f.type is not None and f.type == 1:
                        c = norm_hex(f.fore_color)
                        k = classify(c)
                        if k and k != "neutral":
                            hues[k] += 1
                            color_detail[c] += 1
            except Exception:
                pass
            # 文字
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        txt = r.text or ""
                        total += len(txt)
                        if txt.strip():
                            editable_text_boxes += 1
                        for fn in run_fonts(r):
                            if fn:
                                fonts[fn] += 1
                        sz = r.font.size.pt if r.font.size else None
                        if sz is not None:
                            if sz < MIN_ANY_PT:
                                tiny.append((idx, round(sz, 1), txt[:24]))
                            elif sz < MIN_BODY_PT and len(txt) > 40:
                                warns.append(
                                    f"[字号] 第 {idx} 页正文 {sz}pt 低于 {MIN_BODY_PT}pt: 「{txt[:24]}…」")
                        c = None
                        try:
                            if r.font.color and r.font.color.rgb is not None:
                                c = str(r.font.color.rgb)
                        except Exception:
                            pass
                        k = classify(c)
                        if k and k != "neutral":
                            hues[k] += 1
                            color_detail[c] += 1
        slide_text_len.append((idx, total))

    non_neutral = len([h for h in hues if h != "neutral"])
    infos.append(f"非中性色相数: {non_neutral} (上限 {MAX_HUES})")
    if non_neutral > MAX_HUES:
        fails.append(
            f"[配色] 检出 {non_neutral} 个非中性色相 -> 违反'色相<=2'。明细: {dict(color_detail.most_common(8))}")
    elif non_neutral == MAX_HUES:
        warns.append(f"[配色] 已用到 {non_neutral} 个色相，接近上限，确认是否必要")

    if tiny:
        fails.append(f"[字号] 有 {len(tiny)} 处文字 < {MIN_ANY_PT}pt: {tiny[:5]}")

    # ---- 字体族（仅允许 微软雅黑 正文 / 楷体 注释）
    infos.append(f"检出字体族: {dict(fonts.most_common(6))}")
    bad_fonts = [f for f in fonts if f and f not in ALLOWED_FONTS]
    if bad_fonts:
        fails.append(
            f"[字体] 检出非白名单字体族 {bad_fonts}（仅允许 微软雅黑 正文/标题、楷体 注释）")

    # ---- 文字密度
    dense = [(i, c) for i, c in slide_text_len if c > MAX_CHARS_PER_SLIDE]
    if dense:
        warns.append(
            f"[密度] {len(dense)} 页文字超过 {MAX_CHARS_PER_SLIDE} 字，可能'排版拥挤'/'纯文字堆砌': {dense[:5]}")
    empty = [i for i, c in slide_text_len if c == 0]
    if empty:
        warns.append(f"[结构] 第 {empty} 页没有任何文字，确认是否为纯图片页并检查可编辑性")

    # ---- 可编辑性
    infos.append(f"可编辑文本片段数: {editable_text_boxes}")
    if editable_text_boxes < n:
        warns.append("[可编辑性] 文本片段数少于页数，可能存在整页图片化的情况")

    # ---- AI 味套话
    ban = ["综上所述", "总而言之", "不容忽视", "具有重要的意义", "发挥着重要作用",
           "随着", "在当今", "谢谢观看", "感谢聆听", "汇报时间", "汇报时长"]
    alltext = []
    for _, shapes in shapes_by_slide:
        for sh in shapes:
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                alltext.append(sh.text_frame.text)
    blob = "\n".join(alltext)
    hitban = [b for b in ban if b in blob]
    if hitban:
        warns.append(f"[AI味/废话] 命中禁用语: {hitban}")

    report = {
        "file": os.path.abspath(path),
        "slides": n,
        "info": infos,
        "failed": fails,
        "warned": warns,
        "pass": not fails,
    }

    if "--json" in sys.argv:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"\n=== PPTX 自检: {os.path.basename(path)} ===")
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
