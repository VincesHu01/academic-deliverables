#!/usr/bin/env python3
"""DOCX 交付物自检 —— 按 Word 交付规格逐条检查。

字体: 宋体小四正文 / 黑体三号四号小四标题 / 楷体五号说明 / Times New Roman 斜体公式
行距: 1.5 倍   颜色: 仅黑色   公式: 必须是 OMML 对象
用法: python3 check_docx.py <file.docx> [--json] [--mode default|reflection]
      --mode reflection  追加"读后感 / 课程感悟 / 读书报告"文体检查
                         规范见 references/reflection-spec.md
退出码: 0 = 通过(可含警告)  1 = 有失败项(含读后感禁用词)
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

# ---------- 读后感文体检查常量（--mode reflection） ----------
REFLECT_HARD_BAN = ["受益匪浅", "干货满满", "深入浅出", "娓娓道来", "令人深思", "意犹未尽",
                    "醍醐灌顶", "豁然开朗", "精彩纷呈", "我收获很大", "学到了很多",
                    "开阔了眼界", "提升了认知", "拓宽了视野", "在老师的悉心教导下",
                    "让我对……有了更深的理解", "对……有了更深刻的认识"]
REFLECT_SOFT_BAN = ["综上所述", "总而言之", "值得我们深思", "具有重要意义", "令人印象深刻",
                    "感触颇深", "深有体会", "收获颇丰", "发人深省"]
EMOJI = re.compile("[\U0001F300-\U0001FAFF\u2600-\u27BF\uFE0F]")
CONCRETE = re.compile(r"\d|《|》|「|」|“|”|\"|元|人|户|年|次|期|页|公里|亩|%|第[一二三四五六七八九十]讲")
ERROR_ADMIT = ["我以前", "过去我", "本科时", "从前我", "我原来的", "我曾", "我一度", "我原来",
               "草草", "敷衍", "功利", "走马观花", "不严谨", "我意识到自己", "我忽视了",
               "我当时只", "我当时把", "我承认"]
CONCEDE = ["或许", "不一定", "并非", "有待", "局限", "边界", "我保留", "但我也", "同时我也",
           "不过我也", "相较而言", "需要补充", "这里我存疑", "也有反例"]
FIVE_STEP_MARKS = {
    "旧认知": ["以前", "以往", "过去", "起初", "我原以为", "本科时", "一开始"],
    "触发": ["老师", "课堂", "讲座", "这本书", "作者", "本讲", "讲到", "报告中"],
    "断裂": ["但这", "然而", "我才", "让我重新", "这让我", "现在想来", "反过来", "恰恰"],
    "新命题": ["于是", "由此", "我更倾向于", "我认为", "我的判断是", "可以理解为", "这意味着"],
    "回扣自身": ["我以后", "今后", "接下来我", "在我的", "我手上", "我自己的", "回到我"],
}


def zh_len(s):
    return len(re.findall(r"[\u4e00-\u9fff]", s))


def para_list(d):
    """提取非空段落，标记是否为标题样式。"""
    out = []
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        style = (p.style.name or "").lower()
        out.append({"text": t, "heading": ("heading" in style) or ("标题" in style)})
    if out:
        out[0]["heading"] = True          # 首段视为标题
    return out


def check_reflection(d, fails, warns, infos):
    """读后感 / 课程感悟 / 读书报告 文体检查（启发式，命中即提示回看原文）。"""
    paras = para_list(d)
    blob = "\n".join(p["text"] for p in paras)
    body = [p for p in paras if not p["heading"] and zh_len(p["text"]) >= 80]

    # 1 绝对禁用词 → 失败
    hits = [w for w in REFLECT_HARD_BAN if w in blob]
    if hits:
        fails.append(f"[禁用词] 命中读后感绝对禁用表达: {hits}")
    # 2 软性套话
    soft = [w for w in REFLECT_SOFT_BAN if w in blob]
    if soft:
        warns.append(f"[套话] 需替换为具体观察: {soft}")

    # 3 机械四段
    four = all(k in blob for k in ("首先", "其次", "再次", "最后"))
    three = all(k in blob for k in ("首先", "其次", "最后"))
    if four:
        fails.append("[结构] 出现「首先/其次/再次/最后」机械四段，改为按认知推进组织")
    elif three:
        warns.append("[结构] 出现「首先/其次/最后」三段式痕迹，检查是否沦为罗列")

    # 4 感叹号
    bangs = blob.count("!") + blob.count("！")
    if bangs > 1:
        warns.append(f"[语气] 感叹号 {bangs} 个，规范要求 ≤1")

    # 5 emoji
    if EMOJI.search(blob):
        fails.append("[语气] 出现 emoji 或装饰符号，读后感不得使用")

    # 6 段内小标题
    extra = [p["text"][:20] for p in paras[1:] if p["heading"]]
    if extra:
        warns.append(f"[排版] 正文含 {len(extra)} 个小标题: {extra[:3]}，读后感应为连续段落")

    # 7 五步单元覆盖
    owned = 0
    for p in body:
        if sum(1 for ws in FIVE_STEP_MARKS.values() if any(w in p["text"] for w in ws)) >= 2:
            owned += 1
    cover = owned / len(body) if body else 0
    if body:
        infos.append(f"五步单元覆盖率: {cover:.0%}（{owned}/{len(body)} 段命中≥2 个环节）")
    if body and cover < 0.5:
        warns.append("[五步单元] 不足半数段落跑完「旧认知→触发→断裂→新命题→回扣自身」，接近总-分-总")

    # 8 具体物密度
    bare = [p["text"][:18] for p in body if zh_len(p["text"]) >= 120 and not CONCRETE.search(p["text"])]
    if bare:
        warns.append(f"[具体物] {len(bare)} 段无数字/引号/量词，示例: {bare[:2]}")

    # 9 认错机制
    if not any(w in blob for w in ERROR_ADMIT):
        warns.append("[认错] 未见具体的「我过去是错的」事件，易被读成客套总结")
    # 10 保留机制
    if not any(w in blob for w in CONCEDE):
        warns.append("[保留] 全文对课上/书中观点没有任何保留或边界，缺少思辨")

    # 11 结尾姿态
    tail = "".join(p["text"] for p in paras[-3:]) if len(paras) >= 3 else blob
    if not (("不是" in tail and "而是" in tail) or "最需要" in tail or "而在于" in tail):
        warns.append("[结尾] 未见「最需要的不是 X，而是 Y」式回落，可能只是复述内容")

    # 12 段长
    lens = [zh_len(p["text"]) for p in paras if not p["heading"]]
    if lens:
        infos.append(f"段落数: {len(lens)}  段长区间: {min(lens)}–{max(lens)} 汉字")
        over = [n for n in lens if n > 600]
        short = [n for n in lens if 0 < n < 60]
        if over:
            warns.append(f"[段长] {len(over)} 段超过 600 字，建议拆到 250–450 字")
        if short:
            warns.append(f"[段长] {len(short)} 段不足 60 字，易显碎片化")


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


def parse_args(argv):
    """解析参数 → (paths, json_out, mode)。"""
    paths, json_out, mode, i = [], False, "default", 0
    while i < len(argv):
        a = argv[i]
        if a == "--json":
            json_out = True
        elif a == "--mode":
            if i + 1 >= len(argv):
                sys.exit("--mode 需要一个值: default | reflection")
            mode = argv[i + 1].lower()
            i += 1
        elif a.startswith("--mode="):
            mode = a.split("=", 1)[1].lower()
        elif a.startswith("--"):
            sys.exit(f"未知参数: {a}")
        else:
            paths.append(a)
        i += 1
    if mode not in ("default", "reflection"):
        sys.exit(f"未知 mode: {mode}（可选: default | reflection）")
    return paths, json_out, mode


def main():
    paths, json_out, mode = parse_args(sys.argv[1:])
    if not paths:
        print(__doc__)
        sys.exit(2)
    path = paths[0]
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

    # ---------- 读后感 / 课程感悟 文体（--mode reflection） ----------
    if mode == "reflection":
        check_reflection(d, fails, warns, infos)

    # ---------- 字数 ----------
    zh = len(re.findall(r"[\u4e00-\u9fff]", blob))
    infos.append(f"中文字符数(段落): {zh}")

    report = {"file": os.path.abspath(path), "mode": mode, "info": infos,
              "failed": fails, "warned": warns, "pass": not fails}
    if json_out:
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
