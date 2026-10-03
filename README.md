# academic-deliverables

A specification skill for producing **academic PPT and Word deliverables**. It converts common rework problems—excessive color, plain-text slides, chart overload, AI-flavored writing, non-editable outputs, and fabricated references—into reusable production and validation rules.

> Trigger: ask an AI assistant 
>
> **"use my academic PPT skill to make ..."**
>
> ; it will load the spec and confirm the key parameters before starting.

## Motivation

Academic deliverables repeatedly fail for a small set of structural, visual, and verification reasons. This skill codifies those failure modes into structured rules and ships two self-check scripts, making "done and verified" the default standard rather than "done, please review."

## Structure



```
├── SKILL.md                  # Main spec: routing, intake, red lines, workflow, self-check
├── references/               # Five spec files: PPT / Word / feedback / anti-AI / intake
├── assets/style-tokens.json  # Machine-readable style tokens
└── scripts/                  # PPTX / DOCX self-check and workspace scaffolding
```

## Install



```
git clone <repo-url> academic-deliverables
cd academic-deliverables
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

## Usage

Run the self-check before every delivery and attach the report:



```
.venv/bin/python scripts/check_pptx.py  slides.pptx    # colors / font / size / density / AI-flavor
.venv/bin/python scripts/check_docx.py  report.docx    # fonts / spacing / color / OMML / AI-flavor
```

### Specification highlights



* **PPT**: at most 2 hues, body font **Microsoft YaHei**, annotation font **KaiTi** (no other font families), minimum font size 10pt, one visual anchor per slide, at most 15 slides, no AI-flavor phrases.

* **Word**: SimSun body (12pt), SimHei headings, KaiTi annotations, 1.5 line spacing, black text only, formulas as OMML objects, academic three-line tables.

Detailed rules live in `references/` and are loaded only when relevant.

## License

MIT © 2026 contributors
