# ClawVision

Generate visual HTML summaries, Markdown, and PowerPoint from OpenClaw session history.

[![ClawHub](https://img.shields.io/badge/ClawHub-clawvision-blue)](https://clawhub.ai/monaxamo/skills/clawvision)
[![Version](https://img.shields.io/badge/version-1.0.3-green)](https://github.com/monaxamo/clawvision/releases)

---

## What it does

ClawVision turns an OpenClaw chat session into a clean, tabbed HTML infographic — like a Codex `$visualize` card, but local. It also exports to Markdown and a redesigned, visual PowerPoint deck.

![ClawVision demo](clawvision_demo_en.png)

---

## Features

- **Self-contained HTML cards** with 4 tabs:
  - Main takeaway
  - Format
  - What we built
  - Next steps
- **Flow diagram** and **metric cards** for each summary.
- **EN / RU / ZH language switcher**.
- **Light / dark theme toggle**.
- **Export buttons** for Markdown and PowerPoint right inside the card.
- **PNG screenshots**: one default + one per tab.
- **Markdown export** and **PowerPoint deck export** with visual, card-based design.
- Fully local: no external APIs for rendering; uses Playwright + `python-pptx`.

---

## Requirements

- Python 3.10+
- Playwright (`pip install playwright && playwright install chromium`)
- `python-pptx` (`pip install python-pptx`)

---

## Usage

```bash
python scripts/generate_visual.py \
  --summary summary.json \
  --output ./out \
  --png --md --pptx \
  --lang en
```

Outputs:

- `out/<slug>.html`
- `out/<slug>.md`
- `out/<slug>.pptx`
- `out/<slug>.png`
- `out/<slug>_tab1.png` … `out/<slug>_tab4.png`

---

## Example summary JSON

```json
{
  "title": "OpenClaw skill design session",
  "subtitle": "Building ClawVision 1.0",
  "main_takeaway": "A local, privacy-first visual summary tool is viable and fast.",
  "format_takeaway": "Structured conversation with clear deliverables.",
  "next_takeaway": "Publish the skill and gather feedback.",
  "flow": [
    {"label": "Idea", "sub": "Codex $visualize"},
    {"label": "→", "sub": ""},
    {"label": "Design", "sub": "HTML card + exports"},
    {"label": "→", "sub": ""},
    {"label": "Build", "sub": "generate_visual.py"}
  ],
  "metrics": [
    {"title": "Goal", "text": "Visual summary from chat"},
    {"title": "Approach", "text": "Local renderer"},
    {"title": "Output", "text": "HTML + MD + PPTX + PNG"}
  ],
  "dos": ["Keep summaries local", "Confirm sensitive content before exporting"],
  "donts": ["Expose private data", "Run on vague requests"],
  "checklist": [
    {"text": "HTML renderer", "status": "ready"},
    {"text": "Markdown export", "status": "ready"},
    {"text": "PowerPoint export", "status": "ready"}
  ],
  "next_steps": ["Publish to ClawHub", "Collect user feedback"]
}
```

See `example_summary_en.json` and `course_session_summary_en.json` for full examples.

---

## Install as an OpenClaw skill

```bash
openclaw skills install clawvision
```

Or install from source:

```bash
git clone https://github.com/monaxamo/clawvision.git
# Place the skill folder in your OpenClaw workspace skills directory
```

---

## Roadmap

- **1.0** — visual summaries, language switcher, theme toggle, export formats.
- **2.0** — session analytics: message stats, tool usage, topic/entity extraction, CSS-only charts, insights.

---

## Related

- **[clawvision-plus](https://github.com/monaxamo/clawvision-plus)** — companion plugin with PDF export, OG image generation, and Telegram sharing.

---

## License

MIT
