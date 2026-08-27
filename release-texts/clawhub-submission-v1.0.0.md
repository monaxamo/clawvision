> **Name:** clawvision  
> **Version:** 1.0.0  
> **License:** MIT  
> **Homepage:** https://github.com/openclaw/skills/clawvision  
> **Tags:** visualization, summary, html, sessions, codex, markdown, powerpoint, sessions_history, node_inference  
> **Author:** Maximius  

## Short description

ClawVision 1.0 turns an OpenClaw chat session into a self-contained, tabbed HTML card — like a Codex `$visualize` card, but local. It also exports the summary to Markdown and PowerPoint.

## What it does

- Generates a 4-tab HTML summary card: Main takeaway, Format, What we built, Next steps.
- Renders a flow diagram, metric cards, and a checklist.
- Supports EN/RU/ZH language switcher and light/dark theme toggle.
- Exports to Markdown and PowerPoint.
- Produces one PNG per tab plus a default first-tab PNG.
- Runs fully locally with Playwright and `python-pptx`.

## Tools used

`sessions_history`, `sessions_list`, `node_inference`, `write`, `read`, `exec`, `skill_workshop`

## Example output

- `clawvision_demo_en_2026-08-27.html` — self-contained card.
- `clawvision_demo_en_2026-08-27_collage.png` — 2×2 collage of all 4 tabs.
- `session_vibe_setup_en_2026-08-27_collage.png` — real session: vibe.egetech.ru setup.
- `session_translation_ru_2026-08-27_collage.png` — real session: book translation.

## Usage

```bash
python scripts/generate_visual.py \
  --summary summary.json \
  --output ./out \
  --png --md --pptx \
  --lang en
```

## Requirements

- Python 3.10+
- Playwright + Chromium
- `python-pptx`

## Roadmap

- **1.0** — visual summaries, language switcher, theme toggle, export formats.
- **2.0** — session analytics: message stats, tool usage, topic/entity extraction, CSS-only charts, insights.
