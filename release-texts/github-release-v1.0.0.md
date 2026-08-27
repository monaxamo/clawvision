# ClawVision 1.0

Generate visual HTML summaries, Markdown, and PowerPoint from OpenClaw session history.

## What’s in 1.0

- **Self-contained HTML cards** with 4 tabs: Main takeaway, Format, What we built, Next steps.
- **Flow diagram** and **metric cards** for each summary.
- **EN / RU / ZH language switcher**.
- **Light / dark theme toggle**.
- **Export buttons** for Markdown and PowerPoint right inside the card.
- **PNG screenshots**: one default + one per tab.
- **Markdown export** and **PowerPoint deck export** (7 slides).
- Fully local: no external APIs for rendering, uses Playwright + `python-pptx`.

## Use it

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

## Requirements

- Python 3.10+
- Playwright (`pip install playwright && playwright install chromium`)
- `python-pptx` (`pip install python-pptx`)

## Examples

See `example_summary_en.json`, `course_session_summary_en.json`, and generated screenshots in the repo.

## Roadmap

- **1.0** — visual summaries, language switcher, theme toggle, export formats.
- **2.0** — session analytics: message stats, tool usage, topic/entity extraction, CSS-only charts, insights.

## License

MIT
