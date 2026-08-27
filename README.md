# ClawVision 1.0

Generate visual HTML summaries, Markdown, and PowerPoint from OpenClaw session history.

## What it does

ClawVision turns an OpenClaw conversation into a self-contained, tabbed HTML card — like a Codex `$visualize` card, but local.

Output format:

- **Main takeaway** — the single most important conclusion.
- **Format** — how the discussion was structured, with do's and don'ts.
- **What we built** — checklist of outcomes.
- **Next steps** — what to do next.

Each card also renders a flow diagram, metric cards, EN/RU/ZH language switcher, light/dark theme toggle, and **Export Markdown / Export PowerPoint** buttons.

## Requirements

- Python 3.10+
- `playwright` (`pip install playwright` and `playwright install chromium`)
- `python-pptx` (`pip install python-pptx`)

## Usage

1. Export your conversation summary as JSON matching the schema below.
2. Run the generator:

```bash
python scripts/generate_visual.py \
  --summary summary.json \
  --output ./out \
  --png --md --pptx \
  --lang en
```

3. Get `out/<slug>.html`, `out/<slug>.md`, `out/<slug>.pptx`, and `out/<slug>_tab*.png`.

## JSON schema

```json
{
  "title": "Short title in conversation language",
  "subtitle": "One-line context",
  "main_takeaway": "The single most important conclusion",
  "format_takeaway": "How the discussion was structured",
  "next_takeaway": "What the next move should be",
  "flow": [
    {"label": "Step 1", "sub": "what happened"},
    {"label": "→", "sub": ""},
    {"label": "Step 2", "sub": "what happened"}
  ],
  "metrics": [
    {"title": "Goal", "text": "..."},
    {"title": "Approach", "text": "..."},
    {"title": "Output", "text": "..."}
  ],
  "dos": ["good practice 1", "good practice 2"],
  "donts": ["risk 1", "risk 2"],
  "checklist": [
    {"text": "Item name", "status": "ready|pending|blocked"}
  ],
  "next_steps": ["action 1", "action 2"]
}
```

## Outputs

- `*.html` — self-contained tabbed card with language switcher and theme toggle.
- `*.md` — plain Markdown summary.
- `*.pptx` — 7-slide PowerPoint deck.
- `*.png` — default screenshot of the first tab.
- `*_tab*.png` — one screenshot per tab.

## Languages

The card UI supports **EN / RU / ZH**. Pass `--lang en|ru|zh` to control tab labels, export buttons, and badge text.

## Version

This is **ClawVision 1.0**. Future version 2.0 will add session analytics: message stats, tool usage, topic extraction, and insights.

## License

MIT
