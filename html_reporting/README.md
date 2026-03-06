# html_reporting — Premium HTML + PDF Report Generation

Force plate assessment reports for **Move, Measure, Analyze LLC**.

Generates scrollable HTML reports and controlled multi-page PDFs from Hawkin Dynamics data using the existing Python computation engine.

## Architecture

```
html_reporting/
    __init__.py
    __main__.py          # python -m html_reporting entry point
    cli.py               # CLI argument parsing and orchestration
    payload.py           # Wraps existing compute logic into structured dict
    render_html.py       # Jinja2 template rendering
    export_pdf.py        # Playwright (Chromium) PDF export
    smoke_test.py        # End-to-end smoke tests
    mock_payload.json    # Sample data for offline testing
    templates/
        report.html      # Main Jinja2 template
        partials/
            page_footer.html
    styles/
        screen.css       # Screen / browser styles
        print.css        # Print / PDF-specific styles
    assets/              # Generated plot images (runtime)
    README.md
```

## Prerequisites

Python 3.9+

### Install dependencies

```bash
pip install jinja2 playwright numpy matplotlib pandas
```

### Install Playwright browser

```bash
playwright install chromium
```

### Existing engine

The module imports `Individual_Report_Generator_v4.py` from the parent directory. Ensure `hdforce` and its dependencies are installed for live API mode:

```bash
pip install hdforce anthropic pymupdf
```

## Usage

### Live mode (Hawkin API)

```bash
python -m html_reporting.cli \
    --athlete-id <ATHLETE_ID> \
    --out ./output \
    --format both
```

With AI interpretations:

```bash
python -m html_reporting.cli \
    --athlete-id <ATHLETE_ID> \
    --out ./output \
    --format both \
    --enable-ai
```

Options:
- `--days-back 60` — lookback window (default: 60)
- `--team "Wildcats"` — team name for the report
- `--sport "Basketball"` — sport label
- `--format html|pdf|both` — output format

### Offline mode (JSON payload)

```bash
python -m html_reporting.cli \
    --input-json html_reporting/mock_payload.json \
    --out ./output \
    --format both
```

### Programmatic usage

```python
from html_reporting.payload import build_payload
from html_reporting.render_html import render_to_file
from html_reporting.export_pdf import export_pdf

# Build payload from your found_tests dict
payload = build_payload(
    athlete_name="Alex Johnson",
    found_tests=found_tests,  # from existing engine
    body_weight_n=810.0,
    team="Wildcats",
    sport="Basketball",
)

# Render HTML
html_path = render_to_file(payload, "report.html", inline_styles=True)

# Export PDF
pdf_path = export_pdf(html_path, "report.pdf")
```

## Running smoke tests

```bash
python -m html_reporting.smoke_test
```

Tests cover:
1. Payload structure validation
2. HTML rendering from mock data
3. PDF export via Playwright

## Report structure

| Page | Content |
|------|---------|
| Cover | Athlete name, team, date, branding |
| Dashboard | Performance snapshot tiles, AI summary |
| Test pages | Force-time plot, key metrics, asymmetry indicators |
| Trend pages | Multi-session trend lines and data tables |
| Asymmetry | Summary table with OK/MONITOR/ADDRESS status |

## PDF quality controls

- Playwright `page.pdf()` with Letter format
- `print.css` enforces page breaks between sections
- `break-inside: avoid` prevents split charts, cards, and tables
- CSS counters for page numbering
- `print-color-adjust: exact` preserves background colors

## Brand

- **Colors**: Black (#0a0a0a), Gold (#f4bd2a), White
- **Headings**: Bebas Neue
- **Body**: Source Sans 3
- **Style**: Premium performance lab, minimalist, card-based
