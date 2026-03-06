# Automated Force Plate Assessment Report Generator

> A premium HTML/PDF reporting system for Hawkin Dynamics force plate data

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.x-lightgrey?logo=flask)
![Playwright](https://img.shields.io/badge/Playwright-Chromium-green?logo=playwright)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## Overview

Sport scientists and strength coaches spend hours manually compiling force plate data into presentable reports. This system automates the entire pipeline — from pulling raw data via the Hawkin Dynamics API to generating premium, branded, multi-page HTML and PDF assessment reports.

The result: comprehensive athlete reports generated in under 30 seconds, complete with force-time visualizations, asymmetry analysis, multi-session trend tracking, derived performance metrics, and optional AI-powered interpretations.

---

## Key Features

- **Interactive Web Configuration** — Flask app for athlete search, test/trial selection, session pairing, custom summaries, and AI toggle
- **Premium Report Design** — Branded HTML reports with Bebas Neue / Source Sans 3 typography and a gold/black/white aesthetic
- **Performance Dashboard** — Snapshot tiles with key metrics, delta indicators, and gauge visualizations for derived metrics (EUR, DSI)
- **Force-Time Visualizations** — Matplotlib-generated force-time curves with left/right asymmetry shading embedded as base64 PNGs
- **Multi-Session Trends** — Session-over-session tracking with trend charts and data tables showing performance changes
- **Asymmetry Analysis** — Comprehensive left/right asymmetry assessment with configurable OK / Monitor / Address thresholds
- **PDF Export** — Playwright-powered headless Chromium rendering to print-ready Letter-format PDFs with controlled page breaks
- **AI Interpretations** — Optional Claude-powered narrative summaries for dashboard overview and asymmetry analysis
- **Derived Metrics** — EUR (Elastic Utilization Ratio) and DSI (Dynamic Strength Index) with visual gauge bars showing performance zones
- **Multiple Input Modes** — Live API, CLI interactive, CLI direct, offline JSON, and web app interfaces

---

## Architecture

```
Hawkin Dynamics API
        │
  hdforce package
        │
  Engine (computation, metrics, plots)
        │
  payload.py (structured dict)
        │
  Jinja2 + CSS (HTML rendering)
        │
   ┌────┴────┐
   │         │
 HTML      PDF
(browser) (Playwright)
```

### Module Breakdown

| Module | Purpose |
|--------|---------|
| `Individual_Report_Generator_v4.py` | Core computation engine — force analysis, asymmetry, trends, metric calculations |
| `html_reporting/payload.py` | Wraps engine output into structured dict; renders Matplotlib plots as base64 PNGs |
| `html_reporting/render_html.py` | Jinja2 templating — converts payload dict into self-contained HTML with inline CSS |
| `html_reporting/export_pdf.py` | Playwright (Chromium) — renders HTML to print-ready multi-page PDF |
| `html_reporting/cli.py` | CLI orchestration — argument parsing, athlete/test selection, pipeline execution |
| `html_reporting/webapp/app.py` | Flask web application — interactive UI for configuration and report preview |
| `html_reporting/webapp/hawkin_service.py` | API service layer — wraps engine for Flask routes, handles authentication |

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Data Source | Hawkin Dynamics REST API (`hdforce` package) |
| Computation | Python 3.11, NumPy, Pandas |
| Visualization | Matplotlib (force-time plots, trend charts) |
| Web App | Flask (routes, API endpoints, interactive UI) |
| Templating | Jinja2 (HTML report generation) |
| PDF Export | Playwright (Chromium sync API, Letter format) |
| AI | Anthropic Claude API (optional interpretations) |
| Styling | Custom CSS (screen + print), Google Fonts (Bebas Neue, Source Sans 3) |
| Brand Colors | Black (`#0a0a0a`), Gold (`#f4bd2a`), White (`#ffffff`) |

---

## Screenshots

> *Screenshots coming soon — will include web app configuration, dashboard page, test assessment, trend analysis, asymmetry summary, and PDF output.*

---

## Installation & Setup

```bash
git clone https://github.com/ferdinanddelgadophd-move/force-plate-report-generator.git
cd force-plate-report-generator
pip install -r requirements.txt
playwright install chromium
```

### Environment Variables

Set these before running:

```bash
export HAWKIN_API_TOKEN="your-hawkin-api-token"
export ANTHROPIC_API_KEY="your-anthropic-key"   # optional, for AI interpretations
```

---

## Usage

### 1. Web App (Recommended)

```bash
python -m html_reporting.webapp
```

Open `http://localhost:5000` in your browser. Search for an athlete, configure test selection, toggle AI, and generate reports interactively.

### 2. CLI Interactive

```bash
python -m html_reporting.cli
```

Walks you through athlete selection, test picking, and report generation step by step.

### 3. CLI Direct

```bash
python -m html_reporting.cli \
    --athlete-id <ATHLETE_ID> \
    --out ./output \
    --format both \
    --enable-ai
```

### 4. Offline (JSON Payload)

```bash
python -m html_reporting.cli \
    --input-json sample_data.json \
    --out ./output \
    --format both
```

### CLI Options

| Flag | Description | Default |
|------|-------------|---------|
| `--athlete-id` | Hawkin athlete ID | Interactive selection |
| `--out` | Output directory | `./output` |
| `--format` | `html`, `pdf`, or `both` | `both` |
| `--enable-ai` | Enable Claude AI interpretations | Off |
| `--days-back` | Lookback window for historical data | `60` |
| `--team` | Team name for branding | — |
| `--sport` | Sport label | — |
| `--input-json` | Path to offline JSON payload | — |

---

## Report Structure

| Page | Content |
|------|---------|
| **Cover** | Athlete name, team, sport, date, Move Measure Analyze branding |
| **Dashboard** | Performance snapshot tiles with key metrics, delta indicators, gauge visualizations (EUR, DSI), and optional AI narrative summary |
| **Test Pages** | Force-time plot, key metrics table, asymmetry indicators (one page per test type) |
| **Trend Pages** | Multi-session trend charts with data tables showing session-over-session changes |
| **Asymmetry Summary** | Comprehensive left/right asymmetry table with OK / Monitor / Address status indicators |

---

## Configuration Options

- **Test Selection** — Choose which test types to include (CMJ, SJ, IMTP, DJ, etc.)
- **Trial Selection** — Best trial from latest session auto-selected
- **Session Pairing** — Select which sessions to compare for derived metrics (EUR, DSI)
- **Custom Summaries** — Add free-text notes to the report
- **AI Toggle** — Enable/disable Claude-powered interpretations
- **Output Format** — HTML only, PDF only, or both
- **Team/Sport Branding** — Customize report header with team and sport labels

---

## PDF Quality Controls

- Playwright `page.pdf()` with Letter format (8.5" x 11")
- `print.css` enforces page breaks between sections
- `break-inside: avoid` prevents split charts, cards, and tables
- CSS counters for automatic page numbering
- `print-color-adjust: exact` preserves background colors and gradients

---

## Project Structure

```
force-plate-report-generator/
├── README.md
├── requirements.txt
├── .gitignore
├── Individual_Report_Generator_v4.py
├── html_reporting/
│   ├── __init__.py
│   ├── __main__.py
│   ├── cli.py
│   ├── payload.py
│   ├── render_html.py
│   ├── export_pdf.py
│   ├── smoke_test.py
│   ├── templates/
│   │   ├── report.html
│   │   └── partials/
│   │       └── page_footer.html
│   ├── styles/
│   │   ├── screen.css
│   │   └── print.css
│   └── webapp/
│       ├── __init__.py
│       ├── __main__.py
│       ├── app.py
│       ├── hawkin_service.py
│       ├── templates/
│       │   ├── layout.html
│       │   ├── index.html
│       │   ├── configure.html
│       │   └── preview.html
│       └── static/
│           ├── app.js
│           └── style.css
└── docs/
    └── screenshots/
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Author

**Ferdinand Delgado, PhD**
[Portfolio](https://ferdinanddelgadophd-move.github.io/) · [GitHub](https://github.com/ferdinanddelgadophd-move)
