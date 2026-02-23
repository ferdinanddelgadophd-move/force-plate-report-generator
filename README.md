# Force Plate Report Generator

Automated report generation system for Hawkin Dynamics force plate assessments. Transforms raw API data into professionally formatted PDF reports for individual athletes and teams.

Built for [Move, Measure, Analyze LLC](https://movemeasure.com) — a mobile sports science service providing force plate assessments for athletes and teams.

---

## Overview

This system pulls athlete data directly from the Hawkin Dynamics API, processes force-time metrics, applies a custom biomechanical framework, and generates client-ready PDF reports in under 60 seconds.

### Two Report Types

| Individual Reports | Team Reports |
|-------------------|--------------|
| 4+ assessments per athlete | As many athletes per session |
| Force-time curve visualizations | Team performance dashboard |
| Asymmetry analysis across all tests | Quadrant-based athlete profiling |
| Benchmark comparisons | Injury risk flagging |
| Personalized recommendations | Leaderboards & full roster |

---

## Sample Reports

### Individual Athlete Report

**Performance Dashboard** — Key metrics at a glance with multiple assessment types:

![Individual Dashboard](samples/individual_dashboard.png)

**Force-Time Curves** — Bilateral force traces with real-time asymmetry tracking:

![CMJ Assessment](samples/individual_cmj.png)

**Asymmetry Summary** — Cross-test comparison with status flags:

![Asymmetry Summary](samples/individual_asymmetry.png)

---

### Team Report

**Team Performance Dashboard** — Baseline metrics with descriptive statistics:

![Team Dashboard](samples/team_dashboard.png)

**Performance Profiles** — Quadrant scatter plots categorizing athletes by explosiveness and force application strategy:

![Performance Profiles](samples/team_profiles.png)

**Injury Risk Assessment** — Landing force and asymmetry flags with actionable status levels:

![Injury Risk](samples/team_injury_risk.png)

**Training Focus Matrix** — Profile-based training recommendations:

![Training Focus](samples/team_training_focus.png)

---

## Technical Implementation

### Architecture

```
Hawkin Dynamics API
        │
        ▼
┌───────────────────┐
│   Data Ingestion  │  ← API authentication, pagination, date filtering
│   & Processing    │  ← Test type detection, metric extraction
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  Analysis Layer   │  ← OUTPUT/STRATEGY/DRIVER framework
│                   │  ← Quadrant classification, asymmetry flags
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  PDF Generation   │  ← ReportLab rendering, matplotlib visualizations
│                   │  ← Custom branding, professional formatting
└───────────────────┘
        │
        ▼
   Client-Ready PDF
```

### Key Features

**Data Pipeline**
- Direct Hawkin Dynamics API integration with token authentication
- Auto-detection of available test types (CMJ, Squat Jump, Multi Rebound, Drop Landing, etc.)
- Best-trial selection logic per assessment type
- Handles pagination and rate limiting

**Analysis**
- Custom **OUTPUT / STRATEGY / DRIVER** metric framework
- Quadrant-based athlete profiling (Explosive, Powerful, Fast/Reactive, Building)
- Asymmetry detection with tiered status levels (OK < 10%, MONITOR 10-15%, ADDRESS > 15%)
- Landing force risk thresholds (> 4.0× body weight flagged)
- Sport-specific normative data (Track & Field built out, extensible architecture)

**Report Generation**
- Professional PDF output with ReportLab
- Force-time curve visualization with matplotlib
- Left/right force traces with shaded asymmetry regions
- Benchmark visualizations with color-coded ranges
- Custom font handling and branded templates

### Tech Stack

- **Python 3.x**
- **hdforce** — Hawkin Dynamics API wrapper
- **pandas / numpy** — Data processing
- **matplotlib** — Force-time visualizations
- **ReportLab** — PDF generation
- **PyPDF2 / pypdf** — PDF manipulation (landscape page insertion)

---

## Metric Framework

Reports organize metrics into three categories:

### OUTPUT — What You Produced
- Jump Height (cm)
- Jump Momentum (kg·m/s)
- Peak Relative Power (W/kg)

### STRATEGY — How You Jumped
- RSImod (explosive efficiency)
- Time to Takeoff (s)
- Countermovement Depth (cm)

### DRIVERS — How Much Force You Used
- Relative Propulsive Impulse (N·s/kg)
- Relative Braking RFD (N/s/kg)

This framework helps athletes and coaches understand not just *what* happened, but *how* and *why*.

---

## Supported Assessments

| Test Type | Individual | Team |
|-----------|:----------:|:----:|
| Countermovement Jump (CMJ) | ✓ | ✓ |
| Squat Jump | ✓ | ✓ |
| Multi Rebound | ✓ | ✓ |
| CMJ Rebound | ✓ | ✓ |
| Drop Landing | ✓ | ✓ |
| Bodyweight Squats | ✓ | — |
| Static Balance | ✓ | — |

---

## Code Statistics

| Component | Lines of Code |
|-----------|---------------|
| Team Report Generator | ~1,700 |
| Individual Report Generator | ~1,200 |
| **Total** | **~2,900** |

---

## Output Examples

Full sample reports are available in the `/samples` directory:

- `Sample_Individual_Report.pdf` — 7-page individual athlete assessment
- `Sample_Team_Report.pdf` — 10-page team performance report

---

## About

Developed by **Ferdinand Delgado, PhD** for Move, Measure, Analyze LLC.

- [Portfolio](https://ferdinanddelgadophd-move.github.io)
- [LinkedIn](https://linkedin.com/in/ferdinanddelgado)
- [Move, Measure, Analyze](https://movemeasure.com)
