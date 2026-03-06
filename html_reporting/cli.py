"""
cli.py -- Command-line interface for HTML/PDF report generation.

Usage:
    Interactive (pick athlete from list):
        python -m html_reporting.cli

    Direct (skip menu):
        python -m html_reporting.cli --athlete-id <id> --out <dir> --format both

    Offline testing:
        python -m html_reporting.cli --input-json mock_payload.json --out <dir> --format html
"""

import argparse
import os
import sys
import json
from datetime import datetime

_PARENT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PARENT_DIR not in sys.path:
    sys.path.insert(0, _PARENT_DIR)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate premium HTML/PDF force plate reports.",
        prog="python -m html_reporting.cli",
    )

    # Data source (all optional now -- interactive mode if none provided)
    parser.add_argument("--athlete-id", type=str, default=None, help="Hawkin athlete ID (skip menu)")
    parser.add_argument("--input-json", type=str, default=None, help="Path to a JSON payload file (offline mode)")
    parser.add_argument("--assessment-id", type=str, help="Hawkin assessment/test ID (optional)")
    parser.add_argument("--days-back", type=int, default=365, help="Days of history to fetch (default: 365)")
    parser.add_argument("--out", type=str, default="./output", help="Output directory")
    parser.add_argument("--format", type=str, choices=["html", "pdf", "both"], default="both",
                        help="Output format (default: both)")
    parser.add_argument("--enable-ai", action="store_true", help="Enable AI interpretations")
    parser.add_argument("--team", type=str, default=None, help="Team name")
    parser.add_argument("--sport", type=str, default=None, help="Sport name")

    return parser.parse_args()


def pick_athlete_interactive(engine):
    """Show numbered list of athletes and let user pick one (or multiple)."""
    print("\nFetching athletes...")
    athletes = engine.get_athletes()

    if athletes is None or athletes.empty:
        print("ERROR: No athletes found.")
        sys.exit(1)

    # Sort by name
    athletes = athletes.sort_values("name").reset_index(drop=True)

    print(f"\n{'='*50}")
    print(f"  ATHLETES ({len(athletes)} total)")
    print(f"{'='*50}")
    for i, (_, row) in enumerate(athletes.iterrows()):
        name = row.get("name", "Unknown")
        print(f"  {i + 1:3d}. {name}")
    print(f"{'='*50}")

    while True:
        choice = input("\nEnter athlete number (or 'q' to quit): ").strip()
        if choice.lower() == 'q':
            print("Exiting.")
            sys.exit(0)
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(athletes):
                row = athletes.iloc[idx]
                athlete_id = row.get("id")
                athlete_name = row.get("name", "Unknown")
                print(f"\nSelected: {athlete_name}")
                return athlete_id, athlete_name
            else:
                print(f"Please enter a number between 1 and {len(athletes)}")
        except ValueError:
            print("Please enter a number")


def discover_available_tests(tests_df, engine):
    """Scan the tests DataFrame and return a list of available test types with metadata."""
    tag_col = engine.detect_tag_column(tests_df)

    all_test_types = [
        ("Drop Landing", "Drop Landing"),
        ("Drop Jump", "Drop Jump"),
        ("Bodyweight Squats", "Bodyweight Squats|Free Run-Bodyweight|Free Run.*Bodyweight"),
        ("Static Balance", "Balance.*Firm|Balance.*Eyes|Free Run-Balance|Free Run.*Balance"),
        ("CMJ Rebound", "CMJ Rebound|Countermovement.*Rebound"),
        ("Countermovement Jump", "Countermovement Jump"),
        ("Squat Jump", "Squat Jump"),
        ("Multi Rebound", "Multi Rebound"),
        ("Isometric Test", "Isometric Test"),
    ]

    available = []  # list of (display_name, pattern, variant_name, base_name, tag, trial_count)

    for display_name, pattern in all_test_types:
        trials = engine.get_all_tests_by_type(tests_df, pattern)
        if trials is None or trials.empty:
            continue

        if "testType_name" in trials.columns:
            variants = trials["testType_name"].unique()
        else:
            variants = [display_name]

        for variant_name in variants:
            base_name, tag = engine.parse_test_variant(variant_name)
            if tag_col and tag is None:
                tag = engine.get_tag_value(trials.iloc[0], tag_col)

            variant_trials = trials[trials.get("testType_name", display_name) == variant_name] if "testType_name" in trials.columns else trials
            if variant_trials.empty:
                continue

            label = engine.make_display_label(base_name, tag)
            count = len(variant_trials)
            available.append({
                "display_name": display_name,
                "pattern": pattern,
                "variant_name": variant_name,
                "base_name": base_name,
                "tag": tag,
                "label": label,
                "trial_count": count,
                "trials": variant_trials,
            })

    return available, tag_col


def pick_tests_interactive(available_tests):
    """Show available tests and let user pick which to include."""
    print(f"\n{'='*50}")
    print(f"  AVAILABLE TESTS")
    print(f"{'='*50}")
    for i, t in enumerate(available_tests):
        print(f"  {i + 1:3d}. {t['label']}  ({t['trial_count']} trials)")
    print(f"{'='*50}")
    print(f"  a = select ALL")
    print(f"{'='*50}")

    while True:
        choice = input("\nEnter test numbers separated by commas (e.g. 1,3,4) or 'a' for all: ").strip()
        if choice.lower() == 'a':
            print("  Selected: ALL tests")
            return list(range(len(available_tests)))

        try:
            indices = [int(x.strip()) - 1 for x in choice.split(",")]
            if all(0 <= idx < len(available_tests) for idx in indices):
                for idx in indices:
                    print(f"  Selected: {available_tests[idx]['label']}")
                return indices
            else:
                print(f"Please enter numbers between 1 and {len(available_tests)}")
        except ValueError:
            print("Please enter numbers separated by commas, or 'a' for all")


def generate_report_for_athlete(athlete_id, athlete_name, args, engine, interactive=False):
    """Core report generation logic for a single athlete."""
    from html_reporting.payload import build_payload
    from html_reporting.render_html import render_to_file
    from html_reporting.export_pdf import export_pdf

    print(f"\n[1/6] Fetching tests for {athlete_name}...")
    tests_df = engine.get_tests(athlete_id, days_back=args.days_back)
    if tests_df is None or isinstance(tests_df, str) or (hasattr(tests_df, 'empty') and tests_df.empty):
        print(f"ERROR: No tests found in the last {args.days_back} days.")
        return

    # Discover what's available
    available_tests, tag_col = discover_available_tests(tests_df, engine)

    if not available_tests:
        print("ERROR: No matching test types found.")
        return

    # Let user pick tests in interactive mode
    if interactive:
        selected_indices = pick_tests_interactive(available_tests)
        selected_tests = [available_tests[i] for i in selected_indices]
    else:
        # Non-interactive: include all
        selected_tests = available_tests

    found_tests = {}
    all_interpretations = {}

    for t in selected_tests:
        base_name = t["base_name"]
        tag = t["tag"]
        variant_trials = t["trials"]

        # Use latest trial as best
        best_row = variant_trials.iloc[0]

        # Get force-time data
        test_id = best_row.get("id") or best_row.get("testId")
        force_data = None
        if test_id:
            try:
                force_data = engine.get_force_time(test_id)
            except Exception as e:
                print(f"  Warning: Could not fetch force data for {t['label']}: {e}")

        # Session grouping
        sessions = engine.group_trials_by_session(variant_trials)
        for sess in sessions:
            sess["avg_metrics"] = engine.compute_session_averages(sess["trials"], base_name)

        # System weight
        sw = best_row.get("system_weight_n", 800)

        test_key = engine.make_found_test_key(base_name, tag)
        label = engine.make_display_label(base_name, tag)

        found_tests[test_key] = {
            "row": best_row,
            "force_data": force_data,
            "tag": tag,
            "display_label": label,
            "base_name": base_name,
            "sessions": sessions,
        }

        print(f"  Processing: {label}")

        # AI interpretation
        if args.enable_ai:
            print(f"  Generating AI interpretation for {label}...")
            interp = engine.generate_test_interpretation_direct(base_name, best_row, sw)
            if interp:
                all_interpretations[test_key] = interp

    if not found_tests:
        print("ERROR: No matching tests found.")
        return

    # AI dashboard summary
    ai_summary = None
    if args.enable_ai:
        print("  Generating dashboard summary...")
        ai_summary = engine.generate_dashboard_summary_direct(found_tests, athlete_name)

    # Get body weight
    first_row = list(found_tests.values())[0]["row"]
    body_weight_n = first_row.get("system_weight_n") if first_row is not None else None

    print("[2/6] Building payload...")
    payload = build_payload(
        athlete_name=athlete_name,
        found_tests=found_tests,
        body_weight_n=body_weight_n,
        team=args.team,
        sport=args.sport,
        ai_summary=ai_summary,
        test_interpretations=all_interpretations,
    )

    # Generate outputs
    os.makedirs(args.out, exist_ok=True)
    safe_name = athlete_name.replace(" ", "_").replace("/", "-")
    timestamp = datetime.now().strftime("%Y%m%d")
    base_filename = f"Report_{safe_name}_{timestamp}"

    if args.format in ("html", "both"):
        html_path = os.path.join(args.out, f"{base_filename}.html")
        print(f"[3/6] Rendering HTML -> {html_path}")
        render_to_file(payload, html_path, inline_styles=True)

    if args.format in ("pdf", "both"):
        html_for_pdf = os.path.join(args.out, f"{base_filename}_print.html")
        render_to_file(payload, html_for_pdf, inline_styles=True)
        pdf_path = os.path.join(args.out, f"{base_filename}.pdf")
        print(f"[4/6] Exporting PDF -> {pdf_path}")
        export_pdf(html_for_pdf, pdf_path)
        if args.format == "pdf":
            os.unlink(html_for_pdf)

    print(f"\nDone! Report for {athlete_name} generated successfully.")
    print(f"Output: {os.path.abspath(args.out)}")


def run_interactive(args):
    """Interactive mode: connect to API, show athlete list, let user pick."""
    import Individual_Report_Generator_v4 as engine

    print("Connecting to Hawkin API...")
    engine.setup_api()

    while True:
        athlete_id, athlete_name = pick_athlete_interactive(engine)
        generate_report_for_athlete(athlete_id, athlete_name, args, engine, interactive=True)

        again = input("\nGenerate another report? (y/n): ").strip().lower()
        if again != 'y':
            print("Done. Goodbye!")
            break


def run_from_api(args):
    """Direct mode: generate report for a specific athlete ID."""
    import Individual_Report_Generator_v4 as engine

    print("Connecting to Hawkin API...")
    engine.setup_api()

    print("Fetching athlete data...")
    athletes = engine.get_athletes()
    athlete_row = athletes[athletes["id"] == args.athlete_id]
    if athlete_row.empty:
        print(f"ERROR: Athlete ID '{args.athlete_id}' not found.")
        sys.exit(1)

    athlete_name = athlete_row.iloc[0].get("name", "Unknown Athlete")
    generate_report_for_athlete(args.athlete_id, athlete_name, args, engine)


def run_from_json(args):
    """Generate report from a JSON payload file (offline mode)."""
    from html_reporting.payload import load_payload_from_json
    from html_reporting.render_html import render_to_file
    from html_reporting.export_pdf import export_pdf

    print(f"Loading payload from {args.input_json}...")
    payload = load_payload_from_json(args.input_json)

    os.makedirs(args.out, exist_ok=True)
    athlete_name = payload.get("athlete", {}).get("name", "Athlete")
    safe_name = athlete_name.replace(" ", "_").replace("/", "-")
    timestamp = datetime.now().strftime("%Y%m%d")
    base_filename = f"Report_{safe_name}_{timestamp}"

    if args.format in ("html", "both"):
        html_path = os.path.join(args.out, f"{base_filename}.html")
        print(f"Rendering HTML -> {html_path}")
        render_to_file(payload, html_path, inline_styles=True)

    if args.format in ("pdf", "both"):
        html_for_pdf = os.path.join(args.out, f"{base_filename}_print.html")
        render_to_file(payload, html_for_pdf, inline_styles=True)
        pdf_path = os.path.join(args.out, f"{base_filename}.pdf")
        print(f"Exporting PDF -> {pdf_path}")
        export_pdf(html_for_pdf, pdf_path)
        if args.format == "pdf":
            os.unlink(html_for_pdf)

    print("\nDone! Report generated successfully.")


def main():
    args = parse_args()

    if args.input_json:
        run_from_json(args)
    elif args.athlete_id:
        run_from_api(args)
    else:
        # Interactive mode -- no athlete ID provided
        run_interactive(args)


if __name__ == "__main__":
    main()
