"""
smoke_test.py — Verify HTML rendering and PDF export work end-to-end.

Usage:
    python -m html_reporting.smoke_test
"""

import os
import sys
import time

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_DIR = os.path.dirname(_MODULE_DIR)
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

MOCK_PAYLOAD_PATH = os.path.join(_MODULE_DIR, "mock_payload.json")
OUTPUT_DIR = os.path.join(_MODULE_DIR, "assets", "smoke_test_output")


def test_html_rendering():
    """Test 1: Render HTML from mock payload."""
    print("=" * 60)
    print("TEST 1: HTML Rendering")
    print("=" * 60)

    from html_reporting.payload import load_payload_from_json
    from html_reporting.render_html import render_to_file

    payload = load_payload_from_json(MOCK_PAYLOAD_PATH)
    assert payload is not None, "Failed to load mock payload"
    assert "athlete" in payload, "Payload missing 'athlete' key"
    assert "tests" in payload, "Payload missing 'tests' key"
    print(f"  Loaded payload: {payload['athlete']['name']}")
    print(f"  Tests: {len(payload['tests'])}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    html_path = os.path.join(OUTPUT_DIR, "smoke_test.html")
    result = render_to_file(payload, html_path, inline_styles=True)

    assert os.path.exists(result), f"HTML file not created: {result}"
    size = os.path.getsize(result)
    assert size > 1000, f"HTML file too small ({size} bytes)"
    print(f"  HTML created: {result} ({size:,} bytes)")

    # Verify key content
    with open(result, "r", encoding="utf-8") as f:
        html = f.read()
    assert "Alex Johnson" in html, "Athlete name not found in HTML"
    assert "PERFORMANCE DASHBOARD" in html, "Dashboard section not found"
    assert "COUNTERMOVEMENT JUMP" in html, "Test section not found"
    assert "ASYMMETRY SUMMARY" in html, "Asymmetry section not found"
    assert "<style>" in html, "Inline styles not found in HTML"

    print("  Content checks passed.")
    print("  PASS\n")
    return html_path


def test_pdf_export(html_path):
    """Test 2: Export PDF from rendered HTML."""
    print("=" * 60)
    print("TEST 2: PDF Export (Playwright)")
    print("=" * 60)

    try:
        from html_reporting.export_pdf import export_pdf
    except ImportError as e:
        print(f"  SKIP: Playwright not available ({e})")
        print("  Install with: pip install playwright && playwright install chromium")
        return

    pdf_path = os.path.join(OUTPUT_DIR, "smoke_test.pdf")

    start = time.time()
    try:
        result = export_pdf(html_path, pdf_path)
    except Exception as e:
        print(f"  FAIL: PDF export error: {e}")
        print("  Make sure Playwright browsers are installed: playwright install chromium")
        return

    elapsed = time.time() - start

    assert os.path.exists(result), f"PDF file not created: {result}"
    size = os.path.getsize(result)
    assert size > 5000, f"PDF file too small ({size} bytes)"
    print(f"  PDF created: {result} ({size:,} bytes) in {elapsed:.1f}s")

    # Verify it's a valid PDF
    with open(result, "rb") as f:
        header = f.read(5)
    assert header == b"%PDF-", "File does not start with PDF header"

    print("  PDF header valid.")
    print("  PASS\n")


def test_payload_structure():
    """Test 3: Verify payload structure and data integrity."""
    print("=" * 60)
    print("TEST 3: Payload Structure")
    print("=" * 60)

    from html_reporting.payload import load_payload_from_json

    payload = load_payload_from_json(MOCK_PAYLOAD_PATH)

    # Top-level keys
    required_keys = ["athlete", "assessment", "snapshot", "tests", "asymmetry_summary", "footer"]
    for key in required_keys:
        assert key in payload, f"Missing top-level key: {key}"
    print(f"  Top-level keys: {list(payload.keys())}")

    # Athlete
    athlete = payload["athlete"]
    assert athlete["name"], "Athlete name is empty"
    print(f"  Athlete: {athlete['name']}")

    # Tests
    for test in payload["tests"]:
        assert "name" in test, "Test missing 'name'"
        assert "metrics" in test, f"Test '{test['name']}' missing 'metrics'"
        print(f"  Test: {test['name']} ({len(test['metrics'])} metrics, "
              f"{len(test.get('asymmetries', []))} asymmetries)")

    # Asymmetry summary
    for a in payload["asymmetry_summary"]:
        assert a["status"] in ("OK", "MONITOR", "ADDRESS"), f"Invalid status: {a['status']}"

    print("  PASS\n")


def main():
    print("\n" + "=" * 60)
    print("  SMOKE TEST: html_reporting module")
    print("=" * 60 + "\n")

    test_payload_structure()
    html_path = test_html_rendering()
    test_pdf_export(html_path)

    print("=" * 60)
    print("  ALL TESTS COMPLETE")
    print("=" * 60)
    print(f"\nOutput files in: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
