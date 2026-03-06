"""
app.py — Flask web application for interactive report configuration.

Usage:
    python -m html_reporting.webapp.app
"""

import os
import sys
import json
import io
from flask import Flask, render_template, request, jsonify, send_file, abort

# ---------------------------------------------------------------------------
# Fix Windows console encoding BEFORE anything prints Unicode
# ---------------------------------------------------------------------------
if sys.platform == "win32":
    # Force UTF-8 for stdout/stderr so engine print statements with
    # Unicode chars (checkmarks, arrows, etc.) don't crash
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, encoding="utf-8", errors="replace"
            )
    if hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            sys.stderr = io.TextIOWrapper(
                sys.stderr.buffer, encoding="utf-8", errors="replace"
            )

# Ensure project root is on path
_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

_WEBAPP_DIR = os.path.dirname(os.path.abspath(__file__))
_OUTPUT_DIR = os.path.join(_PROJECT_DIR, "output")

app = Flask(
    __name__,
    template_folder=os.path.join(_WEBAPP_DIR, "templates"),
    static_folder=os.path.join(_WEBAPP_DIR, "static"),
)
app.config["SECRET_KEY"] = "mma-report-dev-key"


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    """Athlete picker page."""
    return render_template("index.html")


@app.route("/configure/<athlete_id>")
def configure(athlete_id):
    """Test & trial configuration page."""
    return render_template("configure.html", athlete_id=athlete_id)


@app.route("/preview/<filename>")
def preview(filename):
    """Report preview page."""
    return render_template("preview.html", filename=filename)


# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------

@app.route("/api/athletes")
def api_athletes():
    """JSON list of athletes."""
    from html_reporting.webapp.hawkin_service import get_athletes
    try:
        athletes = get_athletes()
        return jsonify({"athletes": athletes})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/tests/<athlete_id>")
def api_tests(athlete_id):
    """JSON: available tests, sessions, trials for an athlete."""
    from html_reporting.webapp.hawkin_service import get_tests_for_athlete
    days_back = request.args.get("days_back", 365, type=int)
    try:
        data = get_tests_for_athlete(athlete_id, days_back=days_back)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Generate report from configuration JSON."""
    from html_reporting.webapp.hawkin_service import generate_report, save_report_config
    try:
        config = request.get_json()
        if not config:
            return jsonify({"error": "No configuration provided"}), 400

        print("[webapp] Starting report generation...")
        result = generate_report(config)

        if "error" in result:
            print(f"[webapp] Generation error: {result['error']}")
            return jsonify(result), 400

        # Save config for later regeneration with summaries
        save_report_config(result["filename"], config)
        print(f"[webapp] Report generated: {result['filename']}")
        return jsonify(result)
    except Exception as e:
        import traceback
        err_str = traceback.format_exc()
        # Print safely (replace any problematic chars)
        try:
            print(f"[webapp] GENERATE ERROR:\n{err_str}")
        except Exception:
            pass
        return jsonify({"error": str(e)}), 500


@app.route("/api/report-config/<filename>")
def api_report_config(filename):
    """Get the test names/keys for a generated report (for summary editing)."""
    from html_reporting.webapp.hawkin_service import get_report_config
    config = get_report_config(filename)
    if not config:
        return jsonify({"error": "Config not found"}), 404
    # Return test display names and keys for the summary textareas
    test_names = []
    test_keys = []
    for sel in config.get("selected_tests", []):
        base = sel.get("base_name", "")
        tag = sel.get("tag", "")
        label = f"{base} [{tag}]" if tag else base
        test_names.append(label)
        # Reconstruct the test_key like the engine does
        key = f"{base}_{tag}" if tag else base
        test_keys.append(key)
    return jsonify({"test_names": test_names, "test_keys": test_keys})


@app.route("/api/regenerate/<filename>", methods=["POST"])
def api_regenerate(filename):
    """Regenerate a report with user summaries added."""
    from html_reporting.webapp.hawkin_service import get_report_config, generate_report
    config = get_report_config(filename)
    if not config:
        return jsonify({"error": "Original config not found"}), 404
    try:
        data = request.get_json()
        user_summaries = data.get("user_summaries", {})
        # Inject user summaries into the config
        config["dashboard_summary"] = user_summaries.pop("__dashboard__", config.get("dashboard_summary", ""))
        for sel in config.get("selected_tests", []):
            base = sel.get("base_name", "")
            tag = sel.get("tag", "")
            key = f"{base}_{tag}" if tag else base
            if key in user_summaries:
                sel["user_summary"] = user_summaries[key]
        result = generate_report(config)
        if "error" in result:
            return jsonify(result), 400
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/output/<filename>")
def serve_output(filename):
    """Serve generated report files from output directory."""
    filepath = os.path.join(_OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        abort(404)
    if filename.endswith(".pdf"):
        return send_file(filepath, mimetype="application/pdf")
    elif filename.endswith(".html"):
        return send_file(filepath, mimetype="text/html")
    else:
        abort(404)


@app.route("/download/<filename>")
def download_file(filename):
    """Download a generated file."""
    filepath = os.path.join(_OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        abort(404)
    return send_file(filepath, as_attachment=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    print("\n" + "=" * 60)
    print("  MOVE, MEASURE, ANALYZE - Report Generator")
    print("=" * 60)
    print(f"\n  Open your browser to: http://localhost:5000")
    print(f"  Press Ctrl+C to stop\n")
    # use_reloader=False is critical: the reloader forks the process,
    # which breaks Playwright's Chromium subprocess for PDF generation.
    # debug=True still gives nice error pages without the reloader.
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)


if __name__ == "__main__":
    main()
