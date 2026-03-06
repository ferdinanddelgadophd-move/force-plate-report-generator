"""
hawkin_service.py — Thin wrapper around the existing engine for Flask routes.

All computation stays in Individual_Report_Generator_v4.py.
This module just calls engine functions and returns JSON-friendly dicts.
"""

import os
import sys
import re
import json
from datetime import datetime

import numpy as np
import pandas as pd

# Ensure parent dir is on path for engine import
_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

import Individual_Report_Generator_v4 as engine

_api_initialized = False


def ensure_api():
    """Initialize the Hawkin API once per process."""
    global _api_initialized
    if not _api_initialized:
        engine.setup_api()
        _api_initialized = True


def get_athletes():
    """Return list of athletes as dicts [{id, name}, ...]."""
    ensure_api()
    df = engine.get_athletes()
    if df is None or df.empty:
        return []
    df = df.sort_values("name").reset_index(drop=True)
    results = []
    for _, row in df.iterrows():
        results.append({
            "id": row.get("id", ""),
            "name": row.get("name", "Unknown"),
        })
    return results


def get_athlete_name(athlete_id):
    """Get athlete name from ID."""
    ensure_api()
    athletes = engine.get_athletes()
    if athletes is None or athletes.empty:
        return "Unknown Athlete"
    row = athletes[athletes["id"] == athlete_id]
    if row.empty:
        return "Unknown Athlete"
    return row.iloc[0].get("name", "Unknown Athlete")


# All supported test types (same as cli.py / engine)
ALL_TEST_TYPES = [
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


# "Best by" options per test type for user-selectable trial selection
BEST_METRIC_OPTIONS = {
    "Countermovement Jump": [
        {"label": "mRSI", "col": "mrsi", "col_alt": None, "direction": "max", "default": True},
        {"label": "Jump Height", "col": "jump_height_imp_mom_cm", "col_alt": "jump_height_m", "direction": "max"},
        {"label": "Jump Momentum", "col": "jump_momentum_kg_m_s", "col_alt": None, "direction": "max"},
    ],
    "Squat Jump": [
        {"label": "Jump Height", "col": "jump_height_imp_mom_cm", "col_alt": "jump_height_m", "direction": "max", "default": True},
        {"label": "Peak Power", "col": "peak_relative_propulsive_power_w_kg", "col_alt": None, "direction": "max"},
    ],
    "CMJ Rebound": [
        {"label": "RSI", "col": "rebound_rsi", "col_alt": "rsi", "direction": "max", "default": True},
        {"label": "Rebound Height", "col": "rebound_jump_height_m", "col_alt": None, "direction": "max"},
    ],
    "Multi Rebound": [
        {"label": "RSI", "col": "rsi", "col_alt": None, "direction": "max", "default": True},
        {"label": "Avg Height", "col": "avg_jump_height_m", "col_alt": None, "direction": "max"},
        {"label": "Avg RSI", "col": "avg_rsi", "col_alt": None, "direction": "max"},
    ],
    "Drop Jump": [
        {"label": "RSI", "col": "rsi", "col_alt": None, "direction": "max", "default": True},
        {"label": "Jump Height", "col": "jump_height_m", "col_alt": None, "direction": "max"},
    ],
    "Drop Landing": [
        {"label": "Lowest Peak Force", "col": "peak_landing_force_n", "col_alt": "peak_force_n", "direction": "min", "default": True},
        {"label": "Fastest Stabilization", "col": "time_to_stabilization_s", "col_alt": "time_to_stabilization_ms", "direction": "min"},
    ],
    "Isometric Test": [
        {"label": "Peak Force", "col": "peak_force_n", "col_alt": None, "direction": "max", "default": True},
        {"label": "Relative Force", "col": "peak_relative_force", "col_alt": None, "direction": "max"},
    ],
    "Bodyweight Squats": [],
    "Static Balance": [],
}


# ---------------------------------------------------------------------------
# Report config persistence (in-memory for dev, filesystem for durability)
# ---------------------------------------------------------------------------
_CONFIG_DIR = os.path.join(_PROJECT_DIR, "output", ".configs")


def save_report_config(filename, config):
    """Save the generation config so we can regenerate with summaries later."""
    os.makedirs(_CONFIG_DIR, exist_ok=True)
    path = os.path.join(_CONFIG_DIR, f"{filename}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, default=str)


def get_report_config(filename):
    """Load a previously saved generation config."""
    path = os.path.join(_CONFIG_DIR, f"{filename}.json")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _safe_timestamp(row):
    """Extract a human-readable timestamp from a trial row."""
    for col in ["timestamp", "test_time", "created_at", "date"]:
        if col in row.index and pd.notna(row[col]):
            ts = row[col]
            try:
                if isinstance(ts, (int, float, np.integer, np.floating)):
                    ts_val = int(ts)
                    if ts_val > 10_000_000_000:
                        ts_val = ts_val // 1000
                    return datetime.fromtimestamp(ts_val).strftime("%m/%d/%Y %H:%M")
                elif isinstance(ts, str):
                    if ts.isdigit():
                        ts_val = int(ts)
                        if ts_val > 10_000_000_000:
                            ts_val = ts_val // 1000
                        return datetime.fromtimestamp(ts_val).strftime("%m/%d/%Y %H:%M")
                    return ts[:16] if len(ts) > 16 else ts
            except Exception:
                pass
    return "Unknown"


def _trial_metrics_summary(row, base_name):
    """Return a short metric string for a single trial."""
    parts = []
    if "Countermovement" in base_name and "Rebound" not in base_name:
        for col in ["jump_height_imp_mom_cm", "jump_height_m"]:
            if col in row.index and pd.notna(row[col]):
                v = row[col] * 100 if "jump_height_m" == col else row[col]
                parts.append(f"JH: {v:.1f} cm")
                break
        if "mrsi" in row.index and pd.notna(row["mrsi"]):
            parts.append(f"mRSI: {row['mrsi']:.2f}")
    elif "Squat Jump" in base_name:
        for col in ["jump_height_imp_mom_cm", "jump_height_m"]:
            if col in row.index and pd.notna(row[col]):
                v = row[col] * 100 if col == "jump_height_m" else row[col]
                parts.append(f"JH: {v:.1f} cm")
                break
    elif "CMJ Rebound" in base_name:
        for col in ["rebound_jump_height_m", "rebound_jump_height_cm"]:
            if col in row.index and pd.notna(row[col]):
                v = row[col] * 100 if "_m" in col and row[col] < 1 else row[col]
                parts.append(f"Reb JH: {v:.1f} cm")
                break
        for col in ["rebound_rsi", "rsi"]:
            if col in row.index and pd.notna(row[col]):
                parts.append(f"RSI: {row[col]:.2f}")
                break
    elif "Multi Rebound" in base_name:
        if "rsi" in row.index and pd.notna(row["rsi"]):
            parts.append(f"RSI: {row['rsi']:.2f}")
    elif "Drop Landing" in base_name:
        for col in ["peak_landing_force_n", "peak_force_n"]:
            if col in row.index and pd.notna(row[col]):
                parts.append(f"Peak: {row[col]:.0f} N")
                break
    elif "Isometric" in base_name:
        for col in ["peak_vertical_force_n", "peak_force_n"]:
            if col in row.index and pd.notna(row[col]):
                parts.append(f"Peak: {row[col]:.0f} N")
                break
    elif "Bodyweight" in base_name or "Balance" in base_name:
        if "lr_avg_force" in row.index and pd.notna(row["lr_avg_force"]):
            parts.append(f"L|R: {row['lr_avg_force']:.1f}%")
    return " | ".join(parts)


def _select_best_trial(variant_trials, base_name, best_metric_col=None, best_direction=None):
    """Select the best trial from a DataFrame based on test type.

    Non-interactive version of the engine's select_trial() logic.
    Supports user-selectable "best by" metric via best_metric_col/best_direction.

    Parameters
    ----------
    variant_trials : pd.DataFrame
    base_name : str
    best_metric_col : str, optional
        Column to use for selection. Overrides default per-test logic.
    best_direction : str, optional
        "max" or "min". Required when best_metric_col is provided.

    Returns
    -------
    tuple : (best_positional_index, metric_name, metric_value)
    """
    n = len(variant_trials)
    best_idx = n - 1  # Default: most recent trial
    best_metric_name = None
    best_metric_value = None

    # --- User-specified "best by" metric ---
    if best_metric_col:
        # Try primary column, then look up col_alt from BEST_METRIC_OPTIONS
        col_to_use = None
        if best_metric_col in variant_trials.columns:
            col_to_use = best_metric_col
        else:
            # Check for col_alt in BEST_METRIC_OPTIONS
            options = BEST_METRIC_OPTIONS.get(base_name, [])
            for opt in options:
                if opt["col"] == best_metric_col and opt.get("col_alt"):
                    if opt["col_alt"] in variant_trials.columns:
                        col_to_use = opt["col_alt"]
                        break

        if col_to_use:
            valid = variant_trials[col_to_use].dropna()
            if len(valid) > 0:
                direction = best_direction or "max"
                if direction == "min":
                    label_idx = variant_trials[col_to_use].idxmin()
                else:
                    label_idx = variant_trials[col_to_use].idxmax()
                best_idx = variant_trials.index.get_loc(label_idx)
                best_metric_name = col_to_use
                best_metric_value = float(variant_trials.iloc[best_idx][col_to_use])
                return best_idx, best_metric_name, best_metric_value

    # --- Default per-test-type logic (CLI fallback) ---
    if "Countermovement" in base_name and "Rebound" not in base_name:
        if "mrsi" in variant_trials.columns:
            valid = variant_trials["mrsi"].dropna()
            if len(valid) > 0:
                label_idx = variant_trials["mrsi"].idxmax()
                best_idx = variant_trials.index.get_loc(label_idx)
                best_metric_name = "mRSI"
                best_metric_value = float(variant_trials.iloc[best_idx]["mrsi"])

    elif "CMJ Rebound" in base_name or (
        "Countermovement" in base_name and "Rebound" in base_name
    ):
        for col in ["rebound_rsi", "rsi"]:
            if col in variant_trials.columns:
                valid = variant_trials[col].dropna()
                if len(valid) > 0:
                    label_idx = variant_trials[col].idxmax()
                    best_idx = variant_trials.index.get_loc(label_idx)
                    best_metric_name = "RSI"
                    best_metric_value = float(variant_trials.iloc[best_idx][col])
                    break

    elif "Squat Jump" in base_name:
        for col in ["jump_height_imp_mom_cm", "jump_height_m"]:
            if col in variant_trials.columns:
                valid = variant_trials[col].dropna()
                if len(valid) > 0:
                    label_idx = variant_trials[col].idxmax()
                    best_idx = variant_trials.index.get_loc(label_idx)
                    best_metric_name = "Jump Height"
                    val = float(variant_trials.iloc[best_idx][col])
                    best_metric_value = val * 100 if col.endswith("_m") else val
                    break

    elif "Multi Rebound" in base_name:
        if "rsi" in variant_trials.columns:
            valid = variant_trials["rsi"].dropna()
            if len(valid) > 0:
                label_idx = variant_trials["rsi"].idxmax()
                best_idx = variant_trials.index.get_loc(label_idx)
                best_metric_name = "RSI"
                best_metric_value = float(variant_trials.iloc[best_idx]["rsi"])

    elif "Drop Landing" in base_name:
        for col in ["peak_landing_force_n", "peak_force_n"]:
            if col in variant_trials.columns:
                valid = variant_trials[col].dropna()
                if len(valid) > 0:
                    label_idx = variant_trials[col].idxmin()
                    best_idx = variant_trials.index.get_loc(label_idx)
                    best_metric_name = "Peak Landing Force"
                    best_metric_value = float(variant_trials.iloc[best_idx][col])
                    break

    elif "Drop Jump" in base_name:
        if "rsi" in variant_trials.columns:
            valid = variant_trials["rsi"].dropna()
            if len(valid) > 0:
                label_idx = variant_trials["rsi"].idxmax()
                best_idx = variant_trials.index.get_loc(label_idx)
                best_metric_name = "RSI"
                best_metric_value = float(variant_trials.iloc[best_idx]["rsi"])

    elif "Isometric" in base_name:
        if "peak_force_n" in variant_trials.columns:
            valid = variant_trials["peak_force_n"].dropna()
            if len(valid) > 0:
                label_idx = variant_trials["peak_force_n"].idxmax()
                best_idx = variant_trials.index.get_loc(label_idx)
                best_metric_name = "Peak Force"
                best_metric_value = float(variant_trials.iloc[best_idx]["peak_force_n"])

    return best_idx, best_metric_name, best_metric_value


def get_tests_for_athlete(athlete_id, days_back=365):
    """Fetch all test data for an athlete and return structured dict for the UI.

    Returns:
        {
            "athlete_name": str,
            "test_types": [
                {
                    "display_name": str,
                    "base_name": str,
                    "tag": str or None,
                    "label": str,
                    "sessions": [
                        {
                            "date": str,
                            "trial_count": int,
                            "avg_metrics": str,
                            "trials": [
                                {"id": str, "time": str, "metrics": str, "index": int}
                            ]
                        }
                    ]
                }
            ]
        }
    """
    ensure_api()
    tests_df = engine.get_tests(athlete_id, days_back=days_back)
    if tests_df is None or isinstance(tests_df, str) or (hasattr(tests_df, "empty") and tests_df.empty):
        return {"athlete_name": get_athlete_name(athlete_id), "test_types": []}

    tag_col = engine.detect_tag_column(tests_df)
    result_types = []

    for display_name, pattern in ALL_TEST_TYPES:
        trials = engine.get_all_tests_by_type(tests_df, pattern)
        if trials is None or trials.empty:
            continue

        # Group by variant (testType_name)
        if "testType_name" in trials.columns:
            variants = trials["testType_name"].unique()
        else:
            variants = [display_name]

        for variant_name in variants:
            base_name, tag = engine.parse_test_variant(variant_name)
            if tag_col and tag is None:
                tag = engine.get_tag_value(trials.iloc[0], tag_col)

            vt = (
                trials[trials["testType_name"] == variant_name]
                if "testType_name" in trials.columns
                else trials
            )
            if vt.empty:
                continue

            label = engine.make_display_label(base_name, tag)

            # Group into sessions first (needed for latest-session selection)
            sessions_raw = engine.group_trials_by_session(vt)

            # Determine best trial from the LATEST session (not all sessions)
            if sessions_raw:
                latest_session_trials = sessions_raw[-1].get("trials", pd.DataFrame())
                best_pos_idx, best_metric_name, best_metric_value = _select_best_trial(latest_session_trials, base_name)
                best_trial_row = latest_session_trials.iloc[best_pos_idx]
            else:
                best_pos_idx, best_metric_name, best_metric_value = _select_best_trial(vt, base_name)
                best_trial_row = vt.iloc[best_pos_idx]
            best_trial_id = str(
                best_trial_row.get("id") or best_trial_row.get("testId") or f"row_{best_pos_idx}"
            )

            # Build human-readable selection reason
            latest_date = sessions_raw[-1].get("date_str", "") if sessions_raw else ""
            if best_metric_name and best_metric_value is not None:
                if "Drop Landing" in base_name:
                    fmt = f"{best_metric_value:.0f}" if best_metric_value > 10 else f"{best_metric_value:.2f}"
                    best_reason = f"Lowest {best_metric_name}: {fmt} (latest session{' ' + latest_date if latest_date else ''})"
                else:
                    fmt = f"{best_metric_value:.0f}" if best_metric_value > 10 else f"{best_metric_value:.2f}"
                    best_reason = f"Highest {best_metric_name}: {fmt} (latest session{' ' + latest_date if latest_date else ''})"
            else:
                best_reason = "Most recent trial"
            sessions_out = []

            for sess in sessions_raw:
                sess_trials = sess.get("trials", pd.DataFrame())
                if hasattr(sess_trials, "empty") and sess_trials.empty:
                    continue

                avg = engine.compute_session_averages(sess_trials, base_name)
                avg_str = " | ".join(
                    f"{k}: {v['value']:.2f} {v['unit']}" for k, v in avg.items()
                ) if avg else ""

                trial_list = []
                for idx in range(len(sess_trials)):
                    row = sess_trials.iloc[idx]
                    trial_id = row.get("id") or row.get("testId") or f"row_{idx}"
                    trial_id_str = str(trial_id)
                    trial_list.append({
                        "id": trial_id_str,
                        "time": _safe_timestamp(row),
                        "metrics": _trial_metrics_summary(row, base_name),
                        "index": idx,
                        "is_best": trial_id_str == best_trial_id,
                    })

                sessions_out.append({
                    "date": sess.get("date_str", "Unknown"),
                    "trial_count": len(sess_trials),
                    "avg_metrics": avg_str,
                    "trials": trial_list,
                })

            # Build best-metric options filtered to available columns
            raw_options = BEST_METRIC_OPTIONS.get(base_name, [])
            available_options = []
            for opt in raw_options:
                col = opt["col"]
                col_alt = opt.get("col_alt")
                if col in vt.columns or (col_alt and col_alt in vt.columns):
                    available_options.append({
                        "label": opt["label"],
                        "col": col if col in vt.columns else col_alt,
                        "direction": opt["direction"],
                        "default": opt.get("default", False),
                    })

            result_types.append({
                "display_name": display_name,
                "base_name": base_name,
                "tag": tag,
                "label": label,
                "pattern": pattern,
                "variant_name": variant_name,
                "sessions": sessions_out,
                "best_trial_id": best_trial_id,
                "best_reason": best_reason,
                "best_metric_options": available_options,
            })

    return {
        "athlete_name": get_athlete_name(athlete_id),
        "test_types": result_types,
    }


def _extract_trial_trend_values(row, base_name):
    """Extract trend metric values from a single trial row.

    Returns dict in same format as compute_session_averages():
    {display_name: {"value": float, "unit": str, "n_trials": 1}}
    """
    metrics_config = None
    for key, cfg in engine.TREND_METRICS.items():
        if key in base_name:
            metrics_config = cfg
            break
    if metrics_config is None:
        return {}

    result = {}
    sw = row.get("system_weight_n", 800) or 800

    for display_name, col_candidates, unit, converter in metrics_config:
        for col in col_candidates:
            if col in row.index and pd.notna(row[col]):
                val = float(row[col])
                if converter == "bw_normalize":
                    val = val / sw
                elif callable(converter):
                    val = converter(val, col)
                result[display_name] = {"value": val, "unit": unit, "n_trials": 1}
                break
    return result


def _build_trial_table(session_trials_df, base_name, best_row):
    """Build trial table data: list of dicts with 2-3 key metrics per trial.

    Uses TREND_METRICS config to determine which metrics to show.

    Returns
    -------
    list[dict]
        [{trial_num, timestamp, metrics: [{label, value, unit}], is_best}]
    """
    if session_trials_df is None or session_trials_df.empty:
        return []

    metrics_config = None
    for key, cfg in engine.TREND_METRICS.items():
        if key in base_name:
            metrics_config = cfg
            break
    if metrics_config is None:
        return []

    best_id = str(best_row.get("id") or best_row.get("testId") or "")
    result = []

    for i in range(len(session_trials_df)):
        row = session_trials_df.iloc[i]
        trial_id = str(row.get("id") or row.get("testId") or f"row_{i}")
        trial_metrics = []

        sw = row.get("system_weight_n", 800) or 800

        for display_name, col_candidates, unit, converter in metrics_config:
            for col in col_candidates:
                if col in row.index and pd.notna(row[col]):
                    val = float(row[col])
                    if converter == "bw_normalize":
                        val = val / sw
                    elif callable(converter):
                        val = converter(val, col)

                    if abs(val) >= 100:
                        fmt_val = f"{val:.0f}"
                    elif abs(val) >= 10:
                        fmt_val = f"{val:.1f}"
                    else:
                        fmt_val = f"{val:.2f}"

                    trial_metrics.append({
                        "label": display_name,
                        "value": fmt_val,
                        "unit": unit,
                    })
                    break

        result.append({
            "trial_num": i + 1,
            "timestamp": _safe_timestamp(row),
            "metrics": trial_metrics,
            "is_best": trial_id == best_id,
        })

    return result


def _compute_deltas(sessions, best_trial_date, base_name, best_metric_col=None, best_direction=None):
    """Compute deltas between current and previous session's best trial.

    Returns
    -------
    dict or None
        {metric_label: {current: float, previous: float, delta_pct: float, direction: str}}
    """
    if not sessions or len(sessions) < 2:
        return None

    # Find current session index and previous
    current_idx = None
    for i, sess in enumerate(sessions):
        if sess.get("date_str") == best_trial_date:
            current_idx = i
            break

    if current_idx is None:
        # best_trial_date didn't match any session; try the last session
        current_idx = len(sessions) - 1

    if current_idx == 0:
        return None   # no previous session to compare against

    prev_sess = sessions[current_idx - 1]
    curr_sess = sessions[current_idx]

    # Find best trial in each session
    curr_best_idx, _, _ = _select_best_trial(
        curr_sess["trials"], base_name, best_metric_col, best_direction
    )
    prev_best_idx, _, _ = _select_best_trial(
        prev_sess["trials"], base_name, best_metric_col, best_direction
    )

    curr_row = curr_sess["trials"].iloc[curr_best_idx]
    prev_row = prev_sess["trials"].iloc[prev_best_idx]

    # Extract trend metrics from both
    metrics_config = None
    for key, cfg in engine.TREND_METRICS.items():
        if key in base_name:
            metrics_config = cfg
            break
    if metrics_config is None:
        return None

    sw_curr = curr_row.get("system_weight_n", 800) or 800
    sw_prev = prev_row.get("system_weight_n", 800) or 800

    deltas = {}
    for display_name, col_candidates, unit, converter in metrics_config:
        curr_val = None
        prev_val = None

        for col in col_candidates:
            if col in curr_row.index and pd.notna(curr_row[col]):
                v = float(curr_row[col])
                if converter == "bw_normalize":
                    curr_val = v / sw_curr
                elif callable(converter):
                    curr_val = converter(v, col)
                else:
                    curr_val = v
                break

        for col in col_candidates:
            if col in prev_row.index and pd.notna(prev_row[col]):
                v = float(prev_row[col])
                if converter == "bw_normalize":
                    prev_val = v / sw_prev
                elif callable(converter):
                    prev_val = converter(v, col)
                else:
                    prev_val = v
                break

        if curr_val is not None and prev_val is not None and prev_val != 0:
            pct = ((curr_val - prev_val) / abs(prev_val)) * 100
            direction = "up" if pct > 0.5 else ("down" if pct < -0.5 else "neutral")
            deltas[display_name] = {
                "current": curr_val,
                "previous": prev_val,
                "delta_pct": round(pct, 1),
                "direction": direction,
            }

    return deltas if deltas else None


def generate_report(config):
    """Generate a report from the web app configuration.

    config = {
        "athlete_id": str,
        "athlete_name": str,
        "selected_tests": [
            {
                "base_name": str,
                "tag": str or None,
                "pattern": str,
                "variant_name": str,
                "user_summary": str,
                "best_metric_col": str,
                "best_direction": str,
            }
        ],
        "team": str,
        "sport": str,
        "days_back": int,
        "format": "html" | "pdf" | "both",
        "enable_ai": bool,
        "dashboard_summary": str,
    }

    Returns {"html_path": str or None, "pdf_path": str or None, "filename": str}
    """
    from html_reporting.payload import build_payload
    from html_reporting.render_html import render_to_file
    from html_reporting.export_pdf import export_pdf

    ensure_api()

    athlete_id = config["athlete_id"]
    athlete_name = config.get("athlete_name", get_athlete_name(athlete_id))
    days_back = config.get("days_back", 365)
    output_format = config.get("format", "both")
    team = config.get("team", "")
    sport = config.get("sport", "")
    enable_ai = config.get("enable_ai", False)

    # Fetch tests
    tests_df = engine.get_tests(athlete_id, days_back=days_back)
    if tests_df is None or isinstance(tests_df, str) or (hasattr(tests_df, "empty") and tests_df.empty):
        return {"error": f"No tests found for {athlete_name} in the last {days_back} days."}

    tag_col = engine.detect_tag_column(tests_df)

    found_tests = {}
    all_interpretations = {}
    user_summaries = {}

    # Dashboard user summary
    ds = config.get("dashboard_summary", "").strip()
    if ds:
        user_summaries["__dashboard__"] = ds

    for sel in config.get("selected_tests", []):
        base_name = sel["base_name"]
        tag = sel.get("tag")
        pattern = sel["pattern"]
        variant_name = sel.get("variant_name", base_name)
        user_best_metric_col = sel.get("best_metric_col", "") or None
        user_best_direction = sel.get("best_direction", "") or None

        trials = engine.get_all_tests_by_type(tests_df, pattern)
        if trials is None or trials.empty:
            continue

        # Filter to specific variant
        if "testType_name" in trials.columns:
            variant_trials = trials[trials["testType_name"] == variant_name]
            if variant_trials.empty:
                variant_trials = trials
        else:
            variant_trials = trials

        # Session grouping (moved up so we can default to latest session)
        sessions = engine.group_trials_by_session(variant_trials)
        for sess in sessions:
            sess["avg_metrics"] = engine.compute_session_averages(sess["trials"], base_name)
            # Best trial per session for trend data
            bt_idx, _, _ = _select_best_trial(
                sess["trials"], base_name, user_best_metric_col, user_best_direction
            )
            bt_row = sess["trials"].iloc[bt_idx]
            sess["best_trial_metrics"] = _extract_trial_trend_values(bt_row, base_name)

        # --- Smart trial selection ---
        selected_trial_id = sel.get("selected_trial_id")

        if selected_trial_id:
            # User explicitly picked a trial on the configure page
            id_col = "id" if "id" in variant_trials.columns else "testId"
            match = variant_trials[variant_trials[id_col].astype(str) == str(selected_trial_id)]
            if not match.empty:
                best_row = match.iloc[0]
                selection_reason = "Selected trial"
            else:
                # Fallback: best from latest session
                if sessions:
                    latest_trials = sessions[-1]["trials"]
                    best_pos_idx, sel_metric, sel_val = _select_best_trial(
                        latest_trials, base_name, user_best_metric_col, user_best_direction
                    )
                    best_row = latest_trials.iloc[best_pos_idx]
                else:
                    best_pos_idx, sel_metric, sel_val = _select_best_trial(
                        variant_trials, base_name, user_best_metric_col, user_best_direction
                    )
                    best_row = variant_trials.iloc[best_pos_idx]
                if sel_metric:
                    qualifier = "Lowest" if "Drop Landing" in base_name else "Highest"
                    fmt = f"{sel_val:.0f}" if sel_val > 10 else f"{sel_val:.2f}"
                    selection_reason = f"Best in latest session ({qualifier} {sel_metric}: {fmt})"
                else:
                    selection_reason = "Most recent trial"
        else:
            # No user override — best trial from the LATEST session
            if sessions:
                latest_trials = sessions[-1]["trials"]
                best_pos_idx, sel_metric, sel_val = _select_best_trial(
                    latest_trials, base_name, user_best_metric_col, user_best_direction
                )
                best_row = latest_trials.iloc[best_pos_idx]
            else:
                best_pos_idx, sel_metric, sel_val = _select_best_trial(
                    variant_trials, base_name, user_best_metric_col, user_best_direction
                )
                best_row = variant_trials.iloc[best_pos_idx]
            if sel_metric:
                qualifier = "Lowest" if "Drop Landing" in base_name else "Highest"
                fmt = f"{sel_val:.0f}" if sel_val > 10 else f"{sel_val:.2f}"
                selection_reason = f"Best in latest session ({qualifier} {sel_metric}: {fmt})"
            else:
                selection_reason = "Most recent trial"

        # Extract session date and trial time for the selected trial
        selected_session_date = sel.get("selected_session_date", "") or _safe_timestamp(best_row).split(" ")[0]
        selected_trial_time = sel.get("selected_trial_time", "") or _safe_timestamp(best_row)

        # Force-time data
        test_id = best_row.get("id") or best_row.get("testId")
        force_data = None
        if test_id:
            try:
                force_data = engine.get_force_time(test_id)
            except Exception:
                pass

        sw = best_row.get("system_weight_n", 800)
        test_key = engine.make_found_test_key(base_name, tag)
        label = engine.make_display_label(base_name, tag)

        # --- Identify which session the best trial belongs to ---
        best_trial_id_str = str(best_row.get("id") or best_row.get("testId") or "")
        best_trial_session = None
        best_trial_session_date_str = ""
        for sess in sessions:
            sess_trials = sess["trials"]
            id_col = "id" if "id" in sess_trials.columns else "testId"
            if id_col in sess_trials.columns:
                if best_trial_id_str in sess_trials[id_col].astype(str).values:
                    best_trial_session = sess
                    best_trial_session_date_str = sess.get("date_str", "")
                    break
        if best_trial_session is None and sessions:
            # Fallback: use the most recent session
            best_trial_session = sessions[-1]
            best_trial_session_date_str = best_trial_session.get("date_str", "")

        # --- Build trial table for the best trial's session ---
        trial_table = []
        session_avg_metrics = {}
        if best_trial_session is not None:
            trial_table = _build_trial_table(
                best_trial_session["trials"], base_name, best_row
            )
            session_avg_metrics = best_trial_session.get("avg_metrics", {})

        # --- Compute deltas (current best vs previous session's best) ---
        delta = _compute_deltas(
            sessions, best_trial_session_date_str, base_name,
            user_best_metric_col, user_best_direction
        )

        found_tests[test_key] = {
            "row": best_row,
            "force_data": force_data,
            "tag": tag,
            "display_label": label,
            "base_name": base_name,
            "sessions": sessions,
            "trial_info": {
                "session_date": selected_session_date,
                "trial_time": selected_trial_time,
                "selection_reason": selection_reason,
                "trial_id": str(test_id or "unknown"),
            },
            "session_avg_metrics": session_avg_metrics,
            "trial_table": trial_table,
            "delta": delta,
            "best_metric_col": user_best_metric_col or "",
            "best_direction": user_best_direction or "",
        }

        # User summary for this test
        us = sel.get("user_summary", "").strip()
        if us:
            user_summaries[test_key] = us

        # Note: AI interpretation removed from individual test pages per user request.
        # Asymmetry AI is generated separately in payload.py.

    if not found_tests:
        return {"error": "No matching tests found with the selected configuration."}

    # Note: AI dashboard summary removed per user request.
    # Dashboard summary only shows user-written text via the "Add Summaries" panel.
    ai_summary = None

    # Body weight
    first_row = list(found_tests.values())[0]["row"]
    body_weight_n = first_row.get("system_weight_n") if first_row is not None else None

    # Session pairing for derived metrics (EUR, DSI)
    session_pairs = config.get("session_pairs", {})

    payload = build_payload(
        athlete_name=athlete_name,
        found_tests=found_tests,
        body_weight_n=body_weight_n,
        team=team,
        sport=sport,
        ai_summary=ai_summary,
        test_interpretations=all_interpretations,
        user_summaries=user_summaries,
        session_pairs=session_pairs,
        enable_ai=enable_ai,
        athlete_name_for_ai=athlete_name,
    )

    # Output
    out_dir = os.path.join(_PROJECT_DIR, "output")
    os.makedirs(out_dir, exist_ok=True)
    safe_name = athlete_name.replace(" ", "_").replace("/", "-")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"Report_{safe_name}_{timestamp}"

    result = {"filename": base_filename, "html_path": None, "pdf_path": None}

    if output_format in ("html", "both"):
        html_path = os.path.join(out_dir, f"{base_filename}.html")
        render_to_file(payload, html_path, inline_styles=True)
        result["html_path"] = html_path

    if output_format in ("pdf", "both"):
        html_for_pdf = os.path.join(out_dir, f"{base_filename}_print.html")
        render_to_file(payload, html_for_pdf, inline_styles=True)
        pdf_path = os.path.join(out_dir, f"{base_filename}.pdf")
        export_pdf(html_for_pdf, pdf_path)
        result["pdf_path"] = pdf_path
        # Clean up intermediate file
        try:
            os.unlink(html_for_pdf)
        except Exception:
            pass

    return result
