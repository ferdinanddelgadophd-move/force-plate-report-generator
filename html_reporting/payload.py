"""
payload.py — Build structured report payload from existing compute logic.

Wraps Individual_Report_Generator_v4 functions WITHOUT duplicating computation.
Produces a single dict that the Jinja2 template consumes.
"""

import os
import sys
import json
import base64
import io
from datetime import datetime
from dataclasses import dataclass, field, asdict
from typing import Optional

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ---------------------------------------------------------------------------
# Import existing report engine (parent directory)
# ---------------------------------------------------------------------------
_PARENT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PARENT_DIR not in sys.path:
    sys.path.insert(0, _PARENT_DIR)

import Individual_Report_Generator_v4 as engine

# Re-export constants from engine
COLOR_GOLD = engine.COLOR_GOLD
COLOR_BLACK = engine.COLOR_BLACK
COLOR_WHITE = engine.COLOR_WHITE
COLOR_GREEN = engine.COLOR_GREEN
COLOR_RED = engine.COLOR_RED
ASYM_OK = engine.ASYM_OK
ASYM_CONCERN = engine.ASYM_CONCERN
TEST_DESCRIPTIONS = engine.TEST_DESCRIPTIONS
TREND_METRICS = engine.TREND_METRICS


# ---------------------------------------------------------------------------
# Plot helpers — render matplotlib to base64 PNG
# ---------------------------------------------------------------------------

def _fig_to_base64(fig, dpi=150):
    """Render a matplotlib figure to a base64-encoded PNG string."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none")
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode("utf-8")
    buf.close()
    return b64


def render_force_time_plot(force_data, system_weight, title, is_single_leg=False):
    """Generate force-time + asymmetry plot and return as base64 PNG.

    Reuses engine.create_force_time_plot() for the actual drawing.
    """
    if force_data is None or len(force_data) == 0:
        return None, 0.0

    fig, (ax_main, ax_asym) = plt.subplots(
        2, 1, figsize=(10, 5), gridspec_kw={"height_ratios": [3, 1]},
        facecolor="white",
    )
    fig.subplots_adjust(hspace=0.15)

    mean_asym = engine.create_force_time_plot(
        ax_main, ax_asym, force_data, system_weight, title,
        is_single_leg=is_single_leg,
    )

    b64 = _fig_to_base64(fig)
    plt.close(fig)
    return b64, mean_asym


def render_trend_plot(sessions, test_type):
    """Generate a trend line chart using best trial per session. Returns base64 PNG or None."""
    if not sessions or len(sessions) < 2:
        return None

    # Find trend metrics config
    metrics_config = None
    for key, config in TREND_METRICS.items():
        if key in test_type:
            metrics_config = config
            break
    if metrics_config is None:
        return None

    # Use best_trial_metrics if available (best trial per session), otherwise fall back to avg
    data_key = "best_trial_metrics"
    # Verify at least some sessions have the key
    has_best = any(data_key in s for s in sessions)
    if not has_best:
        data_key = "avg_metrics"
        # Compute session averages if not already done
        for sess in sessions:
            if "avg_metrics" not in sess:
                sess["avg_metrics"] = engine.compute_session_averages(sess["trials"], test_type)

    # Filter to metrics that have data
    available = []
    for display_name, _, unit, _ in metrics_config:
        values = [s[data_key][display_name]["value"]
                  for s in sessions if display_name in s.get(data_key, {})]
        if values:
            available.append((display_name, unit))

    if not available:
        return None

    n_metrics = min(len(available), 4)
    fig, axes = plt.subplots(n_metrics, 1, figsize=(10, 2.5 * n_metrics),
                             facecolor="white", squeeze=False)

    dates = [s["date_str"] for s in sessions]
    x = list(range(len(dates)))

    for i, (metric_name, unit) in enumerate(available[:n_metrics]):
        ax = axes[i, 0]
        vals = []
        valid_x = []
        for j, s in enumerate(sessions):
            m = s.get(data_key, {}).get(metric_name)
            if m:
                vals.append(m["value"])
                valid_x.append(j)

        ax.plot(valid_x, vals, "-o", color=COLOR_GOLD, linewidth=2,
                markersize=6, markerfacecolor=COLOR_BLACK, markeredgecolor=COLOR_GOLD)

        # Trend line if 3+ points
        if len(vals) >= 3:
            z = np.polyfit(valid_x, vals, 1)
            p = np.poly1d(z)
            ax.plot(valid_x, p(valid_x), "--", color="#999999", linewidth=1, alpha=0.6)

        ax.set_ylabel(f"{metric_name} ({unit})" if unit else metric_name,
                       fontsize=9, fontweight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(dates, fontsize=8, rotation=30, ha="right")
        ax.grid(True, alpha=0.3, linestyle="--")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        # Value labels
        for xi, vi in zip(valid_x, vals):
            ax.annotate(f"{vi:.2f}", (xi, vi), textcoords="offset points",
                        xytext=(0, 8), ha="center", fontsize=7, color=COLOR_BLACK)

        # Delta annotation
        if len(vals) >= 2:
            delta = vals[-1] - vals[0]
            pct = (delta / abs(vals[0]) * 100) if vals[0] != 0 else 0
            color = COLOR_GREEN if delta > 0 else (COLOR_RED if delta < 0 else "#999")
            arrow = "\u2191" if delta > 0 else ("\u2193" if delta < 0 else "\u2192")
            ax.text(0.98, 0.92, f"{arrow} {pct:+.1f}%", transform=ax.transAxes,
                    ha="right", va="top", fontsize=9, fontweight="bold", color=color)

    fig.tight_layout()
    b64 = _fig_to_base64(fig)
    plt.close(fig)
    return b64


# ---------------------------------------------------------------------------
# Metric extraction — wraps existing logic, no duplication
# ---------------------------------------------------------------------------

def _safe_get(row, col):
    """Safely extract a scalar from a pandas row."""
    if row is None:
        return None
    try:
        import pandas as pd
        if col in row.index and pd.notna(row[col]):
            return row[col]
    except Exception:
        pass
    return None


def _get_session_date_str(test_data):
    """Extract the latest session date string from test data for pairing checks."""
    sessions = test_data.get("sessions", [])
    if sessions:
        return sessions[-1].get("date_str", "")
    return ""


def _get_delta_for_tile(found_tests, base_name_match, metric_key):
    """Look up a delta value for a snapshot tile.

    Parameters
    ----------
    found_tests : dict
    base_name_match : str
        Substring to match against found_tests base_name.
    metric_key : str
        Key in the delta dict (from TREND_METRICS display name).

    Returns
    -------
    dict or None
        {"value": "+2.1%", "direction": "up"} or None
    """
    for test_key, td in found_tests.items():
        bn = td.get("base_name", "")
        tag = td.get("tag")
        if tag and engine._is_single_leg_tag(tag):
            continue
        if base_name_match in bn:
            delta_dict = td.get("delta")
            if delta_dict and metric_key in delta_dict:
                d = delta_dict[metric_key]
                pct = d.get("delta_pct", 0)
                direction = d.get("direction", "neutral")
                sign = "+" if pct > 0 else ""
                return {"value": f"{sign}{pct:.1f}%", "direction": direction}
    return None


def extract_snapshot_metrics(found_tests, session_pairs=None, body_weight_n=None):
    """Build top-level performance snapshot tiles from found_tests dict.

    Parameters
    ----------
    found_tests : dict
    session_pairs : dict, optional
        Explicit session pairing for derived metrics.
        Keys: "eur" -> {"cmj_key": str, "sj_key": str, "cmj_session": str, "sj_session": str}
              "dsi" -> {"cmj_key": str, "imtp_key": str, "cmj_session": str, "imtp_session": str}
        If None, auto-pairs tests from the same date.
    body_weight_n : float, optional
        System weight in Newtons (for xBW calculations).

    Returns list of dicts: [{label, value, unit, category, delta}, ...]
    """
    session_pairs = session_pairs or {}
    snapshot = []

    # Collect values from bilateral tests
    cmj_height = sj_height = cmj_rsi = cmj_peak_force = None
    mr_rsi = dl_peak_force = dj_rsi = imtp_peak_force = imtp_rel_force = None
    cmj_date = sj_date = imtp_date = ""

    for test_key, test_data in found_tests.items():
        row = test_data.get("row")
        if row is None:
            continue
        base_name = test_data.get("base_name", test_key)
        tag = test_data.get("tag")
        if tag and engine._is_single_leg_tag(tag):
            continue

        date_str = _get_session_date_str(test_data)

        if "Countermovement Jump" in base_name and "Rebound" not in base_name:
            jh = _safe_get(row, "jump_height_imp_mom_cm")
            if jh is None:
                jh_m = _safe_get(row, "jump_height_m")
                if jh_m:
                    jh = jh_m * 100
            cmj_height = jh
            cmj_rsi = _safe_get(row, "mrsi")
            cmj_peak_force = _safe_get(row, "peak_propulsive_force_n") or _safe_get(row, "peak_force_n")
            cmj_date = date_str

        elif "Squat Jump" in base_name:
            jh = _safe_get(row, "jump_height_imp_mom_cm")
            if jh is None:
                jh_m = _safe_get(row, "jump_height_m")
                if jh_m:
                    jh = jh_m * 100
            sj_height = jh
            sj_date = date_str

        elif "Multi Rebound" in base_name:
            mr_rsi = _safe_get(row, "rsi") or _safe_get(row, "avg_rsi")

        elif "Drop Landing" in base_name:
            pf = _safe_get(row, "peak_force_n") or _safe_get(row, "peak_landing_force_n")
            sw = _safe_get(row, "system_weight_n")
            if pf and sw:
                dl_peak_force = pf / sw

        elif "Drop Jump" in base_name:
            dj_rsi = _safe_get(row, "rsi") or _safe_get(row, "peak_rsi")

        elif "Isometric" in base_name:
            imtp_peak_force = _safe_get(row, "peak_force_n")
            imtp_rel_force = _safe_get(row, "peak_relative_force")
            imtp_date = date_str

    # Build tiles (with deltas from previous session)
    if cmj_height:
        snapshot.append({
            "label": "CMJ Height",
            "value": f"{cmj_height:.1f} / {cmj_height / 2.54:.1f}",
            "unit": "cm / in",
            "category": "jump",
            "delta": _get_delta_for_tile(found_tests, "Countermovement Jump", "Jump Height"),
        })

    if cmj_rsi:
        snapshot.append({
            "label": "RSImod",
            "value": f"{cmj_rsi:.2f}",
            "unit": "explosive index",
            "category": "reactive",
            "delta": _get_delta_for_tile(found_tests, "Countermovement Jump", "RSImod"),
        })

    if sj_height:
        snapshot.append({
            "label": "SJ Height",
            "value": f"{sj_height:.1f} / {sj_height / 2.54:.1f}",
            "unit": "cm / in",
            "category": "jump",
            "delta": _get_delta_for_tile(found_tests, "Squat Jump", "Jump Height"),
        })

    # EUR — only if CMJ and SJ share the same session date (or explicitly paired)
    eur_pair = session_pairs.get("eur", {})
    eur_ok = False
    if eur_pair:
        # Explicit pairing from web UI
        eur_ok = bool(eur_pair.get("cmj_session") and eur_pair.get("sj_session"))
    elif cmj_date and sj_date:
        # Auto-pair: same date
        eur_ok = cmj_date == sj_date

    if cmj_height and sj_height and sj_height > 0 and eur_ok:
        eur = cmj_height / sj_height
        eur_status = "good" if eur >= 1.05 else ("ok" if eur >= 0.95 else "low")
        # Gauge: scale 0.80–1.30
        eur_pct = max(0, min(100, (eur - 0.80) / 0.50 * 100))
        if eur >= 1.10:
            eur_interp, eur_color = "Good elastic utilization", "green"
        elif eur >= 1.00:
            eur_interp, eur_color = "Moderate \u2014 monitor", "amber"
        else:
            eur_interp, eur_color = "Poor utilization \u2014 address", "red"
        snapshot.append({
            "label": "EUR (CMJ\u00f7SJ)",
            "value": f"{eur:.2f}",
            "unit": f"elastic energy ({cmj_date})",
            "category": "eur",
            "status": eur_status,
            "gauge": {
                "marker_pct": round(eur_pct, 1),
                "zones": [
                    {"width": 40, "color": "red"},
                    {"width": 20, "color": "amber"},
                    {"width": 40, "color": "green"},
                ],
                "labels": ["< 1.00", "1.00\u20131.10", "> 1.10"],
                "interpretation": eur_interp,
                "interp_color": eur_color,
            },
        })
    elif cmj_height and sj_height and sj_height > 0 and not eur_ok:
        snapshot.append({
            "label": "EUR (CMJ\u00f7SJ)",
            "value": "N/A",
            "unit": f"sessions differ ({cmj_date} vs {sj_date})",
            "category": "eur",
        })

    # DSI — only if CMJ and IMTP share the same session date (or explicitly paired)
    dsi_pair = session_pairs.get("dsi", {})
    dsi_ok = False
    if dsi_pair:
        dsi_ok = bool(dsi_pair.get("cmj_session") and dsi_pair.get("imtp_session"))
    elif cmj_date and imtp_date:
        dsi_ok = cmj_date == imtp_date

    if cmj_peak_force and imtp_peak_force and imtp_peak_force > 0 and dsi_ok:
        dsi = cmj_peak_force / imtp_peak_force
        # Gauge: scale 0.30–1.10
        dsi_pct = max(0, min(100, (dsi - 0.30) / 0.80 * 100))
        if dsi < 0.60:
            dsi_interp, dsi_color = "Speed-strength focus", "amber"
        elif dsi <= 0.80:
            dsi_interp, dsi_color = "Optimal balance", "green"
        else:
            dsi_interp, dsi_color = "Max strength focus", "amber"
        snapshot.append({
            "label": "DSI (CMJ\u00f7IMTP)",
            "value": f"{dsi:.2f}",
            "unit": f"dynamic strength ({cmj_date})",
            "category": "strength",
            "gauge": {
                "marker_pct": round(dsi_pct, 1),
                "zones": [
                    {"width": 37.5, "color": "amber"},
                    {"width": 25, "color": "green"},
                    {"width": 37.5, "color": "amber"},
                ],
                "labels": ["< 0.60", "0.60\u20130.80", "> 0.80"],
                "interpretation": dsi_interp,
                "interp_color": dsi_color,
            },
        })

    if mr_rsi:
        snapshot.append({
            "label": "Reactive RSI",
            "value": f"{mr_rsi:.2f}",
            "unit": "rebound ability",
            "category": "reactive",
            "delta": _get_delta_for_tile(found_tests, "Multi Rebound", "RSI"),
        })

    if dj_rsi:
        snapshot.append({
            "label": "Drop Jump RSI",
            "value": f"{dj_rsi:.2f}",
            "unit": "reactive strength",
            "category": "reactive",
            "delta": _get_delta_for_tile(found_tests, "Drop Jump", "RSI"),
        })

    if dl_peak_force:
        snapshot.append({
            "label": "Landing Force",
            "value": f"{dl_peak_force:.2f}",
            "unit": "\u00d7BW (peak)",
            "category": "landing",
            "delta": _get_delta_for_tile(found_tests, "Drop Landing", "Peak Force"),
        })

    if imtp_peak_force:
        # Compute ×BW if we have body weight
        sw = body_weight_n or 800
        xbw = imtp_peak_force / sw if sw > 0 else 0
        snapshot.append({
            "label": "IMTP Peak Force",
            "value": f"{imtp_peak_force:.0f}N ({xbw:.2f} \u00d7BW)",
            "unit": f"{imtp_rel_force:.1f} N/kg" if imtp_rel_force else "",
            "category": "strength",
            "delta": _get_delta_for_tile(found_tests, "Isometric", "Peak Force"),
        })
    elif imtp_rel_force:
        snapshot.append({
            "label": "IMTP Peak Force",
            "value": f"{imtp_rel_force:.1f}",
            "unit": "N/kg",
            "category": "strength",
            "delta": _get_delta_for_tile(found_tests, "Isometric", "Peak Force"),
        })

    return snapshot


def extract_test_metrics(test_name, test_row):
    """Extract key metrics for a single test. Returns list of metric dicts.

    Reuses the same column-lookup logic from engine.create_test_page
    without duplicating the computation math.
    """
    metrics = []
    if test_row is None:
        return metrics

    import pandas as pd

    def safe_get(col):
        if col in test_row.index and pd.notna(test_row[col]):
            return test_row[col]
        return None

    bw_n = safe_get("system_weight_n")
    bw_kg = bw_n / 9.81 if bw_n else None

    test_info = TEST_DESCRIPTIONS.get(test_name, {})
    defined_metrics = test_info.get("metrics", [])

    # Try the structured TEST_DESCRIPTIONS first
    for display_name, col_name, unit, description in defined_metrics:
        val = safe_get(col_name)
        if val is None:
            continue

        # Apply known conversions
        formatted = val
        if "jump_height_m" in col_name and unit == "cm":
            formatted = val * 100
        elif "contact_time_s" in col_name and unit == "ms":
            formatted = val * 1000
        elif "time_to_stabilization_s" in col_name and unit == "s":
            formatted = val
        elif "depth_m" in col_name and unit == "cm":
            formatted = val * 100

        # BW normalization for force columns
        if unit == "N" and bw_n and "force" in col_name.lower():
            formatted_bw = val / bw_n
            metrics.append({
                "label": display_name,
                "value": f"{formatted_bw:.2f}",
                "unit": "\u00d7BW",
                "raw_value": float(val),
                "description": description,
            })
            continue

        if isinstance(formatted, float):
            if abs(formatted) >= 100:
                display_val = f"{formatted:.0f}"
            elif abs(formatted) >= 10:
                display_val = f"{formatted:.1f}"
            else:
                display_val = f"{formatted:.2f}"
        else:
            display_val = str(formatted)

        metrics.append({
            "label": display_name,
            "value": display_val,
            "unit": unit,
            "raw_value": float(val) if isinstance(val, (int, float, np.floating)) else val,
            "description": description,
        })

    # Fallback: use the inline extraction from create_test_page for special cases
    if not metrics:
        metrics = _extract_metrics_inline(test_name, test_row, bw_n, bw_kg)

    return metrics


def _extract_metrics_inline(test_name, test_row, bw_n, bw_kg):
    """Fallback metric extraction mirroring create_test_page inline logic."""
    import pandas as pd
    metrics = []

    def safe_get(col):
        if col in test_row.index and pd.notna(test_row[col]):
            return test_row[col]
        return None

    if "Drop Landing" in test_name:
        pf = safe_get("peak_force_n") or safe_get("peak_landing_force_n")
        if pf and bw_n:
            metrics.append({"label": "Peak Force", "value": f"{pf/bw_n:.2f}", "unit": "\u00d7BW", "description": "Maximum impact force"})
        aif = safe_get("avg_impact_force_n") or safe_get("avg_landing_force_n")
        if aif and bw_n:
            metrics.append({"label": "Avg Impact", "value": f"{aif/bw_n:.2f}", "unit": "\u00d7BW", "description": "Average landing force"})
        irfd = safe_get("impact_rfd_n_s") or safe_get("landing_rfd_n_s")
        if irfd and bw_kg:
            metrics.append({"label": "Impact RFD", "value": f"{irfd/bw_kg:.0f}", "unit": "N/s/kg", "description": "Rate of force development"})
        tts = safe_get("time_to_stabilization_ms")
        if tts:
            metrics.append({"label": "Time to Stable", "value": f"{tts/1000:.2f}", "unit": "s", "description": "Time to reach stability"})

    elif "Countermovement Jump" in test_name and "Rebound" not in test_name:
        jh = safe_get("jump_height_imp_mom_cm")
        if jh is None:
            jh_m = safe_get("jump_height_m")
            if jh_m:
                jh = jh_m * 100
        if jh:
            metrics.append({"label": "Jump Height", "value": f"{jh:.1f} / {jh/2.54:.1f}", "unit": "cm / in", "description": "How high you jumped"})
        mrsi = safe_get("mrsi")
        if mrsi:
            metrics.append({"label": "RSImod", "value": f"{mrsi:.2f}", "unit": "", "description": "Explosiveness: height \u00f7 time"})
        ppf = safe_get("peak_propulsive_force_n")
        if ppf and bw_n:
            metrics.append({"label": "Peak Prop. Force", "value": f"{ppf/bw_n:.2f}", "unit": "\u00d7BW", "description": "Maximum propulsive force"})

    elif "Squat Jump" in test_name:
        jh = safe_get("jump_height_imp_mom_cm")
        if jh is None:
            jh_m = safe_get("jump_height_m")
            if jh_m:
                jh = jh_m * 100
        if jh:
            metrics.append({"label": "Jump Height", "value": f"{jh:.1f} / {jh/2.54:.1f}", "unit": "cm / in", "description": "How high you jumped"})
        pp = safe_get("peak_relative_propulsive_power_w_kg")
        if pp:
            metrics.append({"label": "Peak Power", "value": f"{pp:.1f}", "unit": "W/kg", "description": "Maximum power output"})

    elif "Multi Rebound" in test_name:
        rsi = safe_get("rsi") or safe_get("avg_rsi")
        if rsi:
            metrics.append({"label": "RSI", "value": f"{rsi:.2f}", "unit": "", "description": "Reactive strength index"})
        avg_h = safe_get("avg_jump_height_m")
        if avg_h:
            metrics.append({"label": "Avg Height", "value": f"{avg_h*100:.1f}", "unit": "cm", "description": "Average height achieved"})
        avg_ct = safe_get("avg_contact_time_s")
        if avg_ct:
            metrics.append({"label": "Avg Contact", "value": f"{avg_ct*1000:.0f}", "unit": "ms", "description": "Average ground time"})

    elif "Drop Jump" in test_name:
        jh = safe_get("jump_height_m") or safe_get("peak_jump_height_m")
        if jh:
            metrics.append({"label": "Jump Height", "value": f"{jh*100:.1f}", "unit": "cm", "description": "Height after the drop"})
        rsi = safe_get("rsi") or safe_get("peak_rsi")
        if rsi:
            metrics.append({"label": "RSI", "value": f"{rsi:.2f}", "unit": "", "description": "Reactive strength index"})
        ct = safe_get("total_contact_time_s")
        if ct:
            metrics.append({"label": "Contact Time", "value": f"{ct*1000:.0f}", "unit": "ms", "description": "Ground contact time"})
        plf = safe_get("peak_landing_force_n")
        if plf and bw_n:
            metrics.append({"label": "Peak Landing", "value": f"{plf/bw_n:.2f}", "unit": "\u00d7BW", "description": "Impact force"})

    elif "Isometric" in test_name:
        pf = safe_get("peak_force_n")
        if pf:
            xbw_str = ""
            if bw_n and bw_n > 0:
                xbw_str = f" ({pf / bw_n:.2f} \u00d7BW)"
            metrics.append({"label": "Peak Force", "value": f"{pf:.0f}{xbw_str}", "unit": "N", "description": "Maximum force produced"})
        rpf = safe_get("peak_relative_force")
        if rpf:
            metrics.append({"label": "Peak Force Rel.", "value": f"{rpf:.1f}", "unit": "N/kg", "description": "Peak force per kg body mass"})
        imp = safe_get("positive_impulse_n_s")
        if imp:
            metrics.append({"label": "Impulse", "value": f"{imp:.1f}", "unit": "N·s", "description": "Total positive impulse"})
        plf = safe_get("peak_left_force_n")
        prf = safe_get("peak_right_force_n")
        if plf and prf:
            asym = ((plf - prf) / max(plf, prf)) * 100 if max(plf, prf) > 0 else 0
            metrics.append({"label": "L/R Asym", "value": f"{asym:.1f}", "unit": "%", "description": "Left vs right peak force difference"})

    return metrics


def _format_session_avg_metrics(session_avg_metrics, base_name, body_weight_n=None):
    """Convert session_avg_metrics dict into the standard metrics list format.

    Parameters
    ----------
    session_avg_metrics : dict
        {display_name: {"value": float, "unit": str, "n_trials": int}}
    base_name : str
    body_weight_n : float, optional
        System weight in Newtons (for xBW calculations on IMTP).

    Returns
    -------
    list[dict]
        [{label, value, unit, description, n_trials}, ...]
    """
    if not session_avg_metrics:
        return []

    test_info = TEST_DESCRIPTIONS.get(base_name, {})
    desc_map = {}
    for entry in test_info.get("metrics", []):
        desc_map[entry[0]] = entry[3]  # display_name -> description

    metrics = []
    n_trials = 1
    for display_name, data in session_avg_metrics.items():
        val = data.get("value", 0)
        unit = data.get("unit", "")
        n_trials = data.get("n_trials", 1)

        if abs(val) >= 100:
            fmt_val = f"{val:.0f}"
        elif abs(val) >= 10:
            fmt_val = f"{val:.1f}"
        else:
            fmt_val = f"{val:.2f}"

        # IMTP Peak Force: append ×BW
        if "Isometric" in base_name and display_name == "Peak Force" and unit == "N" and body_weight_n and body_weight_n > 0:
            xbw = val / body_weight_n
            fmt_val = f"{fmt_val} ({xbw:.2f} \u00d7BW)"

        metrics.append({
            "label": display_name,
            "value": fmt_val,
            "unit": unit,
            "description": desc_map.get(display_name, ""),
            "n_trials": n_trials,
        })

    return metrics


def extract_test_asymmetries(test_row, test_name):
    """Wrap engine.extract_asymmetries and return structured dicts."""
    raw = engine.extract_asymmetries(test_row, test_name)
    result = []
    for entry in raw:
        # entry = [test_name, label, numeric_val, formatted_val, status]
        val = entry[2]
        side = "Left" if val > 0 else "Right"
        result.append({
            "metric": entry[1],
            "value": float(val),
            "percent": entry[3],
            "side": side,
            "status": entry[4],  # OK / MONITOR / ADDRESS
        })
    return result


# ---------------------------------------------------------------------------
# Build full payload
# ---------------------------------------------------------------------------

def build_payload(
    athlete_name,
    found_tests,
    body_weight_n=None,
    team=None,
    sport=None,
    assessment_date=None,
    assessment_type="Force Plate Assessment",
    ai_summary=None,
    test_interpretations=None,
    user_summaries=None,
    session_pairs=None,
    enable_ai=False,
    athlete_name_for_ai=None,
):
    """Construct the complete payload dict consumed by report.html.

    Parameters
    ----------
    athlete_name : str
    found_tests : dict
        Keyed by test_key, values have 'row', 'force_data', 'tag',
        'display_label', 'base_name', 'sessions'.
    body_weight_n : float, optional
        System weight in Newtons.
    team, sport : str, optional
    assessment_date : str or datetime, optional
    assessment_type : str
    ai_summary : str, optional
        AI-generated dashboard summary.
    test_interpretations : dict, optional
        {test_key: interpretation_text}
    user_summaries : dict, optional
        {test_key: user_written_summary_text}
        User-provided summaries override AI interpretations.
    session_pairs : dict, optional
        Explicit session pairing for EUR/DSI.
        See extract_snapshot_metrics for format.
    """
    if assessment_date is None:
        assessment_date = datetime.now()
    if isinstance(assessment_date, str):
        try:
            assessment_date = datetime.strptime(assessment_date, "%Y-%m-%d")
        except ValueError:
            assessment_date = datetime.now()

    bw_kg = body_weight_n / 9.81 if body_weight_n else None
    bw_lb = bw_kg * 2.20462 if bw_kg else None

    test_interpretations = test_interpretations or {}
    user_summaries = user_summaries or {}

    # --- Athlete ---
    athlete = {
        "name": athlete_name,
        "team": team or "",
        "sport": sport or "",
        "body_weight_kg": round(bw_kg, 1) if bw_kg else None,
        "body_weight_lb": round(bw_lb, 1) if bw_lb else None,
    }

    # --- Assessment ---
    assessment = {
        "date": assessment_date.strftime("%Y-%m-%d"),
        "date_display": assessment_date.strftime("%B %d, %Y"),
        "type": assessment_type,
    }

    # --- Performance Snapshot ---
    snapshot = extract_snapshot_metrics(found_tests, session_pairs=session_pairs, body_weight_n=body_weight_n)

    # --- Test ordering ---
    base_test_order = [
        "Bodyweight Squats", "Drop Landing", "Drop Jump", "Static Balance",
        "Countermovement Jump", "CMJ Rebound", "Squat Jump", "Multi Rebound",
        "Isometric Test",
    ]

    def sort_key(key):
        base = found_tests[key].get("base_name", key)
        try:
            return base_test_order.index(base)
        except ValueError:
            return 99

    ordered_keys = sorted(found_tests.keys(), key=sort_key)

    # --- Build test sections ---
    tests = []
    all_asymmetries = []

    for test_key in ordered_keys:
        td = found_tests[test_key]
        row = td.get("row")
        force_data = td.get("force_data")
        tag = td.get("tag")
        base_name = td.get("base_name", test_key)
        display_label = td.get("display_label", test_key)
        sessions = td.get("sessions", [])
        is_single_leg = engine._is_single_leg_tag(tag)

        # System weight from row
        sw = _safe_get(row, "system_weight_n") or body_weight_n or 800

        # Force-time plot
        ft_title = f"{base_name} Force-Time Curve"
        ft_b64, mean_asym = render_force_time_plot(
            force_data, sw, ft_title, is_single_leg=is_single_leg,
        )

        # Key metrics — prefer session averages, fall back to single-trial extraction
        session_avg = td.get("session_avg_metrics", {})
        if session_avg:
            metrics = _format_session_avg_metrics(session_avg, base_name, body_weight_n=sw)
        else:
            metrics = extract_test_metrics(base_name, row)

        # Individual trial table data
        trial_table_data = td.get("trial_table", [])

        # Asymmetries
        asymmetries = []
        if row is not None and not is_single_leg:
            asymmetries = extract_test_asymmetries(row, base_name)
            all_asymmetries.extend([
                {**a, "test": display_label} for a in asymmetries
            ])

        # Test description
        test_info = TEST_DESCRIPTIONS.get(base_name, {})

        # Trend plot
        trend_b64 = render_trend_plot(sessions, base_name)

        # Trend table data — use best trial per session (fall back to avg)
        trend_table = []
        if sessions and len(sessions) >= 2:
            trend_data_key = "best_trial_metrics" if any(
                "best_trial_metrics" in s for s in sessions
            ) else "avg_metrics"
            for sess in sessions:
                src = sess.get(trend_data_key, sess.get("avg_metrics", {}))
                trend_table.append({
                    "date": sess["date_str"],
                    "metrics": {k: {"value": f"{v['value']:.2f}", "unit": v["unit"],
                                    "n": v.get("n_trials", 1)}
                                for k, v in src.items()},
                })

        # User summary takes priority over AI interpretation
        interpretation = user_summaries.get(test_key, "") or test_interpretations.get(test_key, "")

        # Trial selection metadata (from web app or CLI)
        trial_info = td.get("trial_info", {})

        # N trials count for session averages label
        n_trials_in_session = 0
        if trial_table_data:
            n_trials_in_session = len(trial_table_data)
        elif session_avg:
            # Get from first metric's n_trials
            first_val = next(iter(session_avg.values()), {})
            n_trials_in_session = first_val.get("n_trials", 1)

        tests.append({
            "key": test_key,
            "name": base_name,
            "display_label": display_label,
            "tag": tag,
            "is_single_leg": is_single_leg,
            "title": test_info.get("title", base_name),
            "what": test_info.get("what", ""),
            "why": test_info.get("why", ""),
            "force_time_plot": ft_b64,
            "mean_asymmetry": round(mean_asym, 1) if mean_asym else 0,
            "metrics": metrics,
            "metrics_source": "session_avg" if session_avg else "best_trial",
            "n_trials": n_trials_in_session,
            "trial_table": trial_table_data,
            "asymmetries": asymmetries,
            "interpretation": interpretation,
            "has_trend": trend_b64 is not None,
            "trend_plot": trend_b64,
            "trend_table": trend_table,
            "trial_info": {
                "session_date": trial_info.get("session_date", ""),
                "trial_time": trial_info.get("trial_time", ""),
                "selection_reason": trial_info.get("selection_reason", ""),
            },
        })

    # --- Asymmetry AI interpretation (always generated when data exists) ---
    asym_ai_interpretation = None
    if all_asymmetries:
        try:
            asym_ai_interpretation = engine.generate_asymmetry_interpretation_direct(
                all_asymmetries, athlete_name_for_ai or athlete_name
            )
        except Exception as e:
            print(f"  Asymmetry AI error: {e}")

    # --- Footer ---
    footer = {
        "disclaimer": "This report is generated from force plate data and is intended for informational purposes only. It does not constitute medical advice.",
        "data_source": "Hawkin Dynamics",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "brand": "Move, Measure, Analyze",
    }

    # User dashboard summary overrides AI
    dashboard_summary = user_summaries.get("__dashboard__", "") or ai_summary or ""

    return {
        "athlete": athlete,
        "assessment": assessment,
        "snapshot": snapshot,
        "ai_summary": dashboard_summary,
        "tests": tests,
        "asymmetry_summary": all_asymmetries,
        "asym_ai_interpretation": asym_ai_interpretation,
        "asym_ok": ASYM_OK,
        "asym_concern": ASYM_CONCERN,
        "footer": footer,
    }


# ---------------------------------------------------------------------------
# JSON serialization for offline / mock testing
# ---------------------------------------------------------------------------

def payload_to_json(payload, path=None):
    """Serialize payload to JSON, stripping base64 image data for readability."""
    def _strip(obj):
        if isinstance(obj, dict):
            out = {}
            for k, v in obj.items():
                if k in ("force_time_plot", "trend_plot") and isinstance(v, str) and len(v) > 200:
                    out[k] = v[:80] + "...<base64_truncated>"
                else:
                    out[k] = _strip(v)
            return out
        elif isinstance(obj, list):
            return [_strip(i) for i in obj]
        return obj

    data = _strip(payload)
    text = json.dumps(data, indent=2, default=str)
    if path:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
    return text


def load_payload_from_json(path):
    """Load a payload dict from a JSON file (for offline testing)."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
