"""
Hawkin Dynamics API Report - Base Template
==========================================
Clean template that generates:
- Force-time graph + asymmetry subplot for each test type
- Asymmetry summary table as the final page

No normative comparisons - just the data visualization.
This is the foundation to build upon.

Move, Measure, Analyze LLC
"""

import os
import sys
import glob
import textwrap
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.collections import LineCollection
from matplotlib.patches import Ellipse
from datetime import datetime, timedelta

# Try to import PyMuPDF for editable form fields
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False
    print("Note: PyMuPDF not installed. Editable form fields will not be available.")
    print("Install with: pip install pymupdf")


def add_editable_text_fields(pdf_path, output_path=None, test_page_indices=None):
    """
    Add editable text fields to the PDF for manual entry of interpretations.
    Fields sit below the gold header bars so headers remain visible.
    
    Args:
        test_page_indices: Optional set of 0-based page indices that are test pages
                          (have ABOUT THIS TEST panel). If None, applies to all pages 2..N-2.
    """
    if not PYMUPDF_AVAILABLE:
        print("PyMuPDF not available - cannot add editable fields")
        return None
    
    if output_path is None:
        output_path = pdf_path
    
    doc = fitz.open(pdf_path)
    
    # Page dimensions: 8.5 x 11 inches = 612 x 792 points
    # In fitz, y=0 is at TOP of page
    
    # Page 2 (index 1) is the Performance Dashboard
    if len(doc) >= 2:
        page = doc[1]
        
        # Start below the "ASSESSMENT SUMMARY" gold header bar
        summary_rect = fitz.Rect(50, 365, 560, 645)
        
        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_name = "dashboard_summary"
        widget.rect = summary_rect
        widget.text_maxlen = 2000
        widget.field_flags = fitz.PDF_TX_FIELD_IS_MULTILINE
        widget.text_fontsize = 10
        widget.fill_color = (1, 1, 1)
        widget.text_color = (0.2, 0.2, 0.2)
        
        page.add_widget(widget)
    
    # Test pages only (skip trend, COP, stabilogram, asymmetry pages)
    if test_page_indices is not None:
        target_pages = sorted(test_page_indices)
    else:
        # Fallback: pages 3 through second-to-last (old behavior)
        target_pages = list(range(2, len(doc) - 1))
    
    for page_idx in target_pages:
        if page_idx >= len(doc):
            continue
        page = doc[page_idx]
        
        # Start below the "ABOUT THIS TEST" gold header bar
        about_rect = fitz.Rect(38, 500, 304, 726)
        
        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_name = f"test_summary_page{page_idx+1}"
        widget.rect = about_rect
        widget.text_maxlen = 1500
        widget.field_flags = fitz.PDF_TX_FIELD_IS_MULTILINE
        widget.text_fontsize = 9
        # No fill_color = transparent background, so the static 
        # What/Why text shows through as guide text underneath
        widget.text_color = (0.2, 0.2, 0.2)
        
        page.add_widget(widget)
    
    # Save to a temp file first, then replace original
    import shutil
    
    temp_file = pdf_path.replace('.pdf', '_temp.pdf')
    doc.save(temp_file)
    doc.close()
    
    shutil.move(temp_file, output_path)
    
    print(f"Added editable form fields to: {output_path}")
    return output_path


# ==============================================================================
# AI INTERPRETATION GENERATION
# ==============================================================================

def generate_test_interpretation(test_name, metrics_dict, force_data_summary=None):
    """
    Generate AI interpretation for a specific test using Claude API.
    
    Args:
        test_name: Name of the test (e.g., "Countermovement Jump")
        metrics_dict: Dictionary of metric names and values
        force_data_summary: Optional summary of force-time characteristics
    
    Returns:
        String with AI-generated interpretation, or None if API unavailable
    """
    try:
        from anthropic import Anthropic
    except ImportError:
        print("  Note: anthropic package not installed. Skipping AI interpretation.")
        return None
    
    # Get API key
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        # Try to import from rag_research if available
        try:
            import sys
            sys.path.insert(0, r"E:\Report Generators\AI")
            from rag_research import ANTHROPIC_API_KEY
            api_key = ANTHROPIC_API_KEY
        except:
            pass
    
    if not api_key:
        return None
    
    client = Anthropic(api_key=api_key)
    
    # Format metrics for the prompt
    metrics_text = "\n".join([f"- {k}: {v}" for k, v in metrics_dict.items()])
    
    prompt = f"""You are a sports science expert interpreting force plate assessment results. 
Write a brief, clear interpretation (3-5 sentences) of these {test_name} results for an athlete or coach.

TEST: {test_name}

METRICS:
{metrics_text}

{f"FORCE-TIME OBSERVATIONS: {force_data_summary}" if force_data_summary else ""}

Write in second person ("Your results show..."). Explain what the metrics mean practically, 
highlight any concerns (asymmetries >10%, unusual values), and give one actionable insight.
Keep it concise and avoid jargon. Do not use bullet points."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        print(f"  AI interpretation error: {e}")
        return None


def generate_dashboard_summary(all_metrics, athlete_name=None):
    """
    Generate AI summary for the Performance Dashboard.
    
    Args:
        all_metrics: Dictionary with test names as keys, metric dicts as values
        athlete_name: Optional athlete name for personalization
    
    Returns:
        String with AI-generated summary, or None if API unavailable
    """
    try:
        from anthropic import Anthropic
    except ImportError:
        return None
    
    # Get API key
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        try:
            import sys
            sys.path.insert(0, r"E:\Report Generators\AI")
            from rag_research import ANTHROPIC_API_KEY
            api_key = ANTHROPIC_API_KEY
        except:
            pass
    
    if not api_key:
        return None
    
    client = Anthropic(api_key=api_key)
    
    # Format all metrics
    metrics_text = ""
    for test_name, metrics in all_metrics.items():
        metrics_text += f"\n{test_name}:\n"
        for k, v in metrics.items():
            metrics_text += f"  - {k}: {v}\n"
    
    prompt = f"""You are a sports science expert writing an executive summary of a force plate assessment.
Write a clear, actionable summary (5-7 sentences) of the overall results.

{"ATHLETE: " + athlete_name if athlete_name else ""}

ALL TEST RESULTS:
{metrics_text}

Write in second person. Cover:
1. Overall performance level (jump height, power)
2. Elastic energy utilization (EUR = CMJ÷SJ, values >1.0 are good)
3. Any significant asymmetries (>10% is concerning, >15% needs attention)
4. One or two key recommendations

Be encouraging but honest. Keep it practical for coaches/athletes."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        print(f"  Dashboard AI summary error: {e}")
        return None


def extract_metrics_for_ai(test_name, test_row, body_weight_n=None):
    """
    Extract key metrics from a test row for AI interpretation.
    
    Args:
        test_name: Name of the test
        test_row: Pandas Series with test data
        body_weight_n: Body weight in Newtons for normalization
    
    Returns:
        Dictionary of metric names and formatted values
    """
    metrics = {}
    
    def safe_get(col):
        if col in test_row.index and pd.notna(test_row[col]):
            return test_row[col]
        return None
    
    bw = body_weight_n or safe_get('system_weight_n') or 800
    bw_kg = bw / 9.81
    
    if 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics['Jump Height'] = f"{jh:.1f} cm ({jh/2.54:.1f} inches)"
        
        mrsi = safe_get('mrsi')
        if mrsi:
            metrics['RSImod'] = f"{mrsi:.2f}"
        
        pp = safe_get('peak_propulsive_force_n')
        if pp:
            metrics['Peak Propulsive Force'] = f"{pp/bw:.2f} ×BW"
        
        asym = safe_get('lr_propulsive_impulse_index') or safe_get('cmj_lr_propulsive_impulse_index')
        if asym:
            side = 'Left' if asym > 0 else 'Right'
            metrics['Propulsive Impulse Asymmetry'] = f"{abs(asym):.1f}% {side}"
    
    elif 'Squat Jump' in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics['Jump Height'] = f"{jh:.1f} cm ({jh/2.54:.1f} inches)"
        
        pp = safe_get('peak_power_w')
        if pp:
            metrics['Peak Power'] = f"{pp/bw_kg:.1f} W/kg"
        
        asym = safe_get('lr_propulsive_impulse_index') or safe_get('sj_lr_propulsive_impulse_index')
        if asym:
            side = 'Left' if asym > 0 else 'Right'
            metrics['Impulse Asymmetry'] = f"{abs(asym):.1f}% {side}"
    
    elif 'Multi Rebound' in test_name:
        rsi = safe_get('rsi') or safe_get('avg_rsi')
        if rsi:
            metrics['Reactive Strength Index'] = f"{rsi:.2f}"
        
        avg_h = safe_get('avg_jump_height_cm')
        if avg_h:
            metrics['Avg Jump Height'] = f"{avg_h:.1f} cm"
        
        avg_ct = safe_get('avg_contact_time_ms')
        if avg_ct:
            metrics['Avg Contact Time'] = f"{avg_ct:.0f} ms"
        
        asym = safe_get('lr_peak_force')
        if asym:
            side = 'Left' if asym > 0 else 'Right'
            metrics['Peak Force Asymmetry'] = f"{abs(asym):.1f}% {side}"
    
    elif 'Drop Landing' in test_name:
        pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
        if pf:
            metrics['Peak Landing Force'] = f"{pf/bw:.2f} ×BW"
        
        tts = safe_get('time_to_stabilization_ms')
        if tts:
            metrics['Time to Stabilization'] = f"{tts/1000:.2f} seconds"
        
        asym = safe_get('lr_peak_force') or safe_get('lr_peak_landing_force')
        if asym:
            side = 'Left' if asym > 0 else 'Right'
            metrics['Peak Force Asymmetry'] = f"{abs(asym):.1f}% {side}"
    
    elif 'Bodyweight Squats' in test_name:
        avg_asym = safe_get('lr_avg_force')
        if avg_asym:
            side = 'Left' if avg_asym > 0 else 'Right'
            metrics['Average Force Asymmetry'] = f"{abs(avg_asym):.1f}% {side}"
        
        peak_asym = safe_get('lr_peak_force')
        if peak_asym:
            side = 'Left' if peak_asym > 0 else 'Right'
            metrics['Peak Force Asymmetry'] = f"{abs(peak_asym):.1f}% {side}"
    
    return metrics

# ==============================================================================
# CONFIGURATION
# ==============================================================================
API_TOKEN = os.environ.get("HAWKIN_API_TOKEN", "")  # Set via environment variable

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, "Reports")

# RAG System path (for AI interpretations)
RAG_SCRIPT_PATH = r"E:\Report Generators\AI\rag_research.py"

# Anthropic API Key (from your rag_research.py)
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


# ==============================================================================
# DIRECT CLAUDE API FOR AI SUMMARIES (No RAG Required)
# ==============================================================================

def get_anthropic_client():
    """Get Anthropic client for direct API calls."""
    try:
        from anthropic import Anthropic
        api_key = ANTHROPIC_API_KEY or os.environ.get('ANTHROPIC_API_KEY')
        if api_key:
            return Anthropic(api_key=api_key)
    except ImportError:
        print("  Note: anthropic package not installed. Run: pip install anthropic")
    return None


def generate_dashboard_summary_direct(found_tests, athlete_name=None):
    """
    Generate AI summary for Dashboard using Claude directly (no RAG).
    """
    client = get_anthropic_client()
    if not client:
        return None
    
    # Extract metrics
    metrics_lines = []
    cmj_height = sj_height = None
    
    for test_name, test_data in found_tests.items():
        row = test_data.get('row')
        if row is None:
            continue
        
        def safe_get(col):
            if col in row.index and pd.notna(row[col]):
                return row[col]
            return None
        
        bw_n = safe_get('system_weight_n') or 800
        bw_kg = bw_n / 9.81
        
        if 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                cmj_height = jh
                metrics_lines.append(f"CMJ Height: {jh:.1f} cm ({jh/2.54:.1f} in)")
            mrsi = safe_get('mrsi')
            if mrsi:
                metrics_lines.append(f"RSImod: {mrsi:.2f}")
            ppf = safe_get('peak_propulsive_force_n')
            if ppf:
                metrics_lines.append(f"CMJ Peak Propulsive Force: {ppf/bw_n:.2f} x BW")
            prop_asym = safe_get('lr_propulsive_impulse_index') or safe_get('cmj_lr_propulsive_impulse_index')
            if prop_asym:
                metrics_lines.append(f"CMJ Propulsive Impulse Asymmetry: {prop_asym:.1f}%")
        
        elif 'CMJ Rebound' in test_name:
            cmj_jh = safe_get('cmj_jump_height_m')
            if cmj_jh:
                metrics_lines.append(f"CMJ Rebound - CMJ Height: {cmj_jh*100:.1f} cm")
            reb_jh = safe_get('rebound_jump_height_m')
            if reb_jh:
                metrics_lines.append(f"CMJ Rebound - Rebound Height: {reb_jh*100:.1f} cm")
            rsi = safe_get('rsi') or safe_get('rebound_rsi')
            if rsi:
                metrics_lines.append(f"CMJ Rebound RSI: {rsi:.2f}")
        
        elif 'Squat Jump' in test_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                sj_height = jh
                metrics_lines.append(f"SJ Height: {jh:.1f} cm ({jh/2.54:.1f} in)")
            pp = safe_get('peak_relative_propulsive_power_w_kg')
            if pp:
                metrics_lines.append(f"SJ Peak Power: {pp:.1f} W/kg")
            imp_asym = safe_get('sj_lr_propulsive_impulse_index') or safe_get('lr_propulsive_impulse_index')
            if imp_asym:
                metrics_lines.append(f"SJ Impulse Asymmetry: {imp_asym:.1f}%")
        
        elif 'Multi Rebound' in test_name:
            rsi = safe_get('rsi') or safe_get('avg_rsi')
            if rsi:
                metrics_lines.append(f"Reactive RSI (10-5): {rsi:.2f}")
            avg_h = safe_get('avg_jump_height_m')
            if avg_h:
                metrics_lines.append(f"MR Avg Height: {avg_h*100:.1f} cm")
            avg_ct = safe_get('avg_contact_time_s')
            if avg_ct:
                metrics_lines.append(f"MR Avg Contact Time: {avg_ct*1000:.0f} ms")
            pf_asym = safe_get('lr_peak_force')
            if pf_asym:
                metrics_lines.append(f"MR Peak Force Asymmetry: {pf_asym:.1f}%")
        
        elif 'Drop Landing' in test_name:
            pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
            if pf:
                metrics_lines.append(f"Landing Peak Force: {pf/bw_n:.2f} x BW")
            aif = safe_get('avg_impact_force_n') or safe_get('avg_landing_force_n')
            if aif:
                metrics_lines.append(f"Landing Avg Impact: {aif/bw_n:.2f} x BW")
            irfd = safe_get('impact_rfd_n_s') or safe_get('landing_rfd_n_s')
            if irfd:
                metrics_lines.append(f"Impact RFD: {irfd/bw_kg:.0f} N/s/kg")
            tts = safe_get('time_to_stabilization_ms')
            if tts:
                metrics_lines.append(f"Time to Stabilization: {tts:.0f} ms")
            pf_asym = safe_get('lr_peak_force') or safe_get('lr_peak_landing_force')
            if pf_asym:
                metrics_lines.append(f"Landing Peak Force Asymmetry: {pf_asym:.1f}%")
        
        elif 'Bodyweight Squats' in test_name:
            avg_asym = safe_get('lr_avg_force')
            if avg_asym:
                metrics_lines.append(f"BW Squat Avg Force Asymmetry: {avg_asym:.1f}%")
            peak_asym = safe_get('lr_peak_force')
            if peak_asym:
                metrics_lines.append(f"BW Squat Peak Force Asymmetry: {peak_asym:.1f}%")
    
    # Calculate EUR
    if cmj_height and sj_height and sj_height > 0:
        eur = cmj_height / sj_height
        metrics_lines.append(f"EUR (CMJ÷SJ): {eur:.2f}")
    
    if not metrics_lines:
        return None
    
    prompt = f"""You are an expert sports scientist interpreting force plate assessment results. Write a brief executive summary (4-6 sentences) for this athlete's assessment.

ATHLETE: {athlete_name or "Athlete"}

RESULTS:
{chr(10).join(metrics_lines)}

CRITICAL INTERPRETATION GUIDELINES:
- EUR (Eccentric Utilization Ratio = CMJ height ÷ SJ height):
  * 1.0-1.1 = typical, athlete uses similar power in both
  * >1.1 = good stretch-shortening cycle / elastic energy utilization
  * <0.95 = limited elastic energy use, needs plyometric training
  * >1.3 = CONCERNING - indicates very poor concentric-only strength relative to SSC ability. Athlete needs dedicated strength training (squats, deadlifts).
  * >2.0 = VERY PROBLEMATIC - huge discrepancy, either a bad SJ trial or severe concentric weakness. Priority is strength development.
- RSImod (CMJ): Males: <0.34 = low, 0.35-0.48 = average, >0.49 = high. Females: <0.24 = low, 0.25-0.37 = average, >0.38 = high (Sole et al. 2018)
- Reactive RSI (10-5 repeat jump): Sprinters 1.8-2.5, Gymnastics 1.75-2.2, Hockey 1.35-1.9, Football 1.3-1.9, Rec Field Sports 0.75-1.7
- Asymmetries: <10% = OK, 10-15% = monitor, >15% = address
- Impact RFD: higher values mean faster force application at landing, important for injury risk
- Time to Stabilization: <1000 ms is good, >1500 ms is concerning

Write in second person. Discuss EVERY metric listed. Be encouraging but honest. End with 1-2 actionable recommendations.
Do NOT use bullet points - write in flowing paragraphs only."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        print(f"  Claude API error: {e}")
        return None


def generate_asymmetry_interpretation_direct(asymmetry_data, athlete_name=None):
    """
    Generate AI interpretation specifically for asymmetry data.

    Parameters
    ----------
    asymmetry_data : list[dict]
        Each dict has: test, metric, value, percent, side, status
    athlete_name : str, optional

    Returns
    -------
    str or None
    """
    client = get_anthropic_client()
    if not client or not asymmetry_data:
        return None

    lines = []
    for a in asymmetry_data:
        lines.append(f"- {a['test']} — {a['metric']}: {a['percent']} ({a['side']} dominant) → {a['status']}")

    prompt = f"""You are a sports performance coach interpreting bilateral asymmetry data from a force plate assessment.

ATHLETE: {athlete_name or "Athlete"}

ASYMMETRY DATA:
{chr(10).join(lines)}

GUIDELINES:
- OK = under {ASYM_OK}% difference between left and right — normal range
- MONITOR = {ASYM_OK}–{ASYM_CONCERN}% — worth keeping an eye on
- ADDRESS = over {ASYM_CONCERN}% — meaningful imbalance that should be worked on

Write a brief interpretation (3-5 sentences) in plain language for a coach or athlete.
- Explain which side is doing more work and in what movements
- If multiple asymmetries favor the same side, note the pattern
- If everything is OK, say so clearly
- Give 1-2 practical suggestions (e.g., single-leg exercises, mobility work)
- Write in second person ("Your results show...")
- Do NOT use bullet points — write in flowing sentences
- Avoid science jargon — keep it simple and actionable"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        print(f"  Asymmetry AI error: {e}")
        return None


def generate_test_interpretation_direct(test_name, test_row, system_weight=None):
    """
    Generate AI interpretation for a specific test using Claude directly (no RAG).
    """
    client = get_anthropic_client()
    if not client or test_row is None:
        return None
    
    def safe_get(col):
        if col in test_row.index and pd.notna(test_row[col]):
            return test_row[col]
        return None
    
    bw = system_weight or safe_get('system_weight_n') or 800
    bw_kg = bw / 9.81
    metrics_lines = []
    benchmarks = ""
    
    # Extract ALL metrics based on test type (matching KEY METRICS panel)
    if 'Drop Landing' in test_name:
        pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
        pf_asym = safe_get('lr_peak_force') or safe_get('lr_peak_landing_force')
        if pf:
            metrics_lines.append(f"Peak Landing Force: {pf/bw:.2f} x BW")
        if pf_asym:
            side = 'Left' if pf_asym > 0 else 'Right'
            metrics_lines.append(f"Peak Force Asymmetry: {abs(pf_asym):.1f}% {side}")
        aif = safe_get('avg_impact_force_n') or safe_get('avg_landing_force_n')
        aif_asym = safe_get('lr_avg_impact_force') or safe_get('lr_avg_landing_force')
        if aif:
            metrics_lines.append(f"Avg Impact Force: {aif/bw:.2f} x BW")
        if aif_asym:
            side = 'Left' if aif_asym > 0 else 'Right'
            metrics_lines.append(f"Avg Impact Force Asymmetry: {abs(aif_asym):.1f}% {side}")
        irfd = safe_get('impact_rfd_n_s') or safe_get('landing_rfd_n_s')
        irfd_asym = safe_get('lr_impact_rfd') or safe_get('lr_landing_rfd')
        if irfd:
            metrics_lines.append(f"Impact RFD: {irfd/bw_kg:.0f} N/s/kg")
        if irfd_asym:
            side = 'Left' if irfd_asym > 0 else 'Right'
            metrics_lines.append(f"Impact RFD Asymmetry: {abs(irfd_asym):.1f}% {side}")
        tts = safe_get('time_to_stabilization_ms')
        if tts:
            metrics_lines.append(f"Time to Stabilization: {tts:.0f} ms")
        benchmarks = """Key benchmarks:
- Peak landing force: <2.5 x BW is moderate, >3.0 x BW is high
- Impact RFD: rate of force development during landing - higher = faster impact, potential injury concern
- Time to stabilization: <1000 ms is good, >1500 ms is concerning
- Asymmetries >10% warrant monitoring, >15% need attention"""
    
    elif 'CMJ Rebound' in test_name:
        cmj_jh = safe_get('cmj_jump_height_m')
        if cmj_jh:
            metrics_lines.append(f"CMJ Height: {cmj_jh*100:.1f} cm ({cmj_jh*100/2.54:.1f} in)")
        reb_jh = safe_get('rebound_jump_height_m')
        if reb_jh:
            metrics_lines.append(f"Rebound Height: {reb_jh*100:.1f} cm ({reb_jh*100/2.54:.1f} in)")
        rsi = safe_get('rsi') or safe_get('rebound_rsi')
        if rsi:
            metrics_lines.append(f"RSI: {rsi:.2f}")
        benchmarks = """Key benchmarks (RSI from 10-5 repeat jump protocol, similar reactive demands):
- Sprinters: 1.8-2.5
- Gymnastics: 1.75-2.2
- Hockey: 1.35-1.9
- Football: 1.3-1.9
- Sub Elite Field Sports: 1.1-1.65
- Recreation Field Sports: 0.75-1.7
Also compare rebound height to CMJ height - a higher rebound/CMJ ratio indicates good reactive ability."""
    
    elif 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics_lines.append(f"Jump Height: {jh:.1f} cm ({jh/2.54:.1f} in)")
        mrsi = safe_get('mrsi')
        if mrsi:
            metrics_lines.append(f"RSImod: {mrsi:.2f}")
        ppf = safe_get('peak_propulsive_force_n')
        ppf_asym = safe_get('cmj_lr_peak_propulsive_force') or safe_get('lr_peak_propulsive_force')
        if ppf:
            metrics_lines.append(f"Peak Propulsive Force: {ppf/bw:.2f} x BW")
        if ppf_asym:
            side = 'Left' if ppf_asym > 0 else 'Right'
            metrics_lines.append(f"Peak Propulsive Force Asymmetry: {abs(ppf_asym):.1f}% {side}")
        prop_asym = safe_get('lr_propulsive_impulse_index') or safe_get('cmj_lr_propulsive_impulse_index')
        if prop_asym:
            side = 'Left' if prop_asym > 0 else 'Right'
            metrics_lines.append(f"Propulsive Impulse Asymmetry: {abs(prop_asym):.1f}% {side}")
        benchmarks = """Key benchmarks for RSImod (Sole et al. 2018, NCAA Div 1):
- Males: Low <0.34, Average 0.35-0.48, High >0.49
- Females: Low <0.24, Average 0.25-0.37, High >0.38
Discuss ALL metrics listed including force and asymmetry values."""
    
    elif 'Squat Jump' in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics_lines.append(f"Jump Height: {jh:.1f} cm ({jh/2.54:.1f} in)")
        pp = safe_get('peak_relative_propulsive_power_w_kg')
        if pp:
            metrics_lines.append(f"Peak Power: {pp:.1f} W/kg")
        ppf = safe_get('peak_propulsive_force_n')
        ppf_asym = safe_get('sj_lr_peak_propulsive_force') or safe_get('lr_peak_propulsive_force')
        if ppf:
            metrics_lines.append(f"Peak Propulsive Force: {ppf/bw:.2f} x BW")
        if ppf_asym:
            side = 'Left' if ppf_asym > 0 else 'Right'
            metrics_lines.append(f"Peak Propulsive Force Asymmetry: {abs(ppf_asym):.1f}% {side}")
        imp_asym = safe_get('sj_lr_propulsive_impulse_index') or safe_get('lr_propulsive_impulse_index')
        if imp_asym:
            side = 'Left' if imp_asym > 0 else 'Right'
            metrics_lines.append(f"Impulse Asymmetry: {abs(imp_asym):.1f}% {side}")
        benchmarks = "SJ measures concentric-only power. Compare to CMJ for EUR calculation."
    
    elif 'Multi Rebound' in test_name:
        rsi = safe_get('rsi') or safe_get('avg_rsi')
        if rsi:
            metrics_lines.append(f"RSI: {rsi:.2f}")
        avg_h = safe_get('avg_jump_height_m')
        if avg_h:
            metrics_lines.append(f"Avg Jump Height: {avg_h*100:.1f} cm")
        avg_ct = safe_get('avg_contact_time_s')
        if avg_ct:
            metrics_lines.append(f"Avg Contact Time: {avg_ct*1000:.0f} ms")
        pf_asym = safe_get('lr_peak_force')
        if pf_asym:
            side = 'Left' if pf_asym > 0 else 'Right'
            metrics_lines.append(f"Peak Force Asymmetry: {abs(pf_asym):.1f}% {side}")
        benchmarks = """RSI benchmarks (10-5 repeat jump, Flanagan):
- Sprinters: 1.8-2.5
- Gymnastics: 1.75-2.2
- Hockey: 1.35-1.9
- Football: 1.3-1.9
- Sub Elite Field Sports: 1.1-1.65
- Recreation Field Sports: 0.75-1.7
Contact time <250ms indicates good reactive stiffness. >400ms suggests more ground contact training needed."""
    
    elif 'Bodyweight Squats' in test_name:
        avg_asym = safe_get('lr_avg_force')
        if avg_asym:
            side = 'Left' if avg_asym > 0 else 'Right'
            metrics_lines.append(f"Avg Force Asymmetry: {abs(avg_asym):.1f}% {side}")
        peak_asym = safe_get('lr_peak_force')
        if peak_asym:
            side = 'Left' if peak_asym > 0 else 'Right'
            metrics_lines.append(f"Peak Force Asymmetry: {abs(peak_asym):.1f}% {side}")
        benchmarks = "BW squats screen for bilateral symmetry. <10% asymmetry is acceptable."
    
    if not metrics_lines:
        return None
    
    prompt = f"""You are an expert sports scientist. Write a brief interpretation (3-5 sentences) of this {test_name} force plate assessment.

RESULTS:
{chr(10).join(metrics_lines)}

{benchmarks}

Guidelines:
- Write in second person ("Your results show...")
- Discuss EVERY metric listed above - do not skip any
- Reference the benchmarks to contextualize the values
- Flag asymmetries >10% as worth monitoring, >15% as needing attention
- Give one practical recommendation
- Do NOT use bullet points - write in flowing sentences
- Keep it concise and actionable"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    except Exception as e:
        print(f"  Claude API error for {test_name}: {e}")
        return None


# ==============================================================================
# RAG INTEGRATION FOR AI SUMMARIES
# ==============================================================================

def get_rag_system():
    """
    Initialize and return the RAG system for generating AI interpretations.
    Returns None if RAG is not available.
    """
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("rag_research", RAG_SCRIPT_PATH)
        rag_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(rag_module)
        return rag_module.ResearchRAG()
    except Exception as e:
        print(f"  Note: RAG system not available ({e})")
        return None


def generate_dashboard_ai_summary(found_tests, athlete_name, rag=None):
    """
    Generate AI summary for the Performance Dashboard.
    
    Args:
        found_tests: Dictionary of test data
        athlete_name: Athlete's name
        rag: ResearchRAG instance (optional)
    
    Returns:
        AI-generated summary string, or None if RAG not available
    """
    if rag is None:
        return None
    
    # Extract key metrics
    metrics = []
    cmj_height = sj_height = eur = rsi = mr_rsi = None
    
    for test_name, test_data in found_tests.items():
        row = test_data.get('row')
        if row is None:
            continue
        
        def safe_get(col):
            if col in row.index and pd.notna(row[col]):
                return row[col]
            return None
        
        if 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                cmj_height = jh
                metrics.append(f"CMJ height: {jh:.1f} cm ({jh/2.54:.1f} inches)")
            mrsi = safe_get('mrsi')
            if mrsi:
                rsi = mrsi
                metrics.append(f"RSImod: {mrsi:.2f}")
        
        elif 'Squat Jump' in test_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                sj_height = jh
                metrics.append(f"SJ height: {jh:.1f} cm ({jh/2.54:.1f} inches)")
        
        elif 'Multi Rebound' in test_name:
            r = safe_get('rsi') or safe_get('avg_rsi')
            if r:
                mr_rsi = r
                metrics.append(f"Reactive RSI: {r:.2f}")
        
        elif 'Drop Landing' in test_name:
            pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
            sw = safe_get('system_weight_n')
            if pf and sw:
                metrics.append(f"Landing peak force: {pf/sw:.2f} x bodyweight")
    
    # Calculate EUR
    if cmj_height and sj_height and sj_height > 0:
        eur = cmj_height / sj_height
        metrics.append(f"EUR (CMJ/SJ ratio): {eur:.2f}")
    
    if not metrics:
        return None
    
    # Build the query
    metrics_str = "\n".join(metrics)
    query = f"""Based on the following force plate assessment results for an athlete, provide a brief interpretation (3-5 sentences) of their performance profile. Focus on:
1. Overall explosive power
2. Elastic energy utilization (if EUR available)
3. Any notable strengths or areas for improvement

Assessment Results:
{metrics_str}

Provide practical, actionable insights for a coach or athlete."""

    try:
        response = rag.query(query, n_results=3, show_sources=False)
        # Remove the sources line if present
        if "📚 Sources:" in response:
            response = response.split("📚 Sources:")[0].strip()
        return response
    except Exception as e:
        print(f"  Warning: AI summary generation failed: {e}")
        return None


def generate_test_ai_interpretation(test_name, test_row, force_data, rag=None):
    """
    Generate AI interpretation for a specific test.
    
    Args:
        test_name: Name of the test
        test_row: Pandas Series with test metrics
        force_data: Force-time data DataFrame
        rag: ResearchRAG instance (optional)
    
    Returns:
        AI-generated interpretation string, or None if RAG not available
    """
    if rag is None or test_row is None:
        return None
    
    def safe_get(col):
        if col in test_row.index and pd.notna(test_row[col]):
            return test_row[col]
        return None
    
    # Build metrics description based on test type
    metrics = []
    body_weight = safe_get('system_weight_n')
    
    if 'Drop Landing' in test_name:
        pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
        if pf and body_weight:
            metrics.append(f"Peak landing force: {pf/body_weight:.2f} x BW")
        tts = safe_get('time_to_stabilization_ms')
        if tts:
            metrics.append(f"Time to stabilization: {tts:.0f} ms")
        pf_asym = safe_get('lr_peak_force') or safe_get('lr_peak_landing_force')
        if pf_asym:
            metrics.append(f"Peak force asymmetry: {pf_asym:.1f}%")
    
    elif 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics.append(f"Jump height: {jh:.1f} cm ({jh/2.54:.1f} in)")
        mrsi = safe_get('mrsi')
        if mrsi:
            metrics.append(f"RSImod: {mrsi:.2f}")
        prop_asym = safe_get('lr_propulsive_impulse_index')
        if prop_asym:
            metrics.append(f"Propulsive impulse asymmetry: {prop_asym:.1f}%")
    
    elif 'Squat Jump' in test_name:
        jh = safe_get('jump_height_imp_mom_cm')
        if jh is None:
            jh_m = safe_get('jump_height_m')
            if jh_m: jh = jh_m * 100
        if jh:
            metrics.append(f"Jump height: {jh:.1f} cm ({jh/2.54:.1f} in)")
        pp = safe_get('peak_power_w_kg')
        if pp:
            metrics.append(f"Peak power: {pp:.1f} W/kg")
    
    elif 'Multi Rebound' in test_name:
        rsi = safe_get('rsi') or safe_get('avg_rsi')
        if rsi:
            metrics.append(f"RSI: {rsi:.2f}")
        avg_ct = safe_get('avg_contact_time_ms')
        if avg_ct:
            metrics.append(f"Avg contact time: {avg_ct:.0f} ms")
        pf_asym = safe_get('lr_peak_force')
        if pf_asym:
            metrics.append(f"Peak force asymmetry: {pf_asym:.1f}%")
    
    elif 'Bodyweight Squats' in test_name:
        avg_asym = safe_get('lr_avg_force')
        if avg_asym:
            metrics.append(f"Average force asymmetry: {avg_asym:.1f}%")
        peak_asym = safe_get('lr_peak_force')
        if peak_asym:
            metrics.append(f"Peak force asymmetry: {peak_asym:.1f}%")
    
    if not metrics:
        return None
    
    metrics_str = "\n".join(metrics)
    
    query = f"""For a {test_name} force plate assessment with these results:

{metrics_str}

Provide a brief interpretation (2-3 sentences) explaining:
1. What these results indicate about the athlete's performance
2. Any asymmetries or concerns if present
3. One practical recommendation

Keep it concise and actionable for a coach."""

    try:
        response = rag.query(query, n_results=3, show_sources=False)
        if "📚 Sources:" in response:
            response = response.split("📚 Sources:")[0].strip()
        return response
    except Exception as e:
        print(f"  Warning: Test interpretation failed: {e}")
        return None

# ==============================================================================
# BRANDING
# ==============================================================================
COLOR_GOLD = '#f4bd2a'
COLOR_BLACK = '#000000'
COLOR_WHITE = '#ffffff'
COLOR_LIGHT_GRAY = '#f2f2f2'
COLOR_TEXT_MAIN = '#333333'
COLOR_GREEN = '#28a745'
COLOR_RED = '#dc3545'

# Asymmetry thresholds
ASYM_OK = 10        # <10% acceptable
ASYM_CONCERN = 15   # >15% needs attention

# ==============================================================================
# TEST DESCRIPTIONS (Sport-Agnostic)
# ==============================================================================
TEST_DESCRIPTIONS = {
    'Countermovement Jump': {
        'title': 'Countermovement Jump (CMJ)',
        'what': 'A maximal vertical jump starting from standing, using a quick dip before jumping.',
        'why': 'Measures lower-body power, explosiveness, and how well you use the stretch-shortening cycle. Changes over time can indicate fatigue, recovery status, or training adaptations.',
        'metrics': [
            ('Jump Height', 'jump_height_imp_mom_cm', 'cm', 'How high you jumped'),
            ('RSImod', 'mrsi', '', 'Explosiveness: height ÷ time'),
            ('Time to Takeoff', 'time_to_takeoff_s', 's', 'How long to leave ground'),
            ('Propulsive Impulse', 'relative_propulsive_impulse_n_s_kg', 'N·s/kg', 'Total push applied'),
            ('Braking RFD', 'braking_rfd_relative', 'N/s/kg', 'Force absorption speed'),
            ('Dip Depth', 'countermovement_depth_m', 'cm', 'How deep you dipped'),
        ]
    },
    'Squat Jump': {
        'title': 'Squat Jump (SJ)',
        'what': 'A vertical jump starting from a held squat position with no countermovement.',
        'why': 'Measures concentric-only power without the benefit of the stretch-shortening cycle. Comparing SJ to CMJ helps assess how well you use elastic energy.',
        'metrics': [
            ('Jump Height', 'jump_height_imp_mom_cm', 'cm', 'How high you jumped'),
            ('Time to Takeoff', 'time_to_takeoff_ms', 'ms', 'How long to leave ground'),
            ('Peak Power', 'peak_relative_propulsive_power_w_kg', 'W/kg', 'Maximum power output'),
            ('Propulsive Impulse', 'relative_propulsive_impulse_n_s_kg', 'N·s/kg', 'Total push applied'),
            ('Peak Force', 'peak_propulsive_force_n', 'N', 'Maximum force produced'),
            ('Avg Force', 'avg_propulsive_force_n', 'N', 'Average force during push'),
        ]
    },
    'CMJ Rebound': {
        'title': 'CMJ Rebound',
        'what': 'A countermovement jump followed immediately by a rebound jump upon landing.',
        'why': 'Measures reactive strength and the ability to quickly absorb and redirect force. Important for sports requiring repeated explosive movements.',
        'metrics': [
            ('CMJ Height', 'cmj_jump_height_m', 'cm', 'Initial jump height'),
            ('Rebound Height', 'rebound_jump_height_m', 'cm', 'Second jump height'),
            ('RSI', 'rsi', '', 'Reactive strength index'),
            ('Contact Time', 'contact_time_s', 'ms', 'Time on ground between jumps'),
            ('Peak Landing Force', 'peak_landing_force_n', 'N', 'Impact force absorbed'),
            ('Rebound Ratio', 'rebound_ratio', '%', 'Rebound ÷ CMJ height'),
        ]
    },
    'Multi Rebound': {
        'title': 'Multi Rebound (Repeated Jumps)',
        'what': 'A series of repeated maximal jumps with minimal ground contact time.',
        'why': 'Measures reactive strength, leg stiffness, and ability to use elastic energy repeatedly. Reflects performance in activities requiring repeated jumping or bounding.',
        'metrics': [
            ('RSI', 'rsi', '', 'Reactive strength index'),
            ('Avg Jump Height', 'avg_jump_height_m', 'cm', 'Average height achieved'),
            ('Avg Contact Time', 'avg_contact_time_s', 'ms', 'Average ground time'),
            ('Best Jump Height', 'best_jump_height_m', 'cm', 'Highest single jump'),
            ('Jump Count', 'jump_count', '', 'Number of jumps'),
            ('Stiffness', 'avg_stiffness_kn_m', 'kN/m', 'Leg spring stiffness'),
        ]
    },
    'Drop Landing': {
        'title': 'Drop Landing',
        'what': 'Stepping off a box and landing on both feet, absorbing the impact.',
        'why': 'Assesses landing mechanics, force absorption ability, and bilateral symmetry during impact. Important for injury risk screening and return-to-sport decisions.',
        'metrics': [
            ('Peak Force', 'peak_force_n', 'N', 'Maximum impact force'),
            ('Avg Impact Force', 'avg_impact_force_n', 'N', 'Average landing force'),
            ('Impact RFD', 'impact_rfd_n_s', 'N/s', 'Rate of force development'),
            ('Time to Stabilize', 'time_to_stabilization_s', 's', 'Time to reach stability'),
            ('Stabilization Force', 'avg_stabilization_force_n', 'N', 'Force during stabilization'),
            ('Landing Impulse', 'landing_impulse_n_s', 'N·s', 'Total impact absorbed'),
        ]
    },
    'Bodyweight Squats': {
        'title': 'Bodyweight Squats',
        'what': 'Repeated bodyweight squats performed on dual force plates.',
        'why': 'Assesses bilateral symmetry and movement quality during a controlled, familiar movement pattern. Useful for screening and monitoring.',
        'metrics': [
            ('Avg Force', 'avg_force_n', 'N', 'Average force produced'),
            ('Peak Force', 'peak_force_n', 'N', 'Maximum force produced'),
            ('Force Asymmetry', 'force_asymmetry_percent', '%', 'Left-right difference'),
            ('Rep Count', 'rep_count', '', 'Number of repetitions'),
            ('Avg Rep Time', 'avg_rep_time_s', 's', 'Time per repetition'),
            ('Movement Quality', 'movement_quality_score', '', 'Consistency score'),
        ]
    },
    'Static Balance': {
        'title': 'Static Balance',
        'what': 'Standing as still as possible on the force plates for a set duration.',
        'why': 'Assesses postural control and stability. Changes may indicate fatigue, neurological factors, or injury effects.',
        'metrics': [
            ('Sway Area', 'sway_area_cm2', 'cm²', 'Total movement area'),
            ('Sway Velocity', 'sway_velocity_cm_s', 'cm/s', 'Speed of movement'),
            ('COP Path Length', 'cop_path_length_cm', 'cm', 'Total distance traveled'),
            ('ML Range', 'ml_range_cm', 'cm', 'Side-to-side movement'),
            ('AP Range', 'ap_range_cm', 'cm', 'Front-back movement'),
            ('Stability Score', 'stability_score', '', 'Overall stability'),
        ]
    },
    'Drop Jump': {
        'title': 'Drop Jump',
        'what': 'Stepping off a box and immediately jumping as high as possible upon landing.',
        'why': 'Measures reactive strength, landing-to-takeoff efficiency, and the ability to rapidly absorb and redirect force. Key for plyometric profiling and return-to-sport readiness.',
        'metrics': [
            ('Jump Height', 'jump_height_m', 'cm', 'How high you jumped after the drop'),
            ('RSI', 'rsi', '', 'Reactive strength index (height / contact time)'),
            ('Contact Time', 'total_contact_time_s', 'ms', 'Time on ground before takeoff'),
            ('Peak Landing Force', 'peak_landing_force_n', 'N', 'Impact force at landing'),
            ('Landing Stiffness', 'landing_stiffness_n_m', 'N/m', 'Leg stiffness at landing'),
            ('Drop Height', 'drop_height_m', 'cm', 'Box height for the drop'),
        ]
    },
    'Isometric Test': {
        'title': 'Isometric Mid-Thigh Pull (IMTP)',
        'what': 'A maximal isometric pull against an immovable bar positioned at mid-thigh height.',
        'why': 'Measures peak strength and rate of force development without movement. Useful for tracking maximal strength changes and comparing bilateral force production.',
        'metrics': [
            ('Peak Force', 'peak_force_n', 'N', 'Maximum force produced'),
            ('Peak Force Rel.', 'peak_relative_force', 'N/kg', 'Peak force per kg body mass'),
            ('Impulse', 'positive_impulse_n_s', 'N·s', 'Total positive impulse'),
            ('Net Impulse', 'positive_net_impulse_n_s', 'N·s', 'Net positive impulse'),
            ('Peak Left Force', 'peak_left_force_n', 'N', 'Peak force on left plate'),
            ('Peak Right Force', 'peak_right_force_n', 'N', 'Peak force on right plate'),
        ]
    },
}

# ==============================================================================
# HAWKIN API FUNCTIONS
# ==============================================================================

def setup_api():
    """Connect to Hawkin Dynamics API."""
    try:
        from hdforce import AuthManager
        AuthManager(authMethod="manual", refreshToken=API_TOKEN, env_file_name=None)
        print("✓ Connected to Hawkin API")
        return True
    except ImportError:
        print("ERROR: Run 'pip install hdforce'")
        return False
    except Exception as e:
        print(f"ERROR: {e}")
        return False


def get_athletes():
    """Get athlete list."""
    from hdforce import GetAthletes
    return GetAthletes()


def get_tests(athlete_id, days_back=30, to_date=None):
    """Get tests for an athlete. to_date is an optional YYYY-MM-DD string to cap the date range."""
    from hdforce import GetTests
    from_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y-%m-%d')
    tests = GetTests(athleteId=athlete_id, from_=from_date)
    # Filter to to_date if provided
    if to_date and tests is not None and len(tests) > 0:
        to_dt = datetime.strptime(to_date, '%Y-%m-%d')
        to_ts = int((to_dt + timedelta(days=1)).timestamp())  # include full end day
        for col in ['timestamp', 'test_time', 'created_at', 'date', 'testDate']:
            if col in tests.columns:
                def check_before(val):
                    if pd.isna(val):
                        return True
                    ts = int(val) if isinstance(val, (int, float, np.integer, np.floating)) else 0
                    if ts > 10000000000:
                        ts = ts // 1000
                    return ts < to_ts
                tests = tests[tests[col].apply(check_before)]
                break
    return tests


def get_force_time(test_id):
    """Get raw force-time data for a test."""
    try:
        from hdforce import GetForceTime
        return GetForceTime(testId=test_id)
    except:
        return None


def find_cop_file(athlete_name, test_type, data_dir=None):
    """
    Look for a COP CSV file matching the athlete and test type.
    
    COP file naming convention from Hawkin:
    - COP-[Athlete_Name]_[Test_Type].csv
    - COP-[Athlete_Name]_[Test_Type]_[Tag].csv
    
    Examples:
    - COP-Kristina_Robb_Drop_Landing.csv
    - COP-Holland_Fast_Free_Run_Bodyweight_Squats.csv
    """
    if data_dir is None:
        data_dir = SCRIPT_DIR
    
    # Clean athlete name for file matching
    athlete_clean = athlete_name.replace(' ', '_')
    
    # Map test types to possible file patterns
    test_patterns = {
        'Drop Landing': ['Drop_Landing', 'Drop Landing'],
        'Bodyweight Squats': ['Free_Run', 'Free Run', 'Bodyweight_Squats'],
        'Free Run': ['Free_Run', 'Free Run'],
    }
    
    patterns_to_try = test_patterns.get(test_type, [test_type.replace(' ', '_')])
    
    # Look for matching files
    for pattern in patterns_to_try:
        # Try different file patterns
        search_patterns = [
            os.path.join(data_dir, f"COP-{athlete_clean}_{pattern}*.csv"),
            os.path.join(data_dir, f"COP_{athlete_clean}_{pattern}*.csv"),
            os.path.join(data_dir, f"*COP*{athlete_clean}*{pattern}*.csv"),
        ]
        
        for search in search_patterns:
            matches = glob.glob(search)
            if matches:
                return matches[0]  # Return first match
    
    return None


def load_cop_data(filepath):
    """Load COP data from CSV file."""
    try:
        df = pd.read_csv(filepath)
        print(f"      Loaded COP: {os.path.basename(filepath)}")
        return df
    except Exception as e:
        print(f"      Error loading COP file: {e}")
        return None


# ==============================================================================
# DATA PROCESSING
# ==============================================================================

def smooth_data(data, window=50):
    """Rolling average smoothing."""
    if len(data) < window:
        return data
    return pd.Series(data).rolling(window=window, center=True, min_periods=1).mean().values


def create_stabilogram(ax, cop_x, cop_y, title, leg_color=COLOR_GOLD):
    """
    Create a stabilogram (COP path plot) for one leg.
    
    cop_x, cop_y: arrays in mm (will be converted to cm)
    Returns: (path_length, sway_area) or (None, None) if insufficient data
    """
    from scipy.ndimage import gaussian_filter1d
    
    # Clean data - remove NaN
    valid = np.isfinite(cop_x) & np.isfinite(cop_y)
    if np.sum(valid) < 10:
        ax.text(0.5, 0.5, 'Insufficient COP data', ha='center', va='center',
                fontsize=10, color='gray', transform=ax.transAxes)
        ax.set_title(title, fontsize=10, fontweight='bold')
        return None, None
    
    x_raw = cop_x[valid] / 10  # mm to cm
    y_raw = cop_y[valid] / 10
    
    # Apply Gaussian smoothing to create smoother lines
    # Sigma of 3-5 provides good smoothing without losing shape
    sigma = 4
    x = gaussian_filter1d(x_raw, sigma=sigma)
    y = gaussian_filter1d(y_raw, sigma=sigma)
    
    # Calculate path length from RAW data (smoothing would underestimate)
    dx_raw = np.diff(x_raw)
    dy_raw = np.diff(y_raw)
    path_length = np.sum(np.sqrt(dx_raw**2 + dy_raw**2))
    
    # Plot COP path with gradient coloring (light to dark over time)
    points = np.array([x, y]).T.reshape(-1, 1, 2)
    segments = np.concatenate([points[:-1], points[1:]], axis=1)
    
    n_seg = len(segments)
    colors = []
    for i in range(n_seg):
        alpha = 0.3 + (0.7 * i / n_seg)
        if leg_color == COLOR_GOLD:
            colors.append((244/255, 189/255, 42/255, alpha))
        else:
            colors.append((0.2, 0.2, 0.2, alpha))
    
    lc = LineCollection(segments, colors=colors, linewidth=1.5)
    ax.add_collection(lc)
    
    # Start and end markers (use smoothed positions)
    ax.plot(x[0], y[0], 'o', color='green', markersize=8, label='Start',
            markeredgecolor='white', markeredgewidth=1, zorder=5)
    ax.plot(x[-1], y[-1], 'o', color='red', markersize=8, label='End',
            markeredgecolor='white', markeredgewidth=1, zorder=5)
    
    # Calculate 95% confidence ellipse from RAW data (more accurate)
    sway_area = None
    try:
        mean_x, mean_y = np.mean(x_raw), np.mean(y_raw)
        cov = np.cov(x_raw, y_raw)
        eigenvalues, eigenvectors = np.linalg.eig(cov)
        eigenvalues = np.sqrt(np.abs(eigenvalues)) * 2.447  # 95% confidence
        
        # Ellipse area = π * a * b
        sway_area = np.pi * eigenvalues[0] * eigenvalues[1]
        
        angle = np.rad2deg(np.arctan2(eigenvectors[1, 0], eigenvectors[0, 0]))
        ellipse = Ellipse((mean_x, mean_y), width=eigenvalues[0]*2, height=eigenvalues[1]*2,
                         angle=angle, facecolor='none', edgecolor=COLOR_BLACK,
                         linewidth=2, linestyle='--', label='95% Ellipse')
        ax.add_patch(ellipse)
    except:
        pass
    
    # Formatting - Updated axis labels
    ax.set_xlabel('Left & Right (cm)', fontsize=9, fontweight='bold')
    ax.set_ylabel('Forward & Backward (cm)', fontsize=9, fontweight='bold')
    ax.set_title(title, fontsize=10, fontweight='bold', color=COLOR_TEXT_MAIN)
    ax.legend(loc='upper right', fontsize=8)
    ax.grid(True, alpha=0.3, linestyle=':')
    ax.axis('equal')
    ax.autoscale()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    return path_length, sway_area


def get_test_by_type(tests_df, pattern):
    """Filter tests by type pattern and return most recent."""
    if 'testType_name' not in tests_df.columns:
        return None
    
    filtered = tests_df[tests_df['testType_name'].str.contains(pattern, case=False, na=False, regex=True)]
    if len(filtered) == 0:
        return None
    
    return filtered.iloc[-1]  # Most recent


def get_all_tests_by_type(tests_df, pattern):
    """Filter tests by type pattern and return ALL matching tests (for trial selection)."""
    if 'testType_name' not in tests_df.columns:
        return None
    
    # Use regex=True for patterns with special characters like ^ and $
    filtered = tests_df[tests_df['testType_name'].str.contains(pattern, case=False, na=False, regex=True)]
    if len(filtered) == 0:
        return None
    
    return filtered


def format_trial_info(test_row, test_type, tag_col=None):
    """Format trial information for display during selection."""
    # Get timestamp
    time_str = "Unknown time"
    for col in ['timestamp', 'test_time', 'created_at', 'date']:
        if col in test_row.index and pd.notna(test_row[col]):
            ts = test_row[col]
            try:
                if isinstance(ts, (int, float, np.integer, np.floating)):
                    ts_val = int(ts)
                    if ts_val > 10000000000:
                        ts_val = ts_val // 1000
                    time_str = datetime.fromtimestamp(ts_val).strftime('%m/%d/%Y %H:%M')
                elif isinstance(ts, str):
                    if ts.isdigit():
                        ts_val = int(ts)
                        if ts_val > 10000000000:
                            ts_val = ts_val // 1000
                        time_str = datetime.fromtimestamp(ts_val).strftime('%m/%d/%Y %H:%M')
                    else:
                        time_str = ts[:16] if len(ts) > 16 else ts
                break
            except:
                pass
    
    # Get tag
    tag = get_tag_value(test_row, tag_col)
    tag_str = f" [{tag}]" if tag else ""
    
    # Get key metrics based on test type
    metrics = []
    
    # CMJ Rebound - uses rebound-specific metrics
    if 'CMJ Rebound' in test_type or 'Countermovement' in test_type and 'Rebound' in test_type:
        # Rebound Jump Height
        for col in ['rebound_jump_height_m', 'avg_rebound_jump_height_m', 'rebound_jump_height_cm']:
            if col in test_row.index and pd.notna(test_row[col]):
                val = test_row[col]
                if '_m' in col and val < 1:  # meters
                    metrics.append(f"Reb JH: {val * 100:.1f} cm")
                else:
                    metrics.append(f"Reb JH: {val:.1f} cm")
                break
        # Rebound RSI
        for col in ['rebound_rsi', 'avg_rebound_rsi', 'rsi']:
            if col in test_row.index and pd.notna(test_row[col]):
                metrics.append(f"RSI: {test_row[col]:.2f}")
                break
    
    # Regular CMJ (not rebound)
    elif 'CMJ' in test_type or 'Countermovement' in test_type:
        # Jump Height (check both cm and m versions)
        if 'jump_height_imp_mom_cm' in test_row.index and pd.notna(test_row['jump_height_imp_mom_cm']):
            metrics.append(f"JH: {test_row['jump_height_imp_mom_cm']:.1f} cm")
        elif 'jump_height_cm' in test_row.index and pd.notna(test_row['jump_height_cm']):
            metrics.append(f"JH: {test_row['jump_height_cm']:.1f} cm")
        elif 'jump_height_m' in test_row.index and pd.notna(test_row['jump_height_m']):
            metrics.append(f"JH: {test_row['jump_height_m'] * 100:.1f} cm")
        # mRSI (modified Reactive Strength Index)
        if 'mrsi' in test_row.index and pd.notna(test_row['mrsi']):
            metrics.append(f"mRSI: {test_row['mrsi']:.2f}")
    
    elif 'Squat Jump' in test_type:
        # Jump Height
        if 'jump_height_imp_mom_cm' in test_row.index and pd.notna(test_row['jump_height_imp_mom_cm']):
            metrics.append(f"JH: {test_row['jump_height_imp_mom_cm']:.1f} cm")
        elif 'jump_height_cm' in test_row.index and pd.notna(test_row['jump_height_cm']):
            metrics.append(f"JH: {test_row['jump_height_cm']:.1f} cm")
        elif 'jump_height_m' in test_row.index and pd.notna(test_row['jump_height_m']):
            metrics.append(f"JH: {test_row['jump_height_m'] * 100:.1f} cm")
        # Time to Takeoff
        if 'time_to_takeoff_s' in test_row.index and pd.notna(test_row['time_to_takeoff_s']):
            metrics.append(f"TTT: {test_row['time_to_takeoff_s'] * 1000:.0f} ms")
        elif 'contraction_time_s' in test_row.index and pd.notna(test_row['contraction_time_s']):
            metrics.append(f"TTT: {test_row['contraction_time_s'] * 1000:.0f} ms")
    
    # Multi Rebound (not CMJ Rebound)
    elif 'Multi Rebound' in test_type:
        if 'rsi' in test_row.index and pd.notna(test_row['rsi']):
            metrics.append(f"RSI: {test_row['rsi']:.2f}")
        if 'avg_contact_time_ms' in test_row.index and pd.notna(test_row['avg_contact_time_ms']):
            metrics.append(f"GCT: {test_row['avg_contact_time_ms']:.0f} ms")
    
    elif 'Drop Landing' in test_type:
        if 'peak_landing_force_n' in test_row.index and pd.notna(test_row['peak_landing_force_n']):
            metrics.append(f"Peak Landing: {test_row['peak_landing_force_n']:.0f} N")
        elif 'peak_force_n' in test_row.index and pd.notna(test_row['peak_force_n']):
            metrics.append(f"Peak Force: {test_row['peak_force_n']:.0f} N")

    elif 'Drop Jump' in test_type:
        for col in ['jump_height_m', 'peak_jump_height_m']:
            if col in test_row.index and pd.notna(test_row[col]):
                metrics.append(f"JH: {test_row[col] * 100:.1f} cm")
                break
        if 'rsi' in test_row.index and pd.notna(test_row['rsi']):
            metrics.append(f"RSI: {test_row['rsi']:.2f}")

    elif 'Isometric' in test_type:
        for col in ['peak_vertical_force_n', 'peak_force_n']:
            if col in test_row.index and pd.notna(test_row[col]):
                metrics.append(f"Peak Force: {test_row[col]:.0f} N")
                break
        for col in ['relative_peak_vertical_force_n_kg', 'peak_force_relative_n_kg']:
            if col in test_row.index and pd.notna(test_row[col]):
                metrics.append(f"Rel: {test_row[col]:.1f} N/kg")
                break

    elif 'Bodyweight' in test_type or 'Balance' in test_type or 'Free Run' in test_type:
        if 'lr_avg_force' in test_row.index and pd.notna(test_row['lr_avg_force']):
            metrics.append(f"L|R: {test_row['lr_avg_force']:.1f}%")

    metric_str = " | ".join(metrics) if metrics else ""
    
    return time_str, metric_str, tag_str


def select_trial(trials_df, display_name, pattern, auto_select_best=False, tag_col=None):
    """
    Let user select which trial to use from available trials.
    Always prompts, even for single trials. Allows skipping.
    
    Args:
        trials_df: DataFrame of available trials
        display_name: Name of test for display
        pattern: Regex pattern used to find tests
        auto_select_best: If True, automatically select best trial based on key metric
        tag_col: Column name for tags (optional)
        
    Returns the selected test row, or None if skipped.
    """
    if trials_df is None or len(trials_df) == 0:
        return None
    
    num_trials = len(trials_df)
    
    # Determine best trial based on test type
    best_idx = num_trials - 1  # Default to most recent
    best_metric_name = None
    best_metric_value = None
    
    if 'CMJ' in display_name or 'Countermovement' in display_name:
        if 'Rebound' not in display_name:
            # For regular CMJ, best = highest mRSI
            if 'mrsi' in trials_df.columns:
                valid_mrsi = trials_df['mrsi'].dropna()
                if len(valid_mrsi) > 0:
                    best_idx = trials_df['mrsi'].idxmax()
                    best_idx = trials_df.index.get_loc(best_idx)
                    best_metric_name = 'mRSI'
                    best_metric_value = trials_df.iloc[best_idx]['mrsi']
        else:
            # For CMJ Rebound, best = highest rebound RSI
            for col in ['rebound_rsi', 'rsi']:
                if col in trials_df.columns:
                    valid = trials_df[col].dropna()
                    if len(valid) > 0:
                        best_idx = trials_df[col].idxmax()
                        best_idx = trials_df.index.get_loc(best_idx)
                        best_metric_name = 'RSI'
                        best_metric_value = trials_df.iloc[best_idx][col]
                        break
    
    elif 'Squat Jump' in display_name:
        # Best = highest jump height
        for col in ['jump_height_imp_mom_cm', 'jump_height_m']:
            if col in trials_df.columns:
                valid = trials_df[col].dropna()
                if len(valid) > 0:
                    best_idx = trials_df[col].idxmax()
                    best_idx = trials_df.index.get_loc(best_idx)
                    best_metric_name = 'Jump Height'
                    val = trials_df.iloc[best_idx][col]
                    best_metric_value = val * 100 if col.endswith('_m') else val
                    break
    
    elif 'Multi Rebound' in display_name:
        # Best = highest RSI
        if 'rsi' in trials_df.columns:
            valid = trials_df['rsi'].dropna()
            if len(valid) > 0:
                best_idx = trials_df['rsi'].idxmax()
                best_idx = trials_df.index.get_loc(best_idx)
                best_metric_name = 'RSI'
                best_metric_value = trials_df.iloc[best_idx]['rsi']
    
    elif 'Drop Landing' in display_name:
        # Best = lowest peak landing force (safer landing)
        if 'peak_landing_force_n' in trials_df.columns:
            valid = trials_df['peak_landing_force_n'].dropna()
            if len(valid) > 0:
                best_idx = trials_df['peak_landing_force_n'].idxmin()
                best_idx = trials_df.index.get_loc(best_idx)
                best_metric_name = 'Peak Landing'
                best_metric_value = trials_df.iloc[best_idx]['peak_landing_force_n']

    elif 'Drop Jump' in display_name:
        # Best = highest RSI
        if 'rsi' in trials_df.columns:
            valid = trials_df['rsi'].dropna()
            if len(valid) > 0:
                best_idx = trials_df['rsi'].idxmax()
                best_idx = trials_df.index.get_loc(best_idx)
                best_metric_name = 'RSI'
                best_metric_value = trials_df.iloc[best_idx]['rsi']

    elif 'Isometric' in display_name:
        # Best = highest peak force
        for col in ['peak_force_n']:
            if col in trials_df.columns:
                valid = trials_df[col].dropna()
                if len(valid) > 0:
                    best_idx = trials_df[col].idxmax()
                    best_idx = trials_df.index.get_loc(best_idx)
                    best_metric_name = 'Peak Force'
                    best_metric_value = trials_df.iloc[best_idx][col]
                    break

    # Auto-select if requested
    if auto_select_best:
        if best_metric_name:
            print(f"\n  {display_name}: Auto-selected best trial (#{best_idx + 1}) - {best_metric_name}: {best_metric_value:.2f}")
        else:
            print(f"\n  {display_name}: Auto-selected most recent trial (#{num_trials})")
        return trials_df.iloc[best_idx]
    
    print(f"\n  {display_name}: {num_trials} trial(s) available")
    print("  " + "-" * 50)
    
    for i, (_, row) in enumerate(trials_df.iterrows()):
        time_str, metric_str, tag_str = format_trial_info(row, display_name, tag_col)
        metric_display = f" | {metric_str}" if metric_str else ""
        best_marker = " ★ BEST" if i == best_idx and best_metric_name else ""
        print(f"    {i+1}. {time_str}{tag_str}{metric_display}{best_marker}")
    
    print(f"    [Enter] = Use most recent (#{num_trials})")
    print(f"    [b] = Use best performance (#{best_idx + 1})" + (f" - {best_metric_name}: {best_metric_value:.2f}" if best_metric_name else ""))
    print(f"    [s] = Skip this test (exclude from report)")
    print(f"    [d] = Show debug info (available columns)")
    
    while True:
        selection = input(f"  Select trial (1-{num_trials}, b=best, s=skip): ").strip().lower()
        
        # Skip this test
        if selection == 's' or selection == '0':
            print(f"    → Skipping {display_name}")
            return None
        
        # Default to most recent
        if selection == "":
            print(f"    → Using most recent trial")
            return trials_df.iloc[-1]
        
        # Best performance option
        if selection == 'b':
            if best_metric_name:
                print(f"    → Using best performance (#{best_idx + 1}) - {best_metric_name}: {best_metric_value:.2f}")
            else:
                print(f"    → Using most recent (no metric available to rank)")
            return trials_df.iloc[best_idx]
        
        # Debug option - show available columns
        if selection == 'd':
            print(f"\n    Available columns for {display_name}:")
            row = trials_df.iloc[0]
            # Show columns that might be useful metrics
            interesting_cols = [c for c in row.index if any(x in c.lower() for x in 
                ['height', 'rsi', 'time', 'takeoff', 'contraction', 'force', 'power', 'velocity', 'lr_', 'left', 'right', 'impulse', 'rfd'])]
            for col in sorted(interesting_cols):
                val = row[col]
                if pd.notna(val):
                    print(f"      {col}: {val}")
            print()
            continue
        
        # Validate selection
        if selection.isdigit():
            idx = int(selection) - 1
            if 0 <= idx < num_trials:
                time_str, _, _ = format_trial_info(trials_df.iloc[idx], display_name, tag_col)
                print(f"    → Selected trial from {time_str}")
                return trials_df.iloc[idx]
        
        print(f"    Invalid selection. Enter 1-{num_trials}, 's' to skip, or Enter for most recent.")


def extract_asymmetries(test_row, test_name):
    """Extract ALL L|R asymmetry values from a test row dynamically."""
    asymmetries = []
    
    # Find all columns that contain 'lr_' (asymmetry columns)
    lr_columns = [col for col in test_row.index if 'lr_' in col.lower()]
    
    for col in sorted(lr_columns):
        val = test_row[col]
        if pd.notna(val):
            # Parse the column structure
            col_lower = col.lower()
            
            # Split around 'lr_' to get prefix and suffix
            lr_idx = col_lower.find('lr_')
            prefix = col_lower[:lr_idx].rstrip('_')  # Part before lr_ (e.g., 'cmj', 'rebound', '')
            suffix = col_lower[lr_idx + 3:]  # Part after lr_ (e.g., 'avg_braking_force')
            
            # Create readable label
            label = suffix.replace('_', ' ').title()
            label = label.replace('Avg ', 'Avg. ')
            label = label.replace('Rfd', 'RFD')
            label = label.replace(' Index', ' Idx')
            
            # Add prefix to label if present
            if prefix:
                prefix_label = prefix.title()
                prefix_label = prefix_label.replace('Cmj', 'CMJ')
                label = f"{prefix_label} {label}"
            
            # Determine status
            abs_val = abs(val)
            if abs_val < ASYM_OK:
                status = 'OK'
            elif abs_val < ASYM_CONCERN:
                status = 'MONITOR'
            else:
                status = 'ADDRESS'
            
            # Format: [Test, Metric, Asymmetry (numeric), Asymmetry (string), Status]
            asymmetries.append([test_name, label, val, f'{val:.1f}%', status])
    
    return asymmetries


# ==============================================================================
# TEST TYPE PARSING, TAG DETECTION & SESSION GROUPING
# ==============================================================================


def parse_test_variant(testType_name):
    """Parse testType_name into (base_type, variant/tag).

    Examples:
        'Countermovement Jump' → ('Countermovement Jump', None)
        'Countermovement Jump-Single Side - Left' → ('Countermovement Jump', 'Single Side - Left')
        'Free Run-Bodyweight Squats' → ('Bodyweight Squats', None)
        'Free Run-Balance - Firm Surface - Eyes Open' → ('Static Balance', 'Firm Surface - Eyes Open')
    """
    # Handle Free Run prefix (Hawkin wraps some tests under "Free Run-")
    if testType_name.startswith('Free Run-'):
        remainder = testType_name[len('Free Run-'):]
        if 'Bodyweight' in remainder or 'bodyweight' in remainder:
            return ('Bodyweight Squats', None)
        elif 'Balance' in remainder:
            variant = remainder.replace('Balance - ', '').strip()
            return ('Static Balance', variant if variant else None)
        else:
            return (remainder, None)

    # Handle hyphenated variants (e.g., "CMJ-Single Side - Left")
    # Split on first hyphen that separates test type from variant
    known_bases = [
        'Countermovement Jump', 'Squat Jump', 'Multi Rebound',
        'Drop Landing', 'Drop Jump', 'CMJ Rebound', 'Isometric Test'
    ]
    for base in known_bases:
        if testType_name.startswith(base):
            remainder = testType_name[len(base):].lstrip('-').strip()
            return (base, remainder if remainder else None)

    return (testType_name, None)


# Candidate column names for tags in Hawkin API DataFrames
# NOTE: 'segment' is excluded — it contains trial rep numbering (e.g., "CMJ:1"),
# not meaningful tag/variant information.
TAG_COLUMN_CANDIDATES = ['testType_tag_name', 'tag_name', 'tags', 'tag',
                         'testType_tagName', 'tagName', 'test_type_tag_name',
                         'testType_tag']


def _col_has_scalar_strings(df, col):
    """Check if a column contains scalar string values (not lists/dicts)."""
    vals = df[col].dropna()
    if len(vals) == 0:
        return False
    # Check first non-null value — skip columns with list/dict values (e.g. tag_ids)
    sample = vals.iloc[0]
    if isinstance(sample, (list, dict)):
        return False
    return True


def detect_tag_column(df):
    """Auto-detect the tag column in a Hawkin tests DataFrame.
    Returns the column name if found, else None.
    Skips columns containing list/dict values (e.g. tag_ids)."""
    for candidate in TAG_COLUMN_CANDIDATES:
        if candidate in df.columns and _col_has_scalar_strings(df, candidate):
            return candidate
    # Fallback: look for any column containing 'tag' (case-insensitive)
    # Skip *_ids columns and columns with non-scalar values
    for col in df.columns:
        if 'tag' in col.lower() and col.lower() not in ('testtype_name',) and not col.endswith('_ids'):
            if _col_has_scalar_strings(df, col):
                return col
    return None


def get_tag_value(row, tag_col):
    """Safely extract a tag string from a row. Returns None for empty/NaN."""
    if tag_col is None:
        return None
    if tag_col not in row.index:
        return None
    val = row[tag_col]
    if pd.isna(val):
        return None
    val = str(val).strip()
    if val == '' or val.lower() in ('nan', 'none', 'null'):
        return None
    return val


def group_trials_by_tag(trials_df, tag_col):
    """Group a DataFrame of trials by their tag value.
    Returns dict of {tag_value_or_None: sub_DataFrame}."""
    groups = {}
    for idx, row in trials_df.iterrows():
        tag = get_tag_value(row, tag_col)
        if tag not in groups:
            groups[tag] = []
        groups[tag].append(idx)
    return {tag: trials_df.loc[idxs] for tag, idxs in groups.items()}


def get_trial_date(row):
    """Extract date (as datetime.date) from a trial row timestamp."""
    for col in ['timestamp', 'test_time', 'created_at', 'date', 'testDate']:
        if col in row.index and pd.notna(row[col]):
            ts = row[col]
            try:
                if isinstance(ts, (int, float, np.integer, np.floating)):
                    ts_val = int(ts)
                    if ts_val > 10000000000:
                        ts_val = ts_val // 1000
                    return datetime.fromtimestamp(ts_val).date()
                elif isinstance(ts, str) and ts.isdigit():
                    ts_val = int(ts)
                    if ts_val > 10000000000:
                        ts_val = ts_val // 1000
                    return datetime.fromtimestamp(ts_val).date()
            except:
                pass
    return datetime.now().date()


def group_trials_by_session(trials_df):
    """Group trials by session date (calendar day).
    Returns list of session dicts sorted by date ascending:
    [{'date': date_obj, 'date_str': 'MM/DD/YYYY', 'trials': DataFrame}, ...]
    """
    session_map = {}
    for idx, row in trials_df.iterrows():
        d = get_trial_date(row)
        if d not in session_map:
            session_map[d] = []
        session_map[d].append(idx)

    sessions = []
    for d in sorted(session_map.keys()):
        sessions.append({
            'date': d,
            'date_str': d.strftime('%m/%d/%Y'),
            'trials': trials_df.loc[session_map[d]]
        })
    return sessions


# Metrics to average per session and to plot as trends, keyed by base test type.
# Each entry: (display_name, column_name_or_list, unit, conversion_func_or_None)
TREND_METRICS = {
    'Countermovement Jump': [
        ('Jump Height', ['jump_height_imp_mom_cm', 'jump_height_m'], 'cm',
         lambda v, col: v * 100 if 'jump_height_m' == col else v),
        ('RSImod', ['mrsi'], '', None),
        ('Prop. Impulse', ['relative_propulsive_impulse_n_s_kg'], 'N·s/kg', None),
        ('Braking RFD', ['braking_rfd_relative'], 'N/s/kg', None),
    ],
    'Squat Jump': [
        ('Jump Height', ['jump_height_imp_mom_cm', 'jump_height_m'], 'cm',
         lambda v, col: v * 100 if 'jump_height_m' == col else v),
        ('Peak Power', ['peak_relative_propulsive_power_w_kg'], 'W/kg', None),
        ('Prop. Impulse', ['relative_propulsive_impulse_n_s_kg'], 'N·s/kg', None),
    ],
    'CMJ Rebound': [
        ('CMJ Height', ['cmj_jump_height_m'], 'cm',
         lambda v, col: v * 100),
        ('Rebound Height', ['rebound_jump_height_m'], 'cm',
         lambda v, col: v * 100),
        ('RSI', ['rsi', 'rebound_rsi'], '', None),
    ],
    'Multi Rebound': [
        ('RSI', ['rsi', 'avg_rsi'], '', None),
        ('Avg Height', ['avg_jump_height_m'], 'cm',
         lambda v, col: v * 100),
        ('Avg Contact', ['avg_contact_time_s'], 'ms',
         lambda v, col: v * 1000),
    ],
    'Drop Landing': [
        ('Peak Force', ['peak_force_n', 'peak_landing_force_n'], '×BW', 'bw_normalize'),
        ('Time to Stable', ['time_to_stabilization_ms', 'time_to_stabilization_s'], 'ms',
         lambda v, col: v * 1000 if '_s' in col and v < 10 else v),
    ],
    'Bodyweight Squats': [
        ('Avg Force Asym', ['lr_avg_force'], '%', None),
        ('Peak Force Asym', ['lr_peak_force'], '%', lambda v, col: abs(v)),
    ],
    'Static Balance': [
        ('Sway Area', ['sway_area_cm2', 'cop_sway_area_cm2'], 'cm²', None),
        ('COP Path', ['cop_path_length_cm', 'total_cop_path_length_cm'], 'cm', None),
    ],
    'Drop Jump': [
        ('Jump Height', ['jump_height_m', 'peak_jump_height_m'], 'cm',
         lambda v, col: v * 100),
        ('RSI', ['rsi', 'peak_rsi'], '', None),
        ('Contact Time', ['total_contact_time_s'], 'ms',
         lambda v, col: v * 1000),
    ],
    'Isometric Test': [
        ('Peak Force', ['peak_force_n'], 'N', None),
        ('Rel. Peak Force', ['peak_relative_force'], 'N/kg', None),
        ('Impulse', ['positive_impulse_n_s'], 'N·s', None),
    ],
}


def compute_session_averages(session_trials_df, test_type):
    """Compute average metrics for a session (group of trials on the same day).
    Returns dict of {metric_display_name: {'value': float, 'unit': str}}.
    """
    # Find matching trend metrics config
    metrics_config = None
    for key, config in TREND_METRICS.items():
        if key in test_type:
            metrics_config = config
            break
    if metrics_config is None:
        return {}

    averages = {}
    for display_name, col_candidates, unit, converter in metrics_config:
        values = []
        used_col = None
        for col in col_candidates:
            if col in session_trials_df.columns:
                valid = session_trials_df[col].dropna()
                if len(valid) > 0:
                    used_col = col
                    for v in valid:
                        if converter == 'bw_normalize':
                            # Normalize by body weight
                            bw_col = session_trials_df.get('system_weight_n')
                            if bw_col is not None:
                                bw_vals = bw_col.dropna()
                                bw = bw_vals.iloc[0] if len(bw_vals) > 0 else 800
                            else:
                                bw = 800
                            values.append(v / bw)
                        elif converter and callable(converter):
                            values.append(converter(v, col))
                        else:
                            values.append(v)
                    break

        if values:
            averages[display_name] = {
                'value': np.mean(values),
                'unit': unit,
                'n_trials': len(values),
                'values': values  # individual trial values for scatter
            }

    return averages


def make_display_label(base_name, tag):
    """Create display label like 'CMJ (Single Leg - Right)' from base name + tag."""
    if tag is None:
        return base_name
    return f"{base_name} [{tag}]"


def make_found_test_key(base_name, tag):
    """Create dict key for found_tests. Tag-aware to separate variants."""
    if tag is None:
        return base_name
    return f"{base_name}__{tag}"


# ==============================================================================
# PLOTTING FUNCTIONS
# ==============================================================================

def _is_single_leg_tag(tag):
    """Check if a tag/variant indicates a single-leg test."""
    if tag is None:
        return False
    tag_lower = tag.lower()
    return any(kw in tag_lower for kw in ['single side', 'single leg', 'sl -', 'sl-',
                                           '- left', '- right', '-left', '-right'])


def create_force_time_plot(ax_main, ax_asym, force_df, system_weight, title, is_single_leg=False):
    """
    Create force-time curve with asymmetry subplot below.
    
    Returns mean asymmetry value.
    """
    if force_df is None or len(force_df) == 0:
        ax_main.text(0.5, 0.5, 'Force-time data not available', 
                    ha='center', va='center', fontsize=11, color='gray')
        ax_main.set_title(title, fontsize=11, fontweight='bold')
        ax_asym.axis('off')
        return 0.0
    
    # Find columns (API naming varies)
    cols = force_df.columns.str.lower()
    time_col = force_df.columns[cols.str.contains('time')].tolist()
    left_col = force_df.columns[cols.str.contains('left')].tolist()
    right_col = force_df.columns[cols.str.contains('right')].tolist()
    combined_col = force_df.columns[cols.str.contains('combined|total')].tolist()
    
    # Use first match or positional fallback
    time_col = time_col[0] if time_col else force_df.columns[0]
    left_col = left_col[0] if left_col else (force_df.columns[1] if len(force_df.columns) > 1 else None)
    right_col = right_col[0] if right_col else (force_df.columns[2] if len(force_df.columns) > 2 else None)
    combined_col = combined_col[0] if combined_col else None
    
    time = force_df[time_col].values
    left = force_df[left_col].values if left_col else np.zeros_like(time)
    right = force_df[right_col].values if right_col else np.zeros_like(time)
    combined = force_df[combined_col].values if combined_col else left + right
    
    # Smooth and convert to body weight
    bw = system_weight if system_weight > 0 else 800
    left_bw = smooth_data(left) / bw
    right_bw = smooth_data(right) / bw
    combined_bw = smooth_data(combined) / bw
    
    # === MAIN FORCE-TIME PLOT ===
    ax_main.plot(time, combined_bw, linewidth=1.2, color='gray', 
                linestyle='--', alpha=0.7, label='Combined')
    ax_main.plot(time, left_bw, linewidth=1.5, color=COLOR_GOLD, label='Left')
    ax_main.plot(time, right_bw, linewidth=1.5, color=COLOR_BLACK, label='Right')
    
    ax_main.axhline(y=1.0, color='gray', linestyle=':', linewidth=1, alpha=0.5)
    ax_main.set_ylabel('Force (×BW)', fontsize=10, fontweight='bold')
    ax_main.set_title(title, fontsize=11, fontweight='bold', color=COLOR_TEXT_MAIN)
    ax_main.legend(loc='upper right', fontsize=9)
    ax_main.grid(True, alpha=0.3, linestyle='--')
    ax_main.spines['top'].set_visible(False)
    ax_main.spines['right'].set_visible(False)
    ax_main.set_xlim(time[0], time[-1])
    ax_main.tick_params(labelbottom=False)
    
    # === ASYMMETRY SUBPLOT ===
    # For single-leg tests, asymmetry is meaningless — suppress it
    if is_single_leg:
        ax_asym.axis('off')
        return 0.0
    
    # Contact threshold (10% of body weight)
    contact_mask = combined >= (bw * 0.10)
    
    # Calculate asymmetry: (L - R) / (L + R) * 100
    asymmetry = np.zeros_like(left)
    valid = (left + right) > 0
    asymmetry[valid] = 100 * (left[valid] - right[valid]) / (left[valid] + right[valid])
    
    # Zero out flight phases
    asymmetry[~contact_mask] = 0.0
    asymmetry = smooth_data(asymmetry)
    
    mean_asym = np.mean(asymmetry[contact_mask]) if np.any(contact_mask) else 0.0
    
    # Plot
    ax_asym.plot(time, asymmetry, linewidth=1, color=COLOR_BLACK, alpha=0.8)
    ax_asym.fill_between(time, 0, asymmetry, where=(asymmetry >= 0), 
                         alpha=0.4, color=COLOR_GOLD)
    ax_asym.fill_between(time, 0, asymmetry, where=(asymmetry < 0), 
                         alpha=0.5, color='gray')
    
    # Reference lines
    ax_asym.axhline(y=0, color=COLOR_BLACK, linewidth=0.8)
    ax_asym.axhline(y=10, color='red', linestyle='--', linewidth=0.8, alpha=0.7)
    ax_asym.axhline(y=-10, color='red', linestyle='--', linewidth=0.8, alpha=0.7)
    
    # Mean line
    ax_asym.axhline(mean_asym, color=COLOR_BLACK, linewidth=1.2)
    ax_asym.text(time[0] + (time[-1]-time[0])*0.02, mean_asym + 2, 
                f'Mean: {mean_asym:.1f}%', fontsize=8, color=COLOR_BLACK)
    
    # Labels
    ax_asym.text(0.98, 0.85, 'Left dominant', transform=ax_asym.transAxes, 
                fontsize=8, ha='right', color=COLOR_GOLD)
    ax_asym.text(0.98, 0.15, 'Right dominant', transform=ax_asym.transAxes,
                fontsize=8, ha='right', color='gray')
    
    ax_asym.set_xlabel('Time (s)', fontsize=10, fontweight='bold')
    ax_asym.set_ylabel('Asymmetry (%)', fontsize=9)
    ax_asym.set_ylim(-25, 25)
    ax_asym.set_xlim(time[0], time[-1])
    ax_asym.grid(True, alpha=0.3, linestyle='--')
    ax_asym.spines['top'].set_visible(False)
    ax_asym.spines['right'].set_visible(False)
    
    return mean_asym


# ==============================================================================
# PAGE GENERATORS
# ==============================================================================

def draw_header(fig, y_pos, title):
    """Draw gold section header bar."""
    ax = fig.add_axes([0.10, y_pos, 0.80, 0.04])
    ax.axis('off')
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_GOLD, edgecolor='none'))
    ax.text(0.02, 0.5, title, fontsize=12, fontweight='bold', 
            color=COLOR_BLACK, va='center', transform=ax.transAxes)


def add_page_number(fig, num):
    """Add page number bottom right."""
    ax = fig.add_axes([0.10, 0.02, 0.80, 0.02])
    ax.axis('off')
    ax.text(1.0, 0.5, str(num), ha='right', fontsize=9, color='#888888')


def add_footer(fig):
    """Add company footer."""
    ax = fig.add_axes([0.10, 0.04, 0.80, 0.02])
    ax.axis('off')
    ax.text(0.5, 0.5, 'Move, Measure, Analyze LLC', ha='center', fontsize=8, color='#888888')


def create_cover_page(pdf, athlete_name, body_weight_lb, test_date, page_num, date_range=None):
    """Create title page with logo. date_range is an optional (first_date_str, last_date_str) tuple."""
    fig = plt.figure(figsize=(8.5, 11))
    
    # ==========================================================================
    # LOGO POSITION ADJUSTMENT
    # Format: [left, bottom, width, height] - all values are 0-1 (fraction of page)
    # - left: 0.30 means 30% from left edge
    # - bottom: 0.56 means 56% from bottom (lower = move down, higher = move up)
    # - width: 0.40 means logo takes 40% of page width
    # - height: 0.20 means logo takes 20% of page height
    # ==========================================================================
    LOGO_POSITION = [0.30, 0.56, 0.40, 0.20]  # Lowered from 0.62 to 0.56
    
    logo_path = os.path.join(SCRIPT_DIR, 'Logo_Transparent.png')
    if os.path.exists(logo_path):
        try:
            from matplotlib import image as mpimg
            logo = mpimg.imread(logo_path)
            ax_logo = fig.add_axes(LOGO_POSITION)
            ax_logo.imshow(logo)
            ax_logo.axis('off')
        except Exception as e:
            print(f"  Note: Could not load logo: {e}")
    
    fig.text(0.5, 0.48, 'Move, Measure, Analyze LLC', ha='center', 
             fontsize=22, fontweight='bold', color=COLOR_BLACK)
    fig.text(0.5, 0.42, 'FORCE PLATE ASSESSMENT', ha='center', 
             fontsize=16, color=COLOR_BLACK)
    
    # Gold divider
    ax_div = fig.add_axes([0.25, 0.38, 0.50, 0.008])
    ax_div.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_GOLD))
    ax_div.axis('off')
    
    fig.text(0.5, 0.32, athlete_name, ha='center', fontsize=24, fontweight='bold')
    fig.text(0.5, 0.27, f'Body Weight: {body_weight_lb:.1f} lbs', ha='center', fontsize=12)
    # Show date range if multi-session, otherwise single date
    if date_range and date_range[0] != date_range[1]:
        fig.text(0.5, 0.23, f'Assessment Dates: {date_range[0]} — {date_range[1]}', ha='center', fontsize=12)
    else:
        fig.text(0.5, 0.23, f'Assessment Date: {test_date}', ha='center', fontsize=12)
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)


def create_metrics_summary_page(pdf, found_tests, page_num, ai_summary=None):
    """Create a dashboard-style performance metrics summary with AI-generated analysis.
    
    Args:
        ai_summary: Optional AI-generated summary text for the results
    """
    fig = plt.figure(figsize=(8.5, 11))
    
    # Header
    draw_header(fig, 0.92, 'PERFORMANCE DASHBOARD')
    
    # Extract key metrics from all tests (prefer untagged/bilateral variants for dashboard)
    cmj_height = None
    sj_height = None
    cmj_rsi = None
    mr_rsi = None
    dl_peak_force = None
    
    for test_key, test_data in found_tests.items():
        row = test_data.get('row')
        if row is None:
            continue
        
        # Use base_name if available (new structure), else fall back to key
        base_name = test_data.get('base_name', test_key)
        tag = test_data.get('tag')
        
        # Skip single-leg variants for dashboard tiles (EUR needs bilateral only)
        if tag and ('single side' in tag.lower() or 'single leg' in tag.lower() or 'left' in tag.lower() or 'right' in tag.lower()):
            continue
        
        def safe_get(col):
            if col in row.index and pd.notna(row[col]):
                return row[col]
            return None
        
        if 'Countermovement Jump' in base_name and 'Rebound' not in base_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            cmj_height = jh
            cmj_rsi = safe_get('mrsi')
        
        elif 'Squat Jump' in base_name:
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            sj_height = jh
        
        elif 'Multi Rebound' in base_name:
            mr_rsi = safe_get('rsi') or safe_get('avg_rsi')
        
        elif 'Drop Landing' in base_name:
            pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
            sw = safe_get('system_weight_n')
            if pf and sw:
                dl_peak_force = pf / sw
    
    # Calculate EUR
    eur = None
    if cmj_height and sj_height and sj_height > 0:
        eur = cmj_height / sj_height
    
    # === METRIC TILES - Smaller, higher on page, proper margins ===
    tile_y_top = 0.78
    tile_height = 0.10
    tile_width = 0.26
    gap = 0.03
    x_start = 0.08  # Align with summary box margins
    
    def draw_tile(x, y, w, h, title, value, unit, color=COLOR_GOLD):
        ax = fig.add_axes([x, y, w, h])
        ax.axis('off')
        ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_LIGHT_GRAY, 
                                    edgecolor='#cccccc', linewidth=1.5, transform=ax.transAxes))
        ax.add_patch(plt.Rectangle((0, 0.75), 1, 0.25, facecolor=color, transform=ax.transAxes))
        ax.text(0.5, 0.875, title, ha='center', va='center', fontsize=7, fontweight='bold',
               color=COLOR_BLACK if color == COLOR_GOLD else COLOR_WHITE, transform=ax.transAxes)
        ax.text(0.5, 0.45, value, ha='center', va='center', fontsize=16, fontweight='bold',
               color=COLOR_BLACK, transform=ax.transAxes)
        ax.text(0.5, 0.15, unit, ha='center', va='center', fontsize=7,
               color=COLOR_TEXT_MAIN, transform=ax.transAxes)
    
    # TOP ROW
    if cmj_height:
        inches = cmj_height / 2.54
        draw_tile(x_start, tile_y_top, tile_width, tile_height, 
                 'CMJ HEIGHT', f'{cmj_height:.1f} / {inches:.1f}', 'cm / inches')
    
    if cmj_rsi:
        draw_tile(x_start + tile_width + gap, tile_y_top, tile_width, tile_height,
                 'RSImod', f'{cmj_rsi:.2f}', 'explosive index')
    
    if eur:
        eur_color = COLOR_GREEN if eur >= 1.05 else (COLOR_GOLD if eur >= 0.95 else COLOR_RED)
        draw_tile(x_start + 2*(tile_width + gap), tile_y_top, tile_width, tile_height,
                 'EUR (CMJ÷SJ)', f'{eur:.2f}', 'elastic energy', eur_color)
    
    # SECOND ROW
    tile_y_2 = 0.65
    
    if sj_height:
        inches = sj_height / 2.54
        draw_tile(x_start, tile_y_2, tile_width, tile_height,
                 'SJ HEIGHT', f'{sj_height:.1f} / {inches:.1f}', 'cm / inches')
    
    if mr_rsi:
        draw_tile(x_start + tile_width + gap, tile_y_2, tile_width, tile_height,
                 'REACTIVE RSI', f'{mr_rsi:.2f}', 'rebound ability')
    
    if dl_peak_force:
        draw_tile(x_start + 2*(tile_width + gap), tile_y_2, tile_width, tile_height,
                 'LANDING FORCE', f'{dl_peak_force:.2f}', '×BW (peak)')
    
    # === AI SUMMARY SECTION ===
    summary_top = 0.58
    summary_bottom = 0.18
    
    ax_summary = fig.add_axes([0.08, summary_bottom, 0.84, summary_top - summary_bottom])
    ax_summary.axis('off')
    
    ax_summary.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_WHITE, 
                                        edgecolor='#cccccc', linewidth=1.5, transform=ax_summary.transAxes))
    
    ax_summary.add_patch(plt.Rectangle((0, 0.92), 1, 0.08, facecolor=COLOR_GOLD, transform=ax_summary.transAxes))
    ax_summary.text(0.5, 0.96, 'ASSESSMENT SUMMARY', ha='center', va='center',
                   fontsize=10, fontweight='bold', color=COLOR_BLACK, transform=ax_summary.transAxes)
    
    if ai_summary:
        wrapper = textwrap.TextWrapper(width=100)
        lines = wrapper.fill(ai_summary).split('\n')
        y_text = 0.88
        for i, line in enumerate(lines[:14]):
            ax_summary.text(0.02, y_text - (i * 0.058), line, fontsize=8,
                           color=COLOR_TEXT_MAIN, transform=ax_summary.transAxes)
    else:
        # No AI summary - leave the box empty for editable form field overlay
        pass
    
    # Key at bottom
    fig.text(0.08, 0.13, 'KEY:', fontsize=7, fontweight='bold', color=COLOR_BLACK)
    abbrev = 'CMJ=Countermovement Jump • SJ=Squat Jump • EUR=Eccentric Utilization Ratio • RSI=Reactive Strength Index • BW=Body Weight'
    fig.text(0.13, 0.13, abbrev, fontsize=6, color=COLOR_TEXT_MAIN)
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)


def create_test_page(pdf, test_name, force_data, system_weight, page_num, test_row=None, custom_interpretation=None, sessions=None, tag=None):
    """Create a page with force-time graph, asymmetry subplot, description, and metric tiles.
    
    Args:
        custom_interpretation: Optional string to replace default description (for RAG integration)
        sessions: Optional list of session dicts (used by trend page, not here)
        tag: Optional tag string for display in header
    """
    fig = plt.figure(figsize=(8.5, 11))
    
    # Detect single-leg test
    is_single_leg = _is_single_leg_tag(tag)
    
    # Header - include tag if present
    header_label = test_name.upper()
    if tag:
        header_label = f"{header_label} — {tag.upper()}"
    draw_header(fig, 0.92, f'{header_label} ASSESSMENT')
    
    # Fixed layout (no sparkline strip)
    ft_bottom, ft_height = 0.58, 0.28
    asym_bottom, asym_height = 0.46, 0.10
    desc_bottom, desc_height = 0.08, 0.32
    tiles_bottom, tiles_height = 0.08, 0.32
    
    # Main force-time plot
    ax_main = fig.add_axes([0.10, ft_bottom, 0.82, ft_height])
    
    # Asymmetry subplot
    ax_asym = fig.add_axes([0.10, asym_bottom, 0.82, asym_height])
    
    # Create plots
    mean_asym = 0
    if force_data is not None and not force_data.empty:
        mean_asym = create_force_time_plot(ax_main, ax_asym, force_data, system_weight,
                              f'{test_name} Force-Time Curve', is_single_leg=is_single_leg)
    else:
        ax_main.text(0.5, 0.5, 'Force-time data not available', ha='center', va='center',
                    fontsize=12, color='gray', transform=ax_main.transAxes)
        ax_main.set_xticks([])
        ax_main.set_yticks([])
        ax_asym.set_xticks([])
        ax_asym.set_yticks([])
    
    # === BOTTOM SECTION: Description (left) + Metrics (right) ===
    
    # Get test description
    test_info = TEST_DESCRIPTIONS.get(test_name, {
        'title': test_name,
        'what': 'Force plate assessment.',
        'why': 'Measures performance characteristics.',
        'metrics': []
    })
    
    # --- Description Panel (bottom left) ---
    desc_left = 0.06
    desc_width = 0.44

    ax_desc = fig.add_axes([desc_left, desc_bottom, desc_width, desc_height])
    ax_desc.axis('off')

    ax_desc.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_WHITE,
                                     edgecolor='#cccccc', linewidth=1.5,
                                     transform=ax_desc.transAxes))

    header_height = 0.07
    ax_desc.add_patch(plt.Rectangle((0, 1 - header_height), 1, header_height, facecolor=COLOR_GOLD,
                                     transform=ax_desc.transAxes, clip_on=False))
    ax_desc.text(0.5, 1 - header_height/2, 'ABOUT THIS TEST', ha='center', va='center',
                fontsize=9, fontweight='bold', color=COLOR_BLACK, transform=ax_desc.transAxes)

    wrapper = textwrap.TextWrapper(width=65)

    if custom_interpretation:
        y_text = 0.90
        interp_lines = wrapper.fill(custom_interpretation).split('\n')
        for i, line in enumerate(interp_lines[:15]):
            ax_desc.text(0.02, y_text - (i * 0.055), line, fontsize=6.5,
                        color=COLOR_TEXT_MAIN, transform=ax_desc.transAxes)
    # else: intentionally blank — editable PDF form field overlays this area

    # --- Metric Tiles Panel (bottom right) ---
    tiles_left = 0.52
    tiles_width = 0.44

    ax_tiles = fig.add_axes([tiles_left, tiles_bottom, tiles_width, tiles_height])
    ax_tiles.axis('off')

    ax_tiles.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COLOR_WHITE,
                                     edgecolor='#cccccc', linewidth=1.5,
                                     transform=ax_tiles.transAxes))

    ax_tiles.add_patch(plt.Rectangle((0, 1 - header_height), 1, header_height, facecolor=COLOR_GOLD,
                                      transform=ax_tiles.transAxes, clip_on=False))
    ax_tiles.text(0.5, 1 - header_height/2, 'KEY METRICS', ha='center', va='center',
                 fontsize=9, fontweight='bold', color=COLOR_BLACK, transform=ax_tiles.transAxes)

    # Extract metrics
    metrics_to_show = []

    if test_row is not None:
        def safe_get(col):
            if col in test_row.index and pd.notna(test_row[col]):
                return test_row[col]
            return None

        body_weight_n = safe_get('system_weight_n')
        body_weight_kg = body_weight_n / 9.81 if body_weight_n else None
        # Get key metrics based on test type
        if 'Drop Landing' in test_name:
            # Peak Landing Force (relative) - Hawkin uses peak_force_n or peak_landing_force_n
            pf = safe_get('peak_force_n') or safe_get('peak_landing_force_n')
            pf_asym = safe_get('lr_peak_force') or safe_get('lr_peak_landing_force')
            if pf and body_weight_n:
                metrics_to_show.append(('Peak Force', f'{pf/body_weight_n:.2f}', '×BW', pf_asym))
            
            # Avg Impact Force - Hawkin uses avg_impact_force_n or avg_landing_force_n
            aif = safe_get('avg_impact_force_n') or safe_get('avg_landing_force_n')
            aif_asym = safe_get('lr_avg_impact_force') or safe_get('lr_avg_landing_force')
            if aif and body_weight_n:
                metrics_to_show.append(('Avg Impact', f'{aif/body_weight_n:.2f}', '×BW', aif_asym))
            
            # Impact RFD
            irfd = safe_get('impact_rfd_n_s') or safe_get('landing_rfd_n_s')
            irfd_asym = safe_get('lr_impact_rfd') or safe_get('lr_landing_rfd')
            if irfd and body_weight_kg:
                metrics_to_show.append(('Impact RFD', f'{irfd/body_weight_kg:.0f}', 'N/s/kg', irfd_asym))
            
            # Time to Stabilization
            tts = safe_get('time_to_stabilization_ms')
            if tts:
                metrics_to_show.append(('Time to Stable', f'{tts/1000:.2f}', 's', None))
        
        elif 'Countermovement Jump' in test_name and 'Rebound' not in test_name:
            # Jump Height (no asymmetry)
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                jh_in = jh / 2.54
                metrics_to_show.append(('Jump Height', f'{jh:.1f} / {jh_in:.1f}', 'cm / in', None))
            
            # RSImod (no asymmetry)
            mrsi = safe_get('mrsi')
            if mrsi:
                metrics_to_show.append(('RSImod', f'{mrsi:.2f}', '', None))
            
            # Peak Propulsive Force with asymmetry
            ppf = safe_get('peak_propulsive_force_n')
            ppf_asym = safe_get('cmj_lr_peak_propulsive_force') or safe_get('lr_peak_propulsive_force')
            if ppf and body_weight_n:
                metrics_to_show.append(('Peak Prop Force', f'{ppf/body_weight_n:.2f}', '×BW', ppf_asym))
            
            # Propulsive Impulse Index (asymmetry)
            imp_asym = safe_get('cmj_lr_propulsive_impulse_index') or safe_get('lr_propulsive_impulse_index')
            if imp_asym is not None:
                metrics_to_show.append(('Prop Impulse', f'{abs(imp_asym):.1f}%', 'asym', imp_asym))
        
        elif 'Squat Jump' in test_name:
            # Jump Height
            jh = safe_get('jump_height_imp_mom_cm')
            if jh is None:
                jh_m = safe_get('jump_height_m')
                if jh_m: jh = jh_m * 100
            if jh:
                jh_in = jh / 2.54
                metrics_to_show.append(('Jump Height', f'{jh:.1f} / {jh_in:.1f}', 'cm / in', None))
            
            # Peak Power
            pp = safe_get('peak_relative_propulsive_power_w_kg')
            if pp:
                metrics_to_show.append(('Peak Power', f'{pp:.1f}', 'W/kg', None))
            
            # Peak Propulsive Force with asymmetry
            ppf = safe_get('peak_propulsive_force_n')
            ppf_asym = safe_get('sj_lr_peak_propulsive_force') or safe_get('lr_peak_propulsive_force')
            if ppf and body_weight_n:
                metrics_to_show.append(('Peak Prop Force', f'{ppf/body_weight_n:.2f}', '×BW', ppf_asym))
            
            # Propulsive Impulse Index with asymmetry
            imp_asym = safe_get('sj_lr_propulsive_impulse_index') or safe_get('lr_propulsive_impulse_index')
            if imp_asym is not None:
                metrics_to_show.append(('Impulse Asym', f'{abs(imp_asym):.1f}%', 'L' if imp_asym > 0 else 'R', imp_asym))
        
        elif 'Multi Rebound' in test_name:
            # RSI (key metric)
            rsi = safe_get('rsi') or safe_get('avg_rsi')
            if rsi:
                metrics_to_show.append(('RSI', f'{rsi:.2f}', '', None))
            
            # Avg Jump Height (with inches)
            ajh = safe_get('avg_jump_height_m')
            if ajh:
                ajh_cm = ajh * 100
                ajh_in = ajh_cm / 2.54
                metrics_to_show.append(('Avg Height', f'{ajh_cm:.1f} / {ajh_in:.1f}', 'cm / in', None))
            
            # Avg Contact Time
            act = safe_get('avg_contact_time_s')
            if act:
                metrics_to_show.append(('Avg Contact', f'{act*1000:.0f}', 'ms', None))
            
            # Peak Force asymmetry
            pf_asym = safe_get('lr_peak_force')
            if pf_asym is not None:
                metrics_to_show.append(('Peak Force Asym', f'{abs(pf_asym):.1f}%', 'L' if pf_asym > 0 else 'R', pf_asym))
        
        elif 'Bodyweight Squats' in test_name:
            # For BW Squats, we show the asymmetry metrics from API columns
            avg_asym = safe_get('lr_avg_force')
            peak_asym = safe_get('lr_peak_force')
            
            # If API has asymmetry values, show them
            if avg_asym is not None:
                asym_color_text = 'L' if avg_asym > 0 else 'R' if avg_asym < 0 else ''
                metrics_to_show.append(('Avg Force Asym', f'{abs(avg_asym):.1f}%', asym_color_text, avg_asym))
            
            if peak_asym is not None:
                asym_color_text = 'L' if peak_asym > 0 else 'R' if peak_asym < 0 else ''
                metrics_to_show.append(('Peak Force Asym', f'{abs(peak_asym):.1f}%', asym_color_text, peak_asym))
            
            # If no API asymmetry, use mean_asym from force-time curve
            if not metrics_to_show and mean_asym != 0:
                asym_color_text = 'L' if mean_asym > 0 else 'R' if mean_asym < 0 else ''
                metrics_to_show.append(('Mean Asymmetry', f'{abs(mean_asym):.1f}%', asym_color_text, mean_asym))
        
        elif 'CMJ Rebound' in test_name:
            cmj_jh = safe_get('cmj_jump_height_m')
            if cmj_jh:
                jh_in = (cmj_jh * 100) / 2.54
                metrics_to_show.append(('CMJ Height', f'{cmj_jh*100:.1f} / {jh_in:.1f}', 'cm / in', None))
            
            reb_jh = safe_get('rebound_jump_height_m')
            if reb_jh:
                reb_in = (reb_jh * 100) / 2.54
                metrics_to_show.append(('Rebound Height', f'{reb_jh*100:.1f} / {reb_in:.1f}', 'cm / in', None))
            
            rsi = safe_get('rsi') or safe_get('rebound_rsi')
            if rsi:
                metrics_to_show.append(('RSI', f'{rsi:.2f}', '', None))
    
    # Draw metrics
    if not metrics_to_show:
        ax_tiles.text(0.5, 0.45, 'Metrics not available', ha='center', va='center',
                     fontsize=9, color='gray', transform=ax_tiles.transAxes)
    else:
        # Layout metrics vertically with adaptive spacing
        # Metrics with benchmark bars (RSImod, RSI) need more vertical space
        y_start = 0.84
        
        # Pre-calculate which metrics have benchmark bars
        has_bar = []
        for metric_data in metrics_to_show[:4]:
            mname = metric_data[0]
            if mname == 'RSImod':
                has_bar.append('double')  # Two bars (M+F)
            elif mname == 'RSI' and ('Multi Rebound' in test_name or 'CMJ Rebound' in test_name):
                has_bar.append('single')  # One bar
            else:
                has_bar.append(False)
        
        # Calculate y positions: double-bar metrics get 0.32, single-bar get 0.25, others get 0.16
        y_positions = []
        y_cur = y_start
        for i in range(len(metrics_to_show[:4])):
            y_positions.append(y_cur)
            if has_bar[i] == 'double':
                y_cur -= 0.32  # Two bars + labels + citation
            elif has_bar[i] == 'single':
                y_cur -= 0.25  # One bar + reference text
            else:
                y_cur -= 0.16  # Normal spacing
        
        for i, metric_data in enumerate(metrics_to_show[:4]):  # Max 4 metrics
            name, value, unit, asym = metric_data
            y_pos = y_positions[i]
            
            # Metric name
            ax_tiles.text(0.04, y_pos, name, fontsize=7, fontweight='bold',
                         color=COLOR_TEXT_MAIN, transform=ax_tiles.transAxes)
            
            # Value and unit
            display_val = f'{value} {unit}'.strip()
            ax_tiles.text(0.5, y_pos, display_val, fontsize=12, fontweight='bold',
                         ha='center', color=COLOR_BLACK, transform=ax_tiles.transAxes)
            
            # RSImod benchmark scale (for CMJ) — Two bars: Male + Female
            if name == 'RSImod':
                bar_height = 0.028
                bar_left = 0.04
                bar_width = 0.92
                scale_max = 0.70
                
                # --- Male bar ---
                m_bar_y = y_pos - 0.075
                m_low = 0.34 / scale_max
                m_high = 0.49 / scale_max
                
                # "Males" label
                ax_tiles.text(bar_left, m_bar_y + bar_height + 0.008, 'Males', 
                             fontsize=5.5, fontweight='bold', color='#444', transform=ax_tiles.transAxes)
                
                # Draw male colored zones
                ax_tiles.add_patch(plt.Rectangle((bar_left, m_bar_y), bar_width * m_low, bar_height,
                    facecolor='#e74c3c', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * m_low, m_bar_y), 
                    bar_width * (m_high - m_low), bar_height,
                    facecolor='#f1c40f', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * m_high, m_bar_y), 
                    bar_width * (1 - m_high), bar_height,
                    facecolor='#2ecc71', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left, m_bar_y), bar_width, bar_height,
                    facecolor='none', edgecolor='#666666', linewidth=0.5, transform=ax_tiles.transAxes))
                
                # Male zone value labels inside bar
                ax_tiles.text(bar_left + bar_width * m_low * 0.5, m_bar_y + bar_height/2, 
                             '<0.34', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (m_low + m_high) / 2, m_bar_y + bar_height/2, 
                             '0.35-0.48', fontsize=4.5, ha='center', va='center', color='#333', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (m_high + 1) / 2, m_bar_y + bar_height/2, 
                             '>0.49', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                
                # Male athlete marker
                try:
                    val = float(value)
                    marker_x = bar_left + bar_width * min(val / scale_max, 1.0)
                    ax_tiles.plot([marker_x], [m_bar_y + bar_height + 0.005], marker='v', 
                                color=COLOR_BLACK, markersize=5, transform=ax_tiles.transAxes, clip_on=False)
                except:
                    pass
                
                # Male zone labels below bar
                ax_tiles.text(bar_left, m_bar_y - 0.012, 'Low', fontsize=4.5, color='#e74c3c',
                             transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * 0.5, m_bar_y - 0.012, 'Ave', fontsize=4.5, 
                             ha='center', color='#666', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width, m_bar_y - 0.012, 'High', fontsize=4.5, 
                             ha='right', color='#2ecc71', transform=ax_tiles.transAxes)
                
                # --- Female bar ---
                f_bar_y = m_bar_y - 0.075
                f_low = 0.24 / scale_max
                f_high = 0.38 / scale_max
                
                # "Females" label
                ax_tiles.text(bar_left, f_bar_y + bar_height + 0.008, 'Females', 
                             fontsize=5.5, fontweight='bold', color='#444', transform=ax_tiles.transAxes)
                
                # Draw female colored zones
                ax_tiles.add_patch(plt.Rectangle((bar_left, f_bar_y), bar_width * f_low, bar_height,
                    facecolor='#e74c3c', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * f_low, f_bar_y), 
                    bar_width * (f_high - f_low), bar_height,
                    facecolor='#f1c40f', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * f_high, f_bar_y), 
                    bar_width * (1 - f_high), bar_height,
                    facecolor='#2ecc71', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left, f_bar_y), bar_width, bar_height,
                    facecolor='none', edgecolor='#666666', linewidth=0.5, transform=ax_tiles.transAxes))
                
                # Female zone value labels inside bar
                ax_tiles.text(bar_left + bar_width * f_low * 0.5, f_bar_y + bar_height/2, 
                             '<0.24', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (f_low + f_high) / 2, f_bar_y + bar_height/2, 
                             '0.25-0.37', fontsize=4.5, ha='center', va='center', color='#333', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (f_high + 1) / 2, f_bar_y + bar_height/2, 
                             '>0.38', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                
                # Female athlete marker
                try:
                    val = float(value)
                    marker_x = bar_left + bar_width * min(val / scale_max, 1.0)
                    ax_tiles.plot([marker_x], [f_bar_y + bar_height + 0.005], marker='v', 
                                color=COLOR_BLACK, markersize=5, transform=ax_tiles.transAxes, clip_on=False)
                except:
                    pass
                
                # Female zone labels below bar
                ax_tiles.text(bar_left, f_bar_y - 0.012, 'Low', fontsize=4.5, color='#e74c3c',
                             transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * 0.5, f_bar_y - 0.012, 'Ave', fontsize=4.5, 
                             ha='center', color='#666', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width, f_bar_y - 0.012, 'High', fontsize=4.5, 
                             ha='right', color='#2ecc71', transform=ax_tiles.transAxes)
                
                # Citation
                ax_tiles.text(bar_left, f_bar_y - 0.03, 
                             'Sole et al. 2018 — NCAA Div 1',
                             fontsize=4, color='#999', fontstyle='italic', transform=ax_tiles.transAxes)
            
            # RSI benchmark scale (for Multi Rebound and CMJ Rebound)
            elif name == 'RSI' and ('Multi Rebound' in test_name or 'CMJ Rebound' in test_name):
                bar_y = y_pos - 0.09
                bar_height = 0.035
                bar_left = 0.04
                bar_width = 0.92
                
                # Scale: 0 to 3.0
                scale_max = 3.0
                
                # Zone boundaries
                # Low: 0-1.1, Moderate: 1.1-1.8, High: 1.8+
                low_end = 1.1 / scale_max
                mod_end = 1.8 / scale_max
                
                # Draw zones
                ax_tiles.add_patch(plt.Rectangle((bar_left, bar_y), bar_width * low_end, bar_height,
                    facecolor='#e74c3c', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * low_end, bar_y), 
                    bar_width * (mod_end - low_end), bar_height,
                    facecolor='#f1c40f', transform=ax_tiles.transAxes))
                ax_tiles.add_patch(plt.Rectangle((bar_left + bar_width * mod_end, bar_y), 
                    bar_width * (1 - mod_end), bar_height,
                    facecolor='#2ecc71', transform=ax_tiles.transAxes))
                
                ax_tiles.add_patch(plt.Rectangle((bar_left, bar_y), bar_width, bar_height,
                    facecolor='none', edgecolor='#666666', linewidth=0.5, transform=ax_tiles.transAxes))
                
                # Zone value labels inside bar
                ax_tiles.text(bar_left + bar_width * low_end * 0.5, bar_y + bar_height/2, 
                             '<1.1', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (low_end + mod_end) / 2, bar_y + bar_height/2, 
                             '1.1-1.8', fontsize=4.5, ha='center', va='center', color='#333', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * (mod_end + 1) / 2, bar_y + bar_height/2, 
                             '>1.8', fontsize=4.5, ha='center', va='center', color='white', 
                             fontweight='bold', transform=ax_tiles.transAxes)
                
                # Athlete marker
                try:
                    val = float(value)
                    marker_x = bar_left + bar_width * min(val / scale_max, 1.0)
                    ax_tiles.plot([marker_x], [bar_y + bar_height + 0.01], marker='v', 
                                color=COLOR_BLACK, markersize=6, transform=ax_tiles.transAxes, clip_on=False)
                except:
                    pass
                
                # Zone labels
                ax_tiles.text(bar_left, bar_y - 0.015, 'Low', fontsize=5, color='#e74c3c',
                             transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width * 0.5, bar_y - 0.015, 'Moderate', fontsize=5, 
                             ha='center', color='#666', transform=ax_tiles.transAxes)
                ax_tiles.text(bar_left + bar_width, bar_y - 0.015, 'High', fontsize=5, 
                             ha='right', color='#2ecc71', transform=ax_tiles.transAxes)
                
                # Citation
                ref_y = bar_y - 0.04
                ax_tiles.text(bar_left, ref_y, 
                             '(Adapted from Strength by Numbers)',
                             fontsize=4, color='#999', fontstyle='italic', transform=ax_tiles.transAxes)
            
            # Asymmetry bar if available (for other metrics)
            elif asym is not None and not is_single_leg:
                bar_y = y_pos - 0.08
                bar_height = 0.04
                bar_left = 0.04
                bar_width = 0.92
                
                # Background
                ax_tiles.add_patch(plt.Rectangle((bar_left, bar_y), bar_width, bar_height,
                                                 facecolor=COLOR_LIGHT_GRAY, edgecolor='#999999',
                                                 linewidth=0.5, transform=ax_tiles.transAxes))
                
                # Calculate split (50% = balanced)
                left_pct = 50 + (asym / 2)  # asym is already a percentage
                left_pct = max(10, min(90, left_pct))  # Clamp to visible range
                
                # Left portion (gold)
                left_bar_width = bar_width * (left_pct / 100)
                ax_tiles.add_patch(plt.Rectangle((bar_left, bar_y), left_bar_width, bar_height,
                                                 facecolor=COLOR_GOLD, transform=ax_tiles.transAxes))
                
                # Right portion (black)
                ax_tiles.add_patch(plt.Rectangle((bar_left + left_bar_width, bar_y), 
                                                 bar_width - left_bar_width, bar_height,
                                                 facecolor=COLOR_BLACK, transform=ax_tiles.transAxes))
                
                # Center line (50% mark)
                center_x = bar_left + bar_width * 0.5
                ax_tiles.plot([center_x, center_x], [bar_y - 0.01, bar_y + bar_height + 0.01], 
                             color='white', linestyle=':', linewidth=1.5, transform=ax_tiles.transAxes)
                
                # Asymmetry label
                asym_color = COLOR_GREEN if abs(asym) < 10 else (COLOR_GOLD if abs(asym) < 15 else COLOR_RED)
                side = 'L' if asym > 0 else 'R' if asym < 0 else ''
                ax_tiles.text(0.96, y_pos, f'{abs(asym):.0f}%{side}', fontsize=7, fontweight='bold',
                             ha='right', color=asym_color, transform=ax_tiles.transAxes)
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)
    

def create_trend_page(pdf, test_name, sessions, page_num, tag=None):
    """Create a trend page with compact graphs (top) and data table (bottom).
    
    Graphs show last 5 sessions with SD shading where n>1.
    Table shows ALL sessions with date, n, mean ± SD for each metric.
    """
    if not sessions or len(sessions) < 2:
        return
    
    # Collect all metric names across sessions
    all_metric_names = []
    seen = set()
    for s in sessions:
        if s.get('avg_metrics'):
            for mn in s['avg_metrics'].keys():
                if mn not in seen:
                    all_metric_names.append(mn)
                    seen.add(mn)
    
    if not all_metric_names:
        return
    
    n_metrics = min(len(all_metric_names), 4)
    metric_names = all_metric_names[:n_metrics]
    
    # Graphs show last 5 sessions; table shows all
    graph_sessions = sessions[-5:] if len(sessions) > 5 else sessions
    
    fig = plt.figure(figsize=(8.5, 11))
    
    # Header
    header_label = test_name.upper()
    if tag:
        header_label = f"{header_label} — {tag.upper()}"
    draw_header(fig, 0.92, f'{header_label} TRENDS')
    
    # Subtitle
    n_sess_total = len(sessions)
    n_sess_graph = len(graph_sessions)
    date_range = f"{sessions[0]['date_str']} — {sessions[-1]['date_str']}"
    subtitle = f'{n_sess_total} sessions  •  {date_range}'
    if n_sess_graph < n_sess_total:
        subtitle += f'  (graphs show last {n_sess_graph})'
    fig.text(0.10, 0.895, subtitle,
             fontsize=9, fontstyle='italic', color=COLOR_TEXT_MAIN)
    
    # === TOP HALF: Compact graphs ===
    graph_top = 0.87
    graph_bottom = 0.48
    available = graph_top - graph_bottom
    gap = 0.035
    subplot_height = (available - gap * (n_metrics - 1)) / n_metrics
    
    for mi, mname in enumerate(metric_names):
        y_pos = graph_top - mi * (subplot_height + gap) - subplot_height
        ax = fig.add_axes([0.14, y_pos, 0.78, subplot_height])
        
        dates = []
        avg_vals = []
        sd_vals = []
        trial_scatter_x = []
        trial_scatter_y = []
        n_trials_list = []
        
        for si, s in enumerate(graph_sessions):
            m = s.get('avg_metrics', {}).get(mname)
            if m is not None:
                dates.append(s['date'])
                avg_vals.append(m['value'])
                n_trials_list.append(m.get('n_trials', 1))
                
                vals = m.get('values', [])
                if len(vals) > 1:
                    sd_vals.append(np.std(vals, ddof=1))
                else:
                    sd_vals.append(0.0)
                
                for tv in vals:
                    trial_scatter_x.append(si)
                    trial_scatter_y.append(tv)
        
        if len(dates) < 2:
            ax.text(0.5, 0.5, f'{mname}: insufficient data', ha='center', va='center',
                   fontsize=8, color='#aaa', transform=ax.transAxes)
            ax.axis('off')
            continue
        
        x_idx = list(range(len(dates)))
        avg_arr = np.array(avg_vals)
        sd_arr = np.array(sd_vals)
        
        # SD shading band
        if np.any(sd_arr > 0):
            ax.fill_between(x_idx, avg_arr - sd_arr, avg_arr + sd_arr,
                           alpha=0.15, color=COLOR_GOLD, linewidth=0, label='±1 SD')
        
        # Individual trials
        if trial_scatter_x:
            ax.scatter(trial_scatter_x, trial_scatter_y, color=COLOR_GOLD, alpha=0.35,
                      s=15, zorder=2, label='Individual trials')
        
        # Session means
        ax.plot(x_idx, avg_vals, '-o', color=COLOR_BLACK, linewidth=1.8,
               markersize=5, markerfacecolor=COLOR_GOLD, markeredgecolor=COLOR_BLACK,
               markeredgewidth=1, zorder=4, label='Session mean')
        
        # Annotate each point
        for xi, (val, nt) in enumerate(zip(avg_vals, n_trials_list)):
            label_str = f'{val:.1f}' if nt == 1 else f'{val:.1f}\n(n={nt})'
            ax.annotate(label_str, (xi, val), textcoords='offset points',
                       xytext=(0, 8), ha='center', fontsize=6, color=COLOR_TEXT_MAIN)
        
        # Trend line (3+ points)
        if len(x_idx) >= 3:
            z = np.polyfit(x_idx, avg_vals, 1)
            p = np.poly1d(z)
            ax.plot(x_idx, p(x_idx), '--', color='#cccccc', linewidth=1, zorder=1)
        
        # Unit string
        unit_str = ''
        for s in graph_sessions:
            m = s.get('avg_metrics', {}).get(mname)
            if m:
                unit_str = m.get('unit', '')
                break
        
        # % change (computed from graphed sessions)
        delta = avg_vals[-1] - avg_vals[0]
        pct_change = (delta / avg_vals[0] * 100) if avg_vals[0] != 0 else 0
        arrow = '↑' if delta > 0 else ('↓' if delta < 0 else '→')
        trend_color = COLOR_GREEN if delta > 0 else (COLOR_RED if delta < 0 else '#888')
        
        ax.set_ylabel(f'{mname}\n({unit_str})', fontsize=7, fontweight='bold')
        ax.set_title(f'{mname}  {arrow} {abs(pct_change):.1f}%',
                    fontsize=8, fontweight='bold', color=trend_color, loc='left', pad=3)
        
        ax.set_xticks(x_idx)
        ax.tick_params(axis='y', labelsize=6)
        ax.grid(True, alpha=0.3, linestyle='--')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        # X-axis labels ONLY on the bottom subplot
        is_last = (mi == n_metrics - 1)
        if is_last:
            ax.set_xticklabels([d.strftime('%m/%d/%y') for d in dates], fontsize=7)
        else:
            ax.tick_params(labelbottom=False)
        
        if mi == 0:
            ax.legend(fontsize=5.5, loc='upper right', ncol=3)
    
    # === BOTTOM HALF: Data table (ALL sessions) ===
    # Compute table height based on number of rows
    n_data_rows = len(sessions) + 1  # +1 for Δ Change row
    row_height = 0.04
    table_height = min(0.37, (n_data_rows + 1) * row_height + 0.03)  # +1 for header
    table_bottom = 0.46 - table_height
    
    ax_table = fig.add_axes([0.06, max(table_bottom, 0.07), 0.88, table_height])
    ax_table.axis('off')
    
    col_labels = ['Date', 'n'] + [f'{mn}' for mn in metric_names]
    
    table_data = []
    for s in sessions:
        date_str = s['date'].strftime('%m/%d/%y')
        n_trials = '—'
        for mn in metric_names:
            m = s.get('avg_metrics', {}).get(mn)
            if m is not None:
                n_trials = str(m.get('n_trials', 1))
                break
        
        row = [date_str, n_trials]
        for mn in metric_names:
            m = s.get('avg_metrics', {}).get(mn)
            if m is not None:
                val = m['value']
                vals = m.get('values', [])
                if len(vals) > 1:
                    sd = np.std(vals, ddof=1)
                    row.append(f'{val:.1f} ± {sd:.1f}')
                else:
                    row.append(f'{val:.1f}')
            else:
                row.append('—')
        table_data.append(row)
    
    # Δ Change row (first to last session, all sessions)
    change_row = ['Δ Change', '']
    for mn in metric_names:
        first_val = None
        last_val = None
        for s in sessions:
            m = s.get('avg_metrics', {}).get(mn)
            if m is not None:
                if first_val is None:
                    first_val = m['value']
                last_val = m['value']
        if first_val is not None and last_val is not None and first_val != 0:
            pct = (last_val - first_val) / first_val * 100
            arr = '↑' if pct > 0 else ('↓' if pct < 0 else '→')
            change_row.append(f'{arr} {abs(pct):.1f}%')
        else:
            change_row.append('—')
    table_data.append(change_row)
    
    # Render table
    table = ax_table.table(
        cellText=table_data,
        colLabels=col_labels,
        loc='upper center',
        cellLoc='center'
    )
    
    table.auto_set_font_size(False)
    
    col_widths = [0.10, 0.05] + [0.85 / max(n_metrics, 1)] * n_metrics
    total_w = sum(col_widths)
    col_widths = [w / total_w for w in col_widths]
    
    # Scale row heights to fit
    actual_row_h = min(0.065, 1.0 / (n_data_rows + 2))
    
    for (row_idx, col_idx), cell in table.get_celld().items():
        cell.set_edgecolor('#cccccc')
        cell.set_linewidth(0.5)
        
        if col_idx < len(col_widths):
            cell.set_width(col_widths[col_idx])
        
        cell.set_height(actual_row_h)
        
        if row_idx == 0:
            # Header row
            cell.set_facecolor(COLOR_GOLD)
            cell.set_text_props(fontweight='bold', fontsize=7, color=COLOR_BLACK)
        elif row_idx == len(table_data):
            # Δ Change row
            cell.set_facecolor('#f0f0f0')
            cell.set_text_props(fontweight='bold', fontsize=7)
            if col_idx >= 2:
                text = cell.get_text().get_text()
                if '↑' in text:
                    cell.set_text_props(fontweight='bold', fontsize=7, color=COLOR_GREEN)
                elif '↓' in text:
                    cell.set_text_props(fontweight='bold', fontsize=7, color=COLOR_RED)
        else:
            cell.set_facecolor(COLOR_WHITE if row_idx % 2 == 1 else '#fafafa')
            cell.set_text_props(fontsize=7)
    
    # Table title
    ax_table.text(0.5, 1.02, 'SESSION DATA  (mean ± SD)', ha='center', va='bottom',
                 fontsize=9, fontweight='bold', color=COLOR_TEXT_MAIN, transform=ax_table.transAxes)
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)


def create_stabilogram_page(pdf, test_name, cop_data, page_num):
    """
    Create a page with left and right leg stabilograms (COP paths).
    Charts only - no metrics table.
    
    cop_data: DataFrame from COP CSV file, or None if not available
    """
    fig = plt.figure(figsize=(8.5, 11))
    
    # Header
    draw_header(fig, 0.92, f'{test_name.upper()} - BALANCE & POSTURAL CONTROL')
    
    # Description
    ax_desc = fig.add_axes([0.10, 0.87, 0.80, 0.04])
    ax_desc.axis('off')
    ax_desc.text(0, 0.5, 'Center of pressure tracking for stability assessment',
                fontsize=9, color=COLOR_TEXT_MAIN, va='center', fontstyle='italic')
    
    # Left leg stabilogram area - adjusted margins (within 0.10 left, more vertical space)
    # Position: [left, bottom, width, height]
    ax_left = fig.add_axes([0.10, 0.15, 0.38, 0.68])
    
    # Right leg stabilogram area - adjusted margins
    ax_right = fig.add_axes([0.55, 0.15, 0.38, 0.68])
    
    left_path, left_area = None, None
    right_path, right_area = None, None
    has_cop_data = False
    
    if cop_data is not None and len(cop_data) > 0:
        # COP CSV columns: Time (s), COP X, COP Y, Left COP X, Left COP Y, Right COP X, Right COP Y
        # Values are in mm - create_stabilogram handles conversion to cm
        
        if 'Left COP X' in cop_data.columns and 'Left COP Y' in cop_data.columns:
            left_x = cop_data['Left COP X'].values
            left_y = cop_data['Left COP Y'].values
            
            left_path, left_area = create_stabilogram(
                ax_left, left_x, left_y, 'Left Leg Balance', COLOR_GOLD
            )
            if left_path is not None:
                has_cop_data = True
        
        if 'Right COP X' in cop_data.columns and 'Right COP Y' in cop_data.columns:
            right_x = cop_data['Right COP X'].values
            right_y = cop_data['Right COP Y'].values
            
            right_path, right_area = create_stabilogram(
                ax_right, right_x, right_y, 'Right Leg Balance', COLOR_BLACK
            )
            if right_path is not None:
                has_cop_data = True
    
    # If no COP data, show informative message
    if not has_cop_data:
        for ax, title in [(ax_left, 'Left Leg Balance'), (ax_right, 'Right Leg Balance')]:
            ax.text(0.5, 0.6, 'COP Data Not Available', ha='center', va='center',
                   fontsize=11, fontweight='bold', color=COLOR_TEXT_MAIN,
                   transform=ax.transAxes)
            ax.text(0.5, 0.4, 'Place COP CSV file in\nthe same folder as this script',
                   ha='center', va='center', fontsize=9, color='gray',
                   transform=ax.transAxes, linespacing=1.5)
            ax.set_title(title, fontsize=10, fontweight='bold')
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.axis('off')
        
        # Add note about CSV workflow
        ax_note = fig.add_axes([0.10, 0.05, 0.80, 0.08])
        ax_note.axis('off')
        note_text = """To generate stabilograms: Export COP data from Hawkin Cloud → Tests → Select test → Export → COP"""
        ax_note.text(0.5, 0.5, note_text, ha='center', va='center', fontsize=8,
                    color=COLOR_TEXT_MAIN,
                    bbox=dict(boxstyle='round', facecolor=COLOR_LIGHT_GRAY, edgecolor='gray', alpha=0.5))
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)


def create_asymmetry_page(pdf, asym_data, page_num):
    """Create asymmetry summary - fits all data on one page."""
    
    if not asym_data:
        fig = plt.figure(figsize=(8.5, 11))
        draw_header(fig, 0.92, 'ASYMMETRY SUMMARY')
        fig.text(0.5, 0.60, 'No asymmetry data available', ha='center', fontsize=12, color='gray')
        add_footer(fig)
        add_page_number(fig, page_num)
        pdf.savefig(fig)
        plt.close(fig)
        return page_num
    
    # Prepare table data (4 columns: Test, Metric, Asymmetry, Status)
    table_data = [[row[0], row[1], row[3], row[4]] for row in asym_data]
    total_rows = len(asym_data)
    
    fig = plt.figure(figsize=(8.5, 11))
    draw_header(fig, 0.92, 'ASYMMETRY SUMMARY')
    
    # Subtitle
    fig.text(0.08, 0.87, 'Left-Right imbalances across all assessments (target: <10%)',
             fontsize=9, fontstyle='italic', color=COLOR_TEXT_MAIN)
    
    # Calculate row height based on number of rows - aim to fill the page
    available_height = 0.70  # From 0.12 to 0.82
    row_height_scale = min(1.8, max(1.0, available_height * 40 / (total_rows + 1)))
    font_size = 8 if total_rows <= 20 else (7 if total_rows <= 30 else 6)
    
    # === TABLE ===
    ax_table = fig.add_axes([0.06, 0.12, 0.88, 0.73])
    ax_table.axis('off')
    
    table = ax_table.table(
        cellText=table_data,
        colLabels=['Test', 'Metric', 'Asymmetry', 'Status'],
        cellLoc='center',
        loc='upper center',
        colWidths=[0.15, 0.45, 0.20, 0.20]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    table.scale(1, row_height_scale)
    
    for (i, j), cell in table.get_celld().items():
        if i == 0:  # Header row
            cell.set_facecolor(COLOR_BLACK)
            cell.set_text_props(weight='bold', color=COLOR_GOLD, ha='center')
            cell.set_height(cell.get_height() * 1.2)
        else:
            if j == 0:  # Test column
                cell.set_text_props(ha='center', weight='bold')
            elif j == 1:  # Metric column
                cell.set_text_props(ha='left')
            else:
                cell.set_text_props(ha='center')
            
            if j == 3:  # Status column
                status = table_data[i-1][3]
                if status == 'OK':
                    cell.set_facecolor(COLOR_GREEN)
                    cell.set_text_props(color=COLOR_WHITE, weight='bold', ha='center')
                elif status == 'MONITOR':
                    cell.set_facecolor('#FFD700')
                    cell.set_text_props(color=COLOR_BLACK, weight='bold', ha='center')
                elif status == 'ADDRESS':
                    cell.set_facecolor(COLOR_RED)
                    cell.set_text_props(color=COLOR_WHITE, weight='bold', ha='center')
            else:
                cell.set_facecolor(COLOR_LIGHT_GRAY if i % 2 == 0 else COLOR_WHITE)
        
        cell.set_edgecolor('#cccccc')
    
    # === LEGEND ===
    fig.text(0.08, 0.08, 'OK = <10%  |  MONITOR = 10-15%  |  ADDRESS = >15%  |  (-) = Right dominant  |  (+) = Left dominant',
             fontsize=7, color=COLOR_TEXT_MAIN, ha='left')
    
    add_footer(fig)
    add_page_number(fig, page_num)
    pdf.savefig(fig)
    plt.close(fig)
    
    return page_num


# ==============================================================================
# MAIN REPORT GENERATOR
# ==============================================================================

def get_teams():
    """Get all teams from Hawkin."""
    try:
        from hdforce import GetTeams
        return GetTeams()
    except Exception as e:
        print(f"Error fetching teams: {e}")
        return None


def generate_single_report(athlete_id, athlete_name, selected_test_types, auto_select_best=True, days_back=60, to_date=None, include_cop=False, add_editable_fields=False, enable_ai=False):
    """
    Generate a single report for one athlete.

    Args:
        athlete_id: Hawkin athlete ID
        athlete_name: Athlete's name
        selected_test_types: List of (display_name, pattern) tuples for tests to include
        auto_select_best: If True, auto-select best trial for each test
        days_back: How many days back to look for tests
        to_date: Optional end date (YYYY-MM-DD) to cap the date range
        include_cop: If True, include COP/stabilogram pages (requires manual CSV download)
        add_editable_fields: If True, add editable form fields for AI summaries (requires PyMuPDF)
        enable_ai: If True, generate AI interpretations using RAG system

    Returns:
        True if report generated successfully, False otherwise
    """
    print(f"\n{'='*60}")
    print(f"Generating report for: {athlete_name}")
    print('='*60)
    
    # Initialize RAG system if AI is enabled (optional - will fall back to direct Claude)
    rag = None
    if enable_ai:
        print("  Initializing AI system...")
        rag = get_rag_system()
        if rag:
            print("  ✓ RAG system ready (will use research documents)")
        else:
            print("  → Using direct Claude API (no RAG)")
    
    # Get tests
    tests = get_tests(athlete_id, days_back=days_back, to_date=to_date)
    
    if tests is None or len(tests) == 0:
        print(f"  No tests found for {athlete_name}")
        return False
    
    print(f"  Found {len(tests)} tests")
    
    # Get system weight
    system_weight = 800
    if 'system_weight_n' in tests.columns:
        weights = tests['system_weight_n'].dropna()
        if len(weights) > 0:
            system_weight = weights.iloc[-1]
    
    body_weight_lb = (system_weight / 9.81) * 2.20462
    
    # Get test date
    test_date = datetime.now().strftime('%m/%d/%Y')
    date_cols = ['timestamp', 'test_time', 'created_at', 'date', 'testDate']
    
    for col in date_cols:
        if col in tests.columns:
            ts = tests[col].iloc[-1]
            try:
                if isinstance(ts, (int, float, np.integer, np.floating)):
                    ts_val = int(ts)
                    if ts_val > 10000000000:
                        ts_val = ts_val // 1000
                    test_date = datetime.fromtimestamp(ts_val).strftime('%m/%d/%Y')
                break
            except:
                pass
    
    # Detect tag column
    tag_col = detect_tag_column(tests)
    if tag_col:
        print(f"  Tag column detected: '{tag_col}'")
        try:
            unique_tags = tests[tag_col].dropna().unique()
            if len(unique_tags) > 0:
                print(f"  Tags found: {', '.join(str(t) for t in unique_tags)}")
        except TypeError:
            # Handles unhashable types (e.g., list columns)
            unique_tags = set(str(v) for v in tests[tag_col].dropna() if v)
            if unique_tags:
                print(f"  Tags found: {', '.join(unique_tags)}")
    else:
        print("  No tag column found (all tests treated as untagged)")
    
    # Find available tests — variant-aware and multi-session
    # Uses parse_test_variant() for testType_name variants, falls back to tag column
    found_tests = {}
    all_asymmetries = []

    for display_name, pattern in selected_test_types:
        all_trials = get_all_tests_by_type(tests, pattern)
        if all_trials is None or len(all_trials) == 0:
            continue

        # Group trials by variant: prefer parsed testType_name, fall back to tag column
        variant_groups = {}
        for idx, row in all_trials.iterrows():
            _, parsed_variant = parse_test_variant(row['testType_name'])
            # Use parsed variant if available, otherwise check tag column
            variant = parsed_variant
            if variant is None and tag_col:
                variant = get_tag_value(row, tag_col)
            variant_groups.setdefault(variant, []).append(idx)

        for variant, idxs in variant_groups.items():
            variant_trials_df = all_trials.loc[idxs]
            test_key = make_found_test_key(display_name, variant)
            label = make_display_label(display_name, variant)

            # Group by session (date)
            sessions = group_trials_by_session(variant_trials_df)

            # Compute session averages
            for s in sessions:
                s['avg_metrics'] = compute_session_averages(s['trials'], display_name)

            # Select best trial from latest session for force-time
            latest_session = sessions[-1]
            test_row = select_trial(latest_session['trials'], label, pattern,
                                    auto_select_best=auto_select_best, tag_col=tag_col)

            if test_row is not None:
                test_id = test_row.get('id')
                force_data = get_force_time(test_id) if test_id else None

                found_tests[test_key] = {
                    'row': test_row,
                    'force_data': force_data,
                    'tag': variant,
                    'display_label': label,
                    'base_name': display_name,
                    'sessions': sessions
                }

                # Collect asymmetries (variant-aware abbreviation)
                # Skip single-leg variants — their ~100% asymmetry is meaningless
                if not _is_single_leg_tag(variant):
                    abbrev = display_name.replace('Countermovement Jump', 'CMJ')\
                                         .replace('CMJ Rebound', 'CMJ-R')\
                                         .replace('Squat Jump', 'SJ')\
                                         .replace('Multi Rebound', 'MR')\
                                         .replace('Bodyweight Squats', 'BW Squats')\
                                         .replace('Static Balance', 'Balance')\
                                         .replace('Drop Landing', 'DL')
                    if variant:
                        abbrev = f"{abbrev} [{variant}]"
                    asyms = extract_asymmetries(test_row, abbrev)
                    all_asymmetries.extend(asyms)
            
            # Print session summary
            n_sess = len(sessions)
            total_trials = sum(len(s['trials']) for s in sessions)
            variant_note = f" [{variant}]" if variant else ""
            print(f"  {display_name}{variant_note}: {total_trials} trials across {n_sess} session(s)")
    
    if not found_tests:
        print(f"  No valid tests found for {athlete_name}")
        return False
    
    # Generate AI summaries if enabled
    dashboard_ai_summary = None
    test_ai_interpretations = {}
    
    if enable_ai:
        print("  Generating AI interpretations...")
        
        # Try RAG first, fall back to direct Claude
        rag_worked = False
        if rag:
            dashboard_ai_summary = generate_dashboard_ai_summary(found_tests, athlete_name, rag)
            if dashboard_ai_summary:
                rag_worked = True
                print("    ✓ Dashboard summary (RAG)")
        
        # If RAG didn't work, use direct Claude
        if not dashboard_ai_summary:
            dashboard_ai_summary = generate_dashboard_summary_direct(found_tests, athlete_name)
            if dashboard_ai_summary:
                print("    ✓ Dashboard summary (Claude)")
        
        # Test-specific interpretations
        for test_key, test_data in found_tests.items():
            interp = None
            # Use base_name for type matching, display_label for personalization
            base_name = test_data.get('base_name', test_key)
            label = test_data.get('display_label', test_key)
            
            # Try RAG first
            if rag and rag_worked:
                interp = generate_test_ai_interpretation(
                    base_name, 
                    test_data.get('row'), 
                    test_data.get('force_data'),
                    rag
                )
            
            # Fall back to direct Claude
            if not interp:
                interp = generate_test_interpretation_direct(
                    base_name,
                    test_data.get('row'),
                    system_weight
                )
            
            if interp:
                test_ai_interpretations[test_key] = interp
                print(f"    ✓ {label}")
    
    # Compute session date range across all found tests
    all_session_dates = []
    for test_data in found_tests.values():
        for s in test_data.get('sessions', []):
            if s.get('date'):
                all_session_dates.append(s['date'])
    date_range = None
    if all_session_dates:
        first_date = min(all_session_dates).strftime('%m/%d/%Y')
        last_date = max(all_session_dates).strftime('%m/%d/%Y')
        date_range = (first_date, last_date)

    # Generate PDF
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    safe_name = athlete_name.replace(' ', '_').replace('/', '-')
    output_file = os.path.join(OUTPUT_FOLDER,
                               f"Force_Plate_{safe_name}_{datetime.now().strftime('%Y%m%d')}.pdf")

    print(f"  Creating: {os.path.basename(output_file)}")

    page_num = 0

    with PdfPages(output_file) as pdf:
        # Cover page
        page_num += 1
        create_cover_page(pdf, athlete_name, body_weight_lb, test_date, page_num, date_range=date_range)
        
        # Performance metrics summary page (with AI summary if available)
        # Dashboard uses base test names (prefers untagged variants for EUR calc etc.)
        page_num += 1
        create_metrics_summary_page(pdf, found_tests, page_num, ai_summary=dashboard_ai_summary)
        
        # Test pages + trend pages — ordered by base test type, then tag variants
        base_test_order = ['Bodyweight Squats', 'Drop Landing', 'Static Balance',
                      'Countermovement Jump', 'CMJ Rebound', 'Squat Jump', 'Multi Rebound']
        
        tests_with_stabilograms = ['Drop Landing', 'Bodyweight Squats', 'Static Balance']
        stabilogram_only_tests = ['Static Balance']
        
        # Build ordered list of found_test keys matching the base order
        # Track which pages are test pages (for editable field placement)
        test_page_indices = set()
        
        ordered_keys = []
        for base_name in base_test_order:
            # First add untagged version, then tagged variants
            for key, data in found_tests.items():
                if data['base_name'] == base_name and data['tag'] is None:
                    ordered_keys.append(key)
            for key, data in found_tests.items():
                if data['base_name'] == base_name and data['tag'] is not None:
                    ordered_keys.append(key)
        
        for test_key in ordered_keys:
            test_data = found_tests[test_key]
            base_name = test_data['base_name']
            tag = test_data.get('tag')
            sessions = test_data.get('sessions', [])
            
            if base_name not in stabilogram_only_tests:
                page_num += 1
                test_page_indices.add(page_num - 1)  # 0-based index for PyMuPDF
                custom_interp = test_ai_interpretations.get(test_key)
                create_test_page(pdf, base_name, test_data.get('force_data'), 
                               system_weight, page_num, 
                               test_row=test_data.get('row'),
                               custom_interpretation=custom_interp,
                               sessions=sessions,
                               tag=tag)
                
                # Add dedicated trend page if 2+ sessions with data
                sessions_with_data = [s for s in sessions
                                      if s.get('avg_metrics')]
                if len(sessions_with_data) >= 2:
                    page_num += 1
                    create_trend_page(pdf, base_name, sessions_with_data, page_num, tag=tag)
            
            # COP pages
            if include_cop and base_name in tests_with_stabilograms:
                page_num += 1
                cop_file = find_cop_file(athlete_name, base_name)
                cop_data = load_cop_data(cop_file) if cop_file else None
                create_stabilogram_page(pdf, base_name, cop_data, page_num)
        
        # Asymmetry summary
        page_num += 1
        create_asymmetry_page(pdf, all_asymmetries, page_num)
    
    # Add editable form fields if requested
    if add_editable_fields and PYMUPDF_AVAILABLE:
        add_editable_text_fields(output_file, test_page_indices=test_page_indices)
        print(f"  ✓ Added editable form fields")
    elif add_editable_fields and not PYMUPDF_AVAILABLE:
        print(f"  ⚠ Editable fields not added (PyMuPDF not installed)")
    
    print(f"  ✓ Saved ({page_num} pages)")
    return True


def generate_report():
    """Main function to generate the report."""
    
    print("\n" + "="*60)
    print("FORCE PLATE REPORT GENERATOR")
    print("Move, Measure, Analyze LLC")
    print("="*60)
    
    # Check token
    if not API_TOKEN or API_TOKEN == "YOUR_TOKEN_HERE":
        print("\n⚠ Set your API token in the script first!")
        input("\nPress Enter to exit...")
        return
    
    # Connect
    print("\nConnecting to API...")
    if not setup_api():
        input("\nPress Enter to exit...")
        return
    
    # Get teams
    print("\nFetching teams...")
    teams = get_teams()
    
    # Team selection
    print("\n" + "="*60)
    print("TEAM SELECTION")
    print("="*60)
    
    if teams is not None and len(teams) > 0:
        print("\nAvailable Teams:")
        for i, team in enumerate(teams.itertuples()):
            print(f"  {i+1}. {team.name}")
        print(f"  {len(teams)+1}. [ALL ATHLETES]")
        
        while True:
            selection = input(f"\nSelect team (1-{len(teams)+1}): ").strip()
            if selection.isdigit():
                idx = int(selection) - 1
                if idx == len(teams):
                    selected_team_id = None
                    selected_team_name = "All Athletes"
                    break
                elif 0 <= idx < len(teams):
                    selected_team_id = teams.iloc[idx]['id']
                    selected_team_name = teams.iloc[idx]['name']
                    break
            print("Invalid selection.")
    else:
        print("No teams found. Showing all athletes.")
        selected_team_id = None
        selected_team_name = "All Athletes"
    
    print(f"\n✓ Selected: {selected_team_name}")
    
    # Get athletes
    print("\nFetching athletes...")
    all_athletes = get_athletes()
    
    if all_athletes is None or len(all_athletes) == 0:
        print("No athletes found")
        input("\nPress Enter to exit...")
        return
    
    # Filter by team if selected
    if selected_team_id is not None:
        if 'teams' in all_athletes.columns:
            def check_team(teams_val):
                if isinstance(teams_val, list):
                    for t in teams_val:
                        if isinstance(t, dict) and t.get('id') == selected_team_id:
                            return True
                        elif t == selected_team_id:
                            return True
                return False
            athletes = all_athletes[all_athletes['teams'].apply(check_team)]
        else:
            athletes = all_athletes
        
        if len(athletes) == 0:
            print(f"  No athletes found in {selected_team_name}. Using all athletes.")
            athletes = all_athletes
    else:
        athletes = all_athletes
    
    name_col = 'name' if 'name' in athletes.columns else 'Name'
    id_col = 'id' if 'id' in athletes.columns else 'Id'
    
    print(f"\nAthletes in {selected_team_name} ({len(athletes)}):")
    for i, (_, row) in enumerate(athletes.iterrows()):
        print(f"  {i+1}. {row[name_col]}")
    
    # Generation mode selection
    print("\n" + "="*60)
    print("GENERATION MODE")
    print("="*60)
    print("\nOptions:")
    print("  [Enter] = Generate report for ONE athlete")
    print(f"  [a] = Generate reports for ALL {len(athletes)} athletes (batch mode)")
    
    mode = input("\nSelect mode (Enter or a): ").strip().lower()
    batch_mode = (mode == 'a')
    
    # Test type selection (applies to all reports)
    print("\n" + "="*60)
    print("TEST SELECTION")
    print("="*60)
    print("Select which tests to include in report(s):")
    
    all_test_types = [
        ('Drop Landing', 'Drop Landing'),
        ('Drop Jump', 'Drop Jump'),
        ('Bodyweight Squats', 'Bodyweight Squats|Free Run-Bodyweight|Free Run.*Bodyweight'),
        ('Static Balance', 'Balance.*Firm|Balance.*Eyes|Free Run-Balance|Free Run.*Balance'),
        ('CMJ Rebound', 'CMJ Rebound|Countermovement.*Rebound'),
        ('Countermovement Jump', 'Countermovement Jump'),
        ('Squat Jump', 'Squat Jump'),
        ('Multi Rebound', 'Multi Rebound'),
        ('Isometric Test', 'Isometric Test'),
    ]
    
    selected_test_types = []
    for display_name, pattern in all_test_types:
        include = input(f"  Include {display_name}? (Y/n): ").strip().lower()
        if include != 'n':
            selected_test_types.append((display_name, pattern))
    
    if not selected_test_types:
        print("\nNo tests selected.")
        input("\nPress Enter to exit...")
        return
    
    print(f"\n✓ Selected {len(selected_test_types)} test type(s)")
    
    # Days back selection
    days_input = input("\nHow many days back to search for tests? (default: 60): ").strip()
    days_back = int(days_input) if days_input.isdigit() else 60
    
    if batch_mode:
        # Batch mode - generate for all athletes
        print("\n" + "="*60)
        print(f"BATCH GENERATION: {len(athletes)} athletes")
        print("="*60)
        print("\nAuto-selecting best trial for each test...")
        
        # AI option for batch
        ai_mode = input("\nEnable AI interpretations for all reports? (y/N): ").strip().lower()
        enable_ai = (ai_mode == 'y')
        
        editable_mode = input("Add editable text fields to all reports? (y/N): ").strip().lower()
        add_editable = (editable_mode == 'y')
        
        success_count = 0
        fail_count = 0
        
        for i, (_, row) in enumerate(athletes.iterrows()):
            athlete_id = row[id_col]
            athlete_name = row[name_col]
            
            print(f"\n[{i+1}/{len(athletes)}] Processing {athlete_name}...")
            
            if generate_single_report(athlete_id, athlete_name, selected_test_types, 
                                      auto_select_best=True, days_back=days_back,
                                      enable_ai=enable_ai, add_editable_fields=add_editable):
                success_count += 1
            else:
                fail_count += 1
        
        print("\n" + "="*60)
        print("BATCH GENERATION COMPLETE")
        print("="*60)
        print(f"  ✓ Successful: {success_count}")
        print(f"  ✗ Failed/Skipped: {fail_count}")
        print(f"  Reports saved to: {OUTPUT_FOLDER}")
        
    else:
        # Single athlete mode
        selection = input("\nEnter athlete number or name: ").strip()
        
        if selection.isdigit():
            idx = int(selection) - 1
            if 0 <= idx < len(athletes):
                athlete_id = athletes.iloc[idx][id_col]
                athlete_name = athletes.iloc[idx][name_col]
            else:
                print("Invalid selection")
                input("\nPress Enter to exit...")
                return
        else:
            match = athletes[athletes[name_col].str.contains(selection, case=False, na=False)]
            if len(match) > 0:
                athlete_id = match.iloc[0][id_col]
                athlete_name = match.iloc[0][name_col]
            else:
                print(f"No athlete found matching '{selection}'")
                input("\nPress Enter to exit...")
                return
        
        print(f"\n✓ Selected: {athlete_name}")
        
        # Trial selection mode
        print("\n" + "="*60)
        print("TRIAL SELECTION")
        print("="*60)
        print("\nOptions:")
        print("  [Enter] = Manually select each trial")
        print("  [b] = Auto-select best trial for each test (by mRSI, RSI, etc.)")
        
        auto_mode = input("\nSelection mode (Enter or b): ").strip().lower()
        auto_select_best = (auto_mode == 'b')
        
        # AI interpretation option
        print("\n" + "="*60)
        print("AI INTERPRETATIONS")
        print("="*60)
        print("\nGenerate AI-powered interpretations for the report?")
        print("(Requires research documents in the RAG system)")
        ai_mode = input("Enable AI interpretations? (y/N): ").strip().lower()
        enable_ai = (ai_mode == 'y')
        
        # Editable fields option
        editable_mode = input("Add editable text fields? (y/N): ").strip().lower()
        add_editable = (editable_mode == 'y')
        
        generate_single_report(athlete_id, athlete_name, selected_test_types,
                              auto_select_best=auto_select_best, days_back=days_back,
                              enable_ai=enable_ai, add_editable_fields=add_editable)
    
    input("\nPress Enter to exit...")


# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == '__main__':
    try:
        generate_report()
    except KeyboardInterrupt:
        print("\n\nCancelled.")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")