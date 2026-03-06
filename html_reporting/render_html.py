"""
render_html.py — Render payload to HTML using Jinja2 templates.
"""

import os
from jinja2 import Environment, FileSystemLoader, select_autoescape

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
_TEMPLATES_DIR = os.path.join(_MODULE_DIR, "templates")
_STYLES_DIR = os.path.join(_MODULE_DIR, "styles")


def _read_css(filename):
    """Read a CSS file from the styles directory."""
    path = os.path.join(_STYLES_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def render(payload, inline_styles=True):
    """Render report HTML from a payload dict.

    Parameters
    ----------
    payload : dict
        Output of payload.build_payload().
    inline_styles : bool
        If True, embed CSS directly in the HTML (required for PDF export).
        If False, link to external CSS files (for development).

    Returns
    -------
    str
        Complete HTML document.
    """
    env = Environment(
        loader=FileSystemLoader(_TEMPLATES_DIR),
        autoescape=select_autoescape(["html"]),
    )
    template = env.get_template("report.html")

    # Template context
    ctx = {
        **payload,
        "styles_dir": os.path.relpath(_STYLES_DIR, _TEMPLATES_DIR).replace("\\", "/"),
        "inline_styles": inline_styles,
        "asym_ok": payload.get("asym_ok", 10),
        "asym_concern": payload.get("asym_concern", 15),
    }

    if inline_styles:
        ctx["screen_css"] = _read_css("screen.css")
        ctx["print_css"] = _read_css("print.css")

    return template.render(**ctx)


def render_to_file(payload, output_path, inline_styles=True):
    """Render and write HTML to a file.

    Parameters
    ----------
    payload : dict
    output_path : str
        Path to write the .html file.
    inline_styles : bool

    Returns
    -------
    str
        Absolute path to the written file.
    """
    html = render(payload, inline_styles=inline_styles)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return os.path.abspath(output_path)
