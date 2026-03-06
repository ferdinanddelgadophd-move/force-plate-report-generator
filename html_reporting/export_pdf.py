"""
export_pdf.py — Export HTML report to PDF using Playwright (Chromium).

STRICT REQUIREMENTS:
- Uses page.pdf() only (no browser print dialog, no wkhtmltopdf).
- Letter format, printBackground=True, explicit margins.
"""

import os


def export_pdf(html_path, pdf_path):
    """Export an HTML report to PDF using Playwright Chromium.

    Uses the synchronous Playwright API to avoid event-loop conflicts
    when called from Flask or other synchronous frameworks.

    Parameters
    ----------
    html_path : str
        Path to the rendered HTML file.
    pdf_path : str
        Destination path for the PDF.

    Returns
    -------
    str
        Absolute path to the generated PDF.
    """
    from playwright.sync_api import sync_playwright

    os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Load local HTML file
        file_url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
        page.goto(file_url, wait_until="networkidle")

        # Wait for fonts to load
        page.wait_for_timeout(1500)

        # Export PDF with controlled settings
        page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True,
            margin={
                "top": "0.6in",
                "right": "0.5in",
                "bottom": "0.8in",
                "left": "0.5in",
            },
            display_header_footer=False,
        )

        browser.close()

    return os.path.abspath(pdf_path)


def export_pdf_from_string(html_string, pdf_path):
    """Export PDF from an HTML string (writes to temp file first).

    Parameters
    ----------
    html_string : str
        Complete HTML document string.
    pdf_path : str
        Destination path for the PDF.

    Returns
    -------
    str
        Absolute path to the generated PDF.
    """
    import tempfile

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".html", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(html_string)
        tmp_path = tmp.name

    try:
        result = export_pdf(tmp_path, pdf_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    return result
