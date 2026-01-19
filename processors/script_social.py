import pandas as pd
from pathlib import Path
import html
import os
from utils.file_io import make_output_html_filename

# ---------------------------------------------------------
# Read Excel and return list of (title, description)
# ---------------------------------------------------------
def read_excel_rows(path):
    df = pd.read_excel(path)

    rows = []
    for _, row in df.iterrows():
        title = str(row.get("Title", "Untitled"))
        desc_raw = row.get("Description", "")
        desc = "" if pd.isnull(desc_raw) else str(desc_raw)
        rows.append((title, desc))

    return rows


# ---------------------------------------------------------
# Build one accordion item
# ---------------------------------------------------------
def build_accordion_item(title, body_text, index, accordion_id):
    collapse_id = f"collapse{index}"
    heading_id = f"heading{index}"

    clean_text = body_text.strip()
    safe_text = html.escape(clean_text)

    return f"""
        <div class="accordion-item custom-accordion-item">
            <h2 class="accordion-header" id="{heading_id}">
                <button class="accordion-button custom-accordion-header collapsed"
                        type="button"
                        data-bs-toggle="collapse"
                        data-bs-target="#{collapse_id}"
                        aria-expanded="false"
                        aria-controls="{collapse_id}">
                    {title}
                </button>
            </h2>

            <div id="{collapse_id}"
                 class="accordion-collapse collapse"
                 aria-labelledby="{heading_id}"
                 data-bs-parent="#{accordion_id}">
                <div class="accordion-body custom-accordion-body">
                    {safe_text}
                </div>
            </div>
        </div>
    """


# ---------------------------------------------------------
# Build the full HTML document
# ---------------------------------------------------------
def generate_full_html(excel_path):
    rows = read_excel_rows(excel_path)
    accordion_id = "accordionMaster"

    accordion_items = []
    for idx, (title, text) in enumerate(rows, start=1):
        accordion_items.append(
            build_accordion_item(title, text, idx, accordion_id)
        )

    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Accordion Output</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">

    <style>
        .custom-accordion-header {{
            background-color: #dcdcdc;
            font-weight: 700;
            font-family: Arial, sans-serif;
            font-size: 18px;
            padding: 6px 10px;
            color: #000;
        }}

        .custom-accordion-item {{
            border: 1px solid #bfbfbf;
            margin-bottom: 4px;
            border-radius: 4px;
            overflow: hidden;
        }}

        .custom-accordion-body {{
            background-color: #eef1f7;
            padding: 10px 10px;
            line-height: 1.1;
            font-family: Arial, sans-serif;
            font-size: 15px;
            color: #000;
        }}
    </style>

</head>

<body style="padding: 1px;">

    <div class="accordion" id="{accordion_id}">
        {''.join(accordion_items)}
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>
"""


# ---------------------------------------------------------
# Public run() function for Streamlit integration
# ---------------------------------------------------------
def run(input_path):
    output_path = make_output_html_filename(input_path)
    html_output = generate_full_html(input_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_output)

    return output_path
