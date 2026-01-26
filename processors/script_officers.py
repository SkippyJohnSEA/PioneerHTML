import os
import math
import pandas as pd
from utils.file_io import build_output_filename
from pathlib import Path

num_table_columns = 1

def read_pairs_from_excel(xlsx_path):
    """Read (year, name) pairs from an Excel file with columns 'Name' and 'Office'."""
    df = pd.read_excel(xlsx_path, sheet_name="Officers", engine="openpyxl")

    # Ensure required columns exist
    df = df[["Name", "Office"]]

    # Clean values
    df["Name"] = df["Name"].fillna("N/A").astype(str).str.strip()
    df["Office"] = df["Office"].fillna("").astype(str).str.strip()

    # Convert to list of tuples
    return list(df.itertuples(index=False, name=None))

def split_into_columns(pairs, num_columns):
    """Split the list of pairs evenly across the specified number of columns."""
    total = len(pairs)
    rows = math.ceil(total / num_columns)
    return [pairs[i*rows:(i+1)*rows] for i in range(num_columns)], rows


def generate_html(columns, rows):
    """Generate the full HTML table string."""
    html = []

    html.append("""
<style>
  table.eight-col {
    width: 35%;
    border-collapse: collapse;
    font-family: Arial, sans-serif;
  }

  table.eight-col th {
    padding: 2px 2px;
    border: 2px solid #fff;
    background-color: #c8c09e;
    color: #000;
    font-weight: bold;
    text-align: left;
    font-size: 19px;
  }

  table.eight-col td {
    padding: 4px 4px;
    border: 2px solid #fff;
    font-size: 19px;
  }

  table.eight-col tr:nth-child(odd) td {
    background-color: #c8c09e;
    color: #000;
  }

  table.eight-col tr:nth-child(even) td {
    background-color: #fff;
    color: #000;
  }

  table.eight-col td:nth-child(even),
  table.eight-col th:nth-child(even) {
    width: 55%;
    font-weight: bold;
  }

  table.eight-col td:nth-child(odd),
  table.eight-col th:nth-child(odd) {
    width: 45%;
  }
</style>

<table class="eight-col">
  <tr>
""")

    # Header row
    for _ in range(num_table_columns):
        html.append("    <th>Name</th><th>Office</th>")
    html.append("  </tr>\n")

    # Data rows
    for r in range(rows):
        html.append("  <tr>")
        for col in columns:
            if r < len(col):
                name, office = col[r]
                html.append(f"<td>{name}</td><td>{office}</td>")
            else:
                html.append("<td></td><td></td>")
        html.append("</tr>\n")

    html.append("</table>")
    return "".join(html)


def write_html_to_file(html, output_path):
    """Write the generated HTML to a file."""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------------------------------------------------
# Public run() function for Streamlit integration 
# ---------------------------------------------------------
def run(input_path, original_name):
    output_path = build_output_filename(original_name, ".html")

    pairs = read_pairs_from_excel(input_path)
    columns, rows = split_into_columns(pairs, num_columns=num_table_columns)
    html_output = generate_html(columns, rows)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_output)

    return output_path