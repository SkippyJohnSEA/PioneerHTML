import os
import math
import pandas as pd
from utils.file_io import make_output_html_filename
from pathlib import Path

def read_pairs_from_excel(xlsx_path):
    """Read (year, name) pairs from an Excel file with columns 'Year' and 'Name'."""
    df = pd.read_excel(xlsx_path, engine="openpyxl")

    # Ensure required columns exist
    df = df[["Year", "Name"]]

    # Clean values
    df["Year"] = df["Year"].astype(str).str.strip()
    df["Name"] = df["Name"].fillna("").astype(str).str.strip()

    # Convert to list of tuples
    return list(df.itertuples(index=False, name=None))

def split_into_columns(pairs, num_columns=4):
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
    width: 100%;
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
    font-size: 16px;
  }

  table.eight-col td {
    padding: 4px 4px;
    border: 2px solid #fff;
    font-size: 14px;
  }

  table.eight-col tr:nth-child(odd) td {
    background-color: #fff;
    color: #000;
  }

  table.eight-col tr:nth-child(even) td {
    background-color: #e9e4d0;
    color: #000;
  }

  table.eight-col td:nth-child(even),
  table.eight-col th:nth-child(even) {
    width: 18%;
  }

  table.eight-col td:nth-child(odd),
  table.eight-col th:nth-child(odd) {
    width: 5%;
  }
</style>

<table class="eight-col">
  <tr>
""")

    # Header row
    for _ in range(4):
        html.append("    <th>Year</th><th>Name</th>")
    html.append("  </tr>\n")

    # Data rows
    for r in range(rows):
        html.append("  <tr>")
        for col in columns:
            if r < len(col):
                year, name = col[r]
                html.append(f"<td>{year}</td><td>{name}</td>")
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
def run(input_path):
    output_path = make_output_html_filename(input_path)

    pairs = read_pairs_from_excel(input_path)
    columns, rows = split_into_columns(pairs, num_columns=4)
    html = generate_html(columns, rows)

    html_output = generate_html(input_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_output)

    return output_path
