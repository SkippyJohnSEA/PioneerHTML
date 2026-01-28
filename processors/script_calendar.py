import os
from utils.file_io import build_output_filename
import pandas as pd
from collections import defaultdict
from pathlib import Path

# ---------------------------------------------------------
# Detect if the date field is a date, a range of dates or a string
# ---------------------------------------------------------

def format_date_or_range(start, end):
    if pd.isnull(start):
        return ""

    # Try normal Excel date first
    try:
        start_dt = pd.to_datetime(start,errors="coerce").normalize()
        #if no end date, just a single day event
        if pd.isnull(end):
            # Format: "January 1" (no leading zero)
            return f"{start_dt.strftime('%B')} {start_dt.day}"
        
        end_dt = pd.to_datetime(end,errors="coerce").normalize()
        num_days = (end_dt - start_dt).days 
        
        #if delta is 0 or negative, single day event
        if num_days<1:
            # Format: "January 1" (no leading zero)
            return f"{start_dt.strftime('%B')} {start_dt.day}"            
            
        #if dates in same month
        if start_dt.month == end_dt.month and start_dt.year == end_dt.year:
            return f"{start_dt.strftime('%b')} {start_dt.day}-{end_dt.day}"

        # Cross-month → "Sep 29-Oct 2"
        return f"{start_dt.strftime('%b')} {start_dt.day}-{end_dt.strftime('%b')} {end_dt.day}"
            
    except Exception as e:
        print("EXCEPTION:", type(e), e)
        raise
        #pass

    # Fallback: return raw text
    return str(start)

# ---------------------------------------------------------
# Detect if the day should be a single day or a range
# ---------------------------------------------------------

def format_day_or_range(start, end):
    if pd.isnull(start):
        return ""

    # Try normal Excel date first
    try:
        start_dt = pd.to_datetime(start).normalize()
        #if no end date, just a single day event
        if pd.isnull(end):
            # Format: "January 1" (no leading zero)
            return start_dt.strftime("%A")
        
        end_dt = pd.to_datetime(end).normalize()
        num_days = (end_dt - start_dt).days
        
        if num_days > 7 or num_days < 0:
            return ""
            
        return f"{start_dt.strftime('%a')}-{end_dt.strftime('%a')}"

    except:
        pass

    # Fallback: return raw text
    return str(start)

# ---------------------------------------------------------
# Safe sorting for dates within a month
# ---------------------------------------------------------
def sort_key(row):
    raw_start, raw_end, _, _, _ = row

    start_dt = pd.to_datetime(raw_start, errors="coerce")
    end_dt   = pd.to_datetime(raw_end, errors="coerce")

    # Duration: if no end date, treat as 0-day event
    if pd.isna(start_dt):
        return (pd.Timestamp.max, 999999)  # push invalid dates to bottom

    if pd.isna(end_dt):
        duration = 0
    else:
        duration = (end_dt - start_dt).days

    return (start_dt, duration)

# ---------------------------------------------------------
# Read Excel and group rows by accordion title
# ---------------------------------------------------------
def read_excel_grouped(path):
    df = pd.read_excel(path, sheet_name="Events")

    groups = defaultdict(list)

    for _, row in df.iterrows():
        
        raw_date = row["StartDate"]
        raw_end_date = row["EndDate"]
        
        # Accordion title: Month + Year
        #raw_title = row["AccordionTitle"]
        if pd.notnull(raw_date):
            #Establish Accordian Group
            try:
                title = pd.to_datetime(raw_date).strftime("%B %Y")
            except:
                title = str(raw_date)
                
            #Determine Dates of event
            try:
                date = format_date_or_range(raw_date, raw_end_date)
            except:
                date = str(raw_date)
            try:
                day = format_day_or_range(raw_date, raw_end_date)
            except:
                day = ""
        else:
            title = "Untitled"
            date = ""
            day = ""
        
        desc = str(row["Description"])

        # Store the raw start date so we can sort later
        groups[title].append((raw_date, raw_end_date, date, day, desc))

    # 🔥 Sort each month by actual date and duration
    for title in groups:
        #groups[title].sort(key=lambda x: pd.to_datetime(x[0]))
        groups[title].sort(key=sort_key)

    return groups


# ---------------------------------------------------------
# Build a 3-column table for each accordion body
# ---------------------------------------------------------
def build_table(rows):
    html = []
    html.append("""
        <table class="event-table">
            <thead>
                <tr>
                    <th class="col-date">Date</th>
                    <th class="col-day">Day</th>
                    <th class="col-desc">Description</th>
                </tr>
            </thead>
            <tbody>
    """)

    for i, (_, _, date, day, desc) in enumerate(rows):
        row_class = "even-row" if i % 2 == 0 else "odd-row"
        html.append(f"""
                <tr class="{row_class}">
                    <td class="col-date">{date}</td>
                    <td class="col-day">{day}</td>
                    <td class="col-desc">{desc}</td>
                </tr>
        """)

    html.append("""
            </tbody>
        </table>
    """)
    return "\n".join(html)



# ---------------------------------------------------------
# Build one accordion item
# ---------------------------------------------------------
def build_accordion_item(title, table_html, index, accordion_id):
    collapse_id = f"collapse{index}"
    heading_id = f"heading{index}"

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
                    {table_html}
                </div>
            </div>
        </div>
    """


# ---------------------------------------------------------
# Build the full HTML document
# ---------------------------------------------------------

def safe_title_to_dt(title):
    try:
        return pd.to_datetime(title, format="%B %Y")
    except:
        return pd.Timestamp.max   # push non-dates to the end

def generate_full_html(excel_path):
    groups = read_excel_grouped(excel_path)

    accordion_id = "accordionMaster"

    accordion_items = []

    sorted_groups = sorted(groups.items(), key=lambda item: safe_title_to_dt(item[0]))

    for idx, (title, rows) in enumerate(sorted_groups, start=1):
        table_html = build_table(rows)
        accordion_items.append(build_accordion_item(title, table_html, idx, accordion_id))

    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Accordion Output</title>

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">

    <style>
        .custom-accordion-header {{
            background-color: #DDDDDD;
            font-weight: 700;
            font-family: Arial, sans-serif;
            font-size: 18px;
            padding: 2px 2px;
            color: #000;
        }}

        .custom-accordion-item {{
            border: 1px solid #C9C3AD;
            margin-bottom: 3px;
            border-radius: 3px;
            overflow: hidden;
        }}

        .custom-accordion-body {{
            background-color: #e2e6f1;
            padding: 2px 2px;
            font-family: Arial, sans-serif;
            font-size: 14px;
            color: #000;
        }}

        .event-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 14px;
            table-layout: fixed;
        }}

        .event-table th {{
            background-color: #C9C3AD;
            padding: 2px;
            text-align: left;
            font-weight: 700;
            border-bottom: 2px solid #999;
        }}

        .event-table td {{
            padding: 2px 2px;
            vertical-align: top;
        }}

        .col-date {{
            width: 12%;
        }}

        .col-day {{
            width: 8%;
        }}

        .col-desc {{
            width: 80%;
        }}

        .even-row {{
            background-color: #FFFFFF;
        }}

        .odd-row {{
            background-color: #F2F2F2;
        }}
    </style>

</head>

<body style="padding: 2px;">

    <div class="accordion" id="{accordion_id}">
        {''.join(accordion_items)}
    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>
"""

# ---------------------------------------------------------
# Write output to accordian_out.html
# ---------------------------------------------------------
def run(input_path, original_name):
    output_path = build_output_filename(original_name, ".html")

    html_output = generate_full_html(input_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_output)

    return output_path