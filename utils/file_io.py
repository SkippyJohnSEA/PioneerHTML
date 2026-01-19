import os
import uuid

TEMP_DIR = "temp"

def ensure_temp_dir():
    if not os.path.exists(TEMP_DIR):
        os.makedirs(TEMP_DIR)

def save_uploaded_file(uploaded_file):
    ensure_temp_dir()
    temp_name = os.path.join(TEMP_DIR, f"uploaded_{uuid.uuid4().hex}.xlsx")
    with open(temp_name, "wb") as f:
        f.write(uploaded_file.read())
    return temp_name

def make_output_html_filename(input_path):
    base = os.path.splitext(input_path)[0]
    return base + ".html"