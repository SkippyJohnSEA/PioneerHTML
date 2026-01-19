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

def build_output_filename(original_name, extension=".html"):
    """
    Creates a collision-safe output filename inside temp/.
    Example: SocialList.html, SocialList_1.html, SocialList_2.html
    """
    ensure_temp_dir()

    base, _ = os.path.splitext(original_name)
    candidate = os.path.join(TEMP_DIR, base + extension)

    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(TEMP_DIR, f"{base}_{counter}{extension}")
        counter += 1

    return candidate