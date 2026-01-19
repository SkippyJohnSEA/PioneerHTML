import uuid

def save_uploaded_file(uploaded_file):
    temp_name = f"uploaded_{uuid.uuid4().hex}.xlsx"
    with open(temp_name, "wb") as f:
        f.write(uploaded_file.read())
    return temp_name
