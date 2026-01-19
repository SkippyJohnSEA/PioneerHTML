import streamlit as st
from processors.registry import SCRIPTS
from utils.file_io import save_uploaded_file

st.title("Excel → HTML Processing Tool")

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

script_choice = st.selectbox("Choose script", list(SCRIPTS.keys()))
script_info = SCRIPTS[script_choice]

target_filename = None
if script_info["needs_target"]:
    target_filename = st.text_input("Enter target filename")

if uploaded and st.button("Run"):
    input_path = save_uploaded_file(uploaded)

    func = script_info["func"]

    if script_info["needs_target"]:
        output_path = func(input_path, target_filename)
    else:
        output_path = func(input_path)

    with open(output_path, "rb") as f:
        st.download_button(
            label="Download HTML Output",
            data=f,
            file_name=output_path,
            mime="text/html"
        )
