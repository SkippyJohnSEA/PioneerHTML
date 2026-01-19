import os
import shutil
import streamlit as st
from processors.registry import SCRIPTS
from utils.file_io import save_uploaded_file
import streamlit.components.v1 as components

TEMP_DIR = "temp"

# Clean up previous run’s temp files
if os.path.exists(TEMP_DIR):
    shutil.rmtree(TEMP_DIR)

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

    # Read generated HTML
    with open(output_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # -------------------------
    # TABBED UI
    # -------------------------
    tab_preview, tab_html, tab_download = st.tabs(["Preview", "HTML", "Download"])

    # -------------------------
    # 1. PREVIEW TAB
    # -------------------------
    with tab_preview:
        st.subheader("Live Preview")
        components.html(html_content, height=600, scrolling=True)

    # -------------------------
    # 2. HTML TAB (copy button + code block)
    # -------------------------
    with tab_html:
        st.subheader("HTML Source")

        copy_button = f"""
            <button onclick="navigator.clipboard.writeText(`{html_content}`)"
                    style="
                        background-color:#4CAF50;
                        color:white;
                        padding:8px 14px;
                        border:none;
                        border-radius:4px;
                        cursor:pointer;
                        font-size:14px;
                        margin-bottom:10px;
                    ">
                Copy HTML to Clipboard
            </button>
        """
        components.html(copy_button, height=60)

        st.code(html_content, language="html")

    # -------------------------
    # 3. DOWNLOAD TAB
    # -------------------------
    with tab_download:
        st.subheader("Download HTML File")
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download HTML Output",
                data=f,
                file_name=os.path.basename(output_path),
                mime="text/html"
            )