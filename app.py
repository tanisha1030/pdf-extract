import streamlit as st
import pandas as pd
import json
from main_code import extract_from_pdf, extract_from_docx, extract_from_pptx, extract_from_excel
from io import BytesIO
import base64

st.set_page_config(page_title="Universal Info Extractor", layout="wide")
st.title("📄 Universal File Information Extractor")

@st.cache_resource(show_spinner=True)
def process_pdf(file_bytes):
    return extract_from_pdf(BytesIO(file_bytes))

@st.cache_resource(show_spinner=True)
def process_docx(file_bytes):
    return extract_from_docx(BytesIO(file_bytes))

@st.cache_resource(show_spinner=True)
def process_pptx(file_bytes):
    return extract_from_pptx(BytesIO(file_bytes))

@st.cache_resource(show_spinner=True)
def process_excel(file_bytes):
    return extract_from_excel(BytesIO(file_bytes))

uploaded_file = st.file_uploader(
    "Upload a PDF, Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) file",
    type=["pdf", "docx", "pptx", "xlsx"]
)

summary = {}
json_output = {}

def generate_download_button(data, filename, label):
    b64 = base64.b64encode(data.encode()).decode()
    href = f'<a href="data:application/json;base64,{b64}" download="{filename}">{label}</a>'
    return href

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    file_bytes = uploaded_file.read()

    with st.spinner("Processing file..."):
        if file_type == "pdf":
            pages = process_pdf(file_bytes)
            page_num = st.selectbox("Select Page", range(len(pages)))
            page = pages[page_num]
            st.subheader("Text Content")
            st.write(page["text"])
            st.subheader("Images")
            for img in page["images"]:
                st.image(img, use_container_width=True)
            st.subheader("Tables")
            for i, table in enumerate(page["tables"]):
                st.write(f"Table {i+1}")
                st.dataframe(table)

            summary = {
                "file_type": "PDF",
                "total_pages": len(pages),
                "page_selected": page_num,
                "text_chars": len(page["text"]),
                "image_count": len(page["images"]),
                "table_count": len(page["tables"]),
            }
            json_output = {
                "pages": [
                    {
                        "page": i,
                        "text": p["text"],
                        "tables": [df.to_dict() for df in p["tables"]],
                        "image_count": len(p["images"])
                    } for i, p in enumerate(pages)
                ]
            }

        elif file_type == "docx":
            content = process_docx(file_bytes)
            st.subheader("Text Content")
            st.write(content["text"])
            summary = {
                "file_type": "DOCX",
                "text_chars": len(content["text"]),
                "paragraphs": content["text"].count("\n") + 1
            }
            json_output = {
                "text": content["text"]
            }

        elif file_type == "pptx":
            slides = process_pptx(file_bytes)
            slide_num = st.selectbox("Select Slide", range(len(slides)))
            st.subheader("Slide Text")
            st.write(slides[slide_num]["text"])
            summary = {
                "file_type": "PPTX",
                "total_slides": len(slides),
                "selected_slide": slide_num
            }
            json_output = {
                "slides": slides
            }

        elif file_type == "xlsx":
            sheets = process_excel(file_bytes)
            sheet = st.selectbox("Select Sheet", list(sheets.keys()))
            st.subheader(f"Sheet: {sheet}")
            st.dataframe(sheets[sheet])
            summary = {
                "file_type": "XLSX",
                "sheets": list(sheets.keys()),
                "selected_sheet": sheet,
                "row_count": len(sheets[sheet])
            }
            json_output = {
                "sheets": {k: v.to_dict() for k, v in sheets.items()}
            }

        st.markdown("### 📋 Summary")
        st.json(summary)

        st.markdown("### 📦 Download Extracted JSON")
        json_str = json.dumps(json_output, indent=2)
        st.markdown(generate_download_button(json_str, "extracted_data.json", "📥 Download JSON"), unsafe_allow_html=True)
