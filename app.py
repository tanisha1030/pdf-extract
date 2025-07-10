import streamlit as st
import pandas as pd
from main_code import extract_from_pdf, extract_from_docx, extract_from_pptx, extract_from_excel
from io import BytesIO

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

        elif file_type == "docx":
            content = process_docx(file_bytes)
            st.subheader("Text Content")
            st.write(content["text"])

        elif file_type == "pptx":
            slides = process_pptx(file_bytes)
            slide_num = st.selectbox("Select Slide", range(len(slides)))
            st.subheader("Slide Text")
            st.write(slides[slide_num]["text"])

        elif file_type == "xlsx":
            sheets = process_excel(file_bytes)
            sheet = st.selectbox("Select Sheet", list(sheets.keys()))
            st.subheader(f"Sheet: {sheet}")
            st.dataframe(sheets[sheet])

        else:
            st.error("Unsupported file format")