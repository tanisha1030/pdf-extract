import streamlit as st
import pandas as pd
from main_code import extract_from_pdf, extract_from_docx, extract_from_pptx, extract_from_excel

st.set_page_config(page_title="Universal Info Extractor", layout="wide")
st.title("📄 Universal File Information Extractor")

uploaded_file = st.file_uploader("Upload a PDF, Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) file", type=["pdf", "docx", "pptx", "xlsx"])

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    if file_type == "pdf":
        pages = extract_from_pdf(uploaded_file)
        page_num = st.selectbox("Select Page", range(len(pages)))
        page = pages[page_num]
        st.subheader("Text Content")
        st.write(page["text"])
        st.subheader("Images")
        for img in page["images"]:
            st.image(img, use_column_width=True)
        st.subheader("Tables")
        for i, table in enumerate(page["tables"]):
            st.write(f"Table {i+1}")
            st.dataframe(table)
    elif file_type == "docx":
        content = extract_from_docx(uploaded_file)
        st.subheader("Text Content")
        st.write(content["text"])
    elif file_type == "pptx":
        slides = extract_from_pptx(uploaded_file)
        slide_num = st.selectbox("Select Slide", range(len(slides)))
        st.subheader("Slide Text")
        st.write(slides[slide_num]["text"])
    elif file_type == "xlsx":
        sheets = extract_from_excel(uploaded_file)
        sheet = st.selectbox("Select Sheet", list(sheets.keys()))
        st.subheader(f"Sheet: {sheet}")
        st.dataframe(sheets[sheet])
    else:
        st.error("Unsupported file format")