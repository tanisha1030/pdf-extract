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

            # Build summary data
            summary_data = []
            for idx, page in enumerate(pages):
                text = page["text"]
                num_words = len(text.strip().split())
                num_chars = len(text.strip())
                num_tables = len(page["tables"])
                num_images = len(page["images"])
                summary_data.append({
                    "Page No": idx + 1,
                    "# of Words": num_words,
                    "# of Characters": num_chars,
                    "# of Tables": num_tables,
                    "# of Images": num_images,
                })

            df_summary = pd.DataFrame(summary_data)

            st.subheader("📊 Document Structure Summary Table")
            st.dataframe(df_summary)

            # Excel Download
            @st.cache_data
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Summary")
                processed_data = output.getvalue()
                return processed_data

            excel_data = convert_df_to_excel(df_summary)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=excel_data,
                file_name="document_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Page selector
            page_num = st.selectbox("Select Page to View Details", range(len(pages)), format_func=lambda x: f"Page {x+1}")

            page = pages[page_num]
            st.subheader("📝 Text Content")
            st.write(page["text"])

            st.subheader("🖼️ Images")
            for i, img in enumerate(page["images"]):
                st.image(img, caption=f"Page {page_num +1} - Image {i+1}", use_container_width=True)
                img_buffer = BytesIO()
                img.save(img_buffer, format="PNG")
                st.download_button(
                    label=f"Download Image {i+1}",
                    data=img_buffer.getvalue(),
                    file_name=f"page_{page_num+1}_image_{i+1}.png",
                    mime="image/png"
                )

            st.subheader("📈 Tables")
            for i, table in enumerate(page["tables"]):
                st.write(f"Table {i+1}")
                st.dataframe(table)
                csv_data = table.to_csv(index=False).encode()
                st.download_button(
                    label=f"Download Table {i+1} as CSV",
                    data=csv_data,
                    file_name=f"page_{page_num+1}_table_{i+1}.csv",
                    mime="text/csv"
                )

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
