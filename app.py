
import streamlit as st
import os
import json
import pandas as pd
from pdextract import ComprehensiveDocumentExtractor

st.set_page_config(page_title="SmartDoc Extractor", layout="wide")
extractor = ComprehensiveDocumentExtractor()

st.title("📄 SmartDoc Extractor")
st.write("Upload PDF, Word, PowerPoint, or Excel files to extract text, tables, images, and metadata.")

uploaded_files = st.file_uploader(
    "Upload your files here:",
    type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("Processing files..."):
        saved_file_paths = []
        for uploaded_file in uploaded_files:
            file_path = os.path.join(uploaded_file.name)
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.read())
            saved_file_paths.append(file_path)

        results = extractor.process_files(saved_file_paths)

        if results:
            st.success(f"✅ Processed {len(results)} file(s) successfully!")

            for file_path, data in results.items():
                st.subheader(f"📄 {data['filename']}")
                st.write(f"**Type**: {data['file_type']}")
                st.write(f"**Size**: {data['file_size_mb']} MB")

                if data['file_type'] == 'PDF':
                    st.write(f"**Pages**: {data['page_count']}")
                    st.write(f"**Words**: {data['total_words']}")
                    st.write(f"**Characters**: {data['total_characters']}")
                    st.write(f"**Images**: {data['total_images']}")
                    st.write(f"**Tables**: {len(data.get('extracted_tables', []))}")

                elif data['file_type'] == 'DOCX':
                    st.write(f"**Paragraphs**: {len(data['paragraphs'])}")
                    st.write(f"**Words**: {data['total_words']}")
                    st.write(f"**Tables**: {len(data['tables'])}")

                elif data['file_type'] == 'PPTX':
                    st.write(f"**Slides**: {data['total_slides']}")
                    st.write(f"**Words**: {data['total_words']}")

                elif data['file_type'] == 'Excel':
                    st.write(f"**Sheets**: {len(data['sheets'])}")
                    for sheet in data['sheets']:
                        st.write(f"  • {sheet['sheet_name']} ({sheet['shape'][0]} rows × {sheet['shape'][1]} cols)")

                json_data = json.dumps(data, indent=2, ensure_ascii=False)
                st.download_button(
                    label="⬇️ Download Extraction JSON",
                    data=json_data,
                    file_name=f"{data['filename']}_extraction.json",
                    mime="application/json"
                )

            extractor.create_summary_report(results)
            with open("extraction_summary.txt", "r", encoding="utf-8") as f:
                summary_text = f.read()

            st.download_button(
                label="📄 Download Summary Report",
                data=summary_text,
                file_name="extraction_summary.txt",
                mime="text/plain"
            )
        else:
            st.error("❌ No files could be processed.")
else:
    st.info("⬆️ Upload one or more supported files to begin.")
