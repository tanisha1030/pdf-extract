import streamlit as st
import os
import json
import pandas as pd
from pdextract import ComprehensiveDocumentExtractor

st.set_page_config(page_title="SmartDoc Extractor", layout="wide")
st.markdown("""
    <style>
    .main-title {
        font-size: 2.5em;
        color: #3c3c3c;
        margin-bottom: 0.5em;
    }
    .section-header {
        font-size: 1.5em;
        color: #1a73e8;
        margin-top: 1em;
    }
    .info-box {
        background-color: #f1f3f4;
        padding: 10px;
        border-radius: 8px;
        margin-bottom: 1em;
    }
    </style>
""", unsafe_allow_html=True)

extractor = ComprehensiveDocumentExtractor()

st.markdown("<div class='main-title'>📄 SmartDoc Extractor</div>", unsafe_allow_html=True)
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
                st.markdown(f"<div class='section-header'>📄 {data['filename']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div class='info-box'>**Type**: {data['file_type']}<br>**Size**: {data['file_size_mb']} MB</div>", unsafe_allow_html=True)

                if data['file_type'] == 'PDF':
                    st.write(f"**Pages**: {data['page_count']}")
                    st.write(f"**Words**: {data['total_words']}")
                    st.write(f"**Characters**: {data['total_characters']}")
                    st.write(f"**Images**: {data['total_images']}")
                    st.write(f"**Tables**: {len(data.get('extracted_tables', []))}")

                    page_numbers = list(range(1, data['page_count'] + 1))
                    selected_page = st.selectbox("📄 Select a page to view its content:", page_numbers, key=file_path)
                    selected_page_data = next((p for p in data['pages'] if p['page_number'] == selected_page), None)

                    if selected_page_data:
                        st.markdown(f"### 🧾 Page {selected_page}")
                        st.markdown(f"**Word Count**: {selected_page_data['word_count']}")
                        st.markdown(f"**Character Count**: {selected_page_data['char_count']}")
                        st.markdown("**Text:**")
                        st.code(selected_page_data['text'], language='text')

                        if selected_page_data['images']:
                            st.markdown("**Images on this page:**")
                            for img in selected_page_data['images']:
                                st.image(img['filename'], caption=os.path.basename(img['filename']), use_column_width=True)

                        if selected_page_data['tables']:
                            st.markdown("**Tables on this page:**")
                            for table in selected_page_data['tables']:
                                try:
                                    df = pd.DataFrame([row.split() for row in table['data']])
                                    st.table(df)
                                except:
                                    st.write(table['data'])

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
