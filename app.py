import streamlit as st
import pandas as pd
from main_code import extract_from_pdf, extract_from_docx, extract_from_pptx, extract_from_excel
from io import BytesIO

st.set_page_config(page_title="Universal Info Extractor", layout="wide")
st.title("📄 Universal File Information Extractor")

# Initialize session state for navigation
if 'selected_page' not in st.session_state:
    st.session_state.selected_page = 0
if 'selected_slide' not in st.session_state:
    st.session_state.selected_slide = 0
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None

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

def create_summary_table(content, file_type):
    """Create document structure summary table with clickable links"""
    summary_data = []
    
    if file_type == "pdf":
        for i, page in enumerate(content):
            word_count = len(page["text"].split()) if page["text"] else 0
            char_count = len(page["text"]) if page["text"] else 0
            table_count = len(page["tables"])
            image_count = len(page["images"])
            
            summary_data.append({
                "Page No": i + 1,
                "# of words in page": word_count,
                "# of characters in page": char_count,
                "# of tables in page": table_count,
                "# of images in page": image_count
            })
    
    elif file_type == "pptx":
        for i, slide in enumerate(content):
            word_count = len(slide["text"].split()) if slide["text"] else 0
            char_count = len(slide["text"]) if slide["text"] else 0
            
            summary_data.append({
                "Page No": i + 1,
                "# of words in page": word_count,
                "# of characters in page": char_count,
                "# of tables in page": 0,
                "# of images in page": 0
            })
    
    elif file_type == "docx":
        word_count = len(content["text"].split()) if content["text"] else 0
        char_count = len(content["text"]) if content["text"] else 0
        
        summary_data.append({
            "Page No": 1,
            "# of words in page": word_count,
            "# of characters in page": char_count,
            "# of tables in page": 0,
            "# of images in page": 0
        })
    
    elif file_type == "xlsx":
        for i, (sheet_name, df) in enumerate(content.items()):
            word_count = df.astype(str).apply(lambda x: x.str.split().str.len()).sum().sum()
            char_count = df.astype(str).apply(lambda x: x.str.len()).sum().sum()
            
            summary_data.append({
                "Page No": f"Sheet: {sheet_name}",
                "# of words in page": word_count,
                "# of characters in page": char_count,
                "# of tables in page": 1,
                "# of images in page": 0
            })
    
    return pd.DataFrame(summary_data)

def display_clickable_summary(summary_df, file_type, content=None):
    """Display summary table with clickable navigation"""
    st.subheader("📊 Document Structure Summary")
    
    # Display the summary table
    st.dataframe(summary_df, use_container_width=True)
    
    # Create navigation buttons based on file type
    if file_type == "pdf":
        st.write("**Quick Navigation:**")
        cols = st.columns(min(5, len(summary_df)))
        for i, row in summary_df.iterrows():
            with cols[i % 5]:
                if st.button(f"Page {row['Page No']}", key=f"nav_page_{i}"):
                    st.session_state.selected_page = i
                    st.rerun()
    
    elif file_type == "pptx":
        st.write("**Quick Navigation:**")
        cols = st.columns(min(5, len(summary_df)))
        for i, row in summary_df.iterrows():
            with cols[i % 5]:
                if st.button(f"Slide {row['Page No']}", key=f"nav_slide_{i}"):
                    st.session_state.selected_slide = i
                    st.rerun()
    
    elif file_type == "xlsx":
        st.write("**Quick Navigation:**")
        cols = st.columns(min(4, len(summary_df)))
        for i, row in summary_df.iterrows():
            sheet_name = row['Page No'].replace('Sheet: ', '')
            with cols[i % 4]:
                if st.button(f"{sheet_name}", key=f"nav_sheet_{i}"):
                    st.session_state.selected_sheet = sheet_name
                    st.rerun()

def to_excel(df):
    """Convert DataFrame to Excel bytes"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Document_Summary')
    return output.getvalue()

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
            
            # Create and display summary table
            summary_df = create_summary_table(pages, file_type)
            st.subheader("📊 Document Structure Summary")
            st.dataframe(summary_df, use_container_width=True)
            
            # Download button for summary
            excel_data = to_excel(summary_df)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=excel_data,
                file_name=f"{uploaded_file.name}_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
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
            
            # Create and display summary table
            summary_df = create_summary_table(content, file_type)
            st.subheader("📊 Document Structure Summary")
            st.dataframe(summary_df, use_container_width=True)
            
            # Download button for summary
            excel_data = to_excel(summary_df)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=excel_data,
                file_name=f"{uploaded_file.name}_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.subheader("Text Content")
            st.write(content["text"])

        elif file_type == "pptx":
            slides = process_pptx(file_bytes)
            
            # Create and display summary table
            summary_df = create_summary_table(slides, file_type)
            st.subheader("📊 Document Structure Summary")
            st.dataframe(summary_df, use_container_width=True)
            
            # Download button for summary
            excel_data = to_excel(summary_df)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=excel_data,
                file_name=f"{uploaded_file.name}_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            slide_num = st.selectbox("Select Slide", range(len(slides)))
            st.subheader("Slide Text")
            st.write(slides[slide_num]["text"])

        elif file_type == "xlsx":
            sheets = process_excel(file_bytes)
            
            # Create and display summary table
            summary_df = create_summary_table(sheets, file_type)
            st.subheader("📊 Document Structure Summary")
            st.dataframe(summary_df, use_container_width=True)
            
            # Download button for summary
            excel_data = to_excel(summary_df)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=excel_data,
                file_name=f"{uploaded_file.name}_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            sheet = st.selectbox("Select Sheet", list(sheets.keys()))
            st.subheader(f"Sheet: {sheet}")
            st.dataframe(sheets[sheet])

        else:
            st.error("Unsupported file format")
