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
    """Display summary table"""
    st.subheader("📊 Document Structure Summary")
    
    # Display the summary table
    st.dataframe(summary_df, use_container_width=True)

def to_excel(df):
    """Convert DataFrame to Excel bytes"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Document_Summary')
    return output.getvalue()

def create_json_summary(content, file_type, summary_df, filename):
    """Create comprehensive JSON summary of the document"""
    json_data = {
        "file_info": {
            "filename": filename,
            "file_type": file_type,
            "processed_at": pd.Timestamp.now().isoformat(),
            "total_pages": len(summary_df)
        },
        "summary_statistics": {
            "total_words": int(summary_df["# of words in page"].sum()),
            "total_characters": int(summary_df["# of characters in page"].sum()),
            "total_tables": int(summary_df["# of tables in page"].sum()),
            "total_images": int(summary_df["# of images in page"].sum())
        },
        "page_details": summary_df.to_dict('records'),
        "content": {}
    }
    
    if file_type == "pdf":
        json_data["content"]["pages"] = []
        for i, page in enumerate(content):
            page_data = {
                "page_number": i + 1,
                "text": page["text"],
                "image_count": len(page["images"]),
                "table_count": len(page["tables"]),
                "tables": []
            }
            
            # Add table data
            for j, table in enumerate(page["tables"]):
                table_data = {
                    "table_number": j + 1,
                    "data": table.to_dict('records') if not table.empty else []
                }
                page_data["tables"].append(table_data)
            
            json_data["content"]["pages"].append(page_data)
    
    elif file_type == "pptx":
        json_data["content"]["slides"] = []
        for i, slide in enumerate(content):
            slide_data = {
                "slide_number": i + 1,
                "text": slide["text"]
            }
            json_data["content"]["slides"].append(slide_data)
    
    elif file_type == "docx":
        json_data["content"]["document"] = {
            "text": content["text"]
        }
    
    elif file_type == "xlsx":
        json_data["content"]["sheets"] = {}
        for sheet_name, df in content.items():
            json_data["content"]["sheets"][sheet_name] = {
                "data": df.to_dict('records'),
                "shape": list(df.shape),
                "columns": list(df.columns)
            }
    
    return json_data

def to_json(json_data):
    """Convert JSON data to bytes"""
    import json
    return json.dumps(json_data, indent=2, ensure_ascii=False).encode('utf-8')

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
            
            # Create and display summary table with navigation
            summary_df = create_summary_table(pages, file_type)
            display_clickable_summary(summary_df, file_type, pages)
            
            # Create JSON summary
            json_summary = create_json_summary(pages, file_type, summary_df, uploaded_file.name)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                excel_data = to_excel(summary_df)
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=excel_data,
                    file_name=f"{uploaded_file.name}_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                json_data = to_json(json_summary)
                st.download_button(
                    label="📄 Download Complete Data as JSON",
                    data=json_data,
                    file_name=f"{uploaded_file.name}_complete.json",
                    mime="application/json"
                )
            
            # Use session state for page selection
            page_num = st.selectbox("Select Page", range(len(pages)), index=st.session_state.selected_page)
            st.session_state.selected_page = page_num
            
            page = pages[page_num]
            st.subheader(f"📄 Page {page_num + 1} Content")
            
            # Create tabs for better organization
            tab1, tab2, tab3 = st.tabs(["Text Content", "Images", "Tables"])
            
            with tab1:
                if page["text"]:
                    st.write(page["text"])
                else:
                    st.write("No text content found on this page.")
            
            with tab2:
                if page["images"]:
                    for i, img in enumerate(page["images"]):
                        st.write(f"**Image {i+1}:**")
                        st.image(img, use_container_width=True)
                else:
                    st.write("No images found on this page.")
            
            with tab3:
                if page["tables"]:
                    for i, table in enumerate(page["tables"]):
                        st.write(f"**Table {i+1}:**")
                        st.dataframe(table, use_container_width=True)
                else:
                    st.write("No tables found on this page.")

        elif file_type == "docx":
            content = process_docx(file_bytes)
            
            # Create and display summary table
            summary_df = create_summary_table(content, file_type)
            display_clickable_summary(summary_df, file_type, content)
            
            # Create JSON summary
            json_summary = create_json_summary(content, file_type, summary_df, uploaded_file.name)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                excel_data = to_excel(summary_df)
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=excel_data,
                    file_name=f"{uploaded_file.name}_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                json_data = to_json(json_summary)
                st.download_button(
                    label="📄 Download Complete Data as JSON",
                    data=json_data,
                    file_name=f"{uploaded_file.name}_complete.json",
                    mime="application/json"
                )
            
            st.subheader("📄 Document Content")
            st.write(content["text"])

        elif file_type == "pptx":
            slides = process_pptx(file_bytes)
            
            # Create and display summary table with navigation
            summary_df = create_summary_table(slides, file_type)
            display_clickable_summary(summary_df, file_type, slides)
            
            # Create JSON summary
            json_summary = create_json_summary(slides, file_type, summary_df, uploaded_file.name)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                excel_data = to_excel(summary_df)
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=excel_data,
                    file_name=f"{uploaded_file.name}_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                json_data = to_json(json_summary)
                st.download_button(
                    label="📄 Download Complete Data as JSON",
                    data=json_data,
                    file_name=f"{uploaded_file.name}_complete.json",
                    mime="application/json"
                )
            
            # Use session state for slide selection
            slide_num = st.selectbox("Select Slide", range(len(slides)), index=st.session_state.selected_slide)
            st.session_state.selected_slide = slide_num
            
            st.subheader(f"🎞️ Slide {slide_num + 1} Content")
            st.write(slides[slide_num]["text"])

        elif file_type == "xlsx":
            sheets = process_excel(file_bytes)
            
            # Create and display summary table with navigation
            summary_df = create_summary_table(sheets, file_type)
            display_clickable_summary(summary_df, file_type, sheets)
            
            # Create JSON summary
            json_summary = create_json_summary(sheets, file_type, summary_df, uploaded_file.name)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                excel_data = to_excel(summary_df)
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=excel_data,
                    file_name=f"{uploaded_file.name}_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                json_data = to_json(json_summary)
                st.download_button(
                    label="📄 Download Complete Data as JSON",
                    data=json_data,
                    file_name=f"{uploaded_file.name}_complete.json",
                    mime="application/json"
                )
            
            # Use session state for sheet selection or default to first sheet
            sheet_options = list(sheets.keys())
            if st.session_state.selected_sheet and st.session_state.selected_sheet in sheet_options:
                default_index = sheet_options.index(st.session_state.selected_sheet)
            else:
                default_index = 0
            
            sheet = st.selectbox("Select Sheet", sheet_options, index=default_index)
            st.session_state.selected_sheet = sheet
            
            st.subheader(f"📊 Sheet: {sheet}")
            st.dataframe(sheets[sheet], use_container_width=True)

        else:
            st.error("Unsupported file format")
