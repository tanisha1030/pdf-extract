import streamlit as st
import os
import json
import pandas as pd
import tempfile
import shutil
from datetime import datetime
import base64
from io import BytesIO
from pdfextract import ComprehensiveDocumentExtractor

# Page configuration
st.set_page_config(
    page_title="SmartDoc Extractor", 
    layout="wide",
    page_icon="📄"
)

# Enhanced CSS styling
st.markdown("""
    <style>
    .main-title {
        font-size: 2.5em;
        color: #2c3e50;
        margin-bottom: 0.5em;
        text-align: center;
        font-weight: bold;
    }
    .section-header {
        font-size: 1.5em;
        color: #1a73e8;
        margin-top: 1em;
        border-bottom: 2px solid #1a73e8;
        padding-bottom: 0.2em;
    }
    .info-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 1em;
        border-left: 4px solid #1a73e8;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .metrics-container {
        display: flex;
        justify-content: space-around;
        margin: 1em 0;
    }
    .metric-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        margin: 0 5px;
        flex: 1;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .content-container {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
        border: 1px solid #e9ecef;
    }
    .page-header {
        background: linear-gradient(135deg, #74b9ff 0%, #0984e3 100%);
        color: white;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        text-align: center;
    }
    .table-container {
        background-color: white;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .text-preview {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
        max-height: 200px;
        overflow-y: auto;
        font-family: 'Courier New', monospace;
        font-size: 0.9em;
    }
    .image-gallery {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
        gap: 20px;
        margin: 20px 0;
    }
    .image-item {
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .stAlert {
        margin-top: 1em;
    }
    .download-section {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_results' not in st.session_state:
    st.session_state.processed_results = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False

def cleanup_temp_files():
    """Clean up temporary files"""
    if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
        try:
            shutil.rmtree(st.session_state.temp_dir)
            st.session_state.temp_dir = None
        except Exception as e:
            st.warning(f"Could not clean up temporary files: {str(e)}")

def safe_get(obj, key, default=None):
    """Safely get a value from an object, handling both dict and non-dict types"""
    try:
        if isinstance(obj, dict):
            return obj.get(key, default)
        elif hasattr(obj, key):
            return getattr(obj, key, default)
        else:
            return default
    except Exception:
        return default

def ensure_dict(obj):
    """Ensure object is a dictionary or convert it to one"""
    if isinstance(obj, dict):
        return obj
    elif hasattr(obj, '__dict__'):
        return obj.__dict__
    else:
        return {'data': obj, 'status': 'unknown', 'error': 'Unexpected data format'}

def create_download_link(data, filename, label):
    """Create a download link for data"""
    if isinstance(data, dict):
        data = json.dumps(data, indent=2, ensure_ascii=False, default=str)
    
    b64 = base64.b64encode(data.encode()).decode()
    href = f'<a href="data:application/json;base64,{b64}" download="{filename}">{label}</a>'
    return href

def display_file_summary(data):
    """Display enhanced file summary information"""
    file_type = safe_get(data, 'file_type', 'Unknown')
    
    st.markdown(f"""
    <div class='info-box'>
        <h3>📊 File Summary</h3>
        <strong>File Type:</strong> {file_type}<br>
        <strong>File Size:</strong> {safe_get(data, 'file_size_mb', 0):.2f} MB<br>
        <strong>Processing Time:</strong> {safe_get(data, 'processing_time', 'N/A')}<br>
        <strong>Last Modified:</strong> {safe_get(data, 'last_modified', 'N/A')}
    </div>
    """, unsafe_allow_html=True)
    
    # Create metrics based on file type
    if file_type == 'PDF':
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📃 Pages", safe_get(data, 'page_count', 0))
        with col2:
            st.metric("📝 Words", f"{safe_get(data, 'total_words', 0):,}")
        with col3:
            st.metric("🖼️ Images", safe_get(data, 'total_images', 0))
        with col4:
            st.metric("📋 Tables", len(safe_get(data, 'extracted_tables', [])))
    
    elif file_type == 'DOCX':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Paragraphs", len(safe_get(data, 'paragraphs', [])))
        with col2:
            st.metric("📝 Words", f"{safe_get(data, 'total_words', 0):,}")
        with col3:
            st.metric("📋 Tables", len(safe_get(data, 'tables', [])))
    
    elif file_type == 'PPTX':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Slides", safe_get(data, 'total_slides', 0))
        with col2:
            st.metric("📝 Words", f"{safe_get(data, 'total_words', 0):,}")
        with col3:
            st.metric("📋 Tables", len(safe_get(data, 'tables', [])))
    
    elif file_type == 'Excel':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Sheets", len(safe_get(data, 'sheets', [])))
        with col2:
            sheets = safe_get(data, 'sheets', [])
            total_rows = sum(safe_get(sheet, 'shape', [0, 0])[0] for sheet in sheets if isinstance(sheet, dict))
            st.metric("📊 Total Rows", f"{total_rows:,}")
        with col3:
            sheets = safe_get(data, 'sheets', [])
            total_cols = max((safe_get(sheet, 'shape', [0, 0])[1] for sheet in sheets if isinstance(sheet, dict)), default=0)
            st.metric("📊 Total Columns", f"{total_cols:,}")

def display_text_content(text, title="Text Content", max_chars=1000):
    """Display text content with preview and full view options"""
    if not text or not str(text).strip():
        st.info("No text content available.")
        return
    
    text = str(text)  # Ensure it's a string
    st.markdown(f"**📝 {title}:**")
    
    # Show preview
    preview_text = text[:max_chars] + "..." if len(text) > max_chars else text
    st.markdown(f'<div class="text-preview">{preview_text}</div>', unsafe_allow_html=True)
    
    # Full text in expander
    if len(text) > max_chars:
        with st.expander(f"View full {title.lower()} ({len(text):,} characters)"):
            st.text_area(
                f"Full {title.lower()}", 
                value=text, 
                height=400,
                key=f"full_text_{hash(text)}"
            )
    
    # Text statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Characters", f"{len(text):,}")
    with col2:
        words = len(text.split())
        st.metric("Words", f"{words:,}")
    with col3:
        lines = len(text.split('\n'))
        st.metric("Lines", f"{lines:,}")

def display_pdf_content(data, file_path):
    """Display enhanced PDF content with better navigation"""
    pages = safe_get(data, 'pages', [])
    if not pages:
        st.warning("No page content found in this PDF.")
        return
    
    # Overall PDF statistics
    st.markdown("### 📊 PDF Overview")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Pages", safe_get(data, 'page_count', 0))
    with col2:
        st.metric("Total Words", f"{safe_get(data, 'total_words', 0):,}")
    with col3:
        st.metric("Total Images", safe_get(data, 'total_images', 0))
    with col4:
        st.metric("Total Tables", len(safe_get(data, 'extracted_tables', [])))
    
    # Page navigation
    st.markdown("### 📄 Page Navigation")
    page_count = safe_get(data, 'page_count', len(pages))
    page_numbers = list(range(1, page_count + 1))
    
    # Create page selector with more options
    col1, col2 = st.columns([3, 1])
    with col1:
        selected_page = st.selectbox(
            "Select a page to view its content:", 
            page_numbers, 
            key=f"page_selector_{file_path}"
        )
    with col2:
        if st.button("🔄 Refresh Page", key=f"refresh_{file_path}"):
            st.rerun()
    
    # Find selected page data
    selected_page_data = None
    for page in pages:
        if isinstance(page, dict) and safe_get(page, 'page_number') == selected_page:
            selected_page_data = page
            break
    
    if not selected_page_data:
        st.error(f"Could not find data for page {selected_page}")
        return
    
    # Display page content
    st.markdown(f'<div class="page-header">📄 Page {selected_page} Content</div>', unsafe_allow_html=True)
    
    # Page metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Words", safe_get(selected_page_data, 'word_count', 0))
    with col2:
        st.metric("Characters", safe_get(selected_page_data, 'char_count', 0))
    with col3:
        st.metric("Images", len(safe_get(selected_page_data, 'images', [])))
    with col4:
        st.metric("Tables", len(safe_get(selected_page_data, 'tables', [])))
    
    # Display text content
    page_text = safe_get(selected_page_data, 'text', '')
    if page_text and str(page_text).strip():
        display_text_content(page_text, f"Page {selected_page} Text")
    
    # Display images in gallery format
    images = safe_get(selected_page_data, 'images', [])
    if images:
        st.markdown("### 🖼️ Images Gallery")
        cols = st.columns(3)
        for i, img in enumerate(images):
            with cols[i % 3]:
                try:
                    img_filename = safe_get(img, 'filename', '')
                    if img_filename and os.path.exists(img_filename):
                        st.image(
                            img_filename, 
                            caption=f"Image {i+1}",
                            use_container_width=True
                        )
                        # Image details
                        st.markdown(f"**Filename:** {os.path.basename(img_filename)}")
                        img_size = safe_get(img, 'size', '')
                        if img_size:
                            st.markdown(f"**Size:** {img_size}")
                    else:
                        st.warning(f"Image file not found: {img_filename}")
                except Exception as e:
                    st.error(f"Error displaying image {i+1}: {str(e)}")
    
    # Display tables with enhanced formatting
    tables = safe_get(selected_page_data, 'tables', [])
    if tables:
        st.markdown("### 📋 Tables")
        for i, table in enumerate(tables):
            st.markdown(f'<div class="table-container">', unsafe_allow_html=True)
            st.markdown(f"**Table {i+1}**")
            
            try:
                table_data = safe_get(table, 'data', [])
                if table_data:
                    df = pd.DataFrame(table_data)
                    st.dataframe(df, use_container_width=True)
                    
                    # Table statistics
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Rows", len(df))
                    with col2:
                        st.metric("Columns", len(df.columns))
                else:
                    st.info("No data in this table.")
            except Exception as e:
                st.error(f"Error displaying table {i+1}: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)

def display_docx_content(data):
    """Display enhanced DOCX content"""
    st.markdown("### 📝 Document Content")
    
    # Document statistics
    paragraphs = safe_get(data, 'paragraphs', [])
    tables = safe_get(data, 'tables', [])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Paragraphs", len(paragraphs))
    with col2:
        st.metric("Tables", len(tables))
    with col3:
        st.metric("Total Words", f"{safe_get(data, 'total_words', 0):,}")
    
    # Display paragraphs with style information
    if paragraphs:
        st.markdown("#### 📄 Paragraphs")
        
        # Show paragraph preview
        preview_count = st.slider("Number of paragraphs to preview:", 1, min(20, len(paragraphs)), 5)
        
        for i, para in enumerate(paragraphs[:preview_count]):
            para_dict = ensure_dict(para)
            with st.expander(f"Paragraph {i+1} ({safe_get(para_dict, 'style', 'Normal')})", expanded=False):
                para_text = safe_get(para_dict, 'text', '')
                st.write(para_text)
                
                # Paragraph details
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Words", len(str(para_text).split()))
                with col2:
                    st.metric("Characters", len(str(para_text)))
        
        if len(paragraphs) > preview_count:
            st.info(f"Showing {preview_count} of {len(paragraphs)} paragraphs")
        
        # Full document text
        full_text = '\n\n'.join([str(safe_get(ensure_dict(para), 'text', '')) for para in paragraphs])
        display_text_content(full_text, "Full Document Text")
    
    # Display tables
    if tables:
        st.markdown("#### 📋 Tables")
        for i, table in enumerate(tables):
            st.markdown(f'<div class="table-container">', unsafe_allow_html=True)
            st.markdown(f"**Table {i+1}**")
            
            try:
                table_dict = ensure_dict(table)
                table_data = safe_get(table_dict, 'data', [])
                if table_data:
                    df = pd.DataFrame(table_data)
                    st.dataframe(df, use_container_width=True)
                    
                    # Export table option
                    csv = df.to_csv(index=False)
                    st.download_button(
                        label=f"📥 Download Table {i+1} as CSV",
                        data=csv,
                        file_name=f"table_{i+1}.csv",
                        mime="text/csv",
                        key=f"download_table_{i}"
                    )
                else:
                    st.info("No data in this table.")
            except Exception as e:
                st.error(f"Error displaying table {i+1}: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)

def display_pptx_content(data):
    """Display enhanced PPTX content"""
    st.markdown("### 📊 Presentation Content")
    
    slides = safe_get(data, 'slides', [])
    if not slides:
        st.warning("No slides found in this presentation.")
        return
    
    # Presentation statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Slides", len(slides))
    with col2:
        total_tables = sum(len(safe_get(ensure_dict(slide), 'tables', [])) for slide in slides)
        st.metric("Total Tables", total_tables)
    with col3:
        st.metric("Total Words", f"{safe_get(data, 'total_words', 0):,}")
    
    # Slide navigation
    slide_numbers = list(range(1, len(slides) + 1))
    selected_slide = st.selectbox("🎯 Select a slide to view:", slide_numbers)
    
    if selected_slide <= len(slides):
        slide_data = ensure_dict(slides[selected_slide - 1])
        
        st.markdown(f'<div class="page-header">📊 Slide {selected_slide}</div>', unsafe_allow_html=True)
        
        # Slide content
        text_content = safe_get(slide_data, 'text_content', [])
        if text_content:
            st.markdown("**📝 Text Content:**")
            slide_text = '\n'.join([f"• {content}" for content in text_content])
            display_text_content(slide_text, f"Slide {selected_slide} Text")
        
        # Slide tables
        tables = safe_get(slide_data, 'tables', [])
        if tables:
            st.markdown("**📋 Tables:**")
            for i, table in enumerate(tables):
                st.markdown(f'<div class="table-container">', unsafe_allow_html=True)
                st.markdown(f"**Table {i+1}**")
                
                try:
                    if isinstance(table, list) and table:
                        df = pd.DataFrame(table)
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.info("No data in this table.")
                except Exception as e:
                    st.error(f"Error displaying table {i+1}: {str(e)}")
                
                st.markdown('</div>', unsafe_allow_html=True)

def display_excel_content(data):
    """Display enhanced Excel content"""
    st.markdown("### 📊 Spreadsheet Content")
    
    sheets = safe_get(data, 'sheets', [])
    if not sheets:
        st.warning("No sheets found in this Excel file.")
        return
    
    # Excel statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Sheets", len(sheets))
    with col2:
        total_rows = sum(safe_get(ensure_dict(sheet), 'shape', [0, 0])[0] for sheet in sheets)
        st.metric("Total Rows", f"{total_rows:,}")
    with col3:
        total_cols = max((safe_get(ensure_dict(sheet), 'shape', [0, 0])[1] for sheet in sheets), default=0)
        st.metric("Max Columns", total_cols)
    
    # Sheet navigation
    sheet_names = [safe_get(ensure_dict(sheet), 'sheet_name', f'Sheet {i+1}') for i, sheet in enumerate(sheets)]
    selected_sheet = st.selectbox("📋 Select a sheet to view:", sheet_names)
    
    # Find selected sheet data
    sheet_data = None
    for sheet in sheets:
        sheet_dict = ensure_dict(sheet)
        if safe_get(sheet_dict, 'sheet_name') == selected_sheet:
            sheet_data = sheet_dict
            break
    
    if sheet_data:
        st.markdown(f'<div class="page-header">📋 Sheet: {selected_sheet}</div>', unsafe_allow_html=True)
        
        # Sheet statistics
        col1, col2, col3 = st.columns(3)
        with col1:
            shape = safe_get(sheet_data, 'shape', [0, 0])
            st.metric("Rows", f"{shape[0]:,}")
        with col2:
            st.metric("Columns", shape[1])
        with col3:
            data_records = safe_get(sheet_data, 'data', [])
            non_empty_cells = sum(1 for row in data_records for cell in row if cell) if data_records else 0
            st.metric("Non-empty Cells", f"{non_empty_cells:,}")
        
        # Display data with pagination
        if data_records:
            df = pd.DataFrame(data_records)
            
            # Data preview options
            col1, col2 = st.columns(2)
            with col1:
                rows_to_show = st.slider("Rows to display:", 10, min(500, len(df)), 50)
            with col2:
                show_all_cols = st.checkbox("Show all columns", value=False)
            
            # Display dataframe
            st.dataframe(df.head(rows_to_show), use_container_width=True)
            
            # Data summary
            st.markdown("**📊 Data Summary:**")
            st.write(df.describe())
            
            # Export options
            col1, col2 = st.columns(2)
            with col1:
                csv = df.to_csv(index=False)
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv,
                    file_name=f"{selected_sheet}.csv",
                    mime="text/csv"
                )
            with col2:
                json_data = df.to_json(orient='records', indent=2)
                st.download_button(
                    label="📥 Download as JSON",
                    data=json_data,
                    file_name=f"{selected_sheet}.json",
                    mime="application/json"
                )
        else:
            st.info("No data to display for this sheet.")
        
        # Column information
        columns = safe_get(sheet_data, 'columns', [])
        if columns:
            st.markdown("**📝 Column Information:**")
            cols_df = pd.DataFrame({
                'Column': columns,
                'Index': range(len(columns))
            })
            st.dataframe(cols_df, use_container_width=True)

def display_file_results(data, file_path):
    """Display enhanced results for a single file"""
    # Ensure data is a dictionary
    data = ensure_dict(data)
    
    # File header with enhanced styling
    filename = safe_get(data, 'filename', 'Unknown File')
    st.markdown(f"<div class='section-header'>📄 {filename}</div>", unsafe_allow_html=True)
    
    # Display file summary
    display_file_summary(data)
    
    # Content tabs for better organization
    file_type = safe_get(data, 'file_type', 'Unknown')
    
    # Create tabs for different content types
    content_tabs = ["📊 Overview", "📝 Content", "📥 Downloads"]
    tab1, tab2, tab3 = st.tabs(content_tabs)
    
    with tab1:
        st.markdown("### 📊 File Overview")
        # Display content based on file type
        if file_type == 'PDF':
            display_pdf_content(data, file_path)
        elif file_type == 'DOCX':
            display_docx_content(data)
        elif file_type == 'PPTX':
            display_pptx_content(data)
        elif file_type == 'Excel':
            display_excel_content(data)
    
    with tab2:
        st.markdown("### 📝 Raw Content")
        # Display raw extracted content
        raw_text = safe_get(data, 'raw_text', '')
        if raw_text:
            display_text_content(raw_text, "Raw Extracted Text")
        else:
            st.info("No raw text content available.")
    
    with tab3:
        st.markdown("### 📥 Download Options")
        
        # Individual file download options
        col1, col2 = st.columns(2)
        
        with col1:
            try:
                json_data = json.dumps(data, indent=2, ensure_ascii=False, default=str)
                st.download_button(
                    label=f"⬇️ Download {filename} JSON",
                    data=json_data,
                    file_name=f"{os.path.splitext(filename)[0]}_extraction.json",
                    mime="application/json",
                    key=f"download_{file_path}",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create JSON download: {str(e)}")
        
        with col2:
            # Download extracted text
            raw_text = safe_get(data, 'raw_text', '')
            if raw_text:
                st.download_button(
                    label="📝 Download Extracted Text",
                    data=str(raw_text),
                    file_name=f"{os.path.splitext(filename)[0]}_text.txt",
                    mime="text/plain",
                    key=f"download_text_{file_path}",
                    use_container_width=True
                )

# Main application
def main():
    # Enhanced header
    st.markdown("<div class='main-title'>📄 SmartDoc Extractor</div>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2em;'>
        <p style='font-size: 1.2em; color: #666;'>
            🚀 Upload PDF, Word, PowerPoint, or Excel files to extract text, tables, images, and metadata
        </p>
        <p style='color: #888;'>
            Powered by advanced document processing algorithms
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # File uploader with enhanced styling
    st.markdown("### 📁 File Upload")
    uploaded_files = st.file_uploader(
        "Choose your files:",
        type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Supported formats: PDF, DOCX, PPTX, XLSX, XLS (Max 200MB per file)"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Display uploaded files info
        st.markdown("#### 📋 Uploaded Files:")
        for i, file in enumerate(uploaded_files):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.write(f"**{i+1}.** {file.name}")
            with col2:
                file_size_mb = file.size / (1024 * 1024)
                st.write(f"📏 {file_size_mb:.2f} MB")
            with col3:
                file_type = file.name.split('.')[-1].upper()
                st.write(f"📄 {file_type}")
        
        # Processing options
        st.markdown("### ⚙️ Processing Options")
        
        col1, col2 = st.columns(2)
        with col1:
            extract_images = st.checkbox("🖼️ Extract Images", value=True, help="Extract and save images from documents")
            extract_tables = st.checkbox("📋 Extract Tables", value=True, help="Extract tables and convert to structured data")
        with col2:
            extract_metadata = st.checkbox("📊 Extract Metadata", value=True, help="Extract document metadata and properties")
            high_quality = st.checkbox("🎯 High Quality Mode", value=False, help="Use advanced processing (slower but better results)")
        
        # Processing button
        if st.button("🚀 Process Files", type="primary", use_container_width=True):
            if not st.session_state.processing_complete:
                # Clean up previous temp files
                cleanup_temp_files()
                
                # Create new temp directory
                st.session_state.temp_dir = tempfile.mkdtemp()
                
                # Initialize progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                all_results = {}
                
                try:
                    # Process each file
                    for i, uploaded_file in enumerate(uploaded_files):
                        status_text.text(f"Processing {uploaded_file.name}... ({i+1}/{len(uploaded_files)})")
                        
                        # Save uploaded file to temp directory
                        temp_file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
                        with open(temp_file_path, 'wb') as f:
                            f.write(uploaded_file.read())
                        
                        # Initialize extractor
                        extractor = ComprehensiveDocumentExtractor()
                        
                        # Set processing options
                        processing_options = {
                            'extract_images': extract_images,
                            'extract_tables': extract_tables,
                            'extract_metadata': extract_metadata,
                            'high_quality_mode': high_quality,
                            'output_dir': st.session_state.temp_dir
                        }
                        
                        try:
                            # Process the file
                            result = extractor.process_file(temp_file_path, **processing_options)
                            
                            # Ensure result is properly formatted
                            if result:
                                result_dict = ensure_dict(result)
                                result_dict['filename'] = uploaded_file.name
                                result_dict['file_size_mb'] = uploaded_file.size / (1024 * 1024)
                                result_dict['processing_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                all_results[temp_file_path] = result_dict
                            else:
                                st.error(f"❌ Failed to process {uploaded_file.name}: No result returned")
                                
                        except Exception as e:
                            st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
                            continue
                        
                        # Update progress
                        progress_bar.progress((i + 1) / len(uploaded_files))
                    
                    # Store results in session state
                    st.session_state.processed_results = all_results
                    st.session_state.processing_complete = True
                    
                    # Clear progress indicators
                    progress_bar.empty()
                    status_text.empty()
                    
                    if all_results:
                        st.success(f"✅ Successfully processed {len(all_results)} file(s)!")
                        st.balloons()
                    else:
                        st.error("❌ No files were processed successfully.")
                        
                except Exception as e:
                    st.error(f"❌ Critical error during processing: {str(e)}")
                    progress_bar.empty()
                    status_text.empty()
    
    # Display results if available
    if st.session_state.processed_results:
        st.markdown("---")
        st.markdown("## 📊 Extraction Results")
        
        # Results summary
        total_files = len(st.session_state.processed_results)
        st.markdown(f"""
        <div class='info-box'>
            <h3>🎯 Processing Summary</h3>
            <strong>Files Processed:</strong> {total_files}<br>
            <strong>Processing Status:</strong> ✅ Complete<br>
            <strong>Results Available:</strong> Yes
        </div>
        """, unsafe_allow_html=True)
        
        # Global download options
        st.markdown("### 📥 Global Download Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Download all results as JSON
            try:
                all_json_data = json.dumps(st.session_state.processed_results, indent=2, ensure_ascii=False, default=str)
                st.download_button(
                    label="📦 Download All Results (JSON)",
                    data=all_json_data,
                    file_name=f"extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    key="download_all_json",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create global JSON download: {str(e)}")
        
        with col2:
            # Download all extracted text
            try:
                all_text = ""
                for file_path, data in st.session_state.processed_results.items():
                    data_dict = ensure_dict(data)
                    filename = safe_get(data_dict, 'filename', 'Unknown')
                    raw_text = safe_get(data_dict, 'raw_text', '')
                    all_text += f"\n\n{'='*50}\n"
                    all_text += f"FILE: {filename}\n"
                    all_text += f"{'='*50}\n\n"
                    all_text += str(raw_text)
                
                if all_text.strip():
                    st.download_button(
                        label="📝 Download All Text",
                        data=all_text,
                        file_name=f"all_extracted_text_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain",
                        key="download_all_text",
                        use_container_width=True
                    )
                else:
                    st.info("No text content to download")
            except Exception as e:
                st.error(f"Could not create global text download: {str(e)}")
        
        with col3:
            # Reset/Clear results
            if st.button("🔄 Clear Results", key="clear_results", use_container_width=True):
                st.session_state.processed_results = None
                st.session_state.processing_complete = False
                cleanup_temp_files()
                st.rerun()
        
        # Display individual file results
        st.markdown("### 📋 Individual File Results")
        
        # File selector for detailed view
        file_options = []
        for file_path, data in st.session_state.processed_results.items():
            data_dict = ensure_dict(data)
            filename = safe_get(data_dict, 'filename', os.path.basename(file_path))
            file_options.append((filename, file_path))
        
        if file_options:
            selected_file_name, selected_file_path = st.selectbox(
                "🎯 Select a file to view detailed results:",
                file_options,
                format_func=lambda x: x[0]
            )
            
            if selected_file_path in st.session_state.processed_results:
                file_data = st.session_state.processed_results[selected_file_path]
                display_file_results(file_data, selected_file_path)
        
        # Processing statistics
        st.markdown("---")
        st.markdown("### 📊 Processing Statistics")
        
        try:
            # Calculate statistics
            total_pages = 0
            total_words = 0
            total_images = 0
            total_tables = 0
            
            for file_path, data in st.session_state.processed_results.items():
                data_dict = ensure_dict(data)
                total_pages += safe_get(data_dict, 'page_count', 0)
                total_words += safe_get(data_dict, 'total_words', 0)
                total_images += safe_get(data_dict, 'total_images', 0)
                
                # Count tables from different sources
                if safe_get(data_dict, 'file_type') == 'PDF':
                    total_tables += len(safe_get(data_dict, 'extracted_tables', []))
                else:
                    total_tables += len(safe_get(data_dict, 'tables', []))
            
            # Display statistics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(
                    label="📄 Total Pages",
                    value=f"{total_pages:,}",
                    help="Total number of pages processed across all files"
                )
            
            with col2:
                st.metric(
                    label="📝 Total Words",
                    value=f"{total_words:,}",
                    help="Total number of words extracted from all files"
                )
            
            with col3:
                st.metric(
                    label="🖼️ Total Images",
                    value=f"{total_images:,}",
                    help="Total number of images extracted from all files"
                )
            
            with col4:
                st.metric(
                    label="📋 Total Tables",
                    value=f"{total_tables:,}",
                    help="Total number of tables extracted from all files"
                )
            
        except Exception as e:
            st.error(f"Error calculating statistics: {str(e)}")
    
    # Footer with information
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; margin-top: 2em; padding: 1em; background-color: #f8f9fa; border-radius: 10px;'>
        <p style='color: #666; margin: 0;'>
            🔒 Your files are processed locally and securely. No data is stored permanently.
        </p>
        <p style='color: #888; margin: 0.5em 0 0 0; font-size: 0.9em;'>
            Supported formats: PDF, DOCX, PPTX, XLSX, XLS | Max file size: 200MB per file
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Cleanup on app exit
    if st.session_state.temp_dir:
        # Register cleanup function
        import atexit
        atexit.register(cleanup_temp_files)

# Run the application
if __name__ == "__main__":
    main()