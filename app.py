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

def create_download_link(data, filename, label):
    """Create a download link for data"""
    if isinstance(data, dict):
        data = json.dumps(data, indent=2, ensure_ascii=False, default=str)
    
    b64 = base64.b64encode(data.encode()).decode()
    href = f'<a href="data:application/json;base64,{b64}" download="{filename}">{label}</a>'
    return href

def display_file_summary(data):
    """Display enhanced file summary information"""
    file_type = data.get('file_type', 'Unknown')
    
    st.markdown(f"""
    <div class='info-box'>
        <h3>📊 File Summary</h3>
        <strong>File Type:</strong> {file_type}<br>
        <strong>File Size:</strong> {data.get('file_size_mb', 0):.2f} MB<br>
        <strong>Processing Time:</strong> {data.get('processing_time', 'N/A')}<br>
        <strong>Last Modified:</strong> {data.get('last_modified', 'N/A')}
    </div>
    """, unsafe_allow_html=True)
    
    # Create metrics based on file type
    if file_type == 'PDF':
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📃 Pages", data.get('page_count', 0))
        with col2:
            st.metric("📝 Words", f"{data.get('total_words', 0):,}")
        with col3:
            st.metric("🖼️ Images", data.get('total_images', 0))
        with col4:
            st.metric("📋 Tables", len(data.get('extracted_tables', [])))
    
    elif file_type == 'DOCX':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Paragraphs", len(data.get('paragraphs', [])))
        with col2:
            st.metric("📝 Words", f"{data.get('total_words', 0):,}")
        with col3:
            st.metric("📋 Tables", len(data.get('tables', [])))
    
    elif file_type == 'PPTX':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Slides", data.get('total_slides', 0))
        with col2:
            st.metric("📝 Words", f"{data.get('total_words', 0):,}")
        with col3:
            st.metric("📋 Tables", len(data.get('tables', [])))
    
    elif file_type == 'Excel':
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📃 Sheets", len(data.get('sheets', [])))
        with col2:
            total_rows = sum(sheet.get('shape', [0, 0])[0] for sheet in data.get('sheets', []))
            st.metric("📊 Total Rows", f"{total_rows:,}")
        with col3:
            total_cols = sum(sheet.get('shape', [0, 0])[1] for sheet in data.get('sheets', []))
            st.metric("📊 Total Columns", f"{total_cols:,}")

def display_text_content(text, title="Text Content", max_chars=1000):
    """Display text content with preview and full view options"""
    if not text or not text.strip():
        st.info("No text content available.")
        return
    
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
    if not data.get('pages'):
        st.warning("No page content found in this PDF.")
        return
    
    # Overall PDF statistics
    st.markdown("### 📊 PDF Overview")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Pages", data.get('page_count', 0))
    with col2:
        st.metric("Total Words", f"{data.get('total_words', 0):,}")
    with col3:
        st.metric("Total Images", data.get('total_images', 0))
    with col4:
        st.metric("Total Tables", len(data.get('extracted_tables', [])))
    
    # Page navigation
    st.markdown("### 📄 Page Navigation")
    page_numbers = list(range(1, data['page_count'] + 1))
    
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
    selected_page_data = next(
        (p for p in data['pages'] if p['page_number'] == selected_page), 
        None
    )
    
    if not selected_page_data:
        st.error(f"Could not find data for page {selected_page}")
        return
    
    # Display page content
    st.markdown(f'<div class="page-header">📄 Page {selected_page} Content</div>', unsafe_allow_html=True)
    
    # Page metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Words", selected_page_data.get('word_count', 0))
    with col2:
        st.metric("Characters", selected_page_data.get('char_count', 0))
    with col3:
        st.metric("Images", len(selected_page_data.get('images', [])))
    with col4:
        st.metric("Tables", len(selected_page_data.get('tables', [])))
    
    # Display text content
    page_text = selected_page_data.get('text', '').strip()
    if page_text:
        display_text_content(page_text, f"Page {selected_page} Text")
    
    # Display images in gallery format
    if selected_page_data.get('images'):
        st.markdown("### 🖼️ Images Gallery")
        cols = st.columns(3)
        for i, img in enumerate(selected_page_data['images']):
            with cols[i % 3]:
                try:
                    if os.path.exists(img['filename']):
                        st.image(
                            img['filename'], 
                            caption=f"Image {i+1}",
                            use_container_width=True
                        )
                        # Image details
                        st.markdown(f"**Filename:** {os.path.basename(img['filename'])}")
                        if 'size' in img:
                            st.markdown(f"**Size:** {img['size']}")
                    else:
                        st.warning(f"Image file not found: {img['filename']}")
                except Exception as e:
                    st.error(f"Error displaying image {i+1}: {str(e)}")
    
    # Display tables with enhanced formatting
    if selected_page_data.get('tables'):
        st.markdown("### 📋 Tables")
        for i, table in enumerate(selected_page_data['tables']):
            st.markdown(f'<div class="table-container">', unsafe_allow_html=True)
            st.markdown(f"**Table {i+1}**")
            
            try:
                table_data = table.get('data', [])
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
    paragraphs = data.get('paragraphs', [])
    tables = data.get('tables', [])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Paragraphs", len(paragraphs))
    with col2:
        st.metric("Tables", len(tables))
    with col3:
        st.metric("Total Words", f"{data.get('total_words', 0):,}")
    
    # Display paragraphs with style information
    if paragraphs:
        st.markdown("#### 📄 Paragraphs")
        
        # Show paragraph preview
        preview_count = st.slider("Number of paragraphs to preview:", 1, min(20, len(paragraphs)), 5)
        
        for i, para in enumerate(paragraphs[:preview_count]):
            with st.expander(f"Paragraph {i+1} ({para.get('style', 'Normal')})", expanded=False):
                st.write(para.get('text', ''))
                
                # Paragraph details
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Words", len(para.get('text', '').split()))
                with col2:
                    st.metric("Characters", len(para.get('text', '')))
        
        if len(paragraphs) > preview_count:
            st.info(f"Showing {preview_count} of {len(paragraphs)} paragraphs")
        
        # Full document text
        full_text = '\n\n'.join([para.get('text', '') for para in paragraphs])
        display_text_content(full_text, "Full Document Text")
    
    # Display tables
    if tables:
        st.markdown("#### 📋 Tables")
        for i, table in enumerate(tables):
            st.markdown(f'<div class="table-container">', unsafe_allow_html=True)
            st.markdown(f"**Table {i+1}**")
            
            try:
                table_data = table.get('data', [])
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
    
    slides = data.get('slides', [])
    if not slides:
        st.warning("No slides found in this presentation.")
        return
    
    # Presentation statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Slides", len(slides))
    with col2:
        total_tables = sum(len(slide.get('tables', [])) for slide in slides)
        st.metric("Total Tables", total_tables)
    with col3:
        st.metric("Total Words", f"{data.get('total_words', 0):,}")
    
    # Slide navigation
    slide_numbers = list(range(1, len(slides) + 1))
    selected_slide = st.selectbox("🎯 Select a slide to view:", slide_numbers)
    
    if selected_slide <= len(slides):
        slide_data = slides[selected_slide - 1]
        
        st.markdown(f'<div class="page-header">📊 Slide {selected_slide}</div>', unsafe_allow_html=True)
        
        # Slide content
        text_content = slide_data.get('text_content', [])
        if text_content:
            st.markdown("**📝 Text Content:**")
            slide_text = '\n'.join([f"• {content}" for content in text_content])
            display_text_content(slide_text, f"Slide {selected_slide} Text")
        
        # Slide tables
        tables = slide_data.get('tables', [])
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
    
    sheets = data.get('sheets', [])
    if not sheets:
        st.warning("No sheets found in this Excel file.")
        return
    
    # Excel statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Sheets", len(sheets))
    with col2:
        total_rows = sum(sheet.get('shape', [0, 0])[0] for sheet in sheets)
        st.metric("Total Rows", f"{total_rows:,}")
    with col3:
        total_cols = max(sheet.get('shape', [0, 0])[1] for sheet in sheets) if sheets else 0
        st.metric("Max Columns", total_cols)
    
    # Sheet navigation
    sheet_names = [sheet['sheet_name'] for sheet in sheets]
    selected_sheet = st.selectbox("📋 Select a sheet to view:", sheet_names)
    
    # Find selected sheet data
    sheet_data = next(
        (sheet for sheet in sheets if sheet['sheet_name'] == selected_sheet), 
        None
    )
    
    if sheet_data:
        st.markdown(f'<div class="page-header">📋 Sheet: {selected_sheet}</div>', unsafe_allow_html=True)
        
        # Sheet statistics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Rows", f"{sheet_data['shape'][0]:,}")
        with col2:
            st.metric("Columns", sheet_data['shape'][1])
        with col3:
            non_empty_cells = sum(1 for row in sheet_data.get('data', []) for cell in row if cell)
            st.metric("Non-empty Cells", f"{non_empty_cells:,}")
        
        # Display data with pagination
        data_records = sheet_data.get('data', [])
        if data_records:
            df = pd.DataFrame(data_records)
            
            # Data preview options
            col1, col2 = st.columns(2)
            with col1:
                rows_to_show = st.slider("Rows to display:", 10, min(500, len(df)), 50)
            with col2:
                show_all_cols = st.checkbox("Show all columns", value=False)
            
            # Display dataframe
            if show_all_cols:
                st.dataframe(df.head(rows_to_show), use_container_width=True)
            else:
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
        if sheet_data.get('columns'):
            st.markdown("**📝 Column Information:**")
            cols_df = pd.DataFrame({
                'Column': sheet_data['columns'],
                'Index': range(len(sheet_data['columns']))
            })
            st.dataframe(cols_df, use_container_width=True)

def display_file_results(data, file_path):
    """Display enhanced results for a single file"""
    # File header with enhanced styling
    st.markdown(f"<div class='section-header'>📄 {data['filename']}</div>", unsafe_allow_html=True)
    
    # Display file summary
    display_file_summary(data)
    
    # Content tabs for better organization
    file_type = data.get('file_type', 'Unknown')
    
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
        if data.get('raw_text'):
            display_text_content(data['raw_text'], "Raw Extracted Text")
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
                    label=f"⬇️ Download {data['filename']} JSON",
                    data=json_data,
                    file_name=f"{os.path.splitext(data['filename'])[0]}_extraction.json",
                    mime="application/json",
                    key=f"download_{file_path}",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create JSON download: {str(e)}")
        
        with col2:
            # Download extracted text
            if data.get('raw_text'):
                st.download_button(
                    label="📝 Download Extracted Text",
                    data=data['raw_text'],
                    file_name=f"{os.path.splitext(data['filename'])[0]}_text.txt",
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
            st.markdown(f"• **{file.name}** ({file.size / 1024 / 1024:.2f} MB)")
        
        # Process files button with enhanced styling
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Process Files", type="primary", use_container_width=True):
                # Clean up previous temp files
                cleanup_temp_files()
                
                # Create temporary directory
                st.session_state.temp_dir = tempfile.mkdtemp()
                
                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    # Initialize extractor
                    status_text.text("Initializing document extractor...")
                    extractor = ComprehensiveDocumentExtractor()
                    progress_bar.progress(10)
                    
                    # Save uploaded files to temp directory
                    status_text.text("Saving uploaded files...")
                    saved_file_paths = []
                    for i, uploaded_file in enumerate(uploaded_files):
                        file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
                        with open(file_path, 'wb') as f:
                            f.write(uploaded_file.read())
                        saved_file_paths.append(file_path)
                        progress_bar.progress(10 + (i + 1) * 20 // len(uploaded_files))
                    
                    # Process files
                    status_text.text("Processing files... This may take a while...")
                    progress_bar.progress(40)
                    
                    results = extractor.process_files(saved_file_paths)
                    progress_bar.progress(90)
                    
                    if results:
                        st.session_state.processed_results = results
                        st.session_state.processing_complete = True
                        progress_bar.progress(100)
                        status_text.text("Processing complete!")
                        
                        st.markdown("""
                        <div class='success-message'>
                            <h3>🎉 Processing Complete!</h3>
                            <p>Successfully processed all files. View the results below.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div class='error-message'>
                            <h3>❌ Processing Failed</h3>
                            <p>No files could be processed. Please check your files and try again.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div class='error-message'>
                            <h3>❌ Processing Failed</h3>
                            <p>No files could be processed. Please check your files and try again.</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                except Exception as e:
                    st.error(f"Error processing files: {str(e)}")
                    st.exception(e)
                finally:
                    progress_bar.empty()
                    status_text.empty()
    
    # Display results if processing is complete
    if st.session_state.processing_complete and st.session_state.processed_results:
        st.markdown("---")
        st.markdown("## 📊 Processing Results")
        
        # Summary statistics
        total_files = len(st.session_state.processed_results)
        successful_files = sum(1 for result in st.session_state.processed_results if result.get('status') == 'success')
        failed_files = total_files - successful_files
        
        # Display summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📁 Total Files", total_files)
        with col2:
            st.metric("✅ Successful", successful_files)
        with col3:
            st.metric("❌ Failed", failed_files)
        with col4:
            success_rate = (successful_files / total_files * 100) if total_files > 0 else 0
            st.metric("📊 Success Rate", f"{success_rate:.1f}%")
        
        # Global download section
        st.markdown("### 🗂️ Download All Results")
        st.markdown('<div class="download-section">', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Download all results as JSON
            try:
                all_results_json = json.dumps(st.session_state.processed_results, indent=2, ensure_ascii=False, default=str)
                st.download_button(
                    label="📥 Download All Results (JSON)",
                    data=all_results_json,
                    file_name=f"document_extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    key="download_all_json",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create JSON download: {str(e)}")
        
        with col2:
            # Download all text content
            try:
                all_text = ""
                for result in st.session_state.processed_results:
                    if result.get('status') == 'success' and result.get('raw_text'):
                        all_text += f"\n\n{'='*50}\n"
                        all_text += f"FILE: {result['filename']}\n"
                        all_text += f"{'='*50}\n\n"
                        all_text += result['raw_text']
                
                if all_text:
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
                st.error(f"Could not create text download: {str(e)}")
        
        with col3:
            # Download processing summary
            try:
                summary_data = {
                    'processing_timestamp': datetime.now().isoformat(),
                    'total_files': total_files,
                    'successful_files': successful_files,
                    'failed_files': failed_files,
                    'success_rate': f"{success_rate:.1f}%",
                    'file_summary': []
                }
                
                for result in st.session_state.processed_results:
                    summary_data['file_summary'].append({
                        'filename': result.get('filename', 'Unknown'),
                        'status': result.get('status', 'Unknown'),
                        'file_type': result.get('file_type', 'Unknown'),
                        'file_size_mb': result.get('file_size_mb', 0),
                        'processing_time': result.get('processing_time', 'N/A')
                    })
                
                summary_json = json.dumps(summary_data, indent=2, ensure_ascii=False, default=str)
                st.download_button(
                    label="📊 Download Summary Report",
                    data=summary_json,
                    file_name=f"processing_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    key="download_summary",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create summary download: {str(e)}")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Display individual file results
        st.markdown("### 📋 Individual File Results")
        
        # Filter options
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox(
                "Filter by status:",
                ["All", "Successful", "Failed"],
                key="status_filter"
            )
        with col2:
            file_type_filter = st.selectbox(
                "Filter by file type:",
                ["All"] + list(set(result.get('file_type', 'Unknown') for result in st.session_state.processed_results)),
                key="file_type_filter"
            )
        
        # Apply filters
        filtered_results = st.session_state.processed_results
        
        if status_filter == "Successful":
            filtered_results = [r for r in filtered_results if r.get('status') == 'success']
        elif status_filter == "Failed":
            filtered_results = [r for r in filtered_results if r.get('status') != 'success']
        
        if file_type_filter != "All":
            filtered_results = [r for r in filtered_results if r.get('file_type') == file_type_filter]
        
        # Display filtered results
        if filtered_results:
            for i, result in enumerate(filtered_results):
                if result.get('status') == 'success':
                    display_file_results(result, f"file_{i}")
                else:
                    # Display error information
                    st.markdown(f"<div class='section-header'>❌ {result.get('filename', 'Unknown File')}</div>", unsafe_allow_html=True)
                    st.error(f"Processing failed: {result.get('error', 'Unknown error')}")
                
                # Add separator between files
                if i < len(filtered_results) - 1:
                    st.markdown("---")
        else:
            st.info("No files match the selected filters.")
        
        # Cleanup button
        st.markdown("### 🧹 Cleanup")
        if st.button("🗑️ Clear Results and Clean Up", type="secondary"):
            cleanup_temp_files()
            st.session_state.processed_results = None
            st.session_state.processing_complete = False
            st.success("✅ Results cleared and temporary files cleaned up!")
            st.rerun()
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; margin-top: 2em;'>
        <p>📄 SmartDoc Extractor | Built with Streamlit</p>
        <p style='font-size: 0.9em;'>
            Supports PDF, DOCX, PPTX, XLSX, and XLS files
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    # Clean up temp files on app start
    cleanup_temp_files()
    main()