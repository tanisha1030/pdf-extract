import streamlit as st
import os
import json
import time
import shutil
import zipfile
import tempfile
import pandas as pd
from datetime import datetime
from pathlib import Path
import io
import base64

# Import your existing extractor class
from pdfextract import ComprehensiveDocumentExtractor

# Page configuration
st.set_page_config(
    page_title="Document Extractor Pro",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    
    .stats-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin: 0.5rem 0;
    }
    
    .page-content {
        background: white;
        padding: 2rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        margin: 1rem 0;
        min-height: 400px;
    }
    
    .page-navigation {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        text-align: center;
    }
    
    .content-section {
        margin: 1.5rem 0;
        padding: 1rem;
        border-left: 3px solid #667eea;
        background: #f8f9fa;
        border-radius: 0 8px 8px 0;
    }
    
    .text-content {
        font-family: 'Georgia', serif;
        line-height: 1.6;
        color: #2c3e50;
        text-align: justify;
        padding: 1rem;
        background: white;
        border-radius: 8px;
        border: 1px solid #e9ecef;
    }
    
    .image-gallery {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 1rem 0;
    }
    
    .image-item {
        background: white;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        text-align: center;
    }
    
    .table-container {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        overflow-x: auto;
    }
    
    .success-message {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    
    .error-message {
        background: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #dc3545;
        margin: 1rem 0;
    }
    
    .feature-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 1rem 0;
    }
    
    .slide-content {
        background: white;
        padding: 2rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        margin: 1rem 0;
        min-height: 300px;
    }
    
    .sheet-tab {
        background: #e9ecef;
        color: #495057;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        margin: 0.25rem;
        display: inline-block;
        font-size: 0.9rem;
    }
    
    .sheet-tab.active {
        background: #667eea;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None
if 'stats' not in st.session_state:
    st.session_state.stats = None
if 'selected_file' not in st.session_state:
    st.session_state.selected_file = None
if 'selected_page' not in st.session_state:
    st.session_state.selected_page = 1

def main():
    # Main header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Document Extractor Pro</h1>
        <p>Extract text, images, tables, and metadata from PDF, DOCX, PPTX, and Excel files</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar with features and settings
    with st.sidebar:
        st.markdown("## 🔧 Features")
        st.markdown("""
        - **PDF Processing**: Text, images, tables, metadata
        - **DOCX Support**: Paragraphs, tables, styles
        - **PPTX Support**: Slides, text, tables
        - **Excel Support**: Multiple sheets, data analysis
        - **Advanced Table Detection**: Multiple extraction methods
        - **Page-by-Page Content View**: Navigate through document content
        - **Comprehensive Reports**: Summary and detailed analysis
        """)
        
        st.markdown("## 📊 Supported Formats")
        st.markdown("""
        - **PDF** (.pdf)
        - **Word** (.docx)
        - **PowerPoint** (.pptx)
        - **Excel** (.xlsx, .xls)
        """)
        
        st.markdown("## ⚙️ Settings")
        max_file_size = st.slider("Max file size (MB)", 1, 100, 50)
        show_detailed_logs = st.checkbox("Show detailed processing logs", value=False)
        auto_download = st.checkbox("Auto-download results", value=True)
        show_images = st.checkbox("Display extracted images", value=True)
        show_tables = st.checkbox("Display extracted tables", value=True)
    
    # Check if we have results to display
    if st.session_state.processing_complete and st.session_state.results:
        # Show content viewer
        show_content_viewer(st.session_state.results, show_images, show_tables)
    else:
        # Show upload interface
        show_upload_interface(max_file_size, show_detailed_logs, auto_download)

def show_upload_interface(max_file_size, show_detailed_logs, auto_download):
    """Show the file upload interface"""
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📤 Upload Documents")
        
        # File upload
        uploaded_files = st.file_uploader(
            "Choose files to process",
            type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
            accept_multiple_files=True,
            help="Upload multiple files for batch processing"
        )
        
        if uploaded_files:
            st.markdown("### 📋 Selected Files")
            file_info = []
            total_size = 0
            
            for file in uploaded_files:
                file_size_mb = len(file.getvalue()) / (1024 * 1024)
                total_size += file_size_mb
                file_info.append({
                    'Filename': file.name,
                    'Type': file.type,
                    'Size (MB)': f"{file_size_mb:.2f}"
                })
            
            # Display file information
            df_files = pd.DataFrame(file_info)
            st.dataframe(df_files, use_container_width=True)
            
            # Show total size
            st.info(f"Total size: {total_size:.2f} MB")
            
            # Check file size limits
            if total_size > max_file_size:
                st.error(f"⚠️ Total file size ({total_size:.2f} MB) exceeds limit ({max_file_size} MB)")
                return
            
            # Process button
            if st.button("🚀 Process Documents", type="primary", use_container_width=True):
                process_documents(uploaded_files, show_detailed_logs, auto_download)
    
    with col2:
        st.markdown("## 📈 Quick Stats")
        st.markdown("""
        <div class="feature-card">
            <h4>🎯 How to Use</h4>
            <ol>
                <li>Upload your documents using the file uploader</li>
                <li>Review the selected files</li>
                <li>Click "Process Documents" to start extraction</li>
                <li>Navigate through document content page by page</li>
                <li>Download the results when needed</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

def show_content_viewer(results, show_images, show_tables):
    """Show the content viewer with page navigation"""
    st.markdown("## 📖 Document Content Viewer")
    
    # File selector
    file_names = [data['filename'] for data in results.values()]
    
    # Create columns for file selection and stats
    col1, col2 = st.columns([2, 1])
    
    with col1:
        selected_file = st.selectbox(
            "📁 Select a file to view:",
            file_names,
            key="file_selector"
        )
        
        if selected_file != st.session_state.selected_file:
            st.session_state.selected_file = selected_file
            st.session_state.selected_page = 1
    
    with col2:
        # Quick stats for selected file
        if selected_file:
            file_data = get_file_data(results, selected_file)
            if file_data:
                st.markdown("### 📊 File Stats")
                st.metric("Type", file_data['file_type'])
                st.metric("Words", f"{file_data.get('total_words', 0):,}")
                st.metric("Size", f"{file_data.get('file_size_mb', 0):.2f} MB")
    
    if selected_file:
        file_data = get_file_data(results, selected_file)
        if file_data:
            # Show content based on file type
            if file_data['file_type'] == 'PDF':
                show_pdf_content(file_data, show_images, show_tables)
            elif file_data['file_type'] == 'DOCX':
                show_docx_content(file_data, show_images, show_tables)
            elif file_data['file_type'] == 'PPTX':
                show_pptx_content(file_data, show_images, show_tables)
            elif file_data['file_type'] == 'Excel':
                show_excel_content(file_data, show_tables)
    
    # Add download section
    st.markdown("---")
    create_download_section(results, st.session_state.stats)

def get_file_data(results, filename):
    """Get file data by filename"""
    for data in results.values():
        if data['filename'] == filename:
            return data
    return None

def show_pdf_content(file_data, show_images, show_tables):
    """Show PDF content with page navigation"""
    st.markdown("### 📄 PDF Content")
    
    pages = file_data.get('pages', [])
    if not pages:
        st.warning("No pages found in this PDF file.")
        return
    
    total_pages = len(pages)
    
    # Page navigation
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with col2:
        st.markdown(f"""
        <div class="page-navigation">
            <h4>Page {st.session_state.selected_page} of {total_pages}</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Page selector
        selected_page = st.selectbox(
            "Go to page:",
            range(1, total_pages + 1),
            index=st.session_state.selected_page - 1,
            key="page_selector"
        )
        
        if selected_page != st.session_state.selected_page:
            st.session_state.selected_page = selected_page
            st.rerun()
    
    with col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_pages):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Show current page content
    current_page = pages[st.session_state.selected_page - 1]
    
    st.markdown(f"""
    <div class="page-content">
        <h3>📄 Page {current_page.get('page_number', st.session_state.selected_page)} Content</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Show text content
    page_text = current_page.get('text', '').strip()
    if page_text:
        st.markdown("#### 📝 Text Content")
        st.markdown(f"""
        <div class="text-content">
            {page_text.replace('\n', '<br>')}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("No text content found on this page.")
    
    # Show images if enabled
    if show_images and current_page.get('images'):
        st.markdown("#### 🖼️ Images")
        images = current_page['images']
        
        # Display images in a grid
        cols = st.columns(min(3, len(images)))
        for i, image in enumerate(images):
            with cols[i % len(cols)]:
                st.markdown(f"""
                <div class="image-item">
                    <strong>Image {i+1}</strong><br>
                    Size: {image.get('width', 0)} x {image.get('height', 0)}<br>
                    File: {image.get('size_bytes', 0)} bytes
                </div>
                """, unsafe_allow_html=True)
                
                # Try to display image if file exists
                if image.get('filename') and os.path.exists(image['filename']):
                    try:
                        st.image(image['filename'], use_column_width=True)
                    except Exception as e:
                        st.error(f"Error displaying image: {str(e)}")
                else:
                    st.warning("Image file not found")
    
    # Show tables if enabled
    if show_tables and current_page.get('tables'):
        st.markdown("#### 📊 Tables")
        for i, table in enumerate(current_page['tables']):
            st.markdown(f"**Table {i+1}**")
            try:
                if table.get('data'):
                    df = pd.DataFrame(table['data'])
                    st.dataframe(df, use_container_width=True)
                else:
                    st.info("No table data available")
            except Exception as e:
                st.error(f"Error displaying table: {str(e)}")
    
    # Show page metadata
    with st.expander("📋 Page Metadata"):
        metadata = {
            'Page Number': current_page.get('page_number', st.session_state.selected_page),
            'Word Count': len(page_text.split()) if page_text else 0,
            'Character Count': len(page_text) if page_text else 0,
            'Images': len(current_page.get('images', [])),
            'Tables': len(current_page.get('tables', []))
        }
        
        for key, value in metadata.items():
            st.write(f"**{key}:** {value}")

def show_docx_content(file_data, show_images, show_tables):
    """Show DOCX content with paragraph navigation"""
    st.markdown("### 📄 Word Document Content")
    
    paragraphs = file_data.get('paragraphs', [])
    if not paragraphs:
        st.warning("No paragraphs found in this document.")
        return
    
    # Group paragraphs into pages (simulate pages)
    paragraphs_per_page = 10
    total_pages = (len(paragraphs) + paragraphs_per_page - 1) // paragraphs_per_page
    
    # Page navigation
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with col2:
        st.markdown(f"""
        <div class="page-navigation">
            <h4>Section {st.session_state.selected_page} of {total_pages}</h4>
        </div>
        """, unsafe_allow_html=True)
        
        selected_page = st.selectbox(
            "Go to section:",
            range(1, total_pages + 1),
            index=st.session_state.selected_page - 1,
            key="docx_page_selector"
        )
        
        if selected_page != st.session_state.selected_page:
            st.session_state.selected_page = selected_page
            st.rerun()
    
    with col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_pages):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Show current section content
    start_idx = (st.session_state.selected_page - 1) * paragraphs_per_page
    end_idx = min(start_idx + paragraphs_per_page, len(paragraphs))
    current_paragraphs = paragraphs[start_idx:end_idx]
    
    st.markdown(f"""
    <div class="page-content">
        <h3>📄 Section {st.session_state.selected_page} Content</h3>
        <p>Paragraphs {start_idx + 1} - {end_idx}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show paragraphs
    for i, paragraph in enumerate(current_paragraphs):
        para_text = paragraph.get('text', '').strip()
        if para_text:
            st.markdown(f"**Paragraph {start_idx + i + 1}:**")
            st.markdown(f"""
            <div class="text-content">
                {para_text.replace('\n', '<br>')}
            </div>
            """, unsafe_allow_html=True)
    
    # Show tables if enabled
    if show_tables and file_data.get('tables'):
        st.markdown("#### 📊 Document Tables")
        for i, table in enumerate(file_data['tables']):
            st.markdown(f"**Table {i+1}**")
            try:
                if table.get('data'):
                    df = pd.DataFrame(table['data'])
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Error displaying table: {str(e)}")

def show_pptx_content(file_data, show_images, show_tables):
    """Show PowerPoint content with slide navigation"""
    st.markdown("### 🎞️ PowerPoint Content")
    
    slides = file_data.get('slides', [])
    if not slides:
        st.warning("No slides found in this presentation.")
        return
    
    total_slides = len(slides)
    
    # Slide navigation
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with col2:
        st.markdown(f"""
        <div class="page-navigation">
            <h4>Slide {st.session_state.selected_page} of {total_slides}</h4>
        </div>
        """, unsafe_allow_html=True)
        
        selected_slide = st.selectbox(
            "Go to slide:",
            range(1, total_slides + 1),
            index=st.session_state.selected_page - 1,
            key="slide_selector"
        )
        
        if selected_slide != st.session_state.selected_page:
            st.session_state.selected_page = selected_slide
            st.rerun()
    
    with col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_slides):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Show current slide content
    current_slide = slides[st.session_state.selected_page - 1]
    
    st.markdown(f"""
    <div class="slide-content">
        <h3>🎞️ Slide {current_slide.get('slide_number', st.session_state.selected_page)}</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Show slide title
    slide_title = current_slide.get('title', '').strip()
    if slide_title:
        st.markdown(f"### {slide_title}")
    
    # Show slide content
    slide_content = current_slide.get('content', '').strip()
    if slide_content:
        st.markdown("#### 📝 Content")
        st.markdown(f"""
        <div class="text-content">
            {slide_content.replace('\n', '<br>')}
        </div>
        """, unsafe_allow_html=True)
    
    # Show notes if available
    notes = current_slide.get('notes', '').strip()
    if notes:
        st.markdown("#### 📋 Speaker Notes")
        st.markdown(f"""
        <div class="content-section">
            {notes.replace('\n', '<br>')}
        </div>
        """, unsafe_allow_html=True)
    
    # Show tables if enabled
    if show_tables and current_slide.get('tables'):
        st.markdown("#### 📊 Tables")
        for i, table in enumerate(current_slide['tables']):
            st.markdown(f"**Table {i+1}**")
            try:
                if table.get('data'):
                    df = pd.DataFrame(table['data'])
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Error displaying table: {str(e)}")

def show_excel_content(file_data, show_tables):
    """Show Excel content with sheet navigation"""
    st.markdown("### 📊 Excel Content")
    
    sheets = file_data.get('sheets', [])
    if not sheets:
        st.warning("No sheets found in this Excel file.")
        return
    
    # Sheet selector
    sheet_names = [sheet['name'] for sheet in sheets]
    selected_sheet_name = st.selectbox(
        "📋 Select a sheet:",
        sheet_names,
        key="sheet_selector"
    )
    
    # Find selected sheet
    selected_sheet = None
    for sheet in sheets:
        if sheet['name'] == selected_sheet_name:
            selected_sheet = sheet
            break
    
    if selected_sheet:
        st.markdown(f"""
        <div class="page-content">
            <h3>📊 Sheet: {selected_sheet['name']}</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Show sheet info
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Rows", selected_sheet.get('rows', 0))
        with col2:
            st.metric("Columns", selected_sheet.get('columns', 0))
        with col3:
            st.metric("Data Points", selected_sheet.get('rows', 0) * selected_sheet.get('columns', 0))
        
        # Show sheet data
        if show_tables and selected_sheet.get('data'):
            st.markdown("#### 📈 Data")
            try:
                df = pd.DataFrame(selected_sheet['data'])
                
                # Add pagination for large datasets
                if len(df) > 100:
                    st.info(f"Showing first 100 rows of {len(df)} total rows")
                    st.dataframe(df.head(100), use_container_width=True)
                    
                    # Show data summary
                    st.markdown("#### 📊 Data Summary")
                    st.write(df.describe())
                else:
                    st.dataframe(df, use_container_width=True)
                    
                    # Show data info
                    st.markdown("#### 📊 Data Info")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Data Types:**")
                        st.write(df.dtypes)
                    with col2:
                        st.write("**Statistics:**")
                        st.write(df.describe())
                
            except Exception as e:
                st.error(f"Error displaying sheet data: {str(e)}")
        else:
            st.info("No data available for this sheet.")

def process_documents(uploaded_files, show_detailed_logs, auto_download):
    """Process uploaded documents (same as original)"""
    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Save uploaded files
            file_paths = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, 'wb') as f:
                    f.write(uploaded_file.getvalue())
                file_paths.append(file_path)
            
            # Initialize extractor
            extractor = ComprehensiveDocumentExtractor()
            
            # Setup progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            if show_detailed_logs:
                log_container = st.container()
                log_text = st.empty()
                logs = []
            
            # Process files
            start_time = time.time()
            status_text.text("🔄 Initializing document processing...")
            progress_bar.progress(10)
            
            results = {}
            total_files = len(file_paths)
            
            for i, file_path in enumerate(file_paths):
                file_name = os.path.basename(file_path)
                status_text.text(f"📄 Processing {file_name} ({i+1}/{total_files})")
                
                if show_detailed_logs:
                    logs.append(f"Processing: {file_name}")
                    log_text.text_area("Processing Logs", "\n".join(logs), height=200)
                
                # Determine file type and process
                file_ext = os.path.splitext(file_path)[1].lower()
                
                try:
                    if file_ext == '.pdf':
                        result = extractor.extract_pdf_comprehensive(file_path)
                    elif file_ext == '.docx':
                        result = extractor.extract_docx(file_path)
                    elif file_ext == '.pptx':
                        result = extractor.extract_pptx(file_path)
                    elif file_ext in ['.xlsx', '.xls']:
                        result = extractor.extract_excel(file_path)
                    else:
                        continue
                    
                    if result:
                        results[file_path] = result
                        if show_detailed_logs:
                            logs.append(f"✅ Successfully processed: {file_name}")
                    else:
                        if show_detailed_logs:
                            logs.append(f"❌ Failed to process: {file_name}")
                        
                except Exception as e:
                    if show_detailed_logs:
                        logs.append(f"❌ Error processing {file_name}: {str(e)}")
                    continue
                
                # Update progress
                progress = 10 + (i + 1) / total_files * 80
                progress_bar.progress(int(progress))
                
                if show_detailed_logs:
                    log_text.text_area("Processing Logs", "\n".join(logs), height=200)
            
            # Generate statistics
            status_text.text("📊 Generating statistics and reports...")
            progress_bar.progress(95)
            
            processing_time = time.time() - start_time
            stats = generate_stats(results, processing_time)
            
            # Save results to session state
            st.session_state.results = results
            st.session_state.stats = stats
            st.session_state.processing_complete = True
            
            # Complete
            progress_bar.progress(100)
            status_text.text("✅ Processing complete!")
            
            # Success message
            st.markdown(f"""
            <div class="success-message">
                <h4>🎉 Processing Complete!</h4>
                <p>Successfully processed {stats['total_files']} files in {processing_time:.2f} seconds</p>
                <ul>
                    <li>Words extracted: {stats['total_words']:,}</li>
                    <li>Images found: {stats['total_images']:,}</li>
                    <li>Tables detected: {stats['total_tables']:,}</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
            # Auto-download if enabled
            if auto_download and results:
                create_download_files(results, stats)
            
            # Refresh the page to show results
            st.rerun()
            
        except Exception as e:
            st.error(f"Error during processing: {str(e)}")
            if show_detailed_logs:
                logs.append(f"❌ Fatal error: {str(e)}")
                log_text.text_area("Processing Logs", "\n".join(logs), height=200)
            return


def generate_stats(results, processing_time):
    """Generate comprehensive statistics from processing results"""
    stats = {
        'total_files': len(results),
        'processing_time': processing_time,
        'total_words': 0,
        'total_images': 0,
        'total_tables': 0,
        'total_pages': 0,
        'file_types': {},
        'file_details': []
    }
    
    for file_path, data in results.items():
        file_name = data.get('filename', os.path.basename(file_path))
        file_type = data.get('file_type', 'Unknown')
        
        # Count file types
        if file_type not in stats['file_types']:
            stats['file_types'][file_type] = 0
        stats['file_types'][file_type] += 1
        
        # Extract metrics based on file type
        file_words = 0
        file_images = 0
        file_tables = 0
        file_pages = 0
        
        if file_type == 'PDF':
            pages = data.get('pages', [])
            file_pages = len(pages)
            for page in pages:
                if page.get('text'):
                    file_words += len(page['text'].split())
                if page.get('images'):
                    file_images += len(page['images'])
                if page.get('tables'):
                    file_tables += len(page['tables'])
        
        elif file_type == 'DOCX':
            paragraphs = data.get('paragraphs', [])
            for paragraph in paragraphs:
                if paragraph.get('text'):
                    file_words += len(paragraph['text'].split())
            if data.get('tables'):
                file_tables += len(data['tables'])
        
        elif file_type == 'PPTX':
            slides = data.get('slides', [])
            file_pages = len(slides)
            for slide in slides:
                if slide.get('title'):
                    file_words += len(slide['title'].split())
                if slide.get('content'):
                    file_words += len(slide['content'].split())
                if slide.get('notes'):
                    file_words += len(slide['notes'].split())
                if slide.get('tables'):
                    file_tables += len(slide['tables'])
        
        elif file_type == 'Excel':
            sheets = data.get('sheets', [])
            file_pages = len(sheets)
            for sheet in sheets:
                if sheet.get('data'):
                    file_tables += 1
                    # Count cells as "words" for Excel
                    rows = sheet.get('rows', 0)
                    cols = sheet.get('columns', 0)
                    file_words += rows * cols
        
        # Update totals
        stats['total_words'] += file_words
        stats['total_images'] += file_images
        stats['total_tables'] += file_tables
        stats['total_pages'] += file_pages
        
        # Store file details
        stats['file_details'].append({
            'filename': file_name,
            'type': file_type,
            'words': file_words,
            'images': file_images,
            'tables': file_tables,
            'pages': file_pages,
            'size_mb': data.get('file_size_mb', 0)
        })
    
    return stats


def create_download_section(results, stats):
    """Create download section with various format options"""
    st.markdown("## 📥 Download Results")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # JSON download
        if st.button("📄 Download JSON", type="secondary", use_container_width=True):
            json_data = create_json_export(results, stats)
            st.download_button(
                label="📄 Download Complete Results (JSON)",
                data=json_data,
                file_name=f"extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
    
    with col2:
        # CSV download
        if st.button("📊 Download CSV", type="secondary", use_container_width=True):
            csv_data = create_csv_export(results, stats)
            st.download_button(
                label="📊 Download Summary (CSV)",
                data=csv_data,
                file_name=f"extraction_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with col3:
        # ZIP download
        if st.button("📦 Download ZIP", type="secondary", use_container_width=True):
            zip_data = create_zip_export(results, stats)
            if zip_data:
                st.download_button(
                    label="📦 Download Complete Package (ZIP)",
                    data=zip_data,
                    file_name=f"extraction_package_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )


def create_json_export(results, stats):
    """Create JSON export of all results"""
    export_data = {
        'extraction_metadata': {
            'timestamp': datetime.now().isoformat(),
            'stats': stats
        },
        'results': {}
    }
    
    # Clean results for JSON serialization
    for file_path, data in results.items():
        clean_data = {}
        for key, value in data.items():
            if isinstance(value, (str, int, float, bool, list, dict, type(None))):
                clean_data[key] = value
            else:
                clean_data[key] = str(value)
        export_data['results'][os.path.basename(file_path)] = clean_data
    
    return json.dumps(export_data, indent=2, ensure_ascii=False)


def create_csv_export(results, stats):
    """Create CSV export of summary data"""
    if not stats.get('file_details'):
        return ""
    
    df = pd.DataFrame(stats['file_details'])
    return df.to_csv(index=False)


def create_zip_export(results, stats):
    """Create ZIP export with all files and reports"""
    try:
        # Create in-memory zip file
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Add JSON results
            json_data = create_json_export(results, stats)
            zip_file.writestr("extraction_results.json", json_data)
            
            # Add CSV summary
            csv_data = create_csv_export(results, stats)
            if csv_data:
                zip_file.writestr("extraction_summary.csv", csv_data)
            
            # Add text reports for each file
            for file_path, data in results.items():
                filename = data.get('filename', os.path.basename(file_path))
                safe_filename = "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
                
                # Create text report
                text_report = create_text_report(data)
                zip_file.writestr(f"reports/{safe_filename}_report.txt", text_report)
        
        zip_buffer.seek(0)
        return zip_buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error creating ZIP file: {str(e)}")
        return None


def create_text_report(data):
    """Create a text report for a single file"""
    report = []
    report.append(f"DOCUMENT EXTRACTION REPORT")
    report.append(f"=" * 50)
    report.append(f"File: {data.get('filename', 'Unknown')}")
    report.append(f"Type: {data.get('file_type', 'Unknown')}")
    report.append(f"Size: {data.get('file_size_mb', 0):.2f} MB")
    report.append(f"Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")
    
    file_type = data.get('file_type', '')
    
    if file_type == 'PDF':
        pages = data.get('pages', [])
        report.append(f"PAGES: {len(pages)}")
        report.append("-" * 20)
        
        for i, page in enumerate(pages, 1):
            report.append(f"\nPAGE {i}:")
            report.append("-" * 10)
            page_text = page.get('text', '').strip()
            if page_text:
                report.append(page_text)
            else:
                report.append("(No text content)")
            
            if page.get('images'):
                report.append(f"\nImages found: {len(page['images'])}")
            if page.get('tables'):
                report.append(f"Tables found: {len(page['tables'])}")
    
    elif file_type == 'DOCX':
        paragraphs = data.get('paragraphs', [])
        report.append(f"PARAGRAPHS: {len(paragraphs)}")
        report.append("-" * 20)
        
        for i, paragraph in enumerate(paragraphs, 1):
            para_text = paragraph.get('text', '').strip()
            if para_text:
                report.append(f"\nPARAGRAPH {i}:")
                report.append(para_text)
    
    elif file_type == 'PPTX':
        slides = data.get('slides', [])
        report.append(f"SLIDES: {len(slides)}")
        report.append("-" * 20)
        
        for i, slide in enumerate(slides, 1):
            report.append(f"\nSLIDE {i}:")
            report.append("-" * 10)
            
            if slide.get('title'):
                report.append(f"Title: {slide['title']}")
            if slide.get('content'):
                report.append(f"Content: {slide['content']}")
            if slide.get('notes'):
                report.append(f"Notes: {slide['notes']}")
    
    elif file_type == 'Excel':
        sheets = data.get('sheets', [])
        report.append(f"SHEETS: {len(sheets)}")
        report.append("-" * 20)
        
        for sheet in sheets:
            report.append(f"\nSheet: {sheet.get('name', 'Unknown')}")
            report.append(f"Rows: {sheet.get('rows', 0)}")
            report.append(f"Columns: {sheet.get('columns', 0)}")
    
    return "\n".join(report)


def create_download_files(results, stats):
    """Create downloadable files automatically"""
    # This function is called when auto_download is enabled
    # In a real implementation, you might want to save files to a temporary location
    # For now, we'll just show a message
    st.info("📥 Auto-download enabled. Use the download buttons below to get your results.")


# Add reset functionality
def reset_app():
    """Reset the application state"""
    st.session_state.processing_complete = False
    st.session_state.results = None
    st.session_state.stats = None
    st.session_state.selected_file = None
    st.session_state.selected_page = 1


# Add this to your sidebar in the main function
def add_reset_button():
    """Add reset button to sidebar"""
    if st.session_state.processing_complete:
        st.markdown("---")
        if st.button("🔄 Process New Files", type="primary", use_container_width=True):
            reset_app()
            st.rerun()


# Update the main function to include the reset button
def main():
    # Main header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Document Extractor Pro</h1>
        <p>Extract text, images, tables, and metadata from PDF, DOCX, PPTX, and Excel files</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar with features and settings
    with st.sidebar:
        st.markdown("## 🔧 Features")
        st.markdown("""
        - **PDF Processing**: Text, images, tables, metadata
        - **DOCX Support**: Paragraphs, tables, styles
        - **PPTX Support**: Slides, text, tables
        - **Excel Support**: Multiple sheets, data analysis
        - **Advanced Table Detection**: Multiple extraction methods
        - **Page-by-Page Content View**: Navigate through document content
        - **Comprehensive Reports**: Summary and detailed analysis
        """)
        
        st.markdown("## 📊 Supported Formats")
        st.markdown("""
        - **PDF** (.pdf)
        - **Word** (.docx)
        - **PowerPoint** (.pptx)
        - **Excel** (.xlsx, .xls)
        """)
        
        st.markdown("## ⚙️ Settings")
        max_file_size = st.slider("Max file size (MB)", 1, 100, 50)
        show_detailed_logs = st.checkbox("Show detailed processing logs", value=False)
        auto_download = st.checkbox("Auto-download results", value=True)
        show_images = st.checkbox("Display extracted images", value=True)
        show_tables = st.checkbox("Display extracted tables", value=True)
        
        # Add reset button
        add_reset_button()
    
    # Check if we have results to display
    if st.session_state.processing_complete and st.session_state.results:
        # Show content viewer
        show_content_viewer(st.session_state.results, show_images, show_tables)
    else:
        # Show upload interface
        show_upload_interface(max_file_size, show_detailed_logs, auto_download)


if __name__ == "__main__":
    main()
