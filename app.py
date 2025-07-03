import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime
import zipfile
import tempfile
import shutil
from io import BytesIO
import base64
import time
from collections import Counter
from PIL import Image

# Import the extractor class
from pdfextract import ComprehensiveDocumentExtractor

# Configure Streamlit page
st.set_page_config(
    page_title="Document Extractor Pro",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #667eea;
    }
    
    .success-message {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #c3e6cb;
    }
    
    .error-message {
        background: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #f5c6cb;
    }
    
    .warning-message {
        background: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #ffeaa7;
    }
    
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    .feature-info {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin-bottom: 1rem;
    }
    
    .image-container {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem;
        margin-top: 1rem;
    }
    
    .image-card {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 0.5rem;
        background: white;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
        max-width: 350px;
    }
    
    .image-card img {
        max-width: 300px;
        max-height: 300px;
        border-radius: 4px;
        object-fit: contain;
    }
    
    .page-content {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e9ecef;
        margin-top: 1rem;
    }
    
    .page-selector {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    
    .text-preview {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        font-family: monospace;
        font-size: 0.9em;
        line-height: 1.4;
        max-height: 300px;
        overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'extraction_results' not in st.session_state:
    st.session_state.extraction_results = None
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = []
if 'selected_pages' not in st.session_state:
    st.session_state.selected_pages = {}

# Default processing options
DEFAULT_OPTIONS = {
    'extract_images': True,
    'extract_tables': True,
    'extract_metadata': True,
    'use_tabula': True,
    'use_camelot': True,
    'use_pdfplumber': True,
    'save_as_json': True,
    'save_as_csv': True,
    'create_summary': True
}

def clean_table_data(table_data):
    """Clean table data to handle duplicate/None column names and formatting issues"""
    try:
        if not table_data:
            return None
            
        # Handle different data formats
        if isinstance(table_data, list) and len(table_data) > 0:
            # Check if first row contains headers
            if isinstance(table_data[0], dict):
                # DataFrame format - already has column names
                df = pd.DataFrame(table_data)
            else:
                # List of lists format
                if len(table_data) > 1:
                    headers = table_data[0]
                    data_rows = table_data[1:]
                else:
                    headers = table_data[0] if table_data else []
                    data_rows = []
                
                # Clean headers - handle None values and duplicates
                cleaned_headers = []
                header_counts = {}
                
                for i, header in enumerate(headers):
                    if header is None or header == '' or str(header).strip() == '':
                        clean_header = f"Column_{i+1}"
                    else:
                        clean_header = str(header).strip()
                    
                    # Handle duplicates
                    if clean_header in header_counts:
                        header_counts[clean_header] += 1
                        clean_header = f"{clean_header}_{header_counts[clean_header]}"
                    else:
                        header_counts[clean_header] = 1
                    
                    cleaned_headers.append(clean_header)
                
                # Ensure we have enough columns for all data
                max_cols = max(len(row) for row in data_rows) if data_rows else len(cleaned_headers)
                
                # Pad headers if needed
                while len(cleaned_headers) < max_cols:
                    cleaned_headers.append(f"Column_{len(cleaned_headers)+1}")
                
                # Pad data rows if needed
                padded_data = []
                for row in data_rows:
                    padded_row = list(row) + [None] * (max_cols - len(row))
                    padded_data.append(padded_row[:max_cols])  # Trim if too long
                
                # Create DataFrame
                df = pd.DataFrame(padded_data, columns=cleaned_headers[:max_cols])
        else:
            return None
        
        # Clean DataFrame - handle any remaining issues
        if df is not None and not df.empty:
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Remove completely empty columns
            df = df.dropna(axis=1, how='all')
            
            # Convert all data to string for display consistency
            df = df.astype(str)
            
            # Replace 'nan' and 'None' with empty strings for cleaner display
            df = df.replace(['nan', 'None'], '')
            
            return df
        
        return None
        
    except Exception as e:
        print(f"Error cleaning table data: {str(e)}")
        return None

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Document Extractor Pro</h1>
        <p>Extract text, tables, images, and metadata from PDF, DOCX, PPTX, and Excel files</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Features info
    st.markdown("""
    <div class="feature-info">
        <h3>🚀 Automatic Processing Features</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1rem; margin-top: 1rem;">
            <div>
                <strong>📄 Supported Formats:</strong>
                <ul>
                    <li>PDF files</li>
                    <li>DOCX documents</li>
                    <li>PowerPoint (PPTX)</li>
                    <li>Excel (XLSX/XLS)</li>
                </ul>
            </div>
            <div>
                <strong>🔧 Extraction Capabilities:</strong>
                <ul>
                    <li>Text content with formatting</li>
                    <li>Tables (multiple methods)</li>
                    <li>Images and graphics</li>
                    <li>Metadata information</li>
                </ul>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Upload Documents")
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Choose files to process",
            accept_multiple_files=True,
            type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
            help="Upload one or more documents for processing"
        )
        
        if uploaded_files:
            st.session_state.uploaded_files = uploaded_files
            
            # Display uploaded files
            st.subheader("📋 Uploaded Files")
            for i, file in enumerate(uploaded_files, 1):
                file_size_mb = len(file.getvalue()) / (1024 * 1024)
                st.write(f"{i}. **{file.name}** ({file_size_mb:.2f} MB)")
            
            # Process button
            if st.button("🚀 Start Processing", type="primary", use_container_width=True):
                process_documents(uploaded_files, DEFAULT_OPTIONS)
    
    with col2:
        st.header("📊 Processing Status")
        
        if not uploaded_files:
            st.info("👆 Upload files to get started")
        else:
            # Show file statistics
            total_files = len(uploaded_files)
            total_size_mb = sum(len(f.getvalue()) / (1024 * 1024) for f in uploaded_files)
            
            st.metric("Files to Process", total_files)
            st.metric("Total Size", f"{total_size_mb:.2f} MB")
            
            # File type distribution
            file_types = Counter(f.name.split('.')[-1].upper() for f in uploaded_files)
            st.subheader("File Types")
            for file_type, count in file_types.items():
                st.write(f"📄 {file_type}: {count}")
    
    # Results section
    if st.session_state.processing_complete and st.session_state.extraction_results:
        display_results(st.session_state.extraction_results)

def process_documents(uploaded_files, options):
    """Process uploaded documents"""
    
    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Save uploaded files to temporary directory
            file_paths = []
            for i, uploaded_file in enumerate(uploaded_files):
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, 'wb') as f:
                    f.write(uploaded_file.getvalue())
                file_paths.append(file_path)
                
                progress = (i + 1) / (len(uploaded_files) + 1)
                progress_bar.progress(progress)
                status_text.text(f"Saving file {i + 1}/{len(uploaded_files)}: {uploaded_file.name}")
            
            # Initialize extractor
            status_text.text("Initializing document extractor...")
            extractor = ComprehensiveDocumentExtractor()
            
            # Process files
            status_text.text("Processing documents...")
            progress_bar.progress(0.1)
            
            results = {}
            for i, file_path in enumerate(file_paths):
                file_name = os.path.basename(file_path)
                status_text.text(f"Processing {file_name}...")
                
                try:
                    file_ext = os.path.splitext(file_path)[1].lower()
                    
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
                        results[file_name] = result
                        
                except Exception as e:
                    st.error(f"Error processing {file_name}: {str(e)}")
                    continue
                
                progress = 0.1 + (0.8 * (i + 1) / len(file_paths))
                progress_bar.progress(progress)
            
            # Save results
            if results:
                status_text.text("Saving results...")
                progress_bar.progress(0.9)
                
                # Store results in session state
                st.session_state.extraction_results = results
                st.session_state.processing_complete = True
                
                # Initialize page selectors for PDF files
                for file_name, data in results.items():
                    if data.get('file_type') == 'PDF' and file_name not in st.session_state.selected_pages:
                        st.session_state.selected_pages[file_name] = 1
                
                progress_bar.progress(1.0)
                status_text.text("✅ Processing completed!")
                
                # Auto-refresh to show results
                time.sleep(1)
                st.rerun()
            else:
                st.error("❌ No files were successfully processed.")
                
        except Exception as e:
            st.error(f"❌ Error during processing: {str(e)}")

def display_page_content(file_name, data, selected_page):
    """Display content for a specific page"""
    
    if data.get('file_type') != 'PDF' or not data.get('pages'):
        return
    
    if selected_page <= 0 or selected_page > len(data['pages']):
        st.error("Invalid page number selected")
        return
    
    page_data = data['pages'][selected_page - 1]
    
    st.markdown(f"""
    <div class="page-content">
        <h4>📄 Page {selected_page} Content</h4>
    </div>
    """, unsafe_allow_html=True)
    
    # Create columns for page content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Text content
        page_text = page_data.get('text', '')
        if page_text:
            st.markdown("**📝 Text Content:**")
            st.markdown(f"""
            <div class="text-preview">
                {page_text.replace('\n', '<br>')}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("No text content found on this page")
    
    with col2:
        # Page statistics
        st.markdown("**📊 Page Statistics:**")
        st.write(f"• **Characters:** {len(page_text):,}")
        st.write(f"• **Words:** {len(page_text.split()) if page_text else 0:,}")
        st.write(f"• **Images:** {len(page_data.get('images', []))}")
        
        # Page-specific tables
        page_tables = [table for table in data.get('extracted_tables', []) 
                      if table.get('page') == selected_page]
        st.write(f"• **Tables:** {len(page_tables)}")

def display_results(results):
    """Display extraction results"""
    
    st.header("🎉 Processing Complete!")
    
    # Overall statistics
    total_files = len(results)
    total_words = sum(data.get('total_words', 0) for data in results.values())
    total_chars = sum(data.get('total_characters', 0) for data in results.values())
    total_images = sum(data.get('total_images', 0) for data in results.values())
    total_tables = sum(len(data.get('extracted_tables', [])) for data in results.values())
    
    # Statistics cards
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("📁 Files Processed", total_files)
    with col2:
        st.metric("📝 Total Words", f"{total_words:,}")
    with col3:
        st.metric("🔤 Characters", f"{total_chars:,}")
    with col4:
        st.metric("🖼️ Images", total_images)
    with col5:
        st.metric("📊 Tables", total_tables)
    
    # Detailed results tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📄 File Details", "📋 Page View", "📊 Tables", "🖼️ Images", "📈 Summary"])
    
    with tab1:
        st.subheader("File Processing Details")
        
        for file_name, data in results.items():
            with st.expander(f"📄 {file_name}"):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write(f"**File Type:** {data['file_type']}")
                    st.write(f"**Size:** {data.get('file_size_mb', 0):.2f} MB")
                    
                    if data['file_type'] == 'PDF':
                        st.write(f"**Pages:** {data.get('page_count', 0)}")
                        st.write(f"**Words:** {data.get('total_words', 0):,}")
                        st.write(f"**Images:** {data.get('total_images', 0)}")
                        st.write(f"**Tables:** {len(data.get('extracted_tables', []))}")
                    
                    elif data['file_type'] == 'DOCX':
                        st.write(f"**Paragraphs:** {len(data.get('paragraphs', []))}")
                        st.write(f"**Words:** {data.get('total_words', 0):,}")
                        st.write(f"**Tables:** {len(data.get('tables', []))}")
                    
                    elif data['file_type'] == 'PPTX':
                        st.write(f"**Slides:** {data.get('total_slides', 0)}")
                        st.write(f"**Words:** {data.get('total_words', 0):,}")
                    
                    elif data['file_type'] == 'Excel':
                        st.write(f"**Sheets:** {len(data.get('sheets', []))}")
                
                with col2:
                    # Show metadata
                    if data.get('metadata'):
                        st.write("**Metadata:**")
                        for key, value in data['metadata'].items():
                            if isinstance(value, str) and len(value) > 50:
                                st.write(f"  - {key}: {value[:50]}...")
                            else:
                                st.write(f"  - {key}: {value}")
    
    with tab2:
        st.subheader("📋 Page-by-Page Content View")
        
        # Filter PDF files
        pdf_files = {name: data for name, data in results.items() if data.get('file_type') == 'PDF'}
        
        if pdf_files:
            # File selector
            selected_file = st.selectbox(
                "Select a PDF file to view:",
                options=list(pdf_files.keys()),
                key="page_view_file_selector"
            )
            
            if selected_file:
                file_data = pdf_files[selected_file]
                page_count = file_data.get('page_count', 0)
                
                if page_count > 0:
                    # Enhanced page selector
                    st.markdown('<div class="page-selector">', unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([1, 2, 1])
                    
                    with col1:
                        if st.button("⬅️ Previous", disabled=st.session_state.selected_pages.get(selected_file, 1) <= 1):
                            st.session_state.selected_pages[selected_file] = max(1, st.session_state.selected_pages.get(selected_file, 1) - 1)
                            st.rerun()
                    
                    with col2:
                        current_page = st.session_state.selected_pages.get(selected_file, 1)
                        selected_page = st.selectbox(
                            f"Page (1-{page_count}):",
                            options=range(1, page_count + 1),
                            index=current_page - 1,
                            key=f"page_select_{selected_file}"
                        )
                        st.session_state.selected_pages[selected_file] = selected_page
                    
                    with col3:
                        if st.button("➡️ Next", disabled=st.session_state.selected_pages.get(selected_file, 1) >= page_count):
                            st.session_state.selected_pages[selected_file] = min(page_count, st.session_state.selected_pages.get(selected_file, 1) + 1)
                            st.rerun()
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Display page content
                    current_page = st.session_state.selected_pages.get(selected_file, 1)
                    display_page_content(selected_file, file_data, current_page)
                else:
                    st.warning("No pages found in this PDF file")
        else:
            st.info("No PDF files found. Page view is only available for PDF documents.")
    
    with tab3:
        st.subheader("📊 Extracted Tables")
        
        if total_tables > 0:
            # Table extraction methods summary
            method_counts = Counter()
            for data in results.values():
                for table in data.get('extracted_tables', []):
                    method_counts[table.get('method', 'unknown')] += 1
            
            st.write("**Extraction Methods Used:**")
            for method, count in method_counts.items():
                st.write(f"  - {method}: {count} tables")
            
            # Show tables by file
            for file_name, data in results.items():
                tables = data.get('extracted_tables', [])
                if tables:
                    st.write(f"**📄 {file_name}**")
                    
                    for i, table in enumerate(tables, 1):
                        with st.expander(f"Table {i} ({table.get('method', 'unknown')}) - Page {table.get('page', 'N/A')}"):
                            
                            # Show table metadata
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"**Method:** {table.get('method', 'unknown')}")
                                st.write(f"**Page:** {table.get('page', 'N/A')}")
                                st.write(f"**Confidence:** {table.get('confidence', 'unknown')}")
                            with col2:
                                if 'shape' in table:
                                    st.write(f"**Shape:** {table['shape'][0]} rows × {table['shape'][1]} columns")
                                if 'accuracy' in table:
                                    st.write(f"**Accuracy:** {table['accuracy']:.1f}%")
                            
                            # Show table data with improved error handling
                            try:
                                if table.get('data'):
                                    # Clean the table data
                                    cleaned_df = clean_table_data(table['data'])
                                    
                                    if cleaned_df is not None and not cleaned_df.empty:
                                        st.dataframe(cleaned_df, use_container_width=True)
                                        
                                        # Show additional info
                                        st.info(f"Table size: {len(cleaned_df)} rows × {len(cleaned_df.columns)} columns")
                                        
                                        # Option to download individual table
                                        csv_data = cleaned_df.to_csv(index=False)
                                        st.download_button(
                                            label=f"Download Table {i} as CSV",
                                            data=csv_data,
                                            file_name=f"{os.path.splitext(file_name)[0]}_table_{i}.csv",
                                            mime="text/csv",
                                            key=f"download_table_{file_name}_{i}"
                                        )
                                    else:
                                        st.warning("Table data could not be displayed (empty or invalid format)")
                                        
                                        # Show raw data for debugging
                                        with st.expander("Raw Table Data (Debug)"):
                                            st.json(table['data'])
                                else:
                                    st.warning("No table data available")
                                    
                            except Exception as e:
                                st.error(f"Error displaying table: {str(e)}")
                                
                                # Show raw data for debugging
                                with st.expander("Raw Table Data (Debug)"):
                                    try:
                                        st.json(table.get('data', 'No data available'))
                                    except:
                                        st.text(str(table.get('data', 'No data available')))
                                
                                # Try alternative display methods
                                st.info("Attempting alternative display...")
                                try:
                                    table_data = table.get('data', [])
                                    if isinstance(table_data, list) and table_data:
                                        st.text("Raw table content:")
                                        for row_idx, row in enumerate(table_data[:10]):  # Show first 10 rows
                                            st.text(f"Row {row_idx + 1}: {row}")
                                        if len(table_data) > 10:
                                            st.text(f"... and {len(table_data) - 10} more rows")
                                except:
                                    st.error("Could not display table in any format")
        else:
            st.info("No tables were extracted from the uploaded documents.")
    
    with tab4:
        st.subheader("🖼️ Extracted Images")
        
        if total_images > 0:
            for file_name, data in results.items():
                if data.get('total_images', 0) > 0:
                    st.write(f"**📄 {file_name}**")
                    
                    # Show images by page (for PDFs)
                    if data['file_type'] == 'PDF':
                        for page in data.get('pages', []):
                            images = page.get('images', [])
                            if images:
                                st.write(f"**Page {page['page_number']}:** {len(images)} images")
                                
                                # Display images in columns
                                cols = st.columns(min(3, len(images)))
                                
                                for idx, img in enumerate(images):
                                    with cols[idx % 3]:
                                        try:
                                            # Display image if data is available
                                            if 'image_data' in img:
                                                st.image(
                                                    base64.b64decode(img['image_data']),
                                                    caption=f"{os.path.basename(img['filename'])}",
                                                    width=300
                                                )
                                            else:
                                                # Show placeholder if no image data
                                                st.info(f"📸 {os.path.basename(img['filename'])}")
                                            
                                            # Image metadata
                                            with st.expander(f"Image {idx + 1} Details"):
                                                st.write(f"**Filename:** {os.path.basename(img['filename'])}")
                                                st.write(f"**Size:** {img.get('width', 'N/A')}×{img.get('height', 'N/A')}")
                                                st.write(f"**Colorspace:** {img.get('colorspace', 'N/A')}")
                                                st.write(f"**File size:** {img.get('size_bytes', 'N/A')} bytes")
                                                if 'bbox' in img:
                                                    st.write(f"**Position:** {img['bbox']}")
                                        
                                        except Exception as e:
                                            st.error(f"Error displaying image: {str(e)}")
                    
                    # Show images for other file types
                    elif data['file_type'] in ['DOCX', 'PPTX']:
                        images = data.get('images', [])
                        if images:
                            st.write(f"**Images found:** {len(images)}")
                            
                            cols = st.columns(min(3, len(images)))
                            
                            for idx, img in enumerate(images):
                                with cols[idx % 3]:
                                    try:
                                        if 'image_data' in img:
                                            st.image(
                                                base64.b64decode(img['image_data']),
                                                caption=f"Image {idx + 1}",
                                                width=300
                                            )
                                        else:
                                            st.info(f"📸 Image {idx + 1}")
                                        
                                        with st.expander(f"Image {idx + 1} Details"):
                                            st.write(f"**Type:** {img.get('type', 'N/A')}")
                                            st.write(f"**Size:** {img.get('size_bytes', 'N/A')} bytes")
                                            if 'format' in img:
                                                st.write(f"**Format:** {img['format']}")
                                    
                                    st.error(f"Error displaying image: {str(e)}")
        else:
            st.info("No images were extracted from the uploaded documents.")
    
    with tab5:
        st.subheader("📈 Processing Summary")
        
        # Create summary statistics
        st.write("**📊 Overall Statistics:**")
        
        # Processing time estimation
        processing_time = "Processing completed successfully"
        
        # Create summary DataFrame
        summary_data = []
        for file_name, data in results.items():
            summary_data.append({
                'File Name': file_name,
                'File Type': data['file_type'],
                'Size (MB)': f"{data.get('file_size_mb', 0):.2f}",
                'Pages/Slides/Sheets': data.get('page_count', data.get('total_slides', len(data.get('sheets', [])))),
                'Words': f"{data.get('total_words', 0):,}",
                'Images': data.get('total_images', 0),
                'Tables': len(data.get('extracted_tables', [])),
                'Status': '✅ Success'
            })
        
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True)
        
        # Export options
        st.subheader("📥 Export Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Export summary as CSV
            csv_data = summary_df.to_csv(index=False)
            st.download_button(
                label="📊 Download Summary CSV",
                data=csv_data,
                file_name="document_processing_summary.csv",
                mime="text/csv"
            )
        
        with col2:
            # Export all extracted text
            all_text = ""
            for file_name, data in results.items():
                all_text += f"\n\n--- {file_name} ---\n\n"
                
                if data['file_type'] == 'PDF':
                    for page in data.get('pages', []):
                        all_text += f"Page {page['page_number']}:\n"
                        all_text += page.get('text', '') + "\n\n"
                elif data['file_type'] == 'DOCX':
                    for para in data.get('paragraphs', []):
                        all_text += para + "\n\n"
                elif data['file_type'] == 'PPTX':
                    for slide in data.get('slides', []):
                        all_text += f"Slide {slide['slide_number']}:\n"
                        all_text += slide.get('text', '') + "\n\n"
                elif data['file_type'] == 'Excel':
                    for sheet in data.get('sheets', []):
                        all_text += f"Sheet: {sheet['sheet_name']}\n"
                        all_text += sheet.get('text_content', '') + "\n\n"
            
            st.download_button(
                label="📄 Download All Text",
                data=all_text,
                file_name="extracted_text.txt",
                mime="text/plain"
            )
        
        with col3:
            # Export all tables as Excel
            if total_tables > 0:
                try:
                    # Create Excel file with multiple sheets
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        sheet_count = 0
                        
                        for file_name, data in results.items():
                            tables = data.get('extracted_tables', [])
                            if tables:
                                for i, table in enumerate(tables, 1):
                                    cleaned_df = clean_table_data(table.get('data'))
                                    if cleaned_df is not None and not cleaned_df.empty:
                                        sheet_count += 1
                                        sheet_name = f"{os.path.splitext(file_name)[0]}_T{i}"[:31]  # Excel sheet name limit
                                        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    if sheet_count > 0:
                        st.download_button(
                            label="📊 Download All Tables (Excel)",
                            data=output.getvalue(),
                            file_name="extracted_tables.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("No valid tables to export")
                        
                except Exception as e:
                    st.error(f"Error creating Excel export: {str(e)}")
            else:
                st.info("No tables available for export")
        
        # Processing insights
        st.subheader("💡 Processing Insights")
        
        insights = []
        
        # File type insights
        file_types = Counter(data['file_type'] for data in results.values())
        insights.append(f"Processed {len(file_types)} different file types")
        
        # Content insights
        if total_words > 0:
            avg_words_per_file = total_words / total_files
            insights.append(f"Average words per file: {avg_words_per_file:,.0f}")
        
        if total_images > 0:
            insights.append(f"Found images in {sum(1 for data in results.values() if data.get('total_images', 0) > 0)} files")
        
        if total_tables > 0:
            insights.append(f"Extracted tables from {sum(1 for data in results.values() if data.get('extracted_tables', []))} files")
        
        # Display insights
        for insight in insights:
            st.write(f"• {insight}")
        
        # Quality metrics
        st.subheader("🎯 Quality Metrics")
        
        # Success rate
        success_rate = 100.0  # All files in results are successful
        st.metric("Success Rate", f"{success_rate:.1f}%")
        
        # Processing efficiency
        if total_files > 0:
            avg_size = sum(data.get('file_size_mb', 0) for data in results.values()) / total_files
            st.metric("Average File Size", f"{avg_size:.2f} MB")
        
        # Table extraction quality
        if total_tables > 0:
            tables_with_confidence = [
                table for data in results.values() 
                for table in data.get('extracted_tables', []) 
                if 'confidence' in table
            ]
            
            if tables_with_confidence:
                avg_confidence = sum(
                    float(table['confidence']) for table in tables_with_confidence 
                    if isinstance(table['confidence'], (int, float))
                ) / len(tables_with_confidence)
                st.metric("Average Table Confidence", f"{avg_confidence:.1f}%")

    # Additional actions
    st.header("🔄 Additional Actions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Process New Files", type="secondary", use_container_width=True):
            # Clear session state
            st.session_state.extraction_results = None
            st.session_state.processing_complete = False
            st.session_state.uploaded_files = []
            st.session_state.selected_pages = {}
            st.rerun()
    
    with col2:
        if st.button("📋 Copy Summary to Clipboard", use_container_width=True):
            summary_text = f"""
Document Processing Summary
==========================
Total Files: {total_files}
Total Words: {total_words:,}
Total Characters: {total_chars:,}
Total Images: {total_images}
Total Tables: {total_tables}

File Details:
{summary_df.to_string(index=False)}
"""
            st.code(summary_text, language="text")
            st.success("Summary generated! Copy the text above to your clipboard.")

# Run the main application
if __name__ == "__main__":
    main()
