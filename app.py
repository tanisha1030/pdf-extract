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
    
    .page-selector {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    
    .page-content {
        background: white;
        padding: 1.5rem;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    
    .image-container {
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1rem;
        background: #f8f9fa;
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
if 'selected_page' not in st.session_state:
    st.session_state.selected_page = {}

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

def decode_image_data(image_data):
    """Decode base64 image data to PIL Image"""
    try:
        if isinstance(image_data, str):
            # Remove data URL prefix if present
            if image_data.startswith('data:image'):
                image_data = image_data.split(',')[1]
            
            # Decode base64
            image_bytes = base64.b64decode(image_data)
            image = Image.open(BytesIO(image_bytes))
            return image
        return None
    except Exception as e:
        st.error(f"Error decoding image: {str(e)}")
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
            <div>
                <strong>📊 Table Extraction Methods:</strong>
                <ul>
                    <li>Tabula (structured tables)</li>
                    <li>Camelot (lattice tables)</li>
                    <li>PDFplumber (general purpose)</li>
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
                
                progress_bar.progress(1.0)
                status_text.text("✅ Processing completed!")
                
                # Auto-refresh to show results
                time.sleep(1)
                st.rerun()
            else:
                st.error("❌ No files were successfully processed.")
                
        except Exception as e:
            st.error(f"❌ Error during processing: {str(e)}")

def display_page_content(file_name, data, page_number):
    """Display content for a specific page"""
    pages = data.get('pages', [])
    
    if not pages or page_number < 1 or page_number > len(pages):
        st.error(f"Page {page_number} not found")
        return
    
    page = pages[page_number - 1]
    
    st.markdown(f'<div class="page-content">', unsafe_allow_html=True)
    
    # Page header
    st.subheader(f"📄 Page {page_number}")
    
    # Page statistics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Word Count", page.get('word_count', 0))
    with col2:
        st.metric("Character Count", page.get('char_count', 0))
    with col3:
        st.metric("Images", len(page.get('images', [])))
    with col4:
        st.metric("Tables", len(page.get('tables', [])))
    
    # Page content tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📝 Text", "📊 Tables", "🖼️ Images", "ℹ️ Metadata"])
    
    with tab1:
        st.subheader("Page Text")
        page_text = page.get('text', '')
        if page_text:
            st.text_area("Content", page_text, height=400, disabled=True)
        else:
            st.info("No text found on this page")
    
    with tab2:
        st.subheader("Page Tables")
        page_tables = page.get('tables', [])
        if page_tables:
            for i, table in enumerate(page_tables, 1):
                st.write(f"**Table {i}**")
                try:
                    if isinstance(table, dict) and 'data' in table:
                        table_data = table['data']
                        if isinstance(table_data, list) and table_data:
                            if isinstance(table_data[0], dict):
                                df = pd.DataFrame(table_data)
                            else:
                                df = pd.DataFrame(table_data[1:], columns=table_data[0])
                            st.dataframe(df, use_container_width=True)
                        else:
                            st.write("No table data available")
                    else:
                        st.write("Table format not recognized")
                except Exception as e:
                    st.error(f"Error displaying table: {str(e)}")
        else:
            st.info("No tables found on this page")
    
    with tab3:
        st.subheader("Page Images")
        page_images = page.get('images', [])
        if page_images:
            for i, img in enumerate(page_images, 1):
                st.markdown(f'<div class="image-container">', unsafe_allow_html=True)
                
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.write(f"**Image {i}**")
                    st.write(f"**Filename:** {os.path.basename(img.get('filename', 'unknown'))}")
                    st.write(f"**Size:** {img.get('width', 0)}×{img.get('height', 0)}")
                    st.write(f"**Color:** {img.get('colorspace', 'unknown')}")
                    st.write(f"**File size:** {img.get('size_bytes', 0)} bytes")
                
                with col2:
                    # Display the actual image
                    if 'image_data' in img:
                        try:
                            image = decode_image_data(img['image_data'])
                            if image:
                                st.image(image, caption=f"Image {i}", use_column_width=True)
                                
                                # Download button for individual image
                                if st.button(f"Download Image {i}", key=f"download_img_{file_name}_{page_number}_{i}"):
                                    # Convert PIL image to bytes
                                    img_buffer = BytesIO()
                                    image.save(img_buffer, format='PNG')
                                    img_bytes = img_buffer.getvalue()
                                    
                                    st.download_button(
                                        label="Download PNG",
                                        data=img_bytes,
                                        file_name=f"image_{page_number}_{i}.png",
                                        mime="image/png",
                                        key=f"download_png_{file_name}_{page_number}_{i}"
                                    )
                            else:
                                st.warning("Could not decode image data")
                        except Exception as e:
                            st.error(f"Error displaying image: {str(e)}")
                    else:
                        st.warning("No image data available")
                
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No images found on this page")
    
    with tab4:
        st.subheader("Page Metadata")
        if 'metadata' in page:
            for key, value in page['metadata'].items():
                st.write(f"**{key}:** {value}")
        else:
            st.info("No metadata available for this page")
    
    st.markdown('</div>', unsafe_allow_html=True)

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
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📄 File Details", "📖 Page Viewer", "📊 Tables", "🖼️ Images", "📋 Summary"])
    
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
                    # Show text preview
                    if data['file_type'] == 'PDF' and data.get('pages'):
                        first_page_text = data['pages'][0].get('text', '')[:500]
                        if first_page_text:
                            st.write("**Text Preview:**")
                            st.text(first_page_text + "..." if len(first_page_text) == 500 else first_page_text)
                    
                    # Show metadata
                    if data.get('metadata'):
                        st.write("**Metadata:**")
                        for key, value in data['metadata'].items():
                            st.write(f"  - {key}: {value}")
    
    with tab2:
        st.subheader("📖 Page-by-Page Viewer")
        
        # File selector
        pdf_files = {name: data for name, data in results.items() if data['file_type'] == 'PDF'}
        
        if pdf_files:
            selected_file = st.selectbox(
                "Select PDF file to view:",
                list(pdf_files.keys()),
                key="page_viewer_file_selector"
            )
            
            if selected_file:
                file_data = pdf_files[selected_file]
                pages = file_data.get('pages', [])
                
                if pages:
                    # Page selector
                    st.markdown('<div class="page-selector">', unsafe_allow_html=True)
                    col1, col2, col3 = st.columns([1, 2, 1])
                    
                    with col1:
                        if st.button("⬅️ Previous", key="prev_page"):
                            if selected_file not in st.session_state.selected_page:
                                st.session_state.selected_page[selected_file] = 1
                            if st.session_state.selected_page[selected_file] > 1:
                                st.session_state.selected_page[selected_file] -= 1
                    
                    with col2:
                        if selected_file not in st.session_state.selected_page:
                            st.session_state.selected_page[selected_file] = 1
                        
                        page_number = st.number_input(
                            f"Page (1-{len(pages)})",
                            min_value=1,
                            max_value=len(pages),
                            value=st.session_state.selected_page[selected_file],
                            key=f"page_selector_{selected_file}"
                        )
                        st.session_state.selected_page[selected_file] = page_number
                    
                    with col3:
                        if st.button("Next ➡️", key="next_page"):
                            if selected_file not in st.session_state.selected_page:
                                st.session_state.selected_page[selected_file] = 1
                            if st.session_state.selected_page[selected_file] < len(pages):
                                st.session_state.selected_page[selected_file] += 1
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Display selected page content
                    display_page_content(selected_file, file_data, page_number)
                else:
                    st.info("No pages found in the selected file")
        else:
            st.info("No PDF files found. Page viewer is only available for PDF files.")
    
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
                        with st.expander(f"Table {i} ({table.get('method', 'unknown')})"):
                            
                            # Show table metadata
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"**Method:** {table.get('method', 'unknown')}")
                                st.write(f"**Confidence:** {table.get('confidence', 'unknown')}")
                            with col2:
                                if 'shape' in table:
                                    st.write(f"**Shape:** {table['shape'][0]} rows × {table['shape'][1]} columns")
                                if 'accuracy' in table:
                                    st.write(f"**Accuracy:** {table['accuracy']:.1f}%")
                            
                            # Show table data
                            try:
                                if table.get('data'):
                                    if isinstance(table['data'], list) and isinstance(table['data'][0], dict):
                                        # DataFrame format
                                        df = pd.DataFrame(table['data'])
                                        st.dataframe(df, use_container_width=True)
                                    elif isinstance(table['data'], list):
                                        # List of lists format
                                        if table['data']:
                                            df = pd.DataFrame(table['data'][1:], columns=table['data'][0] if table['data'] else [])
                                            st.dataframe(df, use_container_width=True)
                                    else:
                                        st.write("Table data format not recognized")
                            except Exception as e:
                                st.error(f"Error displaying table: {str(e)}")
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
                                
                                # Display images in a grid
                                cols = st.columns(min(3, len(images)))
                                for idx, img in enumerate(images):
                                    with cols[idx % 3]:
                                        st.markdown(f'<div class="image-container">', unsafe_allow_html=True)
                                        
                                        # Image info
                                        st.write(f"**📸 Image {idx + 1}**")
                                        st.write(f"Size: {img.get('width', 0)}×{img.get('height', 0)}")
                                        st.write(f"Format: {img.get('colorspace', 'unknown')}")
                                        
                                        # Display image
                                        if 'image_data' in img:
                                            try:
                                                image = decode_image_data(img['image_data'])
                                                if image:
                                                    st.image(image, use_column_width=True)
                                                    
                                                    # Download button
                                                    img_buffer = BytesIO()
                                                    image.save(img_buffer, format='PNG')
                                                    img_bytes = img_buffer.getvalue()
                                                    
                                                    st.download_button(
                                                        label="⬇️ Download",
                                                        data=img_bytes,
                                                        file_name=f"{file_name}_page_{page['page_number']}_img_{idx + 1}.png",
                                                        mime="image/png",
                                                        key=f"download_img_{file_name}_{page['page_number']}_{idx}",
                                                        use_container_width=True
                                                    )
                                                else:
                                                    st.warning("Could not decode image")
                                            except Exception as e:
                                                st.error(f"Error: {str(e)}")
                                        else:
                                            st.info("No image data available")
                                        
                                        st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No images were extracted from the uploaded documents.")
    
    with tab5:
        st.subheader("📋 Processing Summary")
        
        # File type distribution
        file_types = Counter(data['file_type'] for data in results.values())
        st.write("**File Type Distribution:**")
        for file_type, count in file_types.items():
            st.write(f"  - {file_type}: {count} files")
        
        # Processing statistics
        st.write("\n**Processing Statistics:**")
        st.write(f"  - Total files processed: {total_files}")
        st.write(f"  - Total words extracted: {total_words:,}")
        st.write(f"  - Total characters: {total_chars:,}")
        st.write(f"  - Total images found: {total_images}")
        st.write(f"  - Total tables extracted: {total_tables}")
        
        # Font information (for PDFs)
        all_fonts = set()
        for data in results.values():
            if data['file_type'] == 'PDF':
                all_fonts.update(data.get('fonts_used', []))
        
        if all_fonts:
            st.write(f"\n**Fonts Used:** {', '.join(sorted(all_fonts))}")
        
        # Success message
        st.success("🎉 All documents processed successfully!")
    
    # Download section
    st.header("⬇️ Download Results")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # JSON download
        if st.button("📄 Download Full Results (JSON)", use_container_width=True):
            json_data = json.dumps(results, indent=2, ensure_ascii=False, default=str)
            st.download_button(
                label="Download JSON",
                data=json_data,
                file_name=f"extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True
            )
    
    with col2:
        # CSV download for tables
                # CSV download for tables
        if st.button("📊 Download All Tables (CSV)", use_container_width=True):
            # Create a zip file containing all tables as CSV files
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for file_name, data in results.items():
                    tables = data.get('extracted_tables', [])
                    if tables:
                        for i, table in enumerate(tables, 1):
                            try:
                                if table.get('data'):
                                    # Create DataFrame from table data
                                    if isinstance(table['data'], list) and isinstance(table['data'][0], dict):
                                        df = pd.DataFrame(table['data'])
                                    elif isinstance(table['data'], list):
                                        df = pd.DataFrame(table['data'][1:], columns=table['data'][0] if table['data'] else [])
                                    else:
                                        continue
                                    
                                    # Save DataFrame to CSV in memory
                                    csv_buffer = BytesIO()
                                    df.to_csv(csv_buffer, index=False)
                                    csv_buffer.seek(0)
                                    
                                    # Add to zip file
                                    base_name = os.path.splitext(file_name)[0]
                                    zip_file.writestr(
                                        f"{base_name}_table_{i}.csv",
                                        csv_buffer.getvalue()
                                    )
                            except Exception as e:
                                st.error(f"Error processing table {i} from {file_name}: {str(e)}")
            
            # Prepare the zip file for download
            zip_buffer.seek(0)
            st.download_button(
                label="Download Tables ZIP",
                data=zip_buffer,
                file_name=f"extracted_tables_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )

    # Images download section
    st.markdown("---")
    st.subheader("🖼️ Download Images")
    
    if total_images > 0:
        # Create a zip file containing all images
        if st.button("📸 Download All Images (ZIP)", use_container_width=True):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for file_name, data in results.items():
                    if data.get('total_images', 0) > 0:
                        # Handle PDF images
                        if data['file_type'] == 'PDF':
                            for page in data.get('pages', []):
                                images = page.get('images', [])
                                for idx, img in enumerate(images, 1):
                                    if 'image_data' in img:
                                        try:
                                            image = decode_image_data(img['image_data'])
                                            if image:
                                                # Save image to buffer
                                                img_buffer = BytesIO()
                                                image.save(img_buffer, format='PNG')
                                                img_buffer.seek(0)
                                                
                                                # Add to zip file
                                                base_name = os.path.splitext(file_name)[0]
                                                zip_file.writestr(
                                                    f"{base_name}_page_{page['page_number']}_img_{idx}.png",
                                                    img_buffer.getvalue()
                                                )
                                        except Exception as e:
                                            st.error(f"Error processing image {idx} from {file_name}: {str(e)}")
            
            # Prepare the zip file for download
            zip_buffer.seek(0)
            st.download_button(
                label="Download Images ZIP",
                data=zip_buffer,
                file_name=f"extracted_images_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )
    else:
        st.info("No images available for download")

if __name__ == "__main__":
    main()
