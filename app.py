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
    }
    
    .image-card img {
        max-width: 300px;
        max-height: 300px;
        border-radius: 4px;
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
    st.session_state.selected_page = 1

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
    tab1, tab2, tab3, tab4 = st.tabs(["📄 File Details", "📊 Tables", "🖼️ Images", "📋 Summary"])
    
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
                    # Page selector for PDF files
                    if data['file_type'] == 'PDF' and data.get('pages'):
                        page_count = data.get('page_count', 1)
                        st.session_state.selected_page = st.selectbox(
                            "Select page to view",
                            options=range(1, page_count + 1),
                            index=st.session_state.selected_page - 1,
                            key=f"page_select_{file_name}"
                        )
                        
                        # Show text preview for selected page
                        selected_page_data = data['pages'][st.session_state.selected_page - 1]
                        page_text = selected_page_data.get('text', '')[:500]
                        if page_text:
                            st.write(f"**Page {st.session_state.selected_page} Text Preview:**")
                            st.text(page_text + "..." if len(page_text) == 500 else page_text)
                    
                    # Show metadata
                    if data.get('metadata'):
                        st.write("**Metadata:**")
                        for key, value in data['metadata'].items():
                            st.write(f"  - {key}: {value}")
    
    with tab2:
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
    
    with tab3:
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
                                st.markdown('<div class="image-container">', unsafe_allow_html=True)
                                
                                for img in images:
                                    try:
                                        # Check if image data is available
                                        if 'image_data' in img:
                                            # Display the image
                                            st.markdown(f"""
                                            <div class="image-card">
                                                <img src="data:image/png;base64,{img['image_data']}">
                                                <p><strong>{os.path.basename(img['filename'])}</strong></p>
                                                <p>Size: {img['width']}×{img['height']}</p>
                                                <p>Color: {img['colorspace']}</p>
                                                <p>File size: {img['size_bytes']} bytes</p>
                                            </div>
                                            """, unsafe_allow_html=True)
                                        else:
                                            # Fallback to just showing image info
                                            st.write(f"📸 {os.path.basename(img['filename'])}")
                                            st.write(f"Size: {img['width']}×{img['height']}")
                                            st.write(f"Color: {img['colorspace']}")
                                            st.write(f"File size: {img['size_bytes']} bytes")
                                    except Exception as e:
                                        st.error(f"Error displaying image: {str(e)}")
                                
                                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No images were extracted from the uploaded documents.")
    
    with tab4:
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
        if total_tables > 0:
            if st.button("📊 Download All Tables (CSV)", use_container_width=True):
                # Create ZIP file with all tables
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    
                    table_count = 0
                    for file_name, data in results.items():
                        file_base = os.path.splitext(file_name)[0]
                        
                        for i, table in enumerate(data.get('extracted_tables', []), 1):
                            try:
                                if table.get('data'):
                                    csv_name = f"{file_base}_table_{i}_{table.get('method', 'unknown')}.csv"
                                    
                                    if isinstance(table['data'], list) and isinstance(table['data'][0], dict):
                                        df = pd.DataFrame(table['data'])
                                    elif isinstance(table['data'], list):
                                        df = pd.DataFrame(table['data'][1:], columns=table['data'][0] if table['data'] else [])
                                    else:
                                        continue
                                    
                                    csv_data = df.to_csv(index=False)
                                    zip_file.writestr(csv_name, csv_data)
                                    table_count += 1
                            except Exception as e:
                                continue
                
                if table_count > 0:
                    st.download_button(
                        label=f"Download {table_count} Tables (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"extracted_tables_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
    
    # Clear results button
    if st.button("🔄 Process New Documents", use_container_width=True):
        st.session_state.extraction_results = None
        st.session_state.processing_complete = False
        st.session_state.uploaded_files = []
        st.session_state.selected_page = 1
        st.rerun()

if __name__ == "__main__":
    main()
