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
from pdfextract import OptimizedPDFExtractor

# Configure Streamlit page
st.set_page_config(
    page_title="PDF Extractor Pro",
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
    
    .table-info {
        background: #f8f9fa;
        padding: 0.5rem;
        border-radius: 4px;
        border-left: 3px solid #667eea;
        margin-bottom: 1rem;
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
    'create_summary': True,
    'max_workers': 4
}

def debug_table_structure(results):
    """Debug function to see exactly what's in your data"""
    print("\n=== DETAILED TABLE STRUCTURE DEBUG ===")
    
    for file_name, data in results.items():
        print(f"\n📄 FILE: {file_name}")
        print(f"Top-level keys: {list(data.keys())}")
        
        # Check every key that might contain tables
        for key in data.keys():
            if 'table' in key.lower():
                value = data[key]
                print(f"  🔍 {key}: {type(value)}")
                if isinstance(value, list):
                    print(f"    - List length: {len(value)}")
                    if len(value) > 0:
                        print(f"    - First item type: {type(value[0])}")
                        if isinstance(value[0], dict):
                            print(f"    - First item keys: {list(value[0].keys())}")
                elif isinstance(value, dict):
                    print(f"    - Dict keys: {list(value.keys())}")
        
        # Check pages structure
        if 'pages' in data and data['pages']:
            print(f"  📑 Pages: {len(data['pages'])} pages")
            
            page_table_counts = []
            for i, page in enumerate(data['pages']):
                page_tables = 0
                for key in page.keys():
                    if 'table' in key.lower():
                        if isinstance(page[key], list):
                            page_tables += len(page[key])
                        elif page[key]:
                            page_tables += 1
                
                if page_tables > 0:
                    page_table_counts.append(f"Page {i+1}: {page_tables}")
            
            if page_table_counts:
                print(f"    - Tables per page: {', '.join(page_table_counts)}")
    
    print("=== END DEBUG ===\n")

def get_tables_from_data_simple(data):
    """Simplified table extraction - only count unique tables once"""
    
    # First, try to find tables at root level
    root_tables = []
    
    # Check common root-level keys
    for key in ['tables', 'extracted_tables', 'all_tables']:
        if key in data and data[key]:
            root_tables = data[key]
            print(f"Found {len(root_tables)} tables in root key: {key}")
            break
    
    # If root tables exist, use only those (they're likely aggregated)
    if root_tables:
        return [t for t in root_tables if t and t != {}]
    
    # Otherwise, collect from pages but be very careful about duplicates
    all_tables = []
    if 'pages' in data:
        for page_num, page in enumerate(data['pages'], 1):
            page_tables = page.get('tables', [])
            if page_tables:
                print(f"Page {page_num} has {len(page_tables)} tables")
                all_tables.extend(page_tables)
    
    return [t for t in all_tables if t and t != {}]

def count_total_tables_debug(results):
    """Count tables with detailed debugging"""
    total_tables = 0
    
    print("\n=== TABLE COUNTING DEBUG ===")
    
    for file_name, data in results.items():
        print(f"\n📄 Processing: {file_name}")
        
        file_tables = get_tables_from_data_simple(data)
        file_count = len(file_tables)
        total_tables += file_count
        
        print(f"✅ Final count for {file_name}: {file_count} tables")
        
        # Show table details
        for i, table in enumerate(file_tables[:5], 1):  # Show first 5
            method = table.get('method', 'unknown')
            page = table.get('page_number', 'unknown')
            confidence = table.get('confidence', 'N/A')
            print(f"  Table {i}: Method={method}, Page={page}, Confidence={confidence}")
    
    print(f"\n🎯 TOTAL TABLES: {total_tables}")
    print("=== END COUNTING DEBUG ===\n")
    
    return total_tables

# Simple replacement functions for your main code
def get_tables_from_data(data):
    """Main function - use this to replace your existing one"""
    return get_tables_from_data_simple(data)

def count_total_tables(results):
    """Main function - use this to replace your existing one"""
    return count_total_tables_debug(results)
    
def convert_table_to_dataframe(table_data):
    """Convert various table formats to pandas DataFrame"""
    try:
        if not table_data:
            return None
        
        # Handle different data formats
        if isinstance(table_data, pd.DataFrame):
            return table_data
        
        elif isinstance(table_data, dict):
            # Handle dict format
            if 'data' in table_data:
                return convert_table_to_dataframe(table_data['data'])
            else:
                return pd.DataFrame(table_data)
        
        elif isinstance(table_data, list):
            if not table_data:
                return None
            
            # Check if it's a list of dictionaries
            if isinstance(table_data[0], dict):
                return pd.DataFrame(table_data)
            
            # Check if it's a list of lists
            elif isinstance(table_data[0], list):
                if len(table_data) > 1:
                    # First row as headers
                    headers = table_data[0]
                    data_rows = table_data[1:]
                    return pd.DataFrame(data_rows, columns=headers)
                else:
                    return pd.DataFrame(table_data)
            
            # Handle single list
            else:
                return pd.DataFrame([table_data])
        
        else:
            # Try to convert to string and then to DataFrame
            return pd.DataFrame([str(table_data)])
    
    except Exception as e:
        st.error(f"Error converting table data: {str(e)}")
        return None

def display_table_with_info(table, table_index, file_name):
    """Display a table with proper formatting and metadata"""
    
    # Table metadata
    st.markdown(f"""
    <div class="table-info">
        <strong>Table {table_index}</strong> - 
        Method: {table.get('method', 'unknown')} | 
        Page: {table.get('page_number', 'unknown')} | 
        Confidence: {table.get('confidence', 'N/A')}
    </div>
    """, unsafe_allow_html=True)
    
    # Convert table data to DataFrame
    table_data = table.get('data')
    if not table_data:
        st.warning("No data found in this table")
        return
    
    df = convert_table_to_dataframe(table_data)
    
    if df is not None and not df.empty:
        # Display table shape
        st.write(f"**Shape:** {df.shape[0]} rows × {df.shape[1]} columns")
        
        # Display the table
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )
        
        # Option to download individual table
        csv_data = df.to_csv(index=False)
        st.download_button(
            label=f"📥 Download Table {table_index} (CSV)",
            data=csv_data,
            file_name=f"{os.path.splitext(file_name)[0]}_table_{table_index}.csv",
            mime="text/csv",
            key=f"download_{file_name}_{table_index}"
        )
        
        # Show raw data for debugging (optional)
        with st.expander("🔍 Show Raw Data (Debug)"):
            st.json(table_data)
    
    else:
        st.error("Could not convert table data to displayable format")
        st.write("Raw table data:")
        st.json(table_data)

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📄 PDF Extractor Pro</h1>
        <p>Extract text, tables, images, and metadata from PDF files with optimized performance</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Features info
    st.markdown("""
    <div class="feature-info">
        <h3>🚀 Advanced PDF Processing Features</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1rem; margin-top: 1rem;">
            <div>
                <strong>📄 PDF Processing:</strong>
                <ul>
                    <li>Optimized for large PDF files</li>
                    <li>Chunked processing for memory efficiency</li>
                    <li>Intelligent caching system</li>
                </ul>
            </div>
            <div>
                <strong>🔧 Extraction Capabilities:</strong>
                <ul>
                    <li>Text content with formatting</li>
                    <li>Tables (multiple extraction methods)</li>
                    <li>Images with quality controls</li>
                    <li>Metadata and document structure</li>
                </ul>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Upload PDF Files")
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Choose PDF files to process",
            accept_multiple_files=True,
            type=['pdf'],
            help="Upload one or more PDF files for processing"
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
                process_pdfs(uploaded_files, DEFAULT_OPTIONS)
    
    with col2:
        st.header("📊 Processing Status")
        
        if not uploaded_files:
            st.info("👆 Upload PDF files to get started")
        else:
            # Show file statistics
            total_files = len(uploaded_files)
            total_size_mb = sum(len(f.getvalue()) / (1024 * 1024) for f in uploaded_files)
            
            st.metric("Files to Process", total_files)
            st.metric("Total Size", f"{total_size_mb:.2f} MB")
            
            # File size warnings
            if total_size_mb > 100:
                st.warning("Large files detected (>100MB total). Processing may take longer.")
    
    # Results section
    if st.session_state.processing_complete and st.session_state.extraction_results:
        display_results(st.session_state.extraction_results)

def process_pdfs(uploaded_files, options):
    """Process uploaded PDF files"""
    
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
            
            # Initialize extractor with default options
            status_text.text("Initializing PDF extractor...")
            extractor = OptimizedPDFExtractor(
                max_workers=options['max_workers'],
                enable_caching=True,
                memory_limit_mb=2048
            )
            
            # Process files
            status_text.text("Processing PDFs...")
            progress_bar.progress(0.1)
            
            results = {}
            for i, file_path in enumerate(file_paths):
                file_name = os.path.basename(file_path)
                status_text.text(f"Processing {file_name}...")
                
                try:
                    result = extractor.extract_pdf_chunked(file_path)
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
                
                # Initialize page selectors
                for file_name, data in results.items():
                    if file_name not in st.session_state.selected_pages:
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
        page_tables = []
        all_tables = get_tables_from_data(data)
        for table in all_tables:
            if table.get('page_number', 1) == selected_page:
                page_tables.append(table)
        
        st.write(f"• **Tables:** {len(page_tables)}")

def display_results(results):
    """Display extraction results"""
    
    st.header("🎉 Processing Complete!")
    
    # Overall statistics
    total_files = len(results)
    total_words = sum(data.get('total_words', 0) for data in results.values())
    total_chars = sum(data.get('total_characters', 0) for data in results.values())
    total_images = sum(data.get('total_images', 0) for data in results.values())
    
    # Fixed table counting using the new function
    total_tables = count_total_tables(results)
    
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
                    st.write(f"**File Type:** PDF")
                    st.write(f"**Size:** {data.get('file_size_mb', 0):.2f} MB")
                    st.write(f"**Pages:** {data.get('page_count', 0)}")
                    st.write(f"**Words:** {data.get('total_words', 0):,}")
                    st.write(f"**Images:** {data.get('total_images', 0)}")
                    
                    # Use the new table counting function
                    file_tables = get_tables_from_data(data)
                    st.write(f"**Tables:** {len(file_tables)}")
                
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
        
        # File selector
        selected_file = st.selectbox(
            "Select a PDF file to view:",
            options=list(results.keys()),
            key="page_view_file_selector"
        )
        
        if selected_file:
            file_data = results[selected_file]
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
    
    with tab3:
        st.subheader("📊 Extracted Tables")
        
        if total_tables > 0:
            # Table extraction methods summary
            method_counts = Counter()
            for data in results.values():
                file_tables = get_tables_from_data(data)
                for table in file_tables:
                    method_counts[table.get('method', 'unknown')] += 1
            
            if method_counts:
                st.write("**Extraction Methods Used:**")
                for method, count in method_counts.items():
                    st.write(f"  - {method}: {count} tables")
            
            # Show tables by file
            for file_name, data in results.items():
                file_tables = get_tables_from_data(data)
                
                if file_tables:
                    st.write(f"**📄 {file_name}**")
                    
                    for i, table in enumerate(file_tables, 1):
                        with st.expander(f"Table {i} - {table.get('method', 'unknown')} (Page {table.get('page_number', 'N/A')})"):
                            display_table_with_info(table, i, file_name)
        else:
            st.info("No tables were extracted from the PDF files.")
    
    with tab4:
        st.subheader("🖼️ Extracted Images")
        
        if total_images > 0:
            for file_name, data in results.items():
                if data.get('total_images', 0) > 0:
                    st.write(f"**📄 {file_name}**")
                    
                    # Show images by page
                    for page in data.get('pages', []):
                        images = page.get('images', [])
                        if images:
                            st.write(f"**Page {page['page_number']}:** {len(images)} images")
                            
                            # Display images in columns
                            cols = st.columns(min(3, len(images)))
                            
                            for idx, img in enumerate(images):
                                with cols[idx % 3]:
                                    try:
                                        # Try to display the image
                                        st.image(
                                            img['filename'],
                                            caption=f"Image {idx + 1}",
                                            width=300
                                        )
                                        
                                        # Image metadata
                                        with st.expander(f"Image {idx + 1} Details"):
                                            st.write(f"**Filename:** {os.path.basename(img['filename'])}")
                                            st.write(f"**Size:** {img.get('width', 'N/A')}×{img.get('height', 'N/A')}")
                                            st.write(f"**Colorspace:** {img.get('colorspace', 'N/A')}")
                                            st.write(f"**File size:** {img.get('size_bytes', 'N/A')} bytes")
                                    
                                    except Exception as e:
                                        st.error(f"Error displaying image: {str(e)}")
        else:
            st.info("No images were extracted from the PDF files.")
    
    with tab5:
        st.subheader("📈 Processing Summary")
        
        # File statistics
        st.write("**Processing Statistics:**")
        st.write(f"  - Total files processed: {total_files}")
        st.write(f"  - Total words extracted: {total_words:,}")
        st.write(f"  - Total characters: {total_chars:,}")
        st.write(f"  - Total images found: {total_images}")
        st.write(f"  - Total tables extracted: {total_tables}")
        
        # Font information
        all_fonts = set()
        for data in results.values():
            all_fonts.update(data.get('fonts_used', []))
        
        if all_fonts:
            st.write(f"\n**Fonts Used:** {', '.join(sorted(all_fonts))}")
        
        # Success message
        st.success("🎉 All PDFs processed successfully!")
    
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
                file_name=f"pdf_extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
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
                        
                        # Handle different possible keys for tables
                        tables = data.get('tables', []) or data.get('extracted_tables', []) or []
                        
                        for i, table in enumerate(tables, 1):
                            try:
                                table_data = table.get('data')
                                if table_data:
                                    csv_name = f"{file_base}_table_{i}_{table.get('method', 'unknown')}.csv"
                                    
                                    df = convert_table_to_dataframe(table_data)
                                    if df is not None and not df.empty:
                                        csv_data = df.to_csv(index=False)
                                        zip_file.writestr(csv_name, csv_data)
                                        table_count += 1
                            except Exception as e:
                                st.error(f"Error exporting table {i} from {file_name}: {str(e)}")
                                continue
                
                if table_count > 0:
                    st.download_button(
                        label=f"Download {table_count} Tables (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"extracted_tables_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                else:
                    st.error("No tables could be exported to CSV format.")
        else:
            st.info("No tables available for download.")
    
    # Text content download
    st.subheader("📝 Text Content Download")
    if st.button("📄 Download All Text Content", use_container_width=True):
        # Combine all text content
        all_text = []
        for file_name, data in results.items():
            all_text.append(f"=== {file_name} ===\n")
            
            for page in data.get('pages', []):
                all_text.append(f"\n--- Page {page['page_number']} ---\n")
                all_text.append(page.get('text', ''))
            
            all_text.append("\n\n")
        
        combined_text = '\n'.join(all_text)
        st.download_button(
            label="Download Combined Text",
            data=combined_text,
            file_name=f"extracted_text_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
