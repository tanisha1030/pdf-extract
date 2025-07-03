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
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None
if 'stats' not in st.session_state:
    st.session_state.stats = None

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
    
    # Main content area
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
        
        # Show processing status
        if st.session_state.processing_complete and st.session_state.stats:
            stats = st.session_state.stats
            
            st.markdown(f"""
            <div class="stats-card">
                <h4>📊 Processing Results</h4>
                <p><strong>Files Processed:</strong> {stats.get('total_files', 0)}</p>
                <p><strong>Total Words:</strong> {stats.get('total_words', 0):,}</p>
                <p><strong>Images Extracted:</strong> {stats.get('total_images', 0)}</p>
                <p><strong>Tables Found:</strong> {stats.get('total_tables', 0)}</p>
                <p><strong>Processing Time:</strong> {stats.get('processing_time', 0):.2f}s</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="feature-card">
                <h4>🎯 How to Use</h4>
                <ol>
                    <li>Upload your documents using the file uploader</li>
                    <li>Review the selected files</li>
                    <li>Click "Process Documents" to start extraction</li>
                    <li>Download the results when processing is complete</li>
                </ol>
            </div>
            """, unsafe_allow_html=True)
    
    # Results section
    if st.session_state.processing_complete and st.session_state.results:
        show_results(st.session_state.results, st.session_state.stats)

def process_documents(uploaded_files, show_detailed_logs, auto_download):
    """Process uploaded documents"""
    
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
            st.markdown(f"""
            <div class="error-message">
                <h4>❌ Processing Error</h4>
                <p>{str(e)}</p>
            </div>
            """, unsafe_allow_html=True)

def generate_stats(results, processing_time):
    """Generate processing statistics"""
    stats = {
        'total_files': len(results),
        'total_words': sum(data.get('total_words', 0) for data in results.values()),
        'total_characters': sum(data.get('total_characters', 0) for data in results.values()),
        'total_images': sum(data.get('total_images', 0) for data in results.values()),
        'total_tables': sum(len(data.get('extracted_tables', [])) for data in results.values()),
        'total_size_mb': sum(data.get('file_size_mb', 0) for data in results.values()),
        'processing_time': processing_time,
        'file_types': {}
    }
    
    # Count file types
    for data in results.values():
        file_type = data.get('file_type', 'Unknown')
        stats['file_types'][file_type] = stats['file_types'].get(file_type, 0) + 1
    
    return stats

def show_results(results, stats):
    """Display processing results"""
    st.markdown("## 📊 Processing Results")
    
    # Create tabs for different views
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Summary", "📄 Files", "📊 Tables", "🖼️ Images"])
    
    with tab1:
        show_summary_tab(results, stats)
    
    with tab2:
        show_files_tab(results)
    
    with tab3:
        show_tables_tab(results)
    
    with tab4:
        show_images_tab(results)
    
    # Download section
    st.markdown("## 📥 Download Results")
    create_download_section(results, stats)

def show_summary_tab(results, stats):
    """Show summary statistics"""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Files Processed", stats['total_files'])
        st.metric("Total Words", f"{stats['total_words']:,}")
    
    with col2:
        st.metric("Images Extracted", stats['total_images'])
        st.metric("Tables Found", stats['total_tables'])
    
    with col3:
        st.metric("Processing Time", f"{stats['processing_time']:.2f}s")
        st.metric("Total Size", f"{stats['total_size_mb']:.2f} MB")
    
    # File type distribution
    if stats['file_types']:
        st.markdown("### 📈 File Type Distribution")
        file_type_df = pd.DataFrame(list(stats['file_types'].items()), 
                                   columns=['File Type', 'Count'])
        st.bar_chart(file_type_df.set_index('File Type'))
    
    # Detailed file information
    st.markdown("### 📋 File Details")
    file_details = []
    for file_path, data in results.items():
        file_details.append({
            'Filename': data['filename'],
            'Type': data['file_type'],
            'Size (MB)': f"{data.get('file_size_mb', 0):.2f}",
            'Words': data.get('total_words', 0),
            'Images': data.get('total_images', 0),
            'Tables': len(data.get('extracted_tables', [])),
            'Pages/Slides': data.get('page_count', data.get('total_slides', 0))
        })
    
    df_details = pd.DataFrame(file_details)
    st.dataframe(df_details, use_container_width=True)

def show_files_tab(results):
    """Show individual file results"""
    st.markdown("### 📄 Individual File Results")
    
    # File selector
    file_names = [data['filename'] for data in results.values()]
    selected_file = st.selectbox("Select a file to view details:", file_names)
    
    if selected_file:
        # Find the selected file data
        file_data = None
        for data in results.values():
            if data['filename'] == selected_file:
                file_data = data
                break
        
        if file_data:
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📊 File Information")
                st.write(f"**Type:** {file_data['file_type']}")
                st.write(f"**Size:** {file_data.get('file_size_mb', 0):.2f} MB")
                st.write(f"**Words:** {file_data.get('total_words', 0):,}")
                st.write(f"**Characters:** {file_data.get('total_characters', 0):,}")
                
                if file_data['file_type'] == 'PDF':
                    st.write(f"**Pages:** {file_data.get('page_count', 0)}")
                    st.write(f"**Images:** {file_data.get('total_images', 0)}")
                    if file_data.get('fonts_used'):
                        st.write(f"**Fonts:** {len(file_data['fonts_used'])}")
                
                elif file_data['file_type'] == 'PPTX':
                    st.write(f"**Slides:** {file_data.get('total_slides', 0)}")
                
                elif file_data['file_type'] == 'Excel':
                    st.write(f"**Sheets:** {len(file_data.get('sheets', []))}")
            
            with col2:
                st.markdown("#### 📝 Content Preview")
                
                # Show text preview
                if file_data['file_type'] == 'PDF' and file_data.get('pages'):
                    first_page = file_data['pages'][0]
                    preview_text = first_page.get('text', '')[:500]
                    if preview_text:
                        st.text_area("Text Preview", preview_text, height=200)
                
                elif file_data['file_type'] == 'DOCX' and file_data.get('paragraphs'):
                    first_paragraphs = file_data['paragraphs'][:3]
                    preview_text = '\n\n'.join([p['text'] for p in first_paragraphs])[:500]
                    if preview_text:
                        st.text_area("Text Preview", preview_text, height=200)
            
            # Show metadata if available
            if file_data.get('metadata'):
                st.markdown("#### 📋 Metadata")
                metadata_df = pd.DataFrame(list(file_data['metadata'].items()), 
                                         columns=['Property', 'Value'])
                st.dataframe(metadata_df, use_container_width=True)

def show_tables_tab(results):
    """Show extracted tables"""
    st.markdown("### 📊 Extracted Tables")
    
    # Collect all tables
    all_tables = []
    for file_path, data in results.items():
        filename = data['filename']
        for i, table in enumerate(data.get('extracted_tables', [])):
            all_tables.append({
                'filename': filename,
                'table_index': i,
                'method': table.get('method', 'unknown'),
                'confidence': table.get('confidence', 'unknown'),
                'shape': table.get('shape', (0, 0)),
                'data': table.get('data', [])
            })
    
    if all_tables:
        # Table selector
        table_options = [f"{t['filename']} - Table {t['table_index']+1} ({t['method']})" 
                        for t in all_tables]
        selected_table = st.selectbox("Select a table to view:", table_options)
        
        if selected_table:
            # Find selected table
            table_index = table_options.index(selected_table)
            table_data = all_tables[table_index]
            
            # Show table info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Extraction Method", table_data['method'])
            with col2:
                st.metric("Confidence", table_data['confidence'])
            with col3:
                st.metric("Shape", f"{table_data['shape'][0]} x {table_data['shape'][1]}")
            
            # Show table data
            if table_data['data']:
                try:
                    # Convert to DataFrame for display
                    if isinstance(table_data['data'], list):
                        if all(isinstance(row, dict) for row in table_data['data']):
                            # List of dictionaries
                            df = pd.DataFrame(table_data['data'])
                        else:
                            # List of lists
                            df = pd.DataFrame(table_data['data'])
                    else:
                        df = pd.DataFrame(table_data['data'])
                    
                    st.dataframe(df, use_container_width=True)
                    
                    # Download button for individual table
                    csv_data = df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Table as CSV",
                        data=csv_data,
                        file_name=f"{table_data['filename']}_table_{table_data['table_index']+1}.csv",
                        mime="text/csv"
                    )
                    
                except Exception as e:
                    st.error(f"Error displaying table: {str(e)}")
                    st.json(table_data['data'])
    else:
        st.info("No tables were extracted from the uploaded documents.")

def show_images_tab(results):
    """Show extracted images"""
    st.markdown("### 🖼️ Extracted Images")
    
    # Collect all images
    all_images = []
    for file_path, data in results.items():
        filename = data['filename']
        if data['file_type'] == 'PDF':
            for page in data.get('pages', []):
                for image in page.get('images', []):
                    all_images.append({
                        'filename': filename,
                        'page': page['page_number'],
                        'image_path': image['filename'],
                        'width': image['width'],
                        'height': image['height'],
                        'size_bytes': image['size_bytes']
                    })
    
    if all_images:
        st.write(f"Found {len(all_images)} images")
        
        # Display images in a grid
        cols = st.columns(3)
        for i, image_info in enumerate(all_images):
            with cols[i % 3]:
                st.markdown(f"**{image_info['filename']}** (Page {image_info['page']})")
                st.write(f"Size: {image_info['width']}x{image_info['height']}")
                st.write(f"File size: {image_info['size_bytes']} bytes")
                
                # Try to display image if file exists
                if os.path.exists(image_info['image_path']):
                    try:
                        st.image(image_info['image_path'], use_column_width=True)
                    except Exception as e:
                        st.error(f"Error displaying image: {str(e)}")
                else:
                    st.warning("Image file not found")
    else:
        st.info("No images were extracted from the uploaded documents.")

def create_download_section(results, stats):
    """Create download section with various export options"""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # JSON download
        json_data = json.dumps(results, indent=2, default=str)
        st.download_button(
            label="📄 Download JSON Results",
            data=json_data,
            file_name=f"extraction_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json"
        )
    
    with col2:
        # Summary report download
        summary_report = create_summary_report(results, stats)
        st.download_button(
            label="📊 Download Summary Report",
            data=summary_report,
            file_name=f"extraction_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )
    
    with col3:
        # Excel download with all tables
        excel_data = create_excel_export(results)
        if excel_data:
            st.download_button(
                label="📈 Download Excel Report",
                data=excel_data,
                file_name=f"extraction_tables_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def create_download_files(results, stats):
    """Create downloadable files"""
    # This function would be called for auto-download
    pass

def create_summary_report(results, stats):
    """Create a text summary report"""
    report = []
    report.append("📋 DOCUMENT EXTRACTION SUMMARY REPORT")
    report.append("=" * 60)
    report.append(f"📅 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append(f"📁 Total Files Processed: {stats['total_files']}")
    report.append("")
    
    # Overall statistics
    report.append("📊 OVERALL STATISTICS:")
    report.append("-" * 30)
    report.append(f"📝 Total Words: {stats['total_words']:,}")
    report.append(f"🔤 Total Characters: {stats['total_characters']:,}")
    report.append(f"🖼️  Total Images: {stats['total_images']:,}")
    report.append(f"📊 Total Tables: {stats['total_tables']:,}")
    report.append(f"💾 Total File Size: {stats['total_size_mb']:.2f} MB")
    report.append(f"⏱️  Processing Time: {stats['processing_time']:.2f} seconds")
    report.append("")
    
    # File type distribution
    if stats['file_types']:
        report.append("📈 FILE TYPE DISTRIBUTION:")
        report.append("-" * 30)
        for file_type, count in stats['file_types'].items():
            report.append(f"{file_type}: {count} files")
        report.append("")
    
    # Detailed file information
    report.append("📄 DETAILED FILE INFORMATION:")
    report.append("=" * 60)
    
    for file_path, data in results.items():
        report.append(f"\n📁 File: {data['filename']}")
        report.append(f"   Type: {data['file_type']}")
        report.append(f"   Size: {data.get('file_size_mb', 0):.2f} MB")
        report.append(f"   Words: {data.get('total_words', 0):,}")
        report.append(f"   Characters: {data.get('total_characters', 0):,}")
        
        if data['file_type'] == 'PDF':
            report.append(f"   Pages: {data.get('page_count', 0)}")
            report.append(f"   Images: {data.get('total_images', 0)}")
            report.append(f"   Tables: {len(data.get('extracted_tables', []))}")
        elif data['file_type'] == 'PPTX':
            report.append(f"   Slides: {data.get('total_slides', 0)}")
        elif data['file_type'] == 'Excel':
            report.append(f"   Sheets: {len(data.get('sheets', []))}")
    
    return "\n".join(report)

def create_excel_export(results):
    """Create Excel file with all extracted tables"""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sheet_num = 1
            
            for file_path, data in results.items():
                filename = data['filename']
                
                # Add summary sheet
                summary_data = {
                    'Filename': [filename],
                    'Type': [data['file_type']],
                    'Size (MB)': [data.get('file_size_mb', 0)],
                    'Words': [data.get('total_words', 0)],
                    'Images': [data.get('total_images', 0)],
                    'Tables': [len(data.get('extracted_tables', []))]
                }
                
                summary_df = pd.DataFrame(summary_data)
                sheet_name = f"Summary_{sheet_num}"
                summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Add table sheets
                for i, table in enumerate(data.get('extracted_tables', [])):
                    try:
                        if table.get('data'):
                            if isinstance(table['data'], list):
                                if all(isinstance(row, dict) for row in table['data']):
                                    df = pd.DataFrame(table['data'])
                                else:
                                    df = pd.DataFrame(table['data'])
                            else:
                                df = pd.DataFrame(table['data'])
                            
                            sheet_name = f"Table_{sheet_num}_{i+1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                    except Exception as e:
                        continue
                
                sheet_num += 1
        
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        st.error(f"Error creating Excel export: {str(e)}")
        return None

# Run the app
if __name__ == "__main__":
    # Clean up old files on startup
    try:
        cleanup_old_files()
    except:
        pass
    
    main()
