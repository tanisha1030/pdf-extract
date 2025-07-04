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
import logging
import threading
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Import your document extractor
from paste import OptimizedDocumentExtractor  # Assuming your code is in paste.py

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(90deg, #f0f8ff 0%, #e6f3ff 100%);
        border-radius: 10px;
        border: 2px solid #1f77b4;
    }
    
    .metric-container {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1f77b4;
        margin: 0.5rem 0;
    }
    
    .status-success {
        color: #28a745;
        font-weight: bold;
    }
    
    .status-error {
        color: #dc3545;
        font-weight: bold;
    }
    
    .status-processing {
        color: #ffc107;
        font-weight: bold;
    }
    
    .file-info {
        background: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        margin: 0.5rem 0;
    }
    
    .download-section {
        background: #e8f5e8;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #28a745;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}
if 'processing_history' not in st.session_state:
    st.session_state.processing_history = []
if 'extractor' not in st.session_state:
    st.session_state.extractor = None

def initialize_extractor():
    """Initialize the document extractor"""
    if st.session_state.extractor is None:
        st.session_state.extractor = OptimizedDocumentExtractor(
            max_workers=st.session_state.get('max_workers', 2),
            chunk_size=st.session_state.get('chunk_size', 25),
            enable_caching=st.session_state.get('enable_caching', True),
            memory_limit_mb=st.session_state.get('memory_limit', 1024)
        )

def get_file_icon(file_type):
    """Get appropriate icon for file type"""
    icons = {
        'PDF': '📄',
        'DOCX': '📝',
        'PPTX': '📊',
        'Excel': '📈',
        'Unknown': '📁'
    }
    return icons.get(file_type, '📁')

def format_file_size(size_bytes):
    """Format file size in human readable format"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"

def create_download_link(data, filename, link_text):
    """Create download link for data"""
    if isinstance(data, dict):
        data = json.dumps(data, indent=2, ensure_ascii=False)
    
    b64_data = base64.b64encode(data.encode()).decode()
    href = f'<a href="data:application/json;base64,{b64_data}" download="{filename}" style="text-decoration: none; color: #28a745; font-weight: bold;">{link_text}</a>'
    return href

def process_files(uploaded_files, progress_bar, status_text):
    """Process uploaded files"""
    initialize_extractor()
    
    # Save uploaded files temporarily
    temp_dir = tempfile.mkdtemp()
    file_paths = []
    
    try:
        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.getvalue())
            file_paths.append(file_path)
        
        # Process files
        total_files = len(file_paths)
        results = {}
        
        for i, file_path in enumerate(file_paths):
            progress = (i + 1) / total_files
            progress_bar.progress(progress)
            status_text.text(f"Processing {os.path.basename(file_path)}... ({i + 1}/{total_files})")
            
            try:
                # Process single file
                file_results = st.session_state.extractor.process_files_optimized([file_path])
                results.update(file_results)
                
            except Exception as e:
                logger.error(f"Error processing {file_path}: {e}")
                st.error(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        
        # Update session state
        st.session_state.processed_files.update(results)
        st.session_state.extractor.extracted_data.update(results)
        
        # Add to processing history
        st.session_state.processing_history.append({
            'timestamp': datetime.now().isoformat(),
            'files_processed': len(results),
            'files_list': [os.path.basename(fp) for fp in file_paths],
            'success': True
        })
        
        return results
        
    finally:
        # Cleanup temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)

def create_summary_charts(results):
    """Create summary charts for processed files"""
    if not results:
        return None, None
    
    # Collect data for charts
    file_types = []
    file_sizes = []
    processing_times = []
    page_counts = []
    word_counts = []
    
    for file_path, data in results.items():
        file_types.append(data.get('file_type', 'Unknown'))
        file_sizes.append(data.get('file_size_mb', 0))
        processing_times.append(data.get('processing_time_seconds', 0))
        
        if data.get('file_type') == 'PDF':
            page_counts.append(data.get('page_count', 0))
            word_counts.append(data.get('total_words', 0))
        elif data.get('file_type') == 'DOCX':
            page_counts.append(len(data.get('paragraphs', [])))
            word_counts.append(data.get('total_words', 0))
        elif data.get('file_type') == 'PPTX':
            page_counts.append(data.get('total_slides', 0))
            word_counts.append(data.get('total_words', 0))
        else:
            page_counts.append(0)
            word_counts.append(0)
    
    # Create charts
    fig1 = make_subplots(
        rows=2, cols=2,
        subplot_titles=('File Types Distribution', 'File Sizes (MB)', 
                       'Processing Times (seconds)', 'Content Volume'),
        specs=[[{"type": "pie"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "scatter"}]]
    )
    
    # File types pie chart
    type_counts = Counter(file_types)
    fig1.add_trace(
        go.Pie(labels=list(type_counts.keys()), values=list(type_counts.values()),
               name="File Types"),
        row=1, col=1
    )
    
    # File sizes bar chart
    fig1.add_trace(
        go.Bar(x=[f"File {i+1}" for i in range(len(file_sizes))], 
               y=file_sizes, name="File Size (MB)"),
        row=1, col=2
    )
    
    # Processing times bar chart
    fig1.add_trace(
        go.Bar(x=[f"File {i+1}" for i in range(len(processing_times))], 
               y=processing_times, name="Processing Time (s)"),
        row=2, col=1
    )
    
    # Content volume scatter plot
    fig1.add_trace(
        go.Scatter(x=page_counts, y=word_counts, mode='markers',
                   text=[f"File {i+1}" for i in range(len(page_counts))],
                   name="Pages vs Words"),
        row=2, col=2
    )
    
    fig1.update_layout(height=600, showlegend=True, title_text="Processing Summary")
    
    return fig1, None

def main():
    """Main application function"""
    
    # Header
    st.markdown('<div class="main-header">📄 Document Extractor Pro</div>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    st.sidebar.header("⚙️ Configuration")
    
    # Extractor settings
    st.sidebar.subheader("Processing Settings")
    max_workers = st.sidebar.slider("Max Workers", 1, 8, 2, help="Number of parallel processing threads")
    chunk_size = st.sidebar.slider("Chunk Size", 10, 100, 25, help="Pages processed per chunk")
    memory_limit = st.sidebar.slider("Memory Limit (MB)", 512, 4096, 1024, help="Maximum memory usage")
    enable_caching = st.sidebar.checkbox("Enable Caching", True, help="Cache results for faster reprocessing")
    
    # Update session state
    st.session_state.max_workers = max_workers
    st.session_state.chunk_size = chunk_size
    st.session_state.memory_limit = memory_limit
    st.session_state.enable_caching = enable_caching
    
    # Sidebar statistics
    if st.session_state.processed_files:
        st.sidebar.subheader("📊 Current Session Stats")
        st.sidebar.metric("Files Processed", len(st.session_state.processed_files))
        
        total_pages = sum(
            data.get('page_count', 0) if data.get('file_type') == 'PDF' else 
            len(data.get('paragraphs', [])) if data.get('file_type') == 'DOCX' else
            data.get('total_slides', 0) if data.get('file_type') == 'PPTX' else 0
            for data in st.session_state.processed_files.values()
        )
        st.sidebar.metric("Total Pages/Slides", total_pages)
        
        total_words = sum(
            data.get('total_words', 0) 
            for data in st.session_state.processed_files.values()
        )
        st.sidebar.metric("Total Words", f"{total_words:,}")
    
    # Main content area
    tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload & Process", "📋 Results", "📊 Analytics", "🔧 Tools"])
    
    with tab1:
        st.header("Upload Documents")
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Choose files to process",
            accept_multiple_files=True,
            type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
            help="Supported formats: PDF, DOCX, PPTX, XLSX, XLS"
        )
        
        if uploaded_files:
            st.subheader("📁 Uploaded Files")
            
            # Display uploaded files info
            for uploaded_file in uploaded_files:
                file_size = len(uploaded_file.getvalue())
                file_type = Path(uploaded_file.name).suffix.upper().replace('.', '')
                
                col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                with col1:
                    st.write(f"**{uploaded_file.name}**")
                with col2:
                    st.write(f"{get_file_icon(file_type)} {file_type}")
                with col3:
                    st.write(format_file_size(file_size))
                with col4:
                    st.write("✅ Ready")
            
            # Process button
            if st.button("🚀 Process Files", type="primary"):
                if len(uploaded_files) > 10:
                    st.warning("⚠️ Processing more than 10 files at once may take significant time.")
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                start_time = time.time()
                
                with st.spinner("Processing files..."):
                    results = process_files(uploaded_files, progress_bar, status_text)
                
                processing_time = time.time() - start_time
                
                if results:
                    st.success(f"✅ Successfully processed {len(results)} files in {processing_time:.2f} seconds!")
                    
                    # Show quick summary
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Files Processed", len(results))
                    with col2:
                        total_pages = sum(
                            data.get('page_count', 0) if data.get('file_type') == 'PDF' else 
                            len(data.get('paragraphs', [])) if data.get('file_type') == 'DOCX' else
                            data.get('total_slides', 0) if data.get('file_type') == 'PPTX' else 0
                            for data in results.values()
                        )
                        st.metric("Total Pages/Slides", total_pages)
                    with col3:
                        total_words = sum(data.get('total_words', 0) for data in results.values())
                        st.metric("Total Words", f"{total_words:,}")
                    with col4:
                        st.metric("Processing Time", f"{processing_time:.2f}s")
                    
                    # Show charts
                    fig, _ = create_summary_charts(results)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.error("❌ No files were successfully processed.")
    
    with tab2:
        st.header("Processing Results")
        
        if st.session_state.processed_files:
            # Results overview
            st.subheader("📊 Overview")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Files", len(st.session_state.processed_files))
            with col2:
                file_types = [data.get('file_type', 'Unknown') for data in st.session_state.processed_files.values()]
                st.metric("File Types", len(set(file_types)))
            with col3:
                total_size = sum(data.get('file_size_mb', 0) for data in st.session_state.processed_files.values())
                st.metric("Total Size", f"{total_size:.2f} MB")
            
            # Detailed results
            st.subheader("📋 Detailed Results")
            
            for file_path, data in st.session_state.processed_files.items():
                with st.expander(f"{get_file_icon(data.get('file_type', 'Unknown'))} {data.get('filename', 'Unknown')}"):
                    
                    # Basic info
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.write("**File Type:**", data.get('file_type', 'Unknown'))
                        st.write("**File Size:**", f"{data.get('file_size_mb', 0):.2f} MB")
                    with col2:
                        st.write("**Processing Time:**", f"{data.get('processing_time_seconds', 0):.2f} seconds")
                        if data.get('file_type') == 'PDF':
                            st.write("**Pages:**", data.get('page_count', 0))
                        elif data.get('file_type') == 'DOCX':
                            st.write("**Paragraphs:**", len(data.get('paragraphs', [])))
                        elif data.get('file_type') == 'PPTX':
                            st.write("**Slides:**", data.get('total_slides', 0))
                    with col3:
                        st.write("**Total Words:**", f"{data.get('total_words', 0):,}")
                        if data.get('file_type') == 'PDF':
                            st.write("**Images:**", data.get('total_images', 0))
                            st.write("**Tables:**", data.get('total_tables', 0))
                    
                    # Type-specific information
                    if data.get('file_type') == 'PDF':
                        if data.get('fonts_used'):
                            st.write("**Fonts Used:**", ", ".join(data.get('fonts_used', [])[:5]))
                        
                        if data.get('pages'):
                            st.write("**Sample Page Content:**")
                            sample_page = data['pages'][0] if data['pages'] else {}
                            sample_text = sample_page.get('text', '')[:500]
                            if sample_text:
                                st.text_area("", sample_text, height=100, disabled=True)
                    
                    elif data.get('file_type') == 'DOCX':
                        if data.get('paragraphs'):
                            st.write("**Sample Paragraph:**")
                            sample_para = data['paragraphs'][0] if data['paragraphs'] else {}
                            sample_text = sample_para.get('text', '')[:500]
                            if sample_text:
                                st.text_area("", sample_text, height=100, disabled=True)
                    
                    elif data.get('file_type') == 'PPTX':
                        if data.get('slides'):
                            st.write("**Sample Slide Content:**")
                            sample_slide = data['slides'][0] if data['slides'] else {}
                            sample_text = " ".join(sample_slide.get('text_content', []))[:500]
                            if sample_text:
                                st.text_area("", sample_text, height=100, disabled=True)
                    
                    elif data.get('file_type') == 'Excel':
                        if data.get('sheets'):
                            st.write("**Sheets:**")
                            for sheet in data['sheets'][:3]:  # Show first 3 sheets
                                st.write(f"- {sheet.get('sheet_name', 'Unknown')}: {sheet.get('shape', (0, 0))[0]} rows × {sheet.get('shape', (0, 0))[1]} columns")
            
            # Download section
            st.subheader("💾 Download Results")
            
            col1, col2 = st.columns(2)
            with col1:
                # Download individual file results
                st.write("**Individual File Results:**")
                for file_path, data in st.session_state.processed_files.items():
                    filename = data.get('filename', 'unknown')
                    json_filename = f"{Path(filename).stem}_results.json"
                    download_link = create_download_link(data, json_filename, f"📥 {filename}")
                    st.markdown(download_link, unsafe_allow_html=True)
            
            with col2:
                # Download all results
                st.write("**Complete Results:**")
                if st.button("📦 Download All Results"):
                    all_results = {
                        'processing_summary': {
                            'total_files': len(st.session_state.processed_files),
                            'processing_date': datetime.now().isoformat(),
                            'file_types': list(set(data.get('file_type', 'Unknown') for data in st.session_state.processed_files.values())),
                            'total_size_mb': sum(data.get('file_size_mb', 0) for data in st.session_state.processed_files.values()),
                            'total_processing_time': sum(data.get('processing_time_seconds', 0) for data in st.session_state.processed_files.values())
                        },
                        'results': st.session_state.processed_files
                    }
                    
                    download_link = create_download_link(all_results, "complete_extraction_results.json", "📥 Download Complete Results")
                    st.markdown(download_link, unsafe_allow_html=True)
        else:
            st.info("👋 No files have been processed yet. Go to the Upload & Process tab to get started!")
    
    with tab3:
        st.header("Analytics Dashboard")
        
        if st.session_state.processed_files:
            # Create and display charts
            fig, _ = create_summary_charts(st.session_state.processed_files)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            
            # Processing history
            if st.session_state.processing_history:
                st.subheader("📈 Processing History")
                
                history_df = pd.DataFrame(st.session_state.processing_history)
                history_df['timestamp'] = pd.to_datetime(history_df['timestamp'])
                
                # Timeline chart
                fig_timeline = px.line(
                    history_df, 
                    x='timestamp', 
                    y='files_processed',
                    title='Files Processed Over Time',
                    markers=True
                )
                st.plotly_chart(fig_timeline, use_container_width=True)
                
                # History table
                st.subheader("📋 Processing History Table")
                display_history = history_df.copy()
                display_history['timestamp'] = display_history['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
                display_history['files_list'] = display_history['files_list'].apply(lambda x: ', '.join(x[:3]) + ('...' if len(x) > 3 else ''))
                st.dataframe(display_history[['timestamp', 'files_processed', 'files_list', 'success']], use_container_width=True)
        else:
            st.info("📊 Analytics will be available after processing some files.")
    
    with tab4:
        st.header("Tools & Utilities")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🔧 System Tools")
            
            # Clear cache
            if st.button("🗑️ Clear Processing Cache"):
                st.session_state.processed_files = {}
                st.session_state.processing_history = []
                st.session_state.extractor = None
                st.success("✅ Cache cleared successfully!")
            
            # Memory usage (if psutil is available)
            try:
                import psutil
                process = psutil.Process()
                memory_info = process.memory_info()
                memory_mb = memory_info.rss / 1024 / 1024
                
                st.metric("Memory Usage", f"{memory_mb:.1f} MB")
                st.metric("CPU Usage", f"{psutil.cpu_percent():.1f}%")
            except ImportError:
                st.info("💡 Install psutil for system monitoring: `pip install psutil`")
        
        with col2:
            st.subheader("📚 Help & Information")
            
            st.markdown("""
            **Supported File Formats:**
            - 📄 PDF (.pdf)
            - 📝 Word Documents (.docx)
            - 📊 PowerPoint Presentations (.pptx)
            - 📈 Excel Spreadsheets (.xlsx, .xls)
            
            **Processing Features:**
            - Text extraction from all pages/slides
            - Image extraction (PDF only)
            - Table detection and extraction
            - Metadata extraction
            - Font analysis (PDF only)
            - Batch processing support
            
            **Performance Tips:**
            - Use caching for repeated processing
            - Adjust chunk size based on file size
            - Monitor memory usage for large files
            - Process similar file types together
            """)
            
            # System information
            st.subheader("🖥️ System Information")
            import platform
            st.write(f"**Platform:** {platform.system()} {platform.release()}")
            st.write(f"**Python:** {platform.python_version()}")
            st.write(f"**Streamlit:** {st.__version__}")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666; font-size: 0.9em;'>"
        "Document Extractor Pro - Powered by Streamlit | "
        f"Session started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
