import streamlit as st
import os
import json
import pandas as pd
import tempfile
import shutil
from pdfextract import ComprehensiveDocumentExtractor

# Page configuration
st.set_page_config(
    page_title="SmartDoc Extractor", 
    layout="wide",
    page_icon="📄"
)

# Custom CSS styling
st.markdown("""
    <style>
    .main-title {
        font-size: 2.5em;
        color: #3c3c3c;
        margin-bottom: 0.5em;
        text-align: center;
    }
    .section-header {
        font-size: 1.5em;
        color: #1a73e8;
        margin-top: 1em;
        border-bottom: 2px solid #1a73e8;
        padding-bottom: 0.2em;
    }
    .info-box {
        background-color: #f1f3f4;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 1em;
        border-left: 4px solid #1a73e8;
    }
    .metrics-container {
        display: flex;
        justify-content: space-around;
        margin: 1em 0;
    }
    .metric-box {
        background-color: #e8f0fe;
        padding: 10px;
        border-radius: 5px;
        text-align: center;
        margin: 0 5px;
        flex: 1;
    }
    .stAlert {
        margin-top: 1em;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_results' not in st.session_state:
    st.session_state.processed_results = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None

def cleanup_temp_files():
    """Clean up temporary files"""
    if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
        try:
            shutil.rmtree(st.session_state.temp_dir)
            st.session_state.temp_dir = None
        except Exception as e:
            st.warning(f"Could not clean up temporary files: {str(e)}")

def display_file_summary(data):
    """Display file summary information"""
    file_type = data.get('file_type', 'Unknown')
    
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
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📃 Slides", data.get('total_slides', 0))
        with col2:
            st.metric("📝 Words", f"{data.get('total_words', 0):,}")
    
    elif file_type == 'Excel':
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📃 Sheets", len(data.get('sheets', [])))
        with col2:
            total_rows = sum(sheet.get('shape', [0, 0])[0] for sheet in data.get('sheets', []))
            st.metric("📊 Total Rows", f"{total_rows:,}")

def display_pdf_content(data, file_path):
    """Display PDF content with page selector"""
    if not data.get('pages'):
        st.warning("No page content found in this PDF.")
        return
    
    # Page selector
    page_numbers = list(range(1, data['page_count'] + 1))
    selected_page = st.selectbox(
        "📄 Select a page to view its content:", 
        page_numbers, 
        key=f"page_selector_{file_path}"
    )
    
    # Find selected page data
    selected_page_data = next(
        (p for p in data['pages'] if p['page_number'] == selected_page), 
        None
    )
    
    if not selected_page_data:
        st.error(f"Could not find data for page {selected_page}")
        return
    
    # Display page info
    st.markdown(f"### 🧾 Page {selected_page}")
    
    # Page metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Words", selected_page_data.get('word_count', 0))
    with col2:
        st.metric("Characters", selected_page_data.get('char_count', 0))
    with col3:
        st.metric("Images", len(selected_page_data.get('images', [])))
    
    # Display text content
    page_text = selected_page_data.get('text', '').strip()
    if page_text:
        st.markdown("**📝 Text Content:**")
        with st.expander("View full text", expanded=False):
            st.text_area(
                "Page text", 
                value=page_text, 
                height=300,
                key=f"text_area_{file_path}_{selected_page}"
            )
    else:
        st.info("No text content found on this page.")
    
    # Display images
    if selected_page_data.get('images'):
        st.markdown("**🖼️ Images on this page:**")
        for i, img in enumerate(selected_page_data['images']):
            try:
                if os.path.exists(img['filename']):
                    st.image(
                        img['filename'], 
                        caption=f"Image {i+1} - {os.path.basename(img['filename'])}",
                        use_container_width=True
                    )
                else:
                    st.warning(f"Image file not found: {img['filename']}")
            except Exception as e:
                st.error(f"Error displaying image {i+1}: {str(e)}")
    
    # Display tables
    if selected_page_data.get('tables'):
        st.markdown("**📋 Tables on this page:**")
        for i, table in enumerate(selected_page_data['tables']):
            try:
                st.markdown(f"*Table {i+1}:*")
                table_data = table.get('data', [])
                if table_data:
                    # Try to create a proper DataFrame
                    df = pd.DataFrame(table_data)
                    st.dataframe(df, use_container_width=True)
                else:
                    st.info("No data in this table.")
            except Exception as e:
                st.error(f"Error displaying table {i+1}: {str(e)}")

def display_docx_content(data):
    """Display DOCX content"""
    paragraphs = data.get('paragraphs', [])
    if paragraphs:
        st.markdown("**📝 Document Content:**")
        with st.expander("View paragraphs", expanded=False):
            for i, para in enumerate(paragraphs[:10]):  # Show first 10 paragraphs
                st.markdown(f"**Paragraph {i+1}** ({para.get('style', 'Normal')}):")
                st.write(para.get('text', ''))
                st.markdown("---")
        
        if len(paragraphs) > 10:
            st.info(f"Showing first 10 paragraphs. Total: {len(paragraphs)}")
    
    # Display tables
    tables = data.get('tables', [])
    if tables:
        st.markdown("**📋 Tables:**")
        for i, table in enumerate(tables):
            try:
                st.markdown(f"*Table {i+1}:*")
                table_data = table.get('data', [])
                if table_data:
                    df = pd.DataFrame(table_data)
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Error displaying table {i+1}: {str(e)}")

def display_pptx_content(data):
    """Display PPTX content"""
    slides = data.get('slides', [])
    if slides:
        st.markdown("**📊 Slides:**")
        slide_numbers = list(range(1, len(slides) + 1))
        selected_slide = st.selectbox("Select a slide to view:", slide_numbers)
        
        if selected_slide <= len(slides):
            slide_data = slides[selected_slide - 1]
            st.markdown(f"### Slide {selected_slide}")
            
            # Display text content
            text_content = slide_data.get('text_content', [])
            if text_content:
                st.markdown("**Text Content:**")
                for content in text_content:
                    st.write(f"• {content}")
            
            # Display tables
            tables = slide_data.get('tables', [])
            if tables:
                st.markdown("**Tables:**")
                for i, table in enumerate(tables):
                    try:
                        st.markdown(f"*Table {i+1}:*")
                        df = pd.DataFrame(table)
                        st.dataframe(df, use_container_width=True)
                    except Exception as e:
                        st.error(f"Error displaying table {i+1}: {str(e)}")

def display_excel_content(data):
    """Display Excel content"""
    sheets = data.get('sheets', [])
    if sheets:
        st.markdown("**📊 Sheets:**")
        sheet_names = [sheet['sheet_name'] for sheet in sheets]
        selected_sheet = st.selectbox("Select a sheet to view:", sheet_names)
        
        # Find selected sheet data
        sheet_data = next(
            (sheet for sheet in sheets if sheet['sheet_name'] == selected_sheet), 
            None
        )
        
        if sheet_data:
            st.markdown(f"### Sheet: {selected_sheet}")
            st.write(f"**Shape**: {sheet_data['shape'][0]:,} rows × {sheet_data['shape'][1]} columns")
            
            # Display data
            data_records = sheet_data.get('data', [])
            if data_records:
                df = pd.DataFrame(data_records)
                st.dataframe(df, use_container_width=True)
            else:
                st.info("No data to display for this sheet.")
            
            # Display column information
            if sheet_data.get('columns'):
                st.markdown("**Columns:**")
                st.write(", ".join(sheet_data['columns']))

# Main application
def main():
    # Header
    st.markdown("<div class='main-title'>📄 SmartDoc Extractor</div>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2em;'>
        <p>Upload PDF, Word, PowerPoint, or Excel files to extract text, tables, images, and metadata.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # File uploader
    uploaded_files = st.file_uploader(
        "📁 Upload your files here:",
        type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Supported formats: PDF, DOCX, PPTX, XLSX, XLS"
    )
    
    if uploaded_files:
        # Process files button
        if st.button("🚀 Process Files", type="primary"):
            # Clean up previous temp files
            cleanup_temp_files()
            
            # Create temporary directory
            st.session_state.temp_dir = tempfile.mkdtemp()
            
            with st.spinner("Processing files... Please wait."):
                try:
                    # Initialize extractor
                    extractor = ComprehensiveDocumentExtractor()
                    
                    # Save uploaded files to temp directory
                    saved_file_paths = []
                    for uploaded_file in uploaded_files:
                        file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
                        with open(file_path, 'wb') as f:
                            f.write(uploaded_file.read())
                        saved_file_paths.append(file_path)
                    
                    # Process files
                    results = extractor.process_files(saved_file_paths)
                    
                    if results:
                        st.session_state.processed_results = results
                        st.success(f"✅ Successfully processed {len(results)} file(s)!")
                    else:
                        st.error("❌ No files could be processed. Please check your files and try again.")
                
                except Exception as e:
                    st.error(f"❌ An error occurred during processing: {str(e)}")
                    st.session_state.processed_results = None
    
    # Display results if available
    if st.session_state.processed_results:
        st.markdown("---")
        st.markdown("## 📊 Processing Results")
        
        # Create tabs for each file
        if len(st.session_state.processed_results) > 1:
            file_tabs = st.tabs([data['filename'] for data in st.session_state.processed_results.values()])
            
            for i, (file_path, data) in enumerate(st.session_state.processed_results.items()):
                with file_tabs[i]:
                    display_file_results(data, file_path)
        else:
            # Single file - display directly
            file_path, data = next(iter(st.session_state.processed_results.items()))
            display_file_results(data, file_path)
        
        # Summary and download section
        st.markdown("---")
        st.markdown("## 📥 Download Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Create and offer summary report
            try:
                extractor = ComprehensiveDocumentExtractor()
                extractor.create_summary_report(st.session_state.processed_results)
                
                if os.path.exists("extraction_summary.txt"):
                    with open("extraction_summary.txt", "r", encoding="utf-8") as f:
                        summary_text = f.read()
                    
                    st.download_button(
                        label="📄 Download Summary Report",
                        data=summary_text,
                        file_name="extraction_summary.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Could not create summary report: {str(e)}")
        
        with col2:
            # Offer complete JSON download
            try:
                json_data = json.dumps(st.session_state.processed_results, indent=2, ensure_ascii=False, default=str)
                st.download_button(
                    label="📋 Download Complete JSON",
                    data=json_data,
                    file_name="complete_extraction_results.json",
                    mime="application/json",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Could not create JSON download: {str(e)}")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; margin-top: 2em;'>
        <p>SmartDoc Extractor - Comprehensive Document Processing Tool</p>
        <p>Supports PDF, Word, PowerPoint, and Excel files</p>
    </div>
    """, unsafe_allow_html=True)

def display_file_results(data, file_path):
    """Display results for a single file"""
    # File header
    st.markdown(f"<div class='section-header'>📄 {data['filename']}</div>", unsafe_allow_html=True)
    
    # File info box
    st.markdown(f"""
    <div class='info-box'>
        <strong>Type:</strong> {data.get('file_type', 'Unknown')}<br>
        <strong>Size:</strong> {data.get('file_size_mb', 0)} MB
    </div>
    """, unsafe_allow_html=True)
    
    # Display file summary
    display_file_summary(data)
    
    # Display content based on file type
    file_type = data.get('file_type', 'Unknown')
    
    if file_type == 'PDF':
        display_pdf_content(data, file_path)
    elif file_type == 'DOCX':
        display_docx_content(data)
    elif file_type == 'PPTX':
        display_pptx_content(data)
    elif file_type == 'Excel':
        display_excel_content(data)
    
    # Individual file download
    st.markdown("### 📥 Download Individual File Data")
    try:
        json_data = json.dumps(data, indent=2, ensure_ascii=False, default=str)
        st.download_button(
            label=f"⬇️ Download {data['filename']} JSON",
            data=json_data,
            file_name=f"{os.path.splitext(data['filename'])[0]}_extraction.json",
            mime="application/json",
            key=f"download_{file_path}"
        )
    except Exception as e:
        st.error(f"Could not create download for this file: {str(e)}")

if __name__ == "__main__":
    main()