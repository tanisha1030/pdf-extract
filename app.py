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
from pdfextract import ComprehensiveDocumentExtractor

# Page configuration
st.set_page_config(
    page_title="Document Content Viewer",
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
    
    .page-viewer {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    
    .text-content {
        white-space: pre-wrap;
        font-family: monospace;
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e1e4e8;
        max-height: 600px;
        overflow-y: auto;
    }
    
    .image-container {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem;
        margin: 1rem 0;
    }
    
    .image-card {
        border: 1px solid #e1e4e8;
        border-radius: 8px;
        padding: 0.5rem;
        background: white;
        max-width: 300px;
    }
    
    .nav-buttons {
        display: flex;
        gap: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'results' not in st.session_state:
    st.session_state.results = None
if 'current_file' not in st.session_state:
    st.session_state.current_file = None
if 'current_page' not in st.session_state:
    st.session_state.current_page = 1

def main():
    # Main header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Document Content Viewer</h1>
        <p>View extracted content from PDF, DOCX, PPTX, and Excel files with page navigation</p>
    </div>
    """, unsafe_allow_html=True)
    
    # File upload section
    uploaded_files = st.file_uploader(
        "Upload documents to extract content",
        type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files and not st.session_state.processing_complete:
        process_documents(uploaded_files)
    
    # Display content if processing is complete
    if st.session_state.processing_complete and st.session_state.results:
        display_document_content()

def process_documents(uploaded_files):
    """Process uploaded documents and extract content"""
    with st.spinner("Processing documents..."):
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
                
                # Process files
                results = {}
                for file_path in file_paths:
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
                            
                    except Exception as e:
                        st.error(f"Error processing {os.path.basename(file_path)}: {str(e)}")
                        continue
                
                # Save results to session state
                st.session_state.results = results
                st.session_state.processing_complete = True
                st.session_state.current_file = next(iter(results.keys())) if results else None
                st.session_state.current_page = 1
                
                st.success("Document processing complete!")
                st.rerun()
                
            except Exception as e:
                st.error(f"Error processing documents: {str(e)}")

def display_document_content():
    """Display the extracted document content with navigation"""
    results = st.session_state.results
    
    # File selector in sidebar
    with st.sidebar:
        st.markdown("## 📂 Document Navigation")
        
        # File selection
        file_options = [data['filename'] for data in results.values()]
        selected_file_name = st.selectbox(
            "Select a document",
            options=file_options,
            index=file_options.index(results[st.session_state.current_file]['filename']) if st.session_state.current_file else 0
        )
        
        # Find the selected file path
        selected_file_path = None
        for file_path, data in results.items():
            if data['filename'] == selected_file_name:
                selected_file_path = file_path
                break
        
        if selected_file_path != st.session_state.current_file:
            st.session_state.current_file = selected_file_path
            st.session_state.current_page = 1
            st.rerun()
        
        # Page navigation for PDF and PPTX
        file_data = results[st.session_state.current_file]
        if file_data['file_type'] in ['PDF', 'PPTX']:
            total_pages = file_data.get('page_count', file_data.get('total_slides', 1))
            
            st.markdown(f"### 📄 Page Navigation ({total_pages} total)")
            
            # Page number input
            page_num = st.number_input(
                "Go to page",
                min_value=1,
                max_value=total_pages,
                value=st.session_state.current_page,
                step=1
            )
            
            if page_num != st.session_state.current_page:
                st.session_state.current_page = page_num
                st.rerun()
            
            # Navigation buttons
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button("⏮️ First"):
                    st.session_state.current_page = 1
                    st.rerun()
            with col2:
                if st.button("⬅️ Previous") and st.session_state.current_page > 1:
                    st.session_state.current_page -= 1
                    st.rerun()
            with col3:
                if st.button("➡️ Next") and st.session_state.current_page < total_pages:
                    st.session_state.current_page += 1
                    st.rerun()
            with col4:
                if st.button("⏭️ Last"):
                    st.session_state.current_page = total_pages
                    st.rerun()
    
    # Main content area
    st.markdown(f"## 📄 {file_data['filename']}")
    
    # Show file info
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("File Type", file_data['file_type'])
    with col2:
        st.metric("File Size", f"{file_data.get('file_size_mb', 0):.2f} MB")
    with col3:
        if file_data['file_type'] == 'PDF':
            st.metric("Total Pages", file_data.get('page_count', 0))
        elif file_data['file_type'] == 'PPTX':
            st.metric("Total Slides", file_data.get('total_slides', 0))
        elif file_data['file_type'] == 'DOCX':
            st.metric("Paragraphs", len(file_data.get('paragraphs', [])))
        elif file_data['file_type'] == 'Excel':
            st.metric("Sheets", len(file_data.get('sheets', [])))
    
    # Display content based on file type
    if file_data['file_type'] == 'PDF':
        display_pdf_content(file_data)
    elif file_data['file_type'] == 'DOCX':
        display_docx_content(file_data)
    elif file_data['file_type'] == 'PPTX':
        display_pptx_content(file_data)
    elif file_data['file_type'] in ['XLSX', 'XLS']:
        display_excel_content(file_data)

def display_pdf_content(file_data):
    """Display content from a PDF file"""
    current_page = st.session_state.current_page
    
    # Check if we have data for the current page
    if len(file_data['pages']) >= current_page:
        page_data = file_data['pages'][current_page - 1]
        
        # Display page info
        st.markdown(f"### Page {current_page}")
        st.write(f"Size: {page_data['page_size']['width']} x {page_data['page_size']['height']} pts")
        st.write(f"Rotation: {page_data['rotation']}°")
        st.write(f"Word count: {page_data['word_count']}")
        
        # Display text content
        st.markdown("#### Text Content")
        st.markdown(f'<div class="text-content">{page_data["text"]}</div>', unsafe_allow_html=True)
        
        # Display images if any
        if page_data['images']:
            st.markdown("#### Extracted Images")
            st.markdown('<div class="image-container">', unsafe_allow_html=True)
            
            for img in page_data['images']:
                if os.path.exists(img['filename']):
                    st.markdown(f"""
                    <div class="image-card">
                        <img src="{img['filename']}" style="max-width: 100%; height: auto;">
                        <p>Size: {img['width']}×{img['height']}</p>
                        <p>File size: {img['size_bytes']} bytes</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Display tables if any
        if page_data['tables']:
            st.markdown("#### Detected Tables")
            for i, table in enumerate(page_data['tables']):
                st.write(f"Table {i+1} (Confidence: {table.get('confidence', 'unknown')})")
                st.dataframe(pd.DataFrame(table['data']))
    
    else:
        st.warning(f"No data available for page {current_page}")

def display_docx_content(file_data):
    """Display content from a DOCX file"""
    st.markdown("### Document Content")
    
    # Display paragraphs
    if file_data.get('paragraphs'):
        paragraphs_to_show = file_data['paragraphs'][:20]  # Limit to first 20 paragraphs
        full_text = "\n\n".join([p['text'] for p in paragraphs_to_show])
        st.markdown(f'<div class="text-content">{full_text}</div>', unsafe_allow_html=True)
        
        if len(file_data['paragraphs']) > 20:
            st.info(f"Showing first 20 of {len(file_data['paragraphs'])} paragraphs")
    else:
        st.warning("No text content found in this document")
    
    # Display tables if any
    if file_data.get('tables'):
        st.markdown("### Tables")
        for i, table in enumerate(file_data['tables']):
            st.write(f"Table {i+1}")
            st.dataframe(pd.DataFrame(table['data']))

def display_pptx_content(file_data):
    """Display content from a PPTX file"""
    current_slide = st.session_state.current_page
    
    # Check if we have data for the current slide
    if len(file_data['slides']) >= current_slide:
        slide_data = file_data['slides'][current_slide - 1]
        
        # Display slide info
        st.markdown(f"### Slide {current_slide}")
        
        # Display text content
        if slide_data['text_content']:
            st.markdown("#### Text Content")
            full_text = "\n\n".join(slide_data['text_content'])
            st.markdown(f'<div class="text-content">{full_text}</div>', unsafe_allow_html=True)
        else:
            st.info("No text content on this slide")
        
        # Display tables if any
        if slide_data['tables']:
            st.markdown("#### Tables")
            for i, table in enumerate(slide_data['tables']):
                st.write(f"Table {i+1}")
                st.dataframe(pd.DataFrame(table))
    else:
        st.warning(f"No data available for slide {current_slide}")

def display_excel_content(file_data):
    """Display content from an Excel file"""
    st.markdown("### Sheets")
    
    # Create tabs for each sheet
    if file_data.get('sheets'):
        tabs = st.tabs([sheet['sheet_name'] for sheet in file_data['sheets']])
        
        for i, tab in enumerate(tabs):
            with tab:
                sheet_data = file_data['sheets'][i]
                
                st.write(f"Shape: {sheet_data['shape'][0]} rows × {sheet_data['shape'][1]} columns")
                
                if sheet_data['data']:
                    st.dataframe(pd.DataFrame(sheet_data['data']))
                else:
                    st.warning("No data in this sheet")
    else:
        st.warning("No sheets found in this Excel file")

if __name__ == "__main__":
    main()
