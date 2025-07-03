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
from typing import Dict, List, Optional, Union
import pdfplumber
from docx import Document
from pptx import Presentation
import openpyxl

class ComprehensiveDocumentExtractor:
    """A class to extract content from various document formats."""
    
    def extract_all(self, file_path: str) -> Dict:
        """Extract all content from a document based on its file type."""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.pdf':
            return self._extract_pdf(file_path)
        elif file_ext == '.docx':
            return self._extract_docx(file_path)
        elif file_ext == '.pptx':
            return self._extract_pptx(file_path)
        elif file_ext in ('.xlsx', '.xls'):
            return self._extract_excel(file_path)
        else:
            raise ValueError(f"Unsupported file type: {file_ext}")
    
    def _extract_pdf(self, file_path: str) -> Dict:
        """Extract content from a PDF file."""
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'PDF',
            'file_size_mb': os.path.getsize(file_path) / (1024 * 1024),
            'pages': [],
            'tables': [],
            'metadata': {},
            'page_count': 0,
            'word_count': 0
        }
        
        try:
            with pdfplumber.open(file_path) as pdf:
                result['page_count'] = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    page_data = {
                        'page_number': i + 1,
                        'text': page.extract_text() or '',
                        'tables': []
                    }
                    
                    # Extract tables
                    tables = page.extract_tables()
                    for table in tables:
                        page_data['tables'].append({'data': table})
                        result['tables'].append({'data': table})
                    
                    # Count words
                    word_count = len(page_data['text'].split())
                    result['word_count'] += word_count
                    
                    result['pages'].append(page_data)
                
                # Extract metadata
                if hasattr(pdf, 'metadata'):
                    result['metadata'] = dict(pdf.metadata)
                
            return result
        except Exception as e:
            raise Exception(f"Error processing PDF: {str(e)}")
    
    def _extract_docx(self, file_path: str) -> Dict:
        """Extract content from a DOCX file."""
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'DOCX',
            'file_size_mb': os.path.getsize(file_path) / (1024 * 1024),
            'paragraphs': [],
            'tables': [],
            'page_count': 1,  # DOCX doesn't have pages in the same way as PDF
            'word_count': 0
        }
        
        try:
            doc = Document(file_path)
            
            # Extract paragraphs
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    result['paragraphs'].append({
                        'text': text,
                        'style': para.style.name
                    })
                    result['word_count'] += len(text.split())
            
            # Extract tables
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text.strip())
                    table_data.append(row_data)
                
                result['tables'].append({'data': table_data})
            
            return result
        except Exception as e:
            raise Exception(f"Error processing DOCX: {str(e)}")
    
    def _extract_pptx(self, file_path: str) -> Dict:
        """Extract content from a PPTX file."""
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'PPTX',
            'file_size_mb': os.path.getsize(file_path) / (1024 * 1024),
            'slides': [],
            'tables': [],
            'slide_count': 0,
            'word_count': 0
        }
        
        try:
            prs = Presentation(file_path)
            result['slide_count'] = len(prs.slides)
            
            for i, slide in enumerate(prs.slides):
                slide_data = {
                    'slide_number': i + 1,
                    'title': '',
                    'content': '',
                    'notes': '',
                    'tables': []
                }
                
                # Extract title and content
                for shape in slide.shapes:
                    if hasattr(shape, 'text'):
                        text = shape.text.strip()
                        if not text:
                            continue
                        
                        if shape == slide.shapes.title:
                            slide_data['title'] = text
                        else:
                            slide_data['content'] += text + '\n'
                
                # Extract notes
                notes_slide = slide.notes_slide
                if notes_slide and notes_slide.notes_text_frame:
                    slide_data['notes'] = notes_slide.notes_text_frame.text.strip()
                
                # Extract tables (simplified)
                for shape in slide.shapes:
                    if shape.has_table:
                        table_data = []
                        for row in shape.table.rows:
                            row_data = []
                            for cell in row.cells:
                                row_data.append(cell.text_frame.text.strip())
                            table_data.append(row_data)
                        
                        slide_data['tables'].append({'data': table_data})
                        result['tables'].append({'data': table_data})
                
                # Count words
                word_count = len(slide_data['content'].split()) + len(slide_data['notes'].split())
                result['word_count'] += word_count
                
                result['slides'].append(slide_data)
            
            return result
        except Exception as e:
            raise Exception(f"Error processing PPTX: {str(e)}")
    
    def _extract_excel(self, file_path: str) -> Dict:
        """Extract content from an Excel file."""
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'Excel',
            'file_size_mb': os.path.getsize(file_path) / (1024 * 1024),
            'sheets': [],
            'sheet_count': 0
        }
        
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            result['sheet_count'] = len(wb.sheetnames)
            
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                sheet_data = {
                    'name': sheet_name,
                    'rows': sheet.max_row,
                    'columns': sheet.max_column,
                    'data': []
                }
                
                # Extract data (limited to first 100 rows for performance)
                for row in sheet.iter_rows(values_only=True):
                    sheet_data['data'].append(list(row))
                
                result['sheets'].append(sheet_data)
            
            return result
        except Exception as e:
            raise Exception(f"Error processing Excel: {str(e)}")

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
    
    .page-selector {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        border: 2px solid #667eea;
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
    
    .page-nav-buttons {
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 1rem;
        margin: 1rem 0;
    }
    
    .page-info {
        background: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        margin: 1rem 0;
        border: 1px solid #2196f3;
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
    
    # Sidebar with features
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
        
        # Add reset button
        add_reset_button()
    
    # Check if we have results to display
    if st.session_state.processing_complete and st.session_state.results:
        # Show content viewer
        show_content_viewer(st.session_state.results)
    else:
        # Show upload interface
        show_upload_interface()

def show_upload_interface():
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
            
            # Process button
            if st.button("🚀 Process Documents", type="primary", use_container_width=True):
                process_documents(uploaded_files)
    
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

def show_content_viewer(results):
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
            key="file_selector",
            help="Choose a file to view its content page by page"
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
                st.metric("Words", f"{file_data.get('word_count', 0):,}")
                st.metric("Size", f"{file_data.get('file_size_mb', 0):.2f} MB")
    
    if selected_file:
        file_data = get_file_data(results, selected_file)
        if file_data:
            # Show content based on file type
            if file_data['file_type'] == 'PDF':
                show_pdf_content(file_data)
            elif file_data['file_type'] == 'DOCX':
                show_docx_content(file_data)
            elif file_data['file_type'] == 'PPTX':
                show_pptx_content(file_data)
            elif file_data['file_type'] == 'Excel':
                show_excel_content(file_data)
    
    # Add download section
    st.markdown("---")
    create_download_section(results, st.session_state.stats)

def get_file_data(results, filename):
    """Get file data by filename"""
    for data in results.values():
        if data['filename'] == filename:
            return data
    return None

def show_pdf_content(file_data):
    """Show PDF content with enhanced page navigation"""
    st.markdown("### 📄 PDF Content")
    
    pages = file_data.get('pages', [])
    if not pages:
        st.warning("No pages found in this PDF file.")
        return
    
    total_pages = len(pages)
    
    # Enhanced page navigation with better layout
    st.markdown("""
    <div class="page-selector">
        <h4>📄 Page Navigation</h4>
    </div>
    """, unsafe_allow_html=True)
    
    # Main navigation controls
    nav_col1, nav_col2, nav_col3 = st.columns([1, 2, 1])
    
    with nav_col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1, use_container_width=True):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with nav_col2:
        # Page selector dropdown
        selected_page = st.selectbox(
            "📄 Go to page:",
            range(1, total_pages + 1),
            index=st.session_state.selected_page - 1,
            key="pdf_page_selector",
            help=f"Select a page to view (1-{total_pages})"
        )
        
        if selected_page != st.session_state.selected_page:
            st.session_state.selected_page = selected_page
            st.rerun()
    
    with nav_col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_pages, use_container_width=True):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Page information display
    st.markdown(f"""
    <div class="page-info">
        <h4>📄 Page {st.session_state.selected_page} of {total_pages}</h4>
        <p>Use the controls above to navigate between pages</p>
    </div>
    """, unsafe_allow_html=True)
    
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
    
    # Show tables
    if current_page.get('tables'):
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
            'Tables': len(current_page.get('tables', []))
        }
        
        for key, value in metadata.items():
            st.write(f"**{key}:** {value}")

def show_docx_content(file_data):
    """Show DOCX content with enhanced paragraph navigation"""
    st.markdown("### 📄 Word Document Content")
    
    paragraphs = file_data.get('paragraphs', [])
    if not paragraphs:
        st.warning("No paragraphs found in this document.")
        return
    
    # Group paragraphs into sections for better navigation
    paragraphs_per_section = 10
    total_sections = (len(paragraphs) + paragraphs_per_section - 1) // paragraphs_per_section
    
    # Enhanced section navigation
    st.markdown("""
    <div class="page-selector">
        <h4>📄 Section Navigation</h4>
    </div>
    """, unsafe_allow_html=True)
    
    nav_col1, nav_col2, nav_col3 = st.columns([1, 2, 1])
    
    with nav_col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1, use_container_width=True):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with nav_col2:
        selected_section = st.selectbox(
            "📄 Go to section:",
            range(1, total_sections + 1),
            index=st.session_state.selected_page - 1,
            key="docx_section_selector",
            help=f"Select a section to view (1-{total_sections})"
        )
        
        if selected_section != st.session_state.selected_page:
            st.session_state.selected_page = selected_section
            st.rerun()
    
    with nav_col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_sections, use_container_width=True):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Show current section content
    start_idx = (st.session_state.selected_page - 1) * paragraphs_per_section
    end_idx = min(start_idx + paragraphs_per_section, len(paragraphs))
    current_paragraphs = paragraphs[start_idx:end_idx]
    
    st.markdown(f"""
    <div class="page-info">
        <h4>📄 Section {st.session_state.selected_page} of {total_sections}</h4>
        <p>Paragraphs {start_idx + 1} - {end_idx}</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="page-content">
        <h3>📄 Section {st.session_state.selected_page} Content</h3>
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
    
    # Show tables
    if file_data.get('tables'):
        st.markdown("#### 📊 Document Tables")
        for i, table in enumerate(file_data['tables']):
            st.markdown(f"**Table {i+1}**")
            try:
                if table.get('data'):
                    df = pd.DataFrame(table['data'])
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Error displaying table: {str(e)}")

def show_pptx_content(file_data):
    """Show PowerPoint content with enhanced slide navigation"""
    st.markdown("### 🎞️ PowerPoint Content")
    
    slides = file_data.get('slides', [])
    if not slides:
        st.warning("No slides found in this presentation.")
        return
    
    total_slides = len(slides)
    
    # Enhanced slide navigation
    st.markdown("""
    <div class="page-selector">
        <h4>🎞️ Slide Navigation</h4>
    </div>
    """, unsafe_allow_html=True)
    
    nav_col1, nav_col2, nav_col3 = st.columns([1, 2, 1])
    
    with nav_col1:
        if st.button("⬅️ Previous", disabled=st.session_state.selected_page <= 1, use_container_width=True):
            st.session_state.selected_page -= 1
            st.rerun()
    
    with nav_col2:
        selected_slide = st.selectbox(
            "🎞️ Go to slide:",
            range(1, total_slides + 1),
            index=st.session_state.selected_page - 1,
            key="pptx_slide_selector",
            help=f"Select a slide to view (1-{total_slides})"
        )
        
        if selected_slide != st.session_state.selected_page:
            st.session_state.selected_page = selected_slide
            st.rerun()
    
    with nav_col3:
        if st.button("Next ➡️", disabled=st.session_state.selected_page >= total_slides, use_container_width=True):
            st.session_state.selected_page += 1
            st.rerun()
    
    # Show current slide content
    current_slide = slides[st.session_state.selected_page - 1]
    
    st.markdown(f"""
    <div class="page-info">
        <h4>🎞️ Slide {st.session_state.selected_page} of {total_slides}</h4>
    </div>
    """, unsafe_allow_html=True)
    
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
    
    # Show tables
    if current_slide.get('tables'):
        st.markdown("#### 📊 Tables")
        for i, table in enumerate(current_slide['tables']):
            st.markdown(f"**Table {i+1}**")
            try:
                if table.get('data'):
                    df = pd.DataFrame(table['data'])
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Error displaying table: {str(e)}")

def show_excel_content(file_data):
    """Show Excel content with enhanced sheet navigation"""
    st.markdown("### 📊 Excel Content")
    
    sheets = file_data.get('sheets', [])
    if not sheets:
        st.warning("No sheets found in this Excel file.")
        return
    
    # Enhanced sheet selector
    st.markdown("""
    <div class="page-selector">
        <h4>📊 Sheet Navigation</h4>
    </div>
    """, unsafe_allow_html=True)
    
    sheet_names = [sheet['name'] for sheet in sheets]
    selected_sheet_name = st.selectbox(
        "📋 Select a sheet:",
        sheet_names,
        key="excel_sheet_selector",
        help="Choose a sheet to view its content"
    )
    
    # Find selected sheet
    selected_sheet = None
    for sheet in sheets:
        if sheet['name'] == selected_sheet_name:
            selected_sheet = sheet
            break
    
    if selected_sheet:
        st.markdown(f"""
        <div class="page-info">
            <h4>📊 Sheet: {selected_sheet['name']}</h4>
        </div>
        """, unsafe_allow_html=True)
        
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
        if selected_sheet.get('data'):
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

def process_documents(uploaded_files):
    """Process uploaded documents"""
    # Create a temporary directory to store uploaded files
    with tempfile.TemporaryDirectory() as temp_dir:
        results = {}
        stats = {
            'total_files': len(uploaded_files),
            'processed_files': 0,
            'total_pages': 0,
            'total_words': 0,
            'total_tables': 0,
            'start_time': datetime.now(),
            'file_types': {}
        }
        
        # Initialize progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Process each file
        for i, uploaded_file in enumerate(uploaded_files):
            try:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Update status
                progress = (i + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Processing {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
                
                # Process the file
                extractor = ComprehensiveDocumentExtractor()
                file_data = extractor.extract_all(file_path)
                
                # Update statistics
                stats['processed_files'] += 1
                stats['total_pages'] += file_data.get('page_count', 1)
                stats['total_words'] += file_data.get('word_count', 0)
                stats['total_tables'] += len(file_data.get('tables', []))
                
                # Track file types
                file_type = file_data.get('file_type', 'Unknown')
                if file_type not in stats['file_types']:
                    stats['file_types'][file_type] = 0
                stats['file_types'][file_type] += 1
                
                # Store results
                results[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {str(e)}")
                continue
        
        # Finalize stats
        stats['end_time'] = datetime.now()
        stats['processing_time'] = (stats['end_time'] - stats['start_time']).total_seconds()
        
        # Update session state
        st.session_state.processing_complete = True
        st.session_state.results = results
        st.session_state.stats = stats
        st.session_state.selected_file = list(results.keys())[0] if results else None
        
        # Show completion message
        progress_bar.empty()
        status_text.empty()
        st.success("✅ Processing complete!")
        
        # Show summary stats
        show_processing_stats(stats)

def show_processing_stats(stats):
    """Show processing statistics"""
    st.markdown("## 📊 Processing Statistics")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Files", stats['total_files'])
        st.metric("Processed Files", stats['processed_files'])
    
    with col2:
        st.metric("Total Pages", stats['total_pages'])
        st.metric("Total Words", f"{stats['total_words']:,}")
    
    with col3:
        st.metric("Total Tables", stats['total_tables'])
        st.metric("Processing Time", f"{stats['processing_time']:.2f} seconds")
    
    # File type distribution
    st.markdown("### 📂 File Type Distribution")
    if stats['file_types']:
        df_types = pd.DataFrame.from_dict(stats['file_types'], orient='index', columns=['Count'])
        st.bar_chart(df_types)
    else:
        st.info("No file type information available")

def create_download_section(results, stats):
    """Create download section for results"""
    st.markdown("## ⬇️ Download Results")
    
    # Create a temporary directory for the zip file
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, "extraction_results.zip")
        
        # Create zip file with all results
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            # Add JSON files for each document
            for filename, data in results.items():
                json_filename = f"{Path(filename).stem}_results.json"
                json_path = os.path.join(temp_dir, json_filename)
                
                with open(json_path, 'w') as f:
                    json.dump(data, f, indent=2)
                
                zipf.write(json_path, json_filename)
            
            # Add summary stats
            stats_filename = "processing_summary.json"
            stats_path = os.path.join(temp_dir, stats_filename)
            
            with open(stats_path, 'w') as f:
                json.dump(stats, f, indent=2)
            
            zipf.write(stats_path, stats_filename)
        
        # Create download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            # Download all results as ZIP
            with open(zip_path, 'rb') as f:
                zip_data = f.read()
            
            st.download_button(
                label="📦 Download All Results (ZIP)",
                data=zip_data,
                file_name="document_extraction_results.zip",
                mime="application/zip",
                help="Download all extracted content as a ZIP file containing JSON files"
            )
        
        with col2:
            # Download summary report
            summary_report = generate_summary_report(results, stats)
            
            st.download_button(
                label="📄 Download Summary Report (TXT)",
                data=summary_report,
                file_name="extraction_summary.txt",
                mime="text/plain",
                help="Download a text summary of the extraction results"
            )
        
        # Individual file downloads
        st.markdown("### 📂 Download Individual Files")
        selected_file = st.selectbox(
            "Select a file to download:",
            list(results.keys()),
            key="download_file_selector"
        )
        
        if selected_file:
            file_data = results[selected_file]
            json_data = json.dumps(file_data, indent=2).encode('utf-8')
            
            st.download_button(
                label=f"⬇️ Download {selected_file} Results (JSON)",
                data=json_data,
                file_name=f"{Path(selected_file).stem}_results.json",
                mime="application/json",
                key=f"download_{selected_file}"
            )

def generate_summary_report(results, stats):
    """Generate a text summary report of the extraction"""
    report = io.StringIO()
    
    report.write("=== DOCUMENT EXTRACTION SUMMARY REPORT ===\n\n")
    report.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    report.write(f"Total processing time: {stats.get('processing_time', 0):.2f} seconds\n")
    report.write(f"Files processed: {stats.get('processed_files', 0)}/{stats.get('total_files', 0)}\n\n")
    
    # File type breakdown
    report.write("=== FILE TYPE BREAKDOWN ===\n")
    for file_type, count in stats.get('file_types', {}).items():
        report.write(f"{file_type}: {count} files\n")
    
    # Overall statistics
    report.write("\n=== OVERALL STATISTICS ===\n")
    report.write(f"Total pages extracted: {stats.get('total_pages', 0)}\n")
    report.write(f"Total words extracted: {stats.get('total_words', 0)}\n")
    report.write(f"Total tables extracted: {stats.get('total_tables', 0)}\n\n")
    
    # Per-file details
    report.write("=== FILE DETAILS ===\n")
    for filename, data in results.items():
        report.write(f"\nFile: {filename}\n")
        report.write(f"Type: {data.get('file_type', 'Unknown')}\n")
        report.write(f"Size: {data.get('file_size_mb', 0):.2f} MB\n")
        
        if data['file_type'] == 'PDF':
            report.write(f"Pages: {data.get('page_count', 0)}\n")
            report.write(f"Words: {data.get('word_count', 0)}\n")
            report.write(f"Tables: {len(data.get('tables', []))}\n")
        elif data['file_type'] == 'DOCX':
            report.write(f"Paragraphs: {len(data.get('paragraphs', []))}\n")
            report.write(f"Words: {data.get('word_count', 0)}\n")
            report.write(f"Tables: {len(data.get('tables', []))}\n")
        elif data['file_type'] == 'PPTX':
            report.write(f"Slides: {data.get('slide_count', 0)}\n")
            report.write(f"Words: {data.get('word_count', 0)}\n")
            report.write(f"Tables: {len(data.get('tables', []))}\n")
        elif data['file_type'] == 'Excel':
            report.write(f"Sheets: {data.get('sheet_count', 0)}\n")
    
    return report.getvalue()

def add_reset_button():
    """Add a reset button to clear session state"""
    if st.sidebar.button("🔄 Reset Application", help="Clear all processed data and start fresh"):
        st.session_state.processing_complete = False
        st.session_state.results = None
        st.session_state.stats = None
        st.session_state.selected_file = None
        st.session_state.selected_page = 1
        st.rerun()

if __name__ == "__main__":
    main()
