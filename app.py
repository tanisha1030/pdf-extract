import streamlit as st
import tempfile
import os
import json
import zipfile
from pathlib import Path
from pdfextract import ComprehensiveDocumentExtractor
from datetime import datetime

def main():
    """Main Streamlit application"""
    st.set_page_config(page_title="Document Extractor", layout="wide")
    st.title("📄 Comprehensive Document Extractor")
    
    # Initialize session state
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
        st.session_state.results = None
        st.session_state.stats = None
    
    # File upload section
    st.markdown("## 📤 Upload Files")
    uploaded_files = st.file_uploader(
        "Upload documents to process (PDF, DOCX, PPTX, XLSX)",
        type=['pdf', 'docx', 'pptx', 'xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files and not st.session_state.processing_complete:
        # Save uploaded files to temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            file_paths = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                file_paths.append(file_path)
            
            # Process files
            if st.button("🚀 Process Files", use_container_width=True):
                with st.spinner("Processing files..."):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Initialize extractor
                    extractor = ComprehensiveDocumentExtractor()
                    
                    # Record start time
                    start_time = datetime.now()
                    
                    # Process files
                    results = extractor.process_files(file_paths)
                    
                    # Record processing time
                    processing_time = (datetime.now() - start_time).total_seconds()
                    
                    if results:
                        # Generate statistics
                        stats = {
                            'total_files': len(results),
                            'processed_files': len(results),
                            'processing_time': processing_time,
                            'start_time': start_time,
                            'end_time': datetime.now(),
                            'total_pages': sum(data.get('page_count', 1) for data in results.values()),
                            'total_words': sum(data.get('total_words', 0) for data in results.values()),
                            'total_tables': sum(len(data.get('extracted_tables', [])) for data in results.values()),
                            'total_images': sum(data.get('total_images', 0) for data in results.values()),
                            'file_types': Counter(data['file_type'] for data in results.values())
                        }
                        
                        # Update session state
                        st.session_state.processing_complete = True
                        st.session_state.results = results
                        st.session_state.stats = stats
                        
                        # Show completion message
                        progress_bar.empty()
                        status_text.empty()
                        st.success("✅ All files processed successfully!")
                        st.rerun()
                    else:
                        st.error("❌ No files were successfully processed.")

    # Show results if processing is complete
    if st.session_state.processing_complete:
        st.markdown("## 📊 Processing Results")
        
        # Display statistics
        st.markdown("### 📈 Summary Statistics")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Files", st.session_state.stats['total_files'])
            st.metric("Total Pages", st.session_state.stats['total_pages'])
        with col2:
            st.metric("Total Words", f"{st.session_state.stats['total_words']:,}")
            st.metric("Total Tables", st.session_state.stats['total_tables'])
        with col3:
            st.metric("Total Images", st.session_state.stats['total_images'])
            st.metric("Processing Time", f"{st.session_state.stats['processing_time']:.2f} sec")
        
        # File type distribution
        st.markdown("### 📁 File Type Distribution")
        file_types = st.session_state.stats['file_types']
        st.bar_chart(file_types)
        
        # Create download section
        create_download_section(st.session_state.results, st.session_state.stats)
        
        # Add reset button
        add_reset_button()

def create_download_section(results, stats):
    """Create the download section for processed files"""
    st.markdown("## 📥 Download Results")
    
    # Create a temporary directory for the zip file
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create individual JSON files for each document
        for filename, data in results.items():
            json_path = os.path.join(temp_dir, f"{Path(filename).stem}.json")
            with open(json_path, 'w') as f:
                json.dump(data, f, indent=2)
        
        # Create a summary report
        summary_report = {
            "processing_summary": {
                "total_files": stats['total_files'],
                "processed_files": stats['processed_files'],
                "processing_time_seconds": stats['processing_time'],
                "start_time": stats['start_time'].isoformat(),
                "end_time": stats['end_time'].isoformat()
            },
            "content_summary": {
                "total_pages": stats['total_pages'],
                "total_words": stats['total_words'],
                "total_tables": stats['total_tables'],
                "total_images": stats['total_images'],
                "file_types": dict(stats['file_types'])
            }
        }
        
        # Save summary report
        summary_path = os.path.join(temp_dir, "summary_report.json")
        with open(summary_path, 'w') as f:
            json.dump(summary_report, f, indent=2)
        
        # Create a zip file
        zip_path = os.path.join(temp_dir, "document_extraction_results.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    if file != "document_extraction_results.zip":
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, temp_dir))
        
        # Create download button
        with open(zip_path, 'rb') as f:
            zip_data = f.read()
        
        st.download_button(
            label="📦 Download All Results (ZIP)",
            data=zip_data,
            file_name="document_extraction_results.zip",
            mime="application/zip",
            use_container_width=True
        )
    
    # Add option to download individual files
    st.markdown("### 📄 Download Individual Files")
    for filename in results.keys():
        file_data = results[filename]
        json_str = json.dumps(file_data, indent=2)
        
        st.download_button(
            label=f"⬇️ Download {Path(filename).name} (JSON)",
            data=json_str,
            file_name=f"{Path(filename).stem}.json",
            mime="application/json",
            key=f"dl_{filename}",
            use_container_width=True
        )

def add_reset_button():
    """Add a reset button to clear session state"""
    if st.button("🔄 Reset Application", use_container_width=True):
        st.session_state.processing_complete = False
        st.session_state.results = None
        st.session_state.stats = None
        st.rerun()

if __name__ == "__main__":
    from collections import Counter  # Import needed for file_types counter
    main()
