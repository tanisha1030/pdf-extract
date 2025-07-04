import os
import json
import pandas as pd
import time
import gc
import sys
from pathlib import Path

# Optional libraries with better error handling
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False
    print("⚠️  PyMuPDF not available - PDF processing disabled")

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("⚠️  python-docx not available - DOCX processing disabled")

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False
    print("⚠️  python-pptx not available - PPTX processing disabled")

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("⚠️  openpyxl not available - Excel processing may be limited")


class FixedDocumentExtractor:
    def __init__(self, timeout_seconds=300):
        self.timeout_seconds = timeout_seconds
        self.supported_formats = ['.pdf', '.docx', '.pptx', '.xlsx', '.xls']
        
        # Check dependencies
        if not any([HAS_FITZ, HAS_DOCX, HAS_PPTX]):
            print("❌ No document processing libraries available!")
            print("Install with: pip install PyMuPDF python-docx python-pptx openpyxl")
            sys.exit(1)
    
    def extract_pdf_safe(self, pdf_path):
        """Safe PDF extraction with timeout and error handling"""
        if not HAS_FITZ:
            print("❌ PyMuPDF not available for PDF processing")
            return None
            
        try:
            print(f"📄 Processing PDF: {os.path.basename(pdf_path)}")
            start_time = time.time()
            
            # Check file size first
            file_size_mb = os.path.getsize(pdf_path) / (1024*1024)
            if file_size_mb > 100:
                print(f"⚠️  Large file ({file_size_mb:.2f} MB) - processing may take longer")
            
            # Open PDF with error handling
            try:
                doc = fitz.open(pdf_path)
            except Exception as e:
                print(f"❌ Cannot open PDF: {str(e)}")
                return None
            
            page_count = len(doc)
            print(f"  📊 File: {file_size_mb:.2f} MB, Pages: {page_count}")
            
            # Limit pages for large documents
            max_pages = min(page_count, 100)
            if max_pages < page_count:
                print(f"  ⚠️  Processing limited to first {max_pages} pages")
            
            pdf_data = {
                'filename': os.path.basename(pdf_path),
                'file_type': 'PDF',
                'pages': [],
                'total_words': 0,
                'total_characters': 0,
                'page_count': page_count,
                'processed_pages': max_pages,
                'file_size_mb': round(file_size_mb, 2)
            }
            
            # Process pages with timeout check
            for page_num in range(max_pages):
                # Check timeout
                if time.time() - start_time > self.timeout_seconds:
                    print(f"⏱️  Timeout reached, stopping at page {page_num}")
                    break
                
                try:
                    page = doc[page_num]
                    
                    # Extract text with timeout protection
                    try:
                        page_text = page.get_text()
                        if not page_text:
                            page_text = ""
                    except Exception as e:
                        print(f"⚠️  Text extraction failed on page {page_num + 1}: {str(e)}")
                        page_text = ""
                    
                    # Count words and characters
                    word_count = len(page_text.split()) if page_text else 0
                    char_count = len(page_text) if page_text else 0
                    
                    # Basic image count (without extraction)
                    image_count = 0
                    try:
                        image_list = page.get_images()
                        image_count = len(image_list) if image_list else 0
                    except:
                        image_count = 0
                    
                    page_data = {
                        'page_number': page_num + 1,
                        'text': page_text[:1000],  # Limit text stored
                        'full_text_length': len(page_text),
                        'word_count': word_count,
                        'char_count': char_count,
                        'image_count': image_count
                    }
                    
                    pdf_data['pages'].append(page_data)
                    pdf_data['total_words'] += word_count
                    pdf_data['total_characters'] += char_count
                    
                    # Progress update
                    if page_num > 0 and page_num % 10 == 0:
                        elapsed = time.time() - start_time
                        print(f"  📄 Progress: {page_num + 1}/{max_pages} pages ({elapsed:.1f}s)")
                    
                except Exception as e:
                    print(f"⚠️  Error on page {page_num + 1}: {str(e)}")
                    continue
            
            # Clean up
            try:
                doc.close()
            except:
                pass
            
            elapsed = time.time() - start_time
            print(f"  ✅ PDF processed in {elapsed:.2f}s")
            
            return pdf_data
            
        except Exception as e:
            print(f"❌ PDF processing error: {str(e)}")
            return None
    
    def extract_docx_safe(self, docx_path):
        """Safe DOCX extraction"""
        if not HAS_DOCX:
            print("❌ python-docx not available for DOCX processing")
            return None
            
        try:
            print(f"📄 Processing DOCX: {os.path.basename(docx_path)}")
            start_time = time.time()
            
            # Check file size
            file_size_mb = os.path.getsize(docx_path) / (1024*1024)
            
            try:
                doc = Document(docx_path)
            except Exception as e:
                print(f"❌ Cannot open DOCX: {str(e)}")
                return None
            
            docx_data = {
                'filename': os.path.basename(docx_path),
                'file_type': 'DOCX',
                'paragraphs': [],
                'total_words': 0,
                'total_characters': 0,
                'file_size_mb': round(file_size_mb, 2)
            }
            
            # Process paragraphs with limits
            paragraph_count = 0
            max_paragraphs = 500  # Limit paragraphs
            
            for para in doc.paragraphs:
                if paragraph_count >= max_paragraphs:
                    break
                
                if para.text.strip():
                    text = para.text.strip()
                    
                    # Limit individual paragraph length
                    display_text = text[:500] if len(text) > 500 else text
                    
                    docx_data['paragraphs'].append({
                        'text': display_text,
                        'full_length': len(text),
                        'word_count': len(text.split())
                    })
                    
                    docx_data['total_words'] += len(text.split())
                    docx_data['total_characters'] += len(text)
                    paragraph_count += 1
            
            elapsed = time.time() - start_time
            print(f"  ✅ DOCX processed in {elapsed:.2f}s: {docx_data['total_words']} words")
            
            return docx_data
            
        except Exception as e:
            print(f"❌ DOCX processing error: {str(e)}")
            return None
    
    def extract_pptx_safe(self, pptx_path):
        """Safe PPTX extraction"""
        if not HAS_PPTX:
            print("❌ python-pptx not available for PPTX processing")
            return None
            
        try:
            print(f"📄 Processing PPTX: {os.path.basename(pptx_path)}")
            start_time = time.time()
            
            file_size_mb = os.path.getsize(pptx_path) / (1024*1024)
            
            try:
                prs = Presentation(pptx_path)
            except Exception as e:
                print(f"❌ Cannot open PPTX: {str(e)}")
                return None
            
            pptx_data = {
                'filename': os.path.basename(pptx_path),
                'file_type': 'PPTX',
                'slides': [],
                'total_slides': len(prs.slides),
                'total_words': 0,
                'total_characters': 0,
                'file_size_mb': round(file_size_mb, 2)
            }
            
            # Process slides with limits
            max_slides = min(50, len(prs.slides))
            
            for slide_idx, slide in enumerate(prs.slides[:max_slides]):
                slide_data = {
                    'slide_number': slide_idx + 1,
                    'text_content': [],
                    'total_text': ''
                }
                
                # Extract text from shapes
                slide_text = ''
                text_count = 0
                
                for shape in slide.shapes:
                    if text_count >= 10:  # Limit text shapes per slide
                        break
                    
                    try:
                        if hasattr(shape, "text") and shape.text.strip():
                            text = shape.text.strip()
                            slide_data['text_content'].append(text[:200])  # Limit text
                            slide_text += text + ' '
                            text_count += 1
                    except:
                        continue
                
                slide_data['total_text'] = slide_text
                pptx_data['slides'].append(slide_data)
                pptx_data['total_words'] += len(slide_text.split())
                pptx_data['total_characters'] += len(slide_text)
            
            elapsed = time.time() - start_time
            print(f"  ✅ PPTX processed in {elapsed:.2f}s: {pptx_data['total_words']} words")
            
            return pptx_data
            
        except Exception as e:
            print(f"❌ PPTX processing error: {str(e)}")
            return None
    
    def extract_excel_safe(self, excel_path):
        """Safe Excel extraction"""
        try:
            print(f"📄 Processing Excel: {os.path.basename(excel_path)}")
            start_time = time.time()
            
            file_size_mb = os.path.getsize(excel_path) / (1024*1024)
            
            excel_data = {
                'filename': os.path.basename(excel_path),
                'file_type': 'Excel',
                'sheets': [],
                'file_size_mb': round(file_size_mb, 2)
            }
            
            # Try to read Excel file
            try:
                # Read with pandas, limiting rows
                max_rows = 50 if file_size_mb > 5 else 100
                
                # Get sheet names first
                xl_file = pd.ExcelFile(excel_path)
                sheet_names = xl_file.sheet_names[:5]  # Limit sheets
                
                for sheet_name in sheet_names:
                    try:
                        # Read limited data
                        df = pd.read_excel(
                            excel_path, 
                            sheet_name=sheet_name, 
                            nrows=max_rows
                        )
                        
                        if not df.empty:
                            # Limit columns
                            if df.shape[1] > 10:
                                df = df.iloc[:, :10]
                            
                            sheet_data = {
                                'sheet_name': sheet_name,
                                'shape': df.shape,
                                'columns': df.columns.tolist(),
                                'sample_data': df.head(5).to_dict('records')
                            }
                            
                            excel_data['sheets'].append(sheet_data)
                        
                    except Exception as e:
                        print(f"    ⚠️  Error reading sheet '{sheet_name}': {str(e)}")
                        continue
                
                xl_file.close()
                
            except Exception as e:
                print(f"❌ Cannot read Excel file: {str(e)}")
                return None
            
            elapsed = time.time() - start_time
            print(f"  ✅ Excel processed in {elapsed:.2f}s: {len(excel_data['sheets'])} sheets")
            
            return excel_data
            
        except Exception as e:
            print(f"❌ Excel processing error: {str(e)}")
            return None
    
    def process_file(self, file_path):
        """Process a single file with error handling"""
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            return None
        
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext not in self.supported_formats:
            print(f"❌ Unsupported format: {file_ext}")
            return None
        
        print(f"\n🔄 Processing: {os.path.basename(file_path)}")
        
        try:
            if file_ext == '.pdf':
                return self.extract_pdf_safe(file_path)
            elif file_ext == '.docx':
                return self.extract_docx_safe(file_path)
            elif file_ext == '.pptx':
                return self.extract_pptx_safe(file_path)
            elif file_ext in ['.xlsx', '.xls']:
                return self.extract_excel_safe(file_path)
            else:
                print(f"❌ Unsupported file type: {file_ext}")
                return None
                
        except Exception as e:
            print(f"❌ Processing error: {str(e)}")
            return None
    
    def save_results(self, results, output_file='extraction_results.json'):
        """Save results to JSON file"""
        if not results:
            print("❌ No results to save")
            return
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False, default=str)
            print(f"✅ Results saved to: {output_file}")
            
            # Create summary
            self.create_summary(results)
            
        except Exception as e:
            print(f"❌ Error saving results: {str(e)}")
    
    def create_summary(self, results):
        """Create a summary report"""
        try:
            summary_file = 'extraction_summary.txt'
            
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write("📋 DOCUMENT EXTRACTION SUMMARY\n")
                f.write("=" * 50 + "\n\n")
                
                f.write(f"📁 Files Processed: {len(results)}\n")
                
                total_words = 0
                total_chars = 0
                total_size = 0
                
                for file_path, data in results.items():
                    if data:
                        total_words += data.get('total_words', 0)
                        total_chars += data.get('total_characters', 0)
                        total_size += data.get('file_size_mb', 0)
                
                f.write(f"📝 Total Words: {total_words:,}\n")
                f.write(f"🔤 Total Characters: {total_chars:,}\n")
                f.write(f"💾 Total Size: {total_size:.2f} MB\n\n")
                
                f.write("📄 FILE DETAILS:\n")
                f.write("-" * 30 + "\n")
                
                for file_path, data in results.items():
                    if data:
                        f.write(f"• {data['filename']} ({data['file_type']})\n")
                        f.write(f"  Words: {data.get('total_words', 0):,}\n")
                        f.write(f"  Size: {data.get('file_size_mb', 0):.2f} MB\n")
                        
                        if data['file_type'] == 'PDF':
                            f.write(f"  Pages: {data.get('page_count', 0)}\n")
                        elif data['file_type'] == 'PPTX':
                            f.write(f"  Slides: {data.get('total_slides', 0)}\n")
                        elif data['file_type'] == 'Excel':
                            f.write(f"  Sheets: {len(data.get('sheets', []))}\n")
                        
                        f.write("\n")
            
            print(f"✅ Summary saved to: {summary_file}")
            
        except Exception as e:
            print(f"❌ Error creating summary: {str(e)}")


def main():
    """Main function with improved user interface"""
    print("🚀 FIXED DOCUMENT EXTRACTOR")
    print("=" * 60)
    print("Supported: PDF, DOCX, PPTX, XLSX, XLS")
    print("⚡ Optimized for reliability and speed")
    print()
    
    # Check if we have any processing capabilities
    available_formats = []
    if HAS_FITZ:
        available_formats.append("PDF")
    if HAS_DOCX:
        available_formats.append("DOCX")
    if HAS_PPTX:
        available_formats.append("PPTX")
    available_formats.append("Excel")  # pandas handles this
    
    print(f"✅ Available formats: {', '.join(available_formats)}")
    print()
    
    # Initialize extractor
    extractor = FixedDocumentExtractor(timeout_seconds=300)
    
    # Get files to process
    print("📁 Enter file paths to process:")
    print("   - Enter one path per line")
    print("   - Type 'done' when finished")
    print("   - Drag and drop files if supported")
    print()
    
    file_paths = []
    while True:
        try:
            user_input = input("File path: ").strip()
            
            if user_input.lower() in ['done', 'exit', 'quit', '']:
                break
            
            # Clean up the path (remove quotes)
            clean_path = user_input.strip('"\'')
            
            if os.path.exists(clean_path):
                file_paths.append(clean_path)
                print(f"  ✅ Added: {os.path.basename(clean_path)}")
            else:
                print(f"  ❌ File not found: {clean_path}")
                
        except KeyboardInterrupt:
            print("\n⏹️  Input cancelled")
            break
        except Exception as e:
            print(f"⚠️  Input error: {str(e)}")
    
    if not file_paths:
        print("❌ No files to process. Exiting.")
        return
    
    print(f"\n📋 Processing {len(file_paths)} file(s)...")
    print("=" * 60)
    
    # Process files
    results = {}
    start_time = time.time()
    
    for i, file_path in enumerate(file_paths, 1):
        print(f"\n[{i}/{len(file_paths)}] {'-' * 40}")
        
        try:
            result = extractor.process_file(file_path)
            if result:
                results[file_path] = result
                print(f"✅ Success: {os.path.basename(file_path)}")
            else:
                print(f"❌ Failed: {os.path.basename(file_path)}")
                
        except Exception as e:
            print(f"❌ Error processing {file_path}: {str(e)}")
        
        # Force cleanup between files
        gc.collect()
    
    total_time = time.time() - start_time
    
    # Save results
    if results:
        print("\n" + "=" * 60)
        print("💾 SAVING RESULTS")
        print("=" * 60)
        
        extractor.save_results(results)
        
        print(f"\n🎉 EXTRACTION COMPLETE!")
        print(f"✅ Files processed: {len(results)}/{len(file_paths)}")
        print(f"⏱️  Total time: {total_time:.2f} seconds")
        print(f"⚡ Average: {total_time/len(file_paths):.2f} seconds per file")
        
    else:
        print("\n❌ No files were processed successfully.")
        print("Check the error messages above for details.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⏹️  Process interrupted by user")
    except Exception as e:
        print(f"\n❌ Fatal error: {str(e)}")
