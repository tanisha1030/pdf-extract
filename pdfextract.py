import concurrent.futures
import os
import json
import pandas as pd
from PIL import Image
import pytesseract
from collections import Counter
import zipfile
import tempfile
import warnings
import re
import fitz  # PyMuPDF
import gc  # Add garbage collection
import psutil  # Add memory monitoring
import time

# Optional libraries
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import tabula
    HAS_TABULA = True
except ImportError:
    HAS_TABULA = False

try:
    import camelot
    HAS_CAMELOT = True
except ImportError:
    HAS_CAMELOT = False

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


class OptimizedDocumentExtractor:
    def __init__(self, max_workers=2, memory_limit_mb=1024, extract_formatting=False):
        self.extracted_data = {}
        self.supported_formats = ['.pdf', '.docx', '.pptx', '.xlsx', '.xls']
        self.memory_limit_mb = memory_limit_mb
        self.extract_formatting = extract_formatting
        # Reduce workers for better performance on single files
        self.max_workers = max_workers
        self.setup_directories()

    def setup_directories(self):
        """Create necessary directories"""
        os.makedirs('extracted_images', exist_ok=True)
        os.makedirs('extracted_tables', exist_ok=True)

    def check_memory_usage(self):
        """Check current memory usage"""
        try:
            process = psutil.Process(os.getpid())
            memory_mb = process.memory_info().rss / 1024 / 1024
            return memory_mb
        except:
            return 0

    def is_memory_limit_exceeded(self):
        """Check if memory limit is exceeded"""
        return self.check_memory_usage() > self.memory_limit_mb

    def extract_pdf_optimized(self, pdf_path):
        """Extract PDF data with major performance optimizations"""
        try:
            print(f"📄 Processing PDF: {os.path.basename(pdf_path)}")
            start_time = time.time()
            
            doc = fitz.open(pdf_path)
            file_size_mb = os.path.getsize(pdf_path) / (1024*1024)
            page_count = len(doc)
            
            print(f"  📊 File size: {file_size_mb:.2f} MB, Pages: {page_count}")
            
            pdf_data = {
                'filename': os.path.basename(pdf_path),
                'file_type': 'PDF',
                'metadata': {},
                'pages': [],
                'total_images': 0,
                'total_tables': 0,
                'fonts_used': [],
                'total_words': 0,
                'total_characters': 0,
                'page_count': page_count,
                'file_size_mb': round(file_size_mb, 2),
                'processing_mode': 'optimized'
            }

            # Extract metadata quickly
            try:
                metadata = doc.metadata
                if metadata:
                    pdf_data['metadata'] = {k: v for k, v in metadata.items() if v}
            except:
                pdf_data['metadata'] = {}

            # Process pages sequentially for better performance on single files
            all_fonts = set()
            total_text = ""
            
            for page_num in range(page_count):
                try:
                    page = doc[page_num]
                    
                    # Fast text extraction without formatting
                    page_text = page.get_text()
                    word_count = len(page_text.split()) if page_text else 0
                    char_count = len(page_text)
                    
                    # Only extract images if there are any (check first)
                    images = []
                    image_list = page.get_images()
                    if image_list:
                        images = self.extract_images_fast(page, page_num + 1, pdf_path, max_images=5)
                    
                    # Simple table detection (much faster than external libraries)
                    tables = self.detect_tables_simple(page_text)
                    
                    page_data = {
                        'page_number': page_num + 1,
                        'text': page_text,
                        'images': len(images),  # Just count
                        'tables': len(tables),   # Just count
                        'word_count': word_count,
                        'char_count': char_count,
                        'page_size': {
                            'width': round(page.rect.width, 2),
                            'height': round(page.rect.height, 2)
                        }
                    }
                    
                    pdf_data['pages'].append(page_data)
                    pdf_data['total_images'] += len(images)
                    pdf_data['total_tables'] += len(tables)
                    pdf_data['total_words'] += word_count
                    pdf_data['total_characters'] += char_count
                    
                    total_text += page_text
                    
                    # Progress indicator
                    if page_num % 10 == 0:
                        elapsed = time.time() - start_time
                        print(f"  📄 Processed {page_num + 1}/{page_count} pages ({elapsed:.1f}s)")
                    
                    # Clean up page resources
                    page.clean_contents()
                    
                except Exception as e:
                    print(f"⚠️  Error on page {page_num + 1}: {str(e)}")
                    continue

            # Fast table extraction only if needed
            if pdf_data['total_tables'] > 0:
                print("  📊 Extracting tables (simplified)...")
                pdf_data['extracted_tables'] = self.extract_tables_fast(pdf_path, max_tables=10)
            else:
                pdf_data['extracted_tables'] = []

            doc.close()
            
            elapsed = time.time() - start_time
            print(f"  ✅ PDF processed in {elapsed:.2f}s: {pdf_data['total_words']} words, {pdf_data['total_images']} images")
            
            return pdf_data

        except Exception as e:
            print(f"❌ Error processing PDF {pdf_path}: {str(e)}")
            return None

    def extract_images_fast(self, page, page_num, pdf_path, max_images=5):
        """Fast image extraction with limits"""
        images = []
        try:
            image_list = page.get_images()
            
            # Limit images to process
            for img_index, img in enumerate(image_list[:max_images]):
                try:
                    xref = img[0]
                    # Just get image info, don't extract actual image data
                    base_image = page.parent.extract_image(xref)
                    
                    # Skip large images
                    if base_image.get("width", 0) * base_image.get("height", 0) > 1000000:
                        continue
                    
                    img_info = {
                        'image_index': img_index,
                        'width': base_image.get("width", 0),
                        'height': base_image.get("height", 0),
                        'colorspace': base_image.get("colorspace", "Unknown")
                    }
                    images.append(img_info)
                    
                except Exception as e:
                    continue
                    
        except Exception as e:
            pass
        
        return images

    def detect_tables_simple(self, text):
        """Simple table detection based on text patterns"""
        if not text:
            return []
        
        lines = text.split('\n')
        tables = []
        
        # Look for lines with consistent separators
        for i, line in enumerate(lines):
            if len(line.strip()) > 10:  # Minimum length
                # Check for common table separators
                separators = ['\t', '  ', '|', ',']
                for sep in separators:
                    if line.count(sep) >= 2:  # At least 2 separators
                        # Check if this could be a table
                        parts = line.split(sep)
                        if len(parts) >= 3:  # At least 3 columns
                            tables.append({
                                'line': i,
                                'separator': sep,
                                'columns': len(parts)
                            })
                            break
        
        return tables

    def extract_tables_fast(self, pdf_path, max_tables=10):
        """Fast table extraction using only one method"""
        tables = []
        
        # Use only pdfplumber for speed (most reliable and fastest)
        if HAS_PDFPLUMBER:
            try:
                print("    📊 Using pdfplumber for table extraction...")
                with pdfplumber.open(pdf_path) as pdf:
                    tables_found = 0
                    
                    for page_num, page in enumerate(pdf.pages):
                        if tables_found >= max_tables:
                            break
                        
                        try:
                            page_tables = page.extract_tables()
                            for table in page_tables:
                                if table and len(table) > 1:
                                    tables.append({
                                        'method': 'pdfplumber',
                                        'page': page_num + 1,
                                        'data': table[:20],  # Limit rows
                                        'shape': (len(table), len(table[0]) if table else 0)
                                    })
                                    tables_found += 1
                                    
                                    if tables_found >= max_tables:
                                        break
                        except:
                            continue
                            
            except Exception as e:
                print(f"    ⚠️  Table extraction error: {str(e)}")
        
        return tables

    def extract_docx_fast(self, docx_path):
        """Fast DOCX extraction"""
        if not HAS_DOCX:
            print("❌ python-docx not available")
            return None
            
        try:
            print(f"📄 Processing DOCX: {os.path.basename(docx_path)}")
            start_time = time.time()
            
            doc = Document(docx_path)
            
            docx_data = {
                'filename': os.path.basename(docx_path),
                'file_type': 'DOCX',
                'paragraphs': [],
                'tables': [],
                'total_words': 0,
                'total_characters': 0,
                'file_size_mb': round(os.path.getsize(docx_path) / (1024*1024), 2)
            }
            
            # Extract paragraphs (limit to avoid memory issues)
            for i, para in enumerate(doc.paragraphs):
                if i >= 1000:  # Limit paragraphs
                    break
                if para.text.strip():
                    docx_data['paragraphs'].append({
                        'text': para.text[:500],  # Limit text length
                        'style': para.style.name if para.style else 'Normal'
                    })
                    docx_data['total_words'] += len(para.text.split())
                    docx_data['total_characters'] += len(para.text)
            
            # Extract tables (limit to avoid memory issues)
            for table_idx, table in enumerate(doc.tables):
                if table_idx >= 10:  # Limit tables
                    break
                
                table_data = []
                for row_idx, row in enumerate(table.rows):
                    if row_idx >= 50:  # Limit rows
                        break
                    row_data = [cell.text.strip()[:100] for cell in row.cells]  # Limit cell text
                    table_data.append(row_data)
                
                if table_data and len(table_data) > 1:
                    docx_data['tables'].append({
                        'table_index': table_idx,
                        'data': table_data,
                        'shape': (len(table_data), len(table_data[0]) if table_data else 0)
                    })
            
            elapsed = time.time() - start_time
            print(f"  ✅ DOCX processed in {elapsed:.2f}s: {docx_data['total_words']} words")
            
            return docx_data
            
        except Exception as e:
            print(f"❌ Error processing DOCX {docx_path}: {str(e)}")
            return None

    def extract_pptx_fast(self, pptx_path):
        """Fast PPTX extraction"""
        if not HAS_PPTX:
            print("❌ python-pptx not available")
            return None
            
        try:
            print(f"📄 Processing PPTX: {os.path.basename(pptx_path)}")
            start_time = time.time()
            
            prs = Presentation(pptx_path)
            
            pptx_data = {
                'filename': os.path.basename(pptx_path),
                'file_type': 'PPTX',
                'slides': [],
                'total_slides': len(prs.slides),
                'total_words': 0,
                'total_characters': 0,
                'file_size_mb': round(os.path.getsize(pptx_path) / (1024*1024), 2)
            }
            
            # Limit slides processed
            max_slides = min(100, len(prs.slides))
            
            for slide_idx, slide in enumerate(prs.slides[:max_slides]):
                slide_data = {
                    'slide_number': slide_idx + 1,
                    'text_content': [],
                    'tables': []
                }
                
                # Extract text from shapes (limit)
                text_count = 0
                for shape in slide.shapes:
                    if text_count >= 20:  # Limit text shapes
                        break
                    
                    try:
                        if hasattr(shape, "text") and shape.text.strip():
                            text = shape.text.strip()[:500]  # Limit text length
                            slide_data['text_content'].append(text)
                            pptx_data['total_words'] += len(text.split())
                            pptx_data['total_characters'] += len(text)
                            text_count += 1
                    except:
                        continue
                
                pptx_data['slides'].append(slide_data)
            
            elapsed = time.time() - start_time
            print(f"  ✅ PPTX processed in {elapsed:.2f}s: {pptx_data['total_words']} words")
            
            return pptx_data
            
        except Exception as e:
            print(f"❌ Error processing PPTX {pptx_path}: {str(e)}")
            return None

    def extract_excel_fast(self, excel_path):
        """Fast Excel extraction"""
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
            
            # Read with strict limits
            max_rows = 100 if file_size_mb > 10 else 500
            
            try:
                xl_file = pd.ExcelFile(excel_path)
                sheet_names = xl_file.sheet_names[:10]  # Limit sheets
                
                for sheet_name in sheet_names:
                    try:
                        # Read only first few rows
                        df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=max_rows)
                        
                        if not df.empty:
                            # Limit columns too
                            if df.shape[1] > 20:
                                df = df.iloc[:, :20]
                            
                            sheet_data = {
                                'sheet_name': sheet_name,
                                'shape': df.shape,
                                'columns': df.columns.tolist(),
                                'data': df.head(10).to_dict('records'),  # Only first 10 rows
                                'processing_note': f'Limited to {max_rows} rows and 20 columns'
                            }
                            
                            excel_data['sheets'].append(sheet_data)
                        
                    except Exception as e:
                        print(f"    ⚠️  Error reading sheet '{sheet_name}': {str(e)}")
                        continue
                
                xl_file.close()
                
            except Exception as e:
                print(f"    ❌ Error opening Excel file: {str(e)}")
                return None
            
            elapsed = time.time() - start_time
            print(f"  ✅ Excel processed in {elapsed:.2f}s: {len(excel_data['sheets'])} sheets")
            
            return excel_data
            
        except Exception as e:
            print(f"❌ Error processing Excel {excel_path}: {str(e)}")
            return None

    def process_files(self, file_paths):
        """Process files sequentially for better performance"""
        results = {}
        
        print(f"\n🔄 Processing {len(file_paths)} file(s)...")
        print("=" * 60)
        
        for i, file_path in enumerate(file_paths, 1):
            file_ext = os.path.splitext(file_path)[1].lower()
            
            print(f"\n[{i}/{len(file_paths)}] Processing: {os.path.basename(file_path)}")
            
            start_time = time.time()
            
            try:
                if file_ext == '.pdf':
                    result = self.extract_pdf_optimized(file_path)
                elif file_ext == '.docx':
                    result = self.extract_docx_fast(file_path)
                elif file_ext == '.pptx':
                    result = self.extract_pptx_fast(file_path)
                elif file_ext in ['.xlsx', '.xls']:
                    result = self.extract_excel_fast(file_path)
                else:
                    print(f"❌ Unsupported file format: {file_ext}")
                    continue
                
                elapsed = time.time() - start_time
                
                if result:
                    results[file_path] = result
                    print(f"✅ File completed in {elapsed:.2f}s")
                else:
                    print(f"❌ Failed to process file")
                    
            except Exception as e:
                print(f"❌ Error processing {file_path}: {str(e)}")
                continue
            
            # Force garbage collection between files
            gc.collect()
        
        return results

    def save_results(self, results, output_format='json'):
        """Save extraction results"""
        if not results:
            print("❌ No results to save.")
            return
        
        try:
            # Save as JSON
            with open('extraction_results.json', 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False, default=str)
            print("✅ Results saved as: extraction_results.json")
            
            # Create simple summary
            self.create_simple_summary(results)
            
        except Exception as e:
            print(f"❌ Error saving results: {str(e)}")

    def create_simple_summary(self, results):
        """Create a simple summary report"""
        try:
            with open('extraction_summary.txt', 'w', encoding='utf-8') as f:
                f.write("📋 DOCUMENT EXTRACTION SUMMARY\n")
                f.write("=" * 50 + "\n")
                f.write(f"📁 Total Files Processed: {len(results)}\n\n")
                
                total_words = sum(data.get('total_words', 0) for data in results.values())
                total_chars = sum(data.get('total_characters', 0) for data in results.values())
                total_size_mb = sum(data.get('file_size_mb', 0) for data in results.values())
                
                f.write("📊 TOTALS:\n")
                f.write(f"📝 Words: {total_words:,}\n")
                f.write(f"🔤 Characters: {total_chars:,}\n")
                f.write(f"💾 File Size: {total_size_mb:.2f} MB\n\n")
                
                f.write("📄 FILES:\n")
                for file_path, data in results.items():
                    f.write(f"- {data['filename']} ({data['file_type']})\n")
                    if data['file_type'] == 'PDF':
                        f.write(f"  Pages: {data.get('page_count', 0)}\n")
                    elif data['file_type'] == 'DOCX':
                        f.write(f"  Paragraphs: {len(data.get('paragraphs', []))}\n")
                    elif data['file_type'] == 'PPTX':
                        f.write(f"  Slides: {data.get('total_slides', 0)}\n")
                    elif data['file_type'] == 'Excel':
                        f.write(f"  Sheets: {len(data.get('sheets', []))}\n")
                    f.write(f"  Words: {data.get('total_words', 0):,}\n\n")
            
            print("✅ Summary saved as: extraction_summary.txt")
            
        except Exception as e:
            print(f"❌ Error creating summary: {str(e)}")


def main():
    """Main function"""
    print("🚀 Optimized Document Extractor")
    print("=" * 60)
    print("Supported formats: PDF, DOCX, PPTX, XLSX, XLS")
    print("⚡ Performance optimized for fast processing")
    print()
    
    # Initialize with optimized settings
    extractor = OptimizedDocumentExtractor(
        max_workers=2,  # Reduced for better single-file performance
        extract_formatting=False  # Disabled for speed
    )
    
    # Get file paths
    print("📁 Enter file paths to process:")
    file_paths = []
    
    while True:
        user_input = input("File path (or 'done' to start): ").strip()
        
        if user_input.lower() in ['done', 'finish', 'exit', '']:
            break
        
        if ',' in user_input:
            paths = [path.strip().strip('"\'') for path in user_input.split(',')]
            file_paths.extend(paths)
        else:
            file_paths.append(user_input.strip('"\''))
    
    if not file_paths:
        print("❌ No files provided. Exiting.")
        return
    
    # Validate files
    valid_paths = []
    for path in file_paths:
        if os.path.exists(path):
            ext = os.path.splitext(path)[1].lower()
            if ext in extractor.supported_formats:
                valid_paths.append(path)
            else:
                print(f"⚠️  Unsupported: {path}")
        else:
            print(f"⚠️  Not found: {path}")
    
    if not valid_paths:
        print("❌ No valid files to process.")
        return
    
    # Process files
    try:
        start_time = time.time()
        results = extractor.process_files(valid_paths)
        total_time = time.time() - start_time
        
        if results:
            extractor.save_results(results)
            
            print("\n" + "=" * 60)
            print("📊 PROCESSING COMPLETE")
            print("=" * 60)
            print(f"✅ Files processed: {len(results)}")
            print(f"⏱️  Total time: {total_time:.2f} seconds")
            print(f"⚡ Average: {total_time/len(results):.2f} seconds per file")
            print("\n🎉 Extraction completed successfully!")
            
        else:
            print("❌ No files were processed successfully.")
            
    except KeyboardInterrupt:
        print("\n⏹️  Process interrupted by user.")
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")


if __name__ == "__main__":
    main()
