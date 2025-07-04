import fitz  # PyMuPDF
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
import time
import gc
import logging
from functools import partial
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from typing import Dict, List, Optional, Tuple, Any
import pickle
import sqlite3
import hashlib
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extraction.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Try importing optional libraries with fallbacks
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False
    logger.warning("pdfplumber not available. Some table extraction features may be limited.")

try:
    import tabula
    HAS_TABULA = True
except ImportError:
    HAS_TABULA = False
    logger.warning("tabula-py not available. Some table extraction features may be limited.")

try:
    import camelot
    HAS_CAMELOT = True
except ImportError:
    HAS_CAMELOT = False
    logger.warning("camelot-py not available. Some table extraction features may be limited.")

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False
    logger.warning("OpenCV not available. Some image processing features may be limited.")

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    logger.warning("python-docx not available. DOCX processing will be skipped.")

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False
    logger.warning("python-pptx not available. PPTX processing will be skipped.")

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    logger.warning("openpyxl not available. Excel processing may be limited.")

class OptimizedDocumentExtractor:
    def __init__(self, max_workers=None, chunk_size=50, enable_caching=True, memory_limit_mb=1024):
        self.extracted_data = {}
        self.supported_formats = ['.pdf', '.docx', '.pptx', '.xlsx', '.xls']
        self.max_workers = max_workers or min(4, multiprocessing.cpu_count())
        self.chunk_size = chunk_size
        self.enable_caching = enable_caching
        self.memory_limit_mb = memory_limit_mb
        self.setup_directories()
        self.setup_database()
        
        # Performance tracking
        self.processing_stats = {
            'pages_processed': 0,
            'images_extracted': 0,
            'tables_extracted': 0,
            'memory_usage_mb': 0,
            'processing_time': 0
        }
    
    def setup_directories(self):
        """Create necessary directories with better organization"""
        base_dirs = ['extracted_images', 'extracted_tables', 'temp_processing', 'cache']
        for directory in base_dirs:
            os.makedirs(directory, exist_ok=True)
    
    def setup_database(self):
        """Setup SQLite database for caching and large data handling"""
        if not self.enable_caching:
            return
            
        try:
            self.db_path = 'extraction_cache.db'
            self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
            self.conn.execute('''
                CREATE TABLE IF NOT EXISTS file_cache (
                    file_hash TEXT PRIMARY KEY,
                    filename TEXT,
                    file_size INTEGER,
                    modification_time REAL,
                    extraction_data BLOB,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            self.conn.execute('''
                CREATE TABLE IF NOT EXISTS page_cache (
                    page_hash TEXT PRIMARY KEY,
                    file_hash TEXT,
                    page_number INTEGER,
                    page_data BLOB,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            self.conn.commit()
            logger.info("Database cache initialized")
        except Exception as e:
            logger.error(f"Failed to setup database: {e}")
            self.enable_caching = False
    
    def get_file_hash(self, file_path: str) -> str:
        """Generate hash for file to enable caching"""
        try:
            stat = os.stat(file_path)
            content = f"{file_path}_{stat.st_size}_{stat.st_mtime}"
            return hashlib.md5(content.encode()).hexdigest()
        except Exception as e:
            logger.error(f"Error generating file hash: {e}")
            return hashlib.md5(file_path.encode()).hexdigest()
    
    def check_cache(self, file_path: str) -> Optional[Dict]:
        """Check if file data is cached"""
        if not self.enable_caching:
            return None
            
        try:
            file_hash = self.get_file_hash(file_path)
            cursor = self.conn.cursor()
            cursor.execute(
                "SELECT extraction_data FROM file_cache WHERE file_hash = ?",
                (file_hash,)
            )
            result = cursor.fetchone()
            if result:
                cached_data = pickle.loads(result[0])
                logger.info(f"Using cached data for {os.path.basename(file_path)}")
                return cached_data
        except Exception as e:
            logger.error(f"Error checking cache: {e}")
        return None
    
    def save_to_cache(self, file_path: str, data: Dict):
        """Save extraction data to cache"""
        if not self.enable_caching:
            return
            
        try:
            file_hash = self.get_file_hash(file_path)
            stat = os.stat(file_path)
            
            pickled_data = pickle.dumps(data)
            
            self.conn.execute('''
                INSERT OR REPLACE INTO file_cache 
                (file_hash, filename, file_size, modification_time, extraction_data)
                VALUES (?, ?, ?, ?, ?)
            ''', (file_hash, os.path.basename(file_path), stat.st_size, stat.st_mtime, pickled_data))
            self.conn.commit()
            logger.info(f"Cached data for {os.path.basename(file_path)}")
        except Exception as e:
            logger.error(f"Error saving to cache: {e}")
    
    def monitor_memory_usage(self):
        """Monitor and manage memory usage"""
        try:
            import psutil
            process = psutil.Process(os.getpid())
            memory_mb = process.memory_info().rss / 1024 / 1024
            self.processing_stats['memory_usage_mb'] = memory_mb
            
            if memory_mb > self.memory_limit_mb:
                logger.warning(f"Memory usage ({memory_mb:.1f}MB) exceeds limit ({self.memory_limit_mb}MB)")
                gc.collect()  # Force garbage collection
                
                # Further memory cleanup if needed
                if memory_mb > self.memory_limit_mb * 1.5:
                    logger.warning("Forcing aggressive memory cleanup")
                    return True  # Signal to reduce processing intensity
                    
        except ImportError:
            logger.warning("psutil not available for memory monitoring")
        except Exception as e:
            logger.error(f"Error monitoring memory: {e}")
        
        return False
    
    def extract_pdf_chunked(self, pdf_path: str) -> Optional[Dict]:
        """Extract PDF data in chunks to handle large files"""
        # Check cache first
        cached_data = self.check_cache(pdf_path)
        if cached_data:
            return cached_data
        
        try:
            logger.info(f"Processing PDF: {os.path.basename(pdf_path)}")
            doc = fitz.open(pdf_path)
            
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
                'page_count': len(doc),
                'file_size_mb': round(os.path.getsize(pdf_path) / (1024*1024), 2)
            }
            
            # Extract metadata safely
            try:
                metadata = doc.metadata
                if metadata:
                    pdf_data['metadata'] = {k: v for k, v in metadata.items() if v}
            except:
                pdf_data['metadata'] = {}
            
            # Process pages in chunks
            total_pages = len(doc)
            processed_pages = 0
            
            for chunk_start in range(0, total_pages, self.chunk_size):
                chunk_end = min(chunk_start + self.chunk_size, total_pages)
                logger.info(f"Processing pages {chunk_start + 1}-{chunk_end} of {total_pages}")
                
                # Check memory before processing chunk
                if self.monitor_memory_usage():
                    # Reduce chunk size if memory is high
                    self.chunk_size = max(10, self.chunk_size // 2)
                    logger.info(f"Reduced chunk size to {self.chunk_size} due to memory constraints")
                
                # Process chunk with threading
                chunk_pages = []
                page_futures = []
                
                with ThreadPoolExecutor(max_workers=min(self.max_workers, chunk_end - chunk_start)) as executor:
                    for page_num in range(chunk_start, chunk_end):
                        try:
                            page = doc[page_num]
                            future = executor.submit(self.extract_page_data_optimized, page, page_num + 1, pdf_path)
                            page_futures.append((page_num, future))
                        except Exception as e:
                            logger.error(f"Error submitting page {page_num + 1}: {e}")
                            continue
                    
                    # Collect results
                    for page_num, future in page_futures:
                        try:
                            page_data = future.result(timeout=30)  # 30 second timeout per page
                            if page_data:
                                chunk_pages.append(page_data)
                                
                                # Update totals
                                pdf_data['total_images'] += len(page_data.get('images', []))
                                pdf_data['total_tables'] += len(page_data.get('tables', []))
                                pdf_data['total_words'] += page_data.get('word_count', 0)
                                pdf_data['total_characters'] += page_data.get('char_count', 0)
                                
                                processed_pages += 1
                                self.processing_stats['pages_processed'] += 1
                                
                        except Exception as e:
                            logger.error(f"Error processing page {page_num + 1}: {e}")
                            continue
                
                # Add chunk pages to main data
                pdf_data['pages'].extend(chunk_pages)
                
                # Periodic cleanup
                if processed_pages % 100 == 0:
                    gc.collect()
                    logger.info(f"Processed {processed_pages}/{total_pages} pages")
            
            # Extract unique fonts
            all_fonts = []
            for page in pdf_data['pages']:
                all_fonts.extend(page.get('fonts', []))
            pdf_data['fonts_used'] = list(set(all_fonts))
            
            # Extract tables using optimized methods
            logger.info("Extracting tables with optimized methods...")
            pdf_data['extracted_tables'] = self.extract_tables_optimized(pdf_path, doc)
            
            doc.close()
            
            # Save to cache
            self.save_to_cache(pdf_path, pdf_data)
            
            logger.info(f"PDF processed successfully: {pdf_data['page_count']} pages, {pdf_data['total_words']} words")
            return pdf_data
            
        except Exception as e:
            logger.error(f"Error processing PDF {pdf_path}: {e}")
            return None
    
    def extract_page_data_optimized(self, page, page_num: int, pdf_path: str) -> Dict:
        """Optimized page data extraction with memory management"""
        page_data = {
            'page_number': page_num,
            'text': '',
            'images': [],
            'tables': [],
            'fonts': [],
            'word_count': 0,
            'char_count': 0,
            'page_size': {
                'width': round(page.rect.width, 2),
                'height': round(page.rect.height, 2)
            },
            'rotation': page.rotation
        }
        
        try:
            # Extract text more efficiently
            text_dict = page.get_text("dict")
            page_text = ""
            fonts_on_page = set()
            
            for block in text_dict.get("blocks", []):
                if "lines" in block:
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span.get("text", "")
                            line_text += text
                            fonts_on_page.add(span.get('font', 'Unknown'))
                        page_text += line_text + " "
            
            page_data['text'] = page_text.strip()
            page_data['fonts'] = list(fonts_on_page)
            page_data['word_count'] = len(page_text.split()) if page_text else 0
            page_data['char_count'] = len(page_text)
            
            # Extract images with size limits
            page_data['images'] = self.extract_images_optimized(page, page_num, pdf_path)
            
            # Extract tables with improved detection
            page_data['tables'] = self.extract_tables_from_page_optimized(page)
            
        except Exception as e:
            logger.error(f"Error extracting from page {page_num}: {e}")
        
        return page_data
    
    def extract_images_optimized(self, page, page_num: int, pdf_path: str) -> List[Dict]:
        """Optimized image extraction with size and quality controls"""
        images = []
        
        try:
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    base_image = page.parent.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # Skip very small images (likely decorative)
                    if base_image["width"] < 50 or base_image["height"] < 50:
                        continue
                    
                    # Skip very large images if memory is constrained
                    if len(image_bytes) > 10 * 1024 * 1024:  # 10MB limit
                        logger.warning(f"Skipping large image ({len(image_bytes)/1024/1024:.1f}MB) on page {page_num}")
                        continue
                    
                    # Create safe filename
                    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                    img_name = f"extracted_images/{base_name}_page_{page_num}_img_{img_index + 1}.png"
                    
                    # Save image with compression
                    try:
                        with Image.open(io.BytesIO(image_bytes)) as pil_img:
                            # Resize if too large
                            if pil_img.width > 2000 or pil_img.height > 2000:
                                pil_img.thumbnail((2000, 2000), Image.Resampling.LANCZOS)
                            
                            pil_img.save(img_name, 'PNG', optimize=True, compress_level=6)
                    except Exception:
                        # Fallback to raw save
                        with open(img_name, "wb") as img_file:
                            img_file.write(image_bytes)
                    
                    img_info = {
                        'filename': img_name,
                        'width': base_image["width"],
                        'height': base_image["height"],
                        'colorspace': base_image.get("colorspace", "Unknown"),
                        'size_bytes': len(image_bytes)
                    }
                    
                    images.append(img_info)
                    self.processing_stats['images_extracted'] += 1
                    
                except Exception as e:
                    logger.error(f"Error extracting image {img_index} from page {page_num}: {e}")
                    continue
                    
        except Exception as e:
            logger.error(f"Error accessing images on page {page_num}: {e}")
        
        return images
    
    def extract_tables_from_page_optimized(self, page) -> List[Dict]:
        """Optimized table detection with better performance"""
        tables = []
        try:
            # Use simplified table detection for large documents
            text_dict = page.get_text("dict")
            
            potential_tables = []
            
            for block in text_dict.get("blocks", []):
                if "lines" in block and len(block["lines"]) >= 3:
                    lines = block["lines"]
                    
                    # Quick check for table-like structure
                    line_texts = []
                    for line in lines:
                        row_text = ""
                        for span in line["spans"]:
                            row_text += span.get("text", "") + " "
                        line_texts.append(row_text.strip())
                    
                    # Filter out empty lines
                    line_texts = [line for line in line_texts if line.strip()]
                    
                    if len(line_texts) >= 2:
                        # Quick heuristic check
                        if self.quick_table_check(line_texts):
                            potential_tables.append({
                                'bbox': block.get('bbox', [0, 0, 0, 0]),
                                'data': line_texts,
                                'method': 'optimized_detection',
                                'confidence': 'medium'
                            })
            
            # Limit number of tables per page to prevent memory issues
            tables = potential_tables[:10]  # Max 10 tables per page
            self.processing_stats['tables_extracted'] += len(tables)
            
        except Exception as e:
            logger.error(f"Error extracting tables from page: {e}")
        
        return tables
    
    def quick_table_check(self, text_lines: List[str]) -> bool:
        """Quick heuristic to identify potential tables"""
        if len(text_lines) < 2:
            return False
        
        # Check for consistent separators
        separators = ['\t', '  ', '|', ',']
        for sep in separators:
            counts = [line.count(sep) for line in text_lines]
            if len(set(counts)) <= 2 and max(counts) > 0:  # Consistent separator usage
                return True
        
        # Check for numeric content
        numeric_lines = sum(1 for line in text_lines if re.search(r'\d+', line))
        if numeric_lines / len(text_lines) > 0.4:
            return True
        
        return False
    
    def extract_tables_optimized(self, pdf_path: str, doc=None) -> List[Dict]:
        """Optimized table extraction for large files"""
        all_tables = []
        
        # For very large files, limit table extraction methods
        file_size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        
        if file_size_mb > 100:  # Files larger than 100MB
            logger.info("Large file detected, using lightweight table extraction")
            return all_tables  # Skip heavy table extraction for very large files
        
        # Method 1: Using tabula-py (most reliable but slower)
        if HAS_TABULA and file_size_mb < 50:  # Only for files < 50MB
            try:
                logger.info("Extracting tables with tabula...")
                tabula_tables = tabula.read_pdf(
                    pdf_path, 
                    pages='all', 
                    multiple_tables=True, 
                    silent=True,
                    pandas_options={'dtype': str}  # Prevent type conversion issues
                )
                
                for i, table in enumerate(tabula_tables):
                    if not table.empty and table.shape[0] > 1 and table.shape[1] > 1:
                        if self.validate_extracted_table(table):
                            all_tables.append({
                                'method': 'tabula',
                                'table_index': i,
                                'data': table.head(100).to_dict('records'),  # Limit rows
                                'shape': table.shape,
                                'confidence': 'high'
                            })
                            
                            # Limit total tables to prevent memory issues
                            if len(all_tables) >= 20:
                                break
                
                logger.info(f"Tabula found {len(all_tables)} tables")
                
            except Exception as e:
                logger.error(f"Tabula extraction failed: {e}")
        
        return all_tables
    
    def validate_extracted_table(self, df) -> bool:
        """Validate if a pandas DataFrame represents a real table"""
        if df.empty or df.shape[0] < 2 or df.shape[1] < 2:
            return False
        
        # Check for meaningful content
        text_content = df.astype(str).values.flatten()
        non_empty_cells = [cell for cell in text_content if cell.strip() and cell.strip() != 'nan']
        
        if len(non_empty_cells) < df.shape[0] * df.shape[1] * 0.2:  # At least 20% cells should have content
            return False
        
        return True
    
    def extract_docx_optimized(self, docx_path: str) -> Optional[Dict]:
        """Optimized DOCX extraction"""
        if not HAS_DOCX:
            logger.error("python-docx not available. Skipping DOCX processing.")
            return None
        
        # Check cache first
        cached_data = self.check_cache(docx_path)
        if cached_data:
            return cached_data
            
        try:
            logger.info(f"Processing DOCX: {os.path.basename(docx_path)}")
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
            
            # Extract paragraphs with limits
            for i, para in enumerate(doc.paragraphs):
                if i >= 10000:  # Limit paragraphs for very large documents
                    logger.warning("Paragraph limit reached, truncating extraction")
                    break
                    
                if para.text.strip():
                    docx_data['paragraphs'].append({
                        'text': para.text[:1000],  # Limit text length
                        'style': para.style.name if para.style else 'Normal'
                    })
                    docx_data['total_words'] += len(para.text.split())
                    docx_data['total_characters'] += len(para.text)
            
            # Extract tables with limits
            for table_idx, table in enumerate(doc.tables):
                if table_idx >= 50:  # Limit tables
                    logger.warning("Table limit reached, truncating extraction")
                    break
                    
                table_data = []
                for row_idx, row in enumerate(table.rows):
                    if row_idx >= 100:  # Limit rows per table
                        break
                    row_data = [cell.text.strip()[:200] for cell in row.cells]  # Limit cell text
                    table_data.append(row_data)
                
                if table_data and len(table_data) > 1:
                    docx_data['tables'].append({
                        'table_index': table_idx,
                        'data': table_data,
                        'shape': (len(table_data), len(table_data[0]) if table_data else 0)
                    })
            
            # Save to cache
            self.save_to_cache(docx_path, docx_data)
            
            logger.info(f"DOCX processed: {len(docx_data['paragraphs'])} paragraphs, {docx_data['total_words']} words")
            return docx_data
            
        except Exception as e:
            logger.error(f"Error processing DOCX {docx_path}: {e}")
            return None
    
    def process_files_optimized(self, file_paths: List[str]) -> Dict:
        """Optimized file processing with better resource management"""
        results = {}
        
        logger.info(f"Processing {len(file_paths)} file(s) with optimized methods...")
        
        for i, file_path in enumerate(file_paths, 1):
            file_ext = os.path.splitext(file_path)[1].lower()
            
            logger.info(f"[{i}/{len(file_paths)}] Processing: {file_path}")
            
            # Check file size and warn for very large files
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > 500:
                logger.warning(f"Very large file detected ({file_size_mb:.1f}MB). Processing may take significant time.")
            
            start_time = time.time()
            
            try:
                if file_ext == '.pdf':
                    result = self.extract_pdf_chunked(file_path)
                elif file_ext == '.docx':
                    result = self.extract_docx_optimized(file_path)
                elif file_ext == '.pptx':
                    result = self.extract_pptx_optimized(file_path)
                elif file_ext in ['.xlsx', '.xls']:
                    result = self.extract_excel_optimized(file_path)
                else:
                    logger.warning(f"Unsupported file format: {file_ext}")
                    continue
                
                processing_time = time.time() - start_time
                
                if result:
                    result['processing_time_seconds'] = round(processing_time, 2)
                    results[file_path] = result
                    logger.info(f"Successfully processed: {file_path} ({processing_time:.2f}s)")
                else:
                    logger.error(f"Failed to process: {file_path}")
                    
            except Exception as e:
                logger.error(f"Error processing {file_path}: {e}")
                continue
        
        return results
    
    def extract_pptx_optimized(self, pptx_path: str) -> Optional[Dict]:
        """Optimized PPTX extraction"""
        if not HAS_PPTX:
            logger.error("python-pptx not available. Skipping PPTX processing.")
            return None
        
        # Check cache first
        cached_data = self.check_cache(pptx_path)
        if cached_data:
            return cached_data
            
        try:
            logger.info(f"Processing PPTX: {os.path.basename(pptx_path)}")
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
            
            # Limit slides for very large presentations
            max_slides = min(500, len(prs.slides))
            if max_slides < len(prs.slides):
                logger.warning(f"Limiting extraction to first {max_slides} slides")
            
            for slide_idx, slide in enumerate(prs.slides[:max_slides]):
                slide_data = {
                    'slide_number': slide_idx + 1,
                    'text_content': [],
                    'tables': []
                }
                
                # Extract text from shapes
                for shape in slide.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text.strip():
                            text = shape.text.strip()[:1000]  # Limit text length
                            slide_data['text_content'].append(text)
                            pptx_data['total_words'] += len(text.split())
                            pptx_data['total_characters'] += len(text)
                        
                        # Extract tables
                        if hasattr(shape, "table") and shape.table:
                            table_data = []
                            for row_idx, row in enumerate(shape.table.rows):
                                if row_idx >= 50:  # Limit rows
                                    break
                                row_data = [cell.text.strip()[:200] for cell in row.cells]
                                table_data.append(row_data)
                            if table_data and len(table_data) > 1:
                                slide_data['tables'].append(table_data)
                    except Exception as e:
                        continue
                
                pptx_data['slides'].append(slide_data)
            
            # Save to cache
            self.save_to_cache(pptx_path, pptx_data)
            
            logger.info(f"PPTX processed: {pptx_data['total_slides']} slides, {pptx_data['total_words']} words")
            return pptx_data
            
        except Exception as e:
            logger.error(f"Error processing PPTX {pptx_path}: {e}")
            return None
    
    def extract_excel_optimized(self, excel_path: str) -> Optional[Dict]:
        """Optimized Excel extraction"""
        # Check cache first
        cached_data = self.check_cache(excel_path)
        if cached_data:
            return cached_data
            
        try:
            logger.info(f"Processing Excel: {os.path.basename(excel_path)}")
            
            excel_data = {
                'filename': os.path.basename(excel_path),
                'file_type': 'Excel',
                'sheets': [],
                'total_sheets': 0,
                'total_rows': 0,
                'total_columns': 0,
                'file_size_mb': round(os.path.getsize(excel_path) / (1024*1024), 2)
            }
            
            # Try different Excel reading methods
            try:
                # Method 1: Using pandas (most reliable for data)
                xl_file = pd.ExcelFile(excel_path)
                sheet_names = xl_file.sheet_names
                excel_data['total_sheets'] = len(sheet_names)
                
                # Limit sheets for very large files
                max_sheets = min(20, len(sheet_names))
                if max_sheets < len(sheet_names):
                    logger.warning(f"Limiting extraction to first {max_sheets} sheets")
                
                for sheet_name in sheet_names[:max_sheets]:
                    try:
                        df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=1000)  # Limit rows
                        
                        if not df.empty:
                            sheet_data = {
                                'sheet_name': sheet_name,
                                'shape': df.shape,
                                'columns': df.columns.tolist()[:50],  # Limit columns
                                'data': df.head(100).to_dict('records'),  # Limit data
                                'has_data': True
                            }
                            excel_data['sheets'].append(sheet_data)
                            excel_data['total_rows'] += df.shape[0]
                            excel_data['total_columns'] += df.shape[1]
                        else:
                            excel_data['sheets'].append({
                                'sheet_name': sheet_name,
                                'shape': (0, 0),
                                'columns': [],
                                'data': [],
                                'has_data': False
                            })
                    except Exception as e:
                        logger.error(f"Error reading sheet {sheet_name}: {e}")
                        continue
                        
            except Exception as e:
                logger.error(f"Error with pandas Excel reading: {e}")
                
                # Fallback method using openpyxl
                if HAS_OPENPYXL:
                    try:
                        from openpyxl import load_workbook
                        wb = load_workbook(excel_path, read_only=True)
                        
                        excel_data['total_sheets'] = len(wb.sheetnames)
                        
                        for sheet_name in wb.sheetnames[:20]:  # Limit sheets
                            ws = wb[sheet_name]
                            
                            # Get sheet dimensions
                            max_row = min(ws.max_row, 1000)  # Limit rows
                            max_col = min(ws.max_column, 50)  # Limit columns
                            
                            sheet_data = {
                                'sheet_name': sheet_name,
                                'shape': (max_row, max_col),
                                'columns': [f'Column_{i}' for i in range(1, max_col + 1)],
                                'data': [],
                                'has_data': max_row > 0 and max_col > 0
                            }
                            
                            # Extract limited data
                            for row in ws.iter_rows(min_row=1, max_row=min(100, max_row), 
                                                   min_col=1, max_col=max_col, values_only=True):
                                row_data = [str(cell) if cell is not None else '' for cell in row]
                                sheet_data['data'].append(row_data)
                            
                            excel_data['sheets'].append(sheet_data)
                            excel_data['total_rows'] += max_row
                            excel_data['total_columns'] += max_col
                            
                        wb.close()
                        
                    except Exception as e:
                        logger.error(f"Error with openpyxl Excel reading: {e}")
                        return None
                else:
                    logger.error("Neither pandas nor openpyxl available for Excel processing")
                    return None
            
            # Save to cache
            self.save_to_cache(excel_path, excel_data)
            
            logger.info(f"Excel processed: {excel_data['total_sheets']} sheets, {excel_data['total_rows']} total rows")
            return excel_data
            
        except Exception as e:
            logger.error(f"Error processing Excel {excel_path}: {e}")
            return None
    
    def save_extracted_data(self, output_path: str = "extracted_data.json"):
        """Save all extracted data to JSON file"""
        try:
            # Convert results to JSON-serializable format
            json_data = {}
            
            for file_path, data in self.extracted_data.items():
                # Create a simplified version for JSON export
                simplified_data = {
                    'filename': data.get('filename', ''),
                    'file_type': data.get('file_type', ''),
                    'processing_time_seconds': data.get('processing_time_seconds', 0),
                    'file_size_mb': data.get('file_size_mb', 0),
                    'summary': self.generate_file_summary(data)
                }
                
                # Add type-specific data
                if data.get('file_type') == 'PDF':
                    simplified_data.update({
                        'page_count': data.get('page_count', 0),
                        'total_words': data.get('total_words', 0),
                        'total_images': data.get('total_images', 0),
                        'total_tables': data.get('total_tables', 0),
                        'fonts_used': data.get('fonts_used', [])[:10]  # Limit fonts
                    })
                elif data.get('file_type') == 'DOCX':
                    simplified_data.update({
                        'paragraph_count': len(data.get('paragraphs', [])),
                        'table_count': len(data.get('tables', [])),
                        'total_words': data.get('total_words', 0)
                    })
                elif data.get('file_type') == 'PPTX':
                    simplified_data.update({
                        'slide_count': data.get('total_slides', 0),
                        'total_words': data.get('total_words', 0)
                    })
                elif data.get('file_type') == 'Excel':
                    simplified_data.update({
                        'sheet_count': data.get('total_sheets', 0),
                        'total_rows': data.get('total_rows', 0),
                        'total_columns': data.get('total_columns', 0)
                    })
                
                json_data[file_path] = simplified_data
            
            # Add processing statistics
            json_data['processing_statistics'] = self.processing_stats
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Extracted data saved to: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Error saving extracted data: {e}")
            return None
    
    def generate_file_summary(self, data: Dict) -> Dict:
        """Generate a summary of extracted file data"""
        summary = {
            'file_type': data.get('file_type', 'Unknown'),
            'processing_success': True,
            'key_metrics': {}
        }
        
        try:
            if data.get('file_type') == 'PDF':
                summary['key_metrics'] = {
                    'pages': data.get('page_count', 0),
                    'words': data.get('total_words', 0),
                    'images': data.get('total_images', 0),
                    'tables': data.get('total_tables', 0),
                    'avg_words_per_page': round(data.get('total_words', 0) / max(1, data.get('page_count', 1)), 2)
                }
            elif data.get('file_type') == 'DOCX':
                summary['key_metrics'] = {
                    'paragraphs': len(data.get('paragraphs', [])),
                    'tables': len(data.get('tables', [])),
                    'words': data.get('total_words', 0)
                }
            elif data.get('file_type') == 'PPTX':
                summary['key_metrics'] = {
                    'slides': data.get('total_slides', 0),
                    'words': data.get('total_words', 0),
                    'avg_words_per_slide': round(data.get('total_words', 0) / max(1, data.get('total_slides', 1)), 2)
                }
            elif data.get('file_type') == 'Excel':
                summary['key_metrics'] = {
                    'sheets': data.get('total_sheets', 0),
                    'rows': data.get('total_rows', 0),
                    'columns': data.get('total_columns', 0)
                }
                
        except Exception as e:
            logger.error(f"Error generating summary: {e}")
            summary['processing_success'] = False
            
        return summary
    
    def cleanup_temp_files(self):
        """Clean up temporary files and directories"""
        try:
            temp_dirs = ['temp_processing']
            for temp_dir in temp_dirs:
                if os.path.exists(temp_dir):
                    import shutil
                    shutil.rmtree(temp_dir)
                    os.makedirs(temp_dir, exist_ok=True)
            
            logger.info("Temporary files cleaned up")
            
        except Exception as e:
            logger.error(f"Error cleaning up temp files: {e}")
    
    def get_processing_report(self) -> Dict:
        """Generate a comprehensive processing report"""
        report = {
            'processing_statistics': self.processing_stats,
            'files_processed': len(self.extracted_data),
            'total_processing_time': sum(
                data.get('processing_time_seconds', 0) 
                for data in self.extracted_data.values()
            ),
            'file_breakdown': {},
            'performance_metrics': {
                'avg_processing_time': 0,
                'pages_per_second': 0,
                'images_per_second': 0,
                'tables_per_second': 0
            }
        }
        
        # Calculate file type breakdown
        file_types = {}
        for data in self.extracted_data.values():
            file_type = data.get('file_type', 'Unknown')
            file_types[file_type] = file_types.get(file_type, 0) + 1
        
        report['file_breakdown'] = file_types
        
        # Calculate performance metrics
        total_time = report['total_processing_time']
        if total_time > 0:
            report['performance_metrics']['avg_processing_time'] = round(
                total_time / max(1, len(self.extracted_data)), 2
            )
            report['performance_metrics']['pages_per_second'] = round(
                self.processing_stats['pages_processed'] / total_time, 2
            )
            report['performance_metrics']['images_per_second'] = round(
                self.processing_stats['images_extracted'] / total_time, 2
            )
            report['performance_metrics']['tables_per_second'] = round(
                self.processing_stats['tables_extracted'] / total_time, 2
            )
        
        return report
    
    def __del__(self):
        """Cleanup when object is destroyed"""
        try:
            if hasattr(self, 'conn') and self.conn:
                self.conn.close()
        except:
            pass

# Example usage and utility functions
def process_directory(directory_path: str, extractor: OptimizedDocumentExtractor) -> Dict:
    """Process all supported files in a directory"""
    file_paths = []
    
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_ext = os.path.splitext(file)[1].lower()
            
            if file_ext in extractor.supported_formats:
                file_paths.append(file_path)
    
    logger.info(f"Found {len(file_paths)} supported files in directory")
    
    if not file_paths:
        logger.warning("No supported files found in directory")
        return {}
    
    results = extractor.process_files_optimized(file_paths)
    extractor.extracted_data.update(results)
    
    return results

def main():
    """Main function demonstrating usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Optimized Document Extractor')
    parser.add_argument('input_path', help='Path to file or directory to process')
    parser.add_argument('--output', '-o', default='extracted_data.json', help='Output JSON file path')
    parser.add_argument('--workers', '-w', type=int, default=None, help='Number of worker threads')
    parser.add_argument('--chunk-size', '-c', type=int, default=50, help='Chunk size for processing')
    parser.add_argument('--memory-limit', '-m', type=int, default=1024, help='Memory limit in MB')
    parser.add_argument('--disable-cache', action='store_true', help='Disable caching')
    parser.add_argument('--report', '-r', action='store_true', help='Generate processing report')
    
    args = parser.parse_args()
    
    # Initialize extractor
    extractor = OptimizedDocumentExtractor(
        max_workers=args.workers,
        chunk_size=args.chunk_size,
        enable_caching=not args.disable_cache,
        memory_limit_mb=args.memory_limit
    )
    
    try:
        start_time = time.time()
        
        # Process input
        if os.path.isfile(args.input_path):
            logger.info(f"Processing single file: {args.input_path}")
            results = extractor.process_files_optimized([args.input_path])
            extractor.extracted_data.update(results)
        elif os.path.isdir(args.input_path):
            logger.info(f"Processing directory: {args.input_path}")
            results = process_directory(args.input_path, extractor)
        else:
            logger.error(f"Input path does not exist: {args.input_path}")
            return
        
        total_time = time.time() - start_time
        
        # Save results
        output_file = extractor.save_extracted_data(args.output)
        
        if output_file:
            logger.info(f"Processing completed in {total_time:.2f} seconds")
            logger.info(f"Results saved to: {output_file}")
            
            # Generate report if requested
            if args.report:
                report = extractor.get_processing_report()
                report_file = args.output.replace('.json', '_report.json')
                
                with open(report_file, 'w', encoding='utf-8') as f:
                    json.dump(report, f, indent=2, ensure_ascii=False)
                
                logger.info(f"Processing report saved to: {report_file}")
                
                # Print summary
                print("\n" + "="*50)
                print("PROCESSING SUMMARY")
                print("="*50)
                print(f"Files processed: {report['files_processed']}")
                print(f"Total time: {report['total_processing_time']:.2f} seconds")
                print(f"Pages processed: {report['processing_statistics']['pages_processed']}")
                print(f"Images extracted: {report['processing_statistics']['images_extracted']}")
                print(f"Tables extracted: {report['processing_statistics']['tables_extracted']}")
                print(f"Average processing time: {report['performance_metrics']['avg_processing_time']:.2f} seconds/file")
                print("="*50)
        
        # Cleanup
        extractor.cleanup_temp_files()
        
    except KeyboardInterrupt:
        logger.info("Processing interrupted by user")
    except Exception as e:
        logger.error(f"Error during processing: {e}")
    finally:
        # Ensure cleanup
        try:
            extractor.cleanup_temp_files()
        except:
            pass

if __name__ == "__main__":
    # Add missing import for io
    import io
    main()
