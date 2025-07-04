import fitz  # PyMuPDF
import os
import json
import pandas as pd
from PIL import Image
import pytesseract
from collections import Counter
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
import io
warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_extraction.log'),
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

class OptimizedPDFExtractor:
    def __init__(self, max_workers=None, chunk_size=50, enable_caching=True, memory_limit_mb=1024):
        self.extracted_data = {}
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
            self.db_path = 'pdf_extraction_cache.db'
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
    
    def process_pdf_files(self, pdf_paths: List[str]) -> Dict:
        """Process multiple PDF files with optimized resource management"""
        results = {}
        
        logger.info(f"Processing {len(pdf_paths)} PDF file(s) with optimized methods...")
        
        for i, pdf_path in enumerate(pdf_paths, 1):
            # Verify it's a PDF file
            if not pdf_path.lower().endswith('.pdf'):
                logger.warning(f"Skipping non-PDF file: {pdf_path}")
                continue
            
            if not os.path.exists(pdf_path):
                logger.error(f"File not found: {pdf_path}")
                continue
            
            logger.info(f"[{i}/{len(pdf_paths)}] Processing: {pdf_path}")
            
            # Check file size and warn for very large files
            file_size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
            if file_size_mb > 500:
                logger.warning(f"Very large file detected ({file_size_mb:.1f}MB). Processing may take significant time.")
            
            start_time = time.time()
            
            try:
                result = self.extract_pdf_chunked(pdf_path)
                processing_time = time.time() - start_time
                
                if result:
                    result['processing_time_seconds'] = round(processing_time, 2)
                    results[pdf_path] = result
                    logger.info(f"Successfully processed: {pdf_path} ({processing_time:.2f}s)")
                else:
                    logger.error(f"Failed to process: {pdf_path}")
                    
            except Exception as e:
                logger.error(f"Error processing {pdf_path}: {e}")
                continue
        
        return results
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract only text from PDF (lightweight method)"""
        try:
            doc = fitz.open(pdf_path)
            full_text = ""
            
            for page in doc:
                full_text += page.get_text() + "\n"
            
            doc.close()
            return full_text.strip()
            
        except Exception as e:
            logger.error(f"Error extracting text from {pdf_path}: {e}")
            return ""
    
    def extract_images_from_pdf(self, pdf_path: str) -> List[str]:
        """Extract only images from PDF (lightweight method)"""
        image_files = []
        
        try:
            doc = fitz.open(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                images = self.extract_images_optimized(page, page_num + 1, pdf_path)
                image_files.extend([img['filename'] for img in images])
            
            doc.close()
            return image_files
            
        except Exception as e:
            logger.error(f"Error extracting images from {pdf_path}: {e}")
            return []
    
    def save_extracted_data(self, output_path: str = "extracted_pdf_data.json"):
        """Save all extracted data to JSON file"""
        try:
            # Convert results to JSON-serializable format
            json_data = {}
            
            for file_path, data in self.extracted_data.items():
                # Create a simplified version for JSON export
                simplified_data = {
                    'filename': data.get('filename', ''),
                    'file_type': 'PDF',
                    'processing_time_seconds': data.get('processing_time_seconds', 0),
                    'file_size_mb': data.get('file_size_mb', 0),
                    'page_count': data.get('page_count', 0),
                    'total_words': data.get('total_words', 0),
                    'total_characters': data.get('total_characters', 0),
                    'total_images': data.get('total_images', 0),
                    'total_tables': data.get('total_tables', 0),
                    'fonts_used': data.get('fonts_used', [])[:10],  # Limit fonts
                    'metadata': data.get('metadata', {}),
                    'summary': self.generate_pdf_summary(data)
                }
                
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
    
    def generate_pdf_summary(self, data: Dict) -> Dict:
        """Generate a summary of extracted PDF data"""
        summary = {
            'processing_success': True,
            'key_metrics': {}
        }
        
        try:
            summary['key_metrics'] = {
                'pages': data.get('page_count', 0),
                'words': data.get('total_words', 0),
                'characters': data.get('total_characters', 0),
                'images': data.get('total_images', 0),
                'tables': data.get('total_tables', 0),
                'fonts': len(data.get('fonts_used', [])),
                'avg_words_per_page': round(data.get('total_words', 0) / max(1, data.get('page_count', 1)), 2),
                'avg_chars_per_page': round(data.get('total_characters', 0) / max(1, data.get('page_count', 1)), 2)
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
            'total_pages': sum(
                data.get('page_count', 0) 
                for data in self.extracted_data.values()
            ),
            'total_words': sum(
                data.get('total_words', 0) 
                for data in self.extracted_data.values()
            ),
            'total_images': sum(
                data.get('total_images', 0) 
                for data in self.extracted_data.values()
            ),
            'total_tables': sum(
                data.get('total_tables', 0) 
                for data in self.extracted_data.values()
            ),
            'performance_metrics': {
                'avg_processing_time': 0,
                'pages_per_second': 0,
                'images_per_second': 0,
                'tables_per_second': 0
            }
        }
        
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

# Utility functions
# Completion of the OptimizedPDFExtractor code

# Complete the process_pdf_directory function
def process_pdf_directory(directory_path: str, extractor: OptimizedPDFExtractor) -> Dict:
    """Process all PDF files in a directory"""
    pdf_paths = []
    
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_paths.append(os.path.join(root, file))
    
    logger.info(f"Found {len(pdf_paths)} PDF files in directory")
    
    if not pdf_paths:
        logger.warning("No PDF files found in directory")
        return {}
    
    # Process all PDFs
    results = extractor.process_pdf_files(pdf_paths)
    extractor.extracted_data.update(results)
    
    return results

def batch_process_pdfs(pdf_list: List[str], output_dir: str = "batch_output", 
                      max_workers: int = 4, chunk_size: int = 50) -> Dict:
    """Process multiple PDFs in batch with optimized settings"""
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize extractor with batch-optimized settings
    extractor = OptimizedPDFExtractor(
        max_workers=max_workers,
        chunk_size=chunk_size,
        enable_caching=True,
        memory_limit_mb=2048  # Higher limit for batch processing
    )
    
    try:
        # Process all PDFs
        results = extractor.process_pdf_files(pdf_list)
        extractor.extracted_data.update(results)
        
        # Save results
        output_file = os.path.join(output_dir, "batch_extraction_results.json")
        extractor.save_extracted_data(output_file)
        
        # Generate and save processing report
        report = extractor.get_processing_report()
        report_file = os.path.join(output_dir, "processing_report.json")
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Batch processing complete. Results saved to: {output_dir}")
        logger.info(f"Processed {len(results)} files successfully")
        
        return {
            'results': results,
            'report': report,
            'output_directory': output_dir
        }
        
    except Exception as e:
        logger.error(f"Error in batch processing: {e}")
        return {}
    
    finally:
        extractor.cleanup_temp_files()

def extract_text_only(pdf_path: str) -> Dict:
    """Quick text-only extraction for lightweight processing"""
    try:
        extractor = OptimizedPDFExtractor(max_workers=1, chunk_size=100, enable_caching=False)
        text = extractor.extract_text_from_pdf(pdf_path)
        
        return {
            'filename': os.path.basename(pdf_path),
            'text': text,
            'word_count': len(text.split()) if text else 0,
            'character_count': len(text),
            'extraction_method': 'text_only'
        }
        
    except Exception as e:
        logger.error(f"Error in text-only extraction: {e}")
        return {}

def extract_images_only(pdf_path: str, output_dir: str = "extracted_images") -> Dict:
    """Quick image-only extraction"""
    try:
        extractor = OptimizedPDFExtractor(max_workers=2, enable_caching=False)
        images = extractor.extract_images_from_pdf(pdf_path)
        
        return {
            'filename': os.path.basename(pdf_path),
            'images': images,
            'image_count': len(images),
            'extraction_method': 'images_only'
        }
        
    except Exception as e:
        logger.error(f"Error in image-only extraction: {e}")
        return {}

def analyze_pdf_structure(pdf_path: str) -> Dict:
    """Analyze PDF structure without full extraction"""
    try:
        doc = fitz.open(pdf_path)
        
        analysis = {
            'filename': os.path.basename(pdf_path),
            'page_count': len(doc),
            'file_size_mb': round(os.path.getsize(pdf_path) / (1024*1024), 2),
            'metadata': doc.metadata or {},
            'page_sizes': [],
            'total_images': 0,
            'estimated_tables': 0,
            'font_analysis': {},
            'text_density': []
        }
        
        # Analyze first 10 pages for structure
        sample_pages = min(10, len(doc))
        font_counter = Counter()
        
        for page_num in range(sample_pages):
            page = doc[page_num]
            
            # Page size
            analysis['page_sizes'].append({
                'page': page_num + 1,
                'width': round(page.rect.width, 2),
                'height': round(page.rect.height, 2),
                'rotation': page.rotation
            })
            
            # Count images
            analysis['total_images'] += len(page.get_images())
            
            # Analyze text and fonts
            text_dict = page.get_text("dict")
            page_text = ""
            
            for block in text_dict.get("blocks", []):
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span.get("text", "")
                            page_text += text
                            font_counter[span.get('font', 'Unknown')] += 1
            
            # Text density (words per page)
            word_count = len(page_text.split()) if page_text else 0
            analysis['text_density'].append({
                'page': page_num + 1,
                'words': word_count,
                'characters': len(page_text)
            })
        
        # Estimate total based on sample
        if sample_pages > 0:
            avg_images_per_page = analysis['total_images'] / sample_pages
            analysis['estimated_total_images'] = round(avg_images_per_page * len(doc))
            
            avg_words_per_page = sum(p['words'] for p in analysis['text_density']) / sample_pages
            analysis['estimated_total_words'] = round(avg_words_per_page * len(doc))
        
        # Font analysis
        analysis['font_analysis'] = {
            'unique_fonts': len(font_counter),
            'most_common_fonts': font_counter.most_common(5)
        }
        
        doc.close()
        
        logger.info(f"PDF structure analysis complete: {analysis['filename']}")
        return analysis
        
    except Exception as e:
        logger.error(f"Error analyzing PDF structure: {e}")
        return {}

def main():
    """Main function for command-line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Optimized PDF Data Extractor")
    parser.add_argument("input", help="PDF file or directory path")
    parser.add_argument("-o", "--output", default="extracted_data", 
                       help="Output directory (default: extracted_data)")
    parser.add_argument("-w", "--workers", type=int, default=4,
                       help="Number of worker threads (default: 4)")
    parser.add_argument("-c", "--chunk-size", type=int, default=50,
                       help="Pages per chunk (default: 50)")
    parser.add_argument("--text-only", action="store_true",
                       help="Extract text only (faster)")
    parser.add_argument("--images-only", action="store_true",
                       help="Extract images only")
    parser.add_argument("--analyze", action="store_true",
                       help="Analyze PDF structure only")
    parser.add_argument("--memory-limit", type=int, default=1024,
                       help="Memory limit in MB (default: 1024)")
    parser.add_argument("--disable-cache", action="store_true",
                       help="Disable caching")
    
    args = parser.parse_args()
    
    # Create output directory
    os.makedirs(args.output, exist_ok=True)
    
    start_time = time.time()
    
    try:
        if args.analyze:
            # Structure analysis only
            if os.path.isfile(args.input):
                result = analyze_pdf_structure(args.input)
                output_file = os.path.join(args.output, "structure_analysis.json")
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                print(f"Structure analysis saved to: {output_file}")
            else:
                print("Structure analysis only works with single PDF files")
                
        elif args.text_only:
            # Text-only extraction
            if os.path.isfile(args.input):
                result = extract_text_only(args.input)
                output_file = os.path.join(args.output, "text_extraction.json")
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                print(f"Text extraction saved to: {output_file}")
            else:
                print("Text-only extraction only works with single PDF files")
                
        elif args.images_only:
            # Images-only extraction
            if os.path.isfile(args.input):
                result = extract_images_only(args.input, args.output)
                output_file = os.path.join(args.output, "image_extraction.json")
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                print(f"Image extraction saved to: {output_file}")
            else:
                print("Image-only extraction only works with single PDF files")
                
        else:
            # Full extraction
            extractor = OptimizedPDFExtractor(
                max_workers=args.workers,
                chunk_size=args.chunk_size,
                enable_caching=not args.disable_cache,
                memory_limit_mb=args.memory_limit
            )
            
            if os.path.isfile(args.input):
                # Single file
                result = extractor.extract_pdf_chunked(args.input)
                if result:
                    extractor.extracted_data[args.input] = result
                    
            elif os.path.isdir(args.input):
                # Directory
                process_pdf_directory(args.input, extractor)
                
            else:
                print(f"Error: {args.input} is not a valid file or directory")
                return
            
            # Save results
            output_file = os.path.join(args.output, "extraction_results.json")
            extractor.save_extracted_data(output_file)
            
            # Generate report
            report = extractor.get_processing_report()
            report_file = os.path.join(args.output, "processing_report.json")
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)
            
            print(f"Extraction complete. Results saved to: {args.output}")
            print(f"Files processed: {len(extractor.extracted_data)}")
            
            # Cleanup
            extractor.cleanup_temp_files()
    
    except Exception as e:
        logger.error(f"Error in main processing: {e}")
        print(f"Error: {e}")
        return 1
    
    finally:
        total_time = time.time() - start_time
        print(f"Total processing time: {total_time:.2f} seconds")
    
    return 0

if __name__ == "__main__":
    exit(main())
