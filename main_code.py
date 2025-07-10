import fitz  # PyMuPDF
import pandas as pd
from docx import Document
from pptx import Presentation
from PIL import Image
from io import BytesIO

def extract_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    pages_info = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        images = []
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            images.append(Image.open(BytesIO(image_bytes)))
        tables = page.find_tables()
        tables_data = []
        for table in tables:
            try:
                df = pd.DataFrame(table.extract())
                tables_data.append(df)
            except:
                pass
        pages_info.append({"text": text, "images": images, "tables": tables_data})
    return pages_info

def extract_from_docx(file):
    doc = Document(file)
    text = "\n".join([para.text for para in doc.paragraphs])
    return {"text": text}

def extract_table_from_pptx_shape(table_shape):
    """Extract table data from a PowerPoint table shape"""
    try:
        table = table_shape.table
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                row_data.append(cell_text)
            data.append(row_data)
        
        if data:
            # Create DataFrame with first row as headers if it looks like headers
            df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(data)
            return df
        return None
    except Exception as e:
        return None

def extract_from_pptx(file):
    prs = Presentation(file)
    slides_data = []
    for slide in prs.slides:
        text = []
        tables = []
        images = []
        
        for shape in slide.shapes:
            # Extract text
            if hasattr(shape, "text"):
                if shape.text.strip():
                    text.append(shape.text)
            
            # Extract tables
            if shape.shape_type == 19:  # Table shape type
                table_df = extract_table_from_pptx_shape(shape)
                if table_df is not None:
                    tables.append(table_df)
            
            # Extract images (basic detection)
            if shape.shape_type == 13:  # Picture shape type
                try:
                    image = shape.image
                    image_bytes = image.blob
                    images.append(Image.open(BytesIO(image_bytes)))
                except:
                    pass
        
        slides_data.append({
            "text": "\n".join(text),
            "tables": tables,
            "images": images
        })
    return slides_data

def extract_from_excel(file):
    xls = pd.ExcelFile(file)
    sheets_data = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        sheets_data[sheet_name] = df
    return sheets_data
