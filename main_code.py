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

def extract_from_pptx(file):
    prs = Presentation(file)
    slides_data = []
    for slide in prs.slides:
        text = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
        slides_data.append({"text": "\n".join(text)})
    return slides_data

def extract_from_excel(file):
    xls = pd.ExcelFile(file)
    sheets_data = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        sheets_data[sheet_name] = df
    return sheets_data