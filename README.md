# Universal File Information Extractor

It is Streamlit web app through which users can upload various kinds of documents (PDF, Word, PowerPoint, Excel) and extract structured info such as text, images, tables, and summaries of metadata. It also provides downloadable summaries in Excel and full content exports in JSON.

## Features

- Support of multiple files: `.pdf`, `.docx`, `.pptx`, `.xlsx`
- Displays word count, character count, table count, and picture count per page/sheet/slide
- Pulls and displays:
  - PDF: text, tables, images
  - Word: entire text
  - PowerPoint: slide-by-slide text
  - Excel: sheet-by-sheet data tables
- Generates:
  - Excel summaries
  - Entire JSON content dumps
- Interactive page, slide, or sheet navigation is supported
