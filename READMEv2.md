# File to PDF Converter

A Python tool to **convert different file formats into PDF** using Microsoft Office COM automation and other libraries.  
Supports Excel, Word, PowerPoint, images, and already existing PDFs. Also allows merging PDFs into a single file.

---

## Features
- Convert:
  - **Excel (.xlsx)** → PDF (all sheets or a specific sheet)
  - **Word (.docx)** → PDF
  - **PowerPoint (.pptx)** → PDF
  - **Images (.png, .jpg, .jpeg)** → PDF
  - **Existing PDFs** (just copies/renames)
- Auto-detects landscape/portrait orientation for Excel sheets.
- Handles duplicate filenames by renaming automatically.
- Skip temporary files (Excel `~` files).
- Merge all PDFs in a directory into one.

---

## Requirements
- Windows with **Microsoft Office** installed
- Python 3.9+
- Dependencies:
  ```bash
  pip install pywin32 img2pdf PyPDF2

---

## Usage
- Windows with **Microsoft Office** installed
- Python 3.9+
- Dependencies:
  ```bash
  pip install pywin32 img2pdf PyPDF2
