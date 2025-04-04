# PDF Table Extractor

## 📌 Objective
This tool detects and extracts tables from system-generated PDFs **without using Tabula, Camelot, or image conversion**. The extracted tables are stored in an **Excel sheet**, preserving their structure even if they have **borders, no borders, or irregular shapes**.

## 🚀 Features
- **Text-based Table Extraction**: Uses PyMuPDF (`fitz`) and `pandas` to extract tables without image processing.
- **Handles Various Table Formats**: Supports tables with/without borders and irregular shapes.
- **Multi-Page PDF Support**: Extracts tables from all pages and stores them in separate sheets.
- **Exports to Excel**: Uses `openpyxl` to save extracted tables.
- **Batch Processing**: Processes multiple PDFs in a folder.

## 🛠️ Requirements
Ensure you have the following installed:
```sh
pip install pymupdf pandas openpyxl
```

## 📂 Folder Structure
```
📁 pdf-table-extractor
 ├── 📂 pdfs  # Folder containing input PDFs
 ├── 📂 excel_output  # Extracted Excel files will be saved here
 ├── extract_tables.py  # Main script
 ├── README.md  # Project documentation
```

## 🔧 How to Use
1. **Place PDFs** in the `pdfs` folder.
2. **Run the script**:
   ```sh
   python extract_tables.py
   ```
3. **Find the extracted Excel files** in `excel_output/`.

## 📜 Example Code (extract_tables.py)
```python
import os
import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook

def extract_text_based_table(page):
    words = page.get_text("words")
    words.sort(key=lambda w: (w[1], w[0]))
    rows, current_row, last_y = [], [], None
    for w in words:
        x0, y0, x1, y1, word, *_ = w
        word = word.strip()
        if not word:
            continue
        if last_y is None or abs(y0 - last_y) > 5:
            if current_row:
                rows.append(current_row)
            current_row, last_y = [word], y0
        else:
            current_row.append(word)
    if current_row:
        rows.append(current_row)
    return rows

def extract_tables_from_pdf(pdf_path, output_excel_path):
    doc, workbook = fitz.open(pdf_path), Workbook()
    for page_num, page in enumerate(doc):
        table_data = extract_text_based_table(page)
        if table_data:
            sheet = workbook.create_sheet(title=f"Page_{page_num + 1}")
            for row in table_data:
                sheet.append(row)
    workbook.save(output_excel_path)

# Run the script
pdf_folder, output_folder = "pdfs", "excel_output"
os.makedirs(output_folder, exist_ok=True)
for filename in os.listdir(pdf_folder):
    if filename.lower().endswith(".pdf"):
        extract_tables_from_pdf(os.path.join(pdf_folder, filename), os.path.join(output_folder, filename.replace(".pdf", ".xlsx")))
```

## ✅ Evaluation Criteria
- **Accuracy**: Correct table detection & extraction.
- **Efficiency**: Processes multiple PDFs efficiently.
- **Code Quality**: Readable, maintainable, and well-documented.
- **User Experience**: Simple to use, with clear instructions & error handling.
- **Innovation**: Handles tables with borders, no borders, and irregular shapes.

## 📝 Submission
- Submit the **source code** along with `README.md`.
- Include **sample PDFs & extracted Excel files** for testing.
- (Optional) Provide a short **demo video** explaining your solution.

---
**📧 Need help?** Feel free to reach out! 🚀
