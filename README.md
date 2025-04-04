
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
 ├── main.py  # Main script
 ├── README.md  # Project documentation
```

## 🔧 How to Use
1. **Place PDFs** in the `pdfs` folder.
2. **Run the script**:
   ```sh
   python main.py
   ```
3. **Find the extracted Excel files** in `excel_output/`.



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
