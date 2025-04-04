
# Table extractor


## 📌 Objective
This tool detects and extracts tables from system-generated PDFs **without using Tabula, Camelot, or image conversion**. The extracted tables are stored in an **Excel sheet**, preserving their structure even if they have **borders, no borders, or irregular shapes**.
## 🎥 Demo Video
- [Watch Demo Video](https://drive.google.com/drive/folders/1e9Om5RWH5P9HA03gN0gdcMkcQVEYHaxo?q=sharedwith:public%20parent:1e9Om5RWH5P9HA03gN0gdcMkcQVEYHaxo)

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
## ⚙️ How It Works
1. **Reads PDF**: The script uses `PyMuPDF` to open and read each PDF file.
2. **Extracts Words**: It extracts words along with their positions from each page.
3. **Identifies Table Structure**: Words are grouped into rows based on their y-coordinates.
4. **Formats Data**: The extracted text is structured into a table format.
5. **Exports to Excel**: The cleaned table data is saved into an Excel file using `openpyxl`.
6. **Handles Multiple Pages**: If a PDF has multiple pages, each page’s table is stored in a separate Excel sheet.
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

   
## ❌ Limitations
- **Corrupted PDFs**: The tool cannot process corrupted or unreadable PDFs. Example: `test6.pdf` is detected as corrupted.
- **Complex Formatting**: Extremely complex tables with merged cells may not be extracted accurately.
- **Handwritten PDFs**: Only system-generated PDFs are supported; handwritten or scanned tables are not processed correctly.
- **Multi-Language Support**: Currently optimized for English; other languages may not be parsed accurately.


 ## 🌍 GitHub Workflow
1. **Initialize a Git repository**:
   ```sh
   git init
   ```
2. **Add files to the repository**:
   ```sh
   git add .
   ```
3. **Commit changes**:
   ```sh
   git commit -m "Initial commit - Added PDF table extractor script"
   ```
4. **Create a new repository on GitHub** and copy the repository URL.

5. **Push code to GitHub**:
   ```sh
   git push -u origin main
   ```
6. **Update repository**:
   - If  make changes, repeat steps **2-6** to push updates.

