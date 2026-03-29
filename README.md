# 📄 PDF Metadata Extractor & Excel Updater

Automates downloading PDFs from Excel URLs, extracting metadata, and updating structured document information.  
Ideal for bulk document processing, data enrichment, and workflow automation.

---

## 🚀 Features

- 📥 Bulk download PDFs from Excel URLs  
- 📄 Extract PDF metadata using `pypdf`  
- 🗂 Automatically rename files based on update date  
- 📊 Update Excel with:
  - Create Date  
  - Update Date  
  - Document Date  
  - Effective Date  
  - File Name & File Path  
- 📅 Automatically calculate **Effective Date (Document Date + 1 day)**  
- ⚠️ Robust error handling for failed downloads or invalid metadata  
- 📁 Auto-create download directory if not present  

---

## 🛠 Tech Stack

- Python  
- pandas  
- requests  
- pypdf  
- datetime  

---

## 📂 Workflow

1. Read Excel file containing document URLs  
2. Download each file  
3. Extract PDF metadata (creation & modification date)  
4. Rename files using extracted metadata  
5. Update Excel with new structured fields  
6. Save updated Excel file  

---

## ⚙️ Configuration

Update these variables in the script before running:

```python
excel_path = "path_to_input_excel.xlsx"
column_name = "URL"
download_folder = "path_to_download_folder"
output_path = "path_to_output_excel.xlsx"
