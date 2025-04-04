# 📄 PDF Table Extractor

Extracts structured tables from system-generated PDFs into Excel — no OCR, no Tabula, no Camelot.

## ✅ Features

- Handles tables with and without borders
- Uses text layout for table reconstruction
- Exports to clean, well-formatted Excel sheets

## 🔧 Installation

```bash
pip install pdfplumber pymupdf openpyxl pandas
```
## 🚀 Usage

1. Update the file paths in the script:

```python
pdf_path = "test2.pdf"
output_excel = "output.xlsx"
```

2. Run the script
```python
python main.py
```


## 📁 Output

- Output is an Excel (.xlsx) file.
- Each table or reconstructed text layout is saved as a separate sheet.
- Layout is preserved with aligned cells, borders, and wrapped text.

## ❌ Limitations

- Only works on system-generated PDFs (not scanned images).
- Does not support merged cells or complex row spans.

## 📄 License

For educational and hackathon use only.
