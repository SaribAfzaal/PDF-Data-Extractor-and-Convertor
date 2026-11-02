# PDF Data Extractor and Converter

Extracts tables and text from PDFs (text-based or scanned) and writes them to a formatted Excel file.

- Text-based pages use `pdfplumber`; pages with almost no extractable text are OCR'd with `pytesseract`. Detection is automatic, per page.
- Tables with the same header (even across pages) are merged into one sheet; different tables get their own sheet (`Page N`).
- Non-table text goes to a `Text` sheet. A `Summary` sheet shows the method used per page.
- Output has bold white-on-blue headers, borders, frozen header row, filters, and auto-sized columns.

## Setup

```
pip install -r requirements.txt
```

OCR also needs the Tesseract engine ([Windows installer](https://github.com/UB-Mannheim/tesseract/wiki), `apt install tesseract-ocr`, or `brew install tesseract`). It is only required for scanned PDFs.

## Usage

```
python extract.py input.pdf output.xlsx
```

Options: `--force-ocr`, `--lang eng`, `--dpi 300`.

## Try it

```
python generate_sample.py
python extract.py sample_input.pdf sample_output.xlsx
python extract.py sample_scanned.pdf sample_scanned_output.xlsx
```

`sample_input.pdf` is a 3-page text PDF (employee table split across two pages, then a sales table). `sample_scanned.pdf` is the same content as images and requires Tesseract.

Corrupted, missing, or non-PDF inputs print an error and exit with code 1.
