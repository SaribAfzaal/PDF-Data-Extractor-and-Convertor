import io
import tempfile
from pathlib import Path

import pandas as pd
import pypdfium2
import streamlit as st

from extract import ExtractionError, extract_pdf, write_workbook

ROOT = Path(__file__).parent
SAMPLES = {
    "Sample: text-based PDF": ROOT / "sample_input.pdf",
    "Sample: scanned PDF (OCR)": ROOT / "sample_scanned.pdf",
}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="PDF Data Extractor", page_icon="📄", layout="wide")


@st.cache_data(show_spinner=False)
def process(pdf_bytes, force_ocr, dpi):
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "input.pdf"
        path.write_bytes(pdf_bytes)
        results, warnings = extract_pdf(path, dpi=dpi, force_ocr=force_ocr)
    if not results:
        raise ExtractionError("Every page failed to process.")
    buffer = io.BytesIO()
    write_workbook(results, buffer)
    return results, warnings, buffer.getvalue()


@st.cache_data(show_spinner=False)
def render_previews(pdf_bytes):
    document = pypdfium2.PdfDocument(pdf_bytes)
    try:
        return [document[i].render(scale=1.5).to_pil() for i in range(len(document))]
    finally:
        document.close()


def load_source():
    choice = st.sidebar.radio("Input", [*SAMPLES, "Upload your own PDF"])
    if choice in SAMPLES:
        path = SAMPLES[choice]
        return path.name, path.read_bytes()
    upload = st.sidebar.file_uploader("PDF file", type="pdf")
    return (upload.name, upload.getvalue()) if upload else (None, None)


st.title("PDF Data Extractor and Converter")
st.caption("Text PDFs are read with pdfplumber, scanned pages with Tesseract OCR. The method is detected per page.")

force_ocr = st.sidebar.checkbox("Force OCR on every page")
dpi = st.sidebar.select_slider("OCR resolution (DPI)", options=[150, 200, 300, 400], value=300)
name, data = load_source()

if data is None:
    st.info("Choose a sample or upload a PDF in the sidebar.")
    st.stop()

try:
    with st.spinner("Extracting..."):
        results, warnings, xlsx_bytes = process(data, force_ocr, dpi)
except ExtractionError as exc:
    st.error(str(exc))
    st.stop()

for warning in warnings:
    st.warning(warning)

total_tables = sum(len(r.tables) for r in results)
total_lines = sum(len(r.text) for r in results)
columns = st.columns(4)
columns[0].metric("Pages", len(results))
columns[1].metric("Tables found", total_tables)
columns[2].metric("Text lines", total_lines)
columns[3].metric("OCR pages", sum(r.method == "ocr" for r in results))

st.download_button(
    "Download Excel file",
    xlsx_bytes,
    file_name=f"{Path(name).stem}.xlsx",
    mime=XLSX_MIME,
    type="primary",
)

left, right = st.columns([1, 1])
with left:
    st.subheader("Source PDF")
    with st.container(height=700):
        for number, image in enumerate(render_previews(data), start=1):
            st.image(image, caption=f"Page {number}", width="stretch")
with right:
    st.subheader("Excel output")
    sheets = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=None)
    for tab, (sheet_name, frame) in zip(st.tabs(list(sheets)), sheets.items()):
        with tab:
            st.dataframe(frame, hide_index=True, width="stretch")
            st.caption(f"{len(frame)} rows × {len(frame.columns)} columns")
