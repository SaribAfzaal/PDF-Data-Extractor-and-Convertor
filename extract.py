import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from statistics import median

import pdfplumber
import pypdfium2
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MIN_TEXT_CHARS = 20
HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUMBER_RE = re.compile(r"^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")


class ExtractionError(Exception):
    pass


@dataclass
class PageResult:
    number: int
    method: str
    tables: list = field(default_factory=list)
    text: list = field(default_factory=list)


def clean_cell(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def to_value(text):
    if NUMBER_RE.match(text) and not (len(text) > 1 and text[0] == "0" and text[1].isdigit()):
        number = text.replace(",", "")
        return float(number) if "." in number else int(number)
    return text


def clean_table(rows):
    cleaned = []
    for row in rows:
        cells = [clean_cell(c) for c in row]
        if not any(cells):
            continue
        width = len(cells)
        if cleaned and not cells[0] and sum(1 for c in cells if c) <= max(1, width // 2):
            previous = cleaned[-1]
            for i, cell in enumerate(cells):
                if cell:
                    previous[i] = f"{previous[i]} {cell}".strip()
            continue
        cleaned.append(cells)
    if not cleaned:
        return []
    width = max(len(r) for r in cleaned)
    cleaned = [r + [""] * (width - len(r)) for r in cleaned]
    keep = [i for i in range(width) if any(r[i] for r in cleaned)]
    return [[r[i] for i in keep] for r in cleaned]


def clean_lines(text):
    lines = (re.sub(r"\s+", " ", line).strip() for line in (text or "").splitlines())
    return [line for line in lines if line]


def extract_text_page(page):
    finder = page.find_tables()
    boxes = [t.bbox for t in finder]
    tables = [clean_table(t.extract()) for t in finder]
    tables = [t for t in tables if t]

    def outside_tables(obj):
        if "x0" not in obj or "top" not in obj:
            return True
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in boxes)

    text = clean_lines(page.filter(outside_tables).extract_text() if boxes else page.extract_text())
    return tables, text


def group_rows(words):
    height = median(w["height"] for w in words)
    rows = []
    for word in sorted(words, key=lambda w: w["top"] + w["height"] / 2):
        center = word["top"] + word["height"] / 2
        if rows and abs(rows[-1]["center"] - center) <= height * 0.6:
            rows[-1]["words"].append(word)
            n = len(rows[-1]["words"])
            rows[-1]["center"] += (center - rows[-1]["center"]) / n
        else:
            rows.append({"center": center, "words": [word]})
    return rows, height


def split_cells(words, height):
    cells = []
    for word in sorted(words, key=lambda w: w["left"]):
        if cells and word["left"] - cells[-1]["right"] <= height:
            cells[-1]["text"] += " " + word["text"]
            cells[-1]["right"] = word["left"] + word["width"]
        else:
            cells.append({"text": word["text"], "left": word["left"], "right": word["left"] + word["width"]})
    return cells


def build_columns(table_rows):
    spans = sorted((c["left"], c["right"]) for row in table_rows for c in row)
    columns = []
    for left, right in spans:
        if columns and left <= columns[-1][1]:
            columns[-1][1] = max(columns[-1][1], right)
        else:
            columns.append([left, right])
    return columns


def assemble_ocr_table(table_rows):
    columns = build_columns(table_rows)
    grid = []
    for row in table_rows:
        line = [""] * len(columns)
        for cell in row:
            index = next(i for i, (left, right) in enumerate(columns) if cell["left"] <= right and cell["right"] >= left)
            line[index] = f"{line[index]} {cell['text']}".strip()
        grid.append(line)
    return clean_table(grid)


def extract_ocr_page(image, lang):
    import pytesseract

    try:
        data = pytesseract.image_to_data(image, lang=lang, output_type=pytesseract.Output.DICT)
    except pytesseract.TesseractNotFoundError:
        raise ExtractionError(
            "Tesseract OCR is not installed or not on PATH. Install it from "
            "https://github.com/UB-Mannheim/tesseract/wiki (Windows) or via your package manager."
        )
    except pytesseract.TesseractError as exc:
        raise ExtractionError(f"Tesseract failed: {exc}")

    words = [
        {"text": data["text"][i].strip(), "left": data["left"][i], "top": data["top"][i],
         "width": data["width"][i], "height": data["height"][i]}
        for i in range(len(data["text"]))
        if data["text"][i].strip() and float(data["conf"][i]) > 0
    ]
    if not words:
        return [], []

    rows, height = group_rows(words)
    split = [split_cells(r["words"], height) for r in rows]

    tables, text, block = [], [], []

    def flush():
        if len(block) >= 2:
            table = assemble_ocr_table(block)
            if table:
                tables.append(table)
        else:
            for row in block:
                text.append(" ".join(c["text"] for c in row))
        block.clear()

    for row in split:
        if len(row) >= 2:
            block.append(row)
        else:
            flush()
            text.append(row[0]["text"])
    flush()
    return tables, clean_lines("\n".join(text))


def render_page(pdf_path, index, dpi):
    document = pypdfium2.PdfDocument(str(pdf_path))
    try:
        return document[index].render(scale=dpi / 72).to_pil()
    finally:
        document.close()


def extract_pdf(pdf_path, dpi=300, lang="eng", force_ocr=False):
    results = []
    warnings = []
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception as exc:
        raise ExtractionError(f"Cannot open PDF (corrupted or encrypted?): {exc}")

    with pdf:
        if not pdf.pages:
            raise ExtractionError("PDF contains no pages.")
        for index, page in enumerate(pdf.pages):
            number = index + 1
            try:
                raw = page.extract_text() or ""
                if not force_ocr and len(raw.strip()) >= MIN_TEXT_CHARS:
                    tables, text = extract_text_page(page)
                    results.append(PageResult(number, "text", tables, text))
                else:
                    tables, text = extract_ocr_page(render_page(pdf_path, index, dpi), lang)
                    results.append(PageResult(number, "ocr", tables, text))
            except ExtractionError:
                raise
            except Exception as exc:
                warnings.append(f"Page {number} skipped: {exc}")
    return results, warnings


def is_header_like(row):
    return all(c and not NUMBER_RE.match(c) for c in row)


def normalize_header(row):
    return tuple(c.lower() for c in row)


def group_tables(results):
    groups = []
    for result in results:
        for position, table in enumerate(result.tables, start=1):
            header = table[0]
            key = normalize_header(header)
            target = next((g for g in groups if g["key"] == key), None)
            if target:
                target["rows"].extend(r for r in table[1:] if normalize_header(r) != key)
            elif groups and not is_header_like(header) and len(header) == len(groups[-1]["header"]):
                groups[-1]["rows"].extend(table)
            else:
                groups.append({
                    "key": key,
                    "header": header,
                    "rows": list(table[1:]),
                    "name": f"Page {result.number}" + (f" Table {position}" if len(result.tables) > 1 else ""),
                })
    return groups


def unique_name(name, used):
    base = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Sheet"
    candidate, counter = base, 2
    while candidate.lower() in used:
        suffix = f" ({counter})"
        candidate = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate.lower())
    return candidate


def write_sheet(sheet, header, rows):
    sheet.append(header)
    for row in rows:
        sheet.append([to_value(c) if isinstance(c, str) else c for c in row])

    for cell in sheet[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        for cell in row:
            cell.border = BORDER
            if cell.row > 1:
                cell.alignment = Alignment(
                    vertical="top",
                    horizontal="right" if isinstance(cell.value, (int, float)) else "left",
                    wrap_text=True,
                )
    for index, column in enumerate(sheet.columns, start=1):
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in column)
        sheet.column_dimensions[get_column_letter(index)].width = min(max(longest + 2, 10), 60)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions


def write_workbook(results, output_path):
    workbook = Workbook()
    workbook.remove(workbook.active)
    used = set()

    for group in group_tables(results):
        sheet = workbook.create_sheet(unique_name(group["name"], used))
        write_sheet(sheet, group["header"], group["rows"])

    text_rows = [[r.number, line] for r in results for line in r.text]
    if text_rows:
        sheet = workbook.create_sheet(unique_name("Text", used))
        write_sheet(sheet, ["Page", "Text"], text_rows)
        sheet.column_dimensions["B"].width = 100

    if not workbook.sheetnames:
        raise ExtractionError("No text or tables could be extracted from this PDF.")

    summary = workbook.create_sheet(unique_name("Summary", used), 0)
    write_sheet(
        summary,
        ["Page", "Method", "Tables", "Text Lines"],
        [[r.number, r.method.upper(), len(r.tables), len(r.text)] for r in results],
    )
    workbook.save(output_path)


def parse_args(argv):
    parser = argparse.ArgumentParser(description="Extract tables and text from a PDF (text-based or scanned) into an Excel file.")
    parser.add_argument("input", type=Path, help="Path to the input PDF")
    parser.add_argument("output", type=Path, help="Path to the output .xlsx file")
    parser.add_argument("--lang", default="eng", help="Tesseract language code (default: eng)")
    parser.add_argument("--dpi", type=int, default=300, help="Rendering DPI for OCR (default: 300)")
    parser.add_argument("--force-ocr", action="store_true", help="Use OCR on every page")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if not args.input.is_file():
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        return 1
    if args.input.suffix.lower() != ".pdf":
        print("Error: input must be a .pdf file", file=sys.stderr)
        return 1
    if args.output.suffix.lower() != ".xlsx":
        print("Error: output must be a .xlsx file", file=sys.stderr)
        return 1

    try:
        results, warnings = extract_pdf(args.input, args.dpi, args.lang, args.force_ocr)
        if not results:
            raise ExtractionError("Every page failed to process.")
        write_workbook(results, args.output)
    except ExtractionError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except PermissionError:
        print(f"Error: cannot write {args.output}. Is it open in Excel?", file=sys.stderr)
        return 1

    for warning in warnings:
        print(f"Warning: {warning}", file=sys.stderr)
    methods = ", ".join(f"page {r.number}: {r.method}" for r in results)
    print(f"Saved {args.output} ({methods})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
