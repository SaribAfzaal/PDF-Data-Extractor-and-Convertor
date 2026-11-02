from pathlib import Path

import pypdfium2
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUT = Path(__file__).parent
SCAN_DPI = 200

EMPLOYEES_HEADER = ["ID", "Name", "Department", "City", "Salary (PKR)"]
EMPLOYEES = [
    [1001, "Ahmed Raza", "Engineering", "Lahore", "185,000"],
    [1002, "Fatima Noor", "Finance", "Karachi", "150,000"],
    [1003, "Hassan Ali", "Marketing", "Islamabad", "135,000"],
    [1004, "Ayesha Siddiqui", "Engineering", "Rawalpindi", "190,000"],
    [1005, "Usman Tariq", "Operations", "Faisalabad", "120,000"],
    [1006, "Zainab Hussain", "HR", "Multan", "110,000"],
    [1007, "Bilal Ahmad", "Engineering", "Peshawar", "175,000"],
    [1008, "Sana Malik", "Finance", "Lahore", "145,000"],
    [1009, "Imran Khan", "Sales", "Quetta", "125,000"],
    [1010, "Mariam Javed", "Marketing", "Karachi", "140,000"],
    [1011, "Hamza Sheikh", "Operations", "Sialkot", "118,000"],
    [1012, "Noor Fatima", "HR", "Islamabad", "105,000"],
]

SALES_HEADER = ["Region", "Product", "Units Sold", "Revenue (PKR)"]
SALES = [
    ["Punjab", "Solar Panel", 420, "8,400,000"],
    ["Sindh", "Inverter", 310, "6,200,000"],
    ["Khyber Pakhtunkhwa", "Battery Pack", 185, "4,625,000"],
    ["Balochistan", "Solar Panel", 95, "1,900,000"],
    ["Islamabad Capital Territory", "Inverter", 150, "3,000,000"],
]


def styled_table(header, rows):
    table = Table([header] + rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return table


def build_text_pdf(path):
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Employee Directory", styles["Title"]),
        Paragraph("Prepared by the Human Resources office in Lahore for the 2026 annual review.", styles["Normal"]),
        Spacer(1, 16),
        styled_table(EMPLOYEES_HEADER, EMPLOYEES[:7]),
        PageBreak(),
        Paragraph("Employee Directory (continued)", styles["Heading2"]),
        Spacer(1, 12),
        styled_table(EMPLOYEES_HEADER, EMPLOYEES[7:]),
        PageBreak(),
        Paragraph("Regional Sales Summary", styles["Title"]),
        Paragraph("Sales figures for the last quarter, grouped by region and product.", styles["Normal"]),
        Spacer(1, 16),
        styled_table(SALES_HEADER, SALES),
    ]
    SimpleDocTemplate(str(path), pagesize=A4).build(story)


def build_scanned_pdf(source, path):
    document = pypdfium2.PdfDocument(str(source))
    images = [document[i].render(scale=SCAN_DPI / 72).to_pil().convert("L").convert("RGB") for i in range(len(document))]
    document.close()
    Image.preinit()
    images[0].save(path, save_all=True, append_images=images[1:], resolution=SCAN_DPI)


if __name__ == "__main__":
    text_pdf = OUT / "sample_input.pdf"
    build_text_pdf(text_pdf)
    build_scanned_pdf(text_pdf, OUT / "sample_scanned.pdf")
    print("Created sample_input.pdf and sample_scanned.pdf")
