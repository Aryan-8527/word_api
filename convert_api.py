from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PyPDF2 import PdfMerger, PdfReader

router = APIRouter()

# =========================
# FONT SETUP (SAFE)
# =========================
try:
    base_path = os.getcwd()

    font_path = os.path.join(base_path, "calibri.ttf")
    bold_font_path = os.path.join(base_path, "calibri-bold.ttf")

    pdfmetrics.registerFont(TTFont('Calibri', font_path))
    pdfmetrics.registerFont(TTFont('Calibri-Bold', bold_font_path))

    FONT = "Calibri"
    FONT_BOLD = "Calibri-Bold"

    print("✅ Calibri font loaded")

except Exception as e:
    print("❌ Font load failed:", e)

    FONT = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"


@router.post("/convert-to-pdf")
async def convert_to_pdf(
    file: UploadFile = File(...),
    document_code: str = Form(...),
    client_name: str = Form(...),
    department: str = Form(...),
    document_type: str = Form(...),
    purpose: str = Form(...),
    created_on: str = Form(...),
    created_by: str = Form(...)
):

    # =========================
    # SAVE FILE
    # =========================
    input_path = f"/tmp/{file.filename}"
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # =========================
    # CONVERT TO PDF
    # =========================
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        input_path,
        "--outdir",
        "/tmp"
    ], check=True)

    pdf_file = "/tmp/" + os.path.splitext(file.filename)[0] + ".pdf"

    reader = PdfReader(pdf_file)

    # =========================
    # DETECT ORIENTATION
    # =========================
    first_page = reader.pages[0]

    width_pdf = float(first_page.mediabox.width)
    height_pdf = float(first_page.mediabox.height)

    is_landscape = width_pdf > height_pdf

    # =========================
    # CREATE DETAILS PAGE
    # =========================
    details_pdf = "/tmp/details_page.pdf"

    page_size = landscape(A4) if is_landscape else A4

    c = canvas.Canvas(details_pdf, pagesize=page_size)
    width, height = page_size

    # =========================
    # TITLE
    # =========================
    c.setFont(FONT_BOLD, 12)
    c.drawCentredString(width / 2, height - 60, "Document Control Information")

    # Line below title
    c.setLineWidth(1)
    c.line(80, height - 70, width - 80, height - 70)

    # =========================
    # TABLE DATA
    # =========================
    data = [
        ["Document Number", document_code],
        ["Client Name", client_name],
        ["Department", department],
        ["Document Type", document_type],
        ["Purpose", purpose],
        ["Created On", created_on],
        ["Created By", created_by],
    ]

    table = Table(data, colWidths=[200, 300])

    table.setStyle(TableStyle([

        ("FONTNAME", (0,0), (-1,-1), FONT),

        ("FONTSIZE", (0,0), (-1,-1), 11),

        ("ALIGN", (0,0), (-1,-1), "LEFT"),

        ("GRID", (0,0), (-1,-1), 1, colors.black),

        ("BACKGROUND", (0,0), (-1,-1), colors.white),

        # Padding
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),

    ]))

    # =========================
    # CENTER TABLE
    # =========================
    table_width, table_height = table.wrap(0, 0)

    x = (width - table_width) / 2
    y = height - 150

    table.drawOn(c, x, y)

    c.save()

    # =========================
    # MERGE PDF
    # =========================
    merger = PdfMerger()

    total_pages = len(reader.pages)

    # First page
    merger.append(pdf_file, pages=(0, 1))

    # Insert Details Page
    merger.append(details_pdf)

    # Remaining pages
    if total_pages > 1:
        merger.append(pdf_file, pages=(1, total_pages))

    final_pdf = "/tmp/final_output.pdf"

    merger.write(final_pdf)
    merger.close()

    # =========================
    # RETURN FILE
    # =========================
    return FileResponse(
        final_pdf,
        media_type="application/pdf",
        filename="document.pdf"
    )
