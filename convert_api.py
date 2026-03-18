from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PyPDF2 import PdfMerger, PdfReader

router = APIRouter()

# Register Calibri Fonts
pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
pdfmetrics.registerFont(TTFont('Calibri-Bold', 'Calibri-Bold.ttf'))


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

    # Save uploaded file
    input_path = f"/tmp/{file.filename}"
    with open(input_path, "wb") as f:
        f.write(await file.read())

    # Convert to PDF using LibreOffice
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
    # CREATE DETAILS PAGE (A4 FIXED)
    # =========================
    details_pdf = "/tmp/details_page.pdf"

    c = canvas.Canvas(details_pdf, pagesize=A4)
    width, height = A4

    # Title (Centered)
    c.setFont("Calibri-Bold", 12)
    c.drawCentredString(width / 2, height - 80, "Document Control Information")

    # Line below title
    c.setLineWidth(1)
    c.line(100, height - 90, width - 100, height - 90)

    # Table Data
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

        ("FONTNAME", (0,0), (-1,0), "Calibri-Bold"),
        ("FONTNAME", (0,1), (-1,-1), "Calibri"),

        ("FONTSIZE", (0,0), (-1,0), 12),
        ("FONTSIZE", (0,1), (-1,-1), 11),

        ("ALIGN", (0,0), (-1,-1), "LEFT"),

        ("GRID", (0,0), (-1,-1), 1, colors.black),

        ("BACKGROUND", (0,0), (-1,-1), colors.white),

        # Padding (important)
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),

    ]))

    # Center Table
    table_width, table_height = table.wrap(0, 0)
    x = (width - table_width) / 2
    y = height - 250

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

    return FileResponse(
        final_pdf,
        media_type="application/pdf",
        filename="document.pdf"
    )
