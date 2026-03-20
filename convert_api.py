from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os

from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor

from PyPDF2 import PdfMerger, PdfReader

router = APIRouter()


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
    # ORIGINAL PAGE SIZE (KEY FIX)
    # =========================
    first_page = reader.pages[0]

    width = float(first_page.mediabox.width)
    height = float(first_page.mediabox.height)

    # 👉 USE ORIGINAL SIZE (NO A4 FORCE)
    page_size = (width, height)

    # =========================
    # CREATE DETAILS PAGE
    # =========================
    details_pdf = "/tmp/details_page.pdf"

    c = canvas.Canvas(details_pdf, pagesize=page_size)

    # =========================
    # TITLE
    # =========================
    c.setFont("Helvetica-Bold", 16)
    c.setFillColor(HexColor("#003366"))
    c.drawCentredString(width / 2, height , "Document Control Information")

    # Divider
    c.setLineWidth(1.5)
    c.line(80, height - 60, width - 80, height - 60)

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

    # 👉 Dynamic width (MAIN FIX)
    table_width_available = width - 120

    col1 = table_width_available * 0.35
    col2 = table_width_available * 0.65

    table = Table(data, colWidths=[col1, col2])

    table.setStyle(TableStyle([

        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),

        ("FONTSIZE", (0,0), (-1,-1), 11),

        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),

        ("GRID", (0,0), (-1,-1), 1, colors.black),

        ("BACKGROUND", (0,0), (-1,-1), colors.white),

        # Padding
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),

    ]))

    # =========================
    # CENTER TABLE (FIX)
    # =========================
    table_width, table_height = table.wrap(0, 0)

    x = (width - table_width) / 2
    y = height - 180

    table.drawOn(c, x, y)

    c.save()

    # =========================
    # MERGE PDF
    # =========================
    merger = PdfMerger()

    total_pages = len(reader.pages)

    # First page
    merger.append(pdf_file, pages=(0, 1))

    # Details page
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
