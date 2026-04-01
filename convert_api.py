from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os

from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import getSampleStyleSheet

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
    # GET PAGE SIZE
    # =========================
    first_page = reader.pages[0]
    width = float(first_page.mediabox.width)
    height = float(first_page.mediabox.height)

    # =========================
    # CREATE DETAILS PAGE
    # =========================
    details_pdf = "/tmp/details_page.pdf"
    c = canvas.Canvas(details_pdf, pagesize=(width, height))

    # =========================
    # WATERMARK (BOTTOM CENTER)
    # =========================
    logo_path = "logo.png"

    if os.path.exists(logo_path):
        img = ImageReader(logo_path)
        img_w, img_h = img.getSize()

        ratio = img_w / img_h

        draw_width = width * 0.5
        draw_height = draw_width / ratio

        x = (width - draw_width) / 2
        y = 80

        c.saveState()
        c.setFillAlpha(0.08)
        c.drawImage(img, x, y, width=draw_width, height=draw_height, mask='auto')
        c.restoreState()

    # =========================
    # TITLE
    # =========================
    c.setFont("Helvetica-Bold", 16)
    c.drawString(80, height - 80, "Document Control Information")

    c.setLineWidth(1)
    c.line(80, height - 90, width - 80, height - 90)

    # =========================
    # STYLE (FOR TEXT WRAP)
    # =========================
    styles = getSampleStyleSheet()
    normal_style = styles["Normal"]
    normal_style.wordWrap = 'CJK'   # ✅ important for wrapping
    normal_style.leading = 14
    normal_style.fontSize = 10

    # =========================
    # TABLE DATA (WRAP ENABLED)
    # =========================
    data = [
        ["Document Number", Paragraph(document_code, normal_style)],
        ["Client Name", Paragraph(client_name, normal_style)],
        ["Department", Paragraph(department, normal_style)],
        ["Document Type", Paragraph(document_type, normal_style)],
        ["Purpose", Paragraph(purpose, normal_style)],  # ✅ wrap fix
        ["Created On", Paragraph(created_on, normal_style)],
        ["Created By", Paragraph(created_by, normal_style)],
    ]

    table_width = width - 160

    col_widths = [
        table_width * 0.35,
        table_width * 0.65
    ]

    table = Table(data, colWidths=col_widths)

    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),

        ("FONTSIZE", (0,0), (-1,-1), 11),

        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),

        ("GRID", (0,0), (-1,-1), 1, colors.black),

        ("BACKGROUND", (0,0), (-1,-1), colors.white),

        ("LEFTPADDING", (0,0), (-1,-1), 12),
        ("RIGHTPADDING", (0,0), (-1,-1), 12),
        ("TOPPADDING", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ]))

    # =========================
    # DRAW TABLE
    # =========================
    table.wrapOn(c, width, height)
    table.drawOn(c, 80, height - 350)

    c.save()

    # =========================
    # MERGE PDF
    # =========================
    merger = PdfMerger()

    total_pages = len(reader.pages)

    merger.append(pdf_file, pages=(0, 1))   # first page
    merger.append(details_pdf)              # second page

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
