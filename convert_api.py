from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
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

    input_path = f"/tmp/{file.filename}"

    with open(input_path, "wb") as f:
        f.write(await file.read())

    # Convert to PDF
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
    first_page = reader.pages[0]

    width = float(first_page.mediabox.width)
    height = float(first_page.mediabox.height)

    # Detect orientation
    is_landscape = width > height

    # Create Document Details Page
    details_pdf = "/tmp/details_page.pdf"

    page_size = landscape(A4) if is_landscape else A4

    c = canvas.Canvas(details_pdf, pagesize=page_size)

    width, height = page_size

    # Title
    c.setFont("Helvetica-Bold", 22)
    c.setFillColor(HexColor("#003366"))
    c.drawString(50, height-80, "Document Details")

    # Divider
    c.setLineWidth(2)
    c.line(50, height-90, width-50, height-90)

    c.setFont("Helvetica", 13)
    y = height - 150

    fields = [
        ("Document Code", document_code),
        ("Client Name", client_name),
        ("Department", department),
        ("Document Type", document_type),
        ("Purpose", purpose),
        ("Created On", created_on),
        ("Created By", created_by),
    ]

    for label, value in fields:

        c.setFillColor(HexColor("#444444"))
        c.drawString(70, y, f"{label}")

        c.setFillColor(HexColor("#000000"))
        c.drawString(250, y, f": {value}")

        y -= 40

    c.save()

    # Merge PDFs
    merger = PdfMerger()

    total_pages = len(reader.pages)

    merger.append(pdf_file, pages=(0,1))
    merger.append(details_pdf)

    if total_pages > 1:
        merger.append(pdf_file, pages=(1,total_pages))

    final_pdf = "/tmp/final_output.pdf"

    merger.write(final_pdf)
    merger.close()

    return FileResponse(
        final_pdf,
        media_type="application/pdf",
        filename="document.pdf"
    )
