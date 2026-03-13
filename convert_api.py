from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import FileResponse
import subprocess
import os
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger

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

    # Convert using LibreOffice
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        input_path,
        "--outdir",
        "/tmp"
    ])

    pdf_file = os.path.splitext(input_path)[0] + ".pdf"

    # Create details page
    details_pdf = "/tmp/details_page.pdf"

    c = canvas.Canvas(details_pdf)

    c.setFont("Helvetica-Bold", 18)
    c.drawString(200, 800, "Document Details")

    c.setFont("Helvetica", 12)

    y = 720

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
        c.drawString(100, y, f"{label}: {value}")
        y -= 30

    c.save()

    # Merge PDFs
    merger = PdfMerger()

    merger.append(pdf_file, pages=(0,1))  # cover page
    merger.append(details_pdf)            # new page
    merger.append(pdf_file, pages=(1,None))  # rest pages

    final_pdf = "/tmp/final_output.pdf"

    merger.write(final_pdf)
    merger.close()

    return FileResponse(
        final_pdf,
        media_type="application/pdf",
        filename="document.pdf"
    )
