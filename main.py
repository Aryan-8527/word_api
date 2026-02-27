from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pptx import Presentation
import tempfile
import shutil
import os
import subprocess
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter

app = FastAPI()


@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),
    document_code: str = Form(""),
    client_name: str = Form(""),
    department: str = Form(""),
    document_type: str = Form(""),
    purpose: str = Form(""),
    created_on: str = Form(""),
    created_by: str = Form("")
):
    try:
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)

        # Save uploaded file
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        ext = os.path.splitext(file.filename)[1].lower()

        # =====================================================
        # ===================== WORD ==========================
        # =====================================================
        if ext == ".docx":

            # STEP 1: Convert Word → PDF using LibreOffice
            subprocess.run([
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_dir,
                input_path
            ], check=True)

            pdf_name = os.path.splitext(file.filename)[0] + ".pdf"
            original_pdf_path = os.path.join(temp_dir, pdf_name)

            # STEP 2: Create Details PDF page
            details_pdf_path = os.path.join(temp_dir, "details_page.pdf")
            c = canvas.Canvas(details_pdf_path)

            c.setFont("Helvetica-Bold", 16)
            c.drawString(200, 800, "Document Details")

            c.setFont("Helvetica", 12)

            lines = [
                f"Document Code: {document_code}",
                f"Client Name: {client_name}",
                f"Department: {department}",
                f"Document Type: {document_type}",
                f"Purpose: {purpose}",
                f"Created On: {created_on}",
                f"Created By: {created_by}",
            ]

            y = 760
            for line in lines:
                c.drawString(80, y, line)
                y -= 25

            c.save()

            # STEP 3: Merge PDFs (Insert at page 2)
            writer = PdfWriter()
            original_reader = PdfReader(original_pdf_path)
            details_reader = PdfReader(details_pdf_path)

            # Page 1
            writer.add_page(original_reader.pages[0])

            # Page 2 (Details)
            writer.add_page(details_reader.pages[0])

            # Remaining pages
            for i in range(1, len(original_reader.pages)):
                writer.add_page(original_reader.pages[i])

            final_pdf_path = os.path.join(temp_dir, "final_output.pdf")

            with open(final_pdf_path, "wb") as f:
                writer.write(f)

            return FileResponse(
                final_pdf_path,
                media_type="application/pdf",
                headers={
                    "Content-Disposition": 'attachment; filename="final_output.pdf"'
                }
            )

        # =====================================================
        # ===================== PPT ===========================
        # =====================================================
        elif ext == ".pptx":

            # 👇 PPT CODE BILKUL SAME HAI (TOUCH NAHI KIYA)
            prs = Presentation(input_path)

            layout = prs.slide_layouts[1]
            detail_slide = prs.slides.add_slide(layout)

            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)
            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            detail_slide.shapes.title.text = "Document Details"

            tf = detail_slide.placeholders[1].text_frame
            tf.clear()

            details = [
                f"Document Code: {document_code}",
                f"Client Name: {client_name}",
                f"Department: {department}",
                f"Document Type: {document_type}",
                f"Purpose: {purpose}",
                f"Created On: {created_on}",
                f"Created By: {created_by}",
            ]

            for d in details:
                p = tf.add_paragraph()
                p.text = d

            output_path = os.path.join(temp_dir, file.filename)
            prs.save(output_path)

            return FileResponse(
                output_path,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={
                    "Content-Disposition": f'attachment; filename="{file.filename}"'
                }
            )

        else:
            raise HTTPException(status_code=400, detail="Only DOCX and PPTX supported")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
