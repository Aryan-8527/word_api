from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docxcompose.composer import Composer
from pptx import Presentation
import tempfile
import shutil
import os

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

            original_doc = Document(input_path)

            # Create new details page
            details_doc = Document()
            details_doc.add_heading("Document Details", level=1)
            details_doc.add_paragraph(f"Document Code: {document_code}")
            details_doc.add_paragraph(f"Client Name: {client_name}")
            details_doc.add_paragraph(f"Department: {department}")
            details_doc.add_paragraph(f"Document Type: {document_type}")
            details_doc.add_paragraph(f"Purpose: {purpose}")
            details_doc.add_paragraph(f"Created On: {created_on}")
            details_doc.add_paragraph(f"Created By: {created_by}")

            details_doc.add_page_break()

            # Merge
            composer = Composer(original_doc)

            # Insert at second position
            composer.insert(1, details_doc)

            output_path = os.path.join(temp_dir, file.filename)
            composer.save(output_path)

            return FileResponse(
                output_path,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={
                    "Content-Disposition": f'attachment; filename="{file.filename}"'
                }
            )

        # =====================================================
        # ===================== PPT ===========================
        # =====================================================
        elif ext == ".pptx":

            # 👇 Ye wala code SAME rakha gaya hai (touch nahi kiya)
            prs = Presentation(input_path)

            layout = prs.slide_layouts[1]
            detail_slide = prs.slides.add_slide(layout)

            # Move slide to 2nd position
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
