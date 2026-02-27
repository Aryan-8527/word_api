from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile
import os
import shutil

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
        # ================= DOCX ===============================
        # =====================================================
        if ext == ".docx":

            doc = Document(input_path)

            # Create Document Details content in memory
            detail_doc = Document()
            detail_doc.add_heading("Document Details", level=1)

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
                detail_doc.add_paragraph(d)

            # Insert page break after first page
            doc.paragraphs[0].insert_paragraph_after().add_run().add_break()

            # Insert detail page elements after first paragraph
            for element in detail_doc.element.body:
                doc.element.body.insert(1, element)

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # =====================================================
        # ================= PPTX ===============================
        # =====================================================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            # Add new slide using existing layout
            layout = prs.slide_layouts[1]
            detail_slide = prs.slides.add_slide(layout)

            # Move slide to second position
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

        else:
            raise HTTPException(status_code=400, detail="Only DOCX and PPTX supported")

        return FileResponse(
            output_path,
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
