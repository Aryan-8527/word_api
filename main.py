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

            if len(doc.paragraphs) == 0:
                raise HTTPException(status_code=400, detail="Empty document")

            body = doc._body._element

            # STEP 1: First paragraph ke baad blank lines force karo
            first_para = doc.paragraphs[0]
            for _ in range(40):   # adjust if needed
                first_para.add_run("\n")

            # STEP 2: Details insert karo second position par
            insert_position = 1

            details_lines = [
                "Document Details",
                "",
                f"Document Code: {document_code}",
                f"Client Name: {client_name}",
                f"Department: {department}",
                f"Document Type: {document_type}",
                f"Purpose: {purpose}",
                f"Created On: {created_on}",
                f"Created By: {created_by}",
            ]

            for line in reversed(details_lines):
                para = doc.add_paragraph(line)
                body.remove(para._p)
                body.insert(insert_position, para._p)

            # STEP 3: Details ke baad bhi blank lines push karo
            details_para = doc.paragraphs[insert_position]
            for _ in range(35):  # adjust if needed
                details_para.add_run("\n")

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

            return FileResponse(
                output_path,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={
                    "Content-Disposition": f'attachment; filename="{file.filename}"'
                }
            )

        # =====================================================
        # ================= PPTX ===============================
        # =====================================================
        elif ext == ".pptx":

            # ⚠ PPT CODE BILKUL SAME RAKHA HAI
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
