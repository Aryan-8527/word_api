from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.enum.text import WD_BREAK
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

            # ---- Step 1: Create page break ----
            page_break_para = doc.add_paragraph()
            page_break_para.add_run().add_break(WD_BREAK.PAGE)

            # ---- Step 2: Create Document Details content ----
            new_paragraphs = []

            title_para = doc.add_paragraph()
            title_run = title_para.add_run("Document Details")
            title_run.bold = True
            new_paragraphs.append(title_para)

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
                new_paragraphs.append(doc.add_paragraph(d))

            # ---- Step 3: Move new content after first paragraph ----
            insert_position = 1  # After first paragraph

            all_new_content = [page_break_para] + new_paragraphs

            for para in reversed(all_new_content):
                body.remove(para._p)
                body.insert(insert_position, para._p)

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # =====================================================
        # ================= PPTX ===============================
        # =====================================================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            # Use same layout as first slide
            first_layout = prs.slides[0].slide_layout
            detail_slide = prs.slides.add_slide(first_layout)

            # Move slide to 2nd position
            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)
            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            # Set title
            if detail_slide.shapes.title:
                detail_slide.shapes.title.text = "Document Details"

            # Add content textbox
            left = prs.slide_width * 0.1
            top = prs.slide_height * 0.3
            width = prs.slide_width * 0.8
            height = prs.slide_height * 0.5

            textbox = detail_slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
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
            media_type="application/octet-stream",
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
