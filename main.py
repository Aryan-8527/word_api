from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.enum.section import WD_SECTION
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

        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        ext = os.path.splitext(file.filename)[1].lower()

        # ================= DOCX =================
        if ext == ".docx":

            doc = Document(input_path)

            # STEP 1: Create new section (new page)
            section = doc.add_section(WD_SECTION.NEW_PAGE)

            # STEP 2: Move this section to second position
            body = doc._body._element
            new_section = body[-1]
            body.remove(new_section)
            body.insert(1, new_section)

            # STEP 3: Add only details in that section
            p = doc.add_paragraph()
            p._p.getparent().remove(p._p)
            body.insert(2, p._p)

            p.add_run("Document Details\n\n").bold = True
            p.add_run(f"Document Code: {document_code}\n")
            p.add_run(f"Client Name: {client_name}\n")
            p.add_run(f"Department: {department}\n")
            p.add_run(f"Document Type: {document_type}\n")
            p.add_run(f"Purpose: {purpose}\n")
            p.add_run(f"Created On: {created_on}\n")
            p.add_run(f"Created By: {created_by}\n")

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # ================= PPTX =================
        elif ext == ".pptx":

            prs = Presentation(input_path)
            layout = prs.slides[0].slide_layout
            detail_slide = prs.slides.add_slide(layout)

            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)
            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            if detail_slide.shapes.title:
                detail_slide.shapes.title.text = "Document Details"

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
