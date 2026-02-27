from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.enum.text import WD_BREAK
from pptx import Presentation
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

            original_doc = Document(input_path)
            new_doc = Document()

            # ---- Copy entire original document ----
            for element in original_doc.element.body:
                new_doc.element.body.append(element)

            body = new_doc._body._element

            # ---- Create page break XML ----
            p = OxmlElement("w:p")
            r = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            r.append(br)
            p.append(r)

            # ---- Insert page break after first element ----
            body.insert(1, p)

            # ---- Insert Document Details page ----
            insert_position = 2

            title_para = new_doc.add_paragraph()
            title_run = title_para.add_run("Document Details")
            title_run.bold = True

            body.remove(title_para._p)
            body.insert(insert_position, title_para._p)
            insert_position += 1

            detail_lines = [
                f"Document Code: {document_code}",
                f"Client Name: {client_name}",
                f"Department: {department}",
                f"Document Type: {document_type}",
                f"Purpose: {purpose}",
                f"Created On: {created_on}",
                f"Created By: {created_by}",
            ]

            for line in detail_lines:
                para = new_doc.add_paragraph(line)
                body.remove(para._p)
                body.insert(insert_position, para._p)
                insert_position += 1

            output_path = os.path.join(temp_dir, file.filename)
            new_doc.save(output_path)

        # =====================================================
        # ================= PPTX ===============================
        # =====================================================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            layout = prs.slides[0].slide_layout
            detail_slide = prs.slides.add_slide(layout)

            # Move slide to 2nd position
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
