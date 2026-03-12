from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile
import shutil
import os

app = FastAPI()

DOCX_TEMPLATE = "DCS_TEMPLATE.docx"
PPT_TEMPLATE = "DCS_TEMPLATE.pptx"


# -----------------------------
# REPLACE PLACEHOLDERS DOCX
# -----------------------------
def replace_docx_placeholders(doc, data):

    for p in doc.paragraphs:
        for key, value in data.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in data.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)


# -----------------------------
# APPEND DOCX DOCUMENT
# -----------------------------
def append_docx(template_doc, uploaded_doc):

    for element in uploaded_doc.element.body:
        template_doc.element.body.append(element)


# -----------------------------
# REPLACE PPT PLACEHOLDERS
# -----------------------------
def replace_ppt_placeholders(prs, data):

    slide = prs.slides[0]

    for shape in slide.shapes:

        if not shape.has_text_frame:
            continue

        for paragraph in shape.text_frame.paragraphs:

            for key, value in data.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)


# -----------------------------
# COPY SLIDES FROM USER PPT
# -----------------------------
def copy_user_slides(template_prs, user_prs):

    for slide in user_prs.slides:

        layout = template_prs.slide_layouts[6]  # blank layout
        new_slide = template_prs.slides.add_slide(layout)

        for shape in slide.shapes:

            if shape.has_text_frame:
                textbox = new_slide.shapes.add_textbox(
                    shape.left,
                    shape.top,
                    shape.width,
                    shape.height
                )

                tf = textbox.text_frame

                for paragraph in shape.text_frame.paragraphs:
                    p = tf.add_paragraph()
                    p.text = paragraph.text
                    p.level = paragraph.level


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

        data = {
            "{{DOCUMENT_CODE}}": document_code,
            "{{CLIENT_NAME}}": client_name,
            "{{DEPARTMENT}}": department,
            "{{DOCUMENT_TYPE}}": document_type,
            "{{PURPOSE}}": purpose,
            "{{CREATED_ON}}": created_on,
            "{{CREATED_BY}}": created_by
        }

        # =========================
        # WORD DOCUMENT
        # =========================
        if ext == ".docx":

            template_doc = Document(DOCX_TEMPLATE)
            user_doc = Document(input_path)

            replace_docx_placeholders(template_doc, data)

            template_doc.add_page_break()

            append_docx(template_doc, user_doc)

            output_path = os.path.join(temp_dir, file.filename)

            template_doc.save(output_path)

        # =========================
        # POWERPOINT
        # =========================
        elif ext == ".pptx":

            template_prs = Presentation(PPT_TEMPLATE)
            user_prs = Presentation(input_path)

            replace_ppt_placeholders(template_prs, data)

            copy_user_slides(template_prs, user_prs)

            output_path = os.path.join(temp_dir, file.filename)

            template_prs.save(output_path)

        else:
            raise HTTPException(
                status_code=400,
                detail="Only DOCX and PPTX supported"
            )

        return FileResponse(
            output_path,
            media_type="application/octet-stream",
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
