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


def replace_docx_placeholders(doc, data):
    for p in doc.paragraphs:
        for k, v in data.items():
            if k in p.text:
                p.text = p.text.replace(k, v)


def replace_ppt_placeholders(prs, data):
    slide = prs.slides[0]

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        for paragraph in shape.text_frame.paragraphs:
            for k, v in data.items():
                if k in paragraph.text:
                    paragraph.text = paragraph.text.replace(k, v)


def insert_docx_page(template_path, user_path, data):

    template = Document(template_path)
    user = Document(user_path)

    replace_docx_placeholders(template, data)

    new_doc = Document()

    # Copy first page of user document
    for element in user.element.body:
        new_doc.element.body.append(element)
        if element.tag.endswith("sectPr"):
            break

    # Insert template page
    for element in template.element.body:
        new_doc.element.body.append(element)

    # Append rest of user document
    body_started = False
    for element in user.element.body:
        if body_started:
            new_doc.element.body.append(element)

        if element.tag.endswith("sectPr"):
            body_started = True

    return new_doc


def insert_ppt_slide(template_path, user_path, data):

    template = Presentation(template_path)
    user = Presentation(user_path)

    replace_ppt_placeholders(template, data)

    new_ppt = Presentation()

    # copy first slide
    slide = user.slides[0]
    layout = new_ppt.slide_layouts[6]
    new_slide = new_ppt.slides.add_slide(layout)

    for shape in slide.shapes:
        el = shape.element
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    # insert template slide
    temp_slide = template.slides[0]
    layout = new_ppt.slide_layouts[6]
    new_slide = new_ppt.slides.add_slide(layout)

    for shape in temp_slide.shapes:
        el = shape.element
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    # append remaining slides
    for slide in user.slides[1:]:
        layout = new_ppt.slide_layouts[6]
        new_slide = new_ppt.slides.add_slide(layout)

        for shape in slide.shapes:
            el = shape.element
            new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    return new_ppt


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

        output_path = os.path.join(temp_dir, file.filename)

        if ext == ".docx":

            merged = insert_docx_page(DOCX_TEMPLATE, input_path, data)
            merged.save(output_path)

        elif ext == ".pptx":

            merged = insert_ppt_slide(PPT_TEMPLATE, input_path, data)
            merged.save(output_path)

        else:
            raise HTTPException(
                status_code=400,
                detail="Only DOCX and PPTX supported"
            )

        return FileResponse(
            output_path,
            media_type="application/octet-stream",
            headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
