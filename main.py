from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile
import shutil
import zipfile
import os

app = FastAPI()

DOCX_TEMPLATE = "DCS_TEMPLATE.docx"
PPT_TEMPLATE = "DCS_TEMPLATE.pptx"


# -----------------------------
# Replace placeholders in DOCX template
# -----------------------------
def prepare_docx_template(data):

    doc = Document(DOCX_TEMPLATE)

    for p in doc.paragraphs:
        for k, v in data.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)

    return tmp.name


# -----------------------------
# Replace placeholders in PPT template
# -----------------------------
def prepare_ppt_template(data):

    prs = Presentation(PPT_TEMPLATE)

    slide = prs.slides[0]

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        for p in shape.text_frame.paragraphs:
            for k, v in data.items():
                if k in p.text:
                    p.text = p.text.replace(k, v)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)

    return tmp.name


# -----------------------------
# Insert DOCX page after first page
# -----------------------------
def merge_docx(user_docx, template_docx, output):

    main = Document(user_docx)
    template = Document(template_docx)

    new = Document()

    # Copy first page (cover)
    for element in main.element.body[:]:
        new.element.body.append(element)
        if element.tag.endswith("sectPr"):
            break

    # Insert template page
    for element in template.element.body:
        new.element.body.append(element)

    # Append rest of document
    body_started = False

    for element in main.element.body:
        if body_started:
            new.element.body.append(element)

        if element.tag.endswith("sectPr"):
            body_started = True

    new.save(output)


# -----------------------------
# Insert PPT slide at position 2
# -----------------------------
def merge_ppt(user_ppt, template_ppt, output):

    prs_user = Presentation(user_ppt)
    prs_template = Presentation(template_ppt)

    new = Presentation(user_ppt)

    # insert template slide after slide 1
    template_slide = prs_template.slides[0]

    slide_layout = new.slide_layouts[6]
    inserted_slide = new.slides.add_slide(slide_layout)

    for shape in template_slide.shapes:
        el = shape.element
        inserted_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    # move inserted slide to index 1
    slide_ids = new.slides._sldIdLst
    slides = list(slide_ids)

    slide_ids.remove(slides[-1])
    slide_ids.insert(1, slides[-1])

    new.save(output)


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

            template_ready = prepare_docx_template(data)
            merge_docx(input_path, template_ready, output_path)

        elif ext == ".pptx":

            template_ready = prepare_ppt_template(data)
            merge_ppt(input_path, template_ready, output_path)

        else:
            raise HTTPException(status_code=400, detail="Only DOCX and PPTX supported")

        return FileResponse(
            output_path,
            media_type="application/octet-stream",
            headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
