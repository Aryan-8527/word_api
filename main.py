from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

        # ===============================
        # WORD DOCUMENT PROCESSING
        # ===============================
        if ext == ".docx":

            doc = Document(input_path)

            template_path = "DCS_TEMPLATE.docx"
            template_doc = Document(template_path)

            replacements = {
                "{{DOCUMENT_CODE}}": document_code,
                "{{CLIENT_NAME}}": client_name,
                "{{DEPARTMENT}}": department,
                "{{DOCUMENT_TYPE}}": document_type,
                "{{PURPOSE}}": purpose,
                "{{CREATED_ON}}": created_on,
                "{{CREATED_BY}}": created_by
            }

            for paragraph in template_doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)

            # PAGE BREAK
            page_break = OxmlElement("w:p")
            run = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run.append(br)
            page_break.append(run)

            body = doc._element.body

            # insert page break after first page
            body.insert(1, page_break)

            insert_index = 2

            for element in template_doc.element.body:
                body.insert(insert_index, element)
                insert_index += 1

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # ===============================
        # POWERPOINT PROCESSING
        # ===============================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            template_path = "DCS_TEMPLATE.pptx"
            template_prs = Presentation(template_path)

            template_slide = template_prs.slides[0]

            new_slide = prs.slides.add_slide(prs.slide_layouts[6])

            for shape in template_slide.shapes:

                if shape.has_text_frame:

                    new_shape = new_slide.shapes.add_textbox(
                        shape.left,
                        shape.top,
                        shape.width,
                        shape.height
                    )

                    tf = new_shape.text_frame
                    tf.clear()

                    for paragraph in shape.text_frame.paragraphs:

                        text = paragraph.text

                        text = text.replace("{{DOCUMENT_CODE}}", document_code)
                        text = text.replace("{{CLIENT_NAME}}", client_name)
                        text = text.replace("{{DEPARTMENT}}", department)
                        text = text.replace("{{DOCUMENT_TYPE}}", document_type)
                        text = text.replace("{{PURPOSE}}", purpose)
                        text = text.replace("{{CREATED_ON}}", created_on)
                        text = text.replace("{{CREATED_BY}}", created_by)

                        p = tf.add_paragraph()
                        p.text = text
                        p.level = paragraph.level

                elif shape.shape_type == 13:  # picture
                    image_stream = shape.image.blob

                    image_path = os.path.join(temp_dir, "temp_img.png")

                    with open(image_path, "wb") as img:
                        img.write(image_stream)

                    new_slide.shapes.add_picture(
                        image_path,
                        shape.left,
                        shape.top,
                        shape.width,
                        shape.height
                    )

            # MOVE SLIDE TO POSITION 2
            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)

            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            output_path = os.path.join(temp_dir, file.filename)
            prs.save(output_path)

        else:
            raise HTTPException(
                status_code=400,
                detail="Only DOCX and PPTX files are supported"
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
