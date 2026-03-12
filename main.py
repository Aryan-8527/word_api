from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt  # UPDATED
from docx.oxml import OxmlElement  # UPDATED
from docx.oxml.ns import qn  # UPDATED
from pptx import Presentation
from pptx.util import Pt as PPTPt  # UPDATED
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

            # STEP 3: Create Heading paragraph
            heading = doc.add_paragraph()
            heading._p.getparent().remove(heading._p)
            body.insert(2, heading._p)

            run = heading.add_run("DOCUMENT DETAILS")
            run.bold = True
            run.underline = True
            run.font.name = "Arial"  # UPDATED
            run.font.size = Pt(22)  # UPDATED
            heading.alignment = 1  # center  # UPDATED

            # force Arial (Word XML compatibility)
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')  # UPDATED

            # STEP 4: Create details paragraph
            details = doc.add_paragraph()
            details._p.getparent().remove(details._p)
            body.insert(3, details._p)

            details_list = [
                f"Document Code : {document_code}",
                f"Client Name : {client_name}",
                f"Department : {department}",
                f"Document Type : {document_type}",
                f"Purpose : {purpose}",
                f"Created On : {created_on}",
                f"Created By : {created_by}"
            ]

            for d in details_list:

                run = details.add_run(d + "\n")
                run.font.name = "Arial"  # UPDATED
                run.font.size = Pt(16)  # UPDATED

                r = run._element
                r.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')  # UPDATED

            # STEP 5: Set line spacing 2.0
            details.paragraph_format.line_spacing = 2  # UPDATED

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # ================= PPTX =================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            # UPDATED → Use Title and Content layout
            layout = prs.slide_layouts[1]  # Title and Content

            detail_slide = prs.slides.add_slide(layout)

            # Move slide to second position
            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)
            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            # ================= TITLE =================
            if detail_slide.shapes.title:

                title = detail_slide.shapes.title
                title.text = "DOCUMENT DETAILS"

                for paragraph in title.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Arial"  # UPDATED
                        run.font.size = PPTPt(40)  # UPDATED

            # ================= CONTENT =================
            body_shape = detail_slide.placeholders[1]
            tf = body_shape.text_frame
            tf.clear()

            details = [
                f"Document Code : {document_code}",
                f"Client Name : {client_name}",
                f"Department : {department}",
                f"Document Type : {document_type}",
                f"Purpose : {purpose}",
                f"Created On : {created_on}",
                f"Created By : {created_by}",
            ]

            for i, d in enumerate(details):

                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()

                p.text = d
                p.level = 0  # bullet

                for run in p.runs:
                    run.font.name = "Arial"  # UPDATED
                    run.font.size = PPTPt(24)  # UPDATED

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
