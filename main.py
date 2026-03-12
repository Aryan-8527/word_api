from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Pt as PPTPt
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

        # ================= WORD =================
        if ext == ".docx":

            doc = Document(input_path)

            body = doc._element.body

            # -------------------------
            # PAGE BREAK ELEMENT
            # -------------------------
            page_break = OxmlElement("w:p")
            run = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run.append(br)
            page_break.append(run)

            # -------------------------
            # HEADING
            # -------------------------
            heading_para = doc.add_paragraph()
            heading_para.text = "DOCUMENT DETAILS"

            run = heading_para.runs[0]
            run.bold = True
            run.underline = True
            run.font.name = "Arial"
            run.font.size = Pt(22)

            heading_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # -------------------------
            # DETAILS
            # -------------------------
            details_para = doc.add_paragraph()
            details_para.paragraph_format.line_spacing = 2

            details_list = [
                f"Document Code : {document_code}",
                f"Client Name : {client_name}",
                f"Department : {department}",
                f"Document Type : {document_type}",
                f"Purpose : {purpose}",
                f"Created On : {created_on}",
                f"Created By : {created_by}",
            ]

            for d in details_list:
                r = details_para.add_run(d + "\n")
                r.font.name = "Arial"
                r.font.size = Pt(16)

            # -------------------------
            # MOVE ELEMENTS TO TOP
            # -------------------------
            body.insert(0, page_break)
            body.insert(0, details_para._p)
            body.insert(0, heading_para._p)

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # ================= PPT =================
        elif ext == ".pptx":

            prs = Presentation(input_path)

            layout = prs.slide_layouts[1]
            detail_slide = prs.slides.add_slide(layout)

            slide_ids = prs.slides._sldIdLst
            slides = list(slide_ids)

            slide_ids.remove(slides[-1])
            slide_ids.insert(1, slides[-1])

            if detail_slide.shapes.title:

                title = detail_slide.shapes.title
                title.text = "DOCUMENT DETAILS"

                for paragraph in title.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Arial"
                        run.font.size = PPTPt(40)

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
                p.level = 0

                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = PPTPt(24)

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
