from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile, os, shutil

app = FastAPI()

@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),
    document_code: str = Form(""),
    client_name: str = Form(""),
    customer: str = Form(""),
    contractor: str = Form(""),
    nature: str = Form(""),
    purpose: str = Form(""),
    created_on: str = Form(""),
    created_by: str = Form(""),
):
    temp_dir = tempfile.mkdtemp()
    src_path = os.path.join(temp_dir, file.filename)

    with open(src_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    ext = os.path.splitext(file.filename)[1].lower()
    out_path = os.path.join(temp_dir, file.filename)

    # ================= WORD =================
    if ext == ".docx":
        src = Document(src_path)
        out = Document()

        # copy original content (format safe)
        for p in src.paragraphs:
            new_p = out.add_paragraph(p.text)
            new_p.style = p.style

        # insert details as new section
        out.add_page_break()
        out.add_heading("Document Details", level=1)

        fields = [
            ("Document Code", document_code),
            ("Client Name", client_name),
            ("Customer", customer),
            ("Contractor", contractor),
            ("Nature", nature),
            ("Purpose", purpose),
            ("Created On", created_on),
            ("Created By", created_by),
        ]

        for k, v in fields:
            out.add_paragraph(f"{k}: {v}")

        out.save(out_path)

    # ================= PPT =================
    elif ext == ".pptx":
        src = Presentation(src_path)
        out = Presentation()

        # remove default slide
        while out.slides:
            out.slides._sldIdLst.remove(out.slides._sldIdLst[0])

        # slide 1
        s = src.slides[0]
        slide = out.slides.add_slide(out.slide_layouts[6])
        for sh in s.shapes:
            if sh.has_text_frame:
                box = slide.shapes.add_textbox(
                    sh.left, sh.top, sh.width, sh.height
                )
                box.text_frame.text = sh.text

        # slide 2 (details)
        d = out.slides.add_slide(out.slide_layouts[1])
        d.shapes.title.text = "Document Details"
        tf = d.placeholders[1].text_frame
        tf.text = "\n".join(
            f"{k}: {v}" for k, v in fields
        )

        # remaining slides
        for s in src.slides[1:]:
            slide = out.slides.add_slide(out.slide_layouts[6])
            for sh in s.shapes:
                if sh.has_text_frame:
                    box = slide.shapes.add_textbox(
                        sh.left, sh.top, sh.width, sh.height
                    )
                    box.text_frame.text = sh.text

        out.save(out_path)

    else:
        raise Exception("Unsupported file type")

    return FileResponse(
        out_path,
        headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
    )
