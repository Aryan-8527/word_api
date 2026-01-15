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
    input_path = os.path.join(temp_dir, file.filename)

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    ext = file.filename.lower().split(".")[-1]

    # ===================== WORD =====================
    if ext == "docx":
        original = Document(input_path)
        final_doc = Document()

        # Page 1 (original first page)
        for el in original.element.body:
            final_doc.element.body.append(el)
            if el.tag.endswith("sectPr"):
                break

        final_doc.add_page_break()

        # Page 2 (details)
        final_doc.add_heading("Document Details", level=1)
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
            final_doc.add_paragraph(f"{k}: {v}")

        final_doc.add_page_break()

        # Remaining pages
        remaining = False
        for el in original.element.body:
            if remaining:
                final_doc.element.body.append(el)
            if el.tag.endswith("sectPr"):
                remaining = True

        output_path = os.path.join(temp_dir, file.filename)
        final_doc.save(output_path)

    # ===================== POWERPOINT =====================
    elif ext == "pptx":
        original = Presentation(input_path)
        final_ppt = Presentation()

        # Copy slide layouts
        def copy_slide(slide):
            layout = final_ppt.slide_layouts[slide.slide_layout.slide_layout_id]
            new_slide = final_ppt.slides.add_slide(layout)
            for shape in slide.shapes:
                if shape.has_text_frame:
                    new_slide.shapes.title.text = shape.text

        # Slide 1
        copy_slide(original.slides[0])

        # Slide 2 – Details
        detail_slide = final_ppt.slides.add_slide(final_ppt.slide_layouts[1])
        detail_slide.shapes.title.text = "Document Details"
        body = detail_slide.placeholders[1].text_frame
        body.clear()

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
            body.add_paragraph().text = f"{k}: {v}"

        # Remaining slides
        for i in range(1, len(original.slides)):
            copy_slide(original.slides[i])

        output_path = os.path.join(temp_dir, file.filename)
        final_ppt.save(output_path)

    else:
        raise Exception("Unsupported file type")

    return FileResponse(
        output_path,
        media_type=file.content_type,
        headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
    )
