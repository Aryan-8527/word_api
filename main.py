from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
from pptx.util import Inches
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

    # =====================================================
    # ================= WORD ==============================
    # =====================================================
    if ext == ".docx":
        src = Document(src_path)
        out = Document()

        # Copy ALL paragraphs first (Word decides pages)
        for p in src.paragraphs:
            out.add_paragraph(p.text)

        # Insert form details AFTER first logical section
        out.add_page_break()
        out.add_heading("Document Details", level=1)

        def add(label, val):
            out.add_paragraph(f"{label}: {val}")

        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

        out.save(out_path)

    # =====================================================
    # ================= PPT ===============================
    # =====================================================
    elif ext == ".pptx":
        src = Presentation(src_path)
        out = Presentation()

        # Remove default slide
        while out.slides:
            out.slides._sldIdLst.remove(out.slides._sldIdLst[0])

        # Slide 1 → Original slide 1
        slide = src.slides[0]
        new_slide = out.slides.add_slide(out.slide_layouts[6])

        for shape in slide.shapes:
            if shape.has_text_frame:
                textbox = new_slide.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                textbox.text_frame.text = shape.text

        # Slide 2 → Form details
        details = out.slides.add_slide(out.slide_layouts[1])
        details.shapes.title.text = "Document Details"

        tf = details.placeholders[1].text_frame
        tf.text = (
            f"Document Code: {document_code}\n"
            f"Client Name: {client_name}\n"
            f"Customer: {customer}\n"
            f"Contractor: {contractor}\n"
            f"Nature: {nature}\n"
            f"Purpose: {purpose}\n"
            f"Created On: {created_on}\n"
            f"Created By: {created_by}"
        )

        # Remaining slides
        for slide in src.slides[1:]:
            new_slide = out.slides.add_slide(out.slide_layouts[6])
            for shape in slide.shapes:
                if shape.has_text_frame:
                    textbox = new_slide.shapes.add_textbox(
                        shape.left, shape.top, shape.width, shape.height
                    )
                    textbox.text_frame.text = shape.text

        out.save(out_path)

    else:
        return {"error": "Unsupported file type"}

    return FileResponse(
        out_path,
        headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
    )
