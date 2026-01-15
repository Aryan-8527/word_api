from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile, os, shutil

app = FastAPI()

@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),
    document_code: str = Form(None),
    client_name: str = Form(None),
    customer: str = Form(None),
    contractor: str = Form(None),
    nature: str = Form(None),
    purpose: str = Form(None),
    created_on: str = Form(None),
    created_by: str = Form(None),
):
    temp_dir = tempfile.mkdtemp()
    uploaded_path = os.path.join(temp_dir, file.filename)

    with open(uploaded_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    ext = os.path.splitext(file.filename)[1].lower()

    # ================= WORD =================
    if ext == ".docx":
        original = Document(uploaded_path)
        final_doc = Document()

        # Page 1: original first page
        for el in original.element.body:
            final_doc.element.body.append(el)
            if el.tag.endswith("sectPr"):
                break

        final_doc.add_page_break()

        # Page 2: form details
        final_doc.add_heading("Document Details", level=1)

        def add(label, val):
            final_doc.add_paragraph(f"{label}: {val or ''}")

        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

        final_doc.add_page_break()

        # Remaining pages
        remaining = False
        for el in original.element.body:
            if remaining:
                final_doc.element.body.append(el)
            if el.tag.endswith("sectPr"):
                remaining = True

        final_path = os.path.join(temp_dir, file.filename)
        final_doc.save(final_path)

    # ================= PPT =================
    elif ext == ".pptx":
        original = Presentation(uploaded_path)
        final_ppt = Presentation()

        # Copy slide 1
        final_ppt.slides.add_slide(original.slides[0].slide_layout)
        final_ppt.slides[-1]._element.extend(original.slides[0]._element)

        # Insert details slide (slide 2)
        slide = final_ppt.slides.add_slide(final_ppt.slide_layouts[1])
        slide.shapes.title.text = "Document Details"

        content = slide.placeholders[1].text_frame
        content.clear()

        def add(text):
            p = content.add_paragraph()
            p.text = text

        add(f"Document Code: {document_code}")
        add(f"Client Name: {client_name}")
        add(f"Customer: {customer}")
        add(f"Contractor: {contractor}")
        add(f"Nature: {nature}")
        add(f"Purpose: {purpose}")
        add(f"Created On: {created_on}")
        add(f"Created By: {created_by}")

        # Remaining slides
        for i in range(1, len(original.slides)):
            s = original.slides[i]
            final_ppt.slides.add_slide(s.slide_layout)
            final_ppt.slides[-1]._element.extend(s._element)

        final_path = os.path.join(temp_dir, file.filename)
        final_ppt.save(final_path)

    # ================= OTHER FILES =================
    else:
        final_path = uploaded_path

    return FileResponse(
        final_path,
        filename=file.filename,
        media_type=file.content_type
    )
