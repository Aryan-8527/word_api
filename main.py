from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
from copy import deepcopy
import tempfile, os, shutil

app = FastAPI()


# ---------- PPT HELPER (SAFE COPY) ----------
def copy_slide(prs, source_slide):
    slide_layout = prs.slide_layouts[6]  # blank
    new_slide = prs.slides.add_slide(slide_layout)

    for shape in source_slide.shapes:
        el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    return new_slide


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

    # Save uploaded file
    with open(uploaded_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()

    # ======================================================
    # ================= WORD (.DOCX) =======================
    # ======================================================
    if ext == ".docx":
        original = Document(uploaded_path)
        final_doc = Document()

        # --- PAGE 1: original first page ---
        for element in original.element.body:
            final_doc.element.body.append(deepcopy(element))
            if element.tag.endswith('sectPr'):
                break

        final_doc.add_page_break()

        # --- PAGE 2: form details ---
        final_doc.add_heading("Document Details", level=1)

        def add(label, value):
            final_doc.add_paragraph(f"{label}: {value or ''}")

        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

        final_doc.add_page_break()

        # --- Remaining pages ---
        remaining = False
        for element in original.element.body:
            if remaining:
                final_doc.element.body.append(deepcopy(element))
            if element.tag.endswith('sectPr'):
                remaining = True

        final_path = os.path.join(temp_dir, filename)
        final_doc.save(final_path)

    # ======================================================
    # ================= PPT (.PPTX) ========================
    # ======================================================
    elif ext == ".pptx":
        original = Presentation(uploaded_path)
        final_ppt = Presentation()

        # Slide 1: original first slide
        copy_slide(final_ppt, original.slides[0])

        # Slide 2: form details
        details_slide = final_ppt.slides.add_slide(final_ppt.slide_layouts[1])
        details_slide.shapes.title.text = "Document Details"

        tf = details_slide.placeholders[1].text_frame
        tf.clear()

        def add_ppt(text):
            p = tf.add_paragraph()
            p.text = text

        add_ppt(f"Document Code: {document_code}")
        add_ppt(f"Client Name: {client_name}")
        add_ppt(f"Customer: {customer}")
        add_ppt(f"Contractor: {contractor}")
        add_ppt(f"Nature: {nature}")
        add_ppt(f"Purpose: {purpose}")
        add_ppt(f"Created On: {created_on}")
        add_ppt(f"Created By: {created_by}")

        # Remaining slides
        for i in range(1, len(original.slides)):
            copy_slide(final_ppt, original.slides[i])

        final_path = os.path.join(temp_dir, filename)
        final_ppt.save(final_path)

    else:
        raise Exception("Unsupported file type")

    # ======================================================
    # ================= RESPONSE ===========================
    # ======================================================
    return FileResponse(
        final_path,
        media_type=file.content_type,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )
