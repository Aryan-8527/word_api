from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
from copy import deepcopy
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
    final_path = os.path.join(temp_dir, file.filename)

    # =====================================================
    # ================= WORD FILE =========================
    # =====================================================
    if ext == ".docx":
        original = Document(uploaded_path)
        final_doc = Document()

        elements = list(original.element.body)

        # PAGE 1 → Original first page
        for el in elements:
            final_doc.element.body.append(deepcopy(el))
            if el.tag.endswith("sectPr"):
                break

        # PAGE 2 → Form details
        final_doc.add_page_break()
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

        # PAGE 3+ → Remaining original pages
        final_doc.add_page_break()
        started = False
        for el in elements:
            if started:
                final_doc.element.body.append(deepcopy(el))
            if el.tag.endswith("sectPr"):
                started = True

        final_doc.save(final_path)

    # =====================================================
    # ================= PPT FILE ==========================
    # =====================================================
    elif ext == ".pptx":
        original = Presentation(uploaded_path)
        final_ppt = Presentation()

        # Remove default slide
        while len(final_ppt.slides) > 0:
            rId = final_ppt.slides._sldIdLst[0].rId
            final_ppt.part.drop_rel(rId)
            del final_ppt.slides._sldIdLst[0]

        # SLIDE 1 → Original slide 1
        slide = original.slides[0]
        layout = final_ppt.slide_layouts[slide.slide_layout.slide_layout_id]
        new_slide = final_ppt.slides.add_slide(layout)

        for shape in slide.shapes:
            new_slide.shapes._spTree.insert_element_before(
                deepcopy(shape.element), 'p:extLst'
            )

        # SLIDE 2 → Form details
        layout = final_ppt.slide_layouts[1]
        details = final_ppt.slides.add_slide(layout)

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

        # SLIDE 3+ → Remaining slides
        for slide in original.slides[1:]:
            layout = final_ppt.slide_layouts[slide.slide_layout.slide_layout_id]
            new_slide = final_ppt.slides.add_slide(layout)
            for shape in slide.shapes:
                new_slide.shapes._spTree.insert_element_before(
                    deepcopy(shape.element), 'p:extLst'
                )

        final_ppt.save(final_path)

    else:
        return {"error": "Unsupported file type"}

    return FileResponse(
        final_path,
        media_type=file.content_type,
        headers={
            "Content-Disposition": f'attachment; filename="{file.filename}"'
        }
    )
