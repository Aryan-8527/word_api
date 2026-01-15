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
    created_by: str = Form("")
):
    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, file.filename)

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    ext = os.path.splitext(file.filename)[1].lower()

    # =====================================================
    # ================= WORD (.DOCX) ======================
    # =====================================================
    if ext == ".docx":
        original = Document(input_path)
        result = Document(input_path)   # reuse original to avoid blank page

        body = result.element.body

        # Find page break index
        page_break_index = None
        for i, el in enumerate(body):
            if el.tag.endswith("p") and el.xpath(".//w:br[@w:type='page']"):
                page_break_index = i
                break

        # Insert form page after page 1
        insert_pos = page_break_index + 1 if page_break_index else len(body)

        form_doc = Document()
        form_doc.add_page_break()
        form_doc.add_heading("Document Details", level=1)

        def add(label, val):
            form_doc.add_paragraph(f"{label}: {val}")

        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

        for el in reversed(form_doc.element.body):
            body.insert(insert_pos, el)

        output_path = os.path.join(temp_dir, file.filename)
        result.save(output_path)

    # =====================================================
    # ================= PPT (.PPTX) =======================
    # =====================================================
    elif ext == ".pptx":
        src = Presentation(input_path)
        out = Presentation()

        # Slide 1: original slide 1
        layout = out.slide_layouts[6]
        s1 = out.slides.add_slide(layout)
        for shp in src.slides[0].shapes:
            if shp.has_text_frame:
                s1.shapes.add_textbox(
                    shp.left, shp.top, shp.width, shp.height
                ).text_frame.text = shp.text

        # Slide 2: form details
        s2 = out.slides.add_slide(out.slide_layouts[1])
        s2.shapes.title.text = "Document Details"
        tf = s2.placeholders[1].text_frame
        tf.clear()

        def add_ppt(t):
            p = tf.add_paragraph()
            p.text = t

        add_ppt(f"Document Code: {document_code}")
        add_ppt(f"Client Name: {client_name}")
        add_ppt(f"Customer: {customer}")
        add_ppt(f"Contractor: {contractor}")
        add_ppt(f"Nature: {nature}")
        add_ppt(f"Purpose: {purpose}")
        add_ppt(f"Created On: {created_on}")
        add_ppt(f"Created By: {created_by}")

        # Remaining slides
        for i in range(1, len(src.slides)):
            s = out.slides.add_slide(layout)
            for shp in src.slides[i].shapes:
                if shp.has_text_frame:
                    s.shapes.add_textbox(
                        shp.left, shp.top, shp.width, shp.height
                    ).text_frame.text = shp.text

        output_path = os.path.join(temp_dir, file.filename)
        out.save(output_path)

    else:
        raise Exception("Unsupported file type")

    return FileResponse(
        output_path,
        media_type=file.content_type,
        headers={"Content-Disposition": f'attachment; filename="{file.filename}"'}
    )
