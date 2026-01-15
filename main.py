from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile, os, shutil

app = FastAPI()

# =========================
# PPT HELPER (KEEP FORMATTING)
# =========================
def copy_textbox(src_shape, dst_slide):
    tb = dst_slide.shapes.add_textbox(
        src_shape.left, src_shape.top,
        src_shape.width, src_shape.height
    )
    tf = tb.text_frame
    tf.clear()

    for p in src_shape.text_frame.paragraphs:
        new_p = tf.add_paragraph()
        new_p.alignment = p.alignment

        for r in p.runs:
            new_r = new_p.add_run()
            new_r.text = r.text
            new_r.font.name = r.font.name
            new_r.font.size = r.font.size
            new_r.font.bold = r.font.bold
            new_r.font.italic = r.font.italic
            new_r.font.underline = r.font.underline
            if r.font.color.rgb:
                new_r.font.color.rgb = r.font.color.rgb


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
        result = Document(input_path)
        body = result.element.body

        # Find end of page 1
        insert_pos = None
        for i, el in enumerate(body):
            if el.tag.endswith("p") and el.xpath(".//w:br[@w:type='page']"):
                insert_pos = i + 1
                break

        if insert_pos is None:
            insert_pos = len(body)

        # Add ONE page break
        page_break = result.add_page_break()._element
        body.insert(insert_pos, page_break)

        # Form page content (NO page break here)
        form_doc = Document()
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

        # Insert form content
        for el in reversed(form_doc.element.body):
            body.insert(insert_pos + 1, el)

        output_path = os.path.join(temp_dir, file.filename)
        result.save(output_path)

    # =====================================================
    # ================= PPT (.PPTX) =======================
    # =====================================================
    elif ext == ".pptx":
        src = Presentation(input_path)
        out = Presentation()

        blank_layout = out.slide_layouts[6]

        # Slide 1: original slide 1
        s1 = out.slides.add_slide(blank_layout)
        for shp in src.slides[0].shapes:
            if shp.has_text_frame:
                copy_textbox(shp, s1)

        # Slide 2: Document Details
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

        # Remaining original slides
        for i in range(1, len(src.slides)):
            s = out.slides.add_slide(blank_layout)
            for shp in src.slides[i].shapes:
                if shp.has_text_frame:
                    copy_textbox(shp, s)

        output_path = os.path.join(temp_dir, file.filename)
        out.save(output_path)

    else:
        raise Exception("Unsupported file type")

    return FileResponse(
        output_path,
        media_type=file.content_type,
        headers={
            "Content-Disposition": f'attachment; filename="{file.filename}"'
        }
    )
