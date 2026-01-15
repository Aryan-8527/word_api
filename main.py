from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
import tempfile, os, shutil

app = FastAPI()

# =========================
# PPT TEXT COPY (SAFE)
# =========================
def copy_textbox_safe(src_shape, dst_slide):
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
            # DO NOT copy color


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
        src = Document(input_path)
        out = Document()

        page_1 = []
        page_rest = []
        first_page_done = False

        for para in src.paragraphs:
            if not first_page_done:
                page_1.append(para.text)
            else:
                page_rest.append(para.text)

            if para._p.xpath(".//w:br[@w:type='page']"):
                first_page_done = True

        # Page 1
        for line in page_1:
            out.add_paragraph(line)

        out.add_page_break()

        # Page 2 – Document Details
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

        out.add_page_break()

        # Remaining pages
        for line in page_rest:
            out.add_paragraph(line)

        output_path = os.path.join(temp_dir, file.filename)
        out.save(output_path)

    # =====================================================
    # ================= PPT (.PPTX) =======================
    # =====================================================
    elif ext == ".pptx":
        src = Presentation(input_path)
        out = Presentation()

        blank = out.slide_layouts[6]

        # Slide 1
        s1 = out.slides.add_slide(blank)
        for shp in src.slides[0].shapes:
            if shp.has_text_frame:
                copy_textbox_safe(shp, s1)

        # Slide 2 – Document Details
        s2 = out.slides.add_slide(out.slide_layouts[1])
        s2.shapes.title.text = "Document Details"
        tf = s2.placeholders[1].text_frame
        tf.clear()

        for text in [
            f"Document Code: {document_code}",
            f"Client Name: {client_name}",
            f"Customer: {customer}",
            f"Contractor: {contractor}",
            f"Nature: {nature}",
            f"Purpose: {purpose}",
            f"Created On: {created_on}",
            f"Created By: {created_by}",
        ]:
            p = tf.add_paragraph()
            p.text = text

        # Remaining slides
        for i in range(1, len(src.slides)):
            s = out.slides.add_slide(blank)
            for shp in src.slides[i].shapes:
                if shp.has_text_frame:
                    copy_textbox_safe(shp, s)

        output_path = os.path.join(temp_dir, file.filename)
        out.save(output_path)

    else:
        raise Exception("Unsupported file type")

    return FileResponse(
        output_path,
        headers={
            "Content-Disposition": f'attachment; filename="{file.filename}"'
        }
    )
