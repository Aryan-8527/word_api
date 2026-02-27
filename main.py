from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import tempfile, os, shutil

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

        # =====================================================
        # ================= WORD (.DOCX) ======================
        # =====================================================
        if ext == ".docx":

            # 🔥 OPEN ORIGINAL DOCUMENT (keeps images)
            doc = Document(input_path)

            # Add page break at end of first page
            doc.add_page_break()

            # Add Document Details Page
            doc.add_heading("Document Details", level=1)

            def add(label, val):
                doc.add_paragraph(f"{label}: {val}")

            add("Document Code", document_code)
            add("Client Name", client_name)
            add("Department", department)
            add("Document Type", document_type)
            add("Purpose", purpose)
            add("Created On", created_on)
            add("Created By", created_by)

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # =====================================================
        # ================= PPT (.PPTX) =======================
        # =====================================================
        elif ext == ".pptx":

    prs = Presentation(input_path)

    # Create Document Details slide
    layout = prs.slide_layouts[1]  # Title + Content layout
    detail_slide = prs.slides.add_slide(layout)

    # Move slide to position 2 (index 1)
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[-1])
    xml_slides.insert(1, slides[-1])

    # Add content
    detail_slide.shapes.title.text = "Document Details"

    tf = detail_slide.placeholders[1].text_frame
    tf.clear()

    details = [
        f"Document Code: {document_code}",
        f"Client Name: {client_name}",
        f"Department: {department}",
        f"Document Type: {document_type}",
        f"Purpose: {purpose}",
        f"Created On: {created_on}",
        f"Created By: {created_by}",
    ]

    for d in details:
        p = tf.add_paragraph()
        p.text = d

    output_path = os.path.join(temp_dir, file.filename)
    prs.save(output_path)

        else:
            raise HTTPException(status_code=400, detail="Unsupported file type")

        return FileResponse(
            output_path,
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

