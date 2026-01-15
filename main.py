from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
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

    # Save uploaded file
    uploaded_path = os.path.join(temp_dir, file.filename)
    with open(uploaded_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    original = Document(uploaded_path)

    # -------- PAGE 2 (FORM DETAILS) ----------
    details = Document()
    details.add_heading("Document Details", level=1)

    def add(label, value):
        details.add_paragraph(f"{label}: {value or ''}")

    add("Document Code", document_code)
    add("Client Name", client_name)
    add("Customer", customer)
    add("Contractor", contractor)
    add("Nature", nature)
    add("Purpose", purpose)
    add("Created On", created_on)
    add("Created By", created_by)

    details.add_page_break()

    # -------- FINAL DOCUMENT ----------
    final_doc = Document()

    # Page 1 → Original document
    for element in original.element.body:
        final_doc.element.body.append(element)

    final_doc.add_page_break()

    # Page 2 → Form details
    for element in details.element.body:
        final_doc.element.body.append(element)

    final_path = os.path.join(temp_dir, file.filename)
    final_doc.save(final_path)

    return FileResponse(
        final_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=file.filename   # ✅ ORIGINAL NAME
    )
