from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
import tempfile
import os
import shutil

app = FastAPI()

@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),
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

    # --- Create cover page ---
    cover = Document()
    cover.add_heading("Document Details", level=1)

    def add(label, value):
        cover.add_paragraph(f"{label} : {value if value else 'None'}")

    add("Client Name", client_name)
    add("Customer", customer)
    add("Contractor", contractor)
    add("Nature", nature)
    add("Purpose", purpose)
    add("Created On", created_on)
    add("Created By", created_by)

    # Page break
    cover.add_page_break()

    # --- Append uploaded document ---
    original = Document(uploaded_path)
    for element in original.element.body:
        cover.element.body.append(element)

    final_path = os.path.join(temp_dir, "final_document.docx")
    cover.save(final_path)

    return FileResponse(
        final_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Document_With_Cover.docx"
    )
