from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import Response
from docx import Document
import io

app = FastAPI()

@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),

    client_name: str = Form(""),
    customer: str = Form(""),
    contractor: str = Form(""),
    nature: str = Form(""),
    purpose: str = Form(""),
    created_on: str = Form(""),
    created_by: str = Form("")
):
    # ---- COVER PAGE DOCUMENT ----
    cover = Document()

    cover.add_heading("Document Details", level=1)
    cover.add_paragraph(f"Client Name : {client_name}")
    cover.add_paragraph(f"Customer    : {customer}")
    cover.add_paragraph(f"Contractor  : {contractor}")
    cover.add_paragraph(f"Nature      : {nature}")
    cover.add_paragraph(f"Purpose     : {purpose}")
    cover.add_paragraph(f"Created On  : {created_on}")
    cover.add_paragraph(f"Created By  : {created_by}")
    cover.add_page_break()

    # ---- LOAD ORIGINAL UPLOADED DOC ----
    original_bytes = await file.read()
    original = Document(io.BytesIO(original_bytes))

    # ---- APPEND ORIGINAL CONTENT ----
    for element in original.element.body:
        cover.element.body.append(element)

    # ---- SAVE FINAL DOCUMENT ----
    output = io.BytesIO()
    cover.save(output)
    output.seek(0)

    return Response(
        content=output.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{file.filename}"'
        }
    )
