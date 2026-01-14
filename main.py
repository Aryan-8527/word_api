from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from docx import Document
from io import BytesIO

app = FastAPI()

@app.post("/download-doc")
def download_doc(payload: dict):
    """
    payload = {
      client_name, customer, contractor,
      nature, purpose, created_on, created_by,
      file_base64   <-- uploaded file
    }
    """

    # 1️⃣ Load uploaded Word file
    import base64
    original_bytes = base64.b64decode(payload["file_base64"])
    original_doc = Document(BytesIO(original_bytes))

    # 2️⃣ Create cover page
    cover = Document()
    cover.add_heading("Document Details", level=1)
    cover.add_paragraph(f"Client Name : {payload.get('client_name')}")
    cover.add_paragraph(f"Customer    : {payload.get('customer')}")
    cover.add_paragraph(f"Contractor  : {payload.get('contractor')}")
    cover.add_paragraph(f"Nature      : {payload.get('nature')}")
    cover.add_paragraph(f"Purpose     : {payload.get('purpose')}")
    cover.add_paragraph(f"Created On  : {payload.get('created_on')}")
    cover.add_paragraph(f"Created By  : {payload.get('created_by')}")

    cover.add_page_break()

    # 3️⃣ Merge documents (PROPER WAY)
    for element in original_doc.element.body:
        cover.element.body.append(element)

    # 4️⃣ Return merged file
    output = BytesIO()
    cover.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=Document_With_Cover.docx"
        }
    )
