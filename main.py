from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from docx import Document
import io
import requests

app = FastAPI()

# ---------- INPUT MODEL ----------
class DownloadRequest(BaseModel):
    document_id: int
    client_name: str | None = None
    customer: str | None = None
    contractor: str | None = None
    nature: str | None = None
    purpose: str | None = None
    created_on: str | None = None
    created_by: str | None = None


# ---------- FETCH FILE FROM APEX ----------
def fetch_original_doc(document_id: int) -> bytes:
    """
    Call APEX ORDS / download API to fetch original Word file
    """
    url = f"https://YOUR_APEX_HOST/ords/your_schema/documents/{document_id}"

    r = requests.get(url)
    r.raise_for_status()
    return r.content


# ---------- API ----------
@app.post("/download-doc")
def download_doc(data: DownloadRequest):

    # 1️⃣ Fetch original Word document
    original_bytes = fetch_original_doc(data.document_id)
    original_doc = Document(io.BytesIO(original_bytes))

    # 2️⃣ Create cover page
    cover = Document()
    cover.add_heading("Document Details", level=1)

    cover.add_paragraph(f"Client Name : {data.client_name}")
    cover.add_paragraph(f"Customer    : {data.customer}")
    cover.add_paragraph(f"Contractor  : {data.contractor}")
    cover.add_paragraph(f"Nature      : {data.nature}")
    cover.add_paragraph(f"Purpose     : {data.purpose}")
    cover.add_paragraph(f"Created On  : {data.created_on}")
    cover.add_paragraph(f"Created By  : {data.created_by}")

    cover.add_page_break()

    # 3️⃣ Merge original document into cover
    for element in original_doc.element.body:
        cover.element.body.append(element)

    # 4️⃣ Return final Word
    output = io.BytesIO()
    cover.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=Document_With_Cover.docx"
        }
    )
