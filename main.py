from fastapi import FastAPI
from fastapi.responses import Response
from docx import Document
from io import BytesIO
import base64

app = FastAPI()

@app.post("/download-doc")
async def download_doc(payload: dict):

    # ---- Decode file ----
    file_base64 = payload.get("file_base64")
    file_bytes = base64.b64decode(file_base64)

    # Load original document
    original = Document(BytesIO(file_bytes))

    # ---- Create cover page ----
    cover = Document()
    cover.add_heading("Document Details", level=1)

    def add(label, value):
        cover.add_paragraph(f"{label} : {value if value else 'None'}")

    add("Client Name", payload.get("client_name"))
    add("Customer", payload.get("customer"))
    add("Contractor", payload.get("contractor"))
    add("Nature", payload.get("nature"))
    add("Purpose", payload.get("purpose"))
    add("Created On", payload.get("created_on"))
    add("Created By", payload.get("created_by"))

    cover.add_page_break()

    # ---- Append original document ----
    for element in original.element.body:
        cover.element.body.append(element)

    # ---- Return final Word ----
    output = BytesIO()
    cover.save(output)
    output.seek(0)

    return Response(
        content=output.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=Document_With_Cover.docx"
        }
    )
