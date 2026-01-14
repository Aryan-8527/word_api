from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
import tempfile
import os

app = FastAPI()

class CoverData(BaseModel):
    client_name: str | None = None
    customer: str | None = None
    contractor: str | None = None
    nature: str | None = None
    purpose: str | None = None
    created_on: str | None = None
    created_by: str | None = None

@app.post("/cover-doc")
def generate_cover(data: CoverData):
    doc = Document()

    doc.add_heading("Document Details", level=1)

    def add(label, value):
        doc.add_paragraph(f"{label} : {value if value else 'None'}")

    add("Client Name", data.client_name)
    add("Customer", data.customer)
    add("Contractor", data.contractor)
    add("Nature", data.nature)
    add("Purpose", data.purpose)
    add("Created On", data.created_on)
    add("Created By", data.created_by)

    tmp_dir = tempfile.gettempdir()
    file_path = os.path.join(tmp_dir, "cover_page.docx")
    doc.save(file_path)

    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="cover_page.docx"
    )
