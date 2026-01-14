from fastapi import FastAPI, Response
from docx import Document
import io

app = FastAPI()

@app.post("/download-doc")
def download_doc(data: dict):

    doc = Document()

    # Cover Page
    doc.add_heading("Document Details", level=1)
    doc.add_paragraph(f"Client Name : {data.get('client_name','')}")
    doc.add_paragraph(f"Customer    : {data.get('customer','')}")
    doc.add_paragraph(f"Contractor  : {data.get('contractor','')}")
    doc.add_paragraph(f"Nature      : {data.get('nature','')}")
    doc.add_paragraph(f"Purpose     : {data.get('purpose','')}")
    doc.add_paragraph(f"Created On  : {data.get('created_on','')}")
    doc.add_paragraph(f"Created By  : {data.get('created_by','')}")
    doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return Response(
        content=buffer.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=Document_With_Cover.docx"
        }
    )
