from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import Response
from docx import Document
import io
import oracledb

app = FastAPI()

# ---- Oracle connection ----
conn = oracledb.connect(
    user="YOUR_DB_USER",
    password="YOUR_DB_PASSWORD",
    dsn="YOUR_DB_DSN"
)

@app.post("/download-doc")
async def download_doc(request: Request):

    data = await request.json()  # <-- THIS WAS MISSING

    document_id = data.get("document_id")
    if not document_id:
        raise HTTPException(status_code=400, detail="document_id missing")

    # ---- Fetch Word file from DB ----
    cur = conn.cursor()
    cur.execute("""
        SELECT file_blob
        FROM dms_documents
        WHERE document_id = :id
    """, id=document_id)

    row = cur.fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Document not found")

    original_doc = Document(io.BytesIO(row[0].read()))

    # ---- Create cover page ----
    cover = Document()
    cover.add_heading("Document Details", level=1)

    cover.add_paragraph(f"Client Name : {data.get('client_name','')}")
    cover.add_paragraph(f"Customer    : {data.get('customer','')}")
    cover.add_paragraph(f"Contractor  : {data.get('contractor','')}")
    cover.add_paragraph(f"Nature      : {data.get('nature','')}")
    cover.add_paragraph(f"Purpose     : {data.get('purpose','')}")
    cover.add_paragraph(f"Created On  : {data.get('created_on','')}")
    cover.add_paragraph(f"Created By  : {data.get('created_by','')}")

    cover.add_page_break()

    # ---- Append original content ----
    for element in original_doc.element.body:
        cover.element.body.append(element)

    # ---- Return combined DOCX ----
    out = io.BytesIO()
    cover.save(out)
    out.seek(0)

    return Response(
        content=out.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=Document_With_Cover.docx"
        }
    )
