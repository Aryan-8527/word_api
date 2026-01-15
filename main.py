from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from docx.oxml.ns import qn
import tempfile, os, shutil

app = FastAPI()


def has_page_break(paragraph):
    for run in paragraph.runs:
        for br in run._element.findall(".//w:br", run._element.nsmap):
            if br.get(qn("w:type")) == "page":
                return True
    return False


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
    uploaded_path = os.path.join(temp_dir, file.filename)

    with open(uploaded_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    original = Document(uploaded_path)
    final_doc = Document()

    inserted = False

    for para in original.paragraphs:
        new_para = final_doc.add_paragraph()
        new_para.style = para.style

        for run in para.runs:
            r = new_para.add_run(run.text)
            r.bold = run.bold
            r.italic = run.italic
            r.underline = run.underline

        # Detect FIRST page break
        if not inserted and has_page_break(para):
            final_doc.add_page_break()

            final_doc.add_heading("Document Details", level=1)

            def add(label, value):
                final_doc.add_paragraph(f"{label} : {value or 'None'}")

            add("Document Code", document_code)
            add("Client Name", client_name)
            add("Customer", customer)
            add("Contractor", contractor)
            add("Nature", nature)
            add("Purpose", purpose)
            add("Created On", created_on)
            add("Created By", created_by)

            final_doc.add_page_break()
            inserted = True

    # If document has no page breaks (1-page document)
    if not inserted:
        final_doc.add_page_break()
        final_doc.add_heading("Document Details", level=1)
        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

    final_path = os.path.join(temp_dir, file.filename)
    final_doc.save(final_path)

    return FileResponse(
        final_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=file.filename
    )
