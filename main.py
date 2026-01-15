from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import tempfile
import os
import shutil

app = FastAPI()


def add_page_break(paragraph):
    run = paragraph.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


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

    for element in original.element.body:
        final_doc.element.body.append(element)

        # Detect first page break
        if not inserted and element.tag.endswith("p"):
            for child in element.iter():
                if child.tag.endswith("br") and child.get(qn("w:type")) == "page":
                    # 🔹 Insert form-details page AFTER first page
                    p = final_doc.add_paragraph()
                    add_page_break(p)

                    final_doc.add_heading("Document Details", level=1)

                    def add(label, value):
                        final_doc.add_paragraph(f"{label} : {value if value else 'None'}")

                    add("Document Code", document_code)
                    add("Client Name", client_name)
                    add("Customer", customer)
                    add("Contractor", contractor)
                    add("Nature", nature)
                    add("Purpose", purpose)
                    add("Created On", created_on)
                    add("Created By", created_by)

                    p2 = final_doc.add_paragraph()
                    add_page_break(p2)

                    inserted = True
                    break

    # Safety fallback: if document has only 1 page
    if not inserted:
        p = final_doc.add_paragraph()
        add_page_break(p)

        final_doc.add_heading("Document Details", level=1)
        add("Document Code", document_code)
        add("Client Name", client_name)
        add("Customer", customer)
        add("Contractor", contractor)
        add("Nature", nature)
        add("Purpose", purpose)
        add("Created On", created_on)
        add("Created By", created_by)

    final_filename = file.filename
    final_path = os.path.join(temp_dir, final_filename)
    final_doc.save(final_path)

    return FileResponse(
        final_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=final_filename
    )
