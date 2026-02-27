from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import tempfile, os, shutil

app = FastAPI()


@app.post("/download-doc")
async def download_doc(
    file: UploadFile = File(...),
    document_code: str = Form(""),
    client_name: str = Form(""),
    department: str = Form(""),
    document_type: str = Form(""),
    purpose: str = Form(""),
    created_on: str = Form(""),
    created_by: str = Form("")
):
    try:
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)

        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        ext = os.path.splitext(file.filename)[1].lower()

        # =====================================================
        # ================= WORD (.DOCX) ======================
        # =====================================================
        if ext == ".docx":

            # 🔥 OPEN ORIGINAL DOCUMENT (keeps images)
            doc = Document(input_path)

            # Add page break at end of first page
            doc.add_page_break()

            # Add Document Details Page
            doc.add_heading("Document Details", level=1)

            def add(label, val):
                doc.add_paragraph(f"{label}: {val}")

            add("Document Code", document_code)
            add("Client Name", client_name)
            add("Department", department)
            add("Document Type", document_type)
            add("Purpose", purpose)
            add("Created On", created_on)
            add("Created By", created_by)

            output_path = os.path.join(temp_dir, file.filename)
            doc.save(output_path)

        # =====================================================
        # ================= PPT (.PPTX) =======================
        # =====================================================
        elif ext == ".pptx":

            src = Presentation(input_path)
            out = Presentation()

            blank_layout = out.slide_layouts[6]

            # 🔥 COPY ALL ORIGINAL SLIDES (INCLUDING IMAGES)
            for slide in src.slides:
                new_slide = out.slides.add_slide(blank_layout)

                for shape in slide.shapes:

                    # Copy Text
                    if shape.has_text_frame:
                        textbox = new_slide.shapes.add_textbox(
                            shape.left, shape.top,
                            shape.width, shape.height
                        )
                        textbox.text_frame.text = shape.text

                    # Copy Images
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        image = shape.image
                        image_bytes = image.blob

                        temp_img_path = os.path.join(temp_dir, "temp_img.png")
                        with open(temp_img_path, "wb") as f:
                            f.write(image_bytes)

                        new_slide.shapes.add_picture(
                            temp_img_path,
                            shape.left,
                            shape.top,
                            shape.width,
                            shape.height
                        )

            # 🔥 ADD DOCUMENT DETAILS AS SECOND SLIDE
            detail_slide = out.slides.add_slide(out.slide_layouts[1])
            detail_slide.shapes.title.text = "Document Details"

            tf = detail_slide.placeholders[1].text_frame
            tf.clear()

            details = [
                f"Document Code: {document_code}",
                f"Client Name: {client_name}",
                f"Department: {department}",
                f"Document Type: {document_type}",
                f"Purpose: {purpose}",
                f"Created On: {created_on}",
                f"Created By: {created_by}",
            ]

            for d in details:
                p = tf.add_paragraph()
                p.text = d

            output_path = os.path.join(temp_dir, file.filename)
            out.save(output_path)

        else:
            raise HTTPException(status_code=400, detail="Unsupported file type")

        return FileResponse(
            output_path,
            headers={
                "Content-Disposition": f'attachment; filename="{file.filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
