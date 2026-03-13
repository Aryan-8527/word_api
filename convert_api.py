from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import subprocess
import tempfile
import os
import shutil

app = FastAPI()

@app.post("/convert-to-pdf")
async def convert_to_pdf(file: UploadFile = File(...)):

    try:

        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)

        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Convert to PDF
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            temp_dir,
            input_path
        ])

        pdf_name = os.path.splitext(file.filename)[0] + ".pdf"
        pdf_path = os.path.join(temp_dir, pdf_name)

        return FileResponse(
            pdf_path,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{pdf_name}"'
            }
        )

    except Exception as e:
        return {"error": str(e)}
