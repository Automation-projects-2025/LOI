from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import os
import uuid
import subprocess
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import base64
app = FastAPI()

UPLOAD_DIR = "temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)
if not os.path.exists("service_account.json"):
    encoded_creds = os.getenv("GOOGLE_CREDS_B64")
    if encoded_creds:
        with open("service_account.json", "wb") as f:
            f.write(base64.b64decode(encoded_creds))

# Load credentials from Render Environment Variable
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

@app.post("/convert/")
async def convert_docx_to_pdf(file: UploadFile = File(...)):
    file_id = str(uuid.uuid4())
    docx_path = os.path.join(UPLOAD_DIR, f"{file_id}.docx")
    pdf_path = os.path.join(UPLOAD_DIR, f"{file_id}.pdf")

    with open(docx_path, "wb") as f:
        f.write(await file.read())

    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", docx_path,
        "--outdir", UPLOAD_DIR
    ], check=True)

    # Upload to Google Drive
    file_metadata = {"name": f"{file_id}.pdf"}
    media = MediaFileUpload(pdf_path, mimetype="application/pdf")
    uploaded = drive_service.files().create(
        body=file_metadata, media_body=media, fields="id").execute()

    file_id = uploaded.get("id")

    # Make file public
    drive_service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
    ).execute()

    shareable_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

    return JSONResponse(content={"pdf_link": shareable_link})
