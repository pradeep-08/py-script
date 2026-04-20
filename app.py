from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import shutil
import tempfile
from generator import generate_can_from_excel_with_master

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Use system temp directory for Vercel compatibility
TEMP_DIR = tempfile.gettempdir()
UPLOAD_DIR = os.path.join(TEMP_DIR, "uploads")
OUTPUT_DIR = os.path.join(TEMP_DIR, "output")
MASTER_CAN_PATH = os.path.join(os.getcwd(), "master", "ECU_Unlock_Spec 3 (1).can")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.post("/api/convert-excel-to-can")
async def generate_can(file: UploadFile = File(...)):
    try:
        if not file.filename.endswith((".xlsx", ".xlsm", ".xls")):
            raise HTTPException(status_code=400, detail="Only Excel files are allowed")

        uploaded_excel_path = os.path.join(UPLOAD_DIR, file.filename)

        with open(uploaded_excel_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Output format explicitly returned as JSON with dummy generated code and mock metrics
        result = generate_can_from_excel_with_master(
            excel_path=uploaded_excel_path,
            master_can_path=MASTER_CAN_PATH,
            output_dir=OUTPUT_DIR
        )

        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
