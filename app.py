from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
import os
import pandas as pd
from backend import run_hr_batch  # ✅ your main batch function

app = FastAPI(title="HR Resume–JD Match API", version="1.0")

# ✅ Ensure output folder exists
OUTPUT_DIR = os.path.join(os.getcwd(), "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.post("/run-matching")
async def run_matching_api(
    resumes: list[UploadFile] = File(..., description="Upload multiple .pptx resumes"),
    jds: list[UploadFile] = File(..., description="Upload multiple .docx job descriptions")
):
    try:
        upload_dir = os.path.join(OUTPUT_DIR, "uploads")
        os.makedirs(upload_dir, exist_ok=True)

        resume_paths, jd_paths = [], []

        # ✅ Validate and save resumes
        for resume in resumes:
            if not resume.filename.lower().endswith(".pptx"):
                return JSONResponse(
                    status_code=400,
                    content={"error": f"❌ Invalid file '{resume.filename}'. Only .pptx resumes are accepted."}
                )
            rpath = os.path.join(upload_dir, resume.filename)
            with open(rpath, "wb") as f:
                f.write(await resume.read())
            resume_paths.append(rpath)

        # ✅ Validate and save JDs
        for jd in jds:
            if not jd.filename.lower().endswith(".docx"):
                return JSONResponse(
                    status_code=400,
                    content={"error": f"❌ Invalid file '{jd.filename}'. Only .docx JDs are accepted."}
                )
            jpath = os.path.join(upload_dir, jd.filename)
            with open(jpath, "wb") as f:
                f.write(await jd.read())
            jd_paths.append(jpath)

        # ✅ Run your backend logic
        df, csv_file = run_hr_batch(resume_paths, jd_paths)

        # ✅ Save final output
        output_path = os.path.join(OUTPUT_DIR, "final_match_report.csv")
        df.to_csv(output_path, index=False)

        return FileResponse(
            output_path,
            media_type="text/csv",
            filename="final_match_report.csv"
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": f"❌ Internal Server Error: {str(e)}"}
        )


@app.get("/")
def home():
    return {"message": "✅ HR Resume–JD Matching API is running!"}
