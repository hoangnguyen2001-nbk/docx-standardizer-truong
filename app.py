import os
import io
import urllib.parse
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from core.docx_engine import process_document

BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Docx Standardizer")

# Mount static files
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

@app.get("/")
async def get_index():
    index_file = BASE_DIR / "templates" / "index.html"
    return FileResponse(index_file)

@app.post("/api/process")
async def process_file(
    file: UploadFile = File(...),
    decree_30: bool = Form(False),
    science_math: bool = Form(False),
    science_chemistry: bool = Form(False),
    science_informatics: bool = Form(False),
    science_geography: bool = Form(False),
    flow_control: bool = Form(False),
    page_numbering: bool = Form(False),
    figure_captions: bool = Form(False),
    exam_questions: bool = Form(False),
    bullet_hierarchy: bool = Form(False),
    typography_sanitizer: bool = Form(False),
    style_mode: str = Form("decree30"),
    font_key: str = Form("inter"),
    scale_ratio: str = Form("1.25")
):
    try:
        options = {
            "decree_30": decree_30,
            "science_math": science_math,
            "science_chemistry": science_chemistry,
            "science_informatics": science_informatics,
            "science_geography": science_geography,
            "flow_control": flow_control,
            "page_numbering": page_numbering,
            "figure_captions": figure_captions,
            "exam_questions": exam_questions,
            "bullet_hierarchy": bullet_hierarchy,
            "typography_sanitizer": typography_sanitizer,
            "style_mode": style_mode,
            "font_key": font_key,
            "scale_ratio": scale_ratio
        }
        
        file_bytes = await file.read()
        processed_bytes = process_document(file_bytes, options)
        
        # Safe ASCII + RFC 5987 UTF-8 encoded header for Vietnamese filenames
        encoded_filename = urllib.parse.quote(f"standardized_{file.filename}")
        
        return StreamingResponse(
            io.BytesIO(processed_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename=\"standardized_output.docx\"; filename*=UTF-8''{encoded_filename}"
            }
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
