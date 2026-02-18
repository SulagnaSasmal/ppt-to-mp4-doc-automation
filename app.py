import logging
from pathlib import Path
from typing import Dict
from uuid import uuid4
import json
from threading import Thread
import pythoncom
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import shutil
import os
from ppt_pipeline import run_pipeline

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "ppt_to_video.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("ppt-video-api")

jobs: Dict[str, dict] = {}

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
JOBS_DIR = BASE_DIR / "jobs"
templates = Jinja2Templates(directory="templates")

app = FastAPI(
    title="PPT → Animated Video with AI Voice",
    description="""
Convert PowerPoint presentations into animated MP4 videos
using slide animations and Azure AI voice-over from slide notes.

**Pipeline**
1. PowerPoint → Animated MP4
2. Notes → Azure TTS
3. FFmpeg → Final Video
""",
    version="1.0.0"
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

UPLOAD_DIR = "uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)
JOBS_DIR.mkdir(exist_ok=True)


def persist_status(job_id: str) -> None:
    job = jobs.get(job_id)
    if not job:
        return
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    status_path = job_dir / "status.json"
    status_path.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")


def append_log(job_id: str, message: str) -> None:
    job = jobs.get(job_id)
    if not job:
        return
    log_path = job.get("log")
    if not log_path:
        return
    log_file = Path(log_path)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", encoding="utf-8") as f:
        f.write(message.rstrip() + "\n")


@app.get("/", response_class=HTMLResponse)
def upload_page(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )

def update_job(job_id: str, status: str = None, progress: int = None, message: str = None, **kwargs):
    """Update job status and persist changes"""
    job = jobs.get(job_id)
    if not job:
        return
    if status:
        job["status"] = status
    if progress is not None:
        job["progress"] = progress
    if message:
        job["message"] = message
        append_log(job_id, message)
    job.update(kwargs)
    persist_status(job_id)


def run_conversion_async(job_id: str, ppt_path: str, job_dir: str):
    """Run conversion in background thread with incremental status updates"""
    # Initialize COM for this thread (required for win32com)
    pythoncom.CoInitialize()
    
    try:
        logger.info("[%s] Step 1: Exporting animated video", job_id)
        update_job(job_id, status="processing", progress=10, message="Exporting animated video")

        logger.info("[%s] Step 2: Generating animated video via PowerPoint", job_id)
        update_job(job_id, progress=30, message="Generating animated video via PowerPoint")

        logger.info("[%s] Step 3: Generating Azure TTS from notes", job_id)
        update_job(job_id, progress=60, message="Generating Azure TTS from notes")

        logger.info("[%s] Step 4: FFmpeg muxing", job_id)
        update_job(job_id, progress=85, message="FFmpeg muxing")

        # Run the actual pipeline
        output_video = run_pipeline(ppt_path, job_dir)

        logger.info("[%s] Conversion completed successfully", job_id)
        final_video_path = os.path.join(job_dir, "final.mp4")
        update_job(
            job_id,
            status="completed",
            progress=100,
            message="Video ready for download",
            output=output_video,
            final_video_path=final_video_path
        )

    except Exception as e:
        logger.exception("[%s] Conversion failed", job_id)
        update_job(
            job_id,
            status="error",
            progress=100,
            message=f"Conversion failed: {str(e)}",
            output=None
        )
    finally:
        # Uninitialize COM
        pythoncom.CoUninitialize()


@app.post(
    "/convert",
    summary="Convert PPT to animated video with voice-over",
    description="Uploads a PPTX file and returns a job ID to track progress."
)
async def convert_ppt(file: UploadFile = File(...)):
    logger.info("Received PPT upload: %s", file.filename)
    job_id = str(uuid4())
    job_dir = os.path.join(JOBS_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    # Initialize job
    jobs[job_id] = {
        "status": "processing",
        "stage": "upload",
        "progress": 5,
        "message": "Saving PPT",
        "output": None,
        "log": str(JOBS_DIR / job_id / "status.log")
    }
    persist_status(job_id)
    append_log(job_id, "Job created")

    # Save uploaded file
    ppt_path = os.path.join(UPLOAD_DIR, f"{job_id}_{file.filename}")
    with open(ppt_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    append_log(job_id, f"Saved PPT: {file.filename}")

    # Start background thread
    thread = Thread(target=run_conversion_async, args=(job_id, ppt_path, job_dir))
    thread.daemon = True
    thread.start()

    logger.info("[%s] Background conversion started for: %s", job_id, file.filename)

    return {
        "job_id": job_id,
        "status_url": f"/status/{job_id}",
        "download_url": f"/download/{job_id}",
        "ui_url": f"/jobs/{job_id}"
    }


@app.post("/convert-ui")
async def convert_ppt_ui(request: Request, ppt: UploadFile = File(...)):
    result = await convert_ppt(ppt)
    return JSONResponse(result)


@app.get("/status/{job_id}")
async def get_status(job_id: str):
    job = jobs.get(job_id)
    if job:
        return job
    status_path = JOBS_DIR / job_id / "status.json"
    if status_path.exists():
        return json.loads(status_path.read_text(encoding="utf-8"))
    raise HTTPException(status_code=404, detail="Job not found")


@app.get("/logs/{job_id}")
def get_logs(job_id: str, download: bool = False):
    status_path = JOBS_DIR / job_id / "status.json"
    if not status_path.exists():
        raise HTTPException(status_code=404, detail="Job not found")
    data = json.loads(status_path.read_text(encoding="utf-8"))
    log_file = data.get("log")
    if not log_file or not Path(log_file).exists():
        return PlainTextResponse("")
    if download:
        return FileResponse(
            log_file,
            filename=f"{job_id}_logs.txt",
            media_type="text/plain"
        )
    return PlainTextResponse(Path(log_file).read_text(encoding="utf-8"))


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_page(request: Request, job_id: str):
    job_dir = f"jobs/{job_id}"
    status_file = f"{job_dir}/status.json"

    status = "processing"
    if os.path.exists(status_file):
        with open(status_file, encoding="utf-8") as f:
            status = json.load(f).get("status", "processing")

    return templates.TemplateResponse(
        "job.html",
        {
            "request": request,
            "job_id": job_id,
            "status": status
        }
    )


@app.get("/download/{job_id}")
def download_video(job_id: str):
    """
    Download the final MP4 for a completed job
    """
    video_path = os.path.join("jobs", job_id, "final.mp4")

    if not os.path.exists(video_path):
        return {
            "status": "error",
            "message": "Video not found. Job may not be completed yet."
        }

    return FileResponse(
        path=video_path,
        media_type="video/mp4",
        filename="presentation.mp4"
    )
