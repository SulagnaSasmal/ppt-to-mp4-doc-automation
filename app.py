import json
import logging
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from threading import Thread
from typing import Any, Dict, List
from uuid import uuid4

import pythoncom
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ppt_pipeline import (
    DEFAULT_PIPELINE_SETTINGS,
    extract_slide_notes,
    normalize_pipeline_settings,
    run_pipeline,
)

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "ppt_to_video.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

logger = logging.getLogger("ppt-video-api")

jobs: Dict[str, dict] = {}

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
JOBS_DIR = BASE_DIR / "jobs"
UPLOAD_DIR = BASE_DIR / "uploads"
ENV_FILE = BASE_DIR / ".env"

if not ENV_FILE.exists():
    raise RuntimeError(
        "Missing .env file. Copy .env.example to .env and set AZURE_TTS_KEY before starting the server."
    )

load_dotenv(ENV_FILE)

if not os.environ.get("AZURE_TTS_KEY") or not os.environ.get("AZURE_TTS_REGION"):
    raise RuntimeError(
        "Missing required Azure settings in .env. Set AZURE_TTS_KEY and AZURE_TTS_REGION."
    )

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

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
    version="1.1.0",
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
JOBS_DIR.mkdir(parents=True, exist_ok=True)


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_pipeline_settings(
    voice: str,
    speaking_rate: str,
    resolution: int,
    fps: int,
    quality: int,
) -> Dict[str, Any]:
    return normalize_pipeline_settings(
        {
            "voice": voice,
            "speaking_rate": speaking_rate,
            "resolution": resolution,
            "fps": fps,
            "quality": quality,
        }
    )


def persist_status(job_id: str) -> None:
    job = jobs.get(job_id)
    if not job:
        return
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    status_path = job_dir / "status.json"
    status_path.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")


def load_job(job_id: str) -> dict | None:
    if job_id in jobs:
        return jobs[job_id]
    status_path = JOBS_DIR / job_id / "status.json"
    if not status_path.exists():
        return None
    data = json.loads(status_path.read_text(encoding="utf-8"))
    jobs[job_id] = data
    return data


def list_recent_jobs(limit: int = 25) -> List[dict]:
    records: List[dict] = []
    for job_dir in JOBS_DIR.iterdir():
        if not job_dir.is_dir():
            continue
        status_path = job_dir / "status.json"
        if not status_path.exists():
            continue
        try:
            data = json.loads(status_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        data.setdefault("job_id", job_dir.name)
        records.append(data)

    records.sort(key=lambda item: item.get("updated_at", item.get("created_at", "")), reverse=True)
    return records[:limit]


def append_log(job_id: str, message: str) -> None:
    job = jobs.get(job_id)
    if not job:
        return
    log_path = job.get("log")
    if not log_path:
        return
    log_file = Path(log_path)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", encoding="utf-8") as file_obj:
        file_obj.write(message.rstrip() + "\n")


def update_job(
    job_id: str,
    status: str | None = None,
    progress: int | None = None,
    message: str | None = None,
    **kwargs,
) -> None:
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
    job["updated_at"] = utc_now_iso()
    persist_status(job_id)


def run_conversion_async(job_id: str, ppt_path: str, job_dir: str, settings: Dict[str, Any]) -> None:
    pythoncom.CoInitialize()

    try:
        def on_progress(stage: str, progress: int, message: str) -> None:
            update_job(job_id, stage=stage, progress=progress, message=message)

        result = run_pipeline(
            ppt_path,
            job_dir,
            settings=settings,
            progress_cb=on_progress,
        )

        output_video = result["final_video"]
        telemetry = result.get("telemetry", {})

        logger.info("[%s] Conversion completed successfully", job_id)
        update_job(
            job_id,
            status="completed",
            progress=100,
            message="Video ready for download",
            output=output_video,
            final_video_path=str(Path(job_dir) / "final.mp4"),
            telemetry=telemetry,
        )
    except Exception as exc:
        logger.exception("[%s] Conversion failed", job_id)
        update_job(
            job_id,
            status="error",
            progress=100,
            message=f"Conversion failed: {exc}",
            output=None,
        )
    finally:
        pythoncom.CoUninitialize()


@app.get("/", response_class=HTMLResponse)
def upload_page(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "default_settings": DEFAULT_PIPELINE_SETTINGS,
        },
    )


@app.get("/history", response_class=HTMLResponse)
def history_page(request: Request):
    records = list_recent_jobs(limit=50)
    return templates.TemplateResponse(
        "history.html",
        {
            "request": request,
            "jobs": records,
        },
    )


@app.get("/api/history")
def history_api(limit: int = 25):
    limit = max(1, min(100, limit))
    return {"jobs": list_recent_jobs(limit=limit)}


@app.post("/preview-notes")
async def preview_notes(
    ppt: UploadFile = File(...),
    voice: str = Form(DEFAULT_PIPELINE_SETTINGS["voice"]),
    speaking_rate: str = Form(DEFAULT_PIPELINE_SETTINGS["speaking_rate"]),
    resolution: int = Form(DEFAULT_PIPELINE_SETTINGS["resolution"]),
    fps: int = Form(DEFAULT_PIPELINE_SETTINGS["fps"]),
    quality: int = Form(DEFAULT_PIPELINE_SETTINGS["quality"]),
):
    settings = parse_pipeline_settings(voice, speaking_rate, resolution, fps, quality)

    preview_id = str(uuid4())
    preview_path = UPLOAD_DIR / f"preview_{preview_id}_{ppt.filename}"
    with preview_path.open("wb") as file_obj:
        shutil.copyfileobj(ppt.file, file_obj)

    try:
        notes = extract_slide_notes(str(preview_path))
        notes_with_text = [item for item in notes if item["has_notes"]]
        return {
            "slides_total": len(notes),
            "slides_with_notes": len(notes_with_text),
            "settings": settings,
            "notes": notes,
            "can_convert": len(notes_with_text) > 0,
        }
    finally:
        try:
            preview_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.post("/convert")
async def convert_ppt(
    ppt: UploadFile = File(...),
    voice: str = Form(DEFAULT_PIPELINE_SETTINGS["voice"]),
    speaking_rate: str = Form(DEFAULT_PIPELINE_SETTINGS["speaking_rate"]),
    resolution: int = Form(DEFAULT_PIPELINE_SETTINGS["resolution"]),
    fps: int = Form(DEFAULT_PIPELINE_SETTINGS["fps"]),
    quality: int = Form(DEFAULT_PIPELINE_SETTINGS["quality"]),
):
    logger.info("Received PPT upload: %s", ppt.filename)

    settings = parse_pipeline_settings(voice, speaking_rate, resolution, fps, quality)
    job_id = str(uuid4())
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    jobs[job_id] = {
        "job_id": job_id,
        "status": "processing",
        "stage": "upload",
        "progress": 5,
        "message": "Saving PPT",
        "output": None,
        "filename": ppt.filename,
        "settings": settings,
        "telemetry": {},
        "created_at": utc_now_iso(),
        "updated_at": utc_now_iso(),
        "log": str(job_dir / "status.log"),
    }
    persist_status(job_id)
    append_log(job_id, "Job created")

    ppt_path = UPLOAD_DIR / f"{job_id}_{ppt.filename}"
    with ppt_path.open("wb") as file_obj:
        shutil.copyfileobj(ppt.file, file_obj)
    append_log(job_id, f"Saved PPT: {ppt.filename}")

    thread = Thread(target=run_conversion_async, args=(job_id, str(ppt_path), str(job_dir), settings))
    thread.daemon = True
    thread.start()

    logger.info("[%s] Background conversion started for: %s", job_id, ppt.filename)
    return {
        "job_id": job_id,
        "status_url": f"/status/{job_id}",
        "download_url": f"/download/{job_id}",
        "logs_url": f"/logs/{job_id}",
        "history_url": "/history",
    }


@app.post("/convert-ui")
async def convert_ppt_ui(
    ppt: UploadFile = File(...),
    voice: str = Form(DEFAULT_PIPELINE_SETTINGS["voice"]),
    speaking_rate: str = Form(DEFAULT_PIPELINE_SETTINGS["speaking_rate"]),
    resolution: int = Form(DEFAULT_PIPELINE_SETTINGS["resolution"]),
    fps: int = Form(DEFAULT_PIPELINE_SETTINGS["fps"]),
    quality: int = Form(DEFAULT_PIPELINE_SETTINGS["quality"]),
):
    result = await convert_ppt(ppt, voice, speaking_rate, resolution, fps, quality)
    return JSONResponse(result)


@app.get("/status/{job_id}")
def get_status(job_id: str):
    job = load_job(job_id)
    if job:
        return job
    raise HTTPException(status_code=404, detail="Job not found")


@app.get("/logs/{job_id}")
def get_logs(job_id: str, download: bool = False):
    job = load_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")

    log_file = job.get("log")
    if not log_file or not Path(log_file).exists():
        return PlainTextResponse("")

    if download:
        return FileResponse(
            log_file,
            filename=f"{job_id}_logs.txt",
            media_type="text/plain",
        )

    return PlainTextResponse(Path(log_file).read_text(encoding="utf-8"))


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_page(request: Request, job_id: str):
    job = load_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")

    return templates.TemplateResponse(
        "job.html",
        {
            "request": request,
            "job_id": job_id,
            "status": job.get("status", "processing"),
        },
    )


@app.get("/download/{job_id}")
def download_video(job_id: str):
    video_path = JOBS_DIR / job_id / "final.mp4"
    if not video_path.exists():
        return {
            "status": "error",
            "message": "Video not found. Job may not be completed yet.",
        }

    return FileResponse(
        path=str(video_path),
        media_type="video/mp4",
        filename="presentation.mp4",
    )
