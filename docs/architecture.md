# Architecture

## Goal
Convert PPT/PPTX files into downloadable MP4 videos using an API-first pipeline with asynchronous job tracking.

## Components
- **UI Layer**: HTML pages in `templates/` for upload, status polling, and download.
- **API Layer**: FastAPI app (`app.py`) exposes convert/status/download/log endpoints.
- **Pipeline Engine**: `ppt_pipeline.py` orchestrates rendering, narration, and muxing.
- **Storage**: `uploads/`, `build/`, and `jobs/` hold inputs, intermediate assets, outputs, and status.
- **External Services**: PowerPoint COM automation, Azure TTS, FFmpeg.

## Processing Flow
1. User uploads PPT via UI/API.
2. Server creates a `job_id` and status file under `jobs/<job_id>/status.json`.
3. Pipeline renders slides, generates audio, and assembles MP4.
4. Logs and status are updated during processing.
5. User downloads output once job state is `completed`.

## Runtime Notes
- Windows environment is required for PowerPoint COM automation.
- Jobs run asynchronously to avoid blocking the API thread.
- Filesystem-backed job artifacts allow simple traceability and troubleshooting.