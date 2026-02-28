<div align="center">

# PPT → MP4 Documentation Automation

**Convert PowerPoint presentations into narrated MP4 videos — fully automated.**

No screen recording. No manual voiceover. Just upload a `.pptx` and download a finished `.mp4`.

[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.100+-009688?style=flat-square&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com)
[![Azure TTS](https://img.shields.io/badge/Azure-Text--to--Speech-0078D4?style=flat-square&logo=microsoft-azure&logoColor=white)](https://azure.microsoft.com/en-us/products/ai-services/text-to-speech)
[![FFmpeg](https://img.shields.io/badge/FFmpeg-video%20mux-007808?style=flat-square&logo=ffmpeg&logoColor=white)](https://ffmpeg.org)
[![Platform](https://img.shields.io/badge/Platform-Windows%20only-0078D4?style=flat-square&logo=windows&logoColor=white)](https://github.com/SulagnaSasmal/ppt-to-mp4-doc-automation)

</div>

---

## What it does

Documentation teams spend hours recording walkthrough videos manually — screen capture, voiceover recording, editing, re-recording when slides change.

This tool eliminates all of that.

Write your slide notes once. Run the pipeline. Get a narrated, animated MP4 video back — synced to your slides, with AI voice, ready to publish.

---

## How the pipeline works

```
Upload .pptx file
  │
  ├─► [1] Extract slide notes        (python-pptx)
  │         ↓
  ├─► [2] Generate AI narration      (Azure TTS — Jenny Neural voice)
  │         ↓ per-slide .mp3 files
  ├─► [3] Set slide timings          (PowerPoint COM automation via win32com)
  │         ↓ timings matched to audio duration
  ├─► [4] Export animated video      (PowerPoint CreateVideo API)
  │         ↓ video_raw.mp4
  ├─► [5] Mux video + audio          (FFmpeg)
  │         ↓
  └─► Download final.mp4
```

---

## Features

- **Web UI** — Upload PPT, track job progress in real time, download finished video
- **AI narration** — Azure Text-to-Speech with Jenny Neural voice (natural, clear, professional)
- **Slide animations preserved** — Uses PowerPoint's own export engine, not screenshots
- **Per-slide timing** — Each slide advances exactly when the narration finishes
- **Job history** — Review all past conversions with status and telemetry
- **Preview notes** — See extracted slide notes before committing to conversion
- **Configurable** — Adjust voice, speaking rate, resolution (up to 1080p), FPS, quality

---

## Tech stack

| Component | Technology |
|---|---|
| Web framework | FastAPI + Jinja2 templates |
| PPT note extraction | python-pptx |
| AI voice generation | Azure Cognitive Services — Text-to-Speech |
| PowerPoint control | win32com (PowerPoint COM automation) |
| Video muxing | FFmpeg |
| Job tracking | Thread-based async with JSON persistence |
| Environment | Python 3.11+ on Windows (COM automation requires Windows + PowerPoint) |

---

## Requirements

- Windows 10 or 11
- Microsoft PowerPoint (installed, licensed)
- Python 3.11+
- FFmpeg (installed and on PATH)
- Azure Cognitive Services key (Text-to-Speech)

---

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/SulagnaSasmal/ppt-to-mp4-doc-automation.git
cd ppt-to-mp4-doc-automation

# 2. Create a virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment
copy .env.example .env
# Edit .env — set AZURE_TTS_KEY and AZURE_TTS_REGION

# 5. Start the server
.\start.ps1
# or: .venv\Scripts\python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Open `http://localhost:8000` in your browser.

---

## Running a demo (Cloudflare public URL)

To share a live demo without deploying to a server:

```bash
# In terminal 1 — start the app
.\start.ps1

# In terminal 2 — open a public tunnel
cloudflared tunnel --url http://localhost:8000
```

You'll get a public URL like `https://xxxx.trycloudflare.com` — share it for your demo.
Close the terminal to take it offline.

> See [DEMO-GUIDE.md](DEMO-GUIDE.md) for step-by-step instructions (no technical knowledge needed).

---

## How to prepare your PowerPoint

The narration is generated from your **slide notes** — the text panel at the bottom of each slide in PowerPoint.

- Slides **with** notes → AI voice reads the notes as narration
- Slides **without** notes → slide plays silently (default 2 seconds)
- Notes can be as long as needed — slide timing adjusts automatically

---

## API endpoints

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | Upload UI |
| `POST` | `/convert` | Submit a PPT for conversion (async) |
| `GET` | `/status/{job_id}` | Poll job progress (0–100%) |
| `GET` | `/download/{job_id}` | Download finished MP4 |
| `POST` | `/preview-notes` | Extract and preview slide notes before converting |
| `GET` | `/history` | Job history UI |
| `GET` | `/logs/{job_id}` | View conversion logs |

---

## Use cases

| Who | What they use it for |
|---|---|
| Documentation teams | Feature walkthrough videos for release notes |
| Product managers | Customer enablement and onboarding content |
| Training teams | Narrated training modules from existing decks |
| Demo engineers | Quick product demo videos without recording setup |

---

## Why I built this

Technical Writers spend significant time producing video documentation manually — recording screens, narrating slides, editing, re-recording after updates. This tool removes that overhead entirely: update your slide notes, re-run the pipeline, get a new video.

It also demonstrates something I believe strongly: **technical writers should be able to build the tools their teams depend on**, not just document them.

---

## Author

**Sulagna Sasmal** — Documentation Architect & Senior Technical Writer
[Portfolio](https://sulagnasasmal.github.io/sulagnasasmal-site/) · [GitHub](https://github.com/SulagnaSasmal) · [LinkedIn](https://www.linkedin.com/in/sulagnasasmal)
