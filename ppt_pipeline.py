import win32com.client
import subprocess
import os
import time
from pathlib import Path
import requests

# Use numeric constant to avoid makepy dependency when constants are unavailable
PPT_ADVANCE_USE_TIMINGS = 3  # ppSlideShowUseSlideTimings
DEFAULT_SILENCE_HEAD = 0.0
DEFAULT_SILENCE_TAIL = 0.0
VOICE = "en-US-JennyNeural"


def probe_duration_seconds(path: Path) -> float:
    """Return media duration in seconds using ffprobe."""
    result = subprocess.check_output([
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=nk=1:nw=1",
        str(path)
    ], text=True).strip()
    return float(result)


def wait_for_readable(path: Path, attempts: int = 30, delay: float = 1.0) -> None:
    """Wait until a file can be opened for reading (handles lingering locks)."""
    for _ in range(attempts):
        try:
            with open(path, "rb"):
                return
        except OSError:
            time.sleep(delay)
    raise FileNotFoundError(f"File not readable (locked?): {path}")


def safe_unlink(path: Path, attempts: int = 5, delay: float = 0.5) -> None:
    """Delete a file with retries to handle transient locks."""
    for i in range(attempts):
        if not path.exists():
            return
        try:
            path.unlink()
            return
        except OSError:
            time.sleep(delay)
    raise PermissionError(f"Could not delete locked file: {path}")

def run_pipeline(ppt_path: str, out_dir: str) -> str:
    """
    Runs:
    - PPT → animated MP4
    - Notes → Azure TTS
    - FFmpeg mux
    Returns final MP4 path
    """
    from pptx import Presentation

    ppt_path_obj = Path(ppt_path).resolve()
    out_dir_path = Path(out_dir).resolve()
    out_dir_path.mkdir(parents=True, exist_ok=True)

    video_raw = out_dir_path / "video_raw.mp4"
    narration = out_dir_path / "narration.mp3"
    final_video = out_dir_path / "final.mp4"

    # Check for required Azure credentials
    azure_key = os.environ.get("AZURE_TTS_KEY")
    azure_region = os.environ.get("AZURE_TTS_REGION")
    
    if not azure_key or not azure_region:
        raise ValueError(
            "Missing Azure TTS credentials. Please set environment variables:\n"
            "  AZURE_TTS_KEY=<your-key>\n"
            "  AZURE_TTS_REGION=<region> (e.g., 'eastus')\n"
            "Before starting the server."
        )
    
    tts_url = f"https://{azure_region}.tts.speech.microsoft.com/cognitiveservices/v1"
    tts_headers = {
        "Ocp-Apim-Subscription-Key": azure_key,
        "Content-Type": "application/ssml+xml",
        "X-Microsoft-OutputFormat": "audio-24khz-160kbitrate-mono-mp3",
    }

    # 1️⃣ Extract notes text (python-pptx)
    prs = Presentation(str(ppt_path_obj))
    slide_notes = []
    for i, slide in enumerate(prs.slides, start=1):
        if slide.notes_slide:
            text = slide.notes_slide.notes_text_frame.text.strip()
        else:
            text = ""
        slide_notes.append((i, text))

    # 2️⃣ Azure TTS per slide
    audio_files = []
    audio_durations = {}
    for i, text in slide_notes:
        if not text:
            continue

        ssml = f"""
<speak version='1.0' xml:lang='en-US'>
  <voice name='{VOICE}'>{text}</voice>
</speak>
"""

        out_mp3 = (out_dir_path / f"slide_{i:02}.mp3").resolve()
        safe_unlink(out_mp3)
        r = requests.post(tts_url, headers=tts_headers, data=ssml.encode("utf-8"))
        r.raise_for_status()
        out_mp3.write_bytes(r.content)
        audio_files.append(out_mp3)
        audio_durations[i] = probe_duration_seconds(out_mp3)

    # 3️⃣ Open PPT and set slide timings to match narration durations (plus padding)
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    presentation = ppt.Presentations.Open(
        str(ppt_path_obj),
        WithWindow=True
    )

    for slide_index, slide in enumerate(presentation.Slides, start=1):
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceOnClick = False
        slide.SlideShowTransition.Duration = 0.0
        duration = audio_durations.get(slide_index, 2.0) + DEFAULT_SILENCE_HEAD + DEFAULT_SILENCE_TAIL
        slide.SlideShowTransition.AdvanceTime = duration

    presentation.SlideShowSettings.AdvanceMode = PPT_ADVANCE_USE_TIMINGS
    presentation.SlideShowSettings.ShowWithAnimation = True
    presentation.SlideShowSettings.ShowWithNarration = True

    # Save a clean copy to avoid compatibility/protection quirks, then reopen it
    tmp_pptx = (out_dir_path / f"tmp_for_video_{int(time.time())}.pptx").resolve()
    safe_unlink(tmp_pptx)
    try:
        presentation.SaveCopyAs(str(tmp_pptx))
    except Exception as e:
        presentation.Close()
        ppt.Quit()
        raise RuntimeError(f"PowerPoint could not save a temp copy: {e}") from e

    presentation.Close()
    presentation = ppt.Presentations.Open(str(tmp_pptx), WithWindow=True)

    # Re-apply timings on reopened copy
    for slide_index, slide in enumerate(presentation.Slides, start=1):
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceOnClick = False
        slide.SlideShowTransition.Duration = 0.0
        duration = audio_durations.get(slide_index, 2.0) + DEFAULT_SILENCE_HEAD + DEFAULT_SILENCE_TAIL
        slide.SlideShowTransition.AdvanceTime = duration

    presentation.SlideShowSettings.AdvanceMode = PPT_ADVANCE_USE_TIMINGS
    presentation.SlideShowSettings.ShowWithAnimation = True
    presentation.SlideShowSettings.ShowWithNarration = True

    # Persist timings before rendering
    presentation.Save()

    safe_unlink(video_raw)
    safe_unlink(narration)
    safe_unlink(final_video)

    # 4️⃣ Export video with slide-based timings
    presentation.CreateVideo(
        str(video_raw),
        UseTimingsAndNarrations=True,
        DefaultSlideDuration=1,
        VertResolution=1080,
        FramesPerSecond=30,
        Quality=100
    )

    status = presentation.CreateVideoStatus
    while status == 1:  # In progress
        time.sleep(2)
        status = presentation.CreateVideoStatus

    if status != 3:  # 3 = Success
        # Fallback to SaveAs MP4 (ppSaveAsMP4 = 39). Some builds reject CreateVideo but accept SaveAs.
        try:
            presentation.SaveAs(str(video_raw), FileFormat=39)
            for _ in range(30):  # wait up to ~60s
                if video_raw.exists():
                    status = 3
                    break
                time.sleep(2)
        except Exception as e:
            presentation.Close()
            ppt.Quit()
            raise RuntimeError(
                "PowerPoint failed to render video (status={}). SaveAs fallback also failed: {}".format(status, e)
            ) from e

    if status != 3:
        presentation.Close()
        ppt.Quit()
        raise RuntimeError(
            "PowerPoint failed to render video (status={}). "
            "Tried CreateVideo and SaveAs MP4. Check for hidden dialogs or re-run once manually from UI to unblock.".format(status)
        )

    if not video_raw.exists():
        raise FileNotFoundError(f"PowerPoint did not produce the video file: {video_raw}")
    print("[OK] Animated video created:", video_raw)

    # Wait until the MP4 is actually written and non-empty
    for _ in range(60):  # up to ~120s
        try:
            size = video_raw.stat().st_size
        except FileNotFoundError:
            size = 0
        if size > 0:
            break
        time.sleep(2)
    else:
        try:
            # Fallback: try SaveAs MP4 if CreateVideo wrote a zero-byte file
            presentation.SaveAs(str(video_raw), FileFormat=39)
            for _ in range(30):
                try:
                    size = video_raw.stat().st_size
                except FileNotFoundError:
                    size = 0
                if size > 0:
                    break
                time.sleep(2)
        except Exception:
            presentation.Close()
            ppt.Quit()
            raise FileNotFoundError(f"PowerPoint produced an empty/locked file and SaveAs retry failed: {video_raw}")

        if size <= 0:
            presentation.Close()
            ppt.Quit()
            raise FileNotFoundError(f"PowerPoint produced an empty/locked file: {video_raw}")

    print(f"[OK] Video file size: {video_raw.stat().st_size} bytes")

    # Ensure MP4 is readable (release any lingering lock)
    wait_for_readable(video_raw, attempts=30, delay=1.0)

    # Log intended total timed duration vs audio sum for sanity
    total_slide_seconds = sum(audio_durations.get(i, 2.0) + DEFAULT_SILENCE_HEAD + DEFAULT_SILENCE_TAIL for i, _ in slide_notes)
    try:
        video_seconds = probe_duration_seconds(video_raw)
        print(f"[INFO] Intended slide duration sum: {total_slide_seconds:.2f}s; Rendered video: {video_seconds:.2f}s")
    except Exception:
        pass

    # Close PowerPoint with retry and error handling
    try:
        presentation.Close()
    except Exception as e:
        print(f"[WARNING] Failed to close presentation cleanly: {e}")
        # Try to close without saving
        try:
            presentation.Saved = True  # Mark as saved to avoid prompts
            presentation.Close()
        except Exception:
            pass
    
    try:
        ppt.Quit()
    except Exception as e:
        print(f"[WARNING] Failed to quit PowerPoint cleanly: {e}")

    # 5️⃣ Concatenate audio clips
    audio_list = (out_dir_path / "audio_list.txt").resolve()
    with audio_list.open("w", encoding="utf-8") as f:
        for a in audio_files:
            f.write(f"file '{a.as_posix()}'\n")

    subprocess.run([
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", str(audio_list),
        "-c", "copy",
        str(narration)
    ], check=True)

    # Ensure narration file is present before mux
    if not narration.exists():
        raise FileNotFoundError(f"Narration not found: {narration}")

    # 6️⃣ FFmpeg mux (KEEP ANIMATIONS)
    subprocess.run([
        "ffmpeg", "-y",
        "-i", str(video_raw),
        "-i", str(narration),
        "-map", "0:v:0",
        "-map", "1:a:0",
        "-c:v", "copy",
        "-c:a", "aac",
        "-shortest",
        str(final_video)
    ], check=True)

    print("[OK] DONE ->", final_video)
    return str(final_video)
