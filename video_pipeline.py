import json
import os
import re
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import requests

from ppt_pipeline import (
    ANTHROPIC_MODEL,
    call_claude,
    normalize_pipeline_settings,
    probe_duration_seconds,
    resolve_media_tool,
    safe_unlink,
    strip_code_fences,
    synthesize_azure_tts,
    synthesize_elevenlabs_tts,
)

AZURE_STT_API_VERSION = "2024-11-15"
MAX_TTS_CHUNK_CHARS = 1500
ATEMPO_MIN = 0.85
ATEMPO_MAX = 1.15
MIN_SEGMENT_SECONDS = 0.05


def transcribe_azure(video_path: Path, api_key: str, region: str, out_dir: Path) -> List[Dict[str, Any]]:
    """
    Extract the video's audio and transcribe it via Azure's Fast Transcription REST API.
    Returns a list of {text, start, end} phrase segments (seconds).
    """
    audio_path = out_dir / "source_audio.wav"
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    safe_unlink(audio_path)
    subprocess.run(
        [
            ffmpeg_exe, "-y", "-i", str(video_path),
            "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1",
            str(audio_path),
        ],
        check=True,
    )

    url = (
        f"https://{region}.api.cognitive.microsoft.com/speechtotext/transcriptions:transcribe"
        f"?api-version={AZURE_STT_API_VERSION}"
    )
    headers = {"Ocp-Apim-Subscription-Key": api_key}
    definition = {"locales": ["en-US"]}
    with audio_path.open("rb") as audio_file:
        files = {
            "audio": (audio_path.name, audio_file, "audio/wav"),
            "definition": (None, json.dumps(definition), "application/json"),
        }
        r = requests.post(url, headers=headers, files=files)
    r.raise_for_status()
    data = r.json()

    phrases = data.get("phrases") or data.get("combinedPhrases") or []
    segments = []
    for phrase in phrases:
        text = phrase.get("text", "")
        if not text:
            continue
        offset_ms = phrase.get("offsetMilliseconds", phrase.get("offset", 0))
        duration_ms = phrase.get("durationMilliseconds", phrase.get("duration", 0))
        start = offset_ms / 1000.0
        end = start + duration_ms / 1000.0
        segments.append({"text": text, "start": start, "end": end})
    return segments


def _validate_segments(segments: List[Dict[str, Any]], total_duration: float) -> List[Dict[str, Any]]:
    """Clamp/drop malformed AI-planned segments so downstream ffmpeg calls never see garbage timing."""
    cleaned = []
    for seg in segments:
        try:
            start = max(0.0, min(float(seg["start"]), total_duration))
            end = max(0.0, min(float(seg["end"]), total_duration))
            keep = bool(seg.get("keep", False))
        except (KeyError, TypeError, ValueError):
            continue
        if end - start <= MIN_SEGMENT_SECONDS:
            continue

        entry: Dict[str, Any] = {"start": start, "end": end, "keep": keep}
        if keep:
            narration = str(seg.get("narration") or "").strip()
            if not narration:
                continue
            entry["narration"] = narration
        cleaned.append(entry)

    cleaned.sort(key=lambda s: s["start"])
    if not any(s["keep"] for s in cleaned):
        raise ValueError("AI segment plan produced no usable kept segments.")
    return cleaned


def plan_segments(
    transcript: List[Dict[str, Any]],
    instruction: str,
    total_duration: float,
    api_key: str,
    model: str = ANTHROPIC_MODEL,
) -> List[Dict[str, Any]]:
    """
    Partition the video's timeline into keep/cut segments, writing narration text for each
    kept segment - grounded in the original transcript, shaped by the free-text instruction.

    If instruction is empty, skip the AI call entirely: return one segment spanning the whole
    video, narration = the original transcript's own text (a plain redub-in-new-voice pass).
    """
    if not instruction or not instruction.strip():
        full_text = " ".join(seg["text"].strip() for seg in transcript if seg.get("text", "").strip())
        return [{"start": 0.0, "end": total_duration, "keep": True, "narration": full_text or " "}]

    transcript_payload = [
        {"start": round(seg["start"], 2), "end": round(seg["end"], 2), "text": seg["text"]}
        for seg in transcript
    ]
    prompt = (
        "You are editing a narrated video. You're given its original timestamped transcript and "
        "an instruction describing what to change. Decide how to partition the video's full "
        f"timeline (0 to {total_duration:.2f} seconds) into an ordered list of segments.\n\n"
        f'Instruction: "{instruction}"\n\n'
        "Original timestamped transcript (JSON array of objects with start, end in seconds, and text):\n"
        f"{json.dumps(transcript_payload, ensure_ascii=False)}\n\n"
        "Return ONLY a JSON array of segment objects covering the entire timeline from 0 to "
        f"{total_duration:.2f} with no gaps and no overlaps, in order. Each object has:\n"
        '  "start" and "end" (seconds, numbers, grounded in the transcript\'s own timestamps),\n'
        '  "keep" (boolean - false means this span is cut entirely from the final video),\n'
        '  "narration" (required when keep is true - the text to be spoken by an AI voice for this '
        "span. By default, reuse the original transcript's own words for that span, lightly cleaned "
        "up (remove filler words like um/uh, false starts) - only substantially rewrite where the "
        "instruction specifically asks for it.)\n"
        "No commentary, no code fences - raw JSON array only."
    )
    raw_text = strip_code_fences(call_claude(prompt, api_key, model, max_tokens=8192))
    return _validate_segments(json.loads(raw_text), total_duration)


def resolve_plan(video_path: Path, instruction: str, out_dir: Path) -> List[Dict[str, Any]]:
    """Transcribe the source video and plan its segments (cuts + narration)."""
    azure_key = os.environ.get("AZURE_TTS_KEY")
    azure_region = os.environ.get("AZURE_TTS_REGION")
    if not azure_key or not azure_region:
        raise ValueError("AZURE_TTS_KEY/AZURE_TTS_REGION must be set in .env to transcribe the source video.")

    transcript = transcribe_azure(video_path, azure_key, azure_region, out_dir)
    total_duration = probe_duration_seconds(video_path)

    anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if instruction and instruction.strip() and not anthropic_key:
        raise ValueError("ANTHROPIC_API_KEY must be set in .env to apply an editing instruction.")

    return plan_segments(transcript, instruction, total_duration, anthropic_key)


def _chunk_script(text: str, max_chars: int = MAX_TTS_CHUNK_CHARS) -> List[str]:
    """Split narration text into TTS-provider-safe chunks, preferring paragraph/sentence breaks."""
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if not paragraphs:
        return [text]

    chunks: List[str] = []
    current = ""
    for para in paragraphs:
        candidate = f"{current}\n\n{para}".strip() if current else para
        if len(candidate) <= max_chars:
            current = candidate
            continue

        if current:
            chunks.append(current)
            current = ""

        if len(para) <= max_chars:
            current = para
            continue

        sentences = re.split(r"(?<=[.!?])\s+", para)
        for sent in sentences:
            cand = f"{current} {sent}".strip() if current else sent
            if len(cand) <= max_chars:
                current = cand
            else:
                if current:
                    chunks.append(current)
                current = sent

    if current:
        chunks.append(current)
    return chunks or [text]


def synthesize_narration(script_text: str, settings: Dict[str, Any], out_dir: Path, prefix: str = "narration") -> Path:
    """Synthesize narration text (chunked as needed) into one concatenated mp3."""
    chunks = _chunk_script(script_text)
    provider = settings["tts_provider"]
    chunk_paths = []

    if provider == "elevenlabs":
        api_key = os.environ.get("ELEVENLABS_API_KEY")
        if not api_key:
            raise ValueError("Missing ELEVENLABS_API_KEY in .env")
    else:
        api_key = os.environ.get("AZURE_TTS_KEY")
        region = os.environ.get("AZURE_TTS_REGION")
        if not api_key or not region:
            raise ValueError("Missing AZURE_TTS_KEY/AZURE_TTS_REGION in .env")

    for idx, chunk in enumerate(chunks):
        chunk_path = out_dir / f"{prefix}_chunk_{idx:03}.mp3"
        safe_unlink(chunk_path)
        if provider == "elevenlabs":
            synthesize_elevenlabs_tts(
                chunk, chunk_path, settings["elevenlabs_voice_id"], api_key, settings["elevenlabs_model"]
            )
        else:
            synthesize_azure_tts(chunk, chunk_path, settings["voice"], api_key, region, settings["speaking_rate"])
        chunk_paths.append(chunk_path)

    if len(chunk_paths) == 1:
        return chunk_paths[0]

    concat_list = out_dir / f"{prefix}_chunks.txt"
    with concat_list.open("w", encoding="utf-8") as f:
        for p in chunk_paths:
            f.write(f"file '{p.as_posix()}'\n")

    narration_path = out_dir / f"{prefix}_full.mp3"
    safe_unlink(narration_path)
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    subprocess.run(
        [ffmpeg_exe, "-y", "-f", "concat", "-safe", "0", "-i", str(concat_list), "-c", "copy", str(narration_path)],
        check=True,
    )
    return narration_path


def reconcile_and_mux(
    edited_video_path: Path,
    narration_path: Path,
    final_path: Path,
    out_dir: Path,
    tag: str = "narration",
) -> Dict[str, Any]:
    """
    Time-stretch the narration (ffmpeg atempo) to match the edited video's duration, clamped to
    a natural-sounding band, then mux video + adjusted narration into final_path.
    """
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    video_duration = probe_duration_seconds(edited_video_path)
    narration_duration = probe_duration_seconds(narration_path)

    tempo = narration_duration / video_duration if video_duration > 0 else 1.0
    tempo_clamped = max(ATEMPO_MIN, min(ATEMPO_MAX, tempo))

    adjusted_narration = out_dir / f"{tag}_adjusted.mp3"
    safe_unlink(adjusted_narration)
    subprocess.run(
        [
            ffmpeg_exe, "-y", "-i", str(narration_path),
            "-filter:a", f"atempo={tempo_clamped:.4f}",
            str(adjusted_narration),
        ],
        check=True,
    )

    safe_unlink(final_path)
    subprocess.run(
        [
            ffmpeg_exe, "-y",
            "-i", str(edited_video_path),
            "-i", str(adjusted_narration),
            "-map", "0:v:0", "-map", "1:a:0",
            "-c:v", "copy", "-c:a", "aac",
            str(final_path),
        ],
        check=True,
    )

    return {
        "video_duration": round(video_duration, 2),
        "narration_duration": round(narration_duration, 2),
        "tempo_requested": round(tempo, 4),
        "tempo_applied": round(tempo_clamped, 4),
        "clamped": abs(tempo - tempo_clamped) > 1e-6,
    }


def render_segment(
    video_path: Path,
    segment: Dict[str, Any],
    settings: Dict[str, Any],
    out_dir: Path,
    idx: int,
) -> Tuple[Path, Dict[str, Any]]:
    """Render one kept segment: cut its video span, synthesize+reconcile its narration, mux."""
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    start, end = segment["start"], segment["end"]
    tag = f"seg_{idx:03}"

    seg_video = out_dir / f"{tag}_video.mp4"
    safe_unlink(seg_video)
    subprocess.run(
        [
            ffmpeg_exe, "-y",
            "-ss", f"{start:.3f}", "-to", f"{end:.3f}",
            "-i", str(video_path),
            "-c:v", "libx264", "-c:a", "aac",
            "-avoid_negative_ts", "make_zero",
            str(seg_video),
        ],
        check=True,
    )

    narration_path = synthesize_narration(segment["narration"], settings, out_dir, prefix=f"{tag}_narr")

    seg_final = out_dir / f"{tag}_final.mp4"
    mux_info = reconcile_and_mux(seg_video, narration_path, seg_final, out_dir, tag=tag)
    return seg_final, mux_info


def run_video_pipeline(
    video_path: str,
    instruction: str,
    out_dir: str,
    settings: Optional[Dict[str, Any]] = None,
    progress_cb: Optional[Callable[[str, int, str], None]] = None,
    segments_override: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Runs:
    - Transcribe the source video's original narration (unless segments_override is supplied,
      e.g. from a prior /redub/preview call - then planning is skipped entirely)
    - Plan segments: what to cut, what to keep, and what narration to speak for each kept span
    - Render each kept segment: cut video span, synthesize+reconcile its narration, mux
    - Concatenate all rendered segments into final.mp4
    """
    video_path_obj = Path(video_path).resolve()
    out_dir_path = Path(out_dir).resolve()
    out_dir_path.mkdir(parents=True, exist_ok=True)
    pipeline_settings = normalize_pipeline_settings(settings)

    def report(stage: str, pct: int, message: str) -> None:
        if progress_cb:
            progress_cb(stage, pct, message)

    total_start = time.perf_counter()
    final_video = out_dir_path / "final.mp4"

    if segments_override is not None:
        segments = segments_override
        plan_seconds = 0.0
    else:
        report("plan", 15, "Transcribing original narration and planning edits")
        plan_start = time.perf_counter()
        segments = resolve_plan(video_path_obj, instruction, out_dir_path)
        plan_seconds = round(time.perf_counter() - plan_start, 2)

    kept_segments = [s for s in segments if s.get("keep")]
    if not kept_segments:
        raise ValueError("The plan kept no segments of the video - nothing to render.")

    report("render", 40, f"Rendering {len(kept_segments)} segment(s)")
    render_start = time.perf_counter()
    rendered_paths = []
    per_segment_mux_info = []
    step = 40 / max(1, len(kept_segments))
    for idx, seg in enumerate(kept_segments):
        report("render", int(40 + idx * step), f"Rendering segment {idx + 1}/{len(kept_segments)}")
        seg_path, mux_info = render_segment(video_path_obj, seg, pipeline_settings, out_dir_path, idx)
        rendered_paths.append(seg_path)
        per_segment_mux_info.append(mux_info)
    render_seconds = round(time.perf_counter() - render_start, 2)

    report("concat", 90, "Concatenating segments into final video")
    concat_start = time.perf_counter()
    safe_unlink(final_video)
    if len(rendered_paths) == 1:
        shutil.copy(rendered_paths[0], final_video)
    else:
        concat_list = out_dir_path / "final_segments.txt"
        with concat_list.open("w", encoding="utf-8") as f:
            for p in rendered_paths:
                f.write(f"file '{p.as_posix()}'\n")
        ffmpeg_exe = resolve_media_tool("ffmpeg")
        subprocess.run(
            [
                ffmpeg_exe, "-y",
                "-f", "concat", "-safe", "0",
                "-i", str(concat_list),
                "-c:v", "libx264", "-c:a", "aac",
                str(final_video),
            ],
            check=True,
        )
    concat_seconds = round(time.perf_counter() - concat_start, 2)

    total_seconds = round(time.perf_counter() - total_start, 2)

    return {
        "final_video": str(final_video),
        "telemetry": {
            "plan_seconds": plan_seconds,
            "render_seconds": render_seconds,
            "concat_seconds": concat_seconds,
            "total_seconds": total_seconds,
        },
        "segments": segments,
        "segment_mux_info": per_segment_mux_info,
        "settings": pipeline_settings,
    }
