import json
import os
import re
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
MIN_KEEP_SEGMENT_SECONDS = 0.05

_TIME_RE = r"(?:\d{1,2}:)?\d{1,2}:\d{2}"
_CUT_PATTERN = re.compile(
    rf"\b(?:cut|remove|delete|trim)\b[^0-9]*({_TIME_RE})\s*(?:-|to|through|until|–|—)\s*({_TIME_RE})",
    re.IGNORECASE,
)


def _parse_timestamp(ts: str) -> float:
    parts = [int(p) for p in ts.split(":")]
    if len(parts) == 2:
        minutes, seconds = parts
        return minutes * 60 + seconds
    hours, minutes, seconds = parts
    return hours * 3600 + minutes * 60 + seconds


def parse_explicit_cuts(instruction: str) -> List[Tuple[float, float]]:
    """Pull explicit 'cut/remove/trim X-Y' timestamp ranges out of a free-text instruction."""
    cuts = []
    for match in _CUT_PATTERN.finditer(instruction):
        start = _parse_timestamp(match.group(1))
        end = _parse_timestamp(match.group(2))
        if end > start:
            cuts.append((start, end))
    return sorted(cuts)


def _has_vague_leftover(instruction: str) -> bool:
    """Heuristic: is there instruction text left over once explicit cut phrases are stripped?"""
    stripped = _CUT_PATTERN.sub("", instruction)
    stripped = re.sub(r"\b(cut|remove|delete|trim|and|also|the)\b", "", stripped, flags=re.IGNORECASE)
    stripped = re.sub(r"[,.\s]+", " ", stripped).strip()
    return len(stripped) > 15


def _merge_ranges(ranges: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    if not ranges:
        return []
    ranges = sorted(ranges)
    merged = [ranges[0]]
    for start, end in ranges[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end:
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))
    return merged


def _keep_segments(cuts: List[Tuple[float, float]], total_duration: float) -> List[Tuple[float, float]]:
    merged_cuts = _merge_ranges([(max(0.0, s), min(total_duration, e)) for s, e in cuts])
    keep = []
    cursor = 0.0
    for start, end in merged_cuts:
        if start > cursor:
            keep.append((cursor, start))
        cursor = max(cursor, end)
    if cursor < total_duration:
        keep.append((cursor, total_duration))
    return [(s, e) for s, e in keep if e - s > MIN_KEEP_SEGMENT_SECONDS]


def transcribe_azure(video_path: Path, api_key: str, region: str, out_dir: Path) -> List[Dict[str, Any]]:
    """
    Extract the video's audio and transcribe it via Azure's Fast Transcription REST API.
    Returns a list of {text, start, end} phrase segments (seconds). Only needed when a cut
    instruction has non-timestamp ("vague") content that requires AI judgment.
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


def ai_judgment_cuts(
    transcript: List[Dict[str, Any]],
    instruction: str,
    api_key: str,
    model: str = ANTHROPIC_MODEL,
) -> List[Tuple[float, float]]:
    """Ask Claude to pick cut ranges grounded in a timestamped transcript."""
    if not transcript:
        return []

    transcript_payload = [
        {"start": round(seg["start"], 2), "end": round(seg["end"], 2), "text": seg["text"]}
        for seg in transcript
    ]
    prompt = (
        "You are editing a video transcript to decide what to CUT (remove) from the final video.\n"
        f'Instruction: "{instruction}"\n\n'
        "Timestamped transcript (JSON array of objects with start, end in seconds, and text):\n"
        f"{json.dumps(transcript_payload, ensure_ascii=False)}\n\n"
        "Return ONLY a JSON array of objects with start and end (seconds, numbers) marking ranges "
        "to CUT. Ground start/end strictly in the transcript's own timestamps - never invent times "
        "outside the transcript's range. If nothing should be cut, return an empty array []. "
        "No commentary, no code fences - raw JSON array only."
    )
    raw_text = strip_code_fences(call_claude(prompt, api_key, model))
    cuts = [(float(c["start"]), float(c["end"])) for c in json.loads(raw_text)]
    return sorted(c for c in cuts if c[1] > c[0])


def resolve_cuts(
    video_path: Path,
    cut_instruction: str,
    out_dir: Path,
) -> Tuple[List[Tuple[float, float]], List[Dict[str, Any]]]:
    """
    Resolve a free-text cut instruction into concrete, merged (start, end) ranges.
    Explicit timestamps are parsed directly; any leftover vague instruction text triggers
    transcription + AI judgment. Returns (cut_ranges, transcript_used_or_empty).
    """
    if not cut_instruction or not cut_instruction.strip():
        return [], []

    explicit = parse_explicit_cuts(cut_instruction)
    transcript: List[Dict[str, Any]] = []

    if _has_vague_leftover(cut_instruction):
        azure_key = os.environ.get("AZURE_TTS_KEY")
        azure_region = os.environ.get("AZURE_TTS_REGION")
        if not azure_key or not azure_region:
            raise ValueError(
                "Cut instruction needs AI judgment (non-timestamp text found) but "
                "AZURE_TTS_KEY/AZURE_TTS_REGION are not set in .env for transcription."
            )
        anthropic_key = os.environ.get("ANTHROPIC_API_KEY")
        if not anthropic_key:
            raise ValueError("Cut instruction needs AI judgment but ANTHROPIC_API_KEY is not set in .env.")

        transcript = transcribe_azure(video_path, azure_key, azure_region, out_dir)
        ai_cuts = ai_judgment_cuts(transcript, cut_instruction, anthropic_key)
    else:
        ai_cuts = []

    total_duration = probe_duration_seconds(video_path)
    merged = _merge_ranges(explicit + ai_cuts)
    clamped = [(max(0.0, min(s, total_duration)), max(0.0, min(e, total_duration))) for s, e in merged]
    valid = [(s, e) for s, e in clamped if e > s]
    return _merge_ranges(valid), transcript


def apply_cuts(video_path: Path, cuts: List[Tuple[float, float]], out_path: Path, out_dir: Path) -> Path:
    """Cut the given ranges out of video_path, re-encoding for frame-accurate cuts."""
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    total_duration = probe_duration_seconds(video_path)
    keep = _keep_segments(cuts, total_duration)

    if not keep:
        raise ValueError("Cut instructions removed the entire video - nothing left to keep.")

    safe_unlink(out_path)

    if len(keep) == 1 and abs(keep[0][0]) < 1e-6 and abs(keep[0][1] - total_duration) < 1e-6:
        subprocess.run([ffmpeg_exe, "-y", "-i", str(video_path), "-c", "copy", str(out_path)], check=True)
        return out_path

    segment_paths = []
    for idx, (start, end) in enumerate(keep):
        seg_path = out_dir / f"keep_segment_{idx:03}.mp4"
        safe_unlink(seg_path)
        subprocess.run(
            [
                ffmpeg_exe, "-y",
                "-ss", f"{start:.3f}", "-to", f"{end:.3f}",
                "-i", str(video_path),
                "-c:v", "libx264", "-c:a", "aac",
                "-avoid_negative_ts", "make_zero",
                str(seg_path),
            ],
            check=True,
        )
        segment_paths.append(seg_path)

    concat_list = out_dir / "keep_segments.txt"
    with concat_list.open("w", encoding="utf-8") as f:
        for p in segment_paths:
            f.write(f"file '{p.as_posix()}'\n")

    subprocess.run(
        [
            ffmpeg_exe, "-y",
            "-f", "concat", "-safe", "0",
            "-i", str(concat_list),
            "-c:v", "libx264", "-c:a", "aac",
            str(out_path),
        ],
        check=True,
    )
    return out_path


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


def synthesize_narration(script_text: str, settings: Dict[str, Any], out_dir: Path) -> Path:
    """Synthesize the full narration script (chunked as needed) into one concatenated mp3."""
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
        chunk_path = out_dir / f"narration_chunk_{idx:03}.mp3"
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

    concat_list = out_dir / "narration_chunks.txt"
    with concat_list.open("w", encoding="utf-8") as f:
        for p in chunk_paths:
            f.write(f"file '{p.as_posix()}'\n")

    narration_path = out_dir / "narration_full.mp3"
    safe_unlink(narration_path)
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    subprocess.run(
        [ffmpeg_exe, "-y", "-f", "concat", "-safe", "0", "-i", str(concat_list), "-c", "copy", str(narration_path)],
        check=True,
    )
    return narration_path


def reconcile_and_mux(edited_video_path: Path, narration_path: Path, final_path: Path, out_dir: Path) -> Dict[str, Any]:
    """
    Time-stretch the narration (ffmpeg atempo) to match the edited video's duration, clamped to
    a natural-sounding band, then mux video + adjusted narration into the final MP4.
    """
    ffmpeg_exe = resolve_media_tool("ffmpeg")
    video_duration = probe_duration_seconds(edited_video_path)
    narration_duration = probe_duration_seconds(narration_path)

    tempo = narration_duration / video_duration if video_duration > 0 else 1.0
    tempo_clamped = max(ATEMPO_MIN, min(ATEMPO_MAX, tempo))

    adjusted_narration = out_dir / "narration_adjusted.mp3"
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


def run_video_pipeline(
    video_path: str,
    script_text: str,
    cut_instruction: str,
    out_dir: str,
    settings: Optional[Dict[str, Any]] = None,
    progress_cb: Optional[Callable[[str, int, str], None]] = None,
) -> Dict[str, Any]:
    """
    Runs:
    - Resolve cut instructions (explicit timestamps + AI judgment via transcript)
    - Apply cuts to the source video
    - Generate new narration audio from script_text
    - Reconcile durations (atempo) and mux into final.mp4
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

    if not script_text or not script_text.strip():
        raise ValueError("No narration script provided.")

    report("cuts", 10, "Resolving cut instructions")
    cuts_start = time.perf_counter()
    cut_ranges, _transcript = resolve_cuts(video_path_obj, cut_instruction, out_dir_path)
    cuts_seconds = round(time.perf_counter() - cuts_start, 2)

    report("editing", 30, "Applying cuts to source video")
    edit_start = time.perf_counter()
    edited_video = out_dir_path / "edited_video.mp4"
    apply_cuts(video_path_obj, cut_ranges, edited_video, out_dir_path)
    edit_seconds = round(time.perf_counter() - edit_start, 2)

    report("narration", 55, f"Generating {pipeline_settings['tts_provider']} narration")
    tts_start = time.perf_counter()
    narration_path = synthesize_narration(script_text, pipeline_settings, out_dir_path)
    tts_seconds = round(time.perf_counter() - tts_start, 2)

    report("mux", 85, "Reconciling durations and muxing final video")
    mux_start = time.perf_counter()
    mux_info = reconcile_and_mux(edited_video, narration_path, final_video, out_dir_path)
    mux_seconds = round(time.perf_counter() - mux_start, 2)

    total_seconds = round(time.perf_counter() - total_start, 2)

    return {
        "final_video": str(final_video),
        "telemetry": {
            "cuts_seconds": cuts_seconds,
            "edit_seconds": edit_seconds,
            "tts_seconds": tts_seconds,
            "mux_seconds": mux_seconds,
            "total_seconds": total_seconds,
        },
        "cuts": [{"start": round(s, 2), "end": round(e, 2)} for s, e in cut_ranges],
        "mux_info": mux_info,
        "settings": pipeline_settings,
    }
