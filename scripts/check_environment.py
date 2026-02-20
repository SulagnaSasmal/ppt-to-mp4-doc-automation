import os
import sys
from pathlib import Path

from dotenv import load_dotenv


def check_python() -> tuple[bool, str]:
    major, minor = sys.version_info[:2]
    if major != 3 or minor < 11:
        return False, f"Python 3.11+ required, found {major}.{minor}"
    return True, f"Python {major}.{minor}"


def check_env_vars() -> tuple[bool, str]:
    key = os.environ.get("AZURE_TTS_KEY")
    region = os.environ.get("AZURE_TTS_REGION")
    if not key or not region:
        return False, "AZURE_TTS_KEY/AZURE_TTS_REGION missing in environment (.env not loaded)"
    return True, "Azure env vars present"


def check_ffmpeg() -> tuple[bool, str]:
    try:
        from ppt_pipeline import resolve_media_tool

        ffmpeg = resolve_media_tool("ffmpeg")
        ffprobe = resolve_media_tool("ffprobe")
        return True, f"ffmpeg: {ffmpeg} | ffprobe: {ffprobe}"
    except Exception as exc:
        return False, str(exc)


def check_powerpoint() -> tuple[bool, str]:
    try:
        import win32com.client

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Quit()
        return True, "PowerPoint COM available"
    except Exception as exc:
        return False, f"PowerPoint COM unavailable: {exc}"


def main() -> int:
    workspace = Path(__file__).resolve().parents[1]
    os.chdir(workspace)
    if str(workspace) not in sys.path:
        sys.path.insert(0, str(workspace))
    load_dotenv(workspace / ".env")

    checks = [
        ("Python", check_python),
        ("Azure env", check_env_vars),
        ("FFmpeg", check_ffmpeg),
        ("PowerPoint", check_powerpoint),
    ]

    failed = False
    for label, fn in checks:
        ok, msg = fn()
        marker = "[OK]" if ok else "[FAIL]"
        print(f"{marker} {label}: {msg}")
        failed = failed or (not ok)

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
