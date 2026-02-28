# PPT to Video — Demo Guide

> This guide is written for non-technical use. No coding knowledge needed.
> Come back to this file any time you need to run or restart the demo.

---

## What this tool does

You upload a PowerPoint file → it comes back as a narrated MP4 video.
The narration is read from your slide notes using an AI voice (Azure).

---

## Section 1 — First-Time Setup (do this once only)

You only need to do these steps the first time. After that, skip to Section 2.

### Step 1 — Install Python

1. Go to https://python.org/downloads
2. Download Python 3.11 or newer
3. Run the installer — **tick "Add Python to PATH"** before clicking Install

### Step 2 — Install FFmpeg

1. Open the Start menu → search for **winget** or open **Command Prompt**
2. Paste this and press Enter:
   ```
   winget install ffmpeg
   ```
3. Close and reopen any terminals after this

### Step 3 — Install Cloudflare Tunnel

1. In Command Prompt, paste and press Enter:
   ```
   winget install Cloudflare.cloudflared
   ```

### Step 4 — Set up the project

1. Open Command Prompt
2. Go to this folder:
   ```
   cd /d d:\ppt-video-phase1-API
   ```
3. Create a virtual environment (one-time):
   ```
   python -m venv .venv
   ```
4. Install dependencies:
   ```
   .venv\Scripts\pip install -r requirements.txt
   ```

### Step 5 — Set up your Azure key

1. In the project folder, find the file called `.env.example`
2. Make a copy of it and rename the copy to `.env`
3. Open `.env` in Notepad
4. Fill in your Azure key and region:
   ```
   AZURE_TTS_KEY=paste-your-azure-key-here
   AZURE_TTS_REGION=eastus
   ```
5. Save and close

---

## Section 2 — Running a Demo (every time)

### The easy way — one double-click

1. Open **File Explorer**
2. Go to `d:\ppt-video-phase1-API`
3. Double-click **`start-demo.bat`**

That's it. Two windows will open automatically:

| Window colour | What it is | What to do |
|---|---|---|
| **Blue** | App server | Keep it open. Don't close it. |
| **Yellow** | Your public URL | Look here for the demo link |

In the **yellow window**, look for a line like:
```
Your quick Tunnel has been created! Visit it at:
https://something-something.trycloudflare.com
```

**Copy that URL** — that's the link you share with your audience.

### To use the app yourself

Your browser will also open automatically to `http://localhost:8000` — use this if you're demonstrating from your own screen.

---

## Section 3 — During the Demo

### How to convert a file

1. Open the app in a browser (the trycloudflare URL or localhost)
2. Click **Choose File** → select a `.pptx` file
3. Leave the settings as they are (defaults are fine)
4. Click **Convert**
5. Wait — it takes about 1–3 minutes depending on slide count
6. When done, a **Download** button appears → click to save the MP4

### What your PowerPoint needs

- Slide **notes** must be written (the text in the notes panel at the bottom)
- The notes become the narration — what the AI voice will say
- Slides with no notes will be skipped for audio

---

## Section 4 — Stopping the Demo

When you're done:

1. Close the **yellow** Cloudflare tunnel window → the public URL stops working
2. Close the **blue** server window → the app goes offline
3. Done

---

## Section 5 — Restarting After a Break

If you closed everything and want to demo again:

1. Double-click `start-demo.bat` again
2. You'll get a **new** public URL (different from last time — that's normal)
3. Everything else works the same

---

## Section 6 — Troubleshooting

### "The app won't start" / Blue window closes immediately

- Check that your `.env` file exists and has your Azure key filled in
- Make sure you ran the setup steps in Section 1

### "cloudflared is not recognized"

- Cloudflare Tunnel isn't installed yet
- Run: `winget install Cloudflare.cloudflared` in Command Prompt
- Close and reopen the terminal, then try again

### "The page won't load at the trycloudflare URL"

- The tunnel takes about 30 seconds to become reachable after it starts
- Wait and refresh the browser

### "Conversion failed — no notes found"

- Your PowerPoint has no text in the Notes panel
- Open the file in PowerPoint, click a slide, write something in the Notes section at the bottom, save, and try again

### "Azure TTS error"

- Your Azure key is wrong or expired
- Open `.env` in Notepad and check `AZURE_TTS_KEY` and `AZURE_TTS_REGION`

---

## Quick reference card

| What you want to do | How |
|---|---|
| Start the demo | Double-click `start-demo.bat` |
| Find the public URL | Look at the yellow window |
| Use the app locally | Go to `http://localhost:8000` |
| Stop the demo | Close both windows |
| Restart | Double-click `start-demo.bat` again |

---

*Last updated: February 2026*
