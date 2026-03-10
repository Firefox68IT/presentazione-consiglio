---
name: travel-council-presentation
description: Build a PowerPoint travel presentation from dated stage folders, using Testo.txt for section intro slides, ordering photos and videos chronologically, generating 15-second clips, and configuring embedded videos to autoplay fullscreen after 3 seconds.
---

# Travel Council Presentation

Use this skill when the user wants a PowerPoint presentation generated from a trip folder structured as dated subfolders like `YYYYMMDD Place`, with photos, videos, and optional `Testo.txt` files for the section intro slides.

## What this skill does

- Creates a cover slide plus one intro slide per stage folder.
- Uses `Testo.txt` inside each stage folder when present.
- Orders photos and videos chronologically inside each section.
- Generates 15-second clips from source videos.
- Embeds the clip in PowerPoint, places it inside the correct section, sets fullscreen playback, and delays autoplay by 3 seconds.
- Writes the output `.pptx` and a summary text file in the project folder.

## Expected project structure

The project root should look like this:

```text
project-root/
  MATERIALE...pdf
  20260206 New Delhi Ashalayam tempio/
    Testo.txt
    IMG....jpg
    VID....mp4
  20260207 Another place/
    ...
```

Stage folders are detected from names starting with `YYYYMMDD`.

## Requirements

- Windows
- Microsoft PowerPoint installed locally
- `ffmpeg` available either:
  - in `PATH`, or
  - in `C:\Users\<user>\tools\ffmpeg\bin\ffmpeg.exe`

## Workflow

1. Open the project root.
2. Run the bundled PowerShell script:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .\skills\travel-council-presentation\scripts\generate_presentation.ps1`
3. If needed, pass custom parameters:
   `-ProjectRoot <path> -OutputName <name>.pptx`
4. After generation, verify the resulting deck visually in PowerPoint.

## Notes

- This workflow relies on PowerPoint COM automation. If COM is unavailable in the current session, rerun it in a session where Office automation works.
- If a folder has no media, the skill still creates the intro slide for that stage.
- The output clip folder defaults to `clip_video_15s` under the project root.
- Keep all slide coordinates within the actual PowerPoint canvas. For widescreen in this workflow, the script measures and uses the real slide size instead of assuming a larger layout.

## Bundled resources

- Script: `scripts/generate_presentation.ps1`
- UI metadata: `agents/openai.yaml`
