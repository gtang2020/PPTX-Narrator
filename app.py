"""Free PowerPoint Narrator using Piper TTS."""

import io
import os
import wave
import zipfile
import tempfile
import subprocess
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from pptx import Presentation

app = FastAPI(title="Free PowerPoint Narrator")

ROOT = Path(__file__).parent
MAX_BYTES = 150 * 1024 * 1024

MODEL = ROOT / "models" / "en_US-lessac-low.onnx"


@app.get("/", response_class=HTMLResponse)
def home():
    return (ROOT / "index.html").read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {
        "status": "ok",
        "engine": "Piper TTS",
        "model_exists": MODEL.exists(),
        "api_key_required": False,
    }


def get_notes(slide):
    if not slide.has_notes_slide:
        return ""

    tf = slide.notes_slide.notes_text_frame
    return tf.text.strip() if tf is not None else ""


def synth(text):
    """Generate WAV audio using the local Piper model."""

    if not MODEL.exists():
        raise RuntimeError(
            "Piper voice model not found. Check Docker build."
        )

    with tempfile.NamedTemporaryFile(
        suffix=".wav", delete=False
    ) as f:
        output_path = f.name

    try:
        result = subprocess.run(
            [
                "piper",
                "--model", str(MODEL),
                "--output_file", output_path,
            ],
            input=text,
            text=True,
            capture_output=True,
            timeout=180,
        )

        if result.returncode != 0:
            raise RuntimeError(
                result.stderr[-500:] or "Piper failed"
            )

        audio = Path(output_path).read_bytes()

        if len(audio) < 100:
            raise RuntimeError("Generated audio is empty")

        return audio

    finally:
        Path(output_path).unlink(missing_ok=True)


def add_audio(slide, audio):
    """Embed WAV audio into a PowerPoint slide."""

    with tempfile.NamedTemporaryFile(
        suffix=".wav", delete=False
    ) as f:
        f.write(audio)
        filename = f.name

    try:
        slide.shapes.add_movie(
            filename,
            left=0,
            top=0,
            width=200000,
            height=200000,
            mime_type="audio/wav",
        )
    finally:
        Path(filename).unlink(missing_ok=True)


def download(data, filename, media_type):
    return StreamingResponse(
        io.BytesIO(data),
        media_type=media_type,
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        },
    )


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("existing"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Upload a PPTX file")

    content = await file.read(MAX_BYTES + 1)

    if len(content) > MAX_BYTES:
        raise HTTPException(413, "File exceeds 150 MB")

    try:
        prs = Presentation(io.BytesIO(content))
    except Exception:
        raise HTTPException(400, "Invalid PowerPoint file")

    if limit < 0:
        raise HTTPException(400, "Invalid slide limit")

    n = (
        min(limit, len(prs.slides))
        if limit > 0
        else len(prs.slides)
    )

    notes = [
        get_notes(prs.slides[i])
        for i in range(n)
    ]

    if mode == "notes":
        output = "\n\n".join(
            f"=== Slide {i+1} ===\n{note}"
            for i, note in enumerate(notes)
        )

        return download(
            output.encode("utf-8"),
            "slide_notes.txt",
            "text/plain; charset=utf-8",
        )

    if mode == "existing":
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            has_audio = any(
                name.startswith("ppt/media/")
                and name.lower().endswith(
                    (".mp3", ".m4a", ".wav")
                )
                for name in z.namelist()
            )

        if not has_audio:
            raise HTTPException(
                400, "No existing audio found"
            )

        if 0 < limit < len(prs.slides):
            raise HTTPException(
                400,
                "Existing Audio mode requires all slides"
            )

        return download(
            content,
            "existing_audio.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    if mode != "tts":
        raise HTTPException(400, "Invalid mode")

    if not any(notes):
        raise HTTPException(
            400, "No Speaker Notes found"
        )

    for i, note in enumerate(notes):
        if not note:
            continue

        try:
            audio = synth(note)
            add_audio(prs.slides[i], audio)
        except Exception as exc:
            raise HTTPException(
                500,
                f"Slide {i+1}: {str(exc)[:300]}"
            )

    output = io.BytesIO()
    prs.save(output)

    return download(
        output.getvalue(),
        "narrated_piper_free.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
