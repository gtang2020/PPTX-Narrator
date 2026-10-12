"""iPad PowerPoint Narrator: FastAPI + python-pptx + free Edge TTS.

Edge TTS uses an unofficial online service and can be rate-limited or changed.
"""
import io
import os
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import quote

import edge_tts
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from pptx import Presentation

app = FastAPI(title="iPad PowerPoint Narrator")

ROOT = Path(__file__).parent
MAX_BYTES = 150 * 1024 * 1024
MAX_SLIDE_CHARS = 15000

VOICE_MAP = {
    "alloy": "en-US-AriaNeural",
    "echo": "en-US-GuyNeural",
    "fable": "en-GB-SoniaNeural",
    "onyx": "en-US-ChristopherNeural",
    "nova": "en-US-JennyNeural",
    "shimmer": "en-US-AriaNeural",
}

ALLOWED_VOICES = {
    "en-US-AriaNeural",
    "en-US-JennyNeural",
    "en-US-GuyNeural",
    "en-US-ChristopherNeural",
    "en-GB-SoniaNeural",
    "zh-CN-XiaoxiaoNeural",
    "zh-CN-YunxiNeural",
}


@app.get("/", response_class=HTMLResponse)
def home():
    return (ROOT / "index.html").read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {
        "status": "ok",
        "speech_provider": "edge-tts",
        "api_key_required": False,
    }


def get_notes(slide):
    if not slide.has_notes_slide:
        return ""

    tf = slide.notes_slide.notes_text_frame

    return tf.text.strip() if tf is not None else ""


def add_audio(slide, data, suffix=".mp3"):
    """Embed audio in the PowerPoint slide."""

    with tempfile.NamedTemporaryFile(
        suffix=suffix,
        delete=False
    ) as f:
        f.write(data)
        fn = f.name

    try:
        return slide.shapes.add_movie(
            fn,
            left=0,
            top=0,
            width=200000,
            height=200000,
            poster_frame_image=None,
            mime_type="audio/mpeg",
        )
    finally:
        os.unlink(fn)


async def synth(text, voice):
    """Generate MP3 narration without an OpenAI API key."""

    voice = VOICE_MAP.get(voice, voice)

    if voice not in ALLOWED_VOICES:
        raise HTTPException(
            400,
            f"Unsupported voice: {voice}"
        )

    if len(text) > MAX_SLIDE_CHARS:
        raise HTTPException(
            400,
            f"Slide notes exceed {MAX_SLIDE_CHARS} characters"
        )

    with tempfile.NamedTemporaryFile(
        suffix=".mp3",
        delete=False
    ) as f:
        path = f.name

    try:
        communicate = edge_tts.Communicate(
            text=text,
            voice=voice
        )

        await communicate.save(path)

        data = Path(path).read_bytes()

        if not data:
            raise RuntimeError(
                "Speech service returned empty audio"
            )

        return data

    except Exception as exc:
        raise HTTPException(
            502,
            f"Free speech generation failed: {str(exc)[:250]}"
        ) from exc

    finally:
        Path(path).unlink(missing_ok=True)


def download(data, filename, content_type):
    return StreamingResponse(
        io.BytesIO(data),
        media_type=content_type,
        headers={
            "Content-Disposition":
            f"attachment; filename*=UTF-8''{quote(filename)}"
        },
    )


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("existing"),
    voice: str = Form("en-US-AriaNeural"),
    api_key: str = Form(""),
    limit: int = Form(0),
):

    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(
            400,
            "Upload a .pptx file"
        )

    content = await file.read(MAX_BYTES + 1)

    if len(content) > MAX_BYTES:
        raise HTTPException(
            413,
            "File exceeds 150 MB limit"
        )

    try:
        prs = Presentation(io.BytesIO(content))
    except Exception as exc:
        raise HTTPException(
            400,
            "Invalid PowerPoint file"
        ) from exc

    if limit < 0:
        raise HTTPException(
            400,
            "Slide limit cannot be negative"
        )

    n = (
        min(len(prs.slides), limit)
        if limit > 0
        else len(prs.slides)
    )

    notes = [
        get_notes(prs.slides[i])
        for i in range(n)
    ]

    if mode == "notes":

        output = "\n\n".join(
            f"=== Slide {i+1} ===\n{txt}"
            for i, txt in enumerate(notes)
        )

        return download(
            output.encode("utf-8"),
            "slide_notes.txt",
            "text/plain; charset=utf-8"
        )

    if mode == "existing":

        if 0 < limit < len(prs.slides):
            raise HTTPException(
                400,
                "Existing Audio mode requires All slides (limit = 0)"
            )

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
                    400,
                    "No audio files found in this presentation"
                )

        return download(
            content,
            "narrated_existing_audio.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    if mode != "tts":
        raise HTTPException(
            400,
            "Invalid mode"
        )

    if not any(notes):
        raise HTTPException(
            400,
            "No Speaker Notes found in selected slides"
        )

    for i, note in enumerate(notes):

        if not note:
            continue

        try:
            audio = await synth(note, voice)

            add_audio(
                prs.slides[i],
                audio
            )

        except HTTPException as exc:
            raise HTTPException(
                exc.status_code,
                f"Slide {i+1}: {exc.detail}"
            ) from exc

    output = io.BytesIO()

    prs.save(output)

    return download(
        output.getvalue(),
        "narrated_free_edge_tts.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
