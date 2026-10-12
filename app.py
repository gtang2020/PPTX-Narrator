"""Free PowerPoint Narrator using Piper TTS.

No OpenAI API key required.
Fixes: 'Part' object has no attribute 'sha1'.
"""

import io
import hashlib
import zipfile
import tempfile
import subprocess
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse

from pptx import Presentation
from pptx.package import _MediaParts


# ============================================================
# FIX: Existing PowerPoint media may be loaded as generic Part
# objects without a sha1 property.
# ============================================================

def safe_find_by_sha1(self, sha1):
    """Find existing media by hashing its actual binary content."""

    for media_part in self:
        try:
            existing_sha1 = hashlib.sha1(
                media_part.blob
            ).hexdigest()

            if existing_sha1 == sha1:
                return media_part

        except (AttributeError, TypeError):
            continue

    return None


_MediaParts._find_by_sha1 = safe_find_by_sha1


# ============================================================
# Application settings
# ============================================================

app = FastAPI(title="Free PowerPoint Narrator")

ROOT = Path(__file__).parent

MAX_BYTES = 150 * 1024 * 1024

MODEL = ROOT / "models" / "en_US-lessac-low.onnx"

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.presentation"
)


# ============================================================
# Website
# ============================================================

@app.get("/", response_class=HTMLResponse)
def home():

    return (ROOT / "index.html").read_text(
        encoding="utf-8"
    )


# ============================================================
# Health check
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "ok",
        "engine": "Piper TTS",
        "model_exists": MODEL.exists(),
        "api_key_required": False,
    }


# ============================================================
# Extract PowerPoint speaker notes
# ============================================================

def get_notes(slide):

    if not slide.has_notes_slide:
        return ""

    tf = slide.notes_slide.notes_text_frame

    if tf is None:
        return ""

    return tf.text.strip()


# ============================================================
# Generate speech using free Piper TTS
# ============================================================

def synth(text):
    """Generate WAV speech using the local Piper model."""

    if not MODEL.exists():
        raise RuntimeError(
            "Piper voice model not found. "
            "Check Docker build."
        )

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False
    ) as f:

        output_path = Path(f.name)

    try:

        result = subprocess.run(
            [
                "piper",
                "--model",
                str(MODEL),
                "--output_file",
                str(output_path),
            ],
            input=text,
            text=True,
            capture_output=True,
            timeout=180,
        )

        if result.returncode != 0:

            raise RuntimeError(
                result.stderr[-500:]
                or "Piper speech generation failed"
            )

        audio = output_path.read_bytes()

        if len(audio) < 100:
            raise RuntimeError(
                "Generated audio is empty"
            )

        return audio

    finally:

        output_path.unlink(missing_ok=True)


# ============================================================
# Embed generated audio into PowerPoint
# ============================================================

def add_audio(slide, audio):
    """Embed WAV audio as a playable media shape."""

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False
    ) as f:

        f.write(audio)

        filename = Path(f.name)

    try:

        shape = slide.shapes.add_movie(
            str(filename),
            left=0,
            top=0,
            width=200000,
            height=200000,
            mime_type="audio/wav",
        )

        return shape

    finally:

        filename.unlink(missing_ok=True)


# ============================================================
# Download helper
# ============================================================

def download(data, filename, media_type):

    return StreamingResponse(
        io.BytesIO(data),
        media_type=media_type,
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        },
    )


# ============================================================
# Process PowerPoint
# ============================================================

@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("existing"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):

    # --------------------------------------------------------
    # Validate uploaded file
    # --------------------------------------------------------

    if (
        not file.filename
        or not file.filename.lower().endswith(".pptx")
    ):

        raise HTTPException(
            400,
            "Please upload a PPTX file"
        )

    content = await file.read(MAX_BYTES + 1)

    if len(content) > MAX_BYTES:

        raise HTTPException(
            413,
            "File exceeds 150 MB"
        )

    try:

        prs = Presentation(io.BytesIO(content))

    except Exception:

        raise HTTPException(
            400,
            "Invalid PowerPoint file"
        )

    if limit < 0:

        raise HTTPException(
            400,
            "Invalid slide limit"
        )

    total_slides = len(prs.slides)

    n = (
        min(limit, total_slides)
        if limit > 0
        else total_slides
    )

    notes = [
        get_notes(prs.slides[i])
        for i in range(n)
    ]

    # --------------------------------------------------------
    # MODE 1: Export speaker notes
    # --------------------------------------------------------

    if mode == "notes":

        output = "\n\n".join(
            f"=== Slide {i + 1} ===\n{note}"
            for i, note in enumerate(notes)
        )

        return download(
            output.encode("utf-8"),
            "slide_notes.txt",
            "text/plain; charset=utf-8",
        )

    # --------------------------------------------------------
    # MODE 2: Preserve existing audio
    # --------------------------------------------------------

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
                400,
                "No existing audio files found"
            )

        if 0 < limit < total_slides:

            raise HTTPException(
                400,
                "Existing Audio mode requires all slides"
            )

        return download(
            content,
            "existing_audio.pptx",
            PPTX_MIME,
        )

    # --------------------------------------------------------
    # MODE 3: Free Piper TTS
    # --------------------------------------------------------

    if mode != "tts":

        raise HTTPException(
            400,
            "Invalid processing mode"
        )

    if not any(notes):

        raise HTTPException(
            400,
            "No Speaker Notes found"
        )

    # --------------------------------------------------------
    # Generate and embed narration
    # --------------------------------------------------------

    generated_count = 0

    for i, note in enumerate(notes):

        if not note:
            continue

        try:

            audio = synth(note)

            add_audio(
                prs.slides[i],
                audio
            )

            generated_count += 1

        except Exception as exc:

            raise HTTPException(
                500,
                f"Slide {i + 1}: {str(exc)[:300]}"
            )

    if generated_count == 0:

        raise HTTPException(
            400,
            "No audio generated"
        )

    # --------------------------------------------------------
    # Save narrated PowerPoint
    # --------------------------------------------------------

    try:

        output = io.BytesIO()

        prs.save(output)

    except Exception as exc:

        raise HTTPException(
            500,
            f"PowerPoint saving failed: {str(exc)[:300]}"
        )

    # --------------------------------------------------------
    # Return completed PowerPoint
    # --------------------------------------------------------

    return download(
        output.getvalue(),
        "narrated_piper_free.pptx",
        PPTX_MIME,
    )
