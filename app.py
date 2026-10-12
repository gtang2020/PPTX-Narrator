"""
PPTX Narrator - Free Piper TTS
Version: Safe Audio Replacement

Features:
- Free Piper TTS
- No OpenAI API key
- Extract Speaker Notes
- Remove old slide audio
- Insert new WAV narration
- Preserve PowerPoint content
- Avoid custom autoplay XML
"""

import io
import hashlib
import subprocess
import tempfile
import wave
import zipfile

from pathlib import Path

from fastapi import (
    FastAPI,
    UploadFile,
    File,
    Form,
    HTTPException,
)

from fastapi.responses import (
    HTMLResponse,
    StreamingResponse,
)

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part


# =====================================================
# COMPATIBILITY FIX
# =====================================================

if not hasattr(Part, "sha1"):

    Part.sha1 = property(
        lambda self: hashlib.sha1(
            self.blob
        ).hexdigest()
    )


# =====================================================
# APPLICATION SETTINGS
# =====================================================

app = FastAPI(
    title="Free PowerPoint Narrator"
)

ROOT = Path(__file__).resolve().parent

MODEL = (
    ROOT
    / "models"
    / "en_US-lessac-low.onnx"
)

MAX_BYTES = 150 * 1024 * 1024

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.presentation"
)

P_NS = (
    "http://schemas.openxmlformats.org/"
    "presentationml/2006/main"
)

AUDIO_EXTENSIONS = (
    ".mp3",
    ".m4a",
    ".wav",
    ".wma",
    ".aac",
)


# =====================================================
# WEBSITE
# =====================================================

@app.get("/", response_class=HTMLResponse)
def home():

    return (
        ROOT / "index.html"
    ).read_text(
        encoding="utf-8"
    )


# =====================================================
# HEALTH CHECK
# =====================================================

@app.get("/health")
def health():

    return {
        "status": "ok",
        "engine": "Piper TTS",
        "model_exists": MODEL.exists(),
        "api_key_required": False,
        "version": "safe-audio-replacement",
    }


# =====================================================
# EXTRACT SPEAKER NOTES
# =====================================================

def get_notes(slide):

    if not slide.has_notes_slide:
        return ""

    frame = (
        slide.notes_slide.notes_text_frame
    )

    if frame is None:
        return ""

    return frame.text.strip()


# =====================================================
# GENERATE FREE PIPER SPEECH
# =====================================================

def generate_wav(text):

    if not MODEL.is_file():

        raise RuntimeError(
            "Piper voice model is missing"
        )

    with tempfile.TemporaryDirectory() as td:

        wav_path = (
            Path(td) / "narration.wav"
        )

        result = subprocess.run(
            [
                "piper",
                "--model",
                str(MODEL),
                "--output_file",
                str(wav_path),
            ],
            input=text,
            text=True,
            capture_output=True,
            timeout=240,
        )

        if result.returncode != 0:

            raise RuntimeError(
                result.stderr[-400:]
                or "Piper generation failed"
            )

        if not wav_path.exists():

            raise RuntimeError(
                "Piper did not produce audio"
            )

        with wave.open(
            str(wav_path),
            "rb"
        ) as wav_file:

            duration = (
                wav_file.getnframes()
                / wav_file.getframerate()
            )

        if duration <= 0:

            raise RuntimeError(
                "Generated audio is empty"
            )

        return wav_path.read_bytes()


# =====================================================
# REMOVE OLD AUDIO SHAPES
# =====================================================

def remove_old_audio(slide):

    removed_ids = set()

    for shape in list(slide.shapes):

        if (
            shape.shape_type
            == MSO_SHAPE_TYPE.MEDIA
        ):

            removed_ids.add(
                str(shape.shape_id)
            )

            slide.shapes._spTree.remove(
                shape._element
            )

    # Remove existing media timing information.
    #
    # IMPORTANT:
    # This also removes other animations on a slide
    # that contained media. It is a temporary safety
    # measure until autoplay is rebuilt correctly.

    if removed_ids:

        timing_tag = (
            "{" + P_NS + "}timing"
        )

        for timing in list(
            slide._element.findall(
                timing_tag
            )
        ):

            slide._element.remove(
                timing
            )

    return len(removed_ids)


# =====================================================
# INSERT NEW AUDIO
# =====================================================

def insert_audio(slide, audio):

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False
    ) as temp:

        temp.write(audio)

        filename = Path(temp.name)

    try:

        shape = slide.shapes.add_movie(
            str(filename),
            left=0,
            top=0,
            width=320040,
            height=320040,
            mime_type="audio/wav",
        )

        return shape.shape_id

    finally:

        filename.unlink(
            missing_ok=True
        )


# =====================================================
# DOWNLOAD HELPER
# =====================================================

def download(
    data,
    filename,
    media_type
):

    return StreamingResponse(
        io.BytesIO(data),
        media_type=media_type,
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        },
    )


# =====================================================
# PROCESS POWERPOINT
# =====================================================

@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("existing"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):

    # The existing HTML interface sends voice
    # and api_key fields.
    #
    # Piper ignores these values.

    if (
        not file.filename
        or not file.filename.lower().endswith(
            ".pptx"
        )
    ):

        raise HTTPException(
            400,
            "Please upload a PPTX file"
        )

    if limit < 0:

        raise HTTPException(
            400,
            "Invalid slide count"
        )

    raw = await file.read(
        MAX_BYTES + 1
    )

    if len(raw) > MAX_BYTES:

        raise HTTPException(
            413,
            "File exceeds 150 MB"
        )

    try:

        prs = Presentation(
            io.BytesIO(raw)
        )

    except Exception:

        raise HTTPException(
            400,
            "Invalid PowerPoint file"
        )

    total = len(prs.slides)

    count = (
        min(limit, total)
        if limit > 0
        else total
    )

    notes = [
        get_notes(prs.slides[i])
        for i in range(count)
    ]

    # =============================================
    # MODE 1: EXTRACT NOTES
    # =============================================

    if mode == "notes":

        text = "\n\n".join(
            f"=== Slide {i + 1} ===\n{note}"
            for i, note in enumerate(notes)
        )

        return download(
            text.encode("utf-8"),
            "slide_notes.txt",
            "text/plain; charset=utf-8",
        )

    # =============================================
    # MODE 2: KEEP EXISTING AUDIO
    # =============================================

    if mode == "existing":

        if 0 < limit < total:

            raise HTTPException(
                400,
                "Existing Audio mode requires "
                "all slides"
            )

        with zipfile.ZipFile(
            io.BytesIO(raw)
        ) as archive:

            has_audio = any(
                name.startswith(
                    "ppt/media/"
                )
                and name.lower().endswith(
                    AUDIO_EXTENSIONS
                )
                for name in archive.namelist()
            )

        if not has_audio:

            raise HTTPException(
                400,
                "No existing audio found"
            )

        return download(
            raw,
            "original_audio.pptx",
            PPTX_MIME,
        )

    # =============================================
    # MODE 3: FREE PIPER TTS
    # =============================================

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

    generated = 0

    for i in range(count):

        slide = prs.slides[i]

        note = notes[i]

        if not note:
            continue

        try:

            # Generate new narration first.

            audio = generate_wav(
                note
            )

            # Remove existing media shapes.

            remove_old_audio(
                slide
            )

            # Insert new narration.

            insert_audio(
                slide,
                audio
            )

            generated += 1

        except Exception as exc:

            raise HTTPException(
                500,
                f"Slide {i + 1}: "
                f"{str(exc)[:300]}"
            )

    if generated == 0:

        raise HTTPException(
            400,
            "No audio was generated"
        )

    # =============================================
    # SAVE PRESENTATION
    # =============================================

    try:

        output = io.BytesIO()

        prs.save(
            output
        )

        result = output.getvalue()

        # Basic ZIP integrity check.

        with zipfile.ZipFile(
            io.BytesIO(result)
        ) as archive:

            if archive.testzip() is not None:

                raise RuntimeError(
                    "PPTX ZIP integrity check failed"
                )

        # Verify python-pptx can reopen the file.

        check = Presentation(
            io.BytesIO(result)
        )

        if len(check.slides) != total:

            raise RuntimeError(
                "Slide count changed unexpectedly"
            )

    except Exception as exc:

        raise HTTPException(
            500,
            "PowerPoint validation failed: "
            + str(exc)[:300]
        )

    return download(
        result,
        "narrated_piper_safe.pptx",
        PPTX_MIME,
    )
