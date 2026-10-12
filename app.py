"""Free PPTX Narrator: Piper TTS, replace old audio, autoplay."""

import io
import hashlib
import subprocess
import tempfile
import wave
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part

# Fix compatibility with some python-pptx versions.
if not hasattr(Part, "sha1"):
    Part.sha1 = property(
        lambda self: hashlib.sha1(self.blob).hexdigest()
    )

app = FastAPI(title="Free PowerPoint Narrator")

ROOT = Path(__file__).resolve().parent
MODEL = ROOT / "models" / "en_US-lessac-low.onnx"

MAX_BYTES = 150 * 1024 * 1024

P = "http://schemas.openxmlformats.org/presentationml/2006/main"

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.presentation"
)


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


def notes_of(slide):
    if not slide.has_notes_slide:
        return ""

    tf = slide.notes_slide.notes_text_frame

    return tf.text.strip() if tf is not None else ""


def tts_wav(text):
    """Generate free WAV narration using Piper."""

    if not MODEL.is_file():
        raise RuntimeError(
            "Piper model is missing in /models"
        )

    with tempfile.TemporaryDirectory() as td:

        path = Path(td) / "speech.wav"

        result = subprocess.run(
            [
                "piper",
                "--model",
                str(MODEL),
                "--output_file",
                str(path),
            ],
            input=text,
            text=True,
            capture_output=True,
            timeout=240,
        )

        if result.returncode != 0 or not path.exists():
            raise RuntimeError(
                (result.stderr or "Piper produced no WAV")[-500:]
            )

        with wave.open(str(path), "rb") as wf:
            duration_ms = int(
                wf.getnframes() * 1000 / wf.getframerate()
            )

        if duration_ms <= 0:
            raise RuntimeError("Empty audio")

        return path.read_bytes(), duration_ms


def remove_old_media(slide):
    """Remove old SayTXT audio and media animations."""

    for shape in list(slide.shapes):

        if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
            slide.shapes._spTree.remove(shape._element)

    for timing in slide._element.findall(
        f"{{{P}}}timing"
    ):
        slide._element.remove(timing)


def configure_autoplay(
    slide,
    shape_id,
    duration_ms,
    advance=True,
):
    """Set audio to start automatically."""

    timing = etree.SubElement(
        slide._element,
        f"{{{P}}}timing"
    )

    tn_list = etree.SubElement(
        timing,
        f"{{{P}}}tnLst"
    )

    par = etree.SubElement(
        tn_list,
        f"{{{P}}}par"
    )

    root_tn = etree.SubElement(
        par,
        f"{{{P}}}cTn",
        id="1",
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )

    children = etree.SubElement(
        root_tn,
        f"{{{P}}}childTnLst"
    )

    video = etree.SubElement(
        children,
        f"{{{P}}}video"
    )

    media = etree.SubElement(
        video,
        f"{{{P}}}cMediaNode",
        vol="80000",
    )

    media_tn = etree.SubElement(
        media,
        f"{{{P}}}cTn",
        id="2",
        fill="hold",
        display="0",
    )

    conds = etree.SubElement(
        media_tn,
        f"{{{P}}}stCondLst"
    )

    etree.SubElement(
        conds,
        f"{{{P}}}cond",
        delay="0",
    )

    target = etree.SubElement(
        media,
        f"{{{P}}}tgtEl"
    )

    etree.SubElement(
        target,
        f"{{{P}}}spTgt",
        spid=str(shape_id),
    )

    transition = slide._element.find(
        f"{{{P}}}transition"
    )

    if transition is None:

        transition = etree.Element(
            f"{{{P}}}transition"
        )

        slide._element.insert(
            slide._element.index(timing),
            transition,
        )

    transition.set(
        "advClick",
        "0" if advance else "1"
    )

    if advance:

        transition.set(
            "advTm",
            str(duration_ms + 1000)
        )

    else:

        transition.attrib.pop(
            "advTm",
            None
        )


def add_wav(slide, wav):
    """Embed newly generated narration."""

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False,
    ) as f:

        f.write(wav)
        path = Path(f.name)

    try:

        shape = slide.shapes.add_movie(
            str(path),
            -914400,
            -914400,
            320040,
            320040,
            mime_type="audio/wav",
        )

        return shape.shape_id

    finally:

        path.unlink(missing_ok=True)


def attachment(data, name, mime):
    return StreamingResponse(
        io.BytesIO(data),
        media_type=mime,
        headers={
            "Content-Disposition":
            f'attachment; filename="{name}"'
        },
    )


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("tts"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):

    # voice and api_key are retained for the old HTML form.
    # Piper does not require an API key.

    if (
        not file.filename
        or not file.filename.lower().endswith(".pptx")
    ):
        raise HTTPException(
            400,
            "Upload a .pptx file"
        )

    if limit < 0:
        raise HTTPException(
            400,
            "Slide count must be 0 or greater"
        )

    raw = await file.read(MAX_BYTES + 1)

    if len(raw) > MAX_BYTES:
        raise HTTPException(
            413,
            "PPTX exceeds 150 MB"
        )

    try:

        prs = Presentation(io.BytesIO(raw))

    except Exception:

        raise HTTPException(
            400,
            "Invalid PPTX"
        )

    count = (
        min(limit, len(prs.slides))
        if limit
        else len(prs.slides)
    )

    notes = [
        notes_of(prs.slides[i])
        for i in range(count)
    ]

    if mode == "notes":

        txt = "\n\n".join(
            f"=== Slide {i+1} ===\n{note}"
            for i, note in enumerate(notes)
        )

        return attachment(
            txt.encode("utf-8"),
            "slide_notes.txt",
            "text/plain; charset=utf-8",
        )

    if mode == "existing":

        return attachment(
            raw,
            "original_existing_audio.pptx",
            PPTX_MIME,
        )

    if mode != "tts":

        raise HTTPException(
            400,
            "Invalid mode"
        )

    if not any(notes):

        raise HTTPException(
            400,
            "No speaker notes found"
        )

    for i in range(count):

        slide = prs.slides[i]

        if not notes[i]:
            continue

        try:

            wav, duration_ms = tts_wav(
                notes[i]
            )

            remove_old_media(slide)

            media_id = add_wav(
                slide,
                wav
            )

            configure_autoplay(
                slide,
                media_id,
                duration_ms,
                advance=(
                    i < len(prs.slides) - 1
                ),
            )

        except Exception as exc:

            raise HTTPException(
                500,
                f"Slide {i+1}: {str(exc)[:300]}"
            )

    output = io.BytesIO()

    try:

        prs.save(output)

    except Exception as exc:

        raise HTTPException(
            500,
            f"Could not save PPTX: {str(exc)[:300]}"
        )

    return attachment(
        output.getvalue(),
        "narrated_piper_autoplay.pptx",
        PPTX_MIME,
    )
