"""
PPTX Narrator V3
Free Piper TTS

Features:
1. Free offline Piper voice
2. No API key
3. Complete narration using sentence chunks
4. Replace old audio
5. Automatic audio playback
6. Automatic slide advancement
7. Preserve slide content and notes

Recommended initial test: 3 slides.
"""

import io
import re
import hashlib
import subprocess
import tempfile
import wave
import zipfile

from pathlib import Path
from xml.etree import ElementTree as ET

from fastapi import (
    FastAPI,
    File,
    Form,
    HTTPException,
    UploadFile,
)

from fastapi.responses import (
    HTMLResponse,
    StreamingResponse,
)

from lxml import etree

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part


# =====================================================
# COMPATIBILITY
# =====================================================

if not hasattr(Part, "sha1"):
    Part.sha1 = property(
        lambda self: hashlib.sha1(
            self.blob
        ).hexdigest()
    )


# =====================================================
# SETTINGS
# =====================================================

app = FastAPI(
    title="PPTX Narrator V3"
)

ROOT = Path(__file__).resolve().parent

MODEL = (
    ROOT
    / "models"
    / "en_US-lessac-low.onnx"
)

MAX_BYTES = 150 * 1024 * 1024

PAUSE_AFTER_AUDIO_MS = 2000

P = (
    "http://schemas.openxmlformats.org/"
    "presentationml/2006/main"
)

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.presentation"
)


def ptag(name):
    return f"{{{P}}}{name}"


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


@app.get("/health")
def health():

    return {
        "status": "ok",
        "version": "3.0",
        "engine": "Piper TTS",
        "model_exists": MODEL.exists(),
        "api_key_required": False,
        "autoplay": True,
        "sentence_chunking": True,
    }


# =====================================================
# SPEAKER NOTES
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
# SPLIT NOTES INTO SHORT CHUNKS
# =====================================================

def split_narration(text, max_chars=240):
    """
    Split narration into manageable segments.

    Keep all non-whitespace text.
    Avoid sending an entire long slide to Piper
    in one operation.
    """

    text = text.strip()

    if not text:
        return []

    # Prefer sentence boundaries.
    sentences = re.split(
        r"(?<=[.!?。！？])\s+",
        text
    )

    chunks = []
    current = ""

    for sentence in sentences:

        sentence = sentence.strip()

        if not sentence:
            continue

        if len(current) + len(sentence) + 1 <= max_chars:

            current = (
                current + " " + sentence
            ).strip()

            continue

        if current:
            chunks.append(current)
            current = ""

        # Handle unusually long sentences.
        while len(sentence) > max_chars:

            cut = sentence.rfind(
                " ",
                0,
                max_chars + 1
            )

            if cut <= 0:
                cut = max_chars

            part = sentence[:cut].strip()

            if part:
                chunks.append(part)

            sentence = sentence[cut:].strip()

        current = sentence

    if current:
        chunks.append(current)

    return chunks


# =====================================================
# GENERATE ONE PIPER WAV
# =====================================================

def piper_chunk(text, output_path):

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
        timeout=240,
    )

    if result.returncode != 0:

        raise RuntimeError(
            result.stderr[-500:]
            or "Piper generation failed"
        )

    if not output_path.is_file():

        raise RuntimeError(
            "Piper produced no WAV file"
        )

    if output_path.stat().st_size < 100:

        raise RuntimeError(
            "Piper produced empty audio"
        )


# =====================================================
# GENERATE COMPLETE NARRATION
# =====================================================

def generate_complete_wav(text):
    """
    Generate audio for every text chunk.

    Merge the WAV segments in their original order.

    Raise an error if any segment fails.
    """

    if not MODEL.is_file():

        raise RuntimeError(
            "Piper model is missing"
        )

    chunks = split_narration(text)

    if not chunks:

        raise RuntimeError(
            "Speaker Notes are empty"
        )

    with tempfile.TemporaryDirectory() as td:

        directory = Path(td)

        wav_paths = []

        for index, chunk in enumerate(chunks):

            path = (
                directory
                / f"chunk_{index:04d}.wav"
            )

            piper_chunk(
                chunk,
                path
            )

            wav_paths.append(path)

        # Merge uncompressed WAV audio.
        # All segments must share the same
        # sample format.

        output_path = (
            directory / "complete.wav"
        )

        reference_params = None

        total_frames = 0

        with wave.open(
            str(output_path),
            "wb"
        ) as destination:

            for path in wav_paths:

                with wave.open(
                    str(path),
                    "rb"
                ) as source:

                    params = (
                        source.getnchannels(),
                        source.getsampwidth(),
                        source.getframerate(),
                        source.getcomptype(),
                    )

                    if reference_params is None:

                        reference_params = params

                        destination.setnchannels(
                            params[0]
                        )

                        destination.setsampwidth(
                            params[1]
                        )

                        destination.setframerate(
                            params[2]
                        )

                    elif params != reference_params:

                        raise RuntimeError(
                            "WAV segment formats differ"
                        )

                    frames = source.readframes(
                        source.getnframes()
                    )

                    if not frames:

                        raise RuntimeError(
                            "A narration segment is empty"
                        )

                    destination.writeframes(
                        frames
                    )

                    total_frames += (
                        source.getnframes()
                    )

        if total_frames <= 0:

            raise RuntimeError(
                "Merged narration is empty"
            )

        with wave.open(
            str(output_path),
            "rb"
        ) as audio:

            duration_ms = round(
                1000
                * audio.getnframes()
                / audio.getframerate()
            )

        return (
            output_path.read_bytes(),
            duration_ms,
            len(chunks),
        )


# =====================================================
# REMOVE OLD AUDIO
# =====================================================

def remove_old_audio(slide):
    """
    Remove old media shapes.

    Remove old timing nodes to prevent
    conflicting playback instructions.

    Note:
    Existing slide animations are also
    removed on processed slides.
    """

    for shape in list(slide.shapes):

        if (
            shape.shape_type
            == MSO_SHAPE_TYPE.MEDIA
        ):

            slide.shapes._spTree.remove(
                shape._element
            )

    # Remove all existing timing nodes.
    # We will create exactly one new timing
    # node for the narration.

    for timing in list(
        slide._element.findall(
            ptag("timing")
        )
    ):

        slide._element.remove(
            timing
        )


# =====================================================
# INSERT WAV
# =====================================================

def insert_wav(slide, audio_bytes):

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False
    ) as temporary:

        temporary.write(
            audio_bytes
        )

        path = Path(
            temporary.name
        )

    try:

        shape = slide.shapes.add_movie(
            str(path),
            left=0,
            top=0,
            width=320040,
            height=320040,
            mime_type="audio/wav",
        )

        return shape.shape_id

    finally:

        path.unlink(
            missing_ok=True
        )


# =====================================================
# AUTOMATIC AUDIO PLAYBACK
# =====================================================

def configure_autoplay(
    slide,
    shape_id,
    duration_ms,
    advance=True,
):
    """
    Create one PowerPoint audio timing tree.

    Audio begins when the slide appears.

    The slide advances after narration
    plus an additional pause.

    Uses p:audio rather than p:video.
    """

    # Remove any timing that may have been
    # created during media insertion.

    for old in list(
        slide._element.findall(
            ptag("timing")
        )
    ):

        slide._element.remove(old)

    timing = etree.Element(
        ptag("timing")
    )

    tn_list = etree.SubElement(
        timing,
        ptag("tnLst")
    )

    par = etree.SubElement(
        tn_list,
        ptag("par")
    )

    root = etree.SubElement(
        par,
        ptag("cTn"),
        id="1",
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )

    child_list = etree.SubElement(
        root,
        ptag("childTnLst")
    )

    audio = etree.SubElement(
        child_list,
        ptag("audio")
    )

    media = etree.SubElement(
        audio,
        ptag("cMediaNode"),
        vol="80000",
    )

    media_timing = etree.SubElement(
        media,
        ptag("cTn"),
        id="2",
        fill="hold",
        display="0",
    )

    conditions = etree.SubElement(
        media_timing,
        ptag("stCondLst")
    )

    # Crucial difference:
    # delay=0 instead of delay=indefinite.

    etree.SubElement(
        conditions,
        ptag("cond"),
        delay="0",
    )

    target = etree.SubElement(
        media,
        ptag("tgtEl")
    )

    etree.SubElement(
        target,
        ptag("spTgt"),
        spid=str(shape_id),
    )

    # PowerPoint slide child order:
    #
    # cSld
    # clrMapOvr
    # transition
    # timing
    # extLst

    extension = slide._element.find(
        ptag("extLst")
    )

    if extension is not None:

        slide._element.insert(
            slide._element.index(extension),
            timing
        )

    else:

        slide._element.append(
            timing
        )

    # Configure slide advancement.

    transition = slide._element.find(
        ptag("transition")
    )

    if transition is None:

        transition = etree.Element(
            ptag("transition")
        )

        slide._element.insert(
            slide._element.index(timing),
            transition
        )

    if advance:

        transition.set(
            "advClick",
            "0"
        )

        transition.set(
            "advTm",
            str(
                duration_ms
                + PAUSE_AFTER_AUDIO_MS
            )
        )

    else:

        # Last slide stays visible.

        transition.set(
            "advClick",
            "1"
        )

        transition.attrib.pop(
            "advTm",
            None
        )


# =====================================================
# DOWNLOAD
# =====================================================

def download(
    data,
    filename,
    mime
):

    return StreamingResponse(
        io.BytesIO(data),
        media_type=mime,
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        },
    )


# =====================================================
# PROCESS
# =====================================================

@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("tts"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):

    # voice and api_key are retained for
    # compatibility with the existing website.
    # Piper does not require an API key.

    if (
        not file.filename
        or not file.filename.lower().endswith(
            ".pptx"
        )
    ):

        raise HTTPException(
            400,
            "Upload a PPTX file"
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
            "Invalid PPTX"
        )

    total = len(
        prs.slides
    )

    count = (
        min(limit, total)
        if limit
        else total
    )

    notes = [
        get_notes(
            prs.slides[i]
        )
        for i in range(count)
    ]

    # -----------------------------------------
    # EXTRACT NOTES
    # -----------------------------------------

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

    # -----------------------------------------
    # KEEP EXISTING AUDIO
    # -----------------------------------------

    if mode == "existing":

        return download(
            raw,
            "original_audio.pptx",
            PPTX_MIME,
        )

    # -----------------------------------------
    # GENERATE NEW NARRATION
    # -----------------------------------------

    if mode != "tts":

        raise HTTPException(
            400,
            "Invalid mode"
        )

    if not any(notes):

        raise HTTPException(
            400,
            "No Speaker Notes found"
        )

    for i in range(count):

        slide = prs.slides[i]

        note = notes[i]

        if not note:
            continue

        try:

            # Generate complete narration.

            (
                audio_bytes,
                duration_ms,
                chunk_count,
            ) = generate_complete_wav(
                note
            )

            # Remove old media and timing.

            remove_old_audio(
                slide
            )

            # Insert new audio.

            shape_id = insert_wav(
                slide,
                audio_bytes
            )

            # Configure autoplay.

            configure_autoplay(
                slide,
                shape_id,
                duration_ms,
                advance=(
                    i < total - 1
                ),
            )

            print(
                f"Slide {i + 1}: "
                f"{chunk_count} chunks, "
                f"{duration_ms / 1000:.2f} seconds",
                flush=True,
            )

        except Exception as exc:

            raise HTTPException(
                500,
                f"Slide {i + 1}: "
                f"{str(exc)[:300]}"
            )

    # -----------------------------------------
    # SAVE AND VALIDATE
    # -----------------------------------------

    try:

        output = io.BytesIO()

        prs.save(
            output
        )

        result = output.getvalue()

        with zipfile.ZipFile(
            io.BytesIO(result)
        ) as archive:

            if archive.testzip() is not None:

                raise RuntimeError(
                    "ZIP integrity check failed"
                )

            # Check every processed slide XML
            # for duplicate timing nodes.

            for i in range(count):

                slide_xml = archive.read(
                    f"ppt/slides/slide{i + 1}.xml"
                )

                root = ET.fromstring(
                    slide_xml
                )

                timing_nodes = root.findall(
                    ptag("timing")
                )

                if len(timing_nodes) > 1:

                    raise RuntimeError(
                        f"Slide {i + 1} has "
                        "duplicate timing nodes"
                    )

        # Verify PowerPoint can be parsed again.

        check = Presentation(
            io.BytesIO(result)
        )

        if len(check.slides) != total:

            raise RuntimeError(
                "Slide count changed"
            )

    except Exception as exc:

        raise HTTPException(
            500,
            "PPTX validation failed: "
            + str(exc)[:300]
        )

    return download(
        result,
        "narrated_piper_v3_autoplay.pptx",
        PPTX_MIME,
    )
