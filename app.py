"""
PPTX Narrator V4.6

Step 1:
    Insert slide-numbered DOCX/PDF/TXT narration into
    PowerPoint Speaker Notes.

Step 2:
    Generate Piper WAV narration.
    Preserve individual narration segments.
    Embed merged WAV audio into each slide.
    Configure automatic playback.
    Save PowerPoint checkpoints.
    Support resume and audio ZIP downloads.

V4.5 improvements:
    - Configurable Piper timeout
    - Automatic Piper retries
    - Smaller narration chunks
    - Reuse completed WAV segments
    - Detailed segment progress
    - Safer checkpoint recovery

Requirements:
    fastapi
    python-multipart
    python-pptx
    lxml
    PyMuPDF (optional)

External:
    Piper executable
    Piper ONNX voice model

Important:
    Render free storage is temporary.
    Download checkpoints frequently.
"""

import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
import wave
import zipfile

from concurrent.futures import ThreadPoolExecutor
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
    FileResponse,
    HTMLResponse,
    Response,
    StreamingResponse,
)

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

if not hasattr(Part, "sha1"):
    Part.sha1 = property(
        lambda self: hashlib.sha1(
            self.blob
        ).hexdigest()
    )


app = FastAPI(
    title="PPTX Narrator V4.6"
)

ROOT = Path(__file__).resolve().parent

MODEL = Path(
    os.environ.get(
        "PIPER_MODEL",
        str(
            ROOT
            / "models"
            / "en_US-lessac-low.onnx"
        ),
    )
)

WORKDIR = Path(
    os.environ.get(
        "NARRATOR_WORKDIR",
        str(ROOT / "narrator_jobs"),
    )
)

WORKDIR.mkdir(
    parents=True,
    exist_ok=True,
)

MAX_BYTES = 150 * 1024 * 1024
MAX_NOTES_BYTES = 25 * 1024 * 1024

PAUSE_AFTER_AUDIO_MS = 2000

P = (
    "http://schemas.openxmlformats.org/"
    "presentationml/2006/main"
)

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "presentationml.presentation"
)

MAX_ACTIVE_JOBS = 1

executor = ThreadPoolExecutor(
    max_workers=1
)

active_jobs = set()
active_lock = threading.Lock()


# ============================================================
# PIPER CONFIGURATION
# ============================================================

def env_int(name, default, minimum, maximum):
    try:
        value = int(
            os.environ.get(
                name,
                str(default),
            )
        )
    except (ValueError, TypeError):
        value = default

    return max(
        minimum,
        min(maximum, value),
    )


PIPER_TIMEOUT = env_int(
    "PIPER_TIMEOUT",
    360,
    30,
    1800,
)

PIPER_RETRIES = env_int(
    "PIPER_RETRIES",
    2,
    1,
    5,
)

PIPER_MAX_CHARS = env_int(
    "PIPER_MAX_CHARS",
    140,
    40,
    500,
)


# ============================================================
# BASIC HELPERS
# ============================================================

def ptag(name):
    return f"{{{P}}}{name}"


def save_json(path, data):
    temp = path.with_name(
        path.name + ".tmp"
    )

    temp.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    os.replace(
        temp,
        path,
    )


def read_status(job_id):
    if not re.fullmatch(
        r"[a-f0-9]{32}",
        job_id,
    ):
        raise HTTPException(
            400,
            "Invalid job ID",
        )

    path = (
        WORKDIR
        / job_id
        / "status.json"
    )

    if not path.is_file():
        raise HTTPException(
            404,
            "Job not found; temporary server "
            "storage may have been cleared",
        )

    try:
        return json.loads(
            path.read_text(
                encoding="utf-8"
            )
        )

    except (OSError, ValueError) as exc:
        raise HTTPException(
            500,
            "Job status is unreadable",
        ) from exc


def update_status(job_dir, status):
    status["updated_at"] = int(
        time.time()
    )

    save_json(
        job_dir / "status.json",
        status,
    )


# ============================================================
# WEBSITE
# ============================================================

@app.get(
    "/",
    response_class=HTMLResponse,
)
def home():
    page = (
        ROOT
        / "index_v4.html"
    )

    if not page.is_file():
        raise HTTPException(
            500,
            "index_v4.html not found",
        )

    return page.read_text(
        encoding="utf-8"
    )


@app.get("/health")
def health():
    try:
        import fitz
        pdf_available = True
    except ImportError:
        pdf_available = False

    with active_lock:
        active = len(
            active_jobs
        )

    return {
        "status": "ok",
        "version": "4.6",
        "engine": "Piper",
        "model_exists": MODEL.is_file(),
        "piper_exists": (
            shutil.which("piper")
            is not None
        ),
        "external_notes": True,
        "docx_supported": True,
        "txt_supported": True,
        "pdf_supported": pdf_available,
        "api_key_required": False,
        "autoplay": True,
        "batch_processing": True,
        "segment_audio_zip": True,
        "storage_persistent": False,
        "active_jobs": active,
        "piper_timeout": PIPER_TIMEOUT,
        "piper_retries": PIPER_RETRIES,
        "piper_max_chars": PIPER_MAX_CHARS,
        "segment_resume": True,
        "portable_checkpoint": True,
        "resume_from_uploaded_pptx": True,
    }


# ============================================================
# SPEAKER NOTES
# ============================================================

def get_notes(slide):
    if not slide.has_notes_slide:
        return ""

    frame = (
        slide.notes_slide
        .notes_text_frame
    )

    if frame is None:
        return ""

    return frame.text.strip()


def set_notes(slide, text):
    frame = (
        slide.notes_slide
        .notes_text_frame
    )

    if frame is None:
        raise RuntimeError(
            "Cannot create Speaker Notes"
        )

    frame.text = text


# ============================================================
# TEXT FILE READERS
# ============================================================

def read_txt(data):
    for encoding in (
        "utf-8-sig",
        "utf-16",
        "cp1252",
    ):
        try:
            return data.decode(
                encoding
            )
        except UnicodeError:
            pass

    raise ValueError(
        "Cannot decode TXT file"
    )


def read_docx(data):
    W = (
        "http://schemas.openxmlformats.org/"
        "wordprocessingml/2006/main"
    )

    def wt(name):
        return f"{{{W}}}{name}"

    with zipfile.ZipFile(
        io.BytesIO(data)
    ) as archive:

        root = ET.fromstring(
            archive.read(
                "word/document.xml"
            )
        )

    body = root.find(
        wt("body")
    )

    if body is None:
        return ""

    lines = []

    for paragraph in body.iter(
        wt("p")
    ):
        parts = []

        for node in paragraph.iter():
            if node.tag == wt("t"):
                parts.append(
                    node.text or ""
                )

            elif node.tag == wt("tab"):
                parts.append("\t")

            elif node.tag in (
                wt("br"),
                wt("cr"),
            ):
                parts.append("\n")

        line = "".join(
            parts
        ).strip()

        if line:
            lines.append(
                line
            )

    return "\n".join(
        lines
    )


def read_pdf(data):
    try:
        import fitz
    except ImportError as exc:
        raise ValueError(
            "PDF support requires PyMuPDF"
        ) from exc

    document = fitz.open(
        stream=data,
        filetype="pdf",
    )

    try:
        return "\n".join(
            page.get_text(
                "text",
                sort=True,
            )
            for page in document
        )

    finally:
        document.close()


# ============================================================
# SCRIPT HEADING PARSERS
# ============================================================

SLIDE_HEADING = re.compile(
    r"(?im)^[ \t]*"
    r"(?:slide|幻灯片)"
    r"[ \t]*"
    r"(?:#|no\.?|number)?"
    r"[ \t]*0*(\d{1,4})"
    r"[ \t]*"
    r"(?:[:：.\-–—][^\n]*)?"
    r"[ \t]*$"
)

SEGMENT_HEADING = re.compile(
    r"(?im)^[ \t]*"
    r"(?:(?:narration|audio|speech)"
    r"[ \t]*)?"
    r"(?:segment|part|段落|片段)"
    r"[ \t]*"
    r"(?:#|no\.?)?"
    r"[ \t]*\d{1,4}"
    r"(?:[ \t]*[:：.\-–—][^\n]*)?"
    r"[ \t]*$"
)


def parse_slide_notes(text):
    text = text.replace(
        "\r\n",
        "\n",
    ).replace(
        "\r",
        "\n",
    )

    matches = list(
        SLIDE_HEADING.finditer(
            text
        )
    )

    if not matches:
        raise ValueError(
            "No slide headings found; "
            "use Slide 001, Slide 002, etc."
        )

    mapping = {}

    for index, match in enumerate(
        matches
    ):
        number = int(
            match.group(1)
        )

        if index + 1 < len(matches):
            end = (
                matches[index + 1]
                .start()
            )
        else:
            end = len(
                text
            )

        narration = text[
            match.end():end
        ].strip()

        if number in mapping:
            raise ValueError(
                f"Duplicate Slide {number}"
            )

        if not narration:
            raise ValueError(
                f"Slide {number} "
                "has empty narration"
            )

        mapping[number] = narration

    return mapping


def read_script(data, filename):
    ext = Path(
        filename
    ).suffix.lower()

    readers = {
        ".txt": read_txt,
        ".docx": read_docx,
        ".pdf": read_pdf,
    }

    if ext not in readers:
        raise ValueError(
            "Notes must be DOCX, PDF, or TXT"
        )

    return readers[ext](
        data
    )


def import_external_notes(
    prs,
    data,
    filename,
    existing_notes="replace",
    require_all=False,
):
    mapping = parse_slide_notes(
        read_script(
            data,
            filename,
        )
    )

    total = len(
        prs.slides
    )

    outside = sorted(
        n for n in mapping
        if not 1 <= n <= total
    )

    if outside:
        raise ValueError(
            f"Slide numbers outside PPTX "
            f"(1-{total}): {outside[:12]}"
        )

    missing = sorted(
        set(
            range(
                1,
                total + 1,
            )
        ) - set(mapping)
    )

    if require_all and missing:
        raise ValueError(
            f"Missing narration for "
            f"{len(missing)} slides: "
            f"{missing[:20]}"
        )

    for number, narration in mapping.items():
        slide = prs.slides[
            number - 1
        ]

        old = get_notes(
            slide
        )

        if (
            existing_notes == "keep"
            and old
        ):
            continue

        if (
            existing_notes == "append"
            and old
        ):
            narration = (
                old
                + "\n\n"
                + narration
            )

        set_notes(
            slide,
            narration,
        )

    return len(
        mapping
    )


# ============================================================
# UPLOAD HELPER
# ============================================================

async def save_upload(
    upload,
    destination,
    limit,
):
    size = 0

    with destination.open(
        "wb"
    ) as output:

        while True:
            chunk = await upload.read(
                1024 * 1024
            )

            if not chunk:
                break

            size += len(
                chunk
            )

            if size > limit:
                raise HTTPException(
                    413,
                    "Uploaded file is too large",
                )

            output.write(
                chunk
            )


# ============================================================
# STEP 1: INSERT NOTES
# ============================================================

@app.post("/notes/insert")
async def insert_notes_only(
    file: UploadFile = File(...),
    notes_file: UploadFile = File(...),
    existing_notes: str = Form("replace"),
    require_all: bool = Form(True),
):
    if existing_notes not in (
        "replace",
        "append",
        "keep",
    ):
        raise HTTPException(
            400,
            "Invalid existing_notes mode",
        )

    if (
        not file.filename
        or Path(
            file.filename
        ).suffix.lower() != ".pptx"
    ):
        raise HTTPException(
            400,
            "Please upload a PPTX",
        )

    if (
        not notes_file.filename
        or Path(
            notes_file.filename
        ).suffix.lower()
        not in (
            ".docx",
            ".pdf",
            ".txt",
        )
    ):
        raise HTTPException(
            400,
            "Please upload a DOCX, "
            "PDF, or TXT script",
        )

    with tempfile.TemporaryDirectory(
        prefix="pptx_notes_"
    ) as td:

        source = (
            Path(td)
            / "source.pptx"
        )

        await save_upload(
            file,
            source,
            MAX_BYTES,
        )

        data = await notes_file.read(
            MAX_NOTES_BYTES + 1
        )

        if len(data) > MAX_NOTES_BYTES:
            raise HTTPException(
                413,
                "Notes exceed 25 MB",
            )

        try:
            prs = Presentation(
                str(source)
            )

            count = import_external_notes(
                prs,
                data,
                notes_file.filename,
                existing_notes,
                require_all,
            )

            total = len(
                prs.slides
            )

            output = io.BytesIO()

            prs.save(
                output
            )

            output.seek(0)

            check = Presentation(
                output
            )

            if len(check.slides) != total:
                raise RuntimeError(
                    "Slide count changed"
                )

            output.seek(0)

        except Exception as exc:
            raise HTTPException(
                400,
                "Notes insertion failed: "
                + str(exc)[:450],
            ) from exc

    return Response(
        content=output.getvalue(),
        media_type=PPTX_MIME,
        headers={
            "Content-Disposition":
                'attachment; filename='
                '"PPTX_with_Speaker_Notes_V4_6.pptx"',
            "X-Notes-Imported": str(count),
            "X-Total-Slides": str(total),
        },
    )


# ============================================================
# NARRATION SEGMENTS
# ============================================================

def narration_segments(notes):
    """
    Preserve explicit segment headings.

    Otherwise, use blank-line paragraphs.
    """

    text = notes.replace(
        "\r\n",
        "\n",
    ).replace(
        "\r",
        "\n",
    ).strip()

    if not text:
        return []

    text = SLIDE_HEADING.sub(
        "",
        text,
    ).strip()

    matches = list(
        SEGMENT_HEADING.finditer(
            text
        )
    )

    if matches:
        segments = []

        preface = text[
            :matches[0].start()
        ].strip()

        if preface:
            segments.append(
                preface
            )

        for index, match in enumerate(
            matches
        ):
            if index + 1 < len(matches):
                end = (
                    matches[index + 1]
                    .start()
                )
            else:
                end = len(
                    text
                )

            segment = text[
                match.end():end
            ].strip()

            if segment:
                segments.append(
                    segment
                )

        return segments

    return [
        paragraph.strip()
        for paragraph in re.split(
            r"\n\s*\n+",
            text,
        )
        if paragraph.strip()
    ]


# ============================================================
# TEXT CHUNKING
# ============================================================

def split_narration(
    text,
    max_chars=None,
):
    if max_chars is None:
        max_chars = PIPER_MAX_CHARS

    sentences = re.split(
        r"(?<=[.!?。！？])\s+",
        text.strip(),
    )

    chunks = []
    current = ""

    for sentence in sentences:
        sentence = sentence.strip()

        if not sentence:
            continue

        proposed_length = (
            len(current)
            + len(sentence)
            + (1 if current else 0)
        )

        if proposed_length <= max_chars:
            current = (
                current
                + " "
                + sentence
            ).strip()
            continue

        if current:
            chunks.append(
                current
            )
            current = ""

        while len(sentence) > max_chars:
            cut = sentence.rfind(
                " ",
                0,
                max_chars + 1,
            )

            if cut <= 0:
                cut = max_chars

            chunks.append(
                sentence[:cut].strip()
            )

            sentence = sentence[
                cut:
            ].strip()

        current = sentence

    if current:
        chunks.append(
            current
        )

    return chunks


# ============================================================
# PIPER AUDIO GENERATION
# ============================================================

def piper_chunk(
    text,
    output_path,
):
    """
    Generate one WAV chunk.

    Retry failed synthesis automatically.
    """

    output_path = Path(
        output_path
    )

    last_error = None

    for attempt in range(
        1,
        PIPER_RETRIES + 1,
    ):
        output_path.unlink(
            missing_ok=True
        )

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
                timeout=PIPER_TIMEOUT,
            )

            if result.returncode != 0:
                raise RuntimeError(
                    result.stderr[-500:]
                    or "Piper generation failed"
                )

            if (
                not output_path.is_file()
                or output_path.stat().st_size < 100
            ):
                raise RuntimeError(
                    "Piper produced no audio"
                )

            # Validate WAV header.
            with wave.open(
                str(output_path),
                "rb",
            ) as audio:
                if audio.getnframes() <= 0:
                    raise RuntimeError(
                        "Piper generated empty WAV"
                    )

            return

        except FileNotFoundError as exc:
            raise RuntimeError(
                "Piper executable is not installed"
            ) from exc

        except subprocess.TimeoutExpired:
            last_error = (
                "Piper synthesis timed out "
                f"after {PIPER_TIMEOUT} seconds"
            )

        except Exception as exc:
            last_error = str(
                exc
            )

        output_path.unlink(
            missing_ok=True
        )

        if attempt < PIPER_RETRIES:
            time.sleep(
                3 * attempt
            )

    raise RuntimeError(
        f"{last_error}; "
        f"failed after "
        f"{PIPER_RETRIES} attempts"
    )


# ============================================================
# WAV VALIDATION
# ============================================================

def valid_wav(path):
    path = Path(
        path
    )

    if (
        not path.is_file()
        or path.stat().st_size < 100
    ):
        return False

    try:
        with wave.open(
            str(path),
            "rb",
        ) as audio:
            return (
                audio.getnframes() > 0
                and audio.getframerate() > 0
            )

    except Exception:
        return False


# ============================================================
# CONCATENATE WAV FILES
# ============================================================

def concatenate_wavs(
    paths,
    destination,
):
    if not paths:
        raise ValueError(
            "No WAV files"
        )

    reference = None
    frames_total = 0

    with wave.open(
        str(destination),
        "wb",
    ) as target:

        for path in paths:
            with wave.open(
                str(path),
                "rb",
            ) as source:

                params = (
                    source.getnchannels(),
                    source.getsampwidth(),
                    source.getframerate(),
                    source.getcomptype(),
                )

                if reference is None:
                    reference = params

                    target.setnchannels(
                        params[0]
                    )

                    target.setsampwidth(
                        params[1]
                    )

                    target.setframerate(
                        params[2]
                    )

                elif params != reference:
                    raise RuntimeError(
                        "Piper WAV formats differ"
                    )

                frames = source.readframes(
                    source.getnframes()
                )

                if not frames:
                    raise RuntimeError(
                        "Empty WAV segment"
                    )

                target.writeframes(
                    frames
                )

                frames_total += (
                    len(frames)
                    // (
                        params[0]
                        * params[1]
                    )
                )

    return round(
        1000
        * frames_total
        / reference[2]
    )


# ============================================================
# AUDIO SEGMENT PROGRESS
# ============================================================

def set_audio_progress(
    job_dir,
    status,
    slide_number,
    segment_number,
    segment_total,
    chunk_number,
    chunk_total,
):
    status["audio_slide"] = slide_number

    status["audio_segment"] = (
        segment_number
    )

    status["audio_segment_total"] = (
        segment_total
    )

    status["audio_chunk"] = (
        chunk_number
    )

    status["audio_chunk_total"] = (
        chunk_total
    )

    update_status(
        job_dir,
        status,
    )


# ============================================================
# SYNTHESIZE ONE SLIDE
# ============================================================

def synthesize_slide(
    notes,
    job_dir,
    slide_number,
    status=None,
):
    """
    Generate separate WAV segments and
    one merged WAV for the slide.

    Completed WAV segments are reused
    when the same job resumes.
    """

    if not MODEL.is_file():
        raise RuntimeError(
            "Piper model is missing"
        )

    segments = narration_segments(
        notes
    )

    if not segments:
        raise RuntimeError(
            "Empty narration"
        )

    slide_dir = (
        job_dir
        / "audio"
        / f"slide_{slide_number:04d}"
    )

    slide_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    segment_paths = []

    segment_total = len(
        segments
    )

    for segment_index, segment in enumerate(
        segments,
        1,
    ):
        output = (
            slide_dir
            / f"segment_{segment_index:03d}.wav"
        )

        # Reuse completed WAV segments.
        if valid_wav(output):
            segment_paths.append(
                output
            )

            if status is not None:
                set_audio_progress(
                    job_dir,
                    status,
                    slide_number,
                    segment_index,
                    segment_total,
                    1,
                    1,
                )

            continue

        chunks = split_narration(
            segment
        )

        if not chunks:
            raise RuntimeError(
                "Empty narration segment"
            )

        # Store temporary chunks in a
        # temporary directory.
        with tempfile.TemporaryDirectory(
            prefix="piper_chunk_"
        ) as td:

            chunk_paths = []

            for chunk_index, chunk in enumerate(
                chunks,
                1,
            ):
                if status is not None:
                    set_audio_progress(
                        job_dir,
                        status,
                        slide_number,
                        segment_index,
                        segment_total,
                        chunk_index,
                        len(chunks),
                    )

                chunk_path = (
                    Path(td)
                    / f"chunk_{chunk_index:04d}.wav"
                )

                piper_chunk(
                    chunk,
                    chunk_path,
                )

                chunk_paths.append(
                    chunk_path
                )

            temporary_output = (
                slide_dir
                / f"segment_{segment_index:03d}.tmp.wav"
            )

            concatenate_wavs(
                chunk_paths,
                temporary_output,
            )

            if not valid_wav(
                temporary_output
            ):
                raise RuntimeError(
                    "Generated segment WAV is invalid"
                )

            os.replace(
                temporary_output,
                output,
            )

        segment_paths.append(
            output
        )

    merged = (
        slide_dir
        / "slide_complete.wav"
    )

    temporary_merged = (
        slide_dir
        / "slide_complete.tmp.wav"
    )

    duration_ms = concatenate_wavs(
        segment_paths,
        temporary_merged,
    )

    if not valid_wav(
        temporary_merged
    ):
        raise RuntimeError(
            "Merged WAV is invalid"
        )

    os.replace(
        temporary_merged,
        merged,
    )

    return (
        merged,
        duration_ms,
        len(segments),
    )


# ============================================================
# LEGACY AUDIO HELPER
# ============================================================

def generate_complete_wav(text):
    with tempfile.TemporaryDirectory(
        prefix="legacy_piper_"
    ) as td:

        path, duration, count = synthesize_slide(
            text,
            Path(td),
            1,
        )

        return (
            path.read_bytes(),
            duration,
            count,
        )


# ============================================================
# POWERPOINT AUDIO EMBEDDING
# ============================================================

def remove_old_audio(slide):
    """
    Remove existing media and timing.

    Existing slide animations may be
    affected by this operation.
    """

    for shape in list(
        slide.shapes
    ):
        if (
            shape.shape_type
            == MSO_SHAPE_TYPE.MEDIA
        ):
            slide.shapes._spTree.remove(
                shape._element
            )

    for timing in list(
        slide._element.findall(
            ptag("timing")
        )
    ):
        slide._element.remove(
            timing
        )


def insert_wav(
    slide,
    audio_source,
):
    if isinstance(
        audio_source,
        (str, Path),
    ):
        path = Path(
            audio_source
        )

        shape = slide.shapes.add_movie(
            str(path),
            left=0,
            top=0,
            width=320040,
            height=320040,
            mime_type="audio/wav",
        )

        return shape.shape_id

    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False,
    ) as tmp:

        tmp.write(
            audio_source
        )

        path = Path(
            tmp.name
        )

    try:
        return insert_wav(
            slide,
            path,
        )

    finally:
        path.unlink(
            missing_ok=True
        )


# ============================================================
# AUTOMATIC AUDIO PLAYBACK
# ============================================================

def configure_autoplay(
    slide,
    shape_id,
    duration_ms,
    advance=True,
):
    for old in list(
        slide._element.findall(
            ptag("timing")
        )
    ):
        slide._element.remove(
            old
        )

    timing = etree.Element(
        ptag("timing")
    )

    tn_list = etree.SubElement(
        timing,
        ptag("tnLst"),
    )

    par = etree.SubElement(
        tn_list,
        ptag("par"),
    )

    root = etree.SubElement(
        par,
        ptag("cTn"),
        id="1",
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )

    children = etree.SubElement(
        root,
        ptag("childTnLst"),
    )

    audio = etree.SubElement(
        children,
        ptag("audio"),
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
        ptag("stCondLst"),
    )

    etree.SubElement(
        conditions,
        ptag("cond"),
        delay="0",
    )

    target = etree.SubElement(
        media,
        ptag("tgtEl"),
    )

    etree.SubElement(
        target,
        ptag("spTgt"),
        spid=str(shape_id),
    )

    extension = slide._element.find(
        ptag("extLst")
    )

    if extension is not None:
        slide._element.insert(
            slide._element.index(
                extension
            ),
            timing,
        )
    else:
        slide._element.append(
            timing
        )

    transition = slide._element.find(
        ptag("transition")
    )

    if transition is None:
        transition = etree.Element(
            ptag("transition")
        )

        slide._element.insert(
            slide._element.index(
                timing
            ),
            transition,
        )

    if advance:
        transition.set(
            "advClick",
            "0",
        )

        transition.set(
            "advTm",
            str(
                duration_ms
                + PAUSE_AFTER_AUDIO_MS
            ),
        )

    else:
        transition.set(
            "advClick",
            "1",
        )

        transition.attrib.pop(
            "advTm",
            None,
        )



# ============================================================
# V4.6 PORTABLE CHECKPOINT SUPPORT
# ============================================================

import hashlib
import zipfile
import shutil

PORTABLE_PROGRESS = "pptx_narrator_v46_progress.json"


def notes_digest(notes):
    return hashlib.sha256(
        notes.encode("utf-8")
    ).hexdigest()


def portable_metadata(prs, completed, skipped):
    completed = int(completed)

    return {
        "format": "PPTX_NARRATOR_PORTABLE_CHECKPOINT",
        "version": "4.6",
        "completed": completed,
        "total": len(prs.slides),
        "skipped_empty_notes": int(skipped),
        "notes_sha256": [
            notes_digest(get_notes(prs.slides[i]))
            for i in range(completed)
        ],
        "saved_at": int(time.time()),
    }


def attach_portable_metadata(path, metadata):
    """
    Embed progress information directly inside PPTX.
    """

    temporary = path.with_name(
        path.name + ".portable.tmp"
    )

    try:
        with zipfile.ZipFile(path, "r") as original:
            with zipfile.ZipFile(
                temporary, "w"
            ) as updated:

                for member in original.infolist():
                    if member.filename == PORTABLE_PROGRESS:
                        continue

                    updated.writestr(
                        member,
                        original.read(member.filename),
                    )

                updated.writestr(
                    PORTABLE_PROGRESS,
                    json.dumps(
                        metadata,
                        ensure_ascii=False,
                    ),
                )

        os.replace(temporary, path)

    finally:
        temporary.unlink(missing_ok=True)


def read_portable_metadata(path, prs):
    """
    Validate recovery progress and embedded audio.
    """

    try:
        with zipfile.ZipFile(path, "r") as archive:
            if PORTABLE_PROGRESS not in archive.namelist():
                return None

            info = archive.getinfo(PORTABLE_PROGRESS)

            if info.file_size > 1024 * 1024:
                raise ValueError(
                    "Checkpoint metadata too large"
                )

            raw = archive.read(PORTABLE_PROGRESS)
            meta = json.loads(raw)

    except (
        zipfile.BadZipFile,
        json.JSONDecodeError,
        KeyError,
    ) as exc:
        raise ValueError(
            "Invalid checkpoint metadata"
        ) from exc

    if not isinstance(meta, dict):
        raise ValueError(
            "Checkpoint metadata must be an object"
        )

    if meta.get("format") != (
        "PPTX_NARRATOR_PORTABLE_CHECKPOINT"
    ):
        raise ValueError(
            "Unrecognized checkpoint format"
        )

    if meta.get("version") != "4.6":
        raise ValueError(
            "Unsupported checkpoint version"
        )

    total = len(prs.slides)
    completed = meta.get("completed")
    skipped = meta.get("skipped_empty_notes")
    hashes = meta.get("notes_sha256")

    if (
        type(completed) is not int
        or not 0 <= completed <= total
        or type(skipped) is not int
        or not 0 <= skipped <= completed
        or type(meta.get("total")) is not int
        or meta["total"] != total
        or not isinstance(hashes, list)
        or len(hashes) != completed
    ):
        raise ValueError(
            "Invalid checkpoint progress"
        )

    actual_skipped = 0

    for i in range(completed):
        slide = prs.slides[i]
        notes = get_notes(slide)

        if hashes[i] != notes_digest(notes):
            raise ValueError(
                f"Checkpoint notes changed "
                f"on slide {i + 1}"
            )

        if notes:
            has_audio = any(
                shape.shape_type == MSO_SHAPE_TYPE.MEDIA
                for shape in slide.shapes
            )

            if not has_audio:
                raise ValueError(
                    f"Checkpoint audio missing "
                    f"on slide {i + 1}"
                )
        else:
            actual_skipped += 1

    if actual_skipped != skipped:
        raise ValueError(
            "Skipped-slide count does not match"
        )

    return meta


# ============================================================
# CHECKPOINT VALIDATION
# ============================================================

def validate_pptx(
    path,
    expected_slides,
):
    with zipfile.ZipFile(
        path
    ) as archive:

        if archive.testzip() is not None:
            raise RuntimeError(
                "PPTX ZIP validation failed"
            )

    prs = Presentation(
        str(path)
    )

    if len(prs.slides) != expected_slides:
        raise RuntimeError(
            "Slide count changed"
        )


def checkpoint_path(job_dir):
    return (
        job_dir
        / "checkpoint.pptx"
    )


def checkpoint_meta_path(job_dir):
    return (
        job_dir
        / "checkpoint.json"
    )


def save_checkpoint(
    prs,
    job_dir,
    total,
    completed,
    skipped,
):
    temporary = (
        job_dir
        / "checkpoint.tmp.pptx"
    )

    prs.save(
        str(temporary)
    )

    validate_pptx(
        temporary,
        total,
    )

    attach_portable_metadata(
        temporary,
        portable_metadata(
            prs,
            completed,
            skipped,
        ),
    )

    validate_pptx(
        temporary,
        total,
    )

    os.replace(
        temporary,
        checkpoint_path(job_dir),
    )

    save_json(
        checkpoint_meta_path(job_dir),
        {
            "completed": completed,
            "skipped_empty_notes": skipped,
            "total": total,
            "saved_at": int(time.time()),
        },
    )


def restore_checkpoint_status(
    job_dir,
    status,
):
    checkpoint = checkpoint_path(
        job_dir
    )

    metadata = checkpoint_meta_path(
        job_dir
    )

    if not checkpoint.is_file():
        return status

    if not metadata.is_file():
        try:
            prs = Presentation(str(checkpoint))

            portable = read_portable_metadata(
                checkpoint,
                prs,
            )

            if (
                portable
                and portable["total"] == status["total"]
            ):
                status["completed"] = (
                    portable["completed"]
                )

                status["skipped_empty_notes"] = (
                    portable["skipped_empty_notes"]
                )

                status["checkpoint_available"] = True

        except Exception:
            pass

        return status

    try:
        saved = json.loads(
            metadata.read_text(
                encoding="utf-8"
            )
        )

        completed = int(
            saved["completed"]
        )

        if (
            0 <= completed <= status["target"]
            and int(
                saved["total"]
            ) == status["total"]
        ):
            status["completed"] = completed

            status["skipped_empty_notes"] = int(
                saved.get(
                    "skipped_empty_notes",
                    0,
                )
            )

            status["checkpoint_available"] = True

    except (
        OSError,
        ValueError,
        KeyError,
        TypeError,
    ):
        pass

    return status


# ============================================================
# BACKGROUND JOB PROCESSING
# ============================================================

def process_job(job_id):
    job_dir = (
        WORKDIR
        / job_id
    )

    try:
        status = read_status(
            job_id
        )

    except HTTPException:
        with active_lock:
            active_jobs.discard(
                job_id
            )
        return

    try:
        status = restore_checkpoint_status(
            job_dir,
            status,
        )

        status["state"] = "running"
        status["error"] = None

        update_status(
            job_dir,
            status,
        )

        total = status["total"]
        target = status["target"]
        batch_size = status["batch_size"]

        checkpoint = checkpoint_path(
            job_dir
        )

        source = (
            checkpoint
            if checkpoint.is_file()
            else job_dir / "input.pptx"
        )

        prs = Presentation(
            str(source)
        )

        completed = int(
            status["completed"]
        )

        for i in range(
            completed,
            target,
        ):
            slide_number = i + 1

            status["current_slide"] = (
                slide_number
            )

            status["audio_slide"] = (
                slide_number
            )

            status["audio_segment"] = 0
            status["audio_chunk"] = 0

            update_status(
                job_dir,
                status,
            )

            slide = prs.slides[i]

            note = get_notes(
                slide
            )

            if note:
                (
                    audio_path,
                    duration_ms,
                    segment_count,
                ) = synthesize_slide(
                    note,
                    job_dir,
                    slide_number,
                    status,
                )

                remove_old_audio(
                    slide
                )

                shape_id = insert_wav(
                    slide,
                    audio_path,
                )

                configure_autoplay(
                    slide,
                    shape_id,
                    duration_ms,
                    advance=(
                        i < total - 1
                    ),
                )

                status["last_slide_seconds"] = round(
                    duration_ms / 1000,
                    1,
                )

                status["last_slide_chunks"] = (
                    segment_count
                )

                status["last_slide_segments"] = (
                    segment_count
                )

            else:
                status["skipped_empty_notes"] += 1

            if (
                (
                    i + 1
                    - int(status.get("resumed_from", 0))
                ) % batch_size == 0
                or i + 1 == target
            ):
                save_checkpoint(
                    prs,
                    job_dir,
                    total,
                    i + 1,
                    status["skipped_empty_notes"],
                )

                status["completed"] = (
                    i + 1
                )

                status["checkpoint_available"] = True

            update_status(
                job_dir,
                status,
            )

        status["state"] = "completed"

        status["current_slide"] = target

        update_status(
            job_dir,
            status,
        )

    except Exception as exc:
        status["state"] = "failed"

        status["error"] = (
            f"Slide "
            f"{status.get('current_slide', 0)}: "
            f"{str(exc)[:400]}"
        )

        try:
            update_status(
                job_dir,
                status,
            )
        except Exception:
            pass

    finally:
        with active_lock:
            active_jobs.discard(
                job_id
            )


# ============================================================
# JOB LAUNCH
# ============================================================

def launch_job(job_id):
    with active_lock:
        if job_id in active_jobs:
            return False

        if len(active_jobs) >= MAX_ACTIVE_JOBS:
            raise HTTPException(
                429,
                "Another narration job is running; "
                "please wait",
            )

        active_jobs.add(
            job_id
        )

    try:
        executor.submit(
            process_job,
            job_id,
        )

    except Exception:
        with active_lock:
            active_jobs.discard(
                job_id
            )
        raise

    return True


# ============================================================
# CREATE NARRATION JOB
# ============================================================

@app.post("/jobs")
async def create_job(
    file: UploadFile = File(...),
    notes_file: UploadFile = File(None),
    limit: int = Form(0),
    batch_size: int = Form(5),
):
    if (
        not file.filename
        or Path(
            file.filename
        ).suffix.lower() != ".pptx"
    ):
        raise HTTPException(
            400,
            "Upload a PPTX file",
        )

    if (
        limit < 0
        or not 1 <= batch_size <= 20
    ):
        raise HTTPException(
            400,
            "Invalid slide limit or batch size (1-20)",
        )

    if not MODEL.is_file():
        raise HTTPException(
            500,
            "Piper model is missing",
        )

    if shutil.which("piper") is None:
        raise HTTPException(
            500,
            "Piper executable is missing",
        )

    with active_lock:
        if len(active_jobs) >= MAX_ACTIVE_JOBS:
            raise HTTPException(
                429,
                "Another narration job is running; "
                "please wait",
            )

    job_id = uuid.uuid4().hex

    job_dir = (
        WORKDIR
        / job_id
    )

    job_dir.mkdir(
        parents=True,
        exist_ok=False,
    )

    source = (
        job_dir
        / "input.pptx"
    )

    try:
        await save_upload(
            file,
            source,
            MAX_BYTES,
        )

        try:
            prs = Presentation(
                str(source)
            )

        except Exception as exc:
            raise HTTPException(
                400,
                "Invalid PPTX",
            ) from exc

        total = len(
            prs.slides
        )

        if not total:
            raise HTTPException(
                400,
                "PPTX contains no slides",
            )

        target = (
            min(limit, total)
            if limit
            else total
        )

        try:
            portable = read_portable_metadata(
                source,
                prs,
            )

        except ValueError as exc:
            raise HTTPException(
                400,
                f"Invalid recovery checkpoint: {exc}",
            ) from exc

        recovered = (
            portable["completed"]
            if portable
            else 0
        )

        recovered_skipped = (
            portable["skipped_empty_notes"]
            if portable
            else 0
        )

        if recovered > target:
            raise HTTPException(
                400,
                "Slides to Process must be 0 "
                "or at least the saved slide count",
            )

        if (
            portable
            and notes_file is not None
            and notes_file.filename
        ):
            raise HTTPException(
                400,
                "Do not upload a new script "
                "when resuming a checkpoint",
            )

        imported_count = 0

        if (
            notes_file is not None
            and notes_file.filename
        ):
            data = await notes_file.read(
                MAX_NOTES_BYTES + 1
            )

            if len(data) > MAX_NOTES_BYTES:
                raise HTTPException(
                    413,
                    "Notes exceed 25 MB",
                )

            try:
                imported_count = import_external_notes(
                    prs,
                    data,
                    notes_file.filename,
                )

            except Exception as exc:
                raise HTTPException(
                    400,
                    "Notes import failed: "
                    + str(exc)[:400],
                ) from exc

            prs.save(
                str(source)
            )

        notes_found = sum(
            bool(
                get_notes(
                    prs.slides[i]
                )
            )
            for i in range(target)
        )

        if not notes_found:
            raise HTTPException(
                400,
                "No Speaker Notes found; "
                "run Step 1 or upload a script",
            )

        status = {
            "job_id": job_id,
            "state": "queued",
            "total": total,
            "target": target,
            "completed": recovered,
            "current_slide": recovered,
            "resumed_from": recovered,
            "portable_checkpoint": bool(portable),
            "batch_size": batch_size,
            "imported_notes": imported_count,
            "notes_in_target": notes_found,
            "skipped_empty_notes": recovered_skipped,
            "checkpoint_available": bool(portable),
            "error": None,
            "created_at": int(time.time()),
            "audio_slide": 0,
            "audio_segment": 0,
            "audio_segment_total": 0,
            "audio_chunk": 0,
            "audio_chunk_total": 0,
        }

        if portable:
            shutil.copy2(
                source,
                checkpoint_path(job_dir),
            )

            save_json(
                checkpoint_meta_path(job_dir),
                {
                    "completed": recovered,
                    "skipped_empty_notes": recovered_skipped,
                    "total": total,
                    "saved_at": int(time.time()),
                },
            )

        update_status(
            job_dir,
            status,
        )

        launch_job(
            job_id
        )

        return status

    except Exception:
        shutil.rmtree(
            job_dir,
            ignore_errors=True,
        )
        raise


# ============================================================
# JOB STATUS
# ============================================================

@app.get("/jobs/{job_id}")
def job_status(job_id: str):
    status = read_status(
        job_id
    )

    job_dir = (
        WORKDIR
        / job_id
    )

    status = restore_checkpoint_status(
        job_dir,
        status,
    )

    with active_lock:
        active = (
            job_id
            in active_jobs
        )

    if (
        status["state"]
        in (
            "running",
            "queued",
        )
        and not active
    ):
        status["state"] = "interrupted"

    status["download_url"] = (
        f"/jobs/{job_id}/download"
        if checkpoint_path(
            job_dir
        ).is_file()
        else None
    )

    audio_dir = (
        job_dir
        / "audio"
    )

    status["audio_download_url"] = (
        f"/jobs/{job_id}/audio"
        if audio_dir.is_dir()
        else None
    )

    return status


# ============================================================
# RESUME JOB
# ============================================================

@app.post("/jobs/{job_id}/resume")
def resume_job(job_id: str):
    status = read_status(
        job_id
    )

    if status["state"] == "completed":
        return status

    with active_lock:
        if job_id in active_jobs:
            raise HTTPException(
                409,
                "Job is already running",
            )

    status["state"] = "queued"
    status["error"] = None

    update_status(
        WORKDIR / job_id,
        status,
    )

    launch_job(
        job_id
    )

    return status


# ============================================================
# DOWNLOAD NARRATED POWERPOINT
# ============================================================

@app.get("/jobs/{job_id}/download")
def download_job(job_id: str):
    status = restore_checkpoint_status(
        WORKDIR / job_id,
        read_status(job_id),
    )

    path = checkpoint_path(
        WORKDIR / job_id
    )

    if not path.is_file():
        raise HTTPException(
            404,
            "No checkpoint available",
        )

    return FileResponse(
        str(path),
        media_type=PPTX_MIME,
        filename=(
            f"narrated_v4_6_"
            f"{status['completed']}_of_"
            f"{status['total']}.pptx"
        ),
    )


# ============================================================
# DOWNLOAD WAV SEGMENTS ZIP
# ============================================================

@app.get("/jobs/{job_id}/audio")
def download_audio(job_id: str):
    status = restore_checkpoint_status(
        WORKDIR / job_id,
        read_status(job_id),
    )

    job_dir = (
        WORKDIR
        / job_id
    )

    audio_dir = (
        job_dir
        / "audio"
    )

    if (
        not audio_dir.is_dir()
        or not any(
            audio_dir.rglob(
                "*.wav"
            )
        )
    ):
        raise HTTPException(
            404,
            "No generated audio available",
        )

    archive_path = (
        job_dir
        / "audio_segments.zip"
    )

    temporary = (
        job_dir
        / "audio_segments.tmp.zip"
    )

    with zipfile.ZipFile(
        temporary,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:

        for wav in sorted(
            audio_dir.rglob(
                "*.wav"
            )
        ):
            match = re.match(
                r"slide_(\d+)",
                wav.parent.name,
            )

            if (
                match
                and int(
                    match.group(1)
                )
                <= int(
                    status.get(
                        "completed",
                        0,
                    )
                )
            ):
                archive.write(
                    wav,
                    wav.relative_to(
                        audio_dir
                    ),
                )

    os.replace(
        temporary,
        archive_path,
    )

    return FileResponse(
        str(archive_path),
        media_type="application/zip",
        filename=(
            f"narration_audio_v4_6_"
            f"{status['completed']}_slides.zip"
        ),
    )


# ============================================================
# LEGACY /PROCESS COMPATIBILITY
# ============================================================

@app.post("/process")
async def legacy_process(
    file: UploadFile = File(...),
    mode: str = Form("tts"),
    voice: str = Form("alloy"),
    api_key: str = Form(""),
    limit: int = Form(0),
):
    raw = await file.read(
        MAX_BYTES + 1
    )

    if len(raw) > MAX_BYTES:
        raise HTTPException(
            413,
            "PPTX exceeds 150 MB",
        )

    try:
        prs = Presentation(
            io.BytesIO(raw)
        )

    except Exception as exc:
        raise HTTPException(
            400,
            "Invalid PPTX",
        ) from exc

    if limit < 0:
        raise HTTPException(
            400,
            "Invalid slide limit",
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

    if mode == "notes":
        data = "\n\n".join(
            f"=== Slide {i + 1} ===\n{note}"
            for i, note in enumerate(
                notes
            )
        )

        return StreamingResponse(
            io.BytesIO(
                data.encode(
                    "utf-8"
                )
            ),
            media_type=(
                "text/plain; charset=utf-8"
            ),
            headers={
                "Content-Disposition":
                    'attachment; filename="slide_notes.txt"'
            },
        )

    if mode == "existing":
        return StreamingResponse(
            io.BytesIO(raw),
            media_type=PPTX_MIME,
            headers={
                "Content-Disposition":
                    'attachment; filename="original_audio.pptx"'
            },
        )

    if mode != "tts":
        raise HTTPException(
            400,
            "Invalid mode",
        )

    if not any(notes):
        raise HTTPException(
            400,
            "No Speaker Notes found",
        )

    for i, note in enumerate(
        notes
    ):
        if not note:
            continue

        try:
            audio, duration, _ = (
                generate_complete_wav(
                    note
                )
            )

            slide = prs.slides[i]

            remove_old_audio(
                slide
            )

            shape_id = insert_wav(
                slide,
                audio,
            )

            configure_autoplay(
                slide,
                shape_id,
                duration,
                advance=(
                    i < total - 1
                ),
            )

        except Exception as exc:
            raise HTTPException(
                500,
                f"Slide {i + 1}: "
                f"{str(exc)[:300]}",
            ) from exc

    output = io.BytesIO()

    prs.save(
        output
    )

    output.seek(0)

    return StreamingResponse(
        output,
        media_type=PPTX_MIME,
        headers={
            "Content-Disposition":
                'attachment; filename='
                '"narrated_piper_v4_6.pptx"'
        },
    )

