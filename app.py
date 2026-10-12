"""PPTX Narrator V4 — Piper TTS with background jobs."""

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

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    StreamingResponse,
)

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part


if not hasattr(Part, "sha1"):
    Part.sha1 = property(
        lambda self: hashlib.sha1(self.blob).hexdigest()
    )


ROOT = Path(__file__).resolve().parent
MODEL = ROOT / "models" / "en_US-lessac-low.onnx"

WORKDIR = Path(
    os.environ.get(
        "NARRATOR_WORKDIR",
        str(ROOT / "narrator_jobs"),
    )
)
WORKDIR.mkdir(parents=True, exist_ok=True)

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


app = FastAPI(title="PPTX Narrator V4")

executor = ThreadPoolExecutor(max_workers=1)
active_lock = threading.Lock()
active_jobs = set()


# =====================================================
# JOB STORAGE
# =====================================================

def atomic_json(path, data):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(tmp, path)


def load_status(job_id):
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise HTTPException(400, "Invalid job ID")

    path = WORKDIR / job_id / "status.json"

    if not path.is_file():
        raise HTTPException(
            404,
            "Job not found. Server may have restarted.",
        )

    return json.loads(
        path.read_text(encoding="utf-8")
    )


def update(job_dir, status):
    status["updated_at"] = int(time.time())
    atomic_json(job_dir / "status.json", status)


# =====================================================
# SPEAKER NOTES
# =====================================================

def get_notes(slide):
    if not slide.has_notes_slide:
        return ""

    frame = slide.notes_slide.notes_text_frame

    if frame is None:
        return ""

    return frame.text.strip()


# =====================================================
# SPLIT LONG NARRATION
# =====================================================

def split_narration(text, max_chars=240):
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

        if len(current) + len(sentence) + 1 <= max_chars:
            current = (current + " " + sentence).strip()
            continue

        if current:
            chunks.append(current)
            current = ""

        while len(sentence) > max_chars:
            cut = sentence.rfind(
                " ", 0, max_chars + 1
            )

            if cut <= 0:
                cut = max_chars

            chunks.append(sentence[:cut].strip())
            sentence = sentence[cut:].strip()

        current = sentence

    if current:
        chunks.append(current)

    return chunks


# =====================================================
# PIPER AUDIO GENERATION
# =====================================================

def piper_chunk(text, path):
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

    if (
        result.returncode != 0
        or not path.is_file()
        or path.stat().st_size < 100
    ):
        raise RuntimeError(
            result.stderr[-500:]
            or "Piper failed to create audio"
        )


def generate_complete_wav(text):
    if not MODEL.is_file():
        raise RuntimeError(
            "Piper voice model not found"
        )

    chunks = split_narration(text)

    if not chunks:
        raise RuntimeError(
            "Speaker Notes are empty"
        )

    with tempfile.TemporaryDirectory() as td:
        directory = Path(td)
        paths = []

        for i, chunk in enumerate(chunks):
            path = directory / f"part_{i:04d}.wav"
            piper_chunk(chunk, path)
            paths.append(path)

        output = directory / "merged.wav"

        audio_format = None
        frames_count = 0

        with wave.open(str(output), "wb") as dest:
            for path in paths:
                with wave.open(str(path), "rb") as src:
                    current_format = (
                        src.getnchannels(),
                        src.getsampwidth(),
                        src.getframerate(),
                        src.getcomptype(),
                    )

                    if audio_format is None:
                        audio_format = current_format

                        dest.setnchannels(audio_format[0])
                        dest.setsampwidth(audio_format[1])
                        dest.setframerate(audio_format[2])

                    elif current_format != audio_format:
                        raise RuntimeError(
                            "Audio chunk formats differ"
                        )

                    frames = src.readframes(
                        src.getnframes()
                    )

                    if not frames:
                        raise RuntimeError(
                            "Audio chunk is empty"
                        )

                    dest.writeframes(frames)
                    frames_count += src.getnframes()

        duration_ms = round(
            1000 * frames_count / audio_format[2]
        )

        return (
            output.read_bytes(),
            duration_ms,
            len(chunks),
        )


# =====================================================
# REMOVE OLD AUDIO
# =====================================================

def remove_old_audio(slide):
    # As in V3, this also removes slide animations.

    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
            slide.shapes._spTree.remove(
                shape._element
            )

    for timing in list(
        slide._element.findall(ptag("timing"))
    ):
        slide._element.remove(timing)


# =====================================================
# INSERT AUDIO
# =====================================================

def insert_wav(slide, audio_bytes):
    with tempfile.NamedTemporaryFile(
        suffix=".wav",
        delete=False,
    ) as temp:
        temp.write(audio_bytes)
        path = Path(temp.name)

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
        path.unlink(missing_ok=True)


# =====================================================
# AUTOPLAY AND AUTOMATIC ADVANCEMENT
# =====================================================

def configure_autoplay(
    slide,
    shape_id,
    duration_ms,
    advance=True,
):
    for old in list(
        slide._element.findall(ptag("timing"))
    ):
        slide._element.remove(old)

    timing = etree.Element(ptag("timing"))

    tn_list = etree.SubElement(
        timing, ptag("tnLst")
    )

    par = etree.SubElement(
        tn_list, ptag("par")
    )

    root = etree.SubElement(
        par,
        ptag("cTn"),
        id="1",
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )

    child = etree.SubElement(
        root, ptag("childTnLst")
    )

    audio = etree.SubElement(
        child, ptag("audio")
    )

    media = etree.SubElement(
        audio,
        ptag("cMediaNode"),
        vol="80000",
    )

    ct = etree.SubElement(
        media,
        ptag("cTn"),
        id="2",
        fill="hold",
        display="0",
    )

    conditions = etree.SubElement(
        ct, ptag("stCondLst")
    )

    etree.SubElement(
        conditions,
        ptag("cond"),
        delay="0",
    )

    target = etree.SubElement(
        media, ptag("tgtEl")
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
            slide._element.index(extension),
            timing,
        )
    else:
        slide._element.append(timing)

    transition = slide._element.find(
        ptag("transition")
    )

    if transition is None:
        transition = etree.Element(
            ptag("transition")
        )

        slide._element.insert(
            slide._element.index(timing),
            transition,
        )

    if advance:
        transition.set("advClick", "0")
        transition.set(
            "advTm",
            str(
                duration_ms
                + PAUSE_AFTER_AUDIO_MS
            ),
        )
    else:
        transition.set("advClick", "1")
        transition.attrib.pop("advTm", None)


# =====================================================
# SAVE AND VALIDATE CHECKPOINT
# =====================================================

def validate_pptx(path, total):
    with zipfile.ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise RuntimeError(
                "PPTX ZIP integrity failed"
            )

    if len(Presentation(str(path)).slides) != total:
        raise RuntimeError(
            "PPTX slide count changed"
        )


def save_checkpoint(prs, job_dir, total):
    destination = job_dir / "checkpoint.pptx"
    temporary = job_dir / "checkpoint.tmp.pptx"

    prs.save(str(temporary))
    validate_pptx(temporary, total)

    os.replace(
        temporary,
        destination,
    )


# =====================================================
# BACKGROUND PROCESSING
# =====================================================

def process_job(job_id):
    job_dir = WORKDIR / job_id
    status = load_status(job_id)

    try:
        status["state"] = "running"
        update(job_dir, status)

        total = status["total"]
        count = status["target"]
        batch_size = status["batch_size"]

        checkpoint = job_dir / "checkpoint.pptx"

        source = (
            checkpoint
            if checkpoint.is_file()
            else job_dir / "input.pptx"
        )

        prs = Presentation(str(source))

        completed = int(status["completed"])

        for i in range(completed, count):
            slide = prs.slides[i]
            note = get_notes(slide)

            if note:
                (
                    audio_bytes,
                    duration,
                    chunks,
                ) = generate_complete_wav(note)

                remove_old_audio(slide)

                shape_id = insert_wav(
                    slide,
                    audio_bytes,
                )

                configure_autoplay(
                    slide,
                    shape_id,
                    duration,
                    advance=(i < total - 1),
                )

                status["last_slide_seconds"] = round(
                    duration / 1000, 1
                )

                status["last_slide_chunks"] = chunks

            else:
                status["skipped_empty_notes"] += 1

            status["current_slide"] = i + 1

            if (
                (i + 1) % batch_size == 0
                or i + 1 == count
            ):
                save_checkpoint(
                    prs,
                    job_dir,
                    total,
                )

                status["completed"] = i + 1
                status["checkpoint_available"] = True

            update(job_dir, status)

        status["state"] = "completed"
        status["current_slide"] = count
        update(job_dir, status)

    except Exception as exc:
        status["state"] = "failed"

        status["error"] = (
            f"Slide "
            f"{status.get('current_slide', status['completed']) + 1}: "
            f"{str(exc)[:400]}"
        )

        update(job_dir, status)

    finally:
        with active_lock:
            active_jobs.discard(job_id)


def launch(job_id):
    with active_lock:
        if job_id in active_jobs:
            return False

        active_jobs.add(job_id)

    executor.submit(process_job, job_id)
    return True


# =====================================================
# WEBSITE
# =====================================================

@app.get("/", response_class=HTMLResponse)
def home():
    return (
        ROOT / "index_v4.html"
    ).read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": "4.0",
        "engine": "Piper TTS",
        "model_exists": MODEL.is_file(),
        "api_key_required": False,
        "checkpoint_storage": str(WORKDIR),
        "storage_persistent": False,
    }


# =====================================================
# CREATE JOB
# =====================================================

@app.post("/jobs")
async def create_job(
    file: UploadFile = File(...),
    limit: int = Form(0),
    batch_size: int = Form(5),
):
    if (
        not file.filename
        or not file.filename.lower().endswith(".pptx")
    ):
        raise HTTPException(
            400,
            "Please upload a PPTX",
        )

    if limit < 0 or not 1 <= batch_size <= 20:
        raise HTTPException(
            400,
            "Invalid limit or batch size",
        )

    if not MODEL.is_file():
        raise HTTPException(
            500,
            "Piper model missing",
        )

    job_id = uuid.uuid4().hex
    job_dir = WORKDIR / job_id
    job_dir.mkdir()

    source = job_dir / "input.pptx"

    try:
        size = 0

        with source.open("wb") as output:
            while True:
                chunk = await file.read(
                    1024 * 1024
                )

                if not chunk:
                    break

                size += len(chunk)

                if size > MAX_BYTES:
                    raise HTTPException(
                        413,
                        "PPTX exceeds 150 MB",
                    )

                output.write(chunk)

        try:
            prs = Presentation(str(source))

        except Exception:
            raise HTTPException(
                400,
                "Invalid PPTX",
            )

        total = len(prs.slides)

        if total == 0:
            raise HTTPException(
                400,
                "PPTX has no slides",
            )

        target = (
            min(limit, total)
            if limit
            else total
        )

        if not any(
            get_notes(prs.slides[i])
            for i in range(target)
        ):
            raise HTTPException(
                400,
                "No Speaker Notes found",
            )

        status = {
            "job_id": job_id,
            "state": "queued",
            "total": total,
            "target": target,
            "completed": 0,
            "current_slide": 0,
            "batch_size": batch_size,
            "skipped_empty_notes": 0,
            "checkpoint_available": False,
            "error": None,
            "created_at": int(time.time()),
        }

        update(job_dir, status)
        launch(job_id)

        return status

    except Exception:
        shutil.rmtree(
            job_dir,
            ignore_errors=True,
        )
        raise


# =====================================================
# JOB STATUS
# =====================================================

@app.get("/jobs/{job_id}")
def job_status(job_id: str):
    status = load_status(job_id)

    with active_lock:
        is_active = job_id in active_jobs

    if (
        status["state"] == "running"
        and not is_active
    ):
        status["state"] = "interrupted"

    status["download_url"] = (
        f"/jobs/{job_id}/download"
        if status["checkpoint_available"]
        else None
    )

    return status


# =====================================================
# RESUME JOB
# =====================================================

@app.post("/jobs/{job_id}/resume")
def resume(job_id: str):
    status = load_status(job_id)

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

    update(
        WORKDIR / job_id,
        status,
    )

    launch(job_id)
    return status


# =====================================================
# DOWNLOAD CHECKPOINT
# =====================================================

@app.get("/jobs/{job_id}/download")
def download_checkpoint(job_id: str):
    status = load_status(job_id)

    path = (
        WORKDIR
        / job_id
        / "checkpoint.pptx"
    )

    if (
        not status["checkpoint_available"]
        or not path.is_file()
    ):
        raise HTTPException(
            404,
            "No saved batch yet",
        )

    return FileResponse(
        str(path),
        media_type=PPTX_MIME,
        filename=(
            f"narrated_v4_"
            f"{status['completed']}_"
            f"of_{status['total']}.pptx"
        ),
    )


# =====================================================
# V3-COMPATIBLE PROCESS ENDPOINT
# =====================================================

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
            "File exceeds 150 MB",
        )

    try:
        prs = Presentation(
            io.BytesIO(raw)
        )

    except Exception:
        raise HTTPException(
            400,
            "Invalid PPTX",
        )

    total = len(prs.slides)

    count = (
        min(limit, total)
        if limit
        else total
    )

    notes = [
        get_notes(prs.slides[i])
        for i in range(count)
    ]

    if mode == "notes":
        content = "\n\n".join(
            f"=== Slide {i + 1} ===\n{note}"
            for i, note in enumerate(notes)
        )

        return StreamingResponse(
            io.BytesIO(
                content.encode("utf-8")
            ),
            media_type="text/plain; charset=utf-8",
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

    for i, note in enumerate(notes):
        if not note:
            continue

        try:
            (
                audio,
                duration,
                _,
            ) = generate_complete_wav(note)

            remove_old_audio(
                prs.slides[i]
            )

            shape_id = insert_wav(
                prs.slides[i],
                audio,
            )

            configure_autoplay(
                prs.slides[i],
                shape_id,
                duration,
                advance=(i < total - 1),
            )

        except Exception as exc:
            raise HTTPException(
                500,
                f"Slide {i + 1}: {str(exc)[:300]}",
            )

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type=PPTX_MIME,
        headers={
            "Content-Disposition":
            'attachment; filename="narrated_piper_v4.pptx"'
        },
    )
