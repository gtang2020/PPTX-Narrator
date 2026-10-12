"""
PPTX Narrator V4.6 Upgrade
==========================

Upgrade from V4.5 using original backups.

Required:
    app_v4_5.py
    index_v4_5.html

Generated:
    app.py
    index_v4.html

Features:
    Portable PPTX checkpoint metadata
    Recovery from uploaded PPTX
    Resume without repeating completed slides
    Checkpoint validation
    V4.6 frontend

Run:
    python upgrade_v46.py
"""

from pathlib import Path
import ast
import os
import tempfile


# ============================================================
# FILES
# ============================================================

ROOT = Path(__file__).resolve().parent

BACKEND_BACKUP = ROOT / "app_v4_5.py"
FRONTEND_BACKUP = ROOT / "index_v4_5.html"

BACKEND_OUTPUT = ROOT / "app.py"
FRONTEND_OUTPUT = ROOT / "index_v4.html"


# ============================================================
# READ BACKUPS
# ============================================================

for path in (BACKEND_BACKUP, FRONTEND_BACKUP):
    if not path.is_file():
        raise SystemExit(
            f"Missing required backup: {path.name}"
        )

py = BACKEND_BACKUP.read_text(encoding="utf-8")
html = FRONTEND_BACKUP.read_text(encoding="utf-8")


# ============================================================
# SAFE REPLACEMENT
# ============================================================

def replace_once(source, old, new, label):
    count = source.count(old)

    if count != 1:
        raise RuntimeError(
            f"{label}: expected 1 match, found {count}. "
            "No live files changed."
        )

    return source.replace(old, new, 1)


def replace_version_labels(source, old, new, label):
    count = source.count(old)

    if count < 1:
        raise RuntimeError(
            f"{label}: version label not found. "
            "No live files changed."
        )

    return source.replace(old, new)


# ============================================================
# BACKEND VERSION
# ============================================================

# FIX: V4.5 occurs more than once in the backend.
# Replace all occurrences rather than requiring one.

py = replace_version_labels(
    py,
    "PPTX Narrator V4.5",
    "PPTX Narrator V4.6",
    "Backend version",
)

py = replace_once(
    py,
    '"version": "4.5",',
    '"version": "4.6",',
    "Health version",
)

py = replace_once(
    py,
    '"segment_resume": True,',
    '''"segment_resume": True,
        "portable_checkpoint": True,
        "resume_from_uploaded_pptx": True,''',
    "Health capabilities",
)


# ============================================================
# PORTABLE CHECKPOINT IMPLEMENTATION
# ============================================================

PORTABLE_CODE = '''
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


'''


# ============================================================
# INSERT PORTABLE FUNCTIONS
# ============================================================

MARKER = """# ============================================================
# CHECKPOINT VALIDATION
# ============================================================"""

py = replace_once(
    py,
    MARKER,
    PORTABLE_CODE + MARKER,
    "Portable checkpoint insertion",
)


# ============================================================
# SAVE PORTABLE CHECKPOINT
# ============================================================

py = replace_once(
    py,
    '''    validate_pptx(
        temporary,
        total,
    )

    os.replace(
        temporary,
        checkpoint_path(job_dir),
    )''',
    '''    validate_pptx(
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
    )''',
    "Portable checkpoint saving",
)


# ============================================================
# RECOVER WHEN CHECKPOINT.JSON IS MISSING
# ============================================================

py = replace_once(
    py,
    '''    if (
        not checkpoint.is_file()
        or not metadata.is_file()
    ):
        return status
''',
    '''    if not checkpoint.is_file():
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
''',
    "Checkpoint metadata recovery",
)


# ============================================================
# DETECT UPLOADED RECOVERY PPTX
# ============================================================

py = replace_once(
    py,
    '''        target = (
            min(limit, total)
            if limit
            else total
        )

        imported_count = 0''',
    '''        target = (
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

        imported_count = 0''',
    "Uploaded checkpoint detection",
)


# ============================================================
# RESTORE INITIAL PROGRESS
# ============================================================

py = replace_once(
    py,
    '''            "completed": 0,
            "current_slide": 0,''',
    '''            "completed": recovered,
            "current_slide": recovered,
            "resumed_from": recovered,
            "portable_checkpoint": bool(portable),''',
    "Recovered slide progress",
)

py = replace_once(
    py,
    '''            "skipped_empty_notes": 0,
            "checkpoint_available": False,''',
    '''            "skipped_empty_notes": recovered_skipped,
            "checkpoint_available": bool(portable),''',
    "Recovered skipped slides",
)


# ============================================================
# PRESERVE UPLOADED CHECKPOINT
# ============================================================

py = replace_once(
    py,
    '''        update_status(
            job_dir,
            status,
        )

        launch_job(
            job_id
        )''',
    '''        if portable:
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
        )''',
    "Uploaded checkpoint preservation",
)


# ============================================================
# CHECKPOINT INTERVAL AFTER RECOVERY
# ============================================================

py = replace_once(
    py,
    '''                (i + 1) % batch_size == 0
                or i + 1 == target''',
    '''                (
                    i + 1
                    - int(status.get("resumed_from", 0))
                ) % batch_size == 0
                or i + 1 == target''',
    "Checkpoint interval",
)


# ============================================================
# BACKEND FILENAMES
# ============================================================

py = replace_once(
    py,
    'f"narrated_v4_5_"',
    'f"narrated_v4_6_"',
    "Narrated PPTX filename",
)

py = replace_once(
    py,
    'f"narration_audio_v4_5_"',
    'f"narration_audio_v4_6_"',
    "Audio ZIP filename",
)

py = replace_once(
    py,
    '"PPTX_with_Speaker_Notes_V4_5.pptx"',
    '"PPTX_with_Speaker_Notes_V4_6.pptx"',
    "Notes PPTX filename",
)

py = replace_once(
    py,
    '"narrated_piper_v4_5.pptx"',
    '"narrated_piper_v4_6.pptx"',
    "Legacy PPTX filename",
)


# ============================================================
# FRONTEND VERSION
# ============================================================

# Match specific HTML elements, not generic text.

html = replace_once(
    html,
    "<title>PPTX Narrator V4.5",
    "<title>PPTX Narrator V4.6",
    "HTML title",
)

html = replace_once(
    html,
    "<h1>🎙 PPTX Narrator V4.5</h1>",
    "<h1>🎙 PPTX Narrator V4.6</h1>",
    "HTML heading",
)


# ============================================================
# FRONTEND INSTRUCTIONS
# ============================================================

html = replace_once(
    html,
    "V4.5 includes longer Piper synthesis timeouts and",
    "V4.6 includes portable recovery checkpoints, "
    "longer Piper synthesis timeouts and",
    "Frontend description",
)

html = replace_once(
    html,
    """Render free storage is temporary.
Download completed checkpoints and audio ZIP files promptly.""",
    """Render free storage is temporary.
<strong>Download your recovery PPTX to your iPad after each checkpoint.</strong>
Server-side checkpoints can disappear after restarts.""",
    "Storage warning",
)

html = replace_once(
    html,
    "<strong>3</strong> = first three slides.",
    """<strong>Recovery PPTX: choose 0 to finish remaining slides.</strong>
<br>
<strong>3</strong> = first three slides.""",
    "Slide limit instructions",
)

html = replace_once(
    html,
    """The PPTX download becomes available after a checkpoint
has been saved.""",
    """Download the latest narrated PPTX to your iPad
after each checkpoint.

<br><br>

After a Render restart, upload that saved PPTX
in Step 2 and set Slides to Process = 0.""",
    "Recovery instructions",
)


# ============================================================
# FRONTEND LOCAL STORAGE
# ============================================================

html = replace_once(
    html,
    'const STORAGE_KEY = "pptx_v45_job";',
    'const STORAGE_KEY = "pptx_v46_job";',
    "Local storage key",
)

html = replace_once(
    html,
    'const OLD_STORAGE_KEY = "pptx_v44_job";',
    'const OLD_STORAGE_KEY = "pptx_v45_job";',
    "Previous storage key",
)


# ============================================================
# FRONTEND PROGRESS
# ============================================================

html = replace_once(
    html,
    '''        `Saved slides: ${completed} / ${target}\\n` +''',
    '''        `Saved slides: ${completed} / ${target}\\n` +
        `Recovered from upload: ${status.resumed_from ?? 0} slides\\n` +''',
    "Recovery progress display",
)


# ============================================================
# FRONTEND FILENAMES AND HEALTH
# ============================================================

html = replace_once(
    html,
    '"PPTX_with_Speaker_Notes_V4_5.pptx"',
    '"PPTX_with_Speaker_Notes_V4_6.pptx"',
    "Frontend notes filename",
)

html = replace_once(
    html,
    "Checking V4.5 server configuration…",
    "Checking V4.6 server configuration…",
    "Health placeholder",
)

html = replace_once(
    html,
    'if (version !== "4.5")',
    'if (version !== "4.6")',
    "Backend version comparison",
)

html = replace_once(
    html,
    "Expected backend V4.5.",
    "Expected backend V4.6.",
    "Backend version warning",
)


# ============================================================
# VALIDATION
# ============================================================

ast.parse(py)

required_backend = [
    "def portable_metadata(",
    "def attach_portable_metadata(",
    "def read_portable_metadata(",
    '"version": "4.6"',
]

for marker in required_backend:
    if marker not in py:
        raise RuntimeError(
            f"Generated backend missing: {marker}"
        )

required_frontend = [
    "<script>",
    "</script>",
    "</html>",
    "PPTX Narrator V4.6",
]

for marker in required_frontend:
    if marker not in html:
        raise RuntimeError(
            f"Generated frontend missing: {marker}"
        )


# ============================================================
# STAGE GENERATED FILES
# ============================================================

def stage_text(destination, content):
    fd, filename = tempfile.mkstemp(
        prefix=destination.name + ".",
        suffix=".tmp",
        dir=ROOT,
    )

    temporary = Path(filename)

    try:
        with os.fdopen(
            fd,
            "w",
            encoding="utf-8",
        ) as output:
            output.write(content)

        return temporary

    except Exception:
        temporary.unlink(missing_ok=True)
        raise


staged_backend = None
staged_frontend = None

try:
    staged_backend = stage_text(
        BACKEND_OUTPUT,
        py,
    )

    staged_frontend = stage_text(
        FRONTEND_OUTPUT,
        html,
    )

    os.replace(
        staged_backend,
        BACKEND_OUTPUT,
    )

    os.replace(
        staged_frontend,
        FRONTEND_OUTPUT,
    )

finally:
    if staged_backend is not None:
        staged_backend.unlink(missing_ok=True)

    if staged_frontend is not None:
        staged_frontend.unlink(missing_ok=True)


print("SUCCESS: Generated PPTX Narrator V4.6.")
print("Backend: app.py")
print("Frontend: index_v4.html")
print("V4.5 backups preserved.")
print("Portable checkpoint support added.")
