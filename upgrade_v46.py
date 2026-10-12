"""
PPTX Narrator V4.5 -> V4.6 Upgrade

Portable checkpoint recovery.

This upgrade:
1. Preserves the V4.5 backup files.
2. Generates a V4.6 backend.
3. Generates a V4.6 frontend.
4. Adds portable PPTX checkpoint metadata.
5. Supports resuming from uploaded checkpoints.
6. Checks Python syntax before replacing live files.

Run:
    python upgrade_v46.py

Required backups:
    app_v4_5.py
    index_v4_5.html

Generated:
    app.py
    index_v4.html
"""

from pathlib import Path
import ast
import os
import tempfile


# ============================================================
# FILE CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent

BACKEND_BACKUP = ROOT / "app_v4_5.py"
FRONTEND_BACKUP = ROOT / "index_v4_5.html"

BACKEND_OUTPUT = ROOT / "app.py"
FRONTEND_OUTPUT = ROOT / "index_v4.html"


# ============================================================
# READ ORIGINAL BACKUPS
# ============================================================

if not BACKEND_BACKUP.is_file():
    raise SystemExit(
        "ERROR: app_v4_5.py backup is missing."
    )

if not FRONTEND_BACKUP.is_file():
    raise SystemExit(
        "ERROR: index_v4_5.html backup is missing."
    )

py = BACKEND_BACKUP.read_text(
    encoding="utf-8"
)

html = FRONTEND_BACKUP.read_text(
    encoding="utf-8"
)


# ============================================================
# SAFE REPLACEMENT
# ============================================================

def replace_once(text, old, new, label):
    count = text.count(old)

    if count != 1:
        raise RuntimeError(
            f"{label}: expected exactly one match, "
            f"found {count}. No live files changed."
        )

    return text.replace(
        old,
        new,
        1,
    )


# ============================================================
# BACKEND VERSION
# ============================================================

py = replace_once(
    py,
    "PPTX Narrator V4.5",
    "PPTX Narrator V4.6",
    "Backend documentation",
)

py = replace_once(
    py,
    'title="PPTX Narrator V4.5"',
    'title="PPTX Narrator V4.6"',
    "Application title",
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
# PORTABLE CHECKPOINT SYSTEM
# ============================================================

PORTABLE_CODE = '''
# ============================================================
# V4.6 PORTABLE CHECKPOINT METADATA
# ============================================================

PORTABLE_PROGRESS = (
    "pptx_narrator_v46_progress.json"
)


def notes_digest(notes):
    return hashlib.sha256(
        notes.encode("utf-8")
    ).hexdigest()


def portable_metadata(
    prs,
    completed,
    skipped,
):
    return {
        "format": (
            "PPTX_NARRATOR_PORTABLE_CHECKPOINT"
        ),
        "version": "4.6",
        "completed": int(completed),
        "total": len(prs.slides),
        "skipped_empty_notes": int(skipped),
        "notes_sha256": [
            notes_digest(
                get_notes(prs.slides[i])
            )
            for i in range(completed)
        ],
        "saved_at": int(time.time()),
    }


def attach_portable_metadata(
    path,
    metadata,
):
    """
    Embed progress information inside PPTX.

    The progress travels with the downloaded
    PowerPoint file.
    """

    temporary = path.with_name(
        path.name + ".portable.tmp"
    )

    try:
        with zipfile.ZipFile(
            path,
            "r",
        ) as original:

            with zipfile.ZipFile(
                temporary,
                "w",
            ) as updated:

                for member in original.infolist():

                    if (
                        member.filename
                        == PORTABLE_PROGRESS
                    ):
                        continue

                    updated.writestr(
                        member,
                        original.read(
                            member.filename
                        ),
                    )

                updated.writestr(
                    PORTABLE_PROGRESS,
                    json.dumps(
                        metadata,
                        ensure_ascii=False,
                    ),
                )

        os.replace(
            temporary,
            path,
        )

    finally:
        temporary.unlink(
            missing_ok=True
        )


def read_portable_metadata(
    path,
    prs,
):
    """
    Validate uploaded recovery checkpoints.

    A completed narrated slide must still
    contain its embedded audio.
    """

    try:
        with zipfile.ZipFile(
            path,
            "r",
        ) as archive:

            if (
                PORTABLE_PROGRESS
                not in archive.namelist()
            ):
                return None

            raw = archive.read(
                PORTABLE_PROGRESS
            )

            if len(raw) > 1024 * 1024:
                raise ValueError(
                    "Checkpoint metadata too large"
                )

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

    total = len(prs.slides)

    completed = meta.get(
        "completed"
    )

    skipped = meta.get(
        "skipped_empty_notes"
    )

    hashes = meta.get(
        "notes_sha256"
    )

    if (
        type(completed) is not int
        or not 0 <= completed <= total
        or type(skipped) is not int
        or not 0 <= skipped <= completed
        or meta.get("total") != total
        or not isinstance(hashes, list)
        or len(hashes) != completed
    ):
        raise ValueError(
            "Invalid checkpoint progress"
        )

    actual_skipped = 0

    for i in range(completed):

        slide = prs.slides[i]

        notes = get_notes(
            slide
        )

        if hashes[i] != notes_digest(
            notes
        ):
            raise ValueError(
                f"Checkpoint notes changed "
                f"on slide {i + 1}"
            )

        if notes:

            has_audio = any(
                shape.shape_type
                == MSO_SHAPE_TYPE.MEDIA
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
# INSERT PORTABLE HELPERS
# ============================================================

CHECKPOINT_MARKER = """# ============================================================
# CHECKPOINT VALIDATION
# ============================================================"""

py = replace_once(
    py,
    CHECKPOINT_MARKER,
    PORTABLE_CODE + CHECKPOINT_MARKER,
    "Portable checkpoint helpers",
)


# ============================================================
# EMBED PROGRESS WHEN SAVING
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
    "Checkpoint saving",
)


# ============================================================
# RESTORE PROGRESS FROM PORTABLE FILE
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
            prs = Presentation(
                str(checkpoint)
            )

            portable = read_portable_metadata(
                checkpoint,
                prs,
            )

            if (
                portable
                and portable["total"]
                == status["total"]
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
    "Checkpoint recovery",
)


# ============================================================
# DETECT UPLOADED CHECKPOINT
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
                "Invalid recovery checkpoint: "
                + str(exc),
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
    "Recovered progress",
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
# COPY UPLOADED CHECKPOINT INTO NEW JOB
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
                    "skipped_empty_notes": (
                        recovered_skipped
                    ),
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
    "Initial checkpoint preservation",
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
                    - int(
                        status.get(
                            "resumed_from",
                            0,
                        )
                    )
                ) % batch_size == 0
                or i + 1 == target''',
    "Checkpoint interval",
)


# ============================================================
# BACKEND DOWNLOAD FILENAMES
# ============================================================

py = replace_once(
    py,
    'f"narrated_v4_5_"',
    'f"narrated_v4_6_"',
    "PPTX filename",
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
    "Notes filename",
)

py = replace_once(
    py,
    '"narrated_piper_v4_5.pptx"',
    '"narrated_piper_v4_6.pptx"',
    "Legacy filename",
)


# ============================================================
# FRONTEND VERSION
# ============================================================

# Replace both version labels together.
# This avoids the previous duplicate-replacement bug.

if html.count("PPTX Narrator V4.5") != 2:
    raise RuntimeError(
        "Unexpected HTML version labels. "
        "No live files changed."
    )

html = html.replace(
    "PPTX Narrator V4.5",
    "PPTX Narrator V4.6",
)


# ============================================================
# FRONTEND CHECKPOINT INSTRUCTIONS
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
    "Frontend storage warning",
)

html = replace_once(
    html,
    "<strong>3</strong> = first three slides.",
    """<strong>Recovery PPTX: choose 0 to finish remaining slides.</strong>
<br>
<strong>3</strong> = first three slides.""",
    "Frontend slide limit instructions",
)

html = replace_once(
    html,
    """The PPTX download becomes available after a checkpoint
has been saved.""",
    """⬇️ Download the latest narrated PPTX to your iPad whenever
a new checkpoint becomes available.

<br><br>

After a Render restart, upload that saved PPTX in Step 2
and set Slides to Process = 0.""",
    "Frontend download instructions",
)


# ============================================================
# FRONTEND LOCAL STORAGE
# ============================================================

html = replace_once(
    html,
    'const STORAGE_KEY = "pptx_v45_job";',
    'const STORAGE_KEY = "pptx_v46_job";',
    "Frontend storage key",
)

html = replace_once(
    html,
    'const OLD_STORAGE_KEY = "pptx_v44_job";',
    'const OLD_STORAGE_KEY = "pptx_v45_job";',
    "Frontend previous storage key",
)


# ============================================================
# FRONTEND RECOVERY PROGRESS
# ============================================================

html = replace_once(
    html,
    '''        `Saved slides: ${completed} / ${target}\\n` +''',
    '''        `Saved slides: ${completed} / ${target}\\n` +
        `Recovered from upload: ${status.resumed_from ?? 0} slides\\n` +''',
    "Frontend recovery progress",
)


# ============================================================
# FRONTEND FILE NAMES AND HEALTH
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
    "Frontend health placeholder",
)

html = replace_once(
    html,
    'if (version !== "4.5")',
    'if (version !== "4.6")',
    "Frontend version comparison",
)

html = replace_once(
    html,
    "Expected backend V4.5.",
    "Expected backend V4.6.",
    "Frontend version warning",
)


# ============================================================
# VALIDATE GENERATED FILES
# ============================================================

ast.parse(py)

if "<script>" not in html:
    raise RuntimeError(
        "Generated HTML has no script"
    )

if "</html>" not in html:
    raise RuntimeError(
        "Generated HTML is incomplete"
    )

if "PPTX Narrator V4.6" not in html:
    raise RuntimeError(
        "Frontend version upgrade failed"
    )

if '"version": "4.6"' not in py:
    raise RuntimeError(
        "Backend version upgrade failed"
    )


# ============================================================
# WRITE FILES
# ============================================================

# Stage both outputs before changing live files.

def stage_text(destination, content):
    fd, temporary_name = tempfile.mkstemp(
        prefix=destination.name + ".",
        suffix=".tmp",
        dir=ROOT,
    )

    temporary = Path(temporary_name)

    try:
        with os.fdopen(
            fd,
            "w",
            encoding="utf-8",
        ) as output:
            output.write(content)

        return temporary

    except Exception:
        temporary.unlink(
            missing_ok=True
        )
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

    # Both generated files have passed basic validation.
    # The V4.5 backup files remain untouched.

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
        staged_backend.unlink(
            missing_ok=True
        )

    if staged_frontend is not None:
        staged_frontend.unlink(
            missing_ok=True
        )


print(
    "SUCCESS: PPTX Narrator V4.6 files generated."
)

print(
    "Backend: app.py"
)

print(
    "Frontend: index_v4.html"
)

print(
    "Original V4.5 backups preserved."
)

print(
    "Portable checkpoint recovery enabled."
)
