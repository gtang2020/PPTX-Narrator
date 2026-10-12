"""PPTX Narrator V4.5 -> V4.6. Requires original backed-up source files.

Run in the repository root: python upgrade_v46.py
Fails without modifying live files if the backup structure is unexpected.
"""
from pathlib import Path
import ast

root = Path(__file__).resolve().parent
backend = root / 'app_v4_5.py'
frontend = root / 'index_v4_5.html'
if not backend.is_file() or not frontend.is_file():
    raise SystemExit('Missing app_v4_5.py or index_v4_5.html backup in repository root.')
py = backend.read_text(encoding='utf-8')
html = frontend.read_text(encoding='utf-8')

def replace_once(text, old, new, label):
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f'{label}: expected one matching section, found {count}. No files changed.')
    return text.replace(old, new, 1)

py = replace_once(py, 'PPTX Narrator V4.5', 'PPTX Narrator V4.6', 'backend docstring')
py = replace_once(py, 'title="PPTX Narrator V4.5"', 'title="PPTX Narrator V4.6"', 'app title')
py = replace_once(py, '"version": "4.5",', '"version": "4.6",', 'health version')
py = replace_once(py, '"segment_resume": True,', '"segment_resume": True,\n        "portable_checkpoint": True,\n        "resume_from_uploaded_pptx": True,', 'health flags')

portable = '''# ============================================================
# V4.6 PORTABLE CHECKPOINT METADATA
# ============================================================

PORTABLE_PROGRESS = "pptx_narrator_v46_progress.json"


def notes_digest(notes):
    return hashlib.sha256(notes.encode("utf-8")).hexdigest()


def portable_metadata(prs, completed, skipped):
    """A checkpoint is valid only for the exact slide notes it narrates."""
    return {
        "format": "PPTX_NARRATOR_PORTABLE_CHECKPOINT",
        "version": "4.6",
        "completed": int(completed),
        "total": len(prs.slides),
        "skipped_empty_notes": int(skipped),
        "notes_sha256": [notes_digest(get_notes(prs.slides[i]))
                         for i in range(completed)],
        "saved_at": int(time.time()),
    }


def attach_portable_metadata(path, metadata):
    """Write a portable manifest into the PPTX without changing its slides."""
    temporary = path.with_name(path.name + ".portable.tmp")
    try:
        with zipfile.ZipFile(path, "r") as original, zipfile.ZipFile(
            temporary, "w"
        ) as updated:
            for member in original.infolist():
                if member.filename != PORTABLE_PROGRESS:
                    updated.writestr(member, original.read(member.filename))
            updated.writestr(PORTABLE_PROGRESS,
                             json.dumps(metadata, ensure_ascii=False))
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def read_portable_metadata(path, prs):
    """Reject forged, stale, or incompatible recovery progress."""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            if PORTABLE_PROGRESS not in archive.namelist():
                return None
            raw = archive.read(PORTABLE_PROGRESS)
            if len(raw) > 1024 * 1024:
                raise ValueError("Checkpoint metadata too large")
            meta = json.loads(raw)
    except (zipfile.BadZipFile, json.JSONDecodeError, KeyError) as exc:
        raise ValueError("Invalid checkpoint metadata") from exc
    if meta is None:
        return None
    if not isinstance(meta, dict) or meta.get("format") != \
            "PPTX_NARRATOR_PORTABLE_CHECKPOINT":
        raise ValueError("Unrecognized checkpoint metadata")
    total = len(prs.slides)
    completed = meta.get("completed")
    hashes = meta.get("notes_sha256")
    skipped = meta.get("skipped_empty_notes")
    if (type(completed) is not int or not 0 <= completed <= total or
        type(skipped) is not int or not 0 <= skipped <= completed or
        meta.get("total") != total or
        not isinstance(hashes, list) or len(hashes) != completed):
        raise ValueError("Invalid checkpoint slide count or progress")
    actual_skipped = 0
    for i in range(completed):
        slide = prs.slides[i]
        notes = get_notes(slide)
        if hashes[i] != notes_digest(notes):
            raise ValueError(f"Checkpoint notes changed on slide {i + 1}")
        if notes:
            # Every completed narrated slide must retain embedded audio.
            if not any(shape.shape_type == MSO_SHAPE_TYPE.MEDIA
                       for shape in slide.shapes):
                raise ValueError(f"Checkpoint audio missing on slide {i + 1}")
        else:
            actual_skipped += 1
    if actual_skipped != skipped:
        raise ValueError("Checkpoint skipped-slide count does not match")
    return meta


'''
py = replace_once(py, '# ============================================================\n# CHECKPOINT VALIDATION\n# ============================================================', portable + '# ============================================================\n# CHECKPOINT VALIDATION\n# ============================================================', 'portable helpers')

py = replace_once(py, '''    validate_pptx(
        temporary,
        total,
    )

    os.replace(
        temporary,
        checkpoint_path(job_dir),
    )''', '''    validate_pptx(
        temporary,
        total,
    )

    # Store resume metadata inside the actual downloaded PPTX.
    attach_portable_metadata(
        temporary,
        portable_metadata(prs, completed, skipped),
    )
    validate_pptx(temporary, total)

    os.replace(
        temporary,
        checkpoint_path(job_dir),
    )''', 'save checkpoint')

# In the event checkpoint.json was lost but PPTX survives, read the portable metadata.
py = replace_once(py, '''    if (
        not checkpoint.is_file()
        or not metadata.is_file()
    ):
        return status
''', '''    if not checkpoint.is_file():
        return status

    if not metadata.is_file():
        try:
            prs = Presentation(str(checkpoint))
            portable = read_portable_metadata(checkpoint, prs)
            if portable and portable["total"] == status["total"]:
                status["completed"] = portable["completed"]
                status["skipped_empty_notes"] = portable["skipped_empty_notes"]
                status["checkpoint_available"] = True
        except Exception:
            pass
        return status
''', 'checkpoint restore fallback')

# Do not permit a new narration script when continuing an existing checkpoint.
py = replace_once(py, '''        target = (
            min(limit, total)
            if limit
            else total
        )

        imported_count = 0''', '''        target = (
            min(limit, total)
            if limit
            else total
        )

        try:
            portable = read_portable_metadata(source, prs)
        except ValueError as exc:
            raise HTTPException(400, f"Invalid recovery checkpoint: {exc}") from exc

        recovered = portable["completed"] if portable else 0
        recovered_skipped = portable["skipped_empty_notes"] if portable else 0
        if recovered > target:
            raise HTTPException(400, "Slides to Process must be 0 or at least the saved slide count")
        if portable and notes_file is not None and notes_file.filename:
            raise HTTPException(400, "Do not upload a new script when resuming a checkpoint; notes are already embedded")

        imported_count = 0''', 'upload checkpoint detection')

py = replace_once(py, '''            "completed": 0,
            "current_slide": 0,''', '''            "completed": recovered,
            "current_slide": recovered,
            "resumed_from": recovered,
            "portable_checkpoint": bool(portable),''', 'initial recovered progress')
py = replace_once(py, '''            "skipped_empty_notes": 0,
            "checkpoint_available": False,''', '''            "skipped_empty_notes": recovered_skipped,
            "checkpoint_available": bool(portable),''', 'initial recovered skipped')

py = replace_once(py, '''        update_status(
            job_dir,
            status,
        )

        launch_job(
            job_id
        )''', '''        # A downloaded checkpoint becomes this job's initial checkpoint.
        # Preserve it before background processing starts.
        if portable:
            shutil.copy2(source, checkpoint_path(job_dir))
            save_json(checkpoint_meta_path(job_dir), {
                "completed": recovered,
                "skipped_empty_notes": recovered_skipped,
                "total": total,
                "saved_at": int(time.time()),
            })

        update_status(
            job_dir,
            status,
        )

        launch_job(
            job_id
        )''', 'initial checkpoint copy')

# Fix a corner case: already complete uploaded checkpoint should not be re-saved,
# but the job should return completed with the original checkpoint available.
py = replace_once(py, 'f"narrated_v4_5_"', 'f"narrated_v4_6_"', 'pptx download name')
py = replace_once(py, 'f"narration_audio_v4_5_"', 'f"narration_audio_v4_6_"', 'audio download name')
py = replace_once(py, '"PPTX_with_Speaker_Notes_V4_5.pptx"', '"PPTX_with_Speaker_Notes_V4_6.pptx"', 'step1 name')
py = replace_once(py, '"narrated_piper_v4_5.pptx"', '"narrated_piper_v4_6.pptx"', 'legacy name')

# Change checkpoint batch boundaries relative to recovery point to save after N *new* slides.
py = replace_once(py, '''                (i + 1) % batch_size == 0
                or i + 1 == target''', '''                (i + 1 - int(status.get("resumed_from", 0))) % batch_size == 0
                or i + 1 == target''', 'batch relative resume')

# Frontend upgrades. Keep original functionality and endpoints intact.
html = replace_once(html, 'PPTX Narrator V4.5', 'PPTX Narrator V4.6', 'HTML title')
html = replace_once(html, '🎙 PPTX Narrator V4.5', '🎙 PPTX Narrator V4.6', 'HTML heading')
html = replace_once(html, 'V4.5 includes longer Piper synthesis timeouts and', 'V4.6 includes portable recovery checkpoints, longer Piper synthesis timeouts and', 'warning description')
html = replace_once(html, 'Render free storage is temporary.\nDownload completed checkpoints and audio ZIP files promptly.', 'Render free storage is temporary.\n<strong>Download your recovery PPTX to your iPad after each checkpoint.</strong>\nServer-side checkpoints can disappear after restarts.', 'warning download')
html = replace_once(html, '''<p class="small">
<strong>3</strong> = first three slides.''', '''<p class="small">
<strong>For a recovery PPTX, choose 0 to finish all remaining slides.</strong><br>
<strong>3</strong> = first three slides.''', 'limit instructions')
html = replace_once(html, '''<p class="small">
The PPTX download becomes available after a checkpoint
has been saved.
</p>''', '''<p class="small success">
⬇️ Download the latest narrated PPTX to your iPad whenever
it becomes available. After a Render restart, upload that
file in Step 2 and set Slides to Process = 0.
</p>''', 'download instructions')
html = replace_once(html, 'const STORAGE_KEY = "pptx_v45_job";', 'const STORAGE_KEY = "pptx_v46_job";', 'local storage key')
html = replace_once(html, 'const OLD_STORAGE_KEY = "pptx_v44_job";', 'const OLD_STORAGE_KEY = "pptx_v45_job";', 'old local storage key')
html = replace_once(html, '''        `Saved slides: ${completed} / ${target}\\n` +''', '''        `Saved slides: ${completed} / ${target}\\n` +
        `Recovered from upload: ${status.resumed_from ?? 0} slides\\n` +
        `Next slide after recovery: ${Math.min(target + 1, completed + 1)}\\n` +''', 'recovery status text')
html = replace_once(html, '"PPTX_with_Speaker_Notes_V4_5.pptx"', '"PPTX_with_Speaker_Notes_V4_6.pptx"', 'frontend step1 filename')
html = replace_once(html, 'Checking V4.5 server configuration…', 'Checking V4.6 server configuration…', 'frontend health placeholder')
html = replace_once(html, 'if (version !== "4.5")', 'if (version !== "4.6")', 'frontend version comparison')
html = replace_once(html, '"\\n\\n⚠️ Expected backend V4.5."', '"\\n\\n⚠️ Expected backend V4.6."', 'frontend version warning')

# Check generated code parses and only then write new files.
ast.parse(py)
if '<script>' not in html or '</html>' not in html:
    raise RuntimeError('Invalid generated HTML')
(root / 'app.py').write_text(py, encoding='utf-8')
(root / 'index_v4.html').write_text(html, encoding='utf-8')
print('SUCCESS: Generated app.py and index_v4.html V4.6 from V4.5 backups.')
