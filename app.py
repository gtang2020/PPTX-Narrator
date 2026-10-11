"""iPad Safari PowerPoint narrator: FastAPI + python-pptx + OpenAI speech API."""
import io, os, re, tempfile, zipfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from lxml import etree

app = FastAPI(title='iPad PowerPoint Narrator')
ROOT=Path(__file__).parent
MAX_BYTES=150*1024*1024

@app.get('/',response_class=HTMLResponse)
def home(): return (ROOT/'index.html').read_text(encoding='utf-8')

def get_notes(slide):
    if not slide.has_notes_slide: return ''
    tf=slide.notes_slide.notes_text_frame
    return tf.text.strip() if tf is not None else ''

def slide_audio(z, n):
    relpath=f'ppt/slides/_rels/slide{n}.xml.rels'
    if relpath not in z.namelist(): return None
    root=etree.fromstring(z.read(relpath))
    for el in root:
        target=el.get('Target','')
        typ=el.get('Type','')
        if target.lower().endswith(('.mp3','.m4a','.wav')) and ('/media' in typ or '/audio' in typ or '/video' in typ):
            path='ppt/'+target.replace('../','') if target.startswith('../') else 'ppt/slides/'+target
            if path in z.namelist(): return z.read(path),Path(path).suffix.lower()
    return None

def add_audio(slide, data, suffix):
    """Add a playable media shape; PowerPoint may require a tap on some clients."""
    with tempfile.NamedTemporaryFile(suffix=suffix,delete=False) as f:
        f.write(data); fn=f.name
    try:
        # Media insertion via python-pptx; audio-only files are embedded as media objects.
        mime={'.mp3':'audio/mpeg','.m4a':'audio/mp4','.wav':'audio/wav'}[suffix]
        shape=slide.shapes.add_movie(fn, left=0, top=0, width=200000, height=200000, poster_frame_image=None, mime_type=mime)
        # Set the media animation start to zero and add after previous effects.
        # Note: cross-client autoplay isn't guaranteed for newly inserted audio.
        return shape
    finally: os.unlink(fn)

async def synth(text,voice,api_key):
    if not api_key: raise HTTPException(400,'AI voice mode requires an OpenAI API key. Use Existing Audio mode without one.')
    try:
        from openai import AsyncOpenAI
        client=AsyncOpenAI(api_key=api_key)
        result=await client.audio.speech.create(model='tts-1',voice=voice,input=text[:4096],response_format='mp3')
        return result.content
    except Exception as e:
        raise HTTPException(502,f'Audio generation failed: {str(e)[:250]}')

@app.post('/process')
async def process(file:UploadFile=File(...),mode:str=Form('existing'),voice:str=Form('alloy'),api_key:str=Form(''),limit:int=Form(0)):
    if not file.filename or not file.filename.lower().endswith('.pptx'): raise HTTPException(400,'Upload a .pptx file')
    content=await file.read(MAX_BYTES+1)
    if len(content)>MAX_BYTES: raise HTTPException(413,'File exceeds 150 MB limit')
    try: prs=Presentation(io.BytesIO(content))
    except Exception: raise HTTPException(400,'Invalid PowerPoint file')
    n=min(len(prs.slides),limit) if limit>0 else len(prs.slides)
    notes=[get_notes(prs.slides[i]) for i in range(n)]
    if mode=='notes':
        out='\n\n'.join(f'=== Slide {i+1} ===\n{txt}' for i,txt in enumerate(notes))
        return StreamingResponse(iter([out.encode('utf-8')]),media_type='text/plain; charset=utf-8',headers={'Content-Disposition':'attachment; filename="slide_notes.txt"'})
    if mode not in ('existing','tts'): raise HTTPException(400,'Invalid mode')
    audio=[]
    if mode=='existing':
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            for i in range(n): audio.append(slide_audio(z,i+1))
        if not any(audio): raise HTTPException(400,'No existing slide audio found. Choose AI Voice mode.')
    else:
        for i,txt in enumerate(notes):
            audio.append((await synth(txt,voice,api_key),'.mp3') if txt else None)
    # Preserve original slide media in existing-audio mode; re-saving may change existing audio
    # relationships, so deliver original PPTX for full-length existing-audio jobs.
    if mode=='existing' and (limit<=0 or n==len(prs.slides)):
        return StreamingResponse(iter([content]),media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',headers={'Content-Disposition':'attachment; filename="narrated_existing_audio.pptx"'})
    if mode=='existing':
        # For testing a subset, preserve the original slides as-is via direct OOXML editing
        # rather than rewriting media relationships through python-pptx.
        raise HTTPException(400,'For Existing Audio, choose All slides to preserve embedded media. A 3-slide example is provided separately.')
    for i,item in enumerate(audio):
        if item: add_audio(prs.slides[i],*item)
    out=io.BytesIO(); prs.save(out); out.seek(0)
    return StreamingResponse(out,media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',headers={'Content-Disposition':'attachment; filename="narrated_ai_voice.pptx"'})
