"""Tiny standalone web app: upload a JDE 'A/R Details with Aging' PDF, get back
an Excel workbook in the template1.xlsx layout. No login / database required.

Run:
    .venv\\Scripts\\python.exe ar_app.py
then open  http://127.0.0.1:8001  and drop in the PDF.
"""
from __future__ import annotations

import io
import os
import tempfile

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, StreamingResponse

from ar_to_excel import export, parse

ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(ROOT, "template1.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="A/R PDF → Excel")

PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>A/R Aging · PDF → Excel</title>
<style>
  :root{color-scheme:dark}
  body{font-family:-apple-system,'Segoe UI',Inter,sans-serif;background:#0d0e16;
       color:#f1eee7;display:grid;place-items:center;min-height:100vh;margin:0}
  .card{background:#171823;border:1px solid #2a2d3b;border-radius:16px;
        padding:36px 40px;max-width:520px;width:90%}
  h1{font-family:Georgia,serif;color:#c9a15a;margin:0 0 6px;font-size:26px}
  p.sub{color:#9a9485;margin:0 0 24px;font-style:italic}
  .drop{border:2px dashed #3a3d4e;border-radius:12px;padding:34px;text-align:center;
        cursor:pointer;transition:.15s;background:#0f1018}
  .drop.hover{border-color:#c9a15a;background:#14151f}
  .drop b{color:#f1eee7}
  input[type=file]{display:none}
  button{margin-top:18px;width:100%;padding:12px;border:0;border-radius:10px;
         background:#c9a15a;color:#1a1206;font-weight:700;font-size:15px;cursor:pointer}
  button:disabled{opacity:.5;cursor:default}
  .msg{margin-top:16px;min-height:20px;font-size:14px}
  .ok{color:#5fdfb0}.bad{color:#e9806e}.work{color:#e0bd72}
  .file{margin-top:14px;font-size:13px;color:#b9b3a6}
</style></head><body>
<div class="card">
  <h1>A/R Aging → Excel</h1>
  <p class="sub">Drop the JDE "A/R Details with Aging" PDF. Get the spreadsheet.</p>
  <label class="drop" id="drop">
    <div><b>Click to choose</b> or drop a PDF here</div>
    <input type="file" id="file" accept="application/pdf,.pdf">
    <div class="file" id="fname"></div>
  </label>
  <button id="go" disabled>Convert to Excel</button>
  <div class="msg" id="msg"></div>
</div>
<script>
const drop=document.getElementById('drop'),file=document.getElementById('file'),
      go=document.getElementById('go'),msg=document.getElementById('msg'),
      fname=document.getElementById('fname');
function setFile(f){
  if(!f) return;
  if(!/\\.pdf$/i.test(f.name)){msg.className='msg bad';msg.textContent='Please pick a .pdf file.';return;}
  file._f=f; fname.textContent=f.name+'  ('+(f.size/1024/1024).toFixed(1)+' MB)';
  go.disabled=false; msg.textContent='';
}
drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('hover');});
drop.addEventListener('dragleave',()=>drop.classList.remove('hover'));
drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('hover');setFile(e.dataTransfer.files[0]);});
file.addEventListener('change',()=>setFile(file.files[0]));
go.addEventListener('click',async()=>{
  const f=file._f; if(!f) return;
  go.disabled=true; msg.className='msg work'; msg.textContent='Crunching the report… this can take a few seconds.';
  const fd=new FormData(); fd.append('file',f);
  try{
    const r=await fetch('/convert',{method:'POST',body:fd});
    if(!r.ok){const j=await r.json().catch(()=>({}));throw new Error(j.detail||'Conversion failed.');}
    const blob=await r.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; a.download=f.name.replace(/\\.pdf$/i,'')+'.xlsx'; a.click();
    URL.revokeObjectURL(url);
    msg.className='msg ok'; msg.textContent='Done — your .xlsx downloaded.';
  }catch(e){msg.className='msg bad'; msg.textContent=e.message;}
  go.disabled=false;
});
</script></body></html>"""


@app.get("/")
def index() -> HTMLResponse:
    return HTMLResponse(PAGE)


@app.get("/health")
def health() -> dict:
    return {"ok": True}


def _convert(pdf_bytes: bytes) -> io.BytesIO:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    try:
        data = parse(tmp)
        buf = io.BytesIO()
        export(data, buf, template_path=TEMPLATE, totals=False)
        buf.seek(0)
        return buf
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a .pdf file.")
    pdf_bytes = await file.read()
    if pdf_bytes[:5] != b"%PDF-":
        raise HTTPException(status_code=400, detail="That doesn't look like a PDF.")
    try:
        buf = await run_in_threadpool(_convert, pdf_bytes)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}") from e
    out_name = os.path.splitext(os.path.basename(file.filename))[0] + ".xlsx"
    return StreamingResponse(
        buf, media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


if __name__ == "__main__":
    print("A/R PDF → Excel  ·  open http://127.0.0.1:8001")
    uvicorn.run(app, host="127.0.0.1", port=8001)
