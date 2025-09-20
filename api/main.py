# api/main.py
from __future__ import annotations

import os
import glob
import shutil
import zipfile
import tempfile
import subprocess
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, Form, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

# ----------- CORS (set ALLOW_ORIGINS env var for production) -----------
_allow_origins = os.getenv("ALLOW_ORIGINS", "*").split(",")
app = FastAPI(title="Henbrook Menus API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in _allow_origins],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ----------- Simple upload form (handy for quick manual tests) -----------
@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <html><body style="font-family: system-ui; max-width: 720px; margin: 40px auto;">
      <h2>Henbrook Daily Menu Generator (API)</h2>
      <form action="/generate" method="post" enctype="multipart/form-data">
        <p><label>Weekly DOCX: <input type="file" name="weekly" required /></label></p>
        <fieldset style="border:1px solid #ddd;padding:10px;">
          <legend>Templates (optional – upload ALL three OR use server defaults)</legend>
          <p><label>Standard:  <input type="file" name="standard_tpl" /></label></p>
          <p><label>Vegan:     <input type="file" name="vegan_tpl" /></label></p>
          <p><label>Allergens: <input type="file" name="allergens_tpl" /></label></p>
        </fieldset>
        <p><label>Date (YYYY-MM-DD): <input name="date" placeholder="2025-09-19" /></label></p>
        <p><label><input type="checkbox" name="all_days" value="true" /> Generate all 7 days</label></p>
        <p><button type="submit">Generate</button></p>
      </form>
      <p style="color:#666">If no templates uploaded, the API will use <code>./templates</code> in the repo.</p>
    </body></html>
    """

# ----------- Helpers -----------
def _save_upload(tmpdir: Path, up: UploadFile, name: Optional[str] = None) -> Path:
    """Save an UploadFile to tmpdir, return its path."""
    name = name or up.filename or "upload.bin"
    dest = tmpdir / name
    with dest.open("wb") as f:
        shutil.copyfileobj(up.file, f)
    return dest

def _bundle_many_zips(zips: list[Path], out_path: Path) -> Path:
    """Bundle multiple generated day ZIPs into one ZIP for single download."""
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for zp in zips:
            z.write(zp, arcname=zp.name)
    return out_path

def _to_bool(v: Optional[str | bool]) -> bool:
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    return str(v).lower() in {"1", "true", "yes", "on"}

# ----------- Main endpoint -----------
@app.post("/generate")
async def generate(
    background: BackgroundTasks,
    weekly: UploadFile,
    date: Optional[str] = Form(default=None),
    all_days: Optional[str] = Form(default=None),  # "true"/"false"
    standard_tpl: UploadFile | None = None,
    vegan_tpl: UploadFile | None = None,
    allergens_tpl: UploadFile | None = None,
):
    """
    Accepts:
    - weekly: DOCX weekly grid (required)
    - date: YYYY-MM-DD (optional; omit if generating all 7 days)
    - all_days: "true" to generate the entire week (optional)
    - standard_tpl / vegan_tpl / allergens_tpl: optional custom templates (all three required if any provided)

    Returns:
    - A ZIP file (application/zip). If multiple day ZIPs were produced, they are bundled into one archive.
    """
    tmpdir = Path(tempfile.mkdtemp(prefix="menus_"))

    def _cleanup():
        # Remove temp directory after response has been sent
        shutil.rmtree(tmpdir, ignore_errors=True)

    try:
        # Save weekly file
        weekly_path = _save_upload(tmpdir, weekly, "weekly.docx")

        # Build command
        cmd = ["python", "generate_menus.py", "--weekly", str(weekly_path), "--out", str(tmpdir)]
        if date:
            cmd += ["--date", date]
        if _to_bool(all_days):
            cmd.append("--all-days")

        # Templates: either all three uploaded OR use repo defaults ./templates
        if standard_tpl or vegan_tpl or allergens_tpl:
            if not (standard_tpl and vegan_tpl and allergens_tpl):
                return JSONResponse(
                    {"detail": "If providing custom templates, upload all three: standard, vegan, allergens."},
                    status_code=400,
                )
            std_path = _save_upload(tmpdir, standard_tpl, "standard.docx")
            veg_path = _save_upload(tmpdir, vegan_tpl, "vegan.docx")
            alg_path = _save_upload(tmpdir, allergens_tpl, "allergens.docx")
            cmd += ["--standard_tpl", str(std_path), "--vegan_tpl", str(veg_path), "--allergens_tpl", str(alg_path)]
        else:
            # Use repo templates folder relative to project root
            cmd += ["--templates", "./templates"]

        # Run the generator script
        completed = subprocess.run(cmd, capture_output=True, text=True)
        if completed.returncode != 0:
            # include stderr to help debugging
            return JSONResponse(
                {"detail": "Generation failed", "stderr": completed.stderr, "stdout": completed.stdout},
                status_code=500,
            )

        # Collect produced zips
        produced = [Path(p) for p in glob.glob(str(tmpdir / "*.zip"))]
        if not produced:
            return JSONResponse({"detail": "No ZIP produced. Check the weekly file and templates."}, status_code=400)

        # If multiple, bundle them for convenience
        if len(produced) > 1:
            bundle = tmpdir / "Henbrook-all-days.zip"
            _bundle_many_zips(produced, bundle)
            background.add_task(_cleanup)
            return FileResponse(bundle, filename=bundle.name, media_type="application/zip")

        # Single day — return as-is
        zip_path = produced[0]
        background.add_task(_cleanup)
        return FileResponse(zip_path, filename=zip_path.name, media_type="application/zip")

    except Exception as e:
        background.add_task(_cleanup)
        return JSONResponse({"detail": f"Server error: {e.__class__.__name__}: {e}"}, status_code=500)
