# Henbrook Daily Menu Generator

A single Python script that:

- Parses the **weekly menu DOCX** (your grid).
- Renders **Standard**, **Vegan**, and **Allergens** DOCXs for each day using your templates.
- Writes **only a ZIP per day** containing those three DOCXs:
  - `Residents_DD-MM-YYYY.docx`
  - `Residents_DD-MM-YYYY_vegan.docx`
  - `Allergens_Residents_DD-MM-YYYY.docx`

---

## Folder layout

```
project/
├─ generate_menus.py
├─ templates/
│  ├─ standard.docx
│  ├─ vegan.docx
│  └─ allergens.docx
└─ build/                 # ZIPs are written here (or to --out)
```

> You can keep templates anywhere; the default is `./templates`. You can also point to individual template files with flags.

---

## Install

```powershell
# (optional) create a venv
python -m venv .venv
. .venv/Scripts/Activate.ps1

# install deps
pip install -r requirements.txt
```

`requirements.txt` should include:
```
python-docx
docxtpl
PyYAML
```

---

## Usage

### PowerShell (Windows)

#### Generate **one day** (using the default `templates/` folder)
```powershell
python .\generate_menus.py --weekly ".\Residents Menu WC 15-09-2025.docx" --date 2025-09-18 --templates ".\templates" --out ".\build"
```

#### Generate **all 7 days**
```powershell
python .\generate_menus.py --weekly ".\Residents Menu WC 15-09-2025.docx" --all-days --templates ".\templates" --out ".\build"
```

#### If your templates are not in `templates/` (explicit paths)
```powershell
python .\generate_menus.py `
  --weekly ".\Residents Menu WC 15-09-2025.docx" `
  --standard_tpl ".\templates\standard.docx" `
  --vegan_tpl ".\templates\vegan.docx" `
  --allergens_tpl ".\templates\allergens.docx" `
  --date 2025-09-18 `
  --out ".\build"
```
> PowerShell line-continuation uses the **backtick** (`) — not a backslash.

### CMD (Windows)

```bat
python generate_menus.py ^
  --weekly "Residents Menu WC 15-09-2025.docx" ^
  --all-days ^
  --templates "templates" ^
  --out build
```

### macOS / Linux (bash/zsh)

```bash
python3 generate_menus.py \
  --weekly "Residents Menu WC 15-09-2025.docx" \
  --all-days \
  --templates "./templates" \
  --out "./build"
```

---

## What the script does (rules baked in)

- **Templates & formatting**  
  It never changes template styles; it only fills placeholders so fonts, weights, colours and layout stay identical to your DOCX templates.

- **Standard menu**
  - Vegetarian/veg-friendly lunch main first; meat second.
  - Dessert titles sentence-case + `(V)`.
  - Supper starter fixed as **Chef’s choice soup (V)**.

- **Vegan menu**
  - Lunch always includes **Jacket potato and toppings (Ve)** (short title on menu).
  - All vegan **dessert 1** titles (lunch & supper) end with `(Ve)` and are **highlighted**; ice-cream lines are **not** highlighted.
  - Supper starter **Chef’s choice soup (Ve)** is **highlighted**.
  - **No Milk or Egg** appear in vegan allergens lines (they are scrubbed).
  - Vegan supper special title ends with `(Ve)`.

- **Allergens sheet (table)**
  - Month banner becomes `D Month YYYY` (e.g., `18 September 2025`).
  - Two section rows: `— Standard —`, `— Vegan —`.
  - Order per section:  
    Lunch starter(s) → Lunch mains (veg/vegan first) → Optional sides → Lunch dessert 1 → Supper soup → Supper special → Supper dessert 1.
  - Vegan **lunch starter** uses the **actual weekly title** with `(Ve)`; **only one** vegan supper soup row.
  - Jacket potato row is short-titled and ticks: **Celery, Cereals with Gluten, Mustards, Sulphur**.
  - Vegan items never tick **Milk** or **Eggs**; vegan dessert allergens are always **Gluten, Nuts, Soya, Sulphites**.

---

## Output

Only **ZIPs** are written to `--out` (default `./build`).  
Each ZIP is named:  
`Day-DD-MM-YYYY-menus-and-allergens.zip`  
and contains:
- `Residents_DD-MM-YYYY.docx`
- `Residents_DD-MM-YYYY_vegan.docx`
- `Allergens_Residents_DD-MM-YYYY.docx`

---

## Troubleshooting

- **`unrecognized arguments: \ \ \`** – You pasted bash backslashes into PowerShell/CMD. Use a single-line command or PowerShell’s backtick (`) for multiline.
- **`ModuleNotFoundError: No module named 'docx'`** – Activate your venv and run `pip install -r requirements.txt`.
- **`PermissionError: ... is in use`** – Close any open copy of the output ZIP and re-run.
- **`PackageNotFoundError: Package not found at ...`** – Check the path/filename passed to `--weekly` or the templates flags.
- **`UserWarning: pkg_resources is deprecated`** – Harmless from `docxcompose`; safe to ignore.

---

## Web App (Optional)

Run the project as a **local web app** so non‑technical staff can generate daily ZIPs via a browser.

### Extra dependencies
```bash
pip install fastapi uvicorn python-multipart
```

### Start the server
```bash
uvicorn webapp:app --reload --port 8000
```

### Usage (browser)
1. Open **http://localhost:8000**.
2. Upload the **weekly menu DOCX** and (optionally) a **templates** ZIP/folder path.
3. Choose either **one day** (provide `YYYY-MM-DD`) or **all days**.
4. Click **Generate**. The browser downloads the ZIP named `Day-DD-MM-YYYY-menus-and-allergens.zip` containing:
   - `Residents_DD-MM-YYYY.docx`
   - `Residents_DD-MM-YYYY_vegan.docx`
   - `Allergens_Residents_DD-MM-YYYY.docx`

### File: `webapp.py`
```python
from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import tempfile, shutil, os, subprocess

app = FastAPI()

# Simple upload form (GET /)
@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <html>
      <body style="font-family: system-ui; max-width: 720px; margin: 40px auto;">
        <h2>Henbrook Daily Menu Generator</h2>
        <form action="/generate" method="post" enctype="multipart/form-data">
          <p><label>Weekly DOCX: <input type="file" name="weekly" required /></label></p>
          <p><label>Templates folder/zip (optional): <input type="file" name="templates" /></label></p>
          <p><label>Date (YYYY-MM-DD, optional): <input type="text" name="date" placeholder="2025-09-18"></label></p>
          <p><label><input type="checkbox" name="all_days" value="true"> Generate all 7 days</label></p>
          <p><button type="submit">Generate</button></p>
        </form>
      </body>
    </html>
    """

@app.post("/generate")
async def generate(weekly: UploadFile, templates: UploadFile | None = None,
                   date: str | None = Form(default=None),
                   all_days: bool = Form(default=False)):
    tmpdir = tempfile.mkdtemp()
    weekly_path = os.path.join(tmpdir, weekly.filename)
    with open(weekly_path, "wb") as f:
        shutil.copyfileobj(weekly.file, f)

    cmd = ["python", "generate_menus.py", "--weekly", weekly_path, "--out", tmpdir]
    if date:
        cmd += ["--date", date]
    if all_days:
        cmd.append("--all-days")
    if templates:
        tpl_path = os.path.join(tmpdir, templates.filename)
        with open(tpl_path, "wb") as f:
            shutil.copyfileobj(templates.file, f)
        # If a single uploaded file is a folder/zip of templates, pass path to --templates
        cmd += ["--templates", tpl_path]

    subprocess.run(cmd, check=True)

    # Return the first ZIP generated
    for name in os.listdir(tmpdir):
        if name.endswith(".zip"):
            return FileResponse(os.path.join(tmpdir, name), filename=name)

    # If nothing found, return a simple message
    return {"detail": "No ZIP produced. Check weekly file and templates."}
```

### Notes & deployment tips
- The web app **uses the same rules** and template‑driven formatting as the CLI.
- For staff use on your LAN: `uvicorn webapp:app --host 0.0.0.0 --port 8000`.
- Consider a one‑liner Dockerfile for portability:
  ```dockerfile
  FROM python:3.11-slim
  WORKDIR /app
  COPY . .
  RUN pip install -r requirements.txt && pip install fastapi uvicorn python-multipart
  EXPOSE 8000
  CMD ["uvicorn", "webapp:app", "--host", "0.0.0.0", "--port", "8000"]
  ```
- Security: this is intended for **trusted, local** environments. Do not expose publicly without auth and rate‑limits.
