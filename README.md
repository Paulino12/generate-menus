# Henbrook Daily Menu Generator

Generate resident-friendly **daily menus** and an **allergens sheet** from a single **weekly DOCX grid** using Python. The tool fills your pre‑styled templates and outputs **one ZIP per day**.

- ✅ Standard, Vegan, and Allergens DOCXs rendered from your templates
- ✅ Styles preserved (fonts, weights, colours, layout stay as in the templates)
- ✅ Streamlit web UI or CLI
- ✅ Output naming (fixed):
  - `Residents_DD-MM-YYYY.docx`
  - `Residents_DD-MM-YYYY_vegan.docx`
  - `Allergens_Residents_DD-MM-YYYY.docx`
  - Packaged as `Day-DD-MM-YYYY-menus-and-allergens.zip`

---

## Folder layout

```
project/
├─ generate_menus.py
├─ app.py                     # Streamlit UI
├─ templates/
│  ├─ standard.docx
│  ├─ vegan.docx
│  └─ allergens.docx
├─ examples/                  # (optional) sample weekly DOCXs
└─ build/                     # Generated ZIPs (--out), usually git‑ignored
```

> You can keep templates anywhere; default is `./templates`. You can also pass paths to individual template files with flags.

---

## Requirements

Install once (venv recommended):

```bash
python -m venv .venv
# Windows PowerShell
. .venv/Scripts/Activate.ps1
# macOS/Linux
# source .venv/bin/activate

pip install -r requirements.txt
```

Your `requirements.txt` should include at least:

```
python-docx
docxtpl
PyYAML
streamlit
```

---

## Run with Streamlit (recommended)

Local UI for non‑technical users.

```bash
streamlit run app.py
```

Steps:

1. Upload the **weekly menu DOCX**.
2. (Optional) Upload custom templates or use those in `./templates`.
3. Choose **one day** (pick a date) or **all 7 days**.
4. Click **Generate** → download the ZIP(s).

Hosted version: https://paulino12-generate-menus-app-jl3nqs.streamlit.app/

---

## CLI usage

### Windows PowerShell

Generate **one day**:

```powershell
python .\generate_menus.py --weekly ".\examples\Residents Menu WC 15-09-2025.docx" --date 2025-09-18 --templates ".	emplates" --out ".uild"
```

Generate **all 7 days**:

```powershell
python .\generate_menus.py --weekly ".\examples\Residents Menu WC 15-09-2025.docx" --all-days --templates ".	emplates" --out ".uild"
```

Explicit template paths:

```powershell
python .\generate_menus.py `
  --weekly ".\examples\Residents Menu WC 15-09-2025.docx" `
  --standard_tpl ".	emplates\standard.docx" `
  --vegan_tpl ".	emplatesegan.docx" `
  --allergens_tpl ".	emplatesllergens.docx" `
  --date 2025-09-18 `
  --out ".uild"
```

> PowerShell line continuation uses the **backtick** (`).

### Windows CMD

```bat
python generate_menus.py ^
  --weekly "examples\Residents Menu WC 15-09-2025.docx" ^
  --all-days ^
  --templates "templates" ^
  --out build
```

### macOS / Linux (bash/zsh)

```bash
python3 generate_menus.py   --weekly "examples/Residents Menu WC 15-09-2025.docx"   --all-days   --templates "./templates"   --out "./build"
```

---

## Built‑in rules (baked into the generator)

- **Templates & formatting**: never change styles—only fill placeholders so your document styling remains intact.
- **Standard menu**:
  - Vegetarian/veg‑friendly lunch main first; meat second.
  - Dessert titles sentence‑case and end with `(V)`.
  - Supper starter fixed: **Chef’s choice soup (V)** and highlighted.
- **Vegan menu**:
  - Lunch always includes **Jacket potato and toppings (Ve)**.
  - Vegan **dessert 1** titles (lunch & supper) end with `(Ve)` and are **highlighted** (ice‑cream lines aren’t highlighted).
  - Supper starter: **Chef’s choice soup (Ve)** and **highlighted**.
  - Milk/Egg are scrubbed from vegan allergens.
  - If vegan descriptions are missing, borrow from the standard equivalents and highlight the borrowed description.
- **Allergens sheet**:
  - 2 pages: Page 1 filled ticks, Page 2 blank for day‑of notes.
  - Vegan dessert allergens always **Gluten, Nuts, Soya, Sulphites**.
  - Jacket potato row short‑titled; ticks: **Celery, Cereals with Gluten, Mustard, Sulphur**.
  - Keep footer/legend rows intact.

---

## Output

Only **ZIPs** are written to `--out` (default `./build`).  
Each ZIP is named `Day-DD-MM-YYYY-menus-and-allergens.zip` and contains the three DOCXs listed above.

---

## Troubleshooting

- **`unrecognized arguments: \ \ \`** – You pasted bash backslashes into PowerShell/CMD. Use a single‑line command or PowerShell’s backtick (`) for multiline.
- **`ModuleNotFoundError: No module named 'docx'`** – Activate your venv and run `pip install -r requirements.txt`.
- **`PermissionError: ... is in use`** – Close any open copy of the output ZIP and re‑run.
- **`PackageNotFoundError: Package not found at ...`** – Check the path/filename passed to `--weekly` or template flags.
- **No ZIP produced** – Confirm the weekly DOCX matches the expected grid (rows/labels) and that template files exist at the paths you provided.

---

## Housekeeping

- Git‑ignore local/build artifacts:
  ```
  __pycache__/
  *.pyc
  build/
  *.zip
  .venv/
  .vscode/
  .streamlit/
  .DS_Store
  ```
- Keep only **one README** (this file).

---

## License

Proprietary within Henbrook House / Connaught Care Group. Not for redistribution without permission.
