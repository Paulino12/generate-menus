# Henbrook Daily Menu Generator

A single Python program that:
- Parses a weekly menu DOCX (the grid you maintain).
- Renders **Standard**, **Vegan**, and **Allergens** daily DOCXs from your go‑forward templates.
- Packages each day's three files into a correctly‑named ZIP: `Day-DD-MM-YYYY-menus-and-allergens.zip`.

## Requirements
```bash
pip install -r requirements.txt
```

`requirements.txt` contains:
- `python-docx`
- `docxtpl`
- `PyYAML`

## Usage

### Generate one day
```bash
python generate_menus.py \
  --weekly "Residents Menu WC 15-09-2025.docx" \
  --standard_tpl "Residents_14-09-2025.docx" \
  --vegan_tpl "Residents_14-09-2025_vegan.docx" \
  --allergens_tpl "Allergens_Residents_14-09-2025.docx" \
  --date 2025-09-17 \
  --out build
```

### Generate all 7 days
```bash
python generate_menus.py \
  --weekly "Residents Menu WC 15-09-2025.docx" \
  --standard_tpl "Residents_14-09-2025.docx" \
  --vegan_tpl "Residents_14-09-2025_vegan.docx" \
  --allergens_tpl "Allergens_Residents_14-09-2025.docx" \
  --all-days \
  --out build
```

The script creates:
- `Residents_DD-MM-YYYY.docx`
- `Residents_DD-MM-YYYY_vegan.docx`
- `Allergens_Residents_DD-MM-YYYY.docx`
- A ZIP named `Day-DD-MM-YYYY-menus-and-allergens.zip` with the three files inside.

## Notes
- Supper starter line is fixed as **Chef’s choice soup** (V/Ve as appropriate), and we copy lunch soup allergens across (per your standards).
- Vegan lunch always includes the **Plant Buttered Jacket Potato and Toppings (Ve)** line as the first main, followed by the weekly vegan main.
- Vegan dessert allergens are forced to `Gluten, Nuts, Soya, Sulphites`.
- Allergen sheet excludes the following standard lines on Page 1: 
  Chef’s choice soup (V), Chef’s choice soup (Ve), Ice creams / sorbet (V), Selection of vegan ice creams or sorbet with seasonal fruits (Ve), Henbrook’s assorted sandwich (Ve).

If anything about your weekly grid format changes (extra rows, different ordering), adjust the row index section in `parse_week()` accordingly.
