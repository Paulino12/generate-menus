import os
import io
import tempfile
import datetime as dt
import streamlit as st

# Import your existing generator module
import generate_menus as gm


st.set_page_config(page_title="Henbrook Menu Generator", page_icon="🍽️", layout="centered")
st.title("Daily Menu Generator")

st.markdown("""
Upload your **weekly menu DOCX**, choose templates (built-in or upload custom),
pick a day (or all days), and download the generated **ZIP** per day.
""")

# ----------------------------- Inputs -----------------------------
weekly_file = st.file_uploader("Weekly menu (.docx)", type=["docx"], accept_multiple_files=False)

templates_mode = st.radio("Templates to use", ["Use repository templates (./templates)", "Upload custom templates"])

std_up = veg_up = all_up = None
if templates_mode == "Upload custom templates":
    st.write("Upload each template:")
    std_up = st.file_uploader("Standard template (.docx)", type=["docx"], key="std_tpl")
    veg_up = st.file_uploader("Vegan template (.docx)", type=["docx"], key="veg_tpl")
    all_up = st.file_uploader("Allergens template (.docx)", type=["docx"], key="all_tpl")

mode = st.radio("What to generate?", ["One day", "All 7 days"])

selected_date_iso = None
parsed_days = None

# Try to pre-parse the weekly doc to offer day choices
if weekly_file is not None:
    with tempfile.TemporaryDirectory() as tdir:
        weekly_path = os.path.join(tdir, "weekly.docx")
        with open(weekly_path, "wb") as f:
            f.write(weekly_file.read())
        try:
            week = gm.parse_week(weekly_path)
            parsed_days = week["days"]
            if mode == "One day":
                options = [f'{d["header"]["weekday"]} – {d["header"]["date_iso"]}' for d in parsed_days]
                choice = st.selectbox("Pick a day", options, key="day_pick")
                idx = options.index(choice) if options else 0
                selected_date_iso = parsed_days[idx]["header"]["date_iso"]
        except Exception as e:
            st.error(f"Failed to read weekly grid: {e}")

# ----------------------------- Generate -----------------------------
def _save_upload_to(path: str, uploaded_file) -> str:
    with open(path, "wb") as f:
        f.write(uploaded_file.read())
    return path

def _resolve_templates(tmpdir: str):
    if templates_mode == "Upload custom templates":
        if not (std_up and veg_up and all_up):
            raise RuntimeError("Please upload Standard, Vegan, and Allergens templates.")
        std_path = os.path.join(tmpdir, "standard.docx")
        veg_path = os.path.join(tmpdir, "vegan.docx")
        all_path = os.path.join(tmpdir, "allergens.docx")
        _save_upload_to(std_path, std_up)
        _save_upload_to(veg_path, veg_up)
        _save_upload_to(all_path, all_up)
        return std_path, veg_path, all_path
    # Repository templates (./templates)
    std_path = os.path.join("templates", "standard.docx")
    veg_path = os.path.join("templates", "vegan.docx")
    all_path = os.path.join("templates", "allergens.docx")
    for p in (std_path, veg_path, all_path):
        if not os.path.exists(p):
            raise RuntimeError(f"Template not found: {p}")
    return std_path, veg_path, all_path


def generate_and_collect(weekly_bytes: bytes, date_iso: str | None):
    """Run generation via gm.render_day_to_zip and return (label, zip_bytes) list."""
    outputs: list[tuple[str, bytes]] = []
    with tempfile.TemporaryDirectory() as tdir:
        # Save weekly
        weekly_path = os.path.join(tdir, "weekly.docx")
        with open(weekly_path, "wb") as f:
            f.write(weekly_bytes)

        # Parse week once
        week = gm.parse_week(weekly_path)
        days = week["days"]

        # Resolve templates
        std_tpl, veg_tpl, all_tpl = _resolve_templates(tdir)

        # Decide what to render
        targets = []
        if date_iso:
            t = next((d for d in days if d["header"]["date_iso"] == date_iso), None)
            if not t:
                raise RuntimeError(f"Date {date_iso} not found in weekly grid.")
            targets = [t]
        else:
            targets = days

        # Render each day to ZIP (into a temp out dir), read back as bytes
        outdir = os.path.join(tdir, "out")
        os.makedirs(outdir, exist_ok=True)
        for d in targets:
            zpath = gm.render_day_to_zip(d, std_tpl, veg_tpl, all_tpl, outdir)
            label = os.path.basename(zpath)
            with open(zpath, "rb") as f:
                outputs.append((label, f.read()))
    return outputs


clicked = st.button("Generate")

if clicked:
    if weekly_file is None:
        st.error("Please upload the weekly DOCX.")
    else:
        try:
            weekly_bytes = weekly_file.getvalue()
            date_iso = selected_date_iso if (mode == "One day") else None
            with st.spinner("Generating ZIP(s)..."):
                results = generate_and_collect(weekly_bytes, date_iso)
            st.success("Done.")
            for label, data in results:
                st.download_button(
                    label=f"⬇️ Download {label}",
                    data=data,
                    file_name=label,
                    mime="application/zip",
                )
        except Exception as e:
            st.error(f"Generation failed: {e}")
