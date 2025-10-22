# -*- coding: utf-8 -*-
"""
Utility functions for report generation:
- Evaluate thresholds from Word templates.
- Add figures dynamically with optional text overlays.
- Convert DOCX → PDF (preferring ONLYOFFICE, fallback to LibreOffice).
"""

import os
import glob
import tempfile
import platform
import subprocess
import threading
import shutil
import importlib.resources

from tkinter import messagebox, ttk
from docxtpl import InlineImage
from docx.shared import Mm
from docx import Document
from PIL import Image, ImageDraw, ImageFont

from pathlib import Path

from .figures.proteom_distribution import generate_all_proteom_distributions, DEFAULT_MODELS
# ---------------------------------------------------------------------
# 1️⃣ Evaluate thresholds from Word template
# ---------------------------------------------------------------------
def evaluate_thresholds(mapped_scores):
    """Read thresholds from packaged Word template and evaluate risk sentences."""

    with importlib.resources.path("report_generator_v1.templates", "template_MOS.docx") as tpl_path:
        doc = Document(tpl_path)

    # Locate the table that contains "Normalbereich"
    target_table = next((tbl for tbl in doc.tables if "Normalbereich" in [cell.text.strip() for cell in tbl.rows[0].cells]), None)
    if not target_table:
        return {}

    # Mapping of table rows to score keys
    row_mapping = {
        1: "CKD_score",
        2: "CAD_score",
        3: "HF_score",
        4: "Onkorisk_score"
    }

    sentences = {}
    for row_idx, key in row_mapping.items():
        normal_text = target_table.cell(row_idx, 3).text.strip()
        try:
            threshold = float(normal_text.replace("<", "").strip().replace(",", "."))
        except ValueError:
            threshold = None

        value = float(mapped_scores.get(key, "0").replace(",", "."))
        sentences[key] = "keine" if (threshold is not None and value < threshold) else "eine"

    return {
        "ckd_sentence": sentences.get("CKD_score", "keine"),
        "cad_sentence": sentences.get("CAD_score", "keine"),
        "hf_sentence": sentences.get("HF_score", "keine"),
        "onco_sentence": sentences.get("Onkorisk_score", "keine"),
    }


# ---------------------------------------------------------------------
# 2️⃣ Add figures to report with dynamic overlay text
# ---------------------------------------------------------------------

def stage_proteom_figs_for_add_figures(
    working_dir: Path, data_dir: Path, *,
    languages=("DE",),                     # pass ("DE",) or ("EN",) based on the checkbox
    numbering=None,                        # {"CKD":7,"CAD":8,"HF":9,"Oncorisk":10}
) -> dict:
    """
    Generate CKD/CAD/HF/Oncorisk distribution plots into 'working_dir' with numbered filenames
    so your add_figures() will auto-insert them as fig7..fig10.
    """
    paths = generate_all_proteom_distributions(
        work_dir=working_dir,     # Mustertabelle location
        data_dir=data_dir,        # /data with model XLSX files
        dest_dir=working_dir,     # IMPORTANT: save where add_figures() scans
        languages=languages,
        numbering=numbering or {"CKD":7,"CAD":8,"HF":9,"Oncorisk":10},
    )
    return paths

def add_figures(tpl, working_dir, context, prefix="fig", size_rules=None):
    """
    Add all figures matching '*_[0-9].png' into the Word report context.
    - Automatically supports any number of figures (1–99+)
    - Customizable sizes using 'size_rules' dict
    """

    # Default figure sizes
    default_rules = {
        "1": Mm(150),  # fig1 = large
        "2": Mm(83),
        "3": Mm(90),
        "4": Mm(90),
        "5": Mm(90),
        "6": Mm(90),
    }
    default_width = Mm(80)

    if size_rules:
        default_rules.update(size_rules)

    figure_files = sorted(glob.glob(os.path.join(working_dir, "*_[0-9]*.png")))

    for file in figure_files:
        filename = os.path.basename(file)
        number = filename.split("_")[-1].split(".")[0]  # works for _1.png, _10.png, etc.
        placeholder = f"{prefix}{number}"
        width = default_rules.get(number, default_width)

        # Modify figure #2
        if number == "2":
            file = _overlay_text(file, "IHR PROTEOMPROFIL", (100, 20), "arial.ttf", 18)

        # Modify figures #3–#6
        if number in {"3", "4", "5", "6"}:
            file = _overlay_text(file, "Ihr persönliches Biomarker-Profil", (50, 18), "DejaVuSans.ttf", 20)

        # Insert figure into context
        context[placeholder] = InlineImage(tpl, file, width=width)


def _overlay_text(image_path, text, position, font_name="arial.ttf", font_size=18):
    """Draw text overlay on image and save a temporary copy."""
    img = Image.open(image_path).convert("RGBA")
    draw = ImageDraw.Draw(img)

    # Load font safely (fallback to default)
    try:
        font = ImageFont.truetype(font_name, font_size)
    except IOError:
        font = ImageFont.load_default()

    draw.text(position, text, fill=(255, 255, 255, 255), font=font)

    # Save to temporary file (don’t overwrite original)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        img.save(tmp.name)
        return tmp.name


# ---------------------------------------------------------------------
# 3️⃣ Convert DOCX → PDF with progress bar
# ---------------------------------------------------------------------
def convert_docx_to_pdf_with_progress(root, docx_path, output_dir):
    """
    Convert DOCX → PDF with a Tkinter progress bar.
    Prefers ONLYOFFICE (best quality), falls back to LibreOffice.
    Works on Linux, Windows, macOS.
    """

    progress_win = ttk.Frame(root, padding=20)
    progress_win.place(relx=0.5, rely=0.5, anchor="center")

    ttk.Label(progress_win, text="Generating PDF, please wait...").pack(pady=5)
    bar = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="indeterminate")
    bar.pack(pady=10)
    bar.start(10)

    def task():
        try:
            system = platform.system()
            pdf_path = docx_path.replace(".docx", ".pdf")

            # Windows → docx2pdf (MS Word)
            if system == "Windows":
                from docx2pdf import convert
                convert(docx_path, output_dir)

            # Linux/macOS → ONLYOFFICE preferred
            else:
                onlyoffice_exe = shutil.which("desktopeditors") or shutil.which("onlyoffice-desktopeditors")
                # subprocess.run(
                #     ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                #     check=True,
                # )                
                if onlyoffice_exe:
                    subprocess.run([onlyoffice_exe, "--convert", docx_path, "--output", output_dir], check=True)
                else:
                    subprocess.run(
                        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                        check=True,
                    )

            messagebox.showinfo("Success", f"PDF successfully created:\n{pdf_path}")

        except Exception as e:
            messagebox.showwarning("PDF Conversion Failed", f"Could not generate PDF:\n{e}")

        finally:
            bar.stop()
            progress_win.destroy()

    threading.Thread(target=task, daemon=True).start()
