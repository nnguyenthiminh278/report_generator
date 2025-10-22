import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docxtpl import DocxTemplate
import importlib.resources

from .db_utils import get_patient
from .excel_utils import load_scores
from .context_utils import build_context
from .report_utils import (
    evaluate_thresholds,
    add_figures,
    convert_docx_to_pdf_with_progress,
    stage_proteom_figs_for_add_figures,
)
from pathlib import Path
import os

def _get_package_data_dir() -> Path:
    """
    Returns a concrete filesystem Path to report_generator_v1/data,
    even when installed as a zipped/egg package (uses as_file).
    """
    try:
        res = ir.files("report_generator_v1").joinpath("data")
        with ir.as_file(res) as p:
            return Path(p)
    except Exception:
        # Fallback: resolve relative to this file (dev installs)
        return Path(__file__).resolve().parent / "data"
    
def _load_template_from_package(filename: str) -> DocxTemplate:
    """
    Try to load a template by file name from the packaged templates dir.
    """
    with importlib.resources.path("report_generator_v1.templates", filename) as tpl_path:
        return DocxTemplate(tpl_path)


def _get_template(lang: str, base: str) -> DocxTemplate:
    """
    Look for an EN variant when lang == 'EN', otherwise use the DE base.
    Fallback to DE if the EN file does not exist in the package.
    """
    if lang == "EN":
        # Expect something like template_MOS_EN.docx or template_annex_EN.docx
        en_name = base.replace(".docx", "_EN.docx")
        try:
            return _load_template_from_package(en_name)
        except FileNotFoundError:
            # Fallback to DE/base
            pass
    return _load_template_from_package(base)


class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")
        self.working_dir = None

        # --- UI ---
        tk.Label(root, text="Search by:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.search_type = ttk.Combobox(root, values=["Patient ID", "Sample ID", "Name"], state="readonly")
        self.search_type.current(0)
        self.search_type.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Label(root, text="Value:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.search_value_entry = tk.Entry(root, width=30)
        self.search_value_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Button(root, text="Select Directory", command=self.select_directory).grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.dir_label = tk.Label(root, text="No directory selected", anchor="w")
        self.dir_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # --- Language selector (DE/EN) ---
        tk.Label(root, text="Language:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.lang_var = tk.StringVar(value="DE")
        self.lang_combo = ttk.Combobox(root, values=["DE", "EN"], state="readonly", textvariable=self.lang_var, width=6)
        self.lang_combo.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        tk.Button(root, text="Generate Report", command=self.generate_report).grid(row=4, column=0, columnspan=2, pady=10)
        tk.Button(root, text="Generate Annex", command=self.generate_annex).grid(row=5, column=0, columnspan=2, pady=10)

    def select_directory(self):
        self.working_dir = filedialog.askdirectory()
        if self.working_dir:
            self.dir_label.config(text=self.working_dir)

    # --- Internal helper to stage proteom figs (CKD/CAD/HF/Oncorisk) ---
    def _stage_proteom_figs(self, language_choice: str):
        """
        Generates CKD/CAD/HF/Oncorisk distribution plots into self.working_dir
        with numbered suffixes _7.._10 so add_figures() can auto-pick them up.
        """
        if not self.working_dir:
            return

        data_dir = _get_package_data_dir() 
        if not data_dir.exists():
            # Not fatal—allow reports without those figures if data folder is missing
            messagebox.showwarning("Warning", f"Data folder not found: {data_dir}\nSkipping proteom distribution figures.")
            return

        try:
            # languages expects an iterable; we generate only the selected one
            stage_proteom_figs_for_add_figures(
                working_dir=Path(self.working_dir),
                data_dir=data_dir,
                languages=(language_choice,),  # ("DE",) or ("EN",)
                # numbering defaults to CKD:7, CAD:8, HF:9, Oncorisk:10
            )
        except Exception as e:
            # Non-blocking: continue without those figures
            messagebox.showwarning("Warning", f"Could not generate proteom figures:\n{e}")

    def generate_report(self):
        if not self.working_dir:
            messagebox.showerror("Error", "Please select a working directory first.")
            return

        patient_data = get_patient(self.search_type.get(), self.search_value_entry.get().strip())
        if not patient_data:
            messagebox.showerror("Error", "No patient found")
            return

        context = build_context(patient_data)

        # Map scores & thresholds
        mapped_scores = load_scores(self.working_dir)
        context.update(mapped_scores)
        context.update(evaluate_thresholds(mapped_scores))

        # --- Stage proteom figures (numbered _7.._10 for add_figures) ---
        lang = self.lang_var.get()  # "DE" or "EN"
        self._stage_proteom_figs(lang)

        # --- Load template based on language (fallback-safe) ---
        tpl = _get_template(lang, base="template_MOS.docx")

        # --- Add all figures that match *_[0-9].png (includes the staged proteom figs) ---
        add_figures(tpl, self.working_dir, context, prefix="fig")

        # --- Render & save ---
        first_name, last_name = patient_data[0], patient_data[1]
        output_path = os.path.join(self.working_dir, f"{last_name}_{first_name}_Report_{lang}.docx")
        tpl.render(context)
        tpl.save(output_path)
        messagebox.showinfo("Success", f"Report saved as {output_path}")

        # --- Convert to PDF automatically ---
        convert_docx_to_pdf_with_progress(self.root, output_path, self.working_dir)

    def generate_annex(self):
        if not self.working_dir:
            messagebox.showerror("Error", "Please select a working directory first.")
            return

        patient_data = get_patient(self.search_type.get(), self.search_value_entry.get().strip())
        if not patient_data:
            messagebox.showerror("Error", "No patient found")
            return

        context = build_context(patient_data)

        # Stage proteom figures for annex too (if annex uses numbered placeholders)
        lang = self.lang_var.get()
        self._stage_proteom_figs(lang)

        # Load annex template (language-aware)
        tpl = _get_template(lang, base="template_annex.docx")

        # Add figures with annex prefix (expects {{ annex_fig7 }}, etc. if you use them there)
        add_figures(tpl, self.working_dir, context, prefix="annex_fig")
        tpl.render(context)

        first_name, last_name = patient_data[0], patient_data[1]
        output_path = os.path.join(self.working_dir, f"{last_name}_{first_name}_Annex_{lang}.docx")
        tpl.save(output_path)
        messagebox.showinfo("Success", f"Annex saved as {output_path}")

        # Convert to PDF automatically
        convert_docx_to_pdf_with_progress(self.root, output_path, self.working_dir)


def main():
    root = tk.Tk()
    app = ReportGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
