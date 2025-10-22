import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docxtpl import DocxTemplate
import importlib.resources

from .db_utils import get_patient
from .excel_utils import load_scores
from .context_utils import build_context
from .report_utils import evaluate_thresholds, add_figures, convert_docx_to_pdf_with_progress

import os


class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")
        self.working_dir = None

        # --- UI ---
        tk.Label(root, text="Search by:").grid(row=0, column=0, padx=5, pady=5)
        self.search_type = ttk.Combobox(root, values=["Patient ID", "Sample ID", "Name"])
        self.search_type.current(0)
        self.search_type.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(root, text="Value:").grid(row=1, column=0, padx=5, pady=5)
        self.search_value_entry = tk.Entry(root)
        self.search_value_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(root, text="Select Directory", command=self.select_directory).grid(row=2, column=0, padx=5, pady=5)
        self.dir_label = tk.Label(root, text="No directory selected")
        self.dir_label.grid(row=2, column=1, padx=5, pady=5)

        tk.Button(root, text="Generate Report", command=self.generate_report).grid(row=3, column=0, columnspan=2, pady=10)
        tk.Button(root, text="Generate Annex", command=self.generate_annex).grid(row=4, column=0, columnspan=2, pady=10)

    def select_directory(self):
        self.working_dir = filedialog.askdirectory()
        if self.working_dir:
            self.dir_label.config(text=self.working_dir)

    def generate_report(self):
        patient_data = get_patient(self.search_type.get(), self.search_value_entry.get().strip())
        if not patient_data:
            messagebox.showerror("Error", "No patient found")
            return

        context = build_context(patient_data)
        mapped_scores = load_scores(self.working_dir)
        context.update(mapped_scores)
        context.update(evaluate_thresholds(mapped_scores))

        # Load template from package
        with importlib.resources.path("report_generator_v1.templates", "template_MOS.docx") as tpl_path:
            tpl = DocxTemplate(tpl_path)

        add_figures(tpl, self.working_dir, context, prefix="fig")
        tpl.render(context)

        first_name, last_name = patient_data[0], patient_data[1]
        output_path = os.path.join(self.working_dir, f"{last_name}_{first_name}_Report.docx")
        tpl.save(output_path)
        messagebox.showinfo("Success", f"Report saved as {output_path}")

        # --- Convert to PDF automatically ---
        convert_docx_to_pdf_with_progress(self.root, output_path, self.working_dir)

    def generate_annex(self):
        patient_data = get_patient(self.search_type.get(), self.search_value_entry.get().strip())
        if not patient_data:
            messagebox.showerror("Error", "No patient found")
            return

        context = build_context(patient_data)

        # Load annex template from package
        with importlib.resources.path("report_generator_v1.templates", "template_annex.docx") as tpl_path:
            tpl = DocxTemplate(tpl_path)

        add_figures(tpl, self.working_dir, context, prefix="annex_fig")
        tpl.render(context)

        first_name, last_name = patient_data[0], patient_data[1]
        output_path = os.path.join(self.working_dir, f"{last_name}_{first_name}_Annex.docx")
        tpl.save(output_path)
        messagebox.showinfo("Success", f"Annex saved as {output_path}")
        
        # --- Convert to PDF automatically ---
        convert_docx_to_pdf_with_progress(self.root, output_path, self.working_dir)

def main():
    root = tk.Tk()
    app = ReportGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
