# report_generator_v1/peptide_processor.py
from __future__ import annotations

import os
import glob
from pathlib import Path
import pandas as pd
import numpy as np


class PeptideProcessor:
    def __init__(self, xlsx_path: str, txt_path: str, seq_xls_path: str, output_dir: str):
        self.xlsx_path = xlsx_path
        self.txt_path = txt_path
        self.seq_xls_path = seq_xls_path
        self.output_dir = output_dir
        self.merged_df = None
        self.final_df = None
        self.report_df = None

    # --- Step 1: Load data ---
    def load_data(self):
        # Stats XLSX (male/female) and patient TXT
        self.df_xlsx = pd.read_excel(self.xlsx_path)  # openpyxl
        self.df_txt = pd.read_csv(self.txt_path, sep="\t", skiprows=[1])

    # --- Step 2: Process TXT numeric data ---
    def process_txt_data(self):
        """
        Process numeric columns of df_txt:
        1. Multiply by 10
        2. Replace 0 with NaN
        3. Replace values <1 with 1
        4. Apply log2 transform
        """
        df = self.df_txt.copy()

        # Robustly detect the ID column (case-insensitive, allow variants)
        id_candidates = [c for c in df.columns if str(c).strip().lower().startswith("fid") and "auswertung" in str(c).lower()]
        if not id_candidates:
            raise KeyError(
                "Could not find an ID column like 'fidAuswertung' in TXT. "
                f"Columns were: {list(df.columns)}"
            )
        id_col = id_candidates[0]
        if id_col != "fidAuswertung":
            df = df.rename(columns={id_col: "fidAuswertung"})

        # Numeric columns = everything except the id
        numeric_cols = [c for c in df.columns if c != "fidAuswertung"]

        # Multiply by 10
        df[numeric_cols] = df[numeric_cols] * 10

        # Replace 0 with NaN
        df[numeric_cols] = df[numeric_cols].replace(0, np.nan)

        # Replace values <1 with 1
        df[numeric_cols] = df[numeric_cols].applymap(lambda x: max(x, 1) if pd.notna(x) else x)

        # Apply log2
        df[numeric_cols] = np.log2(df[numeric_cols])

        self.df_txt = df


    # --- Step 3: Merge XLSX and TXT ---
    def merge_xlsx_txt(self):
        for need, dfname in (("fidAuswertung", "df_xlsx"), ("fidAuswertung", "df_txt")):
            if need not in getattr(self, dfname).columns:
                raise KeyError(f"Missing '{need}' column in {dfname}. Columns: {list(getattr(self, dfname).columns)}")

        self.merged_df = pd.merge(self.df_xlsx, self.df_txt, on="fidAuswertung", how="inner")
        if self.merged_df.empty:
            raise ValueError("Merge produced 0 rows. Check that 'fidAuswertung' values overlap between XLSX and TXT.")

        last_txt_col = str(self.merged_df.columns[-1])
        self.merged_df = self.merged_df.rename(columns={last_txt_col: "Your Result"})
        
    # --- Step 4: Add Range column ---
    def add_range(self):
        def check_range(row):
            lower, upper, value = row['2.5%'], row['97.5%'], row["Your Result"]
            if pd.isna(value): return None
            if value > upper: return "up"
            if value < lower: return "down"
            return "within"
        self.merged_df['Range'] = self.merged_df.apply(check_range, axis=1)

    # --- Step 5: Merge sequence info ---
    def merge_sequence_info(self):
        cols = [
            'fidAuswertung_ML2.5',
            'Confidence (based on Xcorr) ML2.5',
            'Sequence ML2.5',
            'Modifications',
            'Master UniProt Name'
        ]
        # NOTE: this is an .xls file -> requires xlrd<2.0
        df_seq = pd.read_excel(self.seq_xls_path, usecols=cols)  # xlrd for .xls
        df_seq = df_seq.rename(columns={'fidAuswertung_ML2.5': 'fidAuswertung'})
        self.final_df = pd.merge(self.merged_df, df_seq, on="fidAuswertung", how="inner")

    # --- Step 6: Mark modified sequences ---
    def mark_modified_sequences(self):
        df = self.final_df
        df['Sequence ML2.5'] = df.apply(
            lambda row: f"{row['Sequence ML2.5']}*"
            if pd.notna(row['Modifications']) and str(row['Modifications']).strip().lower() != "none"
            else row['Sequence ML2.5'],
            axis=1
        )
        self.final_df = df

    # --- Step 7: Filter rows ---
    def filter_rows(self):
        df = self.final_df
        df = df[
            df["Sequence ML2.5"].notna() & df["Sequence ML2.5"].str.strip().ne("") &
            df["Your Result"].notna() & (df["Your Result"] != 0)
        ]
        self.final_df = df

    # --- Step 8: Save final merged result ---
    def save_final(self, filename="final_reference_patient_seqs_filtered.xlsx"):
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        self.final_df.to_excel(str(Path(self.output_dir, filename)), index=False)

    # --- Step 9: Prepare report ---
    def prepare_report(self):
        df = self.final_df.drop(columns=["Confidence (based on Xcorr) ML2.5", "frequency", "mean"], errors='ignore')
        df = df.rename(columns={
            "Sequence ML2.5": "Sequence",
            "Modifications": "Modification",
            "Master UniProt Name": "Protein Name"
        })
        df = df[[
            "fidAuswertung", "Sequence", "Modification", "Protein Name",
            "median", "2.5%", "97.5%", "Your Result", "Range"
        ]]
        self.report_df = df

    # --- Step 10: Save report ---
    def save_report(self, filename="report_peptide_result_final.xlsx"):
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        self.report_df.to_excel(str(Path(self.output_dir, filename)), index=False)

    # --- Run all steps ---
    def run_all(self):
        self.load_data()
        self.process_txt_data()
        self.merge_xlsx_txt()
        self.add_range()
        self.merge_sequence_info()
        self.mark_modified_sequences()
        self.filter_rows()
        self.save_final()
        self.prepare_report()
        self.save_report()


def build_peptide_report_excel(
    *,
    working_dir: str,
    package_data_dir: str,
    gender_raw: str,
    seq_xls_name: str = "ML2.5_21082025_seq.xls",
) -> str:
    """
    Auto-selects:
      - stats XLSX by gender (from package data/)
      - the TXT file in working_dir (prefers '*transposed*.txt', otherwise newest '*.txt')
    Builds report_peptide_result_final.xlsx in working_dir and returns its path.
    """
    g = (gender_raw or "").strip().lower()
    female_markers = {"weiblich", "w", "frau", "female", "f"}
    stats_xlsx = (
        "Healthy_patients_female_n836_log2_stats_nonNeg.xlsx"
        if g in female_markers
        else "Healthy_patients_male_n917_log2_stats_nonNeg.xlsx"
    )

    data_dir = Path(package_data_dir)
    xlsx_path = data_dir / stats_xlsx
    seq_xls_path = data_dir / seq_xls_name

    if not xlsx_path.exists():
        raise FileNotFoundError(f"Stats file not found: {xlsx_path}")
    if not seq_xls_path.exists():
        raise FileNotFoundError(f"Sequence .xls not found: {seq_xls_path}")

    wd = Path(working_dir)

    # Prefer *transposed*.txt, else newest *.txt
    candidates = sorted(wd.glob("*transposed*.txt"), key=os.path.getmtime)
    if not candidates:
        candidates = sorted(wd.glob("*.txt"), key=os.path.getmtime)
    if not candidates:
        raise FileNotFoundError(
            f"No TXT file found in {working_dir}. "
            "Expected a transposed export (e.g., '*transposed*.txt')."
        )
    txt_path = candidates[-1]

    out_xlsx = wd / "report_peptide_result_final.xlsx"
    processor = PeptideProcessor(
        xlsx_path=str(xlsx_path),
        txt_path=str(txt_path),
        seq_xls_path=str(seq_xls_path),
        output_dir=str(wd),
    )
    processor.run_all()

    # sanity check
    if not out_xlsx.exists():
        raise RuntimeError(f"Expected Excel not created: {out_xlsx}")

    return str(out_xlsx)
