# report_generator_v1/peptide_pdf.py
from __future__ import annotations

import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import cm


class ExcelToPDFExporter:
    def __init__(self, excel_file, pdf_file, patient_id,
                 male_female: str = "male", landscape_mode: bool = True):
        self.excel_file = excel_file
        self.pdf_file = pdf_file
        self.patient_id = patient_id
        self.male_female = male_female
        self.landscape_mode = landscape_mode

        # PDF page settings
        self.page_size = landscape(A4) if self.landscape_mode else A4
        self.left_margin = 1.5 * cm
        self.right_margin = 1.5 * cm
        self.top_margin = 2.5 * cm
        self.bottom_margin = 2 * cm

        self.styles = getSampleStyleSheet()
        self.wrap_style = ParagraphStyle(
            "wrap",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            wordWrap="CJK",
        )

    def load_excel(self) -> pd.DataFrame:
        df = pd.read_excel(self.excel_file)
        # Round numeric values
        df = df.round(2)
        return df

    def _auto_column_widths(self, df: pd.DataFrame, available_width: float, min_width: float = 60):
        """Compute proportional column widths with a minimum width."""
        col_lengths = []
        for i, col in enumerate(df.columns):
            max_len = max([len(str(col))] + [len(str(x)) for x in df.iloc[:, i]])
            col_lengths.append(max_len)

        total_length = sum(col_lengths) or 1
        col_widths = [(l / total_length) * available_width for l in col_lengths]
        col_widths = [max(w, min_width) for w in col_widths]

        total_width = sum(col_widths)
        if total_width > available_width:
            scale = available_width / total_width
            col_widths = [w * scale for w in col_widths]

        return col_widths

    def dataframe_to_table(self, df: pd.DataFrame) -> Table:
        available_width = self.page_size[0] - self.left_margin - self.right_margin
        col_widths = self._auto_column_widths(df, available_width)

        wrap_columns = {"Sequence", "Modification", "Protein Name"}
        footnote_cols = {"median", "2.5%", "97.5%"}
        bold_cols = wrap_columns | footnote_cols | {"Your Result", "Range"}

        data = []

        # Header row
        header = []
        for col in df.columns:
            text = str(col)
            if col in footnote_cols:
                text = f"{col}<super>a</super>"

            if col in bold_cols:
                style = ParagraphStyle("bold", fontName="Helvetica-Bold", fontSize=8)
                if col in wrap_columns:
                    style.wordWrap = "CJK"
                header.append(Paragraph(text, style))
            else:
                header.append(Paragraph(text, ParagraphStyle("hdr", fontName="Helvetica", fontSize=8)))
        data.append(header)

        # Data rows
        for _, row in df.iterrows():
            row_cells = []
            for col, val in zip(df.columns, row):
                text = "" if pd.isna(val) else str(val)
                if col in wrap_columns:
                    style = ParagraphStyle("wrap", fontName="Helvetica", fontSize=8, leading=12, wordWrap="CJK")
                else:
                    style = ParagraphStyle("normal", fontName="Helvetica", fontSize=8, leading=12)
                row_cells.append(Paragraph(text, style))
            data.append(row_cells)

        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return table

    def _header_footer(self, canvas, doc):
        canvas.saveState()
        width, height = self.page_size

        # Header
        header_text = f"{self.patient_id}, Page {doc.page}"
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(self.left_margin, height - self.top_margin + 0.5 * cm, header_text)

        # Footer
        footer_text = f"a: Distribution in healthy {self.male_female}"
        canvas.setFont("Helvetica", 8)
        canvas.drawString(self.left_margin, self.bottom_margin - 0.7 * cm, footer_text)

        canvas.restoreState()

    def export(self):
        df = self.load_excel()
        table = self.dataframe_to_table(df)

        doc = SimpleDocTemplate(
            self.pdf_file,
            pagesize=self.page_size,
            leftMargin=self.left_margin,
            rightMargin=self.right_margin,
            topMargin=self.top_margin,
            bottomMargin=self.bottom_margin,
        )

        doc.build([table], onFirstPage=self._header_footer, onLaterPages=self._header_footer)


def export_peptides_pdf(excel_file: str, pdf_file: str, patient_id: str,
                        sex: str = "male", landscape_mode: bool = True):
    ExcelToPDFExporter(
        excel_file=excel_file,
        pdf_file=pdf_file,
        patient_id=patient_id,
        male_female=sex,
        landscape_mode=landscape_mode,
    ).export()
