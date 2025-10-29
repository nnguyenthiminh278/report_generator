# 🧬 Report Generator

A **GUI-based report generator** for patient data that:
- Reads patient information from a local SQLite database.
- Reads analysis scores from Excel files.
- Automatically fills Word report templates (`.docx`).
- Inserts figures dynamically (with custom overlays).
- Generates both **Word** (`.docx`) and **PDF** reports.
- Supports both **Linux** and **Windows** environments.

---

## 📦 Features

- 🧠 Intelligent template filling using [python-docx](https://python-docx.readthedocs.io/) and [docxtpl](https://docxtpl.readthedocs.io/)
- 📊 Excel data extraction using [pandas](https://pandas.pydata.org/)
- 🖼️ Dynamic figure insertion with optional text overlays using [Pillow](https://pillow.readthedocs.io/)
- 🪟 Cross-platform GUI built with Tkinter
- 🧾 PDF export:
  - Uses **ONLYOFFICE** (preferred, best quality)
  - Falls back to **LibreOffice**
  - Uses **docx2pdf** on Windows

---

## 🧰 System Requirements

### Linux
- Python **3.9+**
- **LibreOffice** (for PDF export fallback)
- (Optional but recommended) **ONLYOFFICE Desktop Editors**
- Tkinter (usually included with Python)
- Pandoc (optional fallback converter)

### Windows
- Python **3.9+**
- Microsoft Word (for `docx2pdf` conversion)
- Tkinter (included in standard Python)

---

## 🚀 Installation

### 🐧 On Linux

1. **Clone the repository**
   ```bash
   git clone https://github.com/<your-username>/report_generator.git
   cd report_generator
2. **Set up the virtual environment**
    ```bash
    chmod +x setup_venv.sh
    ./setup_venv.sh
3. **Activate the environment**
    ```bash
    source myenv/bin/activate
4. **Install the package**
    ```bash
    pip install -e .
5. **Run the software**
    ```bash
    report-generator

🪟 On Windows
1. **Install Python 3.9+**
Ensure “Add Python to PATH” is checked during installation.
2. **Clone or extract the project**
3. **Create a virtual environment**
    ```bat
    python -m venv myenv
    myenv\Scripts\activate
4. **Install dependencies**
    ```bat  
    pip install -r requirements.txt
5. **Install the package in editable mode**
    ```bat
    pip install -e .
6. **Run the program**
    ```bat
    report-generator

#### 📂 Project Structure
report_generator/
│
├── setup.py
├── setup_venv.sh
├── requirements.txt
├── README.md
│
├── report_generator_v1/
│   ├── __init__.py
│   ├── main.py                 # GUI entry point
│   ├── db_utils.py             # Database logic
│   ├── excel_utils.py          # Excel data reading
│   ├── report_utils.py         # Word, figure, PDF helpers
│   ├── context_utils.py        # Template context creation
│   ├── peptide_processor.py    # Patient's Peptide List creation
│   ├── peptide_pdf.py          # Patient's Peptide List conversion (to pdf)
│   └── templates/
│   |   ├── template_MOS.docx   # Main report template
│   |   └── template_annex.docx # Annex template
│   └── data/                   # CKD/CAD/HF/Oncorisk score files
│       ├── CKD_273.xlsx
│       ├── CAD_238.xlsx
│       ├── HF2.xlsx
│       └── Oncorisk_norm.xlsx
│       ├── Healthy_patients_female_n836_log2_stats_nonNeg.xlsx
│       ├── Healthy_patients_male_n917_log2_stats_nonNeg.xlsx
│       └── ML2.5_21082025_seq.xls
│
└── tests/
    └── test_basic.py

##### 🧠 Usage Overview
1️⃣ Launch the GUI

Run:

report-generator

2️⃣ Select a Patient

Choose how to search: Patient ID, Sample ID, or Name

Enter the value and click Generate Report

3️⃣ Choose a Working Directory

Select the folder containing your patient’s Excel data and figures.

4️⃣ Generate Report

The app fills your Word template with:

Patient info (from DB)

Analysis scores (from Excel)

Figures (from working directory)

Automatically generates both .docx and .pdf files.

5️⃣ Generate Annex

Click Generate Annex to create an annex document using the second template.

###### 🧪 Troubleshooting
❌ PackageNotFoundError: template_MOS.docx

- Make sure your templates are included in:

report_generator_v1/templates/

and referenced with:

importlib.resources.path("report_generator_v1.templates", "template_MOS.docx")

- Make sure your score of each model are included in:

report_generator_v1/data/

- Make sure your patient data (in excel, txt) is include in the selected dir by the user e.g. input_patient_1234