import os
import pandas as pd

def load_scores(working_dir, filename="Mustertabelle_Klassifikation_alles.xlsx"):
    excel_path = os.path.join(working_dir, filename)
    df = pd.read_excel(excel_path, header=None)

    headers = df.iloc[2].tolist()
    final_row = df[df.iloc[:, 1].astype(str).str.lower().str.contains("final score")]
    if final_row.empty:
        raise ValueError("No 'final score' row found in Excel")

    scores_dict = dict(zip(headers, final_row.iloc[0].tolist()))

    mapping = {
        "CAD238ML1k.mdl": "CAD_score",
        "CKD273ML1hybrid": "CKD_score",
        "HF2_ML1new.mdl": "HF_score",
        "oncoRisk normo": "Onkorisk_score",
        "BioAge": "BioAge_value",
        "LifeSpeed":"LifeSpeed_value"
    }

    return {
        v: f"{scores_dict[k]:.3f}"
        for k, v in mapping.items()
        if k in scores_dict
    }
