# report_generator_v1/figures/proteom_distribution.py
from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, Iterable, Mapping, Tuple, Optional

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import gaussian_kde

DEFAULT_MODELS = ("CKD", "CAD", "HF", "Oncorisk")
DEFAULT_DATA_FILES: Mapping[str, str] = {
    "CKD": "CKD_273.xlsx",
    "CAD": "CAD_238.xlsx",
    "HF": "HF2.xlsx",
    "Oncorisk": "Oncorisk_norm.xlsx",
}
DEFAULT_PREVALENCE: Mapping[str, Tuple[float, float]] = {m: (0.88, 0.12) for m in DEFAULT_MODELS}
OUTLIER_MODELS = {"CAD", "HF", "Oncorisk"}

def _round_axis(val: float, step: float, lower: bool) -> float:
    return (np.floor(val / step) if lower else np.ceil(val / step)) * step

def _remove_outliers(x: np.ndarray, lower_pct: int, upper_pct: int) -> np.ndarray:
    lo = np.percentile(x, lower_pct)
    hi = np.percentile(x, upper_pct)
    return x[(x >= lo) & (x <= hi)]

def _kde(y: np.ndarray, bw_method: float) -> gaussian_kde:
    if y.size < 2 or np.isclose(np.std(y), 0):
        y = y + np.random.normal(scale=1e-6, size=y.size if y.size else 2)
    return gaussian_kde(y, bw_method=bw_method)

def _plot_one(
    df: pd.DataFrame,
    *,
    prev_healthy: float,
    prev_unhealthy: float,
    patient_score: float,
    model: str,
    language: str,
    base_out_path: Path,     # path without suffix/number
    round_to: float,
    lower_pct: int,
    upper_pct: int,
    bw_method: float,
    number_suffix: Optional[int] = None,  # e.g. 7, 8, 9, 10
) -> Path:
    raw = np.concatenate([
        df.loc[df["group"] == 0, "score"].astype(float).to_numpy(),
        df.loc[df["group"] == 1, "score"].astype(float).to_numpy(),
    ])
    x_min = _round_axis(raw.min(), round_to, lower=True)
    x_max = _round_axis(raw.max(), round_to, lower=False)
    x_grid = np.linspace(x_min, x_max, 1000)

    h = df.loc[df["group"] == 0, "score"].astype(float).dropna().to_numpy()
    u = df.loc[df["group"] == 1, "score"].astype(float).dropna().to_numpy()
    removed = False
    if model in OUTLIER_MODELS:
        h = _remove_outliers(h, lower_pct, upper_pct)
        u = _remove_outliers(u, lower_pct, upper_pct)
        removed = True

    x_range = x_max - x_min
    tick_spacing = round_to
    if x_range / tick_spacing > 10:
        tick_spacing = round((x_range / 10) * 2) / 2 or round_to

    kde_h = _kde(h, bw_method)
    kde_u = _kde(u, bw_method)

    fig, ax = plt.subplots(figsize=(8, 6))
    label_h = "Gesund" if language == "DE" else "Healthy"
    label_u = model

    ax.plot(x_grid, kde_h(x_grid) * prev_healthy, color="green", lw=3, label=label_h)
    ax.plot(x_grid, kde_u(x_grid) * prev_unhealthy, color="red", lw=3, label=label_u)

    xlabel = "Proteom-Score" if language == "DE" else "Proteom Score"
    ylabel = "Häufigkeit" if language == "DE" else "Frequency"
    title  = "Verteilung der Proteom-Scores" if language == "DE" else "Distribution of Proteom Scores"

    ax.set_xlabel(xlabel, fontsize=16)
    ax.set_ylabel(ylabel, fontsize=16)
    ax.set_title(title, fontsize=18 if language == "DE" else 16)
    ax.set_xlim(x_min, x_max)
    ax.set_xticks(np.arange(x_min, x_max + tick_spacing, tick_spacing))
    ax.tick_params(labelsize=14)
    ax.grid(alpha=0.3)

    ax.axvline(x=patient_score, color="blue", linestyle="--", linewidth=3.5)
    ytop = ax.get_ylim()[1]
    text_y = ytop * (0.80 if language == "DE" else 0.85)
    text_label = "Aktueller Score" if language == "DE" else "Actual Score"
    ax.text(
        patient_score - 0.05 * (x_max - x_min),
        text_y,
        text_label,
        color="blue",
        rotation=90,
        ha="right",
        va="center",
        fontsize=14,
    )

    ax.legend(fontsize=14)
    fig.tight_layout()

    base_out_path.parent.mkdir(parents=True, exist_ok=True)

    # Assemble final filename:
    # e.g. CAD_distribution_plot_weighted_DE_noExtreme4_8.png
    suffix = f"_noExtreme{lower_pct}" if removed else ""
    stem = base_out_path.stem + suffix
    if number_suffix is not None:
        stem = f"{stem}_{number_suffix}"
    final_path = base_out_path.with_name(stem + base_out_path.suffix)

    fig.savefig(final_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    return final_path

def generate_all_proteom_distributions(
    *,
    work_dir: os.PathLike,                 # where Mustertabelle lives
    data_dir: os.PathLike,                 # where CKD.xlsx, CAD.xlsx, ...
    dest_dir: os.PathLike,                 # where numbered PNGs should be saved (-> your working_dir for add_figures)
    models: Iterable[str] = DEFAULT_MODELS,
    data_files: Mapping[str, str] = DEFAULT_DATA_FILES,
    prevalence: Mapping[str, Tuple[float, float]] = DEFAULT_PREVALENCE,
    languages: Iterable[str] = ("DE",),    # set from checkbox: ("DE",) or ("EN",)
    numbering: Mapping[str, int] | None = None,  # {"CKD":7,"CAD":8,"HF":9,"Oncorisk":10}
    bw_method: float = 0.5,
    round_to: float = 0.5,
    lower_pct: int = 4,
    upper_pct: int = 96,
) -> Dict[str, Path]:
    """
    Returns dict like {"CKD_DE": Path(...), "CAD_DE": Path(...)} with files saved as:
      <Model>_distribution_plot_weighted_<LANG>[_noExtreme<lower_pct>]_<N>.png
    """
    from ..excel_utils import load_scores

    work_dir = Path(work_dir)
    data_dir = Path(data_dir)
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    scores = load_scores(working_dir=str(work_dir))
    def _parse(s):
        try: return float(s)
        except: return None

    score_lookup = {
        "CKD": _parse(scores.get("CKD_score")),
        "CAD": _parse(scores.get("CAD_score")),
        "HF": _parse(scores.get("HF_score")),
        "Oncorisk": _parse(scores.get("Onkorisk_score") or scores.get("Oncorisk_score")),
    }

    # default numbering if none provided
    numbering = numbering or {"CKD": 7, "CAD": 8, "HF": 9, "Oncorisk": 10}
    results: Dict[str, Path] = {}

    languages = tuple(languages)  # guard against sets

    for model in models:
        xlsx_name = data_files.get(model)
        if not xlsx_name:
            continue
        df = pd.read_excel(data_dir / xlsx_name)
        if not {"group", "score"}.issubset(df.columns):
            raise ValueError(f"{xlsx_name} must contain 'group' and 'score' columns.")

        pscore = score_lookup.get(model)
        if pscore is None:
            continue

        prev_h, prev_u = prevalence.get(model, (0.88, 0.12))
        num = numbering.get(model)

        for lang in languages:
            base_name = f"{model}_distribution_plot_weighted_{lang}.png"
            base_out_path = dest_dir / base_name
            final = _plot_one(
                df,
                prev_healthy=prev_h,
                prev_unhealthy=prev_u,
                patient_score=pscore,
                model=model,
                language=lang,
                base_out_path=base_out_path,
                round_to=round_to,
                lower_pct=lower_pct,
                upper_pct=upper_pct,
                bw_method=bw_method,
                number_suffix=num,
            )
            results[f"{model}_{lang}"] = final

    return results
