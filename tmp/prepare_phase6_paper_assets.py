from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import nbformat
import numpy as np
import pandas as pd
from nbformat.v4 import new_code_cell, new_markdown_cell


ROOT = Path(__file__).resolve().parents[1]
RESULTS_DIR = ROOT / "future_forecasting_experiment" / "results"
SYSTEM_RESULTS_PATH = RESULTS_DIR / "phase6_system_results.csv"
SUMMARY_PATH = RESULTS_DIR / "phase6_summary.csv"
PAIRED_PATH = RESULTS_DIR / "phase6_paired_comparison.csv"
FIGURE_PATH = RESULTS_DIR / "phase6_main_comparison_paper.png"
NOTEBOOK_PATH = (
    ROOT
    / "future_forecasting_experiment"
    / "code_future_forecasting.ipynb"
)

MODEL_ORDER = [
    "training_mean",
    "time_solar_only",
    "model_a_weather",
    "model_b_hybrid",
]
MODEL_LABELS = {
    "training_mean": "Training mean",
    "time_solar_only": "Time/solar only",
    "model_a_weather": "Model A: weather",
    "model_b_hybrid": "Model B: hybrid",
}
COLD_START_SYSTEMS = {
    "SQ_Block_P",
    "Shaw_Auditorium",
    "UG_Hall8",
    "UG_Hall9",
}


def validate_inputs(system_results, summary):
    required_system = {
        "model",
        "fold",
        "station",
        "n_test_rows",
        "nRMSE",
        "MAE",
        "MBE",
    }
    required_summary = {
        "model",
        "label",
        "nRMSE_mean",
        "nRMSE_std",
        "nRMSE_median",
        "MAE_mean",
        "MAE_std",
        "MAE_median",
        "MBE_mean",
        "MBE_std",
        "MBE_median",
    }
    if missing := sorted(required_system - set(system_results.columns)):
        raise KeyError(f"Missing system-result columns: {missing}")
    if missing := sorted(required_summary - set(summary.columns)):
        raise KeyError(f"Missing summary columns: {missing}")
    if system_results.isna().any().any() or summary.isna().any().any():
        raise ValueError("Phase 6 results contain missing values")
    if system_results.duplicated(["model", "station"]).any():
        raise ValueError("Duplicate model-station results found")
    model_counts = system_results.groupby("model")["station"].nunique()
    if set(model_counts.index) != set(MODEL_ORDER):
        raise ValueError("Phase 6 results do not contain the four expected models")
    if not model_counts.eq(37).all():
        raise ValueError("Every model must contain 37 station results")

    fold_pairs = system_results.pivot(
        index="station", columns="model", values="fold"
    ).loc[:, MODEL_ORDER]
    row_pairs = system_results.pivot(
        index="station", columns="model", values="n_test_rows"
    ).loc[:, MODEL_ORDER]
    if not fold_pairs.eq(fold_pairs.iloc[:, 0], axis=0).all().all():
        raise ValueError("Fold assignments differ between models")
    if not row_pairs.eq(row_pairs.iloc[:, 0], axis=0).all().all():
        raise ValueError("Test-row counts differ between models")

    recomputed = system_results.groupby("model")[["nRMSE", "MAE", "MBE"]].agg(
        ["mean", "std", "median"]
    )
    for _, row in summary.iterrows():
        for metric in ("nRMSE", "MAE", "MBE"):
            for statistic in ("mean", "std", "median"):
                expected = recomputed.loc[row["model"], (metric, statistic)]
                actual = row[f"{metric}_{statistic}"]
                if not np.isclose(actual, expected, rtol=1e-12, atol=1e-14):
                    raise ValueError(
                        f"Summary mismatch for {row['model']} "
                        f"{metric}_{statistic}"
                    )


def paired_bootstrap(system_results, metric, n_bootstrap=10_000, seed=42):
    paired = system_results.pivot(
        index="station", columns="model", values=metric
    ).loc[:, ["model_a_weather", "model_b_hybrid"]]
    reference = paired["model_a_weather"].to_numpy(dtype=float)
    candidate = paired["model_b_hybrid"].to_numpy(dtype=float)
    difference = candidate - reference

    rng = np.random.default_rng(seed)
    indices = rng.integers(
        0, len(paired), size=(n_bootstrap, len(paired))
    )
    reference_bootstrap = reference[indices].mean(axis=1)
    candidate_bootstrap = candidate[indices].mean(axis=1)
    difference_bootstrap = candidate_bootstrap - reference_bootstrap

    result = {
        "metric": metric,
        "reference_model": "model_a_weather",
        "candidate_model": "model_b_hybrid",
        "n_systems": len(paired),
        "reference_macro_mean": reference.mean(),
        "candidate_macro_mean": candidate.mean(),
        "candidate_minus_reference": difference.mean(),
        "difference_ci95_low": np.quantile(
            difference_bootstrap, 0.025
        ),
        "difference_ci95_high": np.quantile(
            difference_bootstrap, 0.975
        ),
        "candidate_better_systems": int((candidate < reference).sum()),
        "ties": int((candidate == reference).sum()),
        "relative_error_reduction_pct": np.nan,
        "relative_reduction_ci95_low": np.nan,
        "relative_reduction_ci95_high": np.nan,
    }
    if metric in {"nRMSE", "MAE"}:
        relative_bootstrap = (
            (reference_bootstrap - candidate_bootstrap)
            / reference_bootstrap
            * 100.0
        )
        result.update(
            {
                "relative_error_reduction_pct": (
                    (reference.mean() - candidate.mean())
                    / reference.mean()
                    * 100.0
                ),
                "relative_reduction_ci95_low": np.quantile(
                    relative_bootstrap, 0.025
                ),
                "relative_reduction_ci95_high": np.quantile(
                    relative_bootstrap, 0.975
                ),
            }
        )
    return result


def build_figure(system_results, summary):
    summary_plot = (
        summary.set_index("model").loc[MODEL_ORDER].reset_index()
    )
    paired = system_results.pivot(
        index="station", columns="model", values="nRMSE"
    ).loc[:, ["model_a_weather", "model_b_hybrid"]]
    paired["cold_start"] = paired.index.isin(COLD_START_SYSTEMS)

    colors = ["#AEB4BB", "#8C6BBE", "#3478B8", "#D9693A"]
    edge = "#27313A"
    grid = "#D9DEE3"

    fig, axes = plt.subplots(1, 2, figsize=(13.8, 5.15))

    bars = axes[0].barh(
        summary_plot["label"],
        summary_plot["nRMSE_mean"],
        color=colors,
        edgecolor=edge,
        linewidth=0.8,
    )
    axes[0].invert_yaxis()
    axes[0].set_xlim(0, summary_plot["nRMSE_mean"].max() * 1.18)
    axes[0].set_xlabel("Macro nRMSE")
    axes[0].set_title("Macro nRMSE by model", loc="left", fontweight="bold")
    axes[0].bar_label(bars, fmt="%.4f", padding=4, fontsize=9)
    axes[0].grid(axis="x", color=grid, linewidth=0.8)
    axes[0].set_axisbelow(True)

    regular = ~paired["cold_start"]
    axes[1].scatter(
        paired.loc[regular, "model_a_weather"],
        paired.loc[regular, "model_b_hybrid"],
        s=42,
        color="#3478B8",
        edgecolor=edge,
        linewidth=0.55,
        label="Held-out system",
    )
    axes[1].scatter(
        paired.loc[~regular, "model_a_weather"],
        paired.loc[~regular, "model_b_hybrid"],
        s=64,
        facecolor="white",
        edgecolor="#D9693A",
        linewidth=1.6,
        label="No pre-2023 history",
    )
    limit = 1.06 * max(
        paired["model_a_weather"].max(),
        paired["model_b_hybrid"].max(),
    )
    axes[1].plot(
        [0, limit],
        [0, limit],
        color=edge,
        linestyle="--",
        linewidth=1.0,
        label="Equal error",
    )
    axes[1].set_xlim(0, limit)
    axes[1].set_ylim(0, limit)
    axes[1].set_aspect("equal", adjustable="box")
    axes[1].set_xlabel("Weather-only nRMSE")
    axes[1].set_ylabel("Hybrid nRMSE")
    axes[1].set_title(
        "Station-level paired comparison",
        loc="left",
        fontweight="bold",
    )
    axes[1].grid(color="#E4E8EC", linewidth=0.7)
    axes[1].set_axisbelow(True)
    axes[1].legend(frameon=False, fontsize=8, loc="upper left")
    axes[1].text(
        0.98,
        0.04,
        "Hybrid lower on 37 of 37 systems",
        transform=axes[1].transAxes,
        ha="right",
        va="bottom",
        fontsize=8.5,
        color=edge,
    )

    for axis in axes:
        axis.spines[["top", "right"]].set_visible(False)
        axis.tick_params(labelsize=9)

    fig.suptitle(
        "Future-time prediction on unseen PV systems",
        x=0.06,
        ha="left",
        fontsize=15,
        fontweight="bold",
    )
    fig.text(
        0.06,
        0.925,
        "Train through 2022; test on held-out systems in 2023; "
        "n = 37 systems; lower is better",
        ha="left",
        fontsize=9.5,
        color="#5F6368",
    )
    fig.tight_layout(rect=(0.03, 0.03, 0.99, 0.89))
    fig.savefig(FIGURE_PATH, dpi=240, bbox_inches="tight")
    plt.close(fig)


def update_notebook():
    notebook = nbformat.read(NOTEBOOK_PATH, as_version=4)
    notebook.cells = [
        cell
        for cell in notebook.cells
        if cell.get("id")
        not in {"phase6-paired-heading", "phase6-paired-comparison"}
    ]
    run_index = next(
        index
        for index, cell in enumerate(notebook.cells)
        if cell.get("id") == "phase6-run"
    )

    heading = new_markdown_cell(
        """## Paired station-level uncertainty

This analysis reloads the saved station-level CSV when necessary, so it can be
rerun without refitting XGBoost. It resamples the 37 held-out stations, rather
than dependent timestamp rows, and exports the paired confidence intervals used
in the Phase 6 paper update."""
    )
    heading["id"] = "phase6-paired-heading"

    code = new_code_cell(
        """PHASE6_BOOTSTRAP_SAMPLES = 10_000

if "phase6_system_results" not in globals():
    phase6_system_results = pd.read_csv(
        EXPERIMENT_ROOT / "results" / "phase6_system_results.csv"
    )


def phase6_paired_bootstrap(
    system_results,
    metric,
    *,
    n_bootstrap=PHASE6_BOOTSTRAP_SAMPLES,
    random_state=PHASE6_RANDOM_STATE,
):
    paired = system_results.pivot(
        index=GROUP_COL,
        columns="model",
        values=metric,
    ).loc[:, ["model_a_weather", "model_b_hybrid"]]
    reference = paired["model_a_weather"].to_numpy(dtype=float)
    candidate = paired["model_b_hybrid"].to_numpy(dtype=float)

    rng = np.random.default_rng(random_state)
    indices = rng.integers(
        0,
        len(paired),
        size=(n_bootstrap, len(paired)),
    )
    reference_bootstrap = reference[indices].mean(axis=1)
    candidate_bootstrap = candidate[indices].mean(axis=1)
    difference_bootstrap = candidate_bootstrap - reference_bootstrap

    result = {
        "metric": metric,
        "reference_model": "model_a_weather",
        "candidate_model": "model_b_hybrid",
        "n_systems": len(paired),
        "reference_macro_mean": reference.mean(),
        "candidate_macro_mean": candidate.mean(),
        "candidate_minus_reference": (candidate - reference).mean(),
        "difference_ci95_low": np.quantile(
            difference_bootstrap, 0.025
        ),
        "difference_ci95_high": np.quantile(
            difference_bootstrap, 0.975
        ),
        "candidate_better_systems": int(
            (candidate < reference).sum()
        ),
        "ties": int((candidate == reference).sum()),
        "relative_error_reduction_pct": np.nan,
        "relative_reduction_ci95_low": np.nan,
        "relative_reduction_ci95_high": np.nan,
    }
    if metric in {"nRMSE", "MAE"}:
        relative_bootstrap = (
            (reference_bootstrap - candidate_bootstrap)
            / reference_bootstrap
            * 100.0
        )
        result.update({
            "relative_error_reduction_pct": (
                (reference.mean() - candidate.mean())
                / reference.mean()
                * 100.0
            ),
            "relative_reduction_ci95_low": np.quantile(
                relative_bootstrap, 0.025
            ),
            "relative_reduction_ci95_high": np.quantile(
                relative_bootstrap, 0.975
            ),
        })
    return result


phase6_paired_comparison = pd.DataFrame([
    phase6_paired_bootstrap(phase6_system_results, metric)
    for metric in ["nRMSE", "MAE", "MBE"]
])
phase6_paired_comparison.to_csv(
    EXPERIMENT_ROOT / "results" / "phase6_paired_comparison.csv",
    index=False,
)
display(phase6_paired_comparison.round(4))"""
    )
    code["id"] = "phase6-paired-comparison"
    code["execution_count"] = None
    code["outputs"] = []

    notebook.cells[run_index + 1 : run_index + 1] = [heading, code]
    nbformat.validate(notebook)
    nbformat.write(notebook, NOTEBOOK_PATH)


def main():
    system_results = pd.read_csv(SYSTEM_RESULTS_PATH)
    summary = pd.read_csv(SUMMARY_PATH)
    validate_inputs(system_results, summary)

    paired = pd.DataFrame(
        [
            paired_bootstrap(system_results, metric)
            for metric in ("nRMSE", "MAE", "MBE")
        ]
    )
    paired.to_csv(PAIRED_PATH, index=False)
    build_figure(system_results, summary)
    update_notebook()

    print(PAIRED_PATH)
    print(FIGURE_PATH)
    print(NOTEBOOK_PATH)


if __name__ == "__main__":
    main()
