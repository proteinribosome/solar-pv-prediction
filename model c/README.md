# Model C: weather plus ordinary time controls

This folder is an isolated copy of the future-time experiment. It tests whether
the physics-enhanced Model B still outperforms a weather model after ordinary
clock and calendar information is supplied.

The original `future_forecasting_experiment` folder is not modified.

## Contents

- `code_model_c.ipynb`: Colab-ready notebook with Model C added to Phase 6.
- `dataset/site_level_dataset_modified_combined.csv`: copied source dataset.
- `results/`: output directory populated only after training is enabled.
- `requirements.txt`: copied Python requirements.

## Model C features

Model C uses the four Model A weather variables:

- measured GHI
- ambient temperature
- wind speed
- relative humidity

It adds four naive cyclic time controls:

- sine and cosine of local hour of day
- sine and cosine of local day of year

It does not use solar zenith, solar azimuth, clear-sky GHI, clear-sky index, or
estimated cell temperature.

## Run in Colab

1. Open `code_model_c.ipynb` in Colab.
2. Make the source CSV available. The notebook first checks
   `model c/dataset/site_level_dataset_modified_combined.csv`; in a normal
   Colab session, upload the CSV as
   `/content/site_level_dataset_modified_combined.csv`.
3. Run the original data-preparation cells through Phase 1 validation.
4. Run the Phase 2 metric-definition cell.
5. Run the Phase 4 XGBoost import/configuration cell.
6. Run the Phase 6 configuration, split-audit, and evaluator cells.
7. Confirm that the split audit passes.
8. Change `RUN_PHASE6 = False` to `RUN_PHASE6 = True`.
9. Run the Phase 6 training, paired-comparison, and plot cells.

Training is disabled in the delivered notebook. Enabling it runs four XGBoost
feature sets over five folds, for 20 XGBoost fits.

## Main outputs

- `results/phase6_summary.csv`: macro metrics for all five models.
- `results/phase6_system_results.csv`: held-out-system metrics.
- `results/phase6_fold_results.csv`: fold-level metrics.
- `results/phase6_paired_comparison.csv`: Model B versus Model A, Model C versus
  Model A, and Model B versus Model C paired comparisons.
- `results/model_c_future_comparison.png`: summary chart.

The key test is `hybrid_vs_time_control` in
`phase6_paired_comparison.csv`. If Model C approaches Model B, ordinary time
context explains much of the original gain. If Model C remains near Model A,
the physics-specific interpretation is substantially stronger.
