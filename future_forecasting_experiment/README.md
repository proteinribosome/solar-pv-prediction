# Future forecasting experiment

This folder is an isolated copy for testing future-time prediction on unseen
PV systems. The original project files are unchanged.

Contents:

- `code_future_forecasting.ipynb`: original notebook plus an unexecuted Phase 6.
- `dataset/site_level_dataset_modified_combined.csv`: copied source dataset.
- `results/`: Phase 6 output location.

Phase 6 uses data through 2022 from the training stations and evaluates 2023
data from completely held-out stations. It includes a training-mean baseline,
a time/solar-only diagnostic, the weather-only XGBoost model, and the hybrid
XGBoost model.

All 37 systems have eligible 2023 test rows. Four systems have no eligible
pre-2023 rows; Phase 6 reports them explicitly and treats them as valid
cold-start test systems when their fold is evaluated.

Training is disabled by default. First run the original Phase 1 preparation,
the Phase 2 metric-definition cell, and the Phase 4 XGBoost
import/configuration cell. Then run Phase 6, inspect the split audit, and set
`RUN_PHASE6 = True` only when ready.

The experiment uses measured weather at the 2023 prediction timestamps. It is
therefore future-time power prediction conditional on observed weather, not a
day-ahead forecast based on weather forecasts.
