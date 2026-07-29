# Physics-Derived Features for Power Prediction on Unseen Rooftop PV Systems

## Paper
[Read the paper here](https://google.com)

## Abstract
Physics-derived features were most useful in this study as a compact representation rather than as a large source of incremental information. The analysis used 1,370,191 modeling-ready daylight observations from 37 rooftop PV systems. In the primary system- and time-disjoint experiment, models trained through 2022 on other systems were tested on held-out systems in 2023. Weather-plus-physics Model B (W+P) achieved the lowest macro nRMSE at 11.14% of installed capacity, compared with 11.48% for weather-plus-time-plus-physics Model D (W+T+P), 11.56% for weather-plus-time Model C (W+T), and 13.30% for weather-only Model A (W). Model B outperformed Model D on 30 of 37 systems by nRMSE (one-sided exact sign test p = 0.00010) and on 32 of 37 by MAE (p < 0.0001). The pre-specified incremental test, D versus C, reduced nRMSE by only 0.08 percentage points of capacity (0.68% relative); its nRMSE direction was not significant by sign test (23/37, p = 0.094), although its MAE direction was significant (29/37, p = 0.0004). These results favor a representation-efficiency interpretation within this campus dataset and fixed XGBoost configuration.

## Final Version Code
[Link to notebook](https://github.com/proteinribosome/solar-pv-prediction/blob/main/v4%20model%20d/code_model_d.ipynb)
