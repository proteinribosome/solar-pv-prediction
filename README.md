Physics Features and Generalization in Solar Power Prediction

Research question
Do physics-derived features improve ML power predictions on PV systems not seen during training, and by how much?
Why it matters
New solar installations have no historical data, so a useful model must work on systems it wasn't trained on (the "cold-start" problem). The Santos (2025) PhD thesis explicitly names the generalizability of hybrid models as an open question — it claims Physics-Informed ML has "limited" generalization ability but never measures this, and its own experiments were confined to a single site. You're testing an untested claim from the literature, not repeating a settled one.
Dataset
The Hong Kong (HKUST) rooftop PV dataset (Lin et al., Scientific Data, 2025): 60 grid-connected rooftop systems, 2021–2023, 5-minute power data plus on-site weather at 1-minute resolution, with system metadata (capacity, hardware). Chosen because many diverse systems under one weather source make a clean cross-system test possible.
Experimental design
Train two models that differ only in their input features:

* Model A (baseline): weather features only — irradiance, temperature, humidity, wind, etc.
* Model B (hybrid): the same weather features plus physics features computed with `pvlib` — solar zenith/azimuth angle, clear-sky irradiance, clear-sky index, plane-of-array irradiance, estimated cell temperature.
Also run two reference baselines: a trivial predictor (mean/persistence) and a pure-physics model (PVWatts via pvlib), so the final comparison is trivial vs. physics-only vs. ML-only vs. hybrid.
Model choice: XGBoost or Random Forest (best on tabular PV data per the thesis), identical architecture and hyperparameters for A and B.
Evaluation — the core of the whole project

* Split by system ID, not by time or at random: train on ~50 systems, test on ~10 held-out systems the model has never seen. Ideally rotate (leave-systems-out) and average.
* Normalize the target and errors by each system's capacity (kW/kWp) so different-sized systems are comparable.
* Report nRMSE, MAE, and MBE per held-out system.
* The headline result is one number: the error gap between Model A and Model B on unseen systems.
Leakage safeguards
No test-system data anywhere in training; filter to daylight hours (zenith < 90°) rather than padding nights with zeros, which artificially inflates metrics; keep preprocessing fit on training data only.
Limitations (state them upfront)
All 60 systems share one campus, one climate, and one weather station — so this tests cross-system, not cross-climate, generalization. Three years of data; campus-scale rooftops only.
Reproducibility
Public GitHub repo, pinned library versions, fixed random seeds, documented preprocessing.
Stretch goals (only if time allows)

1. Compare hybridization strategies on unseen systems: physics-as-features (Physics-Informed) vs. physics-model-plus-ML-error-correction (Physics-Guided). The thesis predicts the latter generalizes better — testing that prediction head-to-head is a bonus contribution.
2. True cross-location test: train on Hong Kong, test zero-shot on a different-climate public dataset (DKASC Alice Springs or NIST).
What you'll be able to claim at the end
"Physics-derived features reduce prediction error on unseen rooftop PV systems by X% relative to a weather-only model, and the hybrid outperforms/underperforms pure physics by Y%." Any value of X and Y — including zero — is a legitimate finding, because nobody has published this number for individual rooftop systems.
That's the plan. If you want, the next concrete step is small: download the dataset, load one system's CSV, and get pvlib computing clear-sky irradiance for Hong Kong's coordinates — that single script touches every hard part of the pipeline early.