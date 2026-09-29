# Weather Intelligence – Experimentation Repository

**AT2 – Machine Learning as a Service** · 36120 Advanced Machine Learning Applications
**Student:** Tahreema Rahman Zariyat · **Student ID:** 26016360

Two forecasting services for Sydney, NSW (lat −33.8688, lon 151.2093), trained on the
[Open-Meteo Historical Weather API](https://open-meteo.com/en/docs/historical-weather-api):

| Target | Type | Horizon | Notebook | Artefacts |
| --- | --- | --- | --- | --- |
| Climate Comfort Index (CCI, 0–100) | regression (XGBoost, one model per horizon) | t+1, t+2, t+3 | `notebooks/comfort_climate/cci_xgboost_experiment.ipynb` | `models/comfort_climate/` |
| Weather Hazard Category (Low / Moderate / High / Extreme) | 4-class classification (XGBoost, class-weighted) | exactly t+7 | `notebooks/weather_hazard/whc_xgboost_experiment.ipynb` | `models/weather_hazard/` |

All reusable code (data download, daily aggregation, the target formulas, feature
engineering, model helpers, artefact bundles) lives in my custom Python package
**`amla-package-26016360` ≥ 0.2.0**, published on TestPyPI and shared with the
[API repository](https://github.com/tahreemarz/weather-intelligence-api). The notebooks
and the FastAPI service therefore compute targets and features with the *same* code.

## Repository structure (Cookiecutter Data Science layout)

```
├── README.md
├── pyproject.toml            <- project metadata and dependencies (custom package from TestPyPI)
├── requirements.txt          <- same dependencies, pip -r form
├── data
│   ├── raw                   <- cached hourly Open-Meteo download (parquet, created by the notebooks)
│   ├── interim
│   └── processed             <- X_train / X_val / X_test / y_* CSVs + frozen features_list.json per target
│       ├── comfort_climate
│       └── weather_hazard
├── models
│   ├── comfort_climate       <- cci_h1.json, cci_h2.json, cci_h3.json, climatology.json, metadata.json
│   └── weather_hazard        <- whc_h7.json, climatology.json, metadata.json
├── notebooks
│   ├── comfort_climate       <- CCI experiment (template.ipynb structure)
│   └── weather_hazard        <- WHC experiment (template.ipynb structure)
├── reports
│   └── figures               <- every figure produced by the notebooks (png)
├── references                <- assignment brief, APA guide, links
└── src
    └── check_environment.py  <- sanity check: package version, utstd, Open-Meteo reachability
```

## Setup

Requirements: Python 3.10 – 3.12, internet access (Open-Meteo, TestPyPI, PyPI).

```bash
git clone https://github.com/tahreemarz/weather-intelligence-experiments.git
cd weather-intelligence-experiments
python -m venv .venv && source .venv/bin/activate      # Windows: .venv\Scripts\activate

# notebook dependencies (PyPI)
pip install -e ".[notebooks]"         # or: pip install -r requirements.txt
# custom package from TestPyPI (--no-deps: its dependencies were just installed from PyPI)
pip install --no-deps --index-url https://test.pypi.org/simple/ "amla-package-26016360>=0.2.0"

python src/check_environment.py       # should print "environment OK"
```

`utstd` (the `print_tile` helper used by the notebook template) is installed by the
`notebooks` extra.

## Running the experiments

```bash
jupyter lab
```

1. Open `notebooks/comfort_climate/cci_xgboost_experiment.ipynb` and run all cells.
   The first run downloads hourly data for 1990-01-01 → 2025-12-31 (about 36 years,
   ~316k rows, a few minutes on a normal connection) and caches it in
   `data/raw/sydney_hourly_1990-01-01_2025-12-31.parquet`; later runs and the second
   notebook reuse the cache.
2. Open `notebooks/weather_hazard/whc_xgboost_experiment.ipynb` and run all cells.

Each notebook is self-contained and, in order: describes the business case, downloads
and audits the data (`print_data_report` from the package), builds the target with the
package formulas, explores the target and the features of interest, selects features
(filter + embedded/permutation), cleans the data and engineers features with the package
functions, splits chronologically (`split_by_date`: train ≤ 2019, validation 2020–2022,
test 2023–2025), fits a day-of-year climatology on the training years only, runs a
random hyper-parameter search with early stopping, evaluates against persistence /
climatology / majority baselines, analyses the business impact, refits on all data up
to 2025-12-31 and saves the artefact bundle to `models/<target>/`.

* Every explanation cell is an f-string that fills itself with the numbers computed
  above it, so the narrative always matches the run.
* `QUICK_RUN = True` in the package cell runs a reduced search (a smoke test, ~2 min);
  the full run (`QUICK_RUN = False`, default) takes roughly 10–20 min per notebook on a
  laptop.
* Runs are deterministic (`RANDOM_STATE = 42`, XGBoost seed), except for minor
  floating-point differences across CPUs.

## Model artefacts

Models are saved in XGBoost's portable JSON format (no pickles) together with the
fitted climatology table and a `metadata.json` that holds the frozen feature list,
hyper-parameters, feature importance, validation and test metrics, baselines and the
training window — everything the API's `/model-metadata` endpoint publishes.

After running both notebooks, copy the bundles into the API repository:

```bash
cp -r models/comfort_climate models/weather_hazard ../weather-intelligence-api/models/
```

## Final report

`reports/build_report.py` fills the unit's report template (`reports/report_template.docx`)
with the results of the two notebooks: the cover page details and repository links, and
every section's text, tables and figures, all taken from `models/*/metadata.json` and
`reports/figures/`. Sentences that depend on the results (improvement over baselines,
hypothesis outcome, top features, business counts) are generated from the metadata.

```bash
python reports/build_report.py      # after running both notebooks -> reports/AT2_report.docx
```

Before running it, set your GitHub username, repository links and the Render URL at the
top of the script. Open the result in Word and answer "Yes" to updating fields so the
table of contents shows the right page numbers; then read the generated text once. The
script prints the body word count (limit 4,000).

## The custom package

`amla-package-26016360` started as a generic toolkit (dataset splitting, data quality
checks, performance reporting) in AT1. Version 0.2.0 extends it for this project with
the `weather` subpackage, `split_by_date`, `multiclass_scores` and tolerance-aware
`regression_scores`. Source, tests and the publishing workflow (`python -m build`,
`twine upload --repository testpypi`) are in the package repository:
<https://github.com/tahreemarz/36120-26SP-group2-26016360-package>.

Any change to shared logic (formulas, features, cleaning) must be made in the package,
version-bumped and published to TestPyPI, then pinned here **and** in the API repository.

## Data notes

* Source: Open-Meteo Historical Weather API (ERA5 reanalysis), hourly, timezone
  `Australia/Sydney`; 13 hourly variables (7 required by the target formulas + 6 extra
  predictors). Data from 2026-01-01 onwards is never downloaded (production data).
* Daily aggregation follows the assignment exactly: mean for temperature, humidity,
  wind speed and cloud cover; daily total for precipitation and snowfall; daily maximum
  for wind gusts. It is implemented once in `amla_package_26016360.weather.aggregate_daily`.
