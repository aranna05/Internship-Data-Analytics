# Marketing Campaign A/B/n Test: Did the New Promotion Really Drive More Sales?

Statistical analysis of a marketing experiment: three promotions were tested across 137 stores for four weeks. The project answers the Marketing Department's core question, *is the uplift real or just luck?*, using hypothesis tests, confidence intervals, multiple-comparison control, robustness checks and a rollout projection.

> **Dataset note:** the data file in this repo is the *Fast-Food Marketing Campaign A/B Test* dataset (`WA_Marketing-Campaign.csv`). Its outcome is **weekly sales ($K)**, not a binary purchase flag, so the headline analysis compares **mean sales** (Welch t-test / ANOVA) rather than conversion rates. A reusable **conversion-rate** toolkit (two-proportion z-test + chi-square) for the Kaggle *Marketing A/B Testing* dataset (ad vs. PSA) is included in `src/` and unit-tested.

## Key findings (95% confidence level, α = 0.05)

| Comparison | Lift in sales | Difference ($K / store / week), 95% CI | Holm-adjusted p | Verdict |
|---|---|---|---|---|
| Promo 1 vs Promo 2 | **+22.8%** (CI +8.7% to +38.4%) | +10.8 (+4.3, +17.2) | 0.004 | Significant |
| Promo 3 vs Promo 2 | **+17.0%** (CI +3.6% to +32.0%) | +8.0 (+1.7, +14.4) | 0.027 | Significant |
| Promo 1 vs Promo 3 | +4.9% (CI -6.5% to +18.3%) | +2.7 (-4.1, +9.5) | 0.426 | Not significant |

* The omnibus tests reject "all promotions perform equally" (ANOVA p = 0.0037, Kruskal-Wallis p = 0.0002, permutation test p = 0.002), so the differences are **not luck**.
* Promotions 1 and 3 each beat Promotion 2. The result holds after adjusting for market size and store age (OLS, robust SE) and when using all weekly data with store-clustered SEs.
* **Promo 1 vs 3 is inconclusive.** Power is only about 12% for a gap this small; about 550 stores per arm would be needed to confirm it.
* **Rollout projection (bonus):** rolling Promo 1 out to all 137 stores instead of Promo 2 implies about **+$1.48M sales per week** (95% CI $0.60M to $2.36M). For Promo 3 vs 2 it is about +$1.10M per week (CI $0.23M to $1.97M). Annualised figures are in `reports/rollout_projection.csv` and are a linear extrapolation of a 4-week test.

![Lift with confidence intervals](reports/figures/02_lift_confidence_intervals.png)
![Rollout projection](reports/figures/03_rollout_projection.png)

## Methodology

1. **Data quality checks:** no missing values or duplicates; each store receives exactly one promotion.
2. **Unit of analysis:** the *store*. Weekly rows from one store are correlated, so they are averaged into 137 independent observations. Testing the 548 rows directly would be pseudo-replication; the notebook shows it inflates significance (p = 4e-10 vs the valid 0.0013 for Promo 1 vs 2).
3. **Omnibus tests:** one-way ANOVA, Kruskal-Wallis and a permutation test.
4. **Pairwise tests:** Welch's t-test (plus Mann-Whitney as a check), **Holm correction** for 3 comparisons, Cohen's d.
5. **Lift and confidence intervals:** percentile bootstrap (10,000 resamples) for relative lift; Welch CI for absolute difference.
6. **Robustness:** covariate-adjusted OLS (HC3), cluster-robust OLS on weekly data.
7. **Business impact:** incremental chain-wide sales for a 100% rollout, with CIs.
8. **Power analysis:** explains why Promo 1 vs 3 is inconclusive.

## Repository structure

```
.
├── data/WA_Marketing-Campaign.csv           # source data (548 rows = 137 stores x 4 weeks)
├── notebooks/marketing_ab_test_analysis.ipynb   # full executed analysis
├── src/conversion_ab_test.py                # z-test / chi-square toolkit for binary conversion data
├── tests/test_conversion_ab_test.py         # unit tests (pytest)
├── reports/
│   ├── marketing_ab_test_analysis.html      # notebook exported to HTML
│   ├── pairwise_results.csv  rollout_projection.csv  descriptive_stats.csv
│   └── figures/                             # PNG charts
├── requirements.txt
└── README.md
```

## Getting started

```bash
python -m venv .venv && source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
jupyter notebook notebooks/marketing_ab_test_analysis.ipynb
pytest -q                                              # run unit tests
```

### Conversion-rate version (Kaggle *Marketing A/B Testing*)

Download `marketing_AB.csv` from Kaggle (columns `test group` = `ad`/`psa`, `converted` = True/False), then:

```bash
python src/conversion_ab_test.py path/to/marketing_AB.csv --rollout-users 1000000
```

It prints conversion rates, absolute and relative lift with 95% CIs, z-test and chi-square p-values, and the incremental conversions if the new ad reaches the given number of users. The module is verified on synthetic data in the Kaggle schema only; it has not been run on the real Kaggle file.

## Limitations

* The test lasted 4 weeks, so novelty effects and seasonality are not captured.
* Stores were not shown to be randomly assigned within market size; adjusted models mitigate but cannot fully remove this.
* Sales is a proxy for customer purchases; customer counts are not in the data, so impact is in sales dollars.
* The annualised rollout figures assume the effect persists and scales linearly.

## Tech stack

Python, pandas, NumPy, SciPy, statsmodels, matplotlib, Jupyter.
