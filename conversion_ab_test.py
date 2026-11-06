"""Reusable conversion-rate A/B test utilities (two-proportion z-test, chi-square,
confidence intervals for absolute / relative lift, and rollout projection).

Designed for binary-outcome experiments such as the Kaggle "Marketing A/B Testing"
dataset (columns: ``test group`` in {ad, psa}; ``converted`` in {True, False}).

Usage:
    python src/conversion_ab_test.py path/to/marketing_AB.csv [--rollout-users 1000000]
"""
from __future__ import annotations

import argparse
import math
from dataclasses import dataclass

import numpy as np
import pandas as pd
from scipy import stats
from statsmodels.stats.proportion import proportions_ztest


@dataclass
class ConversionResult:
    n_control: int
    n_test: int
    conv_control: int
    conv_test: int
    rate_control: float
    rate_test: float
    abs_lift: float
    abs_lift_ci: tuple
    rel_lift: float
    rel_lift_ci: tuple
    z_stat: float
    z_pvalue: float
    chi2_stat: float
    chi2_pvalue: float
    alpha: float

    @property
    def significant(self) -> bool:
        return self.z_pvalue < self.alpha

    def incremental_conversions(self, population: int) -> dict:
        """Extra conversions if the *test* experience were shown to ``population`` users
        instead of the control experience (point estimate and CI on the absolute lift)."""
        lo, hi = self.abs_lift_ci
        return {
            "point": self.abs_lift * population,
            "ci_low": lo * population,
            "ci_high": hi * population,
        }


def analyze_conversion(conv_control: int, n_control: int, conv_test: int, n_test: int,
                       alpha: float = 0.05) -> ConversionResult:
    """Two-sided two-proportion z-test (pooled SE) + Pearson chi-square (no continuity
    correction, which makes it algebraically equal to z**2). CIs use the unpooled SE."""
    if not (0 <= conv_control <= n_control and 0 <= conv_test <= n_test):
        raise ValueError("conversions must be between 0 and the group size")
    p_c, p_t = conv_control / n_control, conv_test / n_test

    z, p_z = proportions_ztest([conv_test, conv_control], [n_test, n_control])

    table = np.array([[conv_control, n_control - conv_control],
                      [conv_test, n_test - conv_test]])
    chi2, p_chi, _, _ = stats.chi2_contingency(table, correction=False)

    crit = stats.norm.ppf(1 - alpha / 2)
    diff = p_t - p_c
    se = math.sqrt(p_t * (1 - p_t) / n_test + p_c * (1 - p_c) / n_control)
    abs_ci = (diff - crit * se, diff + crit * se)

    # Relative lift (p_t / p_c - 1): Katz log-method CI
    if conv_control > 0 and conv_test > 0:
        rr = p_t / p_c
        se_log = math.sqrt((1 - p_t) / conv_test + (1 - p_c) / conv_control)
        rel_ci = (rr * math.exp(-crit * se_log) - 1, rr * math.exp(crit * se_log) - 1)
        rel = rr - 1
    else:
        rel, rel_ci = float("nan"), (float("nan"), float("nan"))

    return ConversionResult(n_control, n_test, conv_control, conv_test, p_c, p_t,
                            diff, abs_ci, rel, rel_ci, float(z), float(p_z),
                            float(chi2), float(p_chi), alpha)


def analyze_dataframe(df: pd.DataFrame, group_col: str = "test group",
                      outcome_col: str = "converted", control_label: str = "psa",
                      test_label: str = "ad", alpha: float = 0.05) -> ConversionResult:
    c = df.loc[df[group_col] == control_label, outcome_col].astype(int)
    t = df.loc[df[group_col] == test_label, outcome_col].astype(int)
    return analyze_conversion(int(c.sum()), len(c), int(t.sum()), len(t), alpha)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("csv")
    ap.add_argument("--rollout-users", type=int, default=1_000_000)
    ap.add_argument("--alpha", type=float, default=0.05)
    a = ap.parse_args()
    r = analyze_dataframe(pd.read_csv(a.csv), alpha=a.alpha)
    inc = r.incremental_conversions(a.rollout_users)
    print(f"Control rate {r.rate_control:.4%} (n={r.n_control:,}) | Test rate {r.rate_test:.4%} (n={r.n_test:,})")
    print(f"Absolute lift {r.abs_lift:+.4%}  CI {r.abs_lift_ci[0]:+.4%}..{r.abs_lift_ci[1]:+.4%}")
    print(f"Relative lift {r.rel_lift:+.2%}  CI {r.rel_lift_ci[0]:+.2%}..{r.rel_lift_ci[1]:+.2%}")
    print(f"z = {r.z_stat:.3f}, p = {r.z_pvalue:.3g} | chi2 = {r.chi2_stat:.3f}, p = {r.chi2_pvalue:.3g}")
    print("Significant at {:.0%} confidence: {}".format(1 - r.alpha, r.significant))
    print(f"Rollout to {a.rollout_users:,} users: {inc['point']:+,.0f} conversions "
          f"(CI {inc['ci_low']:+,.0f}..{inc['ci_high']:+,.0f})")


if __name__ == "__main__":
    main()
