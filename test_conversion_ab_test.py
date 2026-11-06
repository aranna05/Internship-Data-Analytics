import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from conversion_ab_test import analyze_conversion, analyze_dataframe  # noqa: E402


def test_z_squared_equals_chi2():
    r = analyze_conversion(200, 10_000, 260, 10_000)
    assert r.z_stat ** 2 == pytest.approx(r.chi2_stat, rel=1e-6)
    assert r.z_pvalue == pytest.approx(r.chi2_pvalue, rel=1e-6)


def test_detects_real_effect_and_ci_excludes_zero():
    r = analyze_conversion(200, 10_000, 300, 10_000)
    assert r.significant and r.abs_lift_ci[0] > 0 and r.rel_lift_ci[0] > 0


def test_no_effect_not_significant():
    r = analyze_conversion(200, 10_000, 202, 10_000)
    assert not r.significant and r.abs_lift_ci[0] < 0 < r.abs_lift_ci[1]


def test_known_values():
    r = analyze_conversion(100, 1000, 150, 1000)
    assert r.abs_lift == pytest.approx(0.05)
    assert r.rel_lift == pytest.approx(0.5)
    assert r.incremental_conversions(1_000_000)["point"] == pytest.approx(50_000)


def test_dataframe_schema_of_kaggle_dataset():
    rng = np.random.default_rng(0)
    df = pd.DataFrame({"test group": ["ad"] * 5000 + ["psa"] * 5000,
                       "converted": np.r_[rng.random(5000) < 0.03, rng.random(5000) < 0.02]})
    r = analyze_dataframe(df)
    assert r.rate_test > r.rate_control and r.significant


def test_invalid_input():
    with pytest.raises(ValueError):
        analyze_conversion(11, 10, 1, 10)
