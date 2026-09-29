import math

import numpy as np
import pandas as pd
import pytest

import sampling as s


# ---- factors match the AICPA Audit Sampling guide tables ------------------------
@pytest.mark.parametrize("ria,errors,expected", [
    (0.05, 0, 3.00), (0.10, 0, 2.31), (0.05, 1, 4.75), (0.10, 1, 3.89),
    (0.05, 2, 6.30), (0.20, 0, 1.61),
])
def test_reliability_factors(ria, errors, expected):
    assert s.reliability_factor(ria, errors) == pytest.approx(expected, abs=0.01)


@pytest.mark.parametrize("tdr,exp,ria,n", [
    (0.05, 0.0, 0.05, 59), (0.05, 0.0, 0.10, 45), (0.10, 0.0, 0.05, 29),
    (0.05, 0.01, 0.05, 93), (0.06, 0.01, 0.10, 64),
])
def test_attribute_sample_size_matches_aicpa_table(tdr, exp, ria, n):
    assert s.attribute_sample_size(tdr, exp, ria) == n


def test_attribute_rejects_expected_above_tolerable():
    with pytest.raises(ValueError):
        s.attribute_sample_size(0.02, 0.03, 0.05)


def test_mus_sample_size():
    # BV 5,000,000, TM 150,000, EM 0, RIA 5%: 5e6 * 3.0 / 150,000 = 100
    assert s.mus_sample_size(5_000_000, 150_000, 0.05) == 100
    # with expected misstatement 20,000 -> denominator 150,000 - 20,000*1.6 = 118,000
    assert s.mus_sample_size(5_000_000, 150_000, 0.05, 20_000) == math.ceil(5e6 * 2.9957 / 118_000)
    with pytest.raises(ValueError):
        s.mus_sample_size(5_000_000, 10_000, 0.05, 10_000)


# ---- cleaning -----------------------------------------------------------------
def raw_frame():
    return pd.DataFrame({
        "ACCOUNT ID": [101, 102, 103, None, 104, "105", 106],
        "NAME": ["A", "B", "C", "", "D", "Grand Total", "E"],
        "BAL": ["1,000", "$2,500.50", "(300)", "9999", "abc", "5000", 0],
    })


def test_clean_population_documents_every_exclusion():
    res = s.clean_population(raw_frame(), "ACCOUNT ID", "BAL", "NAME")
    assert list(res.population["NAME"]) == ["A", "B", "E"]
    reasons = dict(zip(res.excluded["NAME"], res.excluded["Exclusion_Reason"]))
    assert reasons == {"C": "Negative (credit) balance", "": "Total / subtotal row",
                       "D": "Amount not numeric", "Grand Total": "Total / subtotal row"}
    assert len(res.population) + len(res.excluded) == len(raw_frame())
    assert res.population[s.AMT].tolist() == [1000.0, 2500.5, 0.0]


def test_normalize_ids_matches_numbers_and_text():
    a = s.normalize_ids(pd.Series([2100458864, 2100458864.0, " 2100458864 "]))
    assert a.nunique() == 1


# ---- selection ----------------------------------------------------------------
def population(n=500, seed=0):
    rng = np.random.default_rng(seed)
    amts = np.round(rng.lognormal(10, 1.5, n), 2)
    df = pd.DataFrame({"ID": range(1, n + 1), "Name": [f"C{i}" for i in range(n)], "Amt": amts})
    return s.clean_population(df, "ID", "Amt", "Name").population


def test_mus_selects_every_key_item_and_is_reproducible():
    pop = population()
    sample, interval = s.mus_select(pop, 60, seed=7)
    big = pop[pop[s.AMT] >= interval]
    assert set(big.index) <= set(sample.index)
    assert (sample.loc[big.index, s.REASON] == "Key item (>= sampling interval)").all()
    assert sample[s.HITS].sum() == 60
    again, _ = s.mus_select(pop, 60, seed=7)
    assert again.index.equals(sample.index)
    other, _ = s.mus_select(pop, 60, seed=8)
    assert not other.index.equals(sample.index)


def test_stratified_sample_size_is_respected():
    pop = population()
    threshold = pop[s.AMT].quantile(0.98)
    pop = s.assign_strata(pop, threshold, 3)
    sample, summary = s.stratified_select(pop, 40, threshold, 3, seed=1)
    n_key = (pop[s.AMT] >= threshold).sum()
    assert len(sample) == n_key + 40
    assert summary.loc[summary["Stratum"] != "Key", "Sample"].sum() == 40
    # strata carry roughly equal value
    vals = summary.loc[summary["Stratum"] != "Key", "Value"]
    assert vals.max() / vals.min() < 2


def test_allocation_caps_at_stratum_size():
    alloc = s.allocate(pd.Series({"1": 2, "2": 50}), pd.Series({"1": 1e6, "2": 1.0}), 10)
    assert alloc["1"] == 2 and alloc.sum() == 10


def test_systematic_is_seeded_and_sized():
    pop = population(103)
    a, k = s.systematic_select(pop, 10, seed=3)
    b, _ = s.systematic_select(pop, 10, seed=3)
    assert len(a) == 10 and a.index.equals(b.index)
    assert k == pytest.approx(10.3)


def test_insider_matching_across_dtypes():
    pop = population(20)
    staff = pd.Series(["3", "7.0", 12, "999"])
    hits = s.insider_items(pop, staff)
    assert sorted(hits["ID"].tolist()) == [3, 7, 12]


def test_combine_keeps_all_reasons():
    pop = population(20)
    a = s.random_select(pop, 5, 1, "Stratified")
    b = s.insider_items(pop, pd.Series(a["ID"].iloc[:1].astype(str)))
    out = s.combine(a, b)
    assert len(out) == 5
    assert out[s.REASON].str.contains("; ").sum() == 1


# ---- evaluation ---------------------------------------------------------------
def test_upper_deviation_rate():
    # 0 deviations in 59 at 95% -> ~4.95%
    assert s.upper_deviation_rate(59, 0, 0.05) == pytest.approx(1 - 0.05 ** (1 / 59), abs=1e-6)
    ev = s.evaluate_attribute(93, 1, 0.05, 0.05)
    assert ev.supports and ev.figures["Upper deviation limit"] < 0.05
    assert not s.evaluate_attribute(59, 2, 0.05, 0.05).supports


def test_mus_evaluation_no_errors_equals_basic_precision():
    sample = pd.DataFrame({s.AMT: [100.0, 200.0, 5000.0], "Err": [0, 0, 0]})
    ev = s.evaluate_mus(sample, "Err", interval=1000, ria=0.05, tolerable=5000)
    assert ev.figures["Upper misstatement limit"] == pytest.approx(2995.7, abs=0.1)
    assert ev.supports


def test_mus_evaluation_with_errors():
    # textbook-style: interval 1,000, RIA 5%
    # key item error 400 (actual); taintings 0.5 and 0.1
    sample = pd.DataFrame({s.AMT: [2000.0, 100.0, 500.0, 300.0],
                           "Err": [400.0, 50.0, 50.0, 0.0]})
    ev = s.evaluate_mus(sample, "Err", interval=1000, ria=0.05, tolerable=10_000)
    f = ev.figures
    assert f["Known misstatement (key items)"] == 400
    assert f["Projected misstatement"] == pytest.approx(600)
    # allowance: (4.744-3.0-1)*500 + (6.296-4.744-1)*100
    expected_allow = (s.reliability_factor(.05, 1) - s.reliability_factor(.05, 0) - 1) * 500 \
        + (s.reliability_factor(.05, 2) - s.reliability_factor(.05, 1) - 1) * 100
    assert f["Incremental allowance"] == pytest.approx(expected_allow)
    assert f["Upper misstatement limit"] == pytest.approx(400 + 2995.73 + 600 + expected_allow, abs=0.1)


def test_ratio_projection_by_stratum():
    pop = pd.DataFrame({s.AMT: [10_000.0, 100.0, 100.0, 100.0, 100.0],
                        s.STRATUM: ["Key", "1", "1", "1", "1"]})
    sample = pop.iloc[[0, 1, 2]].assign(Err=[500.0, 10.0, 0.0])
    ev = s.evaluate_ratio(sample, pop, "Err", tolerable=1000, group_col=s.STRATUM)
    # key: 500 actual; stratum 1: 10/200 * 400 = 20
    assert ev.figures["Projected misstatement"] == pytest.approx(520)
    assert ev.supports
