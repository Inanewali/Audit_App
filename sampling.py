"""Audit sampling engine: population cleaning, sample sizing, selection and evaluation.

Everything here is pure (no Streamlit) so it can be unit tested.

Conventions
-----------
* ``ria`` is the risk of incorrect acceptance as a fraction (0.05 = 5%).
  Confidence = 1 - ria.
* Misstatement amounts are *overstatements*: book value minus audited value.
* Every random draw takes an explicit ``seed`` so a selection can be re-performed
  and documented in the workpaper.

Methods follow the AICPA Audit Guide *Audit Sampling* (Poisson-based monetary
unit sampling, binomial attribute sampling, value-weighted stratification).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# Internal column names added to the population
AMT = "Amt_Clean"
ID_KEY = "_id_key"
STRATUM = "Stratum"
REASON = "Selection_Reason"
METHOD = "Sampling_Method"
HITS = "MUS_Hits"

# AICPA expansion factors for expected misstatement in MUS (keyed by RIA).
_EXPANSION_FACTORS = {0.01: 1.9, 0.05: 1.6, 0.10: 1.5, 0.15: 1.4, 0.20: 1.3,
                      0.25: 1.25, 0.30: 1.2, 0.37: 1.15, 0.50: 1.0}


# --------------------------------------------------------------------------- #
# Probability helpers (no SciPy dependency)
# --------------------------------------------------------------------------- #
def poisson_cdf(k: int, lam: float) -> float:
    """P(X <= k) for X ~ Poisson(lam)."""
    if lam <= 0:
        return 1.0
    term = math.exp(-lam)
    total = term
    for i in range(1, k + 1):
        term *= lam / i
        total += term
    return min(total, 1.0)


def binom_cdf(k: int, n: int, p: float) -> float:
    """P(X <= k) for X ~ Binomial(n, p), computed in log space for stability."""
    if k < 0:
        return 0.0
    if k >= n or p <= 0:
        return 1.0
    if p >= 1:
        return 0.0
    lp, lq = math.log(p), math.log1p(-p)
    total = 0.0
    for i in range(k + 1):
        log_term = (math.lgamma(n + 1) - math.lgamma(i + 1) - math.lgamma(n - i + 1)
                    + i * lp + (n - i) * lq)
        total += math.exp(log_term)
    return min(total, 1.0)


def _bisect(f, lo: float, hi: float, iters: int = 100) -> float:
    """Find x in [lo, hi] where decreasing f crosses zero."""
    for _ in range(iters):
        mid = (lo + hi) / 2
        if f(mid) > 0:
            lo = mid
        else:
            hi = mid
    return hi


def reliability_factor(ria: float, errors: int = 0) -> float:
    """Poisson upper-limit factor: smallest lambda with P(X <= errors) <= ria.

    Matches the AICPA confidence-factor table (e.g. 3.00 for 0 errors at 5%,
    2.31 for 0 errors at 10%, 4.75 for 1 error at 5%).
    """
    if not 0 < ria < 1:
        raise ValueError("RIA must be between 0 and 1")
    if errors == 0:
        return -math.log(ria)
    return _bisect(lambda lam: poisson_cdf(errors, lam) - ria, 0.0, 10.0 * (errors + 5))


def expansion_factor(ria: float) -> float:
    """AICPA expansion factor for expected misstatement (linear interpolation)."""
    keys = sorted(_EXPANSION_FACTORS)
    if ria <= keys[0]:
        return _EXPANSION_FACTORS[keys[0]]
    if ria >= keys[-1]:
        return _EXPANSION_FACTORS[keys[-1]]
    for lo, hi in zip(keys, keys[1:]):
        if lo <= ria <= hi:
            w = (ria - lo) / (hi - lo)
            return _EXPANSION_FACTORS[lo] + w * (_EXPANSION_FACTORS[hi] - _EXPANSION_FACTORS[lo])
    return 1.0  # unreachable


# --------------------------------------------------------------------------- #
# Population cleaning
# --------------------------------------------------------------------------- #
def parse_amounts(series: pd.Series) -> pd.Series:
    """Convert amounts like '$1,234.50', '(500)', '1 000' or '-20' to floats (NaN if unparseable)."""
    s = series.astype(str).str.strip()
    negative = s.str.match(r"^\(.*\)$")
    s = s.str.replace(r"[()]", "", regex=True)
    s = s.str.replace(r"[^\d.\-eE]", "", regex=True)
    out = pd.to_numeric(s, errors="coerce")
    out[negative] = -out[negative].abs()
    return out


def normalize_ids(series: pd.Series) -> pd.Series:
    """Make IDs comparable across sheets: 2100458864, 2100458864.0 and ' 2100458864 ' all match."""
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0+$", "", regex=True)
    return s.where(~series.isna(), "")


@dataclass
class CleanResult:
    population: pd.DataFrame
    excluded: pd.DataFrame  # original rows plus an Exclusion_Reason column


def clean_population(df: pd.DataFrame, id_col: str, amt_col: str,
                     name_col: str | None = None, *,
                     exclude_zero: bool = False) -> CleanResult:
    """Split raw rows into a testable population and a documented list of exclusions.

    Excluded (never silently dropped):
      * total / subtotal rows (blank ID with an amount, or a name/ID containing 'total')
      * amounts that can't be parsed as numbers
      * negative (credit) balances — these need a separate test, not MUS/stratified selection
      * zero balances, if ``exclude_zero``
    """
    work = df.copy()
    work[AMT] = parse_amounts(work[amt_col])
    work[ID_KEY] = normalize_ids(work[id_col])

    reason = pd.Series("", index=work.index, dtype=object)
    text = work[ID_KEY].str.lower()
    if name_col and name_col in work.columns:
        text = text + " " + work[name_col].astype(str).str.lower()
    is_total = text.str.contains(r"\b(?:sub)?total\b", regex=True) | (
        (work[ID_KEY].isin(["", "nan", "none"])) & work[AMT].notna())
    reason[is_total] = "Total / subtotal row"
    reason[(reason == "") & work[AMT].isna()] = "Amount not numeric"
    reason[(reason == "") & (work[AMT] < 0)] = "Negative (credit) balance"
    if exclude_zero:
        reason[(reason == "") & (work[AMT] == 0)] = "Zero balance"

    excluded = work[reason != ""].copy()
    excluded["Exclusion_Reason"] = reason[reason != ""]
    population = work[reason == ""].copy()
    return CleanResult(population=population, excluded=excluded)


# --------------------------------------------------------------------------- #
# Sample sizes
# --------------------------------------------------------------------------- #
def mus_sample_size(book_value: float, tolerable: float, ria: float,
                    expected: float = 0.0) -> int:
    """AICPA MUS sample size: n = BV x RF / (TM - EM x expansion factor)."""
    if tolerable <= 0:
        raise ValueError("Tolerable misstatement must be greater than zero")
    if book_value <= 0:
        return 0
    denom = tolerable - expected * expansion_factor(ria)
    if denom <= 0:
        raise ValueError("Expected misstatement is too close to tolerable misstatement — "
                         "sampling cannot give the required assurance. Lower expected "
                         "misstatement or test 100%.")
    return int(math.ceil(book_value * reliability_factor(ria) / denom))


def attribute_sample_size(tolerable_rate: float, expected_rate: float, ria: float,
                          population_size: int | None = None, max_n: int = 5000) -> int:
    """Smallest n such that, if the expected deviations are found, the upper deviation
    limit at 1 - RIA does not exceed the tolerable rate (binomial; AICPA table basis)."""
    if not 0 < tolerable_rate < 1:
        raise ValueError("Tolerable deviation rate must be between 0 and 1")
    if expected_rate >= tolerable_rate:
        raise ValueError("Expected deviation rate must be below the tolerable rate")
    for n in range(1, max_n + 1):
        k = math.ceil(n * expected_rate - 1e-9) if expected_rate > 0 else 0
        if binom_cdf(k, n, tolerable_rate) <= ria:
            break
    else:
        n = max_n
    if population_size:
        n = min(n, population_size)
    return n


# --------------------------------------------------------------------------- #
# Selection
# --------------------------------------------------------------------------- #
def _tag(frame: pd.DataFrame, reason: str, method: str) -> pd.DataFrame:
    frame = frame.copy()
    frame[REASON] = reason
    frame[METHOD] = method
    return frame


def mus_select(pop: pd.DataFrame, n: int, seed: int) -> tuple[pd.DataFrame, float]:
    """Systematic PPS selection with a random start. Returns (sample, interval).

    Every item >= the sampling interval is certain to be hit and is flagged as a
    key item. Items are hit at most once in the returned sample; ``MUS_Hits``
    records how many selection points fell in each item.
    """
    positive = pop[pop[AMT] > 0]
    total = positive[AMT].sum()
    if n <= 0 or total <= 0:
        return _tag(pop.iloc[0:0], "", "MUS").assign(**{HITS: []}), 0.0
    interval = total / n
    rng = np.random.default_rng(seed)
    start = rng.uniform(0, interval)
    points = start + interval * np.arange(n)
    cum = positive[AMT].cumsum().to_numpy()
    idx = np.searchsorted(cum, points, side="right")
    idx = idx[idx < len(positive)]
    hits = pd.Series(idx).value_counts()
    sample = positive.iloc[hits.index.sort_values()].copy()
    sample[HITS] = hits.sort_index().to_numpy()
    key = sample[AMT] >= interval
    sample[REASON] = np.where(key, "Key item (>= sampling interval)", "MUS selection")
    sample[METHOD] = "MUS"
    return sample, interval


def value_strata(amounts: pd.Series, n_strata: int) -> pd.Series:
    """Assign strata of roughly equal total value (1 = lowest values)."""
    if len(amounts) == 0:
        return pd.Series(dtype=int)
    n_strata = max(1, min(n_strata, len(amounts)))
    order = amounts.sort_values(kind="mergesort")
    share = order.cumsum() / order.sum() if order.sum() > 0 else pd.Series(
        np.linspace(0, 1, len(order), endpoint=False) + 1 / len(order), index=order.index)
    strata = np.minimum((share * n_strata - 1e-9).apply(math.floor) + 1, n_strata)
    return strata.clip(lower=1).reindex(amounts.index).astype(int)


def allocate(sizes: pd.Series, weights: pd.Series, n: int) -> pd.Series:
    """Allocate n across strata proportionally to weight, at least 1 each, capped at stratum size."""
    alloc = pd.Series(0, index=sizes.index)
    n = int(min(n, sizes.sum()))
    if n <= 0:
        return alloc
    w = weights.clip(lower=0)
    w = w / w.sum() if w.sum() > 0 else sizes / sizes.sum()
    raw = w * n
    alloc = np.floor(raw).astype(int).clip(lower=1).combine(sizes, min)
    # hand out any remainder by largest fractional part, respecting caps
    while alloc.sum() < n:
        room = sizes - alloc
        cand = (raw - alloc).where(room > 0)
        if cand.isna().all():
            break
        alloc[cand.idxmax()] += 1
    while alloc.sum() > n:
        alloc[(alloc - raw).idxmax()] -= 1
    return alloc


def assign_strata(pop: pd.DataFrame, key_threshold: float, n_strata: int) -> pd.DataFrame:
    """Label every population item: 'Key' if >= threshold, else value stratum '1'..'n'."""
    out = pop.copy()
    is_key = out[AMT] >= key_threshold
    out[STRATUM] = "Key"
    if (~is_key).any():
        out.loc[~is_key, STRATUM] = value_strata(out.loc[~is_key, AMT], n_strata).astype(str)
    return out


def stratified_select(pop: pd.DataFrame, n_random: int, key_threshold: float,
                      n_strata: int, seed: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Key items (>= threshold) tested 100%; remainder split into value strata and
    sampled at random, allocated by stratum value. Returns (sample, strata_summary).
    ``pop`` should already carry a Stratum column from :func:`assign_strata`."""
    if STRATUM not in pop.columns:
        pop = assign_strata(pop, key_threshold, n_strata)
    key = pop[pop[STRATUM] == "Key"]
    rest = pop[pop[STRATUM] != "Key"]

    summary_rows = []
    picks = [_tag(key, "Key item (>= threshold)", "Stratified").assign(**{STRATUM: "Key"})]
    if len(key):
        summary_rows.append({"Stratum": "Key", "Items": len(key), "Value": key[AMT].sum(),
                             "Lower": key[AMT].min(), "Upper": key[AMT].max(),
                             "Sample": len(key)})
    if len(rest):
        grp = rest.groupby(STRATUM)
        alloc = allocate(grp.size(), grp[AMT].sum(), n_random)
        for i, (stratum, frame) in enumerate(grp):
            k = int(alloc[stratum])
            chosen = frame.sample(n=k, random_state=seed + i) if k else frame.iloc[0:0]
            picks.append(_tag(chosen, f"Random — stratum {stratum}", "Stratified"))
            summary_rows.append({"Stratum": stratum, "Items": len(frame),
                                 "Value": frame[AMT].sum(), "Lower": frame[AMT].min(),
                                 "Upper": frame[AMT].max(), "Sample": k})
    sample = pd.concat(picks)
    return sample, pd.DataFrame(summary_rows)


def systematic_select(pop: pd.DataFrame, n: int, seed: int) -> tuple[pd.DataFrame, float]:
    """Every k-th item (fractional interval) from a random start. Returns (sample, k)."""
    size = len(pop)
    n = min(n, size)
    if n <= 0:
        return _tag(pop.iloc[0:0], "", "Systematic"), 0.0
    k = size / n
    start = np.random.default_rng(seed).uniform(0, k)
    positions = np.floor(start + k * np.arange(n)).astype(int)
    sample = pop.iloc[np.unique(positions)]
    return _tag(sample, f"Systematic (interval {k:.2f})", "Systematic"), k


def random_select(pop: pd.DataFrame, n: int, seed: int, method: str) -> pd.DataFrame:
    n = min(n, len(pop))
    return _tag(pop.sample(n=n, random_state=seed), "Simple random selection", method)


def insider_items(pop: pd.DataFrame, staff_ids: pd.Series) -> pd.DataFrame:
    keys = set(normalize_ids(staff_ids)) - {""}
    return _tag(pop[pop[ID_KEY].isin(keys)], "Staff / insider (targeted)", "Targeted")


def combine(*frames: pd.DataFrame) -> pd.DataFrame:
    """Concatenate selections; an item picked for several reasons keeps all of them."""
    frames = [f for f in frames if len(f)]
    if not frames:
        return pd.DataFrame()
    allrows = pd.concat(frames)
    reasons = allrows.groupby(level=0)[REASON].agg(lambda r: "; ".join(dict.fromkeys(r)))
    methods = allrows.groupby(level=0)[METHOD].agg(lambda r: "; ".join(dict.fromkeys(r)))
    out = allrows[~allrows.index.duplicated(keep="first")].copy()
    out[REASON] = reasons.reindex(out.index)
    out[METHOD] = methods.reindex(out.index)
    return out


# --------------------------------------------------------------------------- #
# Evaluation
# --------------------------------------------------------------------------- #
@dataclass
class Evaluation:
    method: str
    conclusion: str
    supports: bool
    figures: dict = field(default_factory=dict)
    detail: pd.DataFrame | None = None


def evaluate_mus(sample: pd.DataFrame, error_col: str, interval: float, ria: float,
                 tolerable: float) -> Evaluation:
    """Upper misstatement limit = basic precision + projected misstatement + incremental allowance."""
    s = sample.copy()
    s["_err"] = pd.to_numeric(s[error_col], errors="coerce").fillna(0.0)
    key = s[AMT] >= interval
    known = s.loc[key, "_err"].sum()

    nk = s[~key & (s["_err"] != 0) & (s[AMT] > 0)].copy()
    nk["Tainting"] = (nk["_err"] / nk[AMT]).clip(upper=1.0)
    over = nk[nk["Tainting"] > 0].sort_values("Tainting", ascending=False)
    under = nk[nk["Tainting"] < 0]

    basic = reliability_factor(ria, 0) * interval
    rows, projected, allowance = [], 0.0, 0.0
    for rank, (_, r) in enumerate(over.iterrows(), start=1):
        proj = r["Tainting"] * interval
        increment = reliability_factor(ria, rank) - reliability_factor(ria, rank - 1)
        projected += proj
        allowance += (increment - 1) * proj
        rows.append({"Rank": rank, "Book value": r[AMT], "Misstatement": r["_err"],
                     "Tainting": r["Tainting"], "Projected": proj,
                     "Factor increment": increment})
    under_projected = (under["Tainting"] * interval).sum()

    uml = known + basic + projected + allowance
    most_likely = known + projected + under_projected
    supports = uml <= tolerable
    figures = {"Sampling interval": interval, "Known misstatement (key items)": known,
               "Basic precision": basic, "Projected misstatement": projected,
               "Incremental allowance": allowance,
               "Projected understatements (net, not in UML)": under_projected,
               "Most likely misstatement": most_likely,
               "Upper misstatement limit": uml, "Tolerable misstatement": tolerable}
    conclusion = (f"Upper misstatement limit {uml:,.2f} is "
                  f"{'within' if supports else 'ABOVE'} tolerable misstatement {tolerable:,.2f} "
                  f"at {100 * (1 - ria):.0f}% confidence.")
    return Evaluation("MUS", conclusion, supports, figures, pd.DataFrame(rows))


def upper_deviation_rate(n: int, deviations: int, ria: float) -> float:
    """Exact binomial upper limit on the population deviation rate."""
    if n <= 0:
        return 1.0
    if deviations >= n:
        return 1.0
    return _bisect(lambda p: binom_cdf(deviations, n, p) - ria, 0.0, 1.0)


def evaluate_attribute(n: int, deviations: int, ria: float, tolerable_rate: float) -> Evaluation:
    udl = upper_deviation_rate(n, deviations, ria)
    supports = udl <= tolerable_rate
    figures = {"Sample size": n, "Deviations found": deviations,
               "Sample deviation rate": deviations / n if n else 0.0,
               "Upper deviation limit": udl, "Tolerable deviation rate": tolerable_rate}
    conclusion = (f"Upper deviation limit {udl:.2%} is {'within' if supports else 'ABOVE'} "
                  f"the tolerable rate {tolerable_rate:.2%} at {100 * (1 - ria):.0f}% confidence.")
    return Evaluation("Attribute", conclusion, supports, figures)


def evaluate_ratio(sample: pd.DataFrame, population: pd.DataFrame, error_col: str,
                   tolerable: float, group_col: str | None = None) -> Evaluation:
    """Ratio projection (misstatement / book in sample x population book), per stratum.

    Key items are projected at their actual misstatement. This gives the projected
    (most likely) misstatement; the allowance for sampling risk is left to judgment.
    """
    s = sample.copy()
    s["_err"] = pd.to_numeric(s[error_col], errors="coerce").fillna(0.0)
    groups = [(None, s, population)] if group_col is None else [
        (g, s[s[group_col] == g], population[population[group_col] == g])
        for g in population[group_col].dropna().unique()]
    rows, total = [], 0.0
    for g, ss, pp in groups:
        book_s, err_s, book_p = ss[AMT].sum(), ss["_err"].sum(), pp[AMT].sum()
        if g == "Key":
            proj = err_s
        else:
            proj = (err_s / book_s * book_p) if book_s > 0 else 0.0
        total += proj
        rows.append({"Stratum": g if g is not None else "All", "Population value": book_p,
                     "Sample value": book_s, "Misstatement found": err_s,
                     "Projected misstatement": proj})
    supports = total < tolerable
    figures = {"Projected misstatement": total, "Tolerable misstatement": tolerable,
               "Headroom for sampling risk": tolerable - total}
    conclusion = (f"Projected misstatement {total:,.2f} is "
                  f"{'below' if supports else 'AT OR ABOVE'} tolerable misstatement "
                  f"{tolerable:,.2f}. Judge whether the headroom of {tolerable - total:,.2f} "
                  "is enough allowance for sampling risk.")
    return Evaluation("Ratio projection", conclusion, supports, figures,
                      pd.DataFrame(rows).sort_values("Stratum", key=lambda c: c.astype(str)))
