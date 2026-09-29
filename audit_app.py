"""Advanced Audit Sampling & Selection Tool (AAST) — Streamlit UI.

All statistics live in sampling.py (unit tested). This file only handles input,
display and export.
"""

import hashlib
import io
import json
from datetime import datetime

import pandas as pd
import streamlit as st

import sampling as s

st.set_page_config(page_title="Audit Sampling Tool", page_icon="🎯", layout="wide")

METHODS = {
    "Stratified (key items + value strata)": "stratified",
    "Monetary Unit Sampling (MUS)": "mus",
    "Attribute sampling (controls)": "attribute",
    "Systematic": "systematic",
}
RIA_OPTIONS = [0.05, 0.10, 0.15, 0.20, 0.25]
STATUS_OPTIONS = ["Pending", "Tested – no exception", "Tested – exception", "Not testable"]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def read_book(data: bytes, name: str) -> dict[str, pd.DataFrame]:
    if name.lower().endswith(".csv"):
        return {"CSV": pd.read_csv(io.BytesIO(data))}
    return pd.read_excel(io.BytesIO(data), sheet_name=None)


def guess(columns, keywords, fallback=0):
    for kw in keywords:
        for i, c in enumerate(columns):
            if kw in str(c).lower():
                return i
    return min(fallback, len(columns) - 1)


def pct(x):
    return f"{x:.2%}"


# --------------------------------------------------------------------------- #
# Sidebar — parameters
# --------------------------------------------------------------------------- #
sb = st.sidebar
sb.header("⚙️ Sampling parameters")
method_label = sb.radio("Method", list(METHODS), help=(
    "**Stratified** – substantive tests of balances; key items tested 100%, remainder "
    "split into strata of equal value.\n\n**MUS** – substantive tests where overstatement "
    "is the main risk; selection probability proportional to value.\n\n**Attribute** – "
    "tests of controls; items picked at random regardless of value.\n\n**Systematic** – "
    "every k-th item from a random start."))
method = METHODS[method_label]

ria = sb.select_slider("Confidence", RIA_OPTIONS, value=0.05,
                       format_func=lambda r: f"{100 * (1 - r):.0f}% (RIA {r:.0%})",
                       help="Risk of incorrect acceptance (RIA) = 1 − confidence.")
cur = sb.text_input("Currency symbol", value="$", max_chars=4)


def money(x):
    return f"{cur}{x:,.2f}"


monetary = method != "attribute"
if monetary:
    sb.markdown("### Materiality")
    tm_basis = sb.radio("Tolerable misstatement basis", ["% of population value", "Amount"],
                        horizontal=True)
    if tm_basis == "Amount":
        tm_amount = sb.number_input("Tolerable misstatement", min_value=0.0, value=10_000.0,
                                    step=1_000.0, format="%.2f")
    else:
        tm_pct = sb.number_input("Tolerable misstatement (%)", 0.1, 20.0, 2.0, 0.1) / 100
    em_share = sb.slider("Expected misstatement (% of tolerable)", 0, 60, 0, 5,
                         help="Misstatement you expect to find. Larger values increase the "
                              "sample size.") / 100
else:
    sb.markdown("### Deviation rates")
    tdr = sb.number_input("Tolerable deviation rate (%)", 1.0, 20.0, 5.0, 0.5) / 100
    edr = sb.number_input("Expected deviation rate (%)", 0.0, 10.0, 0.0, 0.25) / 100

if method == "stratified":
    sb.markdown("### Strata")
    key_basis = sb.radio("Key item threshold", ["= tolerable misstatement", "Fixed amount",
                                               "Top % of items"])
    if key_basis == "Fixed amount":
        key_amount = sb.number_input("Key item threshold amount", min_value=0.0,
                                     value=100_000.0, step=10_000.0, format="%.2f")
    elif key_basis == "Top % of items":
        key_top = sb.slider("Top % of items by value", 1, 25, 5) / 100
    n_strata = sb.slider("Strata for the remainder", 1, 6, 3)

sb.markdown("### Selection")
seed = int(sb.number_input("Random seed", min_value=0, value=20260929, step=1, help=(
    "Recorded in the workpaper so the selection can be re-performed exactly. "
    "Change it only to draw a different sample.")))
override = sb.checkbox("Override calculated sample size")
if override:
    override_n = int(sb.number_input("Sample size", min_value=1, value=60, step=1))

# --------------------------------------------------------------------------- #
# Data input
# --------------------------------------------------------------------------- #
st.title("🎯 Audit Sampling & Selection Tool")
st.caption("Statistical sample sizes, reproducible selection, an editable workpaper and "
           "evaluation of results — AICPA *Audit Sampling* methodology.")

upload = st.file_uploader("Upload population (Excel or CSV)", type=["xlsx", "csv"])
if not upload:
    st.info("⬆️ Upload a population listing to begin. Each row should be one item with an ID "
            "and an amount. Try `sample_data/loan_population.xlsx` from the repo.")
    st.stop()

file_bytes = upload.getvalue()
try:
    book = read_book(file_bytes, upload.name)
except Exception as e:  # noqa: BLE001 — show any read error to the user
    st.error(f"Couldn't read the file: {e}")
    st.stop()

sheets = list(book)
c1, c2 = st.columns(2)
pop_sheet = c1.selectbox("Population sheet", sheets,
                         index=sheets.index("Loans") if "Loans" in sheets else 0)
staff_candidates = [x for x in sheets if x != pop_sheet]
use_staff = c2.checkbox("Also select staff / insider items from another sheet",
                        value="Staff" in staff_candidates, disabled=not staff_candidates)

raw = book[pop_sheet].copy()
raw.columns = [str(c).strip() for c in raw.columns]
cols = list(raw.columns)

st.markdown("#### Column mapping")
m1, m2, m3 = st.columns(3)
id_col = m1.selectbox("ID column", cols, index=guess(cols, ["account", "id", "no", "number"]))
amt_col = m2.selectbox("Amount column", cols,
                       index=guess(cols, ["outstanding", "balance", "amount", "amt", "value"], 1))
name_col = m3.selectbox("Name / description column", cols,
                        index=guess(cols, ["name", "customer", "description", "desc"], 2))

staff_ids = None
if use_staff and staff_candidates:
    s1, s2 = st.columns(2)
    staff_sheet = s1.selectbox("Staff sheet", staff_candidates,
                               index=staff_candidates.index("Staff")
                               if "Staff" in staff_candidates else 0)
    staff_df = book[staff_sheet]
    staff_df.columns = [str(c).strip() for c in staff_df.columns]
    staff_cols = list(staff_df.columns)
    staff_id_col = s2.selectbox("Staff sheet ID column", staff_cols,
                                index=guess(staff_cols, [id_col.lower(), "account", "id"]))
    staff_ids = staff_df[staff_id_col]

# --------------------------------------------------------------------------- #
# Clean + size + select
# --------------------------------------------------------------------------- #
cleaned = s.clean_population(raw, id_col, amt_col, name_col)
pop, excluded = cleaned.population, cleaned.excluded
if pop.empty:
    st.error("No usable rows — check the ID and amount column mapping.")
    st.stop()

N = len(pop)
BV = pop[s.AMT].sum()
dupes = pop[s.ID_KEY].duplicated(keep=False)

if monetary:
    tolerable = tm_amount if tm_basis == "Amount" else tm_pct * BV
    expected = em_share * tolerable
    if tolerable <= 0:
        st.error("Set a tolerable misstatement greater than zero.")
        st.stop()

strata_summary = None
interval = None
plan_notes = []
try:
    if method == "mus":
        n_calc = s.mus_sample_size(BV, tolerable, ria, expected)
        n = override_n if override else n_calc
        sample, interval = s.mus_select(pop, n, seed)
        eval_pop = pop

    elif method == "stratified":
        if key_basis == "= tolerable misstatement":
            key_threshold = tolerable
        elif key_basis == "Fixed amount":
            key_threshold = key_amount
        else:
            key_threshold = pop[s.AMT].quantile(1 - key_top)
        pop = s.assign_strata(pop, key_threshold, n_strata)
        rest_value = pop.loc[pop[s.STRATUM] != "Key", s.AMT].sum()
        n_calc = s.mus_sample_size(rest_value, tolerable, ria, expected) if rest_value > 0 else 0
        n = override_n if override else n_calc
        sample, strata_summary = s.stratified_select(pop, n, key_threshold, n_strata, seed)
        eval_pop = pop
        plan_notes.append(f"Key item threshold {money(key_threshold)}; random sample of "
                          f"{n} sized on remaining value {money(rest_value)}.")

    elif method == "systematic":
        n_calc = min(s.mus_sample_size(BV, tolerable, ria, expected), N)
        n = override_n if override else n_calc
        sample, k = s.systematic_select(pop, n, seed)
        eval_pop = pop
        plan_notes.append(f"Interval {k:.2f} items from a random start (seed {seed}). "
                          "Population order matters — sort deliberately before uploading.")

    else:  # attribute
        n_calc = s.attribute_sample_size(tdr, edr, ria, N)
        n = override_n if override else n_calc
        sample = s.random_select(pop, n, seed, "Attribute")
        eval_pop = pop
except ValueError as e:
    st.error(str(e))
    st.stop()

if staff_ids is not None:
    sample = s.combine(sample, s.insider_items(pop, staff_ids))

if override:
    plan_notes.append(f"Sample size overridden: calculated {n_calc}, used {n}.")
if n > N and method != "stratified":
    plan_notes.append("Calculated sample exceeds the population — consider testing 100%.")

# --------------------------------------------------------------------------- #
# Population summary strip
# --------------------------------------------------------------------------- #
a, b, c, d = st.columns(4)
a.metric("Population items", f"{N:,}")
b.metric("Population value", money(BV))
c.metric("Sample items", f"{len(sample):,}")
d.metric("Value covered", pct(sample[s.AMT].sum() / BV) if BV else "–")

if len(excluded):
    ex_val = excluded[s.AMT].fillna(0).abs().sum()
    st.warning(f"{len(excluded)} row(s) excluded from the population "
               f"({money(ex_val)} absolute value) — see the Population tab. "
               "Credit balances and unparseable amounts need separate procedures.")
if dupes.any():
    st.warning(f"{dupes.sum()} rows share an ID with another row. They are treated as "
               "separate items; check the ID column mapping if that's unexpected.")
for note in plan_notes:
    st.info(note)

# --------------------------------------------------------------------------- #
# Plan record (used on screen and in the export)
# --------------------------------------------------------------------------- #
plan = {
    "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "Source file": upload.name,
    "Population sheet": pop_sheet,
    "ID / amount / name columns": f"{id_col} / {amt_col} / {name_col}",
    "Method": method_label,
    "Confidence (1 − RIA)": pct(1 - ria),
    "Population items": N,
    "Population value": money(BV),
    "Rows excluded": len(excluded),
    "Calculated sample size": n_calc,
    "Sample size used": n,
    "Items selected (incl. targeted)": len(sample),
    "Random seed": seed,
}
if monetary:
    plan["Tolerable misstatement"] = money(tolerable)
    plan["Expected misstatement"] = money(expected)
    plan["Reliability factor"] = f"{s.reliability_factor(ria):.2f}"
if interval:
    plan["Sampling interval"] = money(interval)
if method == "attribute":
    plan["Tolerable deviation rate"] = pct(tdr)
    plan["Expected deviation rate"] = pct(edr)
if staff_ids is not None:
    plan["Staff sheet"] = staff_sheet
for i, note in enumerate(plan_notes, 1):
    plan[f"Note {i}"] = note
plan_df = pd.DataFrame({"Parameter": list(plan), "Value": [str(v) for v in plan.values()]})

# --------------------------------------------------------------------------- #
# Workpaper state — keyed on everything that determines the selection
# --------------------------------------------------------------------------- #
sig = hashlib.sha1(json.dumps([hashlib.sha1(file_bytes).hexdigest(), pop_sheet, id_col, amt_col,
                               name_col, method, ria, n, seed,
                               sorted(sample.index.tolist())],
                              default=str).encode()).hexdigest()[:12]

base = pd.DataFrame({
    "ID": sample[id_col].values,
    "Name": sample[name_col].values,
    "Book_Value": sample[s.AMT].values,
    "Selection_Reason": sample[s.REASON].values,
    "Stratum": sample[s.STRATUM].values if s.STRATUM in sample else "",
    "Audit_Status": "Pending",
    "Audited_Value": pd.Series([None] * len(sample), dtype="float64").values,
    "Deviation": False,
    "Tested_By": "",
    "Test_Date": pd.Series([pd.NaT] * len(sample), dtype="datetime64[ns]").values,
    "Audit_Notes": "",
}, index=sample.index)
if method == "mus":
    base.insert(4, "MUS_Hits", sample[s.HITS].values)
if method != "stratified":
    base = base.drop(columns="Stratum")
if monetary:
    base = base.drop(columns="Deviation")
else:
    base = base.drop(columns="Audited_Value")

resume_key = f"resume_{sig}"
if resume_key in st.session_state:
    prior = st.session_state[resume_key]
    by_id = prior.set_index(s.normalize_ids(prior["ID"]))
    ids = s.normalize_ids(base["ID"])
    for col in ["Audit_Status", "Audited_Value", "Deviation", "Tested_By", "Test_Date",
                "Audit_Notes"]:
        if col in base and col in by_id:
            vals = ids.map(by_id[col][~by_id.index.duplicated()])
            base[col] = vals.where(vals.notna(), base[col]).astype(base[col].dtype)

# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
t_plan, t_wp, t_eval, t_pop = st.tabs(["📋 Plan", "✍️ Workpaper", "✅ Evaluate", "📊 Population"])

with t_plan:
    st.dataframe(plan_df, hide_index=True, width="stretch")
    if strata_summary is not None and len(strata_summary):
        st.markdown("#### Strata")
        st.dataframe(strata_summary.style.format(
            {"Value": money, "Lower": money, "Upper": money}), hide_index=True,
            width="stretch")
    st.markdown("#### Selection breakdown")
    st.dataframe(sample[s.REASON].value_counts().rename("Items"), width="stretch")

with t_wp:
    st.caption("Record results here. Changing any parameter or the seed redraws the "
               "sample and clears these entries — export first, then re-upload below to resume.")
    with st.expander("Resume from an exported workpaper"):
        prev = st.file_uploader("Exported workpaper (.xlsx or .csv)", type=["xlsx", "csv"],
                                key="resume_upload")
        if prev is not None and st.button("Load results into this workpaper"):
            prior = (pd.read_csv(prev) if prev.name.endswith(".csv")
                     else pd.read_excel(prev, sheet_name="Workpaper"))
            if "Test_Date" in prior:
                prior["Test_Date"] = pd.to_datetime(prior["Test_Date"], errors="coerce")
            st.session_state[resume_key] = prior
            st.rerun()

    col_cfg = {
        "ID": st.column_config.TextColumn(disabled=True),
        "Name": st.column_config.TextColumn(disabled=True),
        "Book_Value": st.column_config.NumberColumn("Book value", disabled=True, format="%.2f"),
        "Selection_Reason": st.column_config.TextColumn("Selection reason", disabled=True),
        "Stratum": st.column_config.TextColumn(disabled=True),
        "MUS_Hits": st.column_config.NumberColumn("Hits", disabled=True),
        "Audit_Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS,
                                                         required=True),
        "Audited_Value": st.column_config.NumberColumn(
            "Audited value", format="%.2f",
            help="Value supported by audit evidence. Misstatement = book − audited."),
        "Deviation": st.column_config.CheckboxColumn("Deviation?"),
        "Tested_By": st.column_config.TextColumn("Tested by"),
        "Test_Date": st.column_config.DateColumn("Test date"),
        "Audit_Notes": st.column_config.TextColumn("Notes", width="large"),
    }
    wp = st.data_editor(base, key=f"editor_{sig}", column_config=col_cfg, hide_index=True,
                        width="stretch", num_rows="fixed")

    wp_out = wp.copy()
    if monetary:
        wp_out["Misstatement"] = (wp_out["Book_Value"] - wp_out["Audited_Value"]).where(
            wp_out["Audited_Value"].notna())
    wp_out["Test_Date"] = pd.to_datetime(wp_out["Test_Date"]).dt.date

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
        wp_out.to_excel(xw, sheet_name="Workpaper", index=False)
        plan_df.to_excel(xw, sheet_name="Sampling Plan", index=False)
        if strata_summary is not None and len(strata_summary):
            strata_summary.to_excel(xw, sheet_name="Strata", index=False)
        if len(excluded):
            excluded.drop(columns=[s.ID_KEY]).to_excel(xw, sheet_name="Excluded Items",
                                                       index=False)
    e1, e2 = st.columns(2)
    e1.download_button("📊 Download Excel workpaper", xbuf.getvalue(),
                       file_name=f"Audit_Workpaper_{stamp}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch")
    e2.download_button("📥 Download CSV", wp_out.to_csv(index=False).encode("utf-8"),
                       file_name=f"Audit_Workpaper_{stamp}.csv", mime="text/csv",
                       width="stretch")

with t_eval:
    done_mask = wp["Audit_Status"].isin(["Tested – no exception", "Tested – exception"])
    if monetary:
        done_mask &= wp["Audited_Value"].notna()
    done, total = int(done_mask.sum()), len(wp)
    st.progress(done / total if total else 0.0, text=f"{done} of {total} items tested")

    preview = False
    if done < total:
        preview = st.checkbox("Preview anyway, treating untested items as correct",
                              help="For a quick look only — not a basis for a conclusion.")
    if done == total or preview:
        targeted = wp["Selection_Reason"] == "Staff / insider (targeted)"
        if method == "attribute":
            stat = wp[~targeted]
            devs = int((stat["Deviation"] & done_mask[~targeted]).sum())
            ev = s.evaluate_attribute(len(stat), devs, ria, tdr)
        else:
            ev_sample = sample.copy()
            ev_sample["Err"] = (wp["Book_Value"] - wp["Audited_Value"]).fillna(0.0)
            ev_sample = ev_sample[~targeted]  # targeted items are concluded on separately
            if method == "mus":
                ev = s.evaluate_mus(ev_sample, "Err", interval, ria, tolerable)
            elif method == "stratified":
                ev = s.evaluate_ratio(ev_sample, eval_pop, "Err", tolerable, s.STRATUM)
            else:
                ev = s.evaluate_ratio(ev_sample, eval_pop, "Err", tolerable)

        (st.success if ev.supports else st.error)(ev.conclusion)
        if preview:
            st.warning("Preview — untested items were treated as correct.")
        def show(key, val):
            if isinstance(val, int):
                return f"{val:,}"
            if method == "attribute":
                return pct(val)
            return money(val)

        fig = pd.DataFrame({"Measure": list(ev.figures),
                            "Value": [show(k, v) for k, v in ev.figures.items()]})
        st.dataframe(fig, hide_index=True, width="stretch")
        if ev.detail is not None and len(ev.detail):
            st.markdown("#### Detail")
            st.dataframe(ev.detail, hide_index=True, width="stretch")
        if targeted.any():
            st.markdown("#### Staff / insider items (targeted, not projected)")
            st.caption("Selected judgmentally, so they are excluded from the statistical "
                       "evaluation above — conclude on them separately.")
            show_cols = ["ID", "Name", "Book_Value", "Audit_Status"] + (
                ["Audited_Value"] if monetary else ["Deviation"]) + ["Audit_Notes"]
            st.dataframe(wp.loc[targeted, show_cols], hide_index=True, width="stretch")
    else:
        st.info("Finish recording results in the Workpaper tab to evaluate the sample.")

with t_pop:
    q = pop[s.AMT].quantile([0, .25, .5, .75, 1]).to_list()
    qs = sample[s.AMT].quantile([0, .25, .5, .75, 1]).to_list()
    dist = pd.DataFrame({"Statistic": ["Minimum", "Q1", "Median", "Q3", "Maximum"],
                         "Population": [money(v) for v in q],
                         "Sample": [money(v) for v in qs]})
    st.dataframe(dist, hide_index=True, width="stretch")
    if len(excluded):
        st.markdown(f"#### Excluded rows ({len(excluded)})")
        st.dataframe(excluded.drop(columns=[s.ID_KEY]), width="stretch")
