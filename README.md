# 🎯 Audit Sampling & Selection Tool

A Streamlit app for planning, selecting, documenting and evaluating audit samples.
It sizes samples with standard statistical methods, draws a **reproducible** selection
from a seed, gives you an editable workpaper, and evaluates the results when testing
is done.

![tests](https://github.com/Inanewali/Audit_App/actions/workflows/tests.yml/badge.svg)

**Try it in your browser:** [inanewali.github.io/Audit_App](https://inanewali.github.io/Audit_App/) — opens on a synthetic loan book; files you upload stay on your machine.

## What it does

1. **Load a population** — Excel or CSV. Pick the sheet and map the ID, amount and
   name columns (the app guesses sensible defaults).
2. **Clean it, visibly** — total/subtotal rows, non-numeric amounts and credit
   balances are excluded and listed with a reason, never silently dropped. Amounts
   like `$1,234.50` or `(300)` are parsed correctly.
3. **Size and select the sample** with one of four methods (below). Staff / insider
   accounts from a second sheet can be added as targeted items; IDs match even when
   one sheet stores them as numbers and the other as text.
4. **Record results** in an editable workpaper — status, audited value (or deviation),
   tester, date, notes. Export to Excel (workpaper, plan, strata and excluded items
   sheets) or CSV, and re-upload an export later to pick up where you left off.
5. **Evaluate** — once every item is tested, the app projects misstatement and says
   whether the results support the balance or control at your chosen confidence.

## Methods

| Method | Use for | Sample size | Evaluation |
|---|---|---|---|
| **Stratified** | Substantive tests of balances | Key items (≥ threshold) tested 100%; remainder sized as `value × RF / (TM − EM × EF)` and allocated across strata of equal value | Ratio projection per stratum |
| **Monetary unit (MUS)** | Overstatement risk, value-concentrated populations | `BV × RF / (TM − EM × EF)`; systematic PPS selection from a random start; every item ≥ interval is a key item | Upper misstatement limit (basic precision + projected + incremental allowance) |
| **Attribute** | Tests of controls | Smallest *n* where the binomial upper deviation limit stays under the tolerable rate | Exact binomial upper deviation limit |
| **Systematic** | Even coverage of an ordered list | MUS formula (overridable) | Ratio projection |

*RF* = Poisson reliability factor for the risk of incorrect acceptance (3.00 at 95%,
2.31 at 90%), *TM* = tolerable misstatement, *EM* = expected misstatement,
*EF* = expansion factor. Factors and sample sizes are unit-tested against the tables in
the AICPA Audit Guide *Audit Sampling*.

Every selection uses the **random seed** shown in the sidebar and recorded in the
plan, so a reviewer can re-perform it exactly.

## Run it

```bash
git clone https://github.com/Inanewali/Audit_App.git
cd Audit_App
python -m venv venv && source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run audit_app.py
```

Try it with `sample_data/loan_population.xlsx` — a synthetic loan book (fictional
names, random balances) with a `Loans` sheet, a `Staff` sheet, a grand-total row and a
few credit balances. Regenerate it with `python scripts/make_sample_data.py`.

**Deploy:** on [share.streamlit.io](https://share.streamlit.io), point a new app at
this repo with main file `audit_app.py`. Exports are built in memory, so users of a
shared deployment never see each other's files.

## Browser edition

`web/index.html` is a single-file JavaScript port of the same methods, published with
GitHub Pages from the `gh-pages` branch. It needs no server: uploads are read in the
visitor's browser and never leave it. Its sample sizes and limits match `sampling.py`;
its random draws differ for the same seed because the generators differ. To publish a
change, copy `web/index.html` to `index.html` on `gh-pages` and push.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

All statistics live in `sampling.py` (no Streamlit imports); `audit_app.py` is only the
interface.

## Project layout

```
audit_app.py              Streamlit UI
sampling.py               cleaning, sample sizes, selection, evaluation
tests/test_sampling.py    unit tests (AICPA table values, reproducibility, edge cases)
web/index.html            browser edition (published to GitHub Pages)
sample_data/              synthetic demo population
scripts/make_sample_data.py
```

## ⚠️ Client data

Don't commit real engagement data to this repository. `.gitignore` excludes exported
workpapers, but it can't catch everything — keep client files outside the repo folder.

## Limitations

* The stratified and systematic evaluations give the projected (most likely)
  misstatement; the allowance for sampling risk is left to the auditor's judgment.
* MUS evaluation handles overstatements statistically; understatements are shown
  separately and not built into the upper limit.
* Sampling doesn't replace professional judgment on materiality, risk or whether a
  population is complete.

## License

MIT — see [LICENSE](LICENSE).
