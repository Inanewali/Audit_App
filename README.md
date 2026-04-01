# 🎯 Advanced Audit Sampling & Selection Tool (AAST)

**Professional-grade audit sampling tool** with multiple statistical methodologies, comprehensive analysis, and audit workpaper generation.

## Overview

AAST is a web-based application designed for auditors to perform statistically valid sample selection and population analysis. It supports multiple sampling methods, customizable risk parameters, and generates professional audit documentation.

**Live Demo:** [Deploy on Streamlit Cloud](https://share.streamlit.io)

---

## ✨ Features

### Sampling Methods
- **Stratified Random Sampling** — Risk-based stratification by item value
- **Monetary Unit Sampling (MUS)** — Probability-proportional-to-size selection
- **Attribute Sampling** — Compliance and operational testing
- **Systematic Sampling** — Every kth item interval selection

### Statistical Analysis
- **Confidence Levels** — 90%, 95%, 99% options
- **Risk Parameters** — Customizable RIA (β) and RIR (α)
- **Sample Size Calculations** — AICPA AU-C 530 compliant
- **Materiality Thresholds** — Dollar or percentage-based
- **Precision Analysis** — Statistical accuracy measures
- **Population Statistics** — Mean, median, quartiles, distribution

### Professional Output
- **4 Analysis Tabs:**
  - Sampling Plan — Complete audit plan documentation
  - Sample Details — Itemized selection list with reasons
  - Statistics — Statistical analysis and risk assessment
  - Workpaper — Editable audit tracking fields

### Export Capabilities
- **CSV Export** — For Excel editing
- **Excel Export** — Dual sheets (Sample + Sampling Plan)
- **Automatic Timestamps** — Audit trail documentation

---

## 🚀 Quick Start

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Inanewali/Audit_App.git
   cd Audit_App
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   streamlit run audit_app.py
   ```

4. **Open in browser:**
   ```
   http://localhost:8501
   ```

### Web Deployment

**Deploy on Streamlit Cloud (Free):**
1. Push code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Select repository: `Inanewali/Audit_App`
4. Main file: `audit_app.py`
5. Click Deploy

---

## 📖 Usage Guide

### Step 1: Configure Audit Parameters
In the left sidebar, set:
- **Sampling Method** — Choose methodology
- **Confidence Level** — 90%, 95%, or 99%
- **Risk of Incorrect Acceptance (RIA)** — Typically 5-10%
- **Risk of Incorrect Rejection (RIR)** — Typically 1-5%
- **Expected Error Rate** — Anticipated % of errors
- **Materiality Threshold** — Maximum acceptable error

### Step 2: Upload Data
1. Click "Upload Master Excel File"
2. Select your audit population file (XLSX format)

### Step 3: Map Columns
Select which columns in your file correspond to:
- **Account/ID Column** — Unique identifier
- **Amount Column** — Value for stratification
- **Name/Description** — Customer or account name
- **Staff Sheet ID** — If testing related parties

### Step 4: Review & Export
- **Sampling Plan Tab** — Review audit parameters
- **Sample Details Tab** — See selected items with reasons
- **Statistics Tab** — Analyze sample composition
- **Workpaper Tab** — Export and document findings

---

## 📊 Sampling Methods Explained

### Stratified Random Sampling
**When to use:** Most common audit sampling method

- Population divided into strata (layers) by value
- High-value items selected separately
- Random selection from remaining population
- Best for: General substantive testing

**Formula:**
```
n = ((z_α + z_β)² × p × (1-p)) / ((1-2p)²)
```
Where: p = expected error rate, z = confidence factor

---

### Monetary Unit Sampling (MUS)
**When to use:** When you want sample size based on dollars, not items

- Selects items proportional to their value
- Higher-value items more likely to be selected
- Automatically includes key items
- Best for: Dollar-based audits, high-risk assertions

**Key Intervals:**
```
Interval = Total Population Value / Sample Size
```

---

### Attribute Sampling
**When to use:** Testing compliance with policies/procedures

- Focuses on high-risk/high-value items first
- Tests presence or absence of attributes
- Lower precision than monetary sampling
- Best for: Control testing, compliance audits

---

### Systematic Sampling
**When to use:** When you need even distribution

- Every kth item selected
- Interval: Population Size / Sample Size
- Random start point prevents bias
- Best for: Ordered populations, initial screening

---

## 📊 Output Statistics

### Sample Composition
- **Population Size** — Total items in audit universe
- **Sample Size** — Items to be tested
- **Coverage %** — Sample as % of population
- **Value Coverage %** — Sample value as % of total

### Risk Assessment
- **RIA (β)** — Risk of accepting incorrect assertion
- **RIR (α)** — Risk of rejecting correct assertion
- **Materiality %** — Threshold as % of total value
- **Precision %** — Acceptable error range

### Distribution Analysis
- **Minimum** — Smallest item value
- **Q1 (25%)** — First quartile
- **Median (50%)** — Middle value
- **Q3 (75%)** — Third quartile
- **Maximum** — Largest item value

---

## 📋 Workpaper Fields

The exported workpaper includes:
| Field | Purpose |
|-------|---------|
| **Account ID** | Unique identifier |
| **Item Description** | Customer/account name |
| **Amount** | Transaction value |
| **Selection Reason** | Why item was selected |
| **Sampling Method** | Method used |
| **Audit Status** | Pending/Completed/Exception |
| **Error Found** | Yes/No indicator |
| **Error Amount** | $ amount of exception |
| **Tested By** | Auditor name |
| **Test Date** | Date of audit procedure |
| **Audit Notes** | Findings/observations |

---

## 🔧 Technical Details

### Requirements
- Python 3.8+
- Streamlit >= 1.28.0
- Pandas >= 2.0.0
- NumPy >= 1.24.0
- OpenPyXL >= 3.1.0

### File Format
- **Input:** Excel (.xlsx) with numeric IDs and amounts
- **Output:** CSV or Excel workpapers
- **Max File Size:** 50MB (Streamlit Cloud limit)

### Auditor Tips
1. **Stratified Sampling** — Best for general substantive testing
2. **MUS** — Use when population has high value concentration
3. **Attribute Sampling** — Use for control testing
4. **Sample Size** — Tool auto-calculates; adjust confidence/risk as needed
5. **Materiality** — Professional judgment required; tool provides framework

---

## 🎓 AICPA Standards

This tool implements principles from:
- **AU-C 530** — Audit Sampling
- **AU-C 320** — Materiality in Planning and Performing an Audit
- **AU-C 315** — Understanding the Entity and Its Environment

---

## 📝 Example Workflow

```
1. Load loan population (500 items, $50M total value)
2. Set confidence = 95%, RIA = 10%, materiality = 2%
3. Select Stratified Random Sampling
4. Tool calculates sample size = 42 items
5. Selects: 8 high-value, 34 random
6. Export to Excel
7. Assign to team for testing
8. Fill in audit findings
9. Archive workpaper
```

---

## 🐛 Troubleshooting

| Issue | Solution |
|-------|----------|
| **File won't upload** | Ensure file is .xlsx (not .xls or .csv) |
| **Columns missing** | Check that column names match invoice/account data |
| **Sample too large** | Increase confidence level or lower materiality |
| **Sample too small** | Decrease confidence level or increase materiality |
| **Error with amounts** | Ensure all values are numeric (remove $ signs, commas) |

---

## 📞 Support & Feedback

- **Issues:** [GitHub Issues](https://github.com/Inanewali/Audit_App/issues)
- **Suggestions:** [GitHub Discussions](https://github.com/Inanewali/Audit_App/discussions)

---

## 📄 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **AICPA** — Audit sampling standards and guidance
- **Streamlit** — Web application framework
- **Pandas** — Data analysis and manipulation

---

**Version:** 2.0 | **Last Updated:** April 1, 2026

*Professional Audit Sampling | Statistical Analysis | Compliance Documentation*
