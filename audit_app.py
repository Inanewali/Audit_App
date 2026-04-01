import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import math

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Advanced Audit Sampling Tool", layout="wide")
st.title("🎯 Advanced Audit Sampling & Selection Tool (AAST)")
st.markdown("Professional-grade audit sampling with multiple methodologies, statistical analysis, and compliance documentation.")

# --- 2. SIDEBAR - AUDIT PARAMETERS ---
st.sidebar.header("⚙️ Audit Configuration")

# Sampling Method Selection
sampling_method = st.sidebar.radio(
    "Sampling Method",
    ["Stratified Random", "Monetary Unit Sampling (MUS)", "Attribute Sampling", "Systematic"],
    help="Choose the appropriate statistical sampling method for your audit"
)

# Statistical Parameters
st.sidebar.markdown("### Statistical Parameters")

confidence_level = st.sidebar.select_slider(
    "Confidence Level",
    options=[90, 95, 99],
    value=95,
    help="Higher confidence = larger sample size required"
)

risk_ia = st.sidebar.slider(
    "Risk of Incorrect Acceptance (RIA)",
    min_value=0.05,
    max_value=0.25,
    value=0.10,
    step=0.01,
    help="β risk: Probability of accepting an incorrect assertion (Type II error)"
)

risk_ir = st.sidebar.slider(
    "Risk of Incorrect Rejection (RIR)",
    min_value=0.01,
    max_value=0.10,
    value=0.05,
    step=0.01,
    help="α risk: Probability of rejecting a correct assertion (Type I error)"
)

# Materiality Settings
st.sidebar.markdown("### Materiality & Error Tolerance")

materiality_type = st.sidebar.radio(
    "Materiality Basis",
    ["Dollar Amount", "Percentage of Population"],
    help="How to define the materiality threshold"
)

if materiality_type == "Dollar Amount":
    materiality = st.sidebar.number_input(
        "Materiality Threshold ($)",
        min_value=0.0,
        value=10000.0,
        step=1000.0,
        help="Maximum acceptable error amount"
    )
else:
    materiality_pct = st.sidebar.slider(
        "Materiality Threshold (%)",
        min_value=0.5,
        max_value=10.0,
        value=2.0,
        step=0.5,
        help="Materiality as % of total population value"
    )
    materiality = None

expected_error_rate = st.sidebar.slider(
    "Expected Error Rate (%)",
    min_value=0.0,
    max_value=5.0,
    value=0.5,
    step=0.1,
    help="Anticipated percentage of items with errors"
)

# Sample Size Constraints
st.sidebar.markdown("### Sample Size Bounds")
min_sample_size = st.sidebar.number_input("Minimum Sample Size", min_value=10, value=30, step=5)
max_sample_size = st.sidebar.number_input("Maximum Sample Size", min_value=50, value=500, step=10)

# --- 3. SAMPLE SIZE CALCULATION FORMULAS ---
def calculate_z_score(confidence_level):
    """Get z-score from confidence level"""
    z_scores = {90: 1.645, 95: 1.96, 99: 2.576}
    return z_scores.get(confidence_level, 1.96)

def calculate_sample_size_stratified(population_size, confidence_level, expected_error_rate, risk_ir, risk_ia):
    """Calculate sample size for stratified sampling (AICPA Standard AU-C 530)"""
    z_alpha = calculate_z_score(confidence_level)
    z_beta = calculate_z_score(100 - (risk_ia * 100))
    
    p = expected_error_rate / 100
    if p == 0:
        p = 0.01  # Avoid division by zero
    
    # Cochran's formula for finite populations
    n = ((z_alpha + z_beta) ** 2 * p * (1 - p)) / ((1 - 2*p) ** 2)
    n = int(math.ceil(n))
    
    # Finite population correction
    if population_size > 0:
        n = int(math.ceil(n / (1 + (n / population_size))))
    
    return max(min_sample_size, min(int(n), max_sample_size))

def calculate_sample_size_mus(book_value, tolerable_error, risk_ir):
    """Calculate sample size for Monetary Unit Sampling (MUS)"""
    # MUS formula: Sample Size = Book Value × Factor / Tolerable Error
    factor = calculate_z_score(100 - (risk_ir * 100))
    sample_size = int(math.ceil((book_value * factor) / tolerable_error)) if tolerable_error > 0 else max_sample_size
    return max(min_sample_size, min(sample_size, max_sample_size))

def calculate_sample_size_attribute(population_size, confidence_level, expected_error_rate, risk_ir):
    """Calculate sample size for Attribute Sampling"""
    z_alpha = calculate_z_score(confidence_level)
    p = expected_error_rate / 100
    e = 0.05  # Acceptable deviation rate
    
    if e <= 0:
        return max_sample_size
    
    n = ((z_alpha ** 2) * p * (1 - p)) / (e ** 2)
    n = int(math.ceil(n))
    
    if population_size > 0:
        n = int(math.ceil(n / (1 + (n / population_size))))
    
    return max(min_sample_size, min(int(n), max_sample_size))

# --- 4. DATA LOADING ---
uploaded_file = st.file_uploader("Upload Master Excel File", type=["xlsx"])

if uploaded_file:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        sheets = excel_obj.sheet_names

        default_loan  = sheets.index("Loans") if "Loans" in sheets else 0
        default_staff = sheets.index("Staff") if "Staff" in sheets else (min(1, len(sheets)-1) if len(sheets) > 1 else None)

        # Sheet Selection
        st.markdown("### 📋 Data Source Configuration")
        col1, col2 = st.columns(2)
        
        with col1:
            loan_sheet  = st.selectbox("Select Loans/Population Sheet", sheets, index=default_loan)
        
        with col2:
            include_staff = st.checkbox("Include Staff/Insider Loans", value=default_staff is not None)

        df = pd.read_excel(uploaded_file, sheet_name=loan_sheet)
        df.columns = df.columns.str.strip()

        staff_df = None
        if include_staff:
            staff_options = ["(None)"] + sheets
            default_staff_idx = sheets.index(default_staff) + 1 if default_staff and default_staff in sheets else 1
            staff_sheet = st.selectbox("Select Staff Sheet", staff_options, index=default_staff_idx)
            if staff_sheet != "(None)":
                staff_df = pd.read_excel(uploaded_file, sheet_name=staff_sheet)
                staff_df.columns = staff_df.columns.str.strip()

        # --- COLUMN SELECTION ---
        st.markdown("### 🔧 Column Mapping")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            ID_COL = st.selectbox("Account/ID Column", df.columns, index=0)
        
        with col2:
            AMT_COL = st.selectbox("Amount Column", df.columns, index=min(1, len(df.columns)-1))
        
        with col3:
            TITLE_COL = st.selectbox("Name/Description Column", df.columns, index=min(2, len(df.columns)-1))
        
        STAFF_ID_COL = ID_COL
        if staff_df is not None and len(staff_df.columns) > 0:
            STAFF_ID_COL = st.selectbox("Staff Sheet ID Column", staff_df.columns, label_visibility="collapsed")

        # Data Cleaning
        df["Amt_Clean"] = pd.to_numeric(
            df[AMT_COL].astype(str).str.replace(r"[\$,]", "", regex=True),
            errors="coerce"
        ).fillna(0)
        
        # Remove negative amounts
        df = df[df["Amt_Clean"] >= 0].copy()
        
        population_size = len(df)
        total_population_value = df["Amt_Clean"].sum()
        max_value = df["Amt_Clean"].max()
        min_value = df[df["Amt_Clean"] > 0]["Amt_Clean"].min() if len(df[df["Amt_Clean"] > 0]) > 0 else 0

        # --- CALCULATE TOLERABLE ERROR ---
        if materiality_type == "Percentage of Population" and materiality is None:
            materiality = (materiality_pct / 100) * total_population_value

        # --- DISPLAY POPULATION STATISTICS ---
        st.sidebar.markdown("---")
        st.sidebar.markdown("### Population Statistics")
        st.sidebar.metric("Population Size", f"{population_size:,}")
        st.sidebar.metric("Total Value", f"${total_population_value:,.2f}")
        st.sidebar.metric("Max Item Value", f"${max_value:,.2f}")
        st.sidebar.metric("Avg Item Value", f"${total_population_value/population_size:,.2f}" if population_size > 0 else "$0")
        st.sidebar.metric("Materiality Threshold", f"${materiality:,.2f}")

        # --- CALCULATE SAMPLE SIZE ---
        if sampling_method == "Stratified Random":
            sample_size_target = calculate_sample_size_stratified(population_size, confidence_level, expected_error_rate, risk_ir, risk_ia)
        elif sampling_method == "Monetary Unit Sampling (MUS)":
            sample_size_target = calculate_sample_size_mus(total_population_value, materiality, risk_ir)
        elif sampling_method == "Attribute Sampling":
            sample_size_target = calculate_sample_size_attribute(population_size, confidence_level, expected_error_rate, risk_ir)
        else:  # Systematic
            sample_size_target = calculate_sample_size_stratified(population_size, confidence_level, expected_error_rate, risk_ir, risk_ia)

        st.sidebar.metric("Calculated Sample Size", sample_size_target)

        # --- 5. SELECTION LOGIC BY METHOD ---
        if sampling_method == "Stratified Random":
            # High value threshold = 75th percentile
            high_val_threshold = df["Amt_Clean"].quantile(0.75)
            
            key_items = df[df["Amt_Clean"] >= high_val_threshold].copy()
            key_items["Selection_Reason"] = "High Value Stratum"
            key_items["Sampling_Method"] = "Stratified"

            # Staff loans
            if staff_df is not None and len(staff_df) > 0:
                insider_loans = df[df[ID_COL].isin(staff_df[STAFF_ID_COL])].copy()
                insider_loans["Selection_Reason"] = "Staff/Insider"
                insider_loans["Sampling_Method"] = "Stratified"
            else:
                insider_loans = pd.DataFrame(columns=df.columns)

            # Random from remainder
            combined = pd.concat([key_items, insider_loans]) if len(insider_loans) > 0 else key_items
            picked_ids = combined[ID_COL].unique()
            remaining_pool = df[~df[ID_COL].isin(picked_ids)].copy()

            already_picked = len(combined.drop_duplicates(subset=[ID_COL]))
            n_random_needed = max(0, sample_size_target - already_picked)
            n_random_actual = min(len(remaining_pool), n_random_needed)

            if n_random_actual > 0:
                random_samples = remaining_pool.sample(n=n_random_actual, random_state=42)
                random_samples["Selection_Reason"] = "Random Sample"
                random_samples["Sampling_Method"] = "Stratified"
            else:
                random_samples = pd.DataFrame(columns=df.columns)

            final_output = pd.concat([key_items, insider_loans, random_samples]).drop_duplicates(subset=[ID_COL])

        elif sampling_method == "Monetary Unit Sampling (MUS)":
            # MUS: Cumulative probability sampling based on amount
            df["Cumulative_Amount"] = df["Amt_Clean"].cumsum()
            df["Cumulative_Pct"] = (df["Cumulative_Amount"] / total_population_value * 100) if total_population_value > 0 else 0
            
            # Select items across cumulative distribution
            interval = total_population_value / sample_size_target if sample_size_target > 0 else 0
            selected_indices = []
            cumulative = 0
            
            for idx, row in df.iterrows():
                if row["Amt_Clean"] > 0:
                    cumulative += row["Amt_Clean"]
                    if (cumulative // interval) > len(selected_indices):
                        selected_indices.append(idx)
                    
                    if len(selected_indices) >= sample_size_target:
                        break
            
            # Also include all items above 2x the interval (key items)
            key_threshold = interval * 2
            key_items = df[df["Amt_Clean"] >= key_threshold].copy()
            
            final_output = df.loc[selected_indices].copy()
            final_output["Selection_Reason"] = "MUS Selection"
            final_output["Sampling_Method"] = "MUS"
            
            # Add key items if not already selected
            key_not_selected = key_items[~key_items[ID_COL].isin(final_output[ID_COL])]
            if len(key_not_selected) > 0:
                key_not_selected["Selection_Reason"] = "Key Item (>2x interval)"
                key_not_selected["Sampling_Method"] = "MUS"
                final_output = pd.concat([final_output, key_not_selected])
            
            final_output = final_output.drop_duplicates(subset=[ID_COL])

        elif sampling_method == "Systematic":
            # Systematic: Every kth item
            k = max(1, population_size // sample_size_target)
            start = np.random.randint(0, k)
            systematic_indices = list(range(start, population_size, k))[:sample_size_target]
            
            final_output = df.iloc[systematic_indices].copy()
            final_output["Selection_Reason"] = f"Systematic (k={k})"
            final_output["Sampling_Method"] = "Systematic"

        else:  # Attribute Sampling
            # Similar to stratified but focused on high-risk items
            high_risk_threshold = df["Amt_Clean"].quantile(0.80)
            high_risk = df[df["Amt_Clean"] >= high_risk_threshold].copy()
            high_risk["Selection_Reason"] = "High Risk Item"
            high_risk["Sampling_Method"] = "Attribute"
            
            remaining = df[~df[ID_COL].isin(high_risk[ID_COL])].sample(
                n=min(sample_size_target - len(high_risk), len(df) - len(high_risk)),
                random_state=42
            )
            remaining["Selection_Reason"] = "Random Selection"
            remaining["Sampling_Method"] = "Attribute"
            
            final_output = pd.concat([high_risk, remaining]).drop_duplicates(subset=[ID_COL])

        # Add audit tracking columns
        final_output["Audit_Status"] = "Pending"
        final_output["Audit_Notes"] = ""
        final_output["Tested_By"] = ""
        final_output["Test_Date"] = ""
        final_output["Error_Found"] = "No"
        final_output["Error_Amount"] = 0.0

        # --- 6. STATISTICAL ANALYSIS ---
        total_sampled = len(final_output)
        sampled_value = final_output["Amt_Clean"].sum()
        coverage_pct = (total_sampled / population_size * 100) if population_size > 0 else 0
        value_coverage_pct = (sampled_value / total_population_value * 100) if total_population_value > 0 else 0
        
        # Precision calculation
        precision = materiality / 2 if materiality > 0 else 0
        precision_pct = (precision / total_population_value * 100) if total_population_value > 0 else 0
        
        # --- 7. UI DISPLAY ---
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Sampling Plan", "🎯 Sample Details", "📈 Statistics", "📄 Workpaper"])
        
        with tab1:
            st.write("### 📋 Audit Sampling Plan Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Sampling Method", sampling_method.split("(")[0].strip())
            col2.metric("Confidence Level", f"{confidence_level}%")
            col3.metric("RIA", f"{risk_ia*100:.1f}%")
            col4.metric("RIR", f"{risk_ir*100:.1f}%")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Population Size", f"{population_size:,}")
            col2.metric("Sample Size", f"{total_sampled:,}")
            col3.metric("Coverage %", f"{coverage_pct:.2f}%")
            col4.metric("Materiality", f"${materiality:,.2f}")
            
            # Sampling Plan Details
            st.markdown("#### Sampling Plan Details")
            sampling_details = {
                "Population Size": population_size,
                "Total Population Value": f"${total_population_value:,.2f}",
                "Sample Size (Calculated)": sample_size_target,
                "Sample Size (Actual)": total_sampled,
                "Sample Value": f"${sampled_value:,.2f}",
                "Sampling Method": sampling_method,
                "Confidence Level": f"{confidence_level}%",
                "Risk of Incorrect Acceptance": f"{risk_ia*100:.2f}%",
                "Risk of Incorrect Rejection": f"{risk_ir*100:.2f}%",
                "Expected Error Rate": f"{expected_error_rate:.2f}%",
                "Materiality Threshold": f"${materiality:,.2f}",
                "Precision": f"${precision:,.2f}",
                "Sampling Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            st.dataframe(pd.DataFrame(sampling_details.items(), columns=["Parameter", "Value"]), hide_index=True, use_container_width=True)
        
        with tab2:
            st.write("### 📋 Sample Items Details")
            
            # Build display dataframe
            display_cols = [ID_COL, TITLE_COL, AMT_COL, "Selection_Reason", "Sampling_Method"]
            display_cols = [col for col in display_cols if col in final_output.columns]
            display_df = final_output[display_cols].reset_index(drop=True)
            display_df.index = display_df.index + 1  # 1-based indexing
            
            st.dataframe(display_df, use_container_width=True)
            
            # Summary by selection reason
            st.markdown("#### Selection Breakdown")
            reason_summary = final_output["Selection_Reason"].value_counts().to_frame().rename(columns={"Selection_Reason": "Count"})
            st.dataframe(reason_summary, use_container_width=True)
        
        with tab3:
            st.write("### 📈 Statistical Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Sample Composition")
                stats_data = {
                    "Metric": [
                        "Population Count",
                        "Sample Count",
                        "Sample %",
                        "Population Value",
                        "Sample Value",
                        "Value %"
                    ],
                    "Value": [
                        f"{population_size:,}",
                        f"{total_sampled:,}",
                        f"{coverage_pct:.2f}%",
                        f"${total_population_value:,.2f}",
                        f"${sampled_value:,.2f}",
                        f"{value_coverage_pct:.2f}%"
                    ]
                }
                st.dataframe(pd.DataFrame(stats_data), hide_index=True, use_container_width=True)
            
            with col2:
                st.markdown("#### Risk Assessment")
                risk_data = {
                    "Risk Factor": [
                        "Risk of Incorrect Acceptance",
                        "Risk of Incorrect Rejection",
                        "Expected Error Rate",
                        "Materiality %",
                        "Precision %"
                    ],
                    "Value": [
                        f"{risk_ia*100:.2f}%",
                        f"{risk_ir*100:.2f}%",
                        f"{expected_error_rate:.2f}%",
                        f"{(materiality/total_population_value*100):.2f}%" if total_population_value > 0 else "0%",
                        f"{precision_pct:.2f}%"
                    ]
                }
                st.dataframe(pd.DataFrame(risk_data), hide_index=True, use_container_width=True)
            
            # Distribution analysis
            st.markdown("#### Sample Distribution by Value")
            
            def get_amount_quantile(df, amt_col):
                return {
                    "Minimum": df[amt_col].min(),
                    "Q1 (25%)": df[amt_col].quantile(0.25),
                    "Median": df[amt_col].quantile(0.50),
                    "Q3 (75%)": df[amt_col].quantile(0.75),
                    "Maximum": df[amt_col].max()
                }
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Population Distribution**")
                pop_dist = get_amount_quantile(df, "Amt_Clean")
                for k, v in pop_dist.items():
                    st.text(f"{k}: ${v:,.2f}")
            
            with col2:
                st.markdown("**Sample Distribution**")
                sample_dist = get_amount_quantile(final_output, "Amt_Clean")
                for k, v in sample_dist.items():
                    st.text(f"{k}: ${v:,.2f}")
        
        with tab4:
            st.write("### ✍️ Digital Audit Workpaper")
            
            # Export-ready workpaper
            workpaper_cols = [ID_COL, TITLE_COL, AMT_COL, "Selection_Reason", "Sampling_Method", 
                            "Audit_Status", "Error_Found", "Error_Amount", "Tested_By", "Test_Date", "Audit_Notes"]
            workpaper_cols = [col for col in workpaper_cols if col in final_output.columns]
            export_df = final_output[workpaper_cols].copy()
            export_df.index = export_df.index + 1
            
            st.dataframe(export_df, use_container_width=True)
            
            # Export buttons
            csv_data = export_df.to_csv().encode("utf-8")
            st.download_button(
                label="📥 Export Workpaper to CSV",
                data=csv_data,
                file_name=f"Audit_Workpaper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            
            xlsx_buffer = pd.ExcelWriter('audit_workpaper.xlsx', engine='openpyxl')
            export_df.to_excel(xlsx_buffer, sheet_name='Sample Items')
            
            # Add summary sheet
            summary_df = pd.DataFrame(sampling_details.items(), columns=["Parameter", "Value"])
            summary_df.to_excel(xlsx_buffer, sheet_name='Sampling Plan', index=False)
            
            xlsx_buffer.close()
            
            with open('audit_workpaper.xlsx', 'rb') as f:
                st.download_button(
                    label="📊 Export to Excel with Summary",
                    data=f.read(),
                    file_name=f"Audit_Workpaper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.exception(e)

else:
    st.info("⬆️ Upload the Excel file to begin your audit sampling process.")