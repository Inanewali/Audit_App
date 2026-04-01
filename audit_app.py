import streamlit as st
import pandas as pd

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="Audit Workpaper Tool", layout="wide")
st.title("📑 Automated Substantive Testing: 60-Item Sample")
st.markdown("Selects up to **60 items** using stratified audit sampling (Key Items → Staff Loans → Random).")

# --- 2. SIDEBAR ---
st.sidebar.header("Audit Parameters")
sample_size_target = 60

# --- 3. DATA LOADING ---
uploaded_file = st.file_uploader("Upload Master Excel File", type=["xlsx"])

if uploaded_file:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        sheets = excel_obj.sheet_names

        default_loan  = sheets.index("Loans") if "Loans" in sheets else 0
        default_staff = sheets.index("Staff") if "Staff" in sheets else min(1, len(sheets)-1)

        loan_sheet  = st.selectbox("Select Loans Sheet",  sheets, index=default_loan)
        staff_sheet = st.selectbox("Select Staff Sheet",  sheets, index=default_staff)

        df = pd.read_excel(
            uploaded_file, sheet_name=loan_sheet,
            dtype={"ACCOUNT ID": str, "CUSTOMER": str}
        )
        df.columns = df.columns.str.strip()

        staff_df = pd.read_excel(
            uploaded_file, sheet_name=staff_sheet,
            dtype={"ACCOUNT ID": str}
        )
        staff_df.columns = staff_df.columns.str.strip()

        ID_COL    = "ACCOUNT ID"
        AMT_COL   = "TOTAL OUTSTANDING"
        TITLE_COL = "CUSTOMER NAME"

        # Validate required columns exist
        missing = [c for c in [ID_COL, AMT_COL, TITLE_COL] if c not in df.columns]
        if missing:
            st.error(f"Missing columns in Loans sheet: {missing}. Found: {df.columns.tolist()}")
            st.stop()

        # Clean amount column
        df["Amt_Clean"] = pd.to_numeric(
            df[AMT_COL].astype(str).str.replace(r"[\$,]", "", regex=True),
            errors="coerce"
        ).fillna(0)

        # --- AUTO THRESHOLD: 80% of highest outstanding balance ---
        max_value          = df["Amt_Clean"].max()
        high_val_threshold = max_value * 0.80

        # Show threshold info in sidebar now that data is loaded
        st.sidebar.markdown("---")
        st.sidebar.metric("Highest Loan Balance", f"${max_value:,.2f}")
        st.sidebar.metric("High Value Threshold (80%)", f"${high_val_threshold:,.2f}")
        st.sidebar.caption("Threshold is automatically set to 80% of the highest loan balance in the population.")

        # --- 4. SELECTION LOGIC ---
        # A. High Value Key Items
        key_items = df[df["Amt_Clean"] >= high_val_threshold].copy()
        key_items["Selection_Reason"] = "High Value Key Item"

        # B. Staff / Insider Loans
        insider_loans = df[df[ID_COL].isin(staff_df[ID_COL])].copy()
        insider_loans["Selection_Reason"] = "Staff Loan"

        # C. Random sample from remaining pool
        picked_ids     = pd.concat([key_items, insider_loans])[ID_COL].unique()
        remaining_pool = df[~df[ID_COL].isin(picked_ids)].copy()

        already_picked  = len(pd.concat([key_items, insider_loans]).drop_duplicates(subset=[ID_COL]))
        n_random_needed = max(0, sample_size_target - already_picked)
        n_random_actual = min(len(remaining_pool), n_random_needed)

        if n_random_actual > 0:
            random_samples = remaining_pool.sample(n=n_random_actual, random_state=42)
            random_samples["Selection_Reason"] = "Random Statistical Sample"
        else:
            random_samples = pd.DataFrame(columns=df.columns)

        # --- 5. CONSOLIDATE ---
        final_output = pd.concat([key_items, insider_loans, random_samples]) \
                         .drop_duplicates(subset=[ID_COL])
        final_output["Audit_Status"] = "Pending"
        final_output["Audit_Notes"]  = ""

        # --- 6. UI DISPLAY ---
        st.write("### 🔍 Population & Sample Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Population", f"{len(df):,}")
        c2.metric("High Value Items",  f"{len(key_items):,}")
        c3.metric("Staff Loan Items",  f"{len(insider_loans):,}")
        c4.metric("Random Sample",     f"{len(random_samples):,}")

        total_sampled = len(final_output)
        coverage_pct  = total_sampled / len(df) * 100 if len(df) else 0

        if total_sampled < sample_size_target:
            st.warning(
                f"⚠️ Sample contains {total_sampled} items (target: {sample_size_target}). "
                f"Only {len(remaining_pool)} rows were left after removing Key/Staff items. "
                f"This is expected when Key Items cover most of the population."
            )
        else:
            st.success(f"✅ {total_sampled}-item sample selected ({coverage_pct:.1f}% population coverage).")

        st.markdown("---")
        st.subheader("✍️ Digital Audit Workpaper")
        st.info("Review the sample below, then export to CSV to fill in your audit notes.")

        # Build display DataFrame — kept separate from st.dataframe() which returns None
        display_cols = [ID_COL, TITLE_COL, AMT_COL, "Selection_Reason", "Audit_Status", "Audit_Notes"]
        export_df = final_output[display_cols].reset_index(drop=True)

        st.dataframe(export_df)

        # --- 7. DOWNLOAD ---
        csv_data = export_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Export Workpaper to CSV",
            data=csv_data,
            file_name="Audit_Workpaper.csv",
            mime="text/csv"
        )

        st.markdown("---")
        st.caption(
            "Tip: Export the CSV, open in Excel, fill in the Audit_Status and Audit_Notes columns, "
            "then save as your final workpaper."
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
        st.exception(e)

else:
    st.info("⬆️ Upload the Excel file to generate your audit sample.")