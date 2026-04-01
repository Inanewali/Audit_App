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
        default_staff = sheets.index("Staff") if "Staff" in sheets else min(1, len(sheets)-1) if len(sheets) > 1 else None

        loan_sheet  = st.selectbox("Select Loans Sheet",  sheets, index=default_loan)
        
        # Optional staff sheet
        include_staff = st.checkbox("Include Staff/Insider Loans", value=default_staff is not None)
        staff_sheet = None
        if include_staff:
            staff_options = ["(None)"] + sheets
            default_staff_idx = sheets.index(default_staff) + 1 if default_staff and default_staff in sheets else 1
            staff_sheet = st.selectbox("Select Staff Sheet", staff_options, index=default_staff_idx)
            if staff_sheet == "(None)":
                staff_sheet = None

        df = pd.read_excel(uploaded_file, sheet_name=loan_sheet)
        df.columns = df.columns.str.strip()

        staff_df = None
        if staff_sheet:
            staff_df = pd.read_excel(uploaded_file, sheet_name=staff_sheet)
            staff_df.columns = staff_df.columns.str.strip()

        # --- COLUMN SELECTION ---
        st.markdown("### 🔧 Column Mapping")
        st.info("Select which columns from your data correspond to these fields:")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            ID_COL = st.selectbox(
                "Account/ID Column",
                df.columns,
                index=0,
                help="Unique identifier for each loan/account"
            )
        
        with col2:
            AMT_COL = st.selectbox(
                "Amount Column",
                df.columns,
                index=min(1, len(df.columns)-1),
                help="The amount field to use for stratification"
            )
        
        with col3:
            TITLE_COL = st.selectbox(
                "Name/Description Column",
                df.columns,
                index=min(2, len(df.columns)-1),
                help="Customer name or other identifier"
            )
        
        if staff_df is not None and len(staff_df.columns) > 0:
            st.markdown("#### Staff Sheet ID Column")
            STAFF_ID_COL = st.selectbox(
                "Staff Sheet ID Column",
                staff_df.columns,
                help="The ID column in the staff sheet to match against loans"
            )
        else:
            STAFF_ID_COL = ID_COL

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
        st.sidebar.metric("Highest Amount", f"${max_value:,.2f}")
        st.sidebar.metric("High Value Threshold (80%)", f"${high_val_threshold:,.2f}")
        st.sidebar.caption("Threshold is automatically set to 80% of the highest amount in the population.")

        # --- 4. SELECTION LOGIC ---
        # A. High Value Key Items
        key_items = df[df["Amt_Clean"] >= high_val_threshold].copy()
        key_items["Selection_Reason"] = "High Value Key Item"

        # B. Staff / Insider Loans
        if staff_df is not None and len(staff_df) > 0:
            insider_loans = df[df[ID_COL].isin(staff_df[STAFF_ID_COL])].copy()
            insider_loans["Selection_Reason"] = "Staff Loan"
        else:
            insider_loans = pd.DataFrame(columns=df.columns)

        # C. Random sample from remaining pool
        combined = pd.concat([key_items, insider_loans]) if len(insider_loans) > 0 else key_items
        picked_ids     = combined[ID_COL].unique()
        remaining_pool = df[~df[ID_COL].isin(picked_ids)].copy()

        already_picked  = len(combined.drop_duplicates(subset=[ID_COL]))
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

        # Build display DataFrame with selected columns
        display_cols = [ID_COL, TITLE_COL, AMT_COL, "Selection_Reason", "Audit_Status", "Audit_Notes"]
        # Only include columns that exist in final_output
        display_cols = [col for col in display_cols if col in final_output.columns]
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