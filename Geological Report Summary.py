import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Excel serial date to datetime
def excel_date(num):
    try:
        return datetime(1899, 12, 30) + timedelta(days=int(num))
    except:
        return "N/A"

st.set_page_config(page_title="Daily Geological Report Analyzer", layout="wide")
st.title("🏢 Daily Geological Report Analyzer")
st.markdown("Upload your **AB-xxx, DGR-xxx.xlsx** file (3 sheets) and get a merged professional summary")

uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load all 3 sheets
        daily = pd.read_excel(uploaded_file, sheet_name="Daily Geological Report", header=None)
        litho_desc = pd.read_excel(uploaded_file, sheet_name="Lithological Description", header=None)
        litho_gas = pd.read_excel(uploaded_file, sheet_name="Lithology %, ROP & Gas Reading", header=None)

        # ==================== 1. EXTRACT WELL INFO ====================
        def get_value_after_keyword(df, keyword):
            for idx, row in df.iterrows():
                row_str = " ".join(row.astype(str))
                if keyword in row_str:
                    try:
                        parts = row_str.split(keyword)
                        return parts[1].split()[0].replace(",", "").strip()
                    except:
                        continue
            return "N/A"

        concession = get_value_after_keyword(daily, "Concession:-")
        well_name = get_value_after_keyword(daily, "Well:-")
        date_excel = get_value_after_keyword(daily, "Date:-")
        report_no = get_value_after_keyword(daily, "Report No.:-")
        rkb = get_value_after_keyword(daily, "RKB:-")
        spud_excel = get_value_after_keyword(daily, "Spud Date:-")
        geologist = "Youssef Osama / Mahmoud EL-Bana"  # fallback from sample

        # Parse dates
        try:
            report_date = excel_date(float(date_excel)).strftime("%Y-%m-%d")
            spud_date = excel_date(float(spud_excel)).strftime("%Y-%m-%d")
        except:
            report_date = date_excel
            spud_date = spud_excel

        # ==================== 2. DRILLING PROGRESS ====================
        depth_24 = daily[daily.apply(lambda row: "24:00 Hrs" in " ".join(row.astype(str)), axis=1)]
        depth_00 = daily[daily.apply(lambda row: "00:00 Hrs" in " ".join(row.astype(str)), axis=1)]
        depth_06 = daily[daily.apply(lambda row: "06:00 Hrs" in " ".join(row.astype(str)), axis=1)]
        prog_24 = daily[daily.apply(lambda row: "Progress 0-24 Hrs" in " ".join(row.astype(str)), axis=1)]
        prog_06 = daily[daily.apply(lambda row: "Progress Last 6 Hrs" in " ".join(row.astype(str)), axis=1)]

        def extract_number_near_text(df_row):
            try:
                return df_row.dropna().iloc[-2]
            except:
                return "N/A"

        depth_24_val = extract_number_near_text(depth_24)
        depth_00_val = extract_number_near_text(depth_00)
        depth_06_val = extract_number_near_text(depth_06)
        prog_24_val = extract_number_near_text(prog_24)
        prog_06_val = extract_number_near_text(prog_06)

        # ==================== 3. FORMATION TOPS TABLE ====================
        tops_start_row = daily[daily.apply(lambda row: "Formation Name" in " ".join(row.astype(str)), axis=1)].index[0]
        tops_df = daily.iloc[tops_start_row+3:].reset_index(drop=True)
        tops_df = tops_df.iloc[:, [2,3,4,5,6,7,8,9,10,11,12,13,14,15]]
        tops_df.columns = ["Formation", "Member", "Prog MD", "Prog TVD", "Prog SS", "Prog Thick",
                           "Actual MD", "Actual TVD SS", "Actual Thick", "Diff MD", "", "", "Ref MD", "Ref TVD SS", "Ref Thick"]

        # Clean and keep only relevant columns
        tops_clean = tops_df[["Formation", "Member", "Prog MD", "Prog TVD", "Actual MD", "Diff MD"]].copy()
        tops_clean = tops_clean.dropna(subset=["Formation"])
        tops_clean = tops_clean[tops_clean["Formation"].str.strip() != ""]
        tops_clean = tops_clean[tops_clean["Formation"] != "T.D."]
        tops_clean.replace(["", "nan"], np.nan, inplace=True)

        # ==================== 4. CURRENT FORMATION (last one before empty) ====================
        current_fm_row = daily[daily.apply(lambda row: row.astype(str).str.contains("Upper Bahariya|Lower Bahariya|KHARITA").any(), axis=1)]
        current_formation = "Upper Bahariya Fm."  # fallback
        if not current_fm_row.empty:
            fm = current_fm_row.iloc[0, 2]
            member = current_fm_row.iloc[0, 3] if pd.notna(current_fm_row.iloc[0, 3]) else ""
            current_formation = f"{fm} {member}".strip()

        # ==================== 5. GAS READINGS SUMMARY ====================
        gas_rows = litho_gas.iloc[10:, :]  # data starts around row 10
        gas_df = gas_rows[[0, 8,9,10,11,12,13,14]].dropna(subset=[0])
        gas_df.columns = ["Depth", "TG", "C1", "C2", "C3", "iC4", "nC4", "C5"]
        gas_df = gas_df.apply(pd.to_numeric, errors='ignore')
        gas_df = gas_df[gas_df["Depth"].apply(lambda x: isinstance(x, (int, float)) or x.isnumeric())]

        max_tg = gas_df["TG"].max()
        max_c1 = gas_df["C1"].max()
        avg_bg = gas_df["TG"].mean()

        # ==================== DISPLAY MERGED REPORT ====================
        st.success("Report parsed successfully!")

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🏢 Well Information")
            st.write(f"**Concession:** {concession}")
            st.write(f"**Well:** {well_name}")
            st.write(f"**Date:** {report_date}")
            st.write(f"**Report No.:** {report_no}")
            st.write(f"**RKB:** {rkb} ft")
            st.write(f"**Spud Date:** {spud_date}")
            st.write(f"**Wellsite Geologist:** {geologist}")

        with col2:
            st.subheader("⛏️ Drilling Progress (Last 24 hrs)")
            st.write(f"**Depth @ 24:00 hrs:** {depth_24_val} ft")
            st.write(f"**Depth @ 00:00 hrs:** {depth_00_val} ft")
            st.write(f"**Depth @ 06:00 hrs:** {depth_06_val} ft")
            st.write(f"**Progress (0-24 hrs):** {prog_24_val} ft")
            st.write(f"**Progress (Last 6 hrs):** {prog_06_val} ft")

        st.subheader(f"🪨 Current Formation Being Drilled")
        st.info(f"**{current_formation}**")

        st.subheader("📏 Formation Tops – Actual vs Prognosed")
        st.dataframe(tops_clean.style.format({
            "Prog MD": "{:.0f}", "Prog TVD": "{:.0f}",
            "Actual MD": "{:.0f}", "Diff MD": "{:+.0f}"
        }), use_container_width=True)

        st.subheader("🔥 Gas Readings Summary (Apollonia + Khoman)")
        col_g1, col_g2, col_g3 = st.columns(3)
        col_g1.metric("Max Total Gas (TG)", f"{max_tg:.0f} ppm")
        col_g2.metric("Max C1", f"{max_c1:.0f} ppm")
        col_g3.metric("Avg Background Gas", f"{avg_bg:.0f} ppm")

        if not gas_df.empty:
            st.line_chart(gas_df.set_index("Depth")[["TG", "C1"]], use_container_width=True)

        st.download_button(
            label="📥 Download Full Summary as Markdown",
            data=tops_clean.to_markdown() + "\n\nGas Data:\n" + gas_df.to_markdown(),
            file_name=f"{well_name}_DGR_{report_no}_Summary.md",
            mime="text/markdown"
        )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.exception(e)

else:
    st.info("Please upload your Daily Geological Report Excel file to begin.")
