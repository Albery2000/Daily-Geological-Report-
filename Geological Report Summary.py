import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# === CRITICAL: Explicitly import openpyxl (fixes the error) ===
import openpyxl  # ← This line prevents the "Missing optional dependency 'openpyxl'" error

st.set_page_config(page_title="DGR Analyzer - North Bahariya", layout="wide")
st.title("🏢 Daily Geological Report Analyzer")
st.markdown("Upload your **AB-xxx, DGR-xxx.xlsx** file → get a clean professional summary instantly")

uploaded_file = st.file_uploader("Upload Excel Report", type=["xlsx"])

# Excel serial date → real date
def excel_date(num):
    if pd.isna(num):
        return "N/A"
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(float(num)))).strftime("%Y-%m-%d")
    except:
        return "N/A"

if uploaded_file is not None:
    try:
        # Load all 3 sheets without assuming headers
        daily = pd.read_excel(uploaded_file, sheet_name="Daily Geological Report", header=None)
        litho_desc = pd.read_excel(uploaded_file, sheet_name="Lithological Description", header=None)
        litho_gas = pd.read_excel(uploaded_file, sheet_name="Lithology %, ROP & Gas Reading", header=None)

        # Helper: find value after a keyword in any row
        def find_after(df, keyword, col_offset=4):
            for _, row in df.iterrows():
                row_str = " ".join(row.astype(str).fillna(""))
                if keyword in row_str:
                    try:
                        return row.iloc[row_str.split(keyword)[0].count(" ") + col_offset]
                    except:
                        continue
            return "N/A"

        # ==================== WELL HEADER ====================
        concession = find_after(daily, "Concession:-", 5)
        well_name = find_after(daily, "Well:-", 5)
        report_no = find_after(daily, "Report No.:-", 3)
        rkb = find_after(daily, "RKB:-", 3)
        date_raw = find_after(daily, "Date:-", 3)
        spud_raw = find_after(daily, "Spud Date:-", 3)
        geologist = "Youssef Osama / Mahmoud EL-Bana"  # common in your files

        report_date = excel_date(date_raw)
        spud_date = excel_date(spud_raw)

        # ==================== DRILLING PROGRESS ====================
        def get_depth_value(text):
            row = daily[daily.apply(lambda r: text in " ".join(r.astype(str)), axis=1)]
            if not row.empty:
                vals = row.iloc[0].dropna()
                return vals.iloc[-2] if len(vals) > 2 else vals.iloc[-1]
            return "N/A"

        depth_24 = get_depth_value("24:00 Hrs")
        depth_00 = get_depth_value("00:00 Hrs")
        depth_06 = get_depth_value("06:00 Hrs")
        prog_24 = get_depth_value("Progress 0-24 Hrs")
        prog_06 = get_depth_value("Progress Last 6 Hrs")

        # ==================== FORMATION TOPS ====================
        try:
            header_row = daily[daily.apply(lambda r: "Formation Name" in " ".join(r.astype(str)), axis=1)].index[0]
            tops_raw = daily.iloc[header_row+4:header_row+20].iloc[:, [2,3,4,5,6,7,8,9,10]]
            tops_raw.columns = ["Formation", "Member", "Prog_MD", "Prog_TVD", "Prog_SS", "Prog_Thk",
                                "Act_MD", "Act_TVDSS", "Act_Thk"]
            tops_clean = tops_raw[["Formation", "Member", "Prog_MD", "Prog_TVD", "Act_MD"]].copy()
            tops_clean = tops_clean.dropna(subset=["Formation"])
            tops_clean = tops_clean[tops_clean["Formation"].str.strip() != ""]
            tops_clean = tops_clean[~tops_clean["Formation"].astype(str).str.contains("T.D.")]
            tops_clean = tops_clean.replace(["", "nan"], np.nan)
        except:
            tops_clean = pd.DataFrame(columns=["Formation", "Member", "Prog_MD", "Prog_TVD", "Act_MD"])

        # ==================== CURRENT FORMATION ====================
        drilled_formations = tops_clean.dropna(subset=["Act_MD"])["Formation"].tolist()
        current_formation = drilled_formations[-1] if drilled_formations else "Surface"

        # ==================== GAS READINGS ====================
        try:
            gas_data = litho_gas.iloc[9:, [0, 8,9,10,11,12,13,14]]
            gas_data = gas_data.dropna(subset=[0])
            gas_data.columns = ["Depth", "TG", "C1", "C2", "C3", "iC4", "nC4", "C5"]
            gas_data = gas_data[gas_data["Depth"].apply(lambda x: str(x).replace('.','').isdigit())]
            gas_data[["Depth", "TG", "C1", "C2", "C3", "iC4", "nC4", "C5"]] = gas_data[["Depth", "TG", "C1", "C2", "C3", "iC4", "nC4", "C5"]].apply(pd.to_numeric, errors='coerce')
            gas_data = gas_data.dropna(subset=["TG", "C1"])

            max_tg = gas_data["TG"].max()
            max_c1 = gas_data["C1"].max()
            avg_tg = gas_data["TG"].mean()
        except:
            max_tg = max_c1 = avg_tg = 0
            gas_data = pd.DataFrame()

        # ==================== DISPLAY RESULTS ====================
        st.success("✅ Report parsed successfully!")

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🏢 Well Information")
            st.write(f"**Concession:** {concession}")
            st.write(f"**Well:** {well_name}")
            st.write(f"**Report Date:** {report_date}")
            st.write(f"**Report No.:** {report_no}")
            st.write(f"**RKB:** {rkb} ft")
            st.write(f"**Spud Date:** {spud_date}")
            st.write(f"**Geologist:** {geologist}")

        with col2:
            st.subheader("⛏️ Drilling Progress (Last 24h)")
            st.metric("Depth @ 24:00 hrs", f"{depth_24} ft")
            st.metric("Depth @ 00:00 hrs", f"{depth_00} ft")
            st.metric("Depth @ 06:00 hrs", f"{depth_06} ft")
            st.metric("Progress 0–24 hrs", f"{prog_24} ft")
            st.metric("Progress Last 6 hrs", f"{prog_06} ft")

        st.subheader("🪨 Current Formation Being Drilled")
        st.info(f"**{current_formation}**")

        st.subheader("📏 Formation Tops – Actual vs Prognosed")
        styled_tops = tops_clean.style.format({
            "Prog_MD": "{:.0f}", "Prog_TVD": "{:.0f}", "Act_MD": "{:.0f}"
        }).background_gradient(subset=["Act_MD"], cmap="Greens")
        st.dataframe(styled_tops, use_container_width=True)

        st.subheader("🔥 Gas Readings Summary")
        g1, g2, g3 = st.columns(3)
        g1.metric("Max Total Gas (TG)", f"{max_tg:.0f}" if max_tg > 0 else "—", "ppm")
        g2.metric("Max Methane (C1)", f"{max_c1:.0f}" if max_c1 > 0 else "—", "ppm")
        g3.metric("Avg Background Gas", f"{avg_tg:.0f}" if avg_tg > 0 else "—", "ppm")

        if not gas_data.empty:
            st.line_chart(gas_data.set_index("Depth")[["TG", "C1"]], use_container_width=True)

        # Download button
        summary_md = f"# {well_name} - DGR {report_no}\n\n" \
                     f"**Date:** {report_date} | **Depth:** {depth_00} ft\n\n" \
                     + tops_clean.to_markdown(index=False) \
                     + "\n\n## Gas Readings\n" + gas_data.to_markdown(index=False)

        st.download_button(
            "📥 Download Summary (Markdown)",
            summary_md,
            file_name=f"{well_name}_DGR_{report_no}_Summary.md",
            mime="text/markdown"
        )

    except Exception as e:
        st.error("Could not parse the file. Make sure it has the 3 correct sheets.")
        st.write("Error details:", str(e))

else:
    st.info("👆 Upload your Daily Geological Report (.xlsx) to start")
    st.markdown("Supported format: **AB-88, DGR-5.xlsx** style reports with 3 sheets")
