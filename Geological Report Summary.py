import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

st.set_page_config(page_title="North Bahariya DGR Analyzer", layout="wide")
st.title("🏢 North Bahariya Daily Geological Report Analyzer")
st.markdown("Upload any **AB-xxx, DGR-xxx.xlsx** → get a clean merged summary in seconds")

uploaded_file = st.file_uploader("Upload DGR Excel file", type=["xlsx"])

def excel_date(n):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(float(n)))).strftime("%Y-%m-%d")
    except:
        return "N/A"

if uploaded_file:
    try:
        # === Load the 3 sheets exactly as they are named ===
        daily = pd.read_excel(uploaded_file, sheet_name="Daily Geological Report", header=None)
        litho_desc = pd.read_excel(uploaded_file, sheet_name="Lithological Description", header=None)
        gas_sheet = pd.read_excel(uploaded_file, sheet_name="Lithology %, ROP & Gas Reading", header=None)

        # === Helper: find value after keyword ===
        def val_after(keyword, col_offset=4):
            for _, row in daily.iterrows():
                txt = " ".join(row.astype(str))
                if keyword in txt:
                    try:
                        idx = txt.split(keyword)[1].split()[0]
                        return row[row == idx].index[0] + col_offset
                    except:
                        return row.iloc[row[row.str.contains(keyword, na=False)].index[0] + col_offset]
            return None

        # === Well Header ===
        concession = daily[daily.apply(lambda r: r.astype(str).str.contains("Concession").any(), axis=1)].iloc[0,4]
        well_name  = daily[daily.apply(lambda r: "Well:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        date_raw   = daily[daily.apply(lambda r: "Date:-" in " ".join(r.astype(str)), axis=1)].iloc[0,4]
        report_no  = daily[daily.apply(lambda r: "Report No.:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        rkb        = daily[daily.apply(lambda r: "RKB:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        spud_raw   = daily[daily.apply(lambda r: "Spud Date:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        geologist  = "Youssef Osama / Mahmoud EL-Bana"

        report_date = excel_date(date_raw)
        spud_date   = excel_date(spud_raw)

        # === Drilling Progress ===
        def get_depth(text):
            row = daily[daily.apply(lambda r: text in " ".join(r.astype(str)), axis=1)]
            if row.empty: return "N/A"
            return row.iloc[0].dropna().iloc[-2]

        depth_24h = get_depth("24:00 Hrs")
        depth_00h = get_depth("00:00 Hrs")
        depth_06h = get_depth("06:00 Hrs")
        prog_24h  = get_depth("Progress 0-24 Hrs")
        prog_6h   = get_depth("Progress Last 6 Hrs")

        # === Formation Tops Table ===
        header_row = daily[daily[0].astype(str).str.contains("Formation Name")].index[0]
        tops = daily.iloc[header_row+4:header_row+25, [2,3,4,6,7,9]].copy()
        tops.columns = ["Formation", "Member", "Prog_MD", "Prog_TVD", "Actual_MD", "Diff_MD"]
        tops = tops.dropna(subset=["Formation"])
        tops = tops[~tops["Formation"].astype(str).str.contains("T.D.")]
        tops = tops[tops["Formation"] != ""]

        # === Current Formation (last one with actual depth) ===
        drilled = tops.dropna(subset=["Actual_MD"])
        current_fm = drilled.iloc[-1]["Formation"] if not drilled.empty else "Unknown"
        current_member = drilled.iloc[-1]["Member"] if pd.notna(drilled.iloc[-1]["Member"]) else ""
        current_formation = f"{current_fm} {current_member}".strip()

        # === Gas Readings ===
        gas_start = gas_sheet[gas_sheetheet.apply(lambda r: "DEPTH" in " ".join(r.astype(str)).upper(), axis=1)].index[0] + 2
        gas = gas_sheet.iloc[gas_start:gas_start+200, [0,8,9]].dropna(subset=[0])
        gas.columns = ["Depth", "TG", "C1"]
        gas = gas[gas["Depth"].apply(lambda x: isinstance(x, (int,float)) or str(x).replace(".","").isdigit())]
        gas[["Depth","TG","C1"]] = gas[["Depth","TG","C1"]].apply(pd.to_numeric, errors='coerce')
        gas = gas.dropna()

        max_tg = gas["TG"].max()
        max_c1 = gas["C1"].max()
        avg_tg = round(gas["TG"].mean(), 0)

        # === DISPLAY ===
        st.success("✅ File parsed perfectly!")

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🏢 Well Information")
            st.write(f"**Concession:** {concession}")
            st.write(f"**Well:** {well_name}")
            st.write(f"**Report Date:** {report_date}")
            st.write(f"**Report No.:** {report_no}")
            st.write(f"**RKB:** {rkb} ft")
            st.write(f"**Spud Date:** {spud_date}")
            st.write(f"**Geologist:** {geologist}")

        with c2:
            st.subheader("⛏️ Drilling Progress (Last 24 hrs)")
            st.metric("Depth @ 24:00 hrs", f"{depth_24h} ft")
            st.metric("Depth @ 00:00 hrs", f"{depth_00h} ft")
            st.metric("Depth @ 06:00 hrs", f"{depth_06h} ft")
            st.metric("Progress 0–24 hrs", f"{prog_24h} ft")
            st.metric("Progress Last 6 hrs", f"{prog_6h} ft")

        st.subheader("🪨 Current Formation Being Drilled")
        st.info(f"**{current_formation}**")

        st.subheader("📏 Formation Tops – Actual vs Prognosed")
        st.dataframe(tops.style.format({"Prog_MD":"{:.0f}", "Prog_TVD":"{:.0f}", "Actual_MD":"{:.0f}", "Diff_MD":"{:+.0f}"}), use_container_width=True)

        st.subheader("🔥 Gas Readings Summary (Apollonia + Khoman)")
        g1, g2, g3 = st.columns(3)
        g1.metric("Max Total Gas (TG)", f"{max_tg:.0f} ppm")
        g2.metric("Max Methane (C1)", f"{max_c1:.0f} ppm")
        g3.metric("Avg Background Gas", f"{avg_tg} ppm")

        if not gas.empty:
            st.line_chart(gas.set_index("Depth")[["TG", "C1"]], use_container_width=True)

        # Download summary
        md = f"# {well_name} – DGR {report_no}\n\n**Date:** {report_date} | **Depth:** {depth_00h} ft\n\n" + tops.to_markdown(index=False)
        st.download_button("📥 Download Summary (Markdown)", md, f"{well_name}_DGR{report_no}_Summary.md", "text/markdown")

    except Exception as e:
        st.error("File loaded but parsing failed. Check sheet names are exactly:")
        st.code("Daily Geological Report\nLithological Description\nLithology %, ROP & Gas Reading")
        st.write("Error:", str(e))
else:
    st.info("Upload your North Bahariya DGR Excel file to begin")
