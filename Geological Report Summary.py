import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

st.set_page_config(page_title="North Bahariya DGR", layout="wide")
st.title("🏢 North Bahariya Daily Geological Report Analyzer")

uploaded_file = st.file_uploader("Upload AB-xxx, DGR-xxx.xlsx", type="xlsx")

def excel_to_date(n):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(float(n)))).strftime("%Y-%m-%d")
    except:
        return "N/A"

if uploaded_file:
    try:
        # Load the three sheets (exact names from your file)
        daily = pd.read_excel(uploaded_file, sheet_name="Daily Geological Report", header=None)
        litho = pd.read_excel(uploaded_file, sheet_name="Lithological Description", header=None)
        gas   = pd.read_excel(uploaded_file, sheet_name="Lithology %, ROP & Gas Reading", header=None)

        # ------------------- Header -------------------
        well_row      = daily[daily.apply(lambda r: "Well:-" in " ".join(r.astype(str)), axis=1)].iloc[0]
        concession    = "North Bahariya"
        well_name     = well_row.iloc[12].strip()
        date_raw      = daily[daily.apply(lambda r: "Date:-" in " ".join(r.astype(str)), axis=1)].iloc[0,4]
        report_no     = daily[daily.apply(lambda r: "Report No.:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        rkb           = daily[daily.apply(lambda r: "RKB:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]
        spud_raw      = daily[daily.apply(lambda r: "Spud Date:-" in " ".join(r.astype(str)), axis=1)].iloc[0,12]

        report_date = excel_to_date(date_raw)
        spud_date   = excel_to_date(spud_raw)

        # ------------------- Progress -------------------
        def get_val(keyword):
            row = daily[daily.apply(lambda r: keyword in " ".join(r.astype(str)), axis=1)]
            if row.empty: return "N/A"
            return row.iloc[0].dropna().values[-2]

        d24 = get_val("24:00 Hrs")
        d00 = get_val("00:00 Hrs")
        d06 = get_val("06:00 Hrs")
        p24 = get_val("Progress 0-24 Hrs")
        p06 = get_val("Progress Last 6 Hrs")

        # ------------------- Formation Tops -------------------
        start = daily[daily.apply(lambda r: "Formation Name" in " ".join(r.astype(str)), axis=1)].index[0] + 4
        tops = daily.iloc[start:start+25, [2,3,4,5,6,7,8,9]].dropna(subset=[2])
        tops.columns = ["Fm", "Member", "Prog MD", "Prog TVD", "Prog SS", "Prog Thk", "Act MD", "Act TVDSS"]
        tops = tops[["Fm", "Member", "Prog MD", "Prog TVD", "Act MD"]]
        tops = tops[tops["Fm"].notna() & (tops["Fm"] != "T.D.")]

        current = tops.dropna(subset=["Act MD"]).iloc[-1]
        current_fm = f"{current['Fm']} {current['Member']}".strip()

        # ------------------- Gas -------------------
        gas_data = gas.iloc[10:, [0,8,9]].dropna(subset=[0])
        gas_data = gas_data[gas_data[0].apply(lambda x: str(x).replace(".", "").isdigit())]
        gas_data.columns = ["Depth", "TG", "C1"]
        gas_data[["Depth","TG","C1"]] = gas_data[["Depth","TG","C1"]].apply(pd.to_numeric)
        max_tg = gas_data["TG"].max()
        max_c1 = gas_data["C1"].max()
        avg_tg = round(gas_data["TG"].mean())

        # ------------------- Display -------------------
        st.success("File parsed perfectly!")

        a, b = st.columns(2)
        with a:
            st.subheader("Well Information")
            st.write(f"**Well:** {well_name}")
            st.write(f"**Concession:** {concession}")
            st.write(f"**Report Date:** {report_date} | Report No. {report_no}")
            st.write(f"**Spud:** {spud_date} | RKB: {rkb} ft")

        with b:
            st.subheader("Drilling Progress (Last 24 h)")
            st.metric("Depth @ 24:00 h", f"{d24} ft")
            st.metric("Depth @ 00:00 h (current)", f"{d00} ft")
            st.metric("Progress 0-24 h", f"{p24} ft")
            st.metric("Progress last 6 h", f"{p06} ft")

        st.info(f"**Current Formation:** {current_fm}")

        st.subheader("Formation Tops")
        st.dataframe(tops.style.format({"Prog MD": "{:.0f}", "Prog TVD": "{:.0f}", "Act MD": "{:.0f}"}), use_container_width=True)

        st.subheader("Gas Readings (Apollonia + Khoman)")
        c1, c2, c3 = st.columns(3)
        c1.metric("Max TG", f"{max_tg:.0f} ppm")
        c2.metric("Max C1", f"{max_c1:.0f} ppm")
        c3.metric("Avg Background", f"{avg_tg} ppm")
        st.line_chart(gas_data.set_index("Depth")[["TG", "C1"]])

        st.download_button("Download Summary (Markdown)", 
                           data=tops.to_markdown(index=False),
                           file_name=f"{well_name}_DGR_{report_no}_Summary.md")

    except ImportError as e:
        if "openpyxl" in str(e):
            st.error("openpyxl is missing → Add 'openpyxl' to your requirements.txt")
        else:
            st.error("Parsing error")
            st.write(e)
    except Exception as e:
        st.error("Wrong file or sheet names")
        st.code("Required sheet names (exact):\nDaily Geological Report\nLithological Description\nLithology %, ROP & Gas Reading")
        st.write(e)
else:
    st.info("Upload your North Bahariya DGR file → works instantly")
