import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

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
        # Load the three sheets using pandas default engine
        daily = pd.read_excel(uploaded_file, sheet_name="Daily Geological Report", header=None)
        litho = pd.read_excel(uploaded_file, sheet_name="Lithological Description", header=None)
        gas   = pd.read_excel(uploaded_file, sheet_name="Lithology %, ROP & Gas Reading", header=None)

        # ------------------- Header -------------------
        # Find well row safely
        well_mask = daily.apply(lambda r: any("Well:-" in str(cell) for cell in r), axis=1)
        if well_mask.any():
            well_row = daily[well_mask].iloc[0]
            well_name = str(well_row.iloc[12]).strip() if len(well_row) > 12 else "N/A"
        else:
            well_name = "N/A"
            
        concession = "North Bahariya"
        
        # Find date row safely
        date_mask = daily.apply(lambda r: any("Date:-" in str(cell) for cell in r), axis=1)
        date_raw = daily[date_mask].iloc[0,4] if date_mask.any() else "N/A"
        
        # Find report number safely
        report_mask = daily.apply(lambda r: any("Report No.:-" in str(cell) for cell in r), axis=1)
        report_no = daily[report_mask].iloc[0,12] if report_mask.any() else "N/A"
        
        # Find RKB safely
        rkb_mask = daily.apply(lambda r: any("RKB:-" in str(cell) for cell in r), axis=1)
        rkb = daily[rkb_mask].iloc[0,12] if rkb_mask.any() else "N/A"
        
        # Find spud date safely
        spud_mask = daily.apply(lambda r: any("Spud Date:-" in str(cell) for cell in r), axis=1)
        spud_raw = daily[spud_mask].iloc[0,12] if spud_mask.any() else "N/A"

        report_date = excel_to_date(date_raw)
        spud_date = excel_to_date(spud_raw)

        # ------------------- Progress -------------------
        def get_val(keyword):
            mask = daily.apply(lambda r: any(keyword in str(cell) for cell in r), axis=1)
            if not mask.any(): 
                return "N/A"
            row = daily[mask].iloc[0]
            non_nan_values = row.dropna().values
            return non_nan_values[-2] if len(non_nan_values) >= 2 else "N/A"

        d24 = get_val("24:00 Hrs")
        d00 = get_val("00:00 Hrs")
        d06 = get_val("06:00 Hrs")
        p24 = get_val("Progress 0-24 Hrs")
        p06 = get_val("Progress Last 6 Hrs")

        # ------------------- Formation Tops -------------------
        tops_mask = daily.apply(lambda r: any("Formation Name" in str(cell) for cell in r), axis=1)
        if tops_mask.any():
            start = daily[tops_mask].index[0] + 4
            tops = daily.iloc[start:start+25, [2,3,4,5,6,7,8,9]].dropna(subset=[2])
            tops.columns = ["Fm", "Member", "Prog MD", "Prog TVD", "Prog SS", "Prog Thk", "Act MD", "Act TVDSS"]
            tops = tops[["Fm", "Member", "Prog MD", "Prog TVD", "Act MD"]]
            tops = tops[tops["Fm"].notna() & (tops["Fm"] != "T.D.")]
            
            if not tops.empty and "Act MD" in tops.columns:
                current = tops.dropna(subset=["Act MD"]).iloc[-1]
                current_fm = f"{current['Fm']} {current['Member']}".strip()
            else:
                current_fm = "N/A"
        else:
            tops = pd.DataFrame()
            current_fm = "N/A"

        # ------------------- Gas -------------------
        if len(gas) > 10:
            gas_data = gas.iloc[10:, [0,8,9]].dropna(subset=[0])
            if not gas_data.empty:
                # Filter numeric depth values
                gas_data = gas_data[gas_data[0].apply(lambda x: str(x).replace(".", "").replace("-", "").isdigit())]
                if not gas_data.empty:
                    gas_data.columns = ["Depth", "TG", "C1"]
                    try:
                        gas_data[["Depth","TG","C1"]] = gas_data[["Depth","TG","C1"]].apply(pd.to_numeric, errors='coerce')
                        gas_data = gas_data.dropna()
                        max_tg = gas_data["TG"].max() if not gas_data.empty else 0
                        max_c1 = gas_data["C1"].max() if not gas_data.empty else 0
                        avg_tg = round(gas_data["TG"].mean()) if not gas_data.empty else 0
                    except:
                        max_tg, max_c1, avg_tg = 0, 0, 0
                else:
                    max_tg, max_c1, avg_tg = 0, 0, 0
            else:
                max_tg, max_c1, avg_tg = 0, 0, 0
        else:
            max_tg, max_c1, avg_tg = 0, 0, 0
            gas_data = pd.DataFrame()

        # ------------------- Display -------------------
        st.success("File parsed successfully!")

        a, b = st.columns(2)
        with a:
            st.subheader("Well Information")
            st.write(f"**Well:** {well_name}")
            st.write(f"**Concession:** {concession}")
            st.write(f"**Report Date:** {report_date} | Report No. {report_no}")
            st.write(f"**Spud:** {spud_date} | RKB: {rkb}")

        with b:
            st.subheader("Drilling Progress (Last 24 h)")
            st.metric("Depth @ 24:00 h", f"{d24}")
            st.metric("Depth @ 00:00 h (current)", f"{d00}")
            st.metric("Progress 0-24 h", f"{p24}")
            st.metric("Progress last 6 h", f"{p06}")

        st.info(f"**Current Formation:** {current_fm}")

        st.subheader("Formation Tops")
        if not tops.empty:
            # Format numeric columns
            numeric_cols = ["Prog MD", "Prog TVD", "Act MD"]
            for col in numeric_cols:
                if col in tops.columns:
                    tops[col] = pd.to_numeric(tops[col], errors='coerce')
            
            st.dataframe(tops.style.format({
                col: "{:.0f}" for col in numeric_cols if col in tops.columns
            }), use_container_width=True)
        else:
            st.write("No formation tops data available")

        st.subheader("Gas Readings (Apollonia + Khoman)")
        c1, c2, c3 = st.columns(3)
        c1.metric("Max TG", f"{max_tg:.0f} ppm")
        c2.metric("Max C1", f"{max_c1:.0f} ppm")
        c3.metric("Avg Background", f"{avg_tg} ppm")
        
        if not gas_data.empty and "Depth" in gas_data.columns:
            st.line_chart(gas_data.set_index("Depth")[["TG", "C1"]])
        else:
            st.write("No gas data available for chart")

        # Download button
        if not tops.empty:
            st.download_button(
                "Download Summary (Markdown)", 
                data=tops.to_markdown(index=False),
                file_name=f"{well_name}_DGR_{report_no}_Summary.md"
            )

    except Exception as e:
        st.error("Error processing the file")
        st.write(f"Error details: {str(e)}")
        st.code("Required sheet names (exact):\n- Daily Geological Report\n- Lithological Description\n- Lithology %, ROP & Gas Reading")
else:
    st.info("Upload your North Bahariya DGR file to begin analysis")
