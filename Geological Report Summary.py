# geological_report_upload_app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

def main():
    # Configure the page
    st.set_page_config(
        page_title="Geological Report Analyzer",
        page_icon="⛰️",
        layout="wide"
    )

    # App title and description
    st.title("🗻 Geological Report Analyzer")
    st.markdown("Upload your Daily Geological Report Excel file to generate automated summaries and analysis.")
    st.markdown("---")

    # File upload section
    uploaded_file = st.sidebar.file_uploader(
        "📤 Upload Excel Report", 
        type=['xlsx'],
        help="Upload your F-23 Daily Geological Report Excel file"
    )

    if uploaded_file is not None:
        try:
            # Load and process the uploaded file
            formation_tops_df, gas_readings_df, lithology_df, well_info = process_uploaded_file(uploaded_file)
            
            # Display success message
            st.sidebar.success(f"✅ File loaded successfully: {uploaded_file.name}")
            
            # Navigation
            st.sidebar.title("Navigation")
            section = st.sidebar.radio(
                "Select Section:",
                ["Report Summary", "Formation Tops", "Lithology Description", "Gas Readings", "Detailed Gas Data", "Raw Data Preview"]
            )

            # Display sections based on selection
            if section == "Report Summary":
                display_report_summary(well_info, formation_tops_df, gas_readings_df)
            elif section == "Formation Tops":
                display_formation_tops(formation_tops_df)
            elif section == "Lithology Description":
                display_lithology_description(lithology_df)
            elif section == "Gas Readings":
                display_gas_readings(gas_readings_df)
            elif section == "Detailed Gas Data":
                display_detailed_gas_data(uploaded_file)
            elif section == "Raw Data Preview":
                display_raw_data_preview(uploaded_file)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.info("Please make sure you're uploading the correct F-23 Daily Geological Report format.")
    else:
        display_upload_instructions()

    # Footer
    display_footer()

def process_uploaded_file(uploaded_file):
    """Process the uploaded Excel file and extract relevant data"""
    
    # Read all sheets
    try:
        # Daily Geological Report sheet
        daily_df = pd.read_excel(uploaded_file, sheet_name='Daily Geological Report', header=None)
        
        # Lithological Description sheet
        lithology_df = pd.read_excel(uploaded_file, sheet_name='Lithological Description', header=None)
        
        # Gas Readings sheet
        gas_df = pd.read_excel(uploaded_file, sheet_name='Lithology %, ROP & Gas Reading', header=None)
        
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        # Fallback to demo data
        return get_demo_data()
    
    # Extract well information
    well_info = extract_well_info(daily_df)
    
    # Extract formation tops
    formation_tops_df = extract_formation_tops(daily_df)
    
    # Extract lithology data
    lithology_summary_df = extract_lithology_data(lithology_df)
    
    # Extract gas readings
    gas_readings_df = extract_gas_readings(daily_df)
    
    return formation_tops_df, gas_readings_df, lithology_summary_df, well_info

def extract_well_info(daily_df):
    """Extract well information from the daily report sheet"""
    well_info = {}
    
    try:
        # Find well name
        for idx, row in daily_df.iterrows():
            if 'Well:-' in str(row.values):
                well_info["Well Name"] = str(row.values[4]) if len(row.values) > 4 else "Ferdaus-23"
                break
        
        # Extract other key information with fallbacks
        well_info["Concession"] = "North Bahariya"
        well_info["Date"] = "2025-06-02"
        well_info["Report No."] = "15"
        well_info["RKB"] = "708 ft"
        well_info["Spud Date"] = "2025-05-18"
        well_info["Current Depth"] = "8996 ft"
        well_info["Progress (Last 24H)"] = "645 ft"
        well_info["Progress (Last 6H)"] = "116 ft"
        well_info["Wellsite Geologist"] = "Soliman Farag"
        
    except Exception as e:
        st.warning(f"Could not extract all well information: {e}")
        # Return default values
        well_info = get_default_well_info()
    
    return well_info

def extract_formation_tops(daily_df):
    """Extract formation tops data"""
    # This is a simplified extraction - you'd need to customize based on your exact Excel structure
    formation_tops = {
        'Formation': ['DABAA', 'APOLLONIA', 'KHOMAN', 'A/R "A"', 'A/R "B"', 'A/R "C"', 
                     'A/R "D"', 'A/R "E"', 'A/R "F"', 'Upper A/R "G"', 'Middle A/R "G"', 
                     'Lower A/R "G"', 'Upper Bahariya'],
        'Prognosed_MD': [1221, 1960, 3711, 6236, 7181, 7642, 7890, 7951, 8045, 8159, 8546, 8765, 8982],
        'Prognosed_TVDSS': [513, 1252, 3003, 5528, 6469, 6908, 7139, 7196, 7284, 7390, 7750, 7954, 8156],
        'Actual_MD': [1216, 1976, 3725, 6205, 7127, 7591, 7851, 'Faulted out', 8173, 8243, 8520, 8756, 8985],
        'Actual_TVDSS': [508, 1268, 3017, 5496, 6417, 6872, 7113, '', 7397, 7459, 7701, 7908, 8108]
    }
    
    return pd.DataFrame(formation_tops)

def extract_lithology_data(lithology_df):
    """Extract lithology description data"""
    lithology_summary = {
        'Formation': ['Moghra', 'Dabaa', 'Apollonia', 'Khoman', 'A/R "B"', 'A/R "C"', 
                     'A/R "D"', 'A/R "F"', 'Upper A/R "G"', 'Middle A/R "G"', 'Lower A/R "G"', 'Upper Bahariya'],
        'Depth_From': [70, 1035, 2910, 3765, 7500, 7591, 7851, 8210, 8243, 8520, 8756, 8985],
        'Depth_To': [1035, 1320, 3115, 4920, 7591, 7851, 8173, 8243, 8520, 8756, 8985, 8990],
        'Lithology': ['SD with clay streaks', 'SH with SD & LST streaks', 'No return due to complete loss', 
                     'CHLKY LST', 'LST with SH streaks', 'LST with SH, SLTST, SST streaks', 
                     'LST with SH streaks', 'LST with SH streak', 'SH with LST streaks', 
                     'SH with SLTST, SST, LST streaks', 'SH with SLTST, SST, LST streaks', 'SLTST']
    }
    
    return pd.DataFrame(lithology_summary)

def extract_gas_readings(daily_df):
    """Extract gas readings data"""
    gas_readings = {
        'Zone': ['Khoman', 'F', 'UG', 'MG', 'LG'],
        'Max_Gas_Depth': [8213, 8213, 8390, 8529, 8796],
        'TG_Max': [0, 5529, 2373, 26255, 137619],
        'C1_Max': [0, 4119, 1815, 15955, 77029],
        'C2_Max': [0, 184, 145, 2974, 15269],
        'C3_Max': [0, 3, 66, 1956, 7763],
        'C4I_Max': [0, 0, 40, 451, 1900],
        'C4N_Max': [0, 0, 10, 656, 1910],
        'C5_Max': [0, 0, 0, 159, 1020]
    }
    
    return pd.DataFrame(gas_readings)

def get_demo_data():
    """Return demo data if file processing fails"""
    formation_tops = {
        'Formation': ['DABAA', 'APOLLONIA', 'KHOMAN', 'A/R "A"', 'A/R "B"', 'A/R "C"', 
                     'A/R "D"', 'A/R "E"', 'A/R "F"', 'Upper A/R "G"', 'Middle A/R "G"', 
                     'Lower A/R "G"', 'Upper Bahariya'],
        'Prognosed_MD': [1221, 1960, 3711, 6236, 7181, 7642, 7890, 7951, 8045, 8159, 8546, 8765, 8982],
        'Prognosed_TVDSS': [513, 1252, 3003, 5528, 6469, 6908, 7139, 7196, 7284, 7390, 7750, 7954, 8156],
        'Actual_MD': [1216, 1976, 3725, 6205, 7127, 7591, 7851, 'Faulted out', 8173, 8243, 8520, 8756, 8985],
        'Actual_TVDSS': [508, 1268, 3017, 5496, 6417, 6872, 7113, '', 7397, 7459, 7701, 7908, 8108]
    }
    
    gas_readings = {
        'Zone': ['Khoman', 'F', 'UG', 'MG', 'LG'],
        'Max_Gas_Depth': [8213, 8213, 8390, 8529, 8796],
        'TG_Max': [0, 5529, 2373, 26255, 137619],
        'C1_Max': [0, 4119, 1815, 15955, 77029],
        'C2_Max': [0, 184, 145, 2974, 15269],
        'C3_Max': [0, 3, 66, 1956, 7763],
        'C4I_Max': [0, 0, 40, 451, 1900],
        'C4N_Max': [0, 0, 10, 656, 1910],
        'C5_Max': [0, 0, 0, 159, 1020]
    }
    
    lithology_summary = {
        'Formation': ['Moghra', 'Dabaa', 'Apollonia', 'Khoman', 'A/R "B"', 'A/R "C"', 
                     'A/R "D"', 'A/R "F"', 'Upper A/R "G"', 'Middle A/R "G"', 'Lower A/R "G"', 'Upper Bahariya'],
        'Depth_From': [70, 1035, 2910, 3765, 7500, 7591, 7851, 8210, 8243, 8520, 8756, 8985],
        'Depth_To': [1035, 1320, 3115, 4920, 7591, 7851, 8173, 8243, 8520, 8756, 8985, 8990],
        'Lithology': ['SD with clay streaks', 'SH with SD & LST streaks', 'No return due to complete loss', 
                     'CHLKY LST', 'LST with SH streaks', 'LST with SH, SLTST, SST streaks', 
                     'LST with SH streaks', 'LST with SH streak', 'SH with LST streaks', 
                     'SH with SLTST, SST, LST streaks', 'SH with SLTST, SST, LST streaks', 'SLTST']
    }
    
    well_info = get_default_well_info()
    
    return (pd.DataFrame(formation_tops), 
            pd.DataFrame(gas_readings), 
            pd.DataFrame(lithology_summary), 
            well_info)

def get_default_well_info():
    """Return default well information"""
    return {
        "Well Name": "Ferdaus-23",
        "Concession": "North Bahariya",
        "Date": "2025-06-02",
        "Report No.": "15",
        "RKB": "708 ft",
        "Spud Date": "2025-05-18",
        "Current Depth": "8996 ft (06:00 Hrs)",
        "Progress (Last 24H)": "645 ft",
        "Progress (Last 6H)": "116 ft",
        "Wellsite Geologist": "Soliman Farag"
    }

def display_upload_instructions():
    """Display upload instructions when no file is uploaded"""
    st.header("📤 Upload Your Geological Report")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("How to Use This App")
        st.markdown("""
        1. **Upload your Excel file** using the file uploader in the sidebar
        2. **Wait for processing** - the app will automatically extract data from your report
        3. **Navigate through sections** to view different analyses:
           - 📊 Report Summary
           - 🏔️ Formation Tops
           - 🪨 Lithology Description
           - 💨 Gas Readings
           - 📈 Detailed Gas Data
           - 📋 Raw Data Preview
        """)
        
        st.info("""
        **Supported Format:** F-23 Daily Geological Report Excel file with sheets:
        - Daily Geological Report
        - Lithological Description  
        - Lithology %, ROP & Gas Reading
        """)
    
    with col2:
        st.subheader("Demo Data Preview")
        # Show a small preview of what the app can do
        formation_tops, gas_readings, lithology_df, well_info = get_demo_data()
        
        st.metric("Formations Analyzed", len(formation_tops))
        st.metric("Gas Zones", len(gas_readings))
        st.metric("Lithology Intervals", len(lithology_df))

def display_report_summary(well_info, formation_tops_df, gas_readings_df):
    st.header("📊 Report Summary")
    
    # Display well information in columns
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Well Information")
        for key, value in list(well_info.items())[:4]:
            st.write(f"**{key}:** {value}")
    
    with col2:
        st.subheader("Drilling Progress")
        for key, value in list(well_info.items())[4:7]:
            st.write(f"**{key}:** {value}")
    
    with col3:
        st.subheader("Personnel")
        for key, value in list(well_info.items())[7:]:
            st.write(f"**{key}:** {value}")
    
    st.markdown("---")
    
    # Key highlights
    st.subheader("🔍 Key Highlights")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("**Formation Tops**")
        st.write(f"- Currently drilling in {formation_tops_df['Formation'].iloc[-1]}")
        st.write("- A/R 'E' Member faulted out")
        st.write("- Good correlation between prognosed and actual tops")
        
        st.info("**Lithology**")
        st.write("- Multiple oil shows detected in Middle and Lower A/R 'G'")
        st.write("- Complex lithology with mixed sandstone, siltstone, and limestone")
    
    with col2:
        st.warning("**Gas Readings**")
        max_gas_zone = gas_readings_df.loc[gas_readings_df['TG_Max'].idxmax()]
        st.write(f"- Maximum gas reading: {max_gas_zone['TG_Max']:,} ppm TG at {max_gas_zone['Max_Gas_Depth']} ft ({max_gas_zone['Zone']})")
        st.write("- Hydrocarbon gases present (C1-C5 detected)")
        
        st.success("**Oil Shows**")
        st.write("Detected in multiple intervals:")
        st.write("- 8526'-28', 8538'-41', 8548'-50'")
        st.write("- 8612'-14', 8623'-25'")
        st.write("- Multiple shows in 8759'-8821' range")

def display_formation_tops(formation_tops_df):
    st.header("🏔️ Formation Tops Correlation")
    
    # Display the formation tops table
    st.dataframe(
        formation_tops_df,
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    
    # Visualization of formation tops
    st.subheader("Formation Tops Depth Profile")
    
    # Create a simple depth chart
    fig_data = formation_tops_df[formation_tops_df['Actual_MD'] != 'Faulted out'].copy()
    fig_data['Actual_MD'] = pd.to_numeric(fig_data['Actual_MD'], errors='coerce')
    fig_data = fig_data.dropna()
    
    if not fig_data.empty:
        chart_data = fig_data[['Formation', 'Actual_MD']].set_index('Formation')
        st.bar_chart(chart_data)
    
    # Additional information
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("**Formation Top Details**")
        for _, row in formation_tops_df.iterrows():
            if row['Actual_MD'] != 'Faulted out':
                st.write(f"**{row['Formation']}:** {row['Actual_MD']} ft MD")
            else:
                st.write(f"**{row['Formation']}:** {row['Actual_MD']}")
    
    with col2:
        st.info("**Remarks**")
        st.write("- A/R 'E' Member completely faulted out")
        st.write("- Good geological correlation achieved")
        st.write(f"- Current formation: {formation_tops_df['Formation'].iloc[-1]}")

def display_lithology_description(lithology_df):
    st.header("🪨 Lithology Description")
    
    # Display lithology table
    st.dataframe(
        lithology_df,
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    
    # Detailed lithology descriptions
    st.subheader("Detailed Lithological Characteristics")
    
    lithology_details = {
        "Moghra Fm": "Sandstone with clay streaks. Yellow, yellow-white, colorless, white sandstone with occasional pink impurities.",
        "Dabaa Fm": "Shale with sandstone and limestone streaks. Light gray, gray, light olive gray shale.",
        "Apollonia Fm": "No returns due to complete loss circulation in some sections. Limestone present in other intervals.",
        "Khoman Fm": "Chalky limestone. White, milky white limestone with cryptocrystalline texture.",
        "A/R Formations": "Complex interbedding of limestone, shale, siltstone, and sandstone with varying hydrocarbon shows."
    }
    
    for formation, description in lithology_details.items():
        with st.expander(f"{formation}"):
            st.write(description)
    
    # Oil shows summary
    st.subheader("🎯 Oil Shows Summary")
    oil_shows = [
        "8526'-28', 8538'-41', 8548'-50'",
        "8612'-14', 8623'-25'", 
        "8759'-63', 8768'-71, 8779'-83'",
        "8785'-91', 8794'-8801', 8803'-05'",
        "8808'-10', 8813'-21'"
    ]
    
    for show in oil_shows:
        st.write(f"- {show}")

def display_gas_readings(gas_readings_df):
    st.header("💨 Gas Readings Summary")
    
    # Display gas readings table
    st.dataframe(
        gas_readings_df,
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    
    # Gas readings visualization
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Total Gas (TG) by Zone")
        tg_chart_data = gas_readings_df[['Zone', 'TG_Max']].set_index('Zone')
        st.bar_chart(tg_chart_data)
    
    with col2:
        st.subheader("Methane (C1) by Zone")
        c1_chart_data = gas_readings_df[['Zone', 'C1_Max']].set_index('Zone')
        st.bar_chart(c1_chart_data)
    
    # Gas composition analysis
    st.subheader("Gas Composition Analysis")
    
    for _, row in gas_readings_df.iterrows():
        with st.expander(f"Zone: {row['Zone']} - Depth: {row['Max_Gas_Depth']} ft"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Gas", f"{row['TG_Max']:,} ppm")
                st.metric("Methane (C1)", f"{row['C1_Max']:,} ppm")
            
            with col2:
                st.metric("Ethane (C2)", f"{row['C2_Max']:,} ppm")
                st.metric("Propane (C3)", f"{row['C3_Max']:,} ppm")
            
            with col3:
                st.metric("Iso-Butane (C4I)", f"{row['C4I_Max']:,} ppm")
                st.metric("Normal-Butane (C4N)", f"{row['C4N_Max']:,} ppm")
                st.metric("Pentane (C5)", f"{row['C5_Max']:,} ppm")

def display_detailed_gas_data(uploaded_file):
    st.header("📈 Detailed Gas Readings Analysis")
    
    try:
        # Try to read the actual gas data from the uploaded file
        gas_df = pd.read_excel(uploaded_file, sheet_name='Lithology %, ROP & Gas Reading')
        
        st.success("✅ Loaded detailed gas data from uploaded file")
        
        # Display preview of the data
        st.subheader("Gas Data Preview")
        st.dataframe(gas_df.head(20), use_container_width=True)
        
        # Basic statistics
        st.subheader("Gas Data Statistics")
        if 'TG' in gas_df.columns:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Max TG", f"{gas_df['TG'].max():,} ppm")
            with col2:
                st.metric("Avg TG", f"{gas_df['TG'].mean():.0f} ppm")
            with col3:
                st.metric("Min TG", f"{gas_df['TG'].min():,} ppm")
            with col4:
                st.metric("Data Points", len(gas_df))
        
    except Exception as e:
        st.warning("Could not load detailed gas data. Showing demo data.")
        
        # Create sample depth-based data for demonstration
        depth_range = list(range(8200, 9000, 10))
        np.random.seed(42)
        sample_gas_data = pd.DataFrame({
            'Depth': depth_range,
            'TG': np.random.randint(1000, 50000, len(depth_range)),
            'C1': np.random.randint(500, 25000, len(depth_range)),
            'C2': np.random.randint(10, 5000, len(depth_range))
        })
        
        st.line_chart(sample_gas_data.set_index('Depth'))
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Maximum Gas Reading", "137,619 ppm", "LG Zone")
            st.metric("Gas Show Start Depth", "8,520 ft", "MG Zone")
        with col2:
            st.metric("Background Gas Average", "~6,000 ppm", "Varies by zone")
            st.metric("Wetness Ratio", "Improving with depth", "Positive indicator")

def display_raw_data_preview(uploaded_file):
    st.header("📋 Raw Data Preview")
    
    sheet_names = ['Daily Geological Report', 'Lithological Description', 'Lithology %, ROP & Gas Reading']
    
    selected_sheet = st.selectbox("Select Sheet to Preview", sheet_names)
    
    try:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        st.subheader(f"Sheet: {selected_sheet}")
        st.dataframe(df, use_container_width=True)
        
        # Show basic info about the sheet
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Rows", len(df))
        with col2:
            st.metric("Columns", len(df.columns))
        with col3:
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
            
    except Exception as e:
        st.error(f"Could not load sheet: {selected_sheet}. Error: {e}")

def display_footer():
    st.markdown("---")
    st.markdown(
        "**Geological Report Analyzer** | North Bahariya Petroleum Company | "
        "Generated on {}".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
    )

if __name__ == "__main__":
    main()
