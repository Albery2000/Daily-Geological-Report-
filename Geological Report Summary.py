import streamlit as st
import pandas as pd
import io
from datetime import datetime
import subprocess
import sys

def install_package(package):
    """Install a package using pip"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except:
        return False

def check_dependencies():
    """Check if required packages are installed, install if missing"""
    required_packages = ['xlrd', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        st.warning(f"Installing missing dependencies: {', '.join(missing_packages)}")
        for package in missing_packages:
            if install_package(package):
                st.success(f"✓ Successfully installed {package}")
            else:
                st.error(f"✗ Failed to install {package}")
                return False
    return True

def parse_excel_file(uploaded_file):
    """Parse the Excel file and extract required data from all three sheets"""
    try:
        # Try different engines to read Excel file
        engines_to_try = ['openpyxl', 'xlrd']
        
        for engine in engines_to_try:
            try:
                df_daily = pd.read_excel(uploaded_file, sheet_name='Daily Geological Report', header=None, engine=engine)
                df_litho_desc = pd.read_excel(uploaded_file, sheet_name='Lithological Description', header=None, engine=engine)
                df_litho_gas = pd.read_excel(uploaded_file, sheet_name='Lithology %, ROP & Gas Reading', header=None, engine=engine)
                st.success(f"✓ Successfully read file using {engine} engine")
                return df_daily, df_litho_desc, df_litho_gas
            except Exception as e:
                st.info(f"Tried {engine} engine: {str(e)}")
                continue
        
        # If no engine worked, try without specifying engine
        try:
            df_daily = pd.read_excel(uploaded_file, sheet_name='Daily Geological Report', header=None)
            df_litho_desc = pd.read_excel(uploaded_file, sheet_name='Lithological Description', header=None)
            df_litho_gas = pd.read_excel(uploaded_file, sheet_name='Lithology %, ROP & Gas Reading', header=None)
            st.success("✓ Successfully read file using default engine")
            return df_daily, df_litho_desc, df_litho_gas
        except Exception as e:
            st.error(f"Default engine also failed: {str(e)}")
            return None, None, None
            
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None, None, None

def extract_well_info(df_daily):
    """Extract well information from the Daily Geological Report sheet"""
    well_info = {}
    
    # Convert to string for easier searching
    df_str = df_daily.astype(str)
    
    try:
        # Extract concession
        concession_mask = df_str.iloc[:, 0].str.contains('Concession', na=False)
        if concession_mask.any():
            concession_row = df_daily[concession_mask].iloc[0]
            well_info['Concession'] = concession_row[concession_row.last_valid_index()] if concession_row.last_valid_index() else 'North Bahariya'
        
        # Extract date
        date_mask = df_str.iloc[:, 0].str.contains('Date', na=False)
        if date_mask.any():
            date_row = df_daily[date_mask].iloc[0]
            date_value = date_row[date_row.last_valid_index()] if date_row.last_valid_index() else ''
            # Clean date format
            if '00:00:00' in str(date_value):
                date_value = str(date_value).split(' ')[0]
            well_info['Date'] = date_value
        
        # Extract report number
        report_mask = df_str.iloc[:, 0].str.contains('Report No.', na=False)
        if report_mask.any():
            report_row = df_daily[report_mask].iloc[0]
            well_info['Report No.'] = report_row[report_row.last_valid_index()] if report_row.last_valid_index() else ''
        
        # Extract RKB
        rkb_mask = df_str.iloc[:, 0].str.contains('RKB', na=False)
        if rkb_mask.any():
            rkb_row = df_daily[rkb_mask].iloc[0]
            well_info['RKB'] = rkb_row[rkb_row.last_valid_index()] if rkb_row.last_valid_index() else ''
        
        # Extract spud date
        spud_mask = df_str.iloc[:, 0].str.contains('Spud Date', na=False)
        if spud_mask.any():
            spud_row = df_daily[spud_mask].iloc[0]
            spud_value = spud_row[spud_row.last_valid_index()] if spud_row.last_valid_index() else ''
            # Clean date format
            if '00:00:00' in str(spud_value):
                spud_value = str(spud_value).split(' ')[0]
            well_info['Spud Date'] = spud_value
        
        # Extract geologist
        geo_mask = df_str.iloc[:, 0].str.contains('Geologist', na=False)
        if geo_mask.any():
            geo_row = df_daily[geo_mask].iloc[0]
            well_info['Geologist'] = geo_row[geo_row.last_valid_index()] if geo_row.last_valid_index() else 'Soliman Farag'
        else:
            # Try to find geologist name from the bottom of the sheet
            wellsites_mask = df_str.iloc[:, 0].str.contains('Wellsite Geologist', na=False)
            if wellsites_mask.any():
                geo_row = df_daily[wellsites_mask].iloc[0]
                well_info['Geologist'] = geo_row[geo_row.last_valid_index()] if geo_row.last_valid_index() else 'Soliman Farag'
            else:
                well_info['Geologist'] = 'Soliman Farag'
                
    except Exception as e:
        st.warning(f"Could not extract all well information: {str(e)}")
    
    return well_info

def extract_drilling_depths(df_daily):
    """Extract drilling depth information"""
    depths = {}
    
    # Convert to string for easier searching
    df_str = df_daily.astype(str)
    
    try:
        # Find depth rows - looking for time patterns like "24:00 Hrs", "00:00 Hrs", "06:00 Hrs"
        depth_24_mask = df_str.iloc[:, 4].str.contains('24:00', na=False)
        depth_00_mask = df_str.iloc[:, 4].str.contains('00:00', na=False)
        depth_06_mask = df_str.iloc[:, 4].str.contains('06:00', na=False)
        
        if depth_24_mask.any():
            row_idx = df_daily[depth_24_mask].index[0]
            depths['24:00 Hrs Depth'] = df_daily.iloc[row_idx, 7] if pd.notna(df_daily.iloc[row_idx, 7]) else 'N/A'
        
        if depth_00_mask.any():
            row_idx = df_daily[depth_00_mask].index[0]
            depths['00:00 Hrs Depth'] = df_daily.iloc[row_idx, 7] if pd.notna(df_daily.iloc[row_idx, 7]) else 'N/A'
        
        if depth_06_mask.any():
            row_idx = df_daily[depth_06_mask].index[0]
            depths['06:00 Hrs Depth'] = df_daily.iloc[row_idx, 7] if pd.notna(df_daily.iloc[row_idx, 7]) else 'N/A'
        
        # Calculate progress
        try:
            if '00:00 Hrs Depth' in depths and '24:00 Hrs Depth' in depths:
                depth_00 = float(depths['00:00 Hrs Depth']) if depths['00:00 Hrs Depth'] != 'N/A' else 0
                depth_24 = float(depths['24:00 Hrs Depth']) if depths['24:00 Hrs Depth'] != 'N/A' else 0
                depths['Progress (Last 24H)'] = f"{depth_00 - depth_24:.1f} ft"
            
            if '06:00 Hrs Depth' in depths and '00:00 Hrs Depth' in depths:
                depth_06 = float(depths['06:00 Hrs Depth']) if depths['06:00 Hrs Depth'] != 'N/A' else 0
                depth_00 = float(depths['00:00 Hrs Depth']) if depths['00:00 Hrs Depth'] != 'N/A' else 0
                depths['Progress (Last 6H)'] = f"{depth_06 - depth_00:.1f} ft"
        except:
            depths['Progress (Last 24H)'] = 'N/A'
            depths['Progress (Last 6H)'] = 'N/A'
            
    except Exception as e:
        st.warning(f"Could not extract all depth information: {str(e)}")
    
    return depths

def extract_formation_tops(df_daily):
    """Extract formation tops information"""
    formations = []
    
    try:
        # Look for the formation table - typically starts after "Fm. Tops Correlation"
        formation_start = None
        for i, row in df_daily.iterrows():
            if pd.notna(row[0]) and 'Fm. Tops Correlation' in str(row[0]):
                formation_start = i + 3  # Usually starts a few rows after the header
                break
        
        if formation_start:
            # Extract formation data
            for i in range(formation_start, min(formation_start + 20, len(df_daily))):
                row = df_daily.iloc[i]
                if pd.notna(row[2]) and isinstance(row[2], str) and row[2].strip():
                    formation = {
                        'Formation': row[2],
                        'Member': row[4] if pd.notna(row[4]) else '',
                        'Prognosed_MD': row[6] if pd.notna(row[6]) and str(row[6]).replace('.', '').isdigit() else 'N/A',
                        'Prognosed_TVDSS': row[7] if pd.notna(row[7]) and str(row[7]).replace('.', '').isdigit() else 'N/A',
                        'Actual_MD': row[9] if pd.notna(row[9]) and str(row[9]).replace('.', '').isdigit() else 'N/A',
                        'Actual_TVDSS': row[10] if pd.notna(row[10]) and str(row[10]).replace('.', '').isdigit() else 'N/A'
                    }
                    formations.append(formation)
    
    except Exception as e:
        st.warning(f"Could not extract formation tops: {str(e)}")
    
    return formations

def extract_gas_readings(df_daily):
    """Extract gas reading summary"""
    gas_readings = []
    
    try:
        # Look for gas reading sections
        df_str = df_daily.astype(str)
        
        # Find all "Max. Gas Reading at:" occurrences
        gas_sections = df_str[df_str.iloc[:, 0].str.contains('Max. Gas Reading at:', na=False)]
        
        for idx in gas_sections.index:
            formation = df_daily.iloc[idx, 0].replace('Max. Gas Reading at:', '').strip()
            depth = df_daily.iloc[idx, 7] if pd.notna(df_daily.iloc[idx, 7]) else 'N/A'
            
            # Get gas values from next few rows
            tg_row = df_daily.iloc[idx + 2] if idx + 2 < len(df_daily) else None
            c1_row = df_daily.iloc[idx + 2, 7] if tg_row is not None and pd.notna(df_daily.iloc[idx + 2, 7]) else 'N/A'
            
            if formation and formation != '':
                gas_readings.append({
                    'Formation': formation,
                    'Depth': depth,
                    'TG': c1_row if c1_row != 'N/A' else 'N/A'
                })
    
    except Exception as e:
        st.warning(f"Could not extract gas readings: {str(e)}")
    
    return gas_readings

def extract_detailed_gas_readings(df_litho_gas):
    """Extract detailed gas readings from the third sheet"""
    detailed_gas = []
    
    try:
        # Find the header row with gas reading columns
        header_found = False
        header_row_idx = 0
        
        for i, row in df_litho_gas.iterrows():
            if any('TG' in str(cell) for cell in row if pd.notna(cell)):
                header_row_idx = i
                header_found = True
                break
        
        if header_found:
            # Extract data rows
            for i in range(header_row_idx + 1, min(header_row_idx + 50, len(df_litho_gas))):
                row = df_litho_gas.iloc[i]
                if pd.notna(row[0]) and str(row[0]).replace('.', '').isdigit():
                    gas_data = {
                        'Depth': row[0],
                        'TG': row[8] if pd.notna(row[8]) else 'N/A',
                        'C1': row[9] if pd.notna(row[9]) else 'N/A',
                        'C2': row[10] if pd.notna(row[10]) else 'N/A',
                        'C3': row[11] if pd.notna(row[11]) else 'N/A',
                        'C4I': row[12] if pd.notna(row[12]) else 'N/A',
                        'C4N': row[13] if pd.notna(row[13]) else 'N/A',
                        'C5': row[14] if pd.notna(row[14]) else 'N/A'
                    }
                    detailed_gas.append(gas_data)
    
    except Exception as e:
        st.warning(f"Could not extract detailed gas readings: {str(e)}")
    
    return detailed_gas

def main():
    st.set_page_config(page_title="Geological Report Analyzer", layout="wide")
    st.title("🏢 Daily Geological Report Analyzer")
    st.write("Upload your Daily Geological Report Excel file to generate a comprehensive analysis")
    
    # Check and install dependencies
    if not check_dependencies():
        st.error("Failed to install required dependencies. Please install them manually:")
        st.code("pip install xlrd openpyxl")
        return
    
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        with st.spinner("Analyzing geological report..."):
            # Parse the Excel file
            df_daily, df_litho_desc, df_litho_gas = parse_excel_file(uploaded_file)
            
            if df_daily is not None:
                # Extract information from all sheets
                well_info = extract_well_info(df_daily)
                drilling_depths = extract_drilling_depths(df_daily)
                formation_tops = extract_formation_tops(df_daily)
                gas_summary = extract_gas_readings(df_daily)
                detailed_gas = extract_detailed_gas_readings(df_litho_gas)
                
                # Display results
                st.success("Report analysis completed!")
                
                # Well Information Section
                st.header("🏢 Well Information")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Concession", well_info.get('Concession', 'N/A'))
                    st.metric("Date", well_info.get('Date', 'N/A'))
                
                with col2:
                    st.metric("Report No.", well_info.get('Report No.', 'N/A'))
                    st.metric("RKB", f"{well_info.get('RKB', 'N/A')} ft")
                
                with col3:
                    st.metric("Spud Date", well_info.get('Spud Date', 'N/A'))
                    st.metric("Geologist", well_info.get('Geologist', 'N/A'))
                
                # Drilling Progress Section
                st.header("⛏️ Drilling Progress")
                depth_col1, depth_col2, depth_col3, depth_col4, depth_col5 = st.columns(5)
                
                with depth_col1:
                    st.metric("24:00 Hrs Depth", f"{drilling_depths.get('24:00 Hrs Depth', 'N/A')} ft")
                with depth_col2:
                    st.metric("00:00 Hrs Depth", f"{drilling_depths.get('00:00 Hrs Depth', 'N/A')} ft")
                with depth_col3:
                    st.metric("06:00 Hrs Depth", f"{drilling_depths.get('06:00 Hrs Depth', 'N/A')} ft")
                with depth_col4:
                    st.metric("Progress (Last 24H)", drilling_depths.get('Progress (Last 24H)', 'N/A'))
                with depth_col5:
                    st.metric("Progress (Last 6H)", drilling_depths.get('Progress (Last 6H)', 'N/A'))
                
                # Formation Tops Section
                st.header("🗻 Formation Tops")
                if formation_tops:
                    formation_data = []
                    for formation in formation_tops:
                        formation_data.append([
                            formation['Formation'],
                            formation['Member'],
                            formation['Prognosed_MD'],
                            formation['Prognosed_TVDSS'],
                            formation['Actual_MD'],
                            formation['Actual_TVDSS']
                        ])
                    
                    formation_df = pd.DataFrame(formation_data, 
                                              columns=['Formation', 'Member', 'Prognosed MD', 'Prognosed TVDSS', 
                                                      'Actual MD', 'Actual TVDSS'])
                    st.dataframe(formation_df, use_container_width=True)
                else:
                    st.info("No formation tops data found in the report")
                
                # Gas Reading Summary Section
                st.header("🔥 Gas Reading Summary")
                if gas_summary:
                    gas_data = []
                    for gas in gas_summary:
                        gas_data.append([
                            gas['Formation'],
                            gas['Depth'],
                            gas['TG']
                        ])
                    
                    gas_df = pd.DataFrame(gas_data, columns=['Formation', 'Depth', 'TG (Max)'])
                    st.dataframe(gas_df, use_container_width=True)
                else:
                    st.info("No gas reading summary found in the report")
                
                # Detailed Gas Readings Section
                st.header("📊 Detailed Gas Readings")
                if detailed_gas:
                    detailed_gas_data = []
                    for gas in detailed_gas:
                        detailed_gas_data.append([
                            gas['Depth'],
                            gas['TG'],
                            gas['C1'],
                            gas['C2'],
                            gas['C3'],
                            gas['C4I'],
                            gas['C4N'],
                            gas['C5']
                        ])
                    
                    detailed_gas_df = pd.DataFrame(detailed_gas_data, 
                                                 columns=['Depth', 'TG', 'C1', 'C2', 'C3', 'C4I', 'C4N', 'C5'])
                    st.dataframe(detailed_gas_df, use_container_width=True)
                else:
                    st.info("No detailed gas readings found in the report")
                
                # Export Section
                st.header("📤 Export Report")
                
                # Create a summary report
                summary_report = f"""
                GEOLOGICAL REPORT SUMMARY
                =========================
                
                Well Information:
                - Concession: {well_info.get('Concession', 'N/A')}
                - Date: {well_info.get('Date', 'N/A')}
                - Report No.: {well_info.get('Report No.', 'N/A')}
                - RKB: {well_info.get('RKB', 'N/A')} ft
                - Spud Date: {well_info.get('Spud Date', 'N/A')}
                - Geologist: {well_info.get('Geologist', 'N/A')}
                
                Drilling Progress:
                - 24:00 Hrs Depth: {drilling_depths.get('24:00 Hrs Depth', 'N/A')} ft
                - 00:00 Hrs Depth: {drilling_depths.get('00:00 Hrs Depth', 'N/A')} ft
                - 06:00 Hrs Depth: {drilling_depths.get('06:00 Hrs Depth', 'N/A')} ft
                - Progress (Last 24H): {drilling_depths.get('Progress (Last 24H)', 'N/A')}
                - Progress (Last 6H): {drilling_depths.get('Progress (Last 6H)', 'N/A')}
                
                Formation Tops: {len(formation_tops)} formations identified
                Gas Readings: {len(gas_summary)} formations with gas shows
                Detailed Gas Data: {len(detailed_gas)} depth points analyzed
                """
                
                st.download_button(
                    label="Download Summary Report",
                    data=summary_report,
                    file_name=f"geological_report_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                    mime="text/plain"
                )
                
            else:
                st.error("Failed to parse the Excel file. Please check the file format and try again.")

if __name__ == "__main__":
    main()
