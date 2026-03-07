import pandas as pd
import json
import os

# File paths
file_primary = 'IPAK Group Dashboard.xlsx'
js_file = 'data.js'

def convert_excel_to_js():
    if not os.path.exists(file_primary):
        print(f"Error: {file_primary} not found.")
        return

    try:
        print(f"Reading {file_primary}...")
        xl_pri = pd.ExcelFile(file_primary)
        
        # --- PROCESS MAIN DATA ---
        main_data = []
        if 'Main Data' in xl_pri.sheet_names:
            main_df = xl_pri.parse('Main Data')
            main_df.columns = main_df.columns.str.strip()
            main_df = main_df.rename(columns={'Date': 'Month'})
            if 'Month' in main_df.columns:
                main_df['Month'] = main_df['Month'].astype(str)
            
            # Fill NaN with 0 for numeric/safer JSON
            main_df = main_df.fillna(0)
            main_data = main_df.to_dict(orient='records')
            print(f"Main Data: {len(main_data)} records")
        else:
            print("Warning: 'Main Data' sheet not found.")

        # --- PROCESS PRODUCT DATA ---
        product_data = []
        if 'Product' in xl_pri.sheet_names:
            product_df = xl_pri.parse('Product')
            product_df.columns = product_df.columns.str.strip()
            product_df = product_df.rename(columns={'Date': 'Month'})
            if 'Month' in product_df.columns:
                product_df['Month'] = product_df['Month'].astype(str)
            
            # Extract Product Type if available
            if 'Product Type' not in product_df.columns:
                print("Warning: 'Product Type' column not found in Product sheet. Defaulting to 'General'.")
                product_df['Product Type'] = 'General'
            else:
                product_df['Product Type'] = product_df['Product Type'].fillna('General').astype(str).str.strip()

            product_df = product_df.fillna(0)
            product_data = product_df.to_dict(orient='records')
            print(f"Product Data: {len(product_data)} records")

        # --- PROCESS EREMA DATA ---
        erema_data = []
        if 'EREMA' in xl_pri.sheet_names:
            erema_df = xl_pri.parse('EREMA')
            erema_df.columns = erema_df.columns.str.strip()
            erema_df = erema_df.rename(columns={'Date': 'Month'})
            if 'Month' in erema_df.columns:
                erema_df['Month'] = erema_df['Month'].astype(str)
            erema_df = erema_df.fillna(0)
            erema_data = erema_df.to_dict(orient='records')
            print(f"Erema Data: {len(erema_data)} records")

        # --- PROCESS TRIALS DATA ---
        trials_data = []
        if 'Trials' in xl_pri.sheet_names:
            trials_df = xl_pri.parse('Trials')
            trials_df.columns = trials_df.columns.str.strip()
            
            # Keep Date as Date, formatted as YYYY-MM-DD
            if 'Date' in trials_df.columns:
                trials_df['Date'] = pd.to_datetime(trials_df['Date']).dt.strftime('%Y-%m-%d')
            
            # Extract Product Type if available
            if 'Product Type' not in trials_df.columns:
                print("Warning: 'Product Type' column not found in Trials sheet. Defaulting to 'General'.")
                trials_df['Product Type'] = 'General'
            else:
                trials_df['Product Type'] = trials_df['Product Type'].fillna('General').astype(str).str.strip()

            # Fill remaining numeric NaNs with 0
            trials_df = trials_df.fillna(0)
            trials_data = trials_df.to_dict(orient='records')
            print(f"Trials Data: {len(trials_data)} records")

        # --- PROCESS OFFLINE DATA ---
        offline_data = []
        if 'OFFLine' in xl_pri.sheet_names:
            offline_df = xl_pri.parse('OFFLine')
            offline_df.columns = offline_df.columns.str.strip()
            offline_df = offline_df.rename(columns={'Date': 'Month'})
            if 'Month' in offline_df.columns:
                offline_df['Month'] = offline_df['Month'].astype(str)
            offline_df = offline_df.fillna(0)
            offline_data = offline_df.to_dict(orient='records')
            print(f"Offline Data: {len(offline_data)} records")

        # --- PROCESS RECON DATA ---
        recon_data = []
        recon_codes = {}
        if 'Recon' in xl_pri.sheet_names:
            # Read header normally (Row 1 in Excel is header)
            recon_full_df = xl_pri.parse('Recon')
            recon_full_df.columns = recon_full_df.columns.str.strip()
            
            # Rows 0-3 (Excel rows 2-5) contain the codes for Plants
            # We enforce a specific structure based on user input
            # Row 0: IPAK codes
            # Row 1: CPAK codes
            # Row 2: GPAK codes
            # Row 3: PETPAK codes
            
            # Extract Codes Metadata
            # Iterate through the first few rows to build the code map
            # We assume the 'Plant' column exists to identify which row belongs to which plant
            
            # Slice the first 4 rows for codes
            metadata_df = recon_full_df.iloc[:4].copy()
            
            for _, row in metadata_df.iterrows():
                plant_name = row.get('Plant')
                if plant_name:
                    plant_codes = {}
                    for col in recon_full_df.columns:
                        if col not in ['Month', 'Plant']:
                            plant_codes[col] = row[col]
                    recon_codes[plant_name] = plant_codes
            
            print(f"Recon Codes extracted for: {list(recon_codes.keys())}")

            # Extract Data (Rows 4 onwards)
            recon_df = recon_full_df.iloc[4:].copy()
            
            # Rename Date to Month if needed
            recon_df = recon_df.rename(columns={'Date': 'Month'})
            
            # Ensure Month is string
            if 'Month' in recon_df.columns:
                recon_df['Month'] = recon_df['Month'].astype(str)
            
            recon_df = recon_df.fillna(0)
            recon_data = recon_df.to_dict(orient='records')
            print(f"Recon Data: {len(recon_data)} records")
            
        # --- SUMMARY KPIS ---
        # No longer reading from secondary file. Keeping empty for now.
        # Dashboard logic will be updated to calculate totals dynamically.
        summary_kpis = {} 

        # Construct the final data object
        dashboard_data = {
            'mainData': main_data,
            'productData': product_data,
            'eremaData': erema_data,
            'trialsData': trials_data,
            'offlineData': offline_data,
            'reconData': recon_data,
            'reconCodes': recon_codes,
            'summaryKPIs': summary_kpis
        }

        # Write to data.js
        with open(js_file, 'w') as f:
            f.write(f"const dashboardData = {json.dumps(dashboard_data, indent=4)};\n")
            
        print(f"Successfully converted to {js_file}")
        
    except Exception as e:
        print(f"Error converting data: {e}")

if __name__ == "__main__":
    convert_excel_to_js()
