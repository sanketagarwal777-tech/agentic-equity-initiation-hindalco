import pandas as pd
import os

# Paths
DATA_DIR = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\data\raw"
OUTPUT_FILE = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\models\three_statement_HINDALCO.xlsx"

def clean_currency(x):
    """ Cleans currency strings to float, handling errors gracefully. """
    if pd.isna(x):
        return 0
    if isinstance(x, (int, float)):
        return x
    if isinstance(x, str):
        x = x.strip().replace(',', '')
        if x in ['', '-', '--']: # Handle common empty/zero representations
            return 0
        try:
            return float(x)
        except ValueError:
            return 0 # Default to 0 if parsing fails
    return 0

def process_csv(filename, required_years):
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        print(f"File {filename} not found")
        return pd.DataFrame()

    try:
        df = pd.read_csv(path)
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return pd.DataFrame()
    
    # Standardize 'Item' column name
    if df.columns[0].startswith('Unnamed'): 
        df.rename(columns={df.columns[0]: 'Item'}, inplace=True)
    elif 'Item' not in df.columns:
         # Best guess: first column is the item description
         df.rename(columns={df.columns[0]: 'Item'}, inplace=True)

    # Filter columns to keep only Item + requested years
    cols_to_keep = ['Item'] + [c for c in required_years if c in df.columns]
    
    if len(cols_to_keep) < 2:
        print(f"Warning: No matching year columns found in {filename}")
        return df # Return original if filtration fails
        
    df = df[cols_to_keep]

    # Clean numeric data column by column
    for col in cols_to_keep[1:]:
        df[col] = df[col].apply(clean_currency)
        
    return df

def main():
    # Ensure output directory exists
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

    # Remove existing file to avoid permission errors if open (though specific error handling is better)
    if os.path.exists(OUTPUT_FILE):
        try:
            os.remove(OUTPUT_FILE)
        except PermissionError:
            print(f"Error: Cannot delete {OUTPUT_FILE}. It might be open.")
            return

    # Define years
    years = ["Mar 2021", "Mar 2022", "Mar 2023", "Mar 2024", "Mar 2025"]

    df_is = process_csv("IS.csv", years)
    df_bs = process_csv("BS.csv", years)
    df_cf = process_csv("CF.csv", years)
    
    # Create Inputs Sheet Data
    inputs_data = {
        "Item": [
            "Revenue Growth %", 
            "EBITDA Margin %", 
            "Effective Tax Rate %", 
            "Capex as % of Sales", 
            "Working Capital Days"
        ]
    }
    # Add empty columns for usage in Excel
    for yr in years:
        inputs_data[yr] = [None] * len(inputs_data["Item"])
        
    df_inputs = pd.DataFrame(inputs_data)

    # Write to Excel using xlsxwriter engine
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
            df_inputs.to_excel(writer, sheet_name='Inputs', index=False)
            df_is.to_excel(writer, sheet_name='IS', index=False)
            df_bs.to_excel(writer, sheet_name='BS', index=False)
            df_cf.to_excel(writer, sheet_name='CF', index=False)
        print(f"Successfully created Excel model at {OUTPUT_FILE}")
    except Exception as e:
        print(f"Failed to create Excel file: {e}")

if __name__ == "__main__":
    main()
