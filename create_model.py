import pandas as pd
import os

# Paths
DATA_DIR = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\data\raw"
OUTPUT_FILE = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\models\three_statement_HINDALCO.xlsx"

def clean_currency(x):
    if isinstance(x, str):
        x = x.replace(',', '')
        if x == '' or x == '-': # Handle empty or dash
            return 0
        try:
            return float(x)
        except ValueError:
            return x
    return x

def process_csv(filename, required_years):
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        print(f"File {filename} not found")
        return pd.DataFrame() # Return empty DF

    df = pd.read_csv(path)
    
    # Rename first column if unnamed
    if df.columns[0].startswith('Unnamed'):
        df.rename(columns={df.columns[0]: 'Item'}, inplace=True)

    # Filter columns
    cols = ['Item'] + [c for c in df.columns if c in required_years]
    # Check if all required years are present, if not, take what's available
    available_cols = ['Item'] + [c for c in required_years if c in df.columns]
    
    if len(available_cols) < 2:
        print(f"Warning: No matching year columns found in {filename}")
        return df # Return original if filtration fails
        
    df = df[available_cols]

    # Clean numeric data
    for col in available_cols[1:]:
        df[col] = df[col].apply(clean_currency)
        
    return df

def main():
    if not os.path.exists(os.path.dirname(OUTPUT_FILE)):
        os.makedirs(os.path.dirname(OUTPUT_FILE))

    # Define years to keep (Last 5 years: FY21-FY25)
    # The CSV headers are "Mar 2021", "Mar 2022" etc.
    years = ["Mar 2021", "Mar 2022", "Mar 2023", "Mar 2024", "Mar 2025"]

    df_is = process_csv("IS.csv", years)
    df_bs = process_csv("BS.csv", years)
    df_cf = process_csv("CF.csv", years)
    
    # Create Inputs Sheet DF
    inputs_data = {
        "Item": [
            "Revenue Growth %", 
            "EBITDA Margin %", 
            "Effective Tax Rate %", 
            "Capex as % of Sales", 
            "Working Capital Days"
        ]
    }
    # Add empty columns for the years
    for yr in years:
        inputs_data[yr] = [None] * 5
        
    df_inputs = pd.DataFrame(inputs_data)

    # Write to Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        df_inputs.to_excel(writer, sheet_name='Inputs', index=False)
        df_is.to_excel(writer, sheet_name='IS', index=False)
        df_bs.to_excel(writer, sheet_name='BS', index=False)
        df_cf.to_excel(writer, sheet_name='CF', index=False)

    print(f"Created Excel model at {OUTPUT_FILE}")
    
    # Summarize key lines
    print("\n--- Key Financials (Mar 2025) ---")
    
    print("Income Statement:")
    if not df_is.empty and 'Mar 2025' in df_is.columns:
        sales = df_is[df_is['Item'].str.contains("Sales", case=False, na=False)]
        net_profit = df_is[df_is['Item'].str.contains("Net Profit", case=False, na=False)]
        print(sales[['Item', 'Mar 2025']].to_string(index=False))
        print(net_profit[['Item', 'Mar 2025']].to_string(index=False))

    print("\nBalance Sheet:")
    if not df_bs.empty and 'Mar 2025' in df_bs.columns:
        total_assets = df_bs[df_bs['Item'].str.contains("Total Assets", case=False, na=False)]
        borrowings = df_bs[df_bs['Item'].str.contains("Borrowings", case=False, na=False)]
        print(total_assets[['Item', 'Mar 2025']].to_string(index=False))
        print(borrowings[['Item', 'Mar 2025']].to_string(index=False))

    print("\nCash Flow:")
    if not df_cf.empty and 'Mar 2025' in df_cf.columns:
        ops_cf = df_cf[df_cf['Item'].str.contains("Cash from Operating", case=False, na=False)]
        print(ops_cf[['Item', 'Mar 2025']].to_string(index=False))

if __name__ == "__main__":
    main()
