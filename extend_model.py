import pandas as pd
import xlsxwriter
import os

# Paths
EXISTING_MODEL = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\models\three_statement_HINDALCO.xlsx"
OUTPUT_FILE = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\models\three_statement_HINDALCO_v2.xlsx"

def main():
    # 1. Read Existing Data
    # We use openpyxl engine just to read the dataframes from the file we just made (or use the raw CSVs again)
    # Using raw CSVs is safer because we know the structure perfectly.
    # But let's verify we can read the file we made.
    try:
        xls = pd.ExcelFile(EXISTING_MODEL)
        df_inputs = pd.read_excel(xls, 'Inputs')
        df_is = pd.read_excel(xls, 'IS')
        df_bs = pd.read_excel(xls, 'BS')
        df_cf = pd.read_excel(xls, 'CF')
    except Exception as e:
        print(f"Error reading existing model: {e}")
        return

    # 2. Prepare Peer Data (Comps)
    peers_data = {
        "Company": ["Hindalco", "Vedanta", "Tata Steel", "NALCO", "JSW Steel", "SAIL"],
        "Ticker": ["HINDALCO", "VEDL", "TATASTEEL", "NATIONALUM", "JSWSTEEL", "SAIL"],
        "Price (INR)": [964, 595, 185, 365, 1250, 161], # Approx recent prices
        "Market Cap (Cr)": [217138, 269836, 259907, 67551, 306024, 66831],
        "EV (Cr)": [280000, 320000, 340000, 65000, 380000, 100000], # Rough estimates based on debt
        "Revenue (Cr)": [238496, 150000, 230000, 14000, 175000, 105000],
        "EBITDA (Cr)": [26000, 35000, 24000, 3500, 29000, 10000],
        "P/E": [12.1, 18.6, 28.3, 12.9, 40.7, 24.0],
        "EV/EBITDA": [8.0, 7.8, 6.8, 7.2, 14.2, 9.1],
        "P/B": [2.2, 2.5, 2.8, 2.1, 3.5, 1.2]
    }
    df_comps = pd.DataFrame(peers_data)

    # 3. Create New Excel File with All Sheets
    workbook = xlsxwriter.Workbook(OUTPUT_FILE)
    
    # --- Formatting ---
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    fmt_number = workbook.add_format({'num_format': '#,##0.00'})
    fmt_currency = workbook.add_format({'num_format': '₹#,##0'})
    fmt_percent = workbook.add_format({'num_format': '0.0%'})
    fmt_bold = workbook.add_format({'bold': True})

    # --- INPUTS Sheet ---
    ws_inputs = workbook.add_worksheet('Inputs')
    # Write Header
    headers = list(df_inputs.columns)
    for col, h in enumerate(headers):
        ws_inputs.write(0, col, h, fmt_header)
    # Write Data
    for row_idx, row in df_inputs.iterrows():
        ws_inputs.write(row_idx+1, 0, row['Item'])
        # Fill forecast years with defaults
        ws_inputs.write(row_idx+1, 1, 0.0) # Historical placeholder
        ws_inputs.write(row_idx+1, 2, 0.0)
        ws_inputs.write(row_idx+1, 3, 0.0)
        ws_inputs.write(row_idx+1, 4, 0.0)
        # FY26-FY30 Assumptions (Hardcoded for now, user can change)
        defaults = [0.08, 0.14, 0.25, 0.05, 45] # Growth, Margin, Tax, Capex, WC Days
        val = defaults[row_idx] if row_idx < 5 else 0
        ws_inputs.write(row_idx+1, 5, val, fmt_percent if row_idx < 4 else fmt_number)
        ws_inputs.write(row_idx+1, 6, val, fmt_percent if row_idx < 4 else fmt_number)
        ws_inputs.write(row_idx+1, 7, val, fmt_percent if row_idx < 4 else fmt_number)
        ws_inputs.write(row_idx+1, 8, val, fmt_percent if row_idx < 4 else fmt_number)
        ws_inputs.write(row_idx+1, 9, val, fmt_percent if row_idx < 4 else fmt_number)

    # WACC Assumptions Table in Inputs
    ws_inputs.write(8, 0, "WACC Assumptions", fmt_bold)
    ws_inputs.write(9, 0, "Risk Free Rate")
    ws_inputs.write(9, 1, 0.067, fmt_percent)
    ws_inputs.write(10, 0, "Beta")
    ws_inputs.write(10, 1, 1.20, fmt_number)
    ws_inputs.write(11, 0, "Market Premium")
    ws_inputs.write(11, 1, 0.05, fmt_percent)
    ws_inputs.write(12, 0, "Cost of Debt (Pre-tax)")
    ws_inputs.write(12, 1, 0.075, fmt_percent)
    ws_inputs.write(13, 0, "Tax Rate")
    ws_inputs.write(13, 1, 0.25, fmt_percent)
    ws_inputs.write(14, 0, "Target Debt/Equity")
    ws_inputs.write(14, 1, 0.50, fmt_percent)
    
    # --- IS, BS, CF Sheets (Write Historical Data) ---
    for name, df in [('IS', df_is), ('BS', df_bs), ('CF', df_cf)]:
        ws = workbook.add_worksheet(name)
        # Write headers
        for col, h in enumerate(df.columns):
            ws.write(0, col, h, fmt_header)
        # Write data
        for row_idx, row in df.iterrows():
            for col_idx, val in enumerate(row):
                ws.write(row_idx+1, col_idx, val, fmt_currency if col_idx > 0 else None)

    # --- DCF Sheet ---
    ws_dcf = workbook.add_worksheet('DCF')
    
    # Setup timeline
    years = ["FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    ws_dcf.write(0, 0, "DCF Model", fmt_bold)
    for i, yr in enumerate(years):
        ws_dcf.write(0, i+1, yr, fmt_header)
        
    # Link Inputs for Growth/Margins
    # Row 1: Revenue
    ws_dcf.write(1, 0, "Revenue")
    # Base revenue from IS last year (Assume column 5 of IS is FY25)
    last_rev = df_is.iloc[1, 5] if len(df_is) > 1 else 238000 # Fallback
    ws_dcf.write(1,1, f"=IS!F3*(1+Inputs!F2)", fmt_currency)
    ws_dcf.write(1,2, f"=B2*(1+Inputs!G2)", fmt_currency)
    ws_dcf.write(1,3, f"=C2*(1+Inputs!H2)", fmt_currency)
    ws_dcf.write(1,4, f"=D2*(1+Inputs!I2)", fmt_currency)
    ws_dcf.write(1,5, f"=E2*(1+Inputs!J2)", fmt_currency)

    # Row 2: EBITDA
    ws_dcf.write(2, 0, "EBITDA")
    col_letters = ['B', 'C', 'D', 'E', 'F']
    for i, col in enumerate(col_letters):
        ws_dcf.write(2, i+1, f"={col}2*Inputs!{col}3", fmt_currency)

    # Row 3: D&A (Assume % of Sales or Flat growth, let's keep it simple: 3% of sales)
    ws_dcf.write(3, 0, "D&A")
    for i, col in enumerate(col_letters):
        ws_dcf.write(3, i+1, f"={col}2*0.03", fmt_currency)

    # Row 4: EBIT
    ws_dcf.write(4, 0, "EBIT")
    for i, col in enumerate(col_letters):
        ws_dcf.write(4, i+1, f"={col}3-{col}4", fmt_currency)

    # Row 5: Tax
    ws_dcf.write(5, 0, "Tax")
    for i, col in enumerate(col_letters):
        ws_dcf.write(5, i+1, f"={col}5*Inputs!B14", fmt_currency) # Fixed tax rate from Inputs

    # Row 6: NOPAT
    ws_dcf.write(6, 0, "NOPAT")
    for i, col in enumerate(col_letters):
        ws_dcf.write(6, i+1, f"={col}5-{col}6", fmt_currency)
        
    # Row 7: Capex
    ws_dcf.write(7, 0, "Capex")
    for i, col in enumerate(col_letters):
        ws_dcf.write(7, i+1, f"={col}2*Inputs!{col}5", fmt_currency)

    # Row 8: Change in WC (simplified)
    ws_dcf.write(8, 0, "Change in WC")
    for i, col in enumerate(col_letters):
        ws_dcf.write(8, i+1, f"={col}2*0.02", fmt_currency) # 2% of sales

    # Row 9: FCFF
    ws_dcf.write(9, 0, "FCFF", fmt_bold)
    for i, col in enumerate(col_letters):
        ws_dcf.write(9, i+1, f"={col}7+{col}4-{col}8-{col}9", fmt_currency)

    # Discounting
    ws_dcf.write(11, 0, "WACC Calculation", fmt_bold)
    ws_dcf.write(12, 0, "Cost of Equity (Ke)")
    ws_dcf.write(12, 1, "=Inputs!B10 + Inputs!B11*Inputs!B12", fmt_percent)
    ws_dcf.write(13, 0, "Cost of Debt (Kd, After-tax)")
    ws_dcf.write(13, 1, "=Inputs!B13*(1-Inputs!B14)", fmt_percent)
    ws_dcf.write(14, 0, "WACC")
    # Simple WACC: Ke * 0.6 + Kd * 0.4 (Assuming 40% Debt / 60% Equity for simplicity or link to D/E)
    ws_dcf.write(14, 1, "=B13*0.6 + B14*0.4", fmt_percent)

    ws_dcf.write(16, 0, "Discount Factor")
    for i, col in enumerate(col_letters):
        ws_dcf.write(16, i+1, f"=1/(1+$B$15)^{i+1}", fmt_number)
        
    ws_dcf.write(17, 0, "Present Value of FCFF")
    for i, col in enumerate(col_letters):
        ws_dcf.write(17, i+1, f"={col}10*{col}17", fmt_currency)
    
    # Terminal Value
    ws_dcf.write(19, 0, "Terminal Value Assumptions", fmt_bold)
    ws_dcf.write(20, 0, "Terminal Growth Rate")
    ws_dcf.write(20, 1, 0.03, fmt_percent)
    ws_dcf.write(21, 0, "Implied TV (Gordon Growth)")
    # TV = Final FCFF * (1+g) / (WACC - g)
    ws_dcf.write(21, 1, "=F10*(1+B21)/(B15-B21)", fmt_currency)
    ws_dcf.write(22, 0, "PV of Terminal Value")
    ws_dcf.write(22, 1, "=B22*F17", fmt_currency)
    
    # Valuation Summary
    ws_dcf.write(24, 0, "Enterprise Value", fmt_bold)
    ws_dcf.write(24, 1, "=SUM(B18:F18)+B23", fmt_currency)
    ws_dcf.write(25, 0, "Less: Net Debt")
    ws_dcf.write(25, 1, 40000, fmt_currency) # Placeholder estimate from BS
    ws_dcf.write(26, 0, "Equity Value", fmt_bold)
    ws_dcf.write(26, 1, "=B25-B26", fmt_currency)
    ws_dcf.write(27, 0, "Shares Outstanding (Cr)")
    ws_dcf.write(27, 1, 222, fmt_number) # Approx shares for Hindalco
    ws_dcf.write(28, 0, "Implied Share Price", fmt_bold)
    ws_dcf.write(28, 1, "=B27/B28", fmt_currency)

    # --- Comps Sheet ---
    ws_comps = workbook.add_worksheet('Comps')
    headers = list(df_comps.columns)
    for col, h in enumerate(headers):
        ws_comps.write(0, col, h, fmt_header)
    for row_idx, row in df_comps.iterrows():
        for col_idx, val in enumerate(row):
            ws_comps.write(row_idx+1, col_idx, val, fmt_number if isinstance(val, float) else None)
            
    # --- Scenarios Sheet ---
    ws_scen = workbook.add_worksheet('Scenarios')
    ws_scen.write(0, 0, "Scenario Analysis", fmt_bold)
    
    scenarios = ["Bear", "Base", "Bull"]
    metrics = ["Revenue Growth", "EBITDA Margin", "WACC"]
    
    ws_scen.write(1, 0, "Metric")
    for i, s in enumerate(scenarios):
        ws_scen.write(1, i+1, s, fmt_header)
        
    values = [
        [0.04, 0.08, 0.12], # Growth
        [0.10, 0.14, 0.18], # Margin
        [0.12, 0.10, 0.09]  # WACC
    ]
    
    for r, metric in enumerate(metrics):
        ws_scen.write(r+2, 0, metric)
        for c, val in enumerate(values[r]):
            ws_scen.write(r+2, c+1, val, fmt_percent)

    ws_scen.write(6, 0, "Instructions: Copy these values to Inputs sheet to test scenarios.")

    workbook.close()
    print(f"Extended model saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
