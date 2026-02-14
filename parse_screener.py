import pandas as pd
from bs4 import BeautifulSoup
import os

# Paths
HTML_FILE = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\data\raw\screener.html"
OUTPUT_DIR = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\data\raw"

def parse_table(soup, section_id):
    section = soup.find('section', id=section_id)
    if not section:
        print(f"Section {section_id} not found")
        return None
    
    table = section.find('table')
    if not table:
        print(f"Table in {section_id} not found")
        return None
        
    # Extract headers
    headers = [th.text.strip() for th in table.find('thead').find_all('th')]
    
    # Extract rows
    rows = []
    for tr in table.find('tbody').find_all('tr'):
        cells = [td.text.strip() for td in tr.find_all('td')]
        rows.append(cells)
        
    # Create DataFrame
    df = pd.DataFrame(rows, columns=headers)
    return df

def main():
    if not os.path.exists(HTML_FILE):
        print(f"File {HTML_FILE} not found")
        return

    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    # Parse P&L
    df_pl = parse_table(soup, 'profit-loss')
    if df_pl is not None:
        csv_path = os.path.join(OUTPUT_DIR, 'IS.csv')
        df_pl.to_csv(csv_path, index=False)
        print(f"Saved {csv_path}")

    # Parse Balance Sheet
    df_bs = parse_table(soup, 'balance-sheet')
    if df_bs is not None:
        csv_path = os.path.join(OUTPUT_DIR, 'BS.csv')
        df_bs.to_csv(csv_path, index=False)
        print(f"Saved {csv_path}")

    # Parse Cash Flow
    df_cf = parse_table(soup, 'cash-flow')
    if df_cf is not None:
        csv_path = os.path.join(OUTPUT_DIR, 'CF.csv')
        df_cf.to_csv(csv_path, index=False)
        print(f"Saved {csv_path}")

if __name__ == "__main__":
    main()
