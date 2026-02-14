import pandas as pd
import os

# Test File Path
TEST_FILE = r"C:\Users\DELL\.gemini\antigravity\scratch\HINDALCO_Equity_Research\models\test_model.xlsx"

def main():
    if not os.path.exists(os.path.dirname(TEST_FILE)):
        os.makedirs(os.path.dirname(TEST_FILE))

    df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})

    try:
        with pd.ExcelWriter(TEST_FILE, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        print(f"Successfully created {TEST_FILE}")
    except Exception as e:
        print(f"Error creating file: {e}")

if __name__ == "__main__":
    main()
