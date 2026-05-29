import pandas as pd
import sys

# Set output encoding to utf-8 for safety
sys.stdout.reconfigure(encoding='utf-8')

file_path = './excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheets: {xl.sheet_names}")
    
    for sheet in ['e-DAHS 전표내역', '세중 발권리스트']:
        if sheet in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            print(f"\n--- {sheet} ---")
            print(f"Columns: {df.columns.tolist()}")
            print(f"Shape: {df.shape}")
            print("First 3 rows:")
            print(df.head(3).to_string())
        else:
            print(f"Sheet '{sheet}' not found!")
            
except Exception as e:
    print(f"Error: {e}")
