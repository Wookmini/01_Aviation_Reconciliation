import pandas as pd
import os

file_path = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheets: {xl.sheet_names}")
    
    for sheet in xl.sheet_names[:4]: # 상위 4개 시트만 확인
        df = pd.read_excel(file_path, sheet_name=sheet).head(3)
        print(f"\n--- Sheet: {sheet} ---")
        print(df.columns.tolist())
except Exception as e:
    print(f"Error: {e}")
