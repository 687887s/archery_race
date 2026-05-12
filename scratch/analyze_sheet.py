import pandas as pd
import sys

file_path = r'2026長庚盃 排名、個人團體對抗表(2026新版).xlsx'
sheet_keywords = ['傳統30', '16強']

try:
    xl = pd.ExcelFile(file_path)
    sheet_name = next((n for n in xl.sheet_names if all(k in n for k in sheet_keywords)), None)
    
    # Read the sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    print(f"--- [ {sheet_name} ] 非空數據分析 ---")
    for r_idx, row in df.iterrows():
        # Filter out NaN values for display
        clean_row = {c_idx: val for c_idx, val in row.items() if pd.notna(val)}
        if clean_row:
            # Format: Row Number | Col: Val | Col: Val ...
            row_str = f"Row {r_idx:2} | " + " | ".join([f"C{c}: {v}" for c, v in clean_row.items()])
            print(row_str)
            
except Exception as e:
    print(f"錯誤: {e}")
