import pandas as pd

file_path = r'2026長庚盃 排名、個人團體對抗表(2026新版).xlsx'

def analyze_sheet(name):
    print(f"\n--- 分析分頁: [{name}] ---")
    try:
        df = pd.read_excel(file_path, sheet_name=name, header=None)
        for r_idx, row in df.iloc[:30].iterrows():
            clean_row = {c_idx: val for c_idx, val in row.items() if pd.notna(val)}
            if clean_row:
                row_str = f"Row {r_idx:2} | " + " | ".join([f"C{c}: {v}" for c, v in clean_row.items()])
                print(row_str)
    except Exception as e:
        print(f"無法讀取 {name}: {e}")

analyze_sheet('傳統30 16強')
analyze_sheet('反曲70 16強')
