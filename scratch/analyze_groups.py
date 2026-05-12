import pandas as pd
import re

file_path = r'2026長庚盃 排名、個人團體對抗表(2026新版).xlsx'
xl = pd.ExcelFile(file_path)

def get_group(name):
    # Simulate the JS logic
    g = name.replace('傳統', '').replace('對抗', '').replace('團體', '').replace('排名', '').replace('賽', '').strip()
    g = re.sub(r'\s*\d+強', '', g).strip()
    return g

groups = {}
for name in xl.sheet_names:
    group = get_group(name)
    if group not in groups:
        groups[group] = []
    groups[group].append(name)

print("--- 系統組別歸類分析 ---")
for g, sheets in groups.items():
    print(f"組別 [{g}] 包含了以下分頁: {sheets}")
