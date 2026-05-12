import pandas as pd
import re

file_path = r'2026長庚盃 排名、個人團體對抗表(2026新版).xlsx'
xl = pd.ExcelFile(file_path)

def get_group(name):
    # Old logic (BROKEN): Strips '傳統', '反曲'
    # g = name.replace('傳統', '').replace('反曲', '')...
    
    # Proposed new logic: Keep the Bow Type
    g = name.replace('對抗', '').replace('排名', '').replace('賽', '').strip()
    g = re.sub(r'\s*\d+強', '', g).strip()
    return g

groups = {}
for name in xl.sheet_names:
    group = get_group(name)
    if group not in groups:
        groups[group] = []
    groups[group].append(name)

with open('scratch/group_analysis.txt', 'w', encoding='utf-8') as f:
    f.write("--- 修正後的組別歸類分析 ---\n")
    for g, sheets in groups.items():
        f.write(f"組別 [{g}] 包含了以下分頁: {sheets}\n")
