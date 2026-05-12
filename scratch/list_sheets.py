import pandas as pd

file_path = r'2026長庚盃 排名、個人團體對抗表(2026新版).xlsx'
xl = pd.ExcelFile(file_path)

with open('scratch/sheets.txt', 'w', encoding='utf-8') as f:
    for name in xl.sheet_names:
        f.write(name + '\n')
