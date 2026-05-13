import openpyxl
import os
import json

file_path = os.path.join(os.path.dirname(__file__), '..', '2026長庚盃 排名、個人團體對抗表(2026新版).xlsx')
wb = openpyxl.load_workbook(file_path, data_only=True)

with open(os.path.join(os.path.dirname(__file__), 'sheets.json'), 'w', encoding='utf-8') as f:
    json.dump(wb.sheetnames, f, ensure_ascii=False, indent=2)
