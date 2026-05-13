import openpyxl
import os

file_path = os.path.join(os.path.dirname(__file__), '..', '2026長庚盃 排名、個人團體對抗表(2026新版).xlsx')
wb = openpyxl.load_workbook(file_path, read_only=True)

print("Sheet names in the workbook:")
for i, name in enumerate(wb.sheetnames):
    print(f"{i}: {name}")
