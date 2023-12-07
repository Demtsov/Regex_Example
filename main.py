import openpyxl
import re


wb = openpyxl.load_workbook('list.xlsx')
sheet = wb.active

for row in sheet.iter_rows(values_only=True):
    for cell in row:
        if type(cell) == str:  # Проверяем, что ячейка содержит строку
            emails = re.findall(r'[\w\.-]+@[\w\.-]+', cell)
            for email in emails:
                print(email)
