import openpyxl as xl
wb = xl.load_workbook('transactions.xlsx')

sheet = wb['Sheet1']  # Access the sheet
cell = sheet['a1']  # Give the coordinate
cell = sheet.cell(1,1)
# print(cell.value)  # Get value of cell at specified position
# print(sheet.max_row)  # Get the maximum row filed

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)  # Get access to cell at row and column = 3
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)  # setting the corrected value
    corrected_price_cell.value = corrected_price  # setting corrected_value at corrected_prince_cell

wb.save("transactions.2.xlsx")  # Workbook.save