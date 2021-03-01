from openpyxl import load_workbook, Workbook

# Loading the full-list file
full_list = load_workbook('W09_sorted.xlsx')['auto']

# Creating the automation-related cases spreadsheet
