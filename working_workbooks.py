import openpyxl

# open subjects_data excel file
subjects_book = openpyxl.load_workbook("subjetcs-data.xlsx")

subjects_sheet = subjects_book["students"]

subjects_sheet["B7"].value = 0
subjects_sheet["C7"] = 0

subjects_book.save("subjetcs-data.xlsx")
subjects_book.close
