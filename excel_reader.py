import openpyxl
import merge_email

def read_excel(File):
    try:
        workbook = openpyxl.load_workbook(filename=File)
    except FileNotFoundError:
        print("File not found")
        return True

    sheet = workbook.active
    max_row = int(sheet.max_row) + 1
    for i in range(1, max_row):
        merge_email.send_email(sheet, i)