import excel_reader

if __name__ == "__main__":
    global Excel_file 
    success = True
    while success == True: 
        Excel_file = input("Enter path of Excel File:\n")
        success = excel_reader.read_excel(f"{Excel_file}.xlsx")

    
