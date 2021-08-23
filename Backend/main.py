import excel_reader
import data
import setData

if __name__ == "__main__":
    success = True
    while success == True:
        email = input("Enter your email:")
        password = input("Enter your password:") 
        setData.setEmail(email, password)
        data.Excel_name = input("Enter name of Excel File:\n")
        success = excel_reader.read_excel(f"{data.Excel_name}.xlsx")

    
