import data


def setEmail(Email, Password):
    '''
    Sets email and password
    '''
    data.Email = Email
    data.Password = Password

def setExcelFileName(Excel_Name):
    '''
    Sets the ExcelFileName
    '''
    data.Excel_file = Excel_Name

def setMessage(msg):
    '''
    Sets the message for the email
    '''
    data.Message = msg

def setSubject(sub):
    '''
    Sets the subject for the email
    '''
    data.Subject = sub