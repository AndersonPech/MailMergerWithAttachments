from main import Excel_file
import smtplib, ssl
import openpyxl

#############################
# Sends email using smtplib
# 
#     
#############################
def send_email(index):
    port = 465  # For SSL
    password = input("Type your password and press enter: ")

    # Create a secure SSL context
    context = ssl.create_default_context()
    message=input("Enter your message")
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login("my@gmail.com", password)

