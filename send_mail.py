import win32com.client as win32
import os 
import xlwings as xw




@xw.func # xlwing function to run from excel
def maill(file_name):

    #wb = xw.Book.caller()
    #sheet = wb.sheets.active

    __location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))

    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'mail to'
    mail.Subject = 'subject'
    #mail.Body = 'How are you'
    mail.HTMLBody = '<h2>HELLLOOOOOOOO FROM PYTHON</h2>' #this field is optional

    
    file_path= __location__ + "\\"+ file_name
    mail.Attachments.Add(file_path)

    mail.display(True)
    mail.Send()
    print("sent")
