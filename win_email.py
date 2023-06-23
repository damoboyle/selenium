import win32com.client as win32

def email(address, subject, message):
    #Open Outlook and compose an email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    mail.To = address
    mail.Subject = subject
    
    #Pulls up default signature of sender from Outlook
    mail.GetInspector
    
    #HTML text to inform receiver that this email is automated
    intro = "<i style='color:gray'>This is an automated message.</i><br><br>"
    
    #Used to merge intro and user message into HTML body properly while including default signature
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + intro + message + mail.HTMLbody[index + 1:]
    
    #Send Email
    #mail.Display(True)  #Displays Email for Testing
    mail.Send()