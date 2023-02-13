import win32com.client as win32

def send_new_email(subject, recipients, content):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = subject
    mail.To = recipients
    mail.HTMLBody = content

    mail.Send()
