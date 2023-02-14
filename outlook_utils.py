import win32com.client as win32


def get_cell_value(sheet, row, col):
    """Access cell values using row and column index in openpyxl"""
    return sheet.cell(row=row, column=col).value


def write_email_content(name, term, course, section):
    """Takes a student name and course information, returns a subject line and body of an email."""

    subject = f"Approved to Register for {term} {course}"

    greeting = f"Hello {name},<br><br>"
    message = f"Thank you for completing the {term} Research Courses form.  The instructor approved your request for <b>{course} at {section}</b>.<br><br>Once FAU registration opens, please add this class to your schedule by following these instructions: <a href='https://www.fau.edu/registrar/registration/#Register'>FAU Guide to Register for Classes</a><br><br>"
    signature = f"We are excited to see you this {term}!<br><br>--<br><b>FAU Lab Schools Research Team</b><br>Florida Atlantic University<br>777 Glades Road, Building 26F, Room 109<br>Boca Raton, Florida 33431"

    body = "<h3 style='font-weight:normal;'>" + greeting + message + signature + "</h3>"

    return subject, body


def send_new_email(subject, recipients, body):
    """Sends a new email using MS Outlook to the specified recipients containing the subject line and body provided."""
    
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    mail.Subject = subject
    mail.To = recipients
    mail.HTMLBody = body

    mail.Send()
