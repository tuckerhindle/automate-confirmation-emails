from openpyxl import load_workbook
from datetime import datetime

from outlook_utils import get_cell_value, write_email_content, send_new_email

# User Inputs
workbook = "C:\\Users\\thindle2016\\Florida Atlantic University\\FAU Lab Schools Research Department - Fall_2024_Registration_Materials\\Fall 2024 Research Courses Form.xlsx"
term = "Fall 2024"
outlook_account = "fauhsresearch@fau.edu"

# Specify column index
names = 8
emails = 9
courses = 12
EDF2910_sections = 13
EDF3912_sections = 14
status_col = 16
last_updated = 17

wb = load_workbook(workbook)
ws = wb.active

row_count = ws.max_row

for row in range(2, row_count + 1):

    status = get_cell_value(ws, row, status_col)

    if status == "Approved":
        name = get_cell_value(ws, row, names)
        recipient = get_cell_value(ws, row, emails)
        course = get_cell_value(ws, row, courses)

        EDF2910 = get_cell_value(ws, row, EDF2910_sections) # course 1
        EDF3912 = get_cell_value(ws, row, EDF3912_sections) # course 2
        section = EDF3912 if EDF2910 is None else EDF2910
        
        subject, body = write_email_content(name, term, course, section)

        send_new_email(outlook_account, recipient, subject, body)

        new_status = "Email Sent"
        ws.cell(row=row, column=status_col).value = new_status
        ws.cell(row=row, column=last_updated).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

wb.save(workbook)
