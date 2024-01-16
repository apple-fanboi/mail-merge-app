import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Load the Excel file
workbook = openpyxl.load_workbook('C:\\PATH\\TO\\THE\\EXCEL\\FILE\\email_list.xlsx')
sheet = workbook.active

# Set up the email details
sender_email = 'email address'
sender_password = 'App Password'
subject = 'Subject Line'
message = '''
.
.
.
.
.
Messaage
.
.
.
.
'''

# Connect to the email server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, sender_password)

for row in sheet.iter_rows(min_row=2, values_only=True):
    recipient_email = row[0]

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    
    # Attach resume (Word file)
    resume_path = 'C:\\PATH\\TO\\THE\\FILE\\filename.docx'
    with open(resume_path, 'rb') as resume_file:
        attach = MIMEApplication(resume_file.read(), _subtype="docx")
        attach.add_header('Content-Disposition', 'attachment', filename='filename.docx')
        msg.attach(attach)

    # Send the email
    server.sendmail(sender_email, recipient_email, msg.as_string())
    print(f"Email sent to {recipient_email}")

#Print after all emails are sent
print("""
🚀📧 All Systems Go!
All Emails Sent Successfully! 📧🚀

Congratulations! Your messages have been delivered to their destinations.
If you have any questions or need further assistance, feel free to reach out.

Thank you for using our email service! ✨
""")


# Disconnect from the server
server.quit()
