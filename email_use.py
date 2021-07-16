import smtplib, email.encoders, email.mime.text, email.mime.base
from email.mime.multipart import MIMEMultipart

monthName = input('Please enter your chosen month (E.g Nov): ')

# Email manager the department list.
smtpserver = 'SERVER.DOMAIN.local'
to = ['RECEIVER_EMAIL']
fromAddr = 'SENDER_EMAIL'
subject = monthName + " Project References allocated to wrong departments"
cc = ''

# Create HTML email
html = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
html += '"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
html += '<body style="font-size:12px;font-family:Verdana"><p>...</p>'
html += "</body></html>"
emailMsg = MIMEMultipart('alternative')
emailMsg['Subject'] = subject
emailMsg['From'] = fromAddr
emailMsg['To'] = ', '.join(to)
emailMsg['Cc'] = ", ".join(cc)
emailMsg.attach(email.mime.text.MIMEText(html, 'html'))

# Attach the file
fileMsg = email.mime.base.MIMEBase('application', 'vnd.ms-excel')
fileMsg.set_payload(open('\\\\SERVER\\DIRECTORY\\' + monthName + ' department.txt', "rb").read())
email.encoders.encode_base64(fileMsg)
fileMsg.add_header('Content-Disposition', 'attachment;filename=' + monthName + ' department.txt')
emailMsg.attach(fileMsg)

# Send email
server = smtplib.SMTP(smtpserver)
server.sendmail(fromAddr, to, emailMsg.as_string())
server.quit()
