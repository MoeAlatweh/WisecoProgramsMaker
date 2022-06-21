# Sending emails without attachments using Python.
# importing the required library.
import smtplib
import keyring

# SERVER NAME
smtp_server = "smtp.office365.com"
port = 587
# PUT THE EMAIL THAT'S USED TO SEND THE EMAIL
sender_email = ""
# USING KEYRING PACKAGE TO GET PASSWORD FROM Windows Credential Manager WHILE THEY ARE SAVING IN USER COMPUTER,
# PUT THE EMAIL THAT'S CONNECT WITH THE COMPUTER AS SECOND PARAMETER
sender_password = keyring.get_password("outlook.office365.com", "")
# sender_password = ""

# PUT THE EMAIL THAT'S NEED TO RECEIVE THE EMAIL
recipient_email = ""

# creates SMTP session
email_Server = smtplib.SMTP(smtp_server, port)

# TLS for security
email_Server.starttls()

print(sender_email)
# print(sender_password)

# authentication
# compiler gives an error for wrong credential.
email_Server.login(sender_email, sender_password)

email_message_list = ["test8"+'\n', "test9"+'\n', "test10"+'\n', "test11"+'\n']

# message to be sent
email_message = f"""From: Alatweh Moemen <*PUT THE SENDER EMAIL*>
To: <*PUT THE RECEIVER EMAIL*>
Subject: Testing Email by python

{"".join(email_message_list)}"""

# print(email_message)

# sending the mail
email_Server.sendmail(sender_email, recipient_email, email_message)

# terminating the session
email_Server.quit()

print("EMAIL HAS BEEN SENT")
print(sender_email)
# print(sender_password)
print(email_message)
