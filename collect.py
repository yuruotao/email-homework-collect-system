# import the necessary dependencies
import email
import imaplib
import json
from pathlib import Path
import smtplib
from base64 import b64decode
from email.header import Header
from email.mime.text import MIMEText
import pandas as pd
import os
from email import utils

# create folders for later convenience
folder_path = str(Path.cwd())
name_list_folder = folder_path + "\\name_list\\"
if not os.path.exists(name_list_folder):
    os.makedirs(name_list_folder)
result_folder = folder_path + '\\result'
if not os.path.exists(result_folder):
    os.makedirs(result_folder)


# workid: the specific homework ID
# reply_check: whether to send reply
download_folder = result_folder
workid = input("Today's Work's ID:")
reply_check = int(input("to send reply, enter 1:"))


# Get student list
students = pd.read_excel(name_list_folder + 'student_info.xlsx')
students['UPLOAD_STATUS'] = students.apply(lambda _: '', axis=1)
students['EMAIL_ADDRESS'] = students.apply(lambda _: '', axis=1)
students['TIME'] = students.apply(lambda _: '', axis=1)


# basic web setup using the information in file config.json
# PASSWORD: the password to your email
# MAIL: mail address
# IMAPSERVER: incoming mail server, in accordance with your email host name
# SMTPSERVER: simple mail transfer protocol server, in accordance with your email host name
# SIGN: used in reply emails, specifying the email sender

PASSWORD, MAIL, IMAPSERVER, SMTPSERVER, SIGN = None, None, None, None, None,
with open("config.json", "r", encoding="utf-8") as config_file:
    config = json.loads(config_file.read())
    PASSWORD = config["PASSWORD"]
    MAIL = config["MAIL"]
    IMAPSERVER = config["IMAPSERVER"]
    SMTPSERVER = config["SMTPSERVER"]
    SIGN = config["SIGN"]
print("MAIL: ", MAIL)
print("IMAPSERVER: ", IMAPSERVER)
print("SMTPSERVER: ", SMTPSERVER)

# mail_obj represents the web-connection object
# smtp_obj represents an SMTP client session object that can be used to send mail
# 993 specifies the port (IMAP over SSL), 
# reference: https://en.wikipedia.org/wiki/Internet_Message_Access_Protocol
mail_obj = imaplib.IMAP4_SSL(IMAPSERVER, "993")

# for 163.com: command SEARCH illegal in state AUTH problem
# ref: http://help.mail.163.com/faqDetail.do?code=d7a5dc8471cd0c0e8b4b8f4f8e49998b374173cfe9171305fa1ce630d7f67ac211b1978002df8b23
if "@163" in MAIL:
    imaplib.Commands['ID'] = 'AUTH'
smtp_obj = smtplib.SMTP()

# log in the mail with mail_obj and smtp_obj
try:
    mail_obj.login(MAIL, PASSWORD)
    smtp_obj.connect(SMTPSERVER, 25)
    smtp_obj.login(MAIL, PASSWORD)

    args = ("name", "gongchuang201", "contact", "gongchuang201@163.com", "version", "1.0.0", "vendor", "myclient")
    typ, dat = mail_obj._simple_command('ID', '("' + '" "'.join(args) + '")')
except Exception as e:
    print('Error: %s' % str(e.args[0], "utf-8"))
    exit()
print("CONNECTED!")

# fetch all the emails
print("Fetching mail list")
mail_obj.select()
typ, received_data = mail_obj.search(None, 'ALL')

# the form of email title should be STUDENT_ID/HOMEWORK_ID
# the received emails are sorted in a reversed order, named as emails
emails = received_data[0].split()[::-1]


# collecting dependencies based on the structure of the email
for i in emails:
    print("Fetching mail %d" % int(i))
    # Internet RFC 822 specification defines an electronic message format consisting of header fields and an optional message body
    typ, received_data = mail_obj.fetch(i, '(RFC822)')

    # UTF-8 decoding may cause problem since not containing Chinese characters
    email_message = email.message_from_string(received_data[0][1].decode('utf-8'))
    email_subtitle = email.header.decode_header(email_message.get('subject'))[0][0]
    if type(email_subtitle) == bytes:
        email_subtitle = email_subtitle.decode('utf-8')

    # examine by the header
    email_headers = email_subtitle.split("/")
    if len(email_headers) == 2:
        if email_headers[1] != workid:
            continue
    else:
        continue
    pass

    # document the subtitle, time, and email address with student ID
    STUDENT_ID = email_headers[0]
    student_index = students[students['STUDENT_ID'] == STUDENT_ID].index
    print("[SUBJECT] ", email_subtitle)
    mailfrom = email.utils.parseaddr(email_message.get('from'))[1]
    students._set_value(student_index, 'EMAIL_ADDRESS', mailfrom)
    print("[   FROM] ", mailfrom)

    # time stamp
    raw = email.message_from_bytes(received_data[0][1])
    datestring = raw['date']
    # Convert to datetime object
    datetime_obj = utils.parsedate_to_datetime(datestring)
    print("[   TIME] ", repr(datetime_obj))
    students._set_value(student_index, 'TIME', repr(datetime_obj))


    # saving files to local and document the upload
    if email_message.get_content_maintype() == 'multipart':
        # loop on the parts of the mail
        for part in email_message.walk():
            # find the attachment part and skip all the other parts
            if part.get_content_maintype() == 'multipart': continue
            if part.get_content_maintype() == 'text': continue
            if part.get('Content-Disposition') == 'inline': continue
            if part.get('Content-Disposition') is None: continue
            name = part.get_filename()
            break
        # fix filename
        if name[0:8] == "=?utf-8?":
            name = name[10:-2]
            name = b64decode(name).decode(encoding="utf-8")
        if name:
            print("[ ATTACH] ", name)
            attach_data = part.get_payload(decode=True)
            # the filename here indicates the name saved in the folder 'result'
            filename = STUDENT_ID + students.at[students.index[student_index], 'NAME']

            # saving attachments to the path
            att_path = os.path.join(download_folder, filename)
            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(attach_data)
                fp.close()
        else:
            pass
    # set the upload status of a student
    students._set_value(student_index, 'UPLOAD_STATUS', 'yes')


    # send the confirmation of submission if the parameter reply_check is 1
    if reply_check == 1:
        mail_msg = "<p>" + \
                   students[STUDENT_ID][0] + \
                   "<br>your homework " + workid + \
                   "is accepted. <br>Thank you!</p><hr>" + SIGN
        message = email.mime.text.MIMEText(mail_msg, 'html', 'utf-8')
        message['Subject'] = Header('Re:' + email_subtitle, 'utf-8')
        message['From'] = Header(MAIL, 'utf-8')
        message['To'] = Header(students[STUDENT_ID][0], 'utf-8')
        try:
            smtp_obj.sendmail(MAIL, [mailfrom], message.as_string())
            print("Reply Sent to ", )
        except smtplib.SMTPException:
            print("Error: Reply not sent")


# log out the email account
mail_obj.logout()
