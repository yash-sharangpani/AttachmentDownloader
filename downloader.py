import imaplib
import os
import email as emp
import patoolib

email_user = input('Email: ')
email_pass = input('Password: ')


basePath = 'C:\\Users\\Manas\\Desktop\\downloadAttachment\\'

mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)

mail.login(email_user, email_pass)
mail.select("inbox")


eMailID = ["abc@gmail.com", "xyz@gmail.com"]
Dict = {"abc@gmail.com": 'abc', "xyz@gmail.com": 'xyz'}


for email in eMailID:
    result, data = mail.search(None, '(UNSEEN)', '(FROM "{emailID}")'.format(emailID=email))
    if result != 'OK':
        print("Authentication failed")
        exit(1)


    mail_ids = data[0]
    id_list = mail_ids.split()
    print("Downloading attachments")
    for num in data[0].split():
        typ, data = mail.fetch(num, '(RFC822)')
        raw_email = data[0][1]
        # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = emp.message_from_string(raw_email_string)
        # downloading attachments
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            if bool(fileName):
                subFolder = Dict[email]
                path = basePath + subFolder
                if not os.path.isdir(basePath):
                    print('Folder created: ' + basePath)
                    os.mkdir(basePath)
                if not os.path.isdir(path):
                    print('Sub folder created: ' + subFolder)
                    os.mkdir(path)
                filePath = os.path.join(path, fileName)

                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

                subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
                print('Downloaded "{file}" from email titled "{subject} and saving to {subFolder}"'.format(file=fileName, subject=subject, subFolder=Dict[email]))
                fileExtension = str(fileName).split(".")[-1]
                print(fileExtension)
                if fileExtension == 'rar' or fileExtension == 'zip':
                    print("Decompressing file: " + fileName)
                    patoolib.extract_archive(filePath, outdir=path)

