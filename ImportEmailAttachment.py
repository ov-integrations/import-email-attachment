import re
import os
import csv
import json
import email
import imaplib
import zipfile
import requests
import onevizion
from requests.auth import HTTPBasicAuth

class Integration(object):
    def __init__(self, url_onevizion='', login_onevizion='', pass_onevizion='', import_name='', login_mail='', pass_mail='', subject_mail=''):
        self.url_onevizion = self.url_setting(url_onevizion)
        self.import_name = import_name
        self.login_mail = login_mail
        self.pass_mail = pass_mail
        self.subject_mail = subject_mail

        self.headers = {'Content-type':'application/json','Content-Encoding':'utf-8'}
        self.auth_onevizion = HTTPBasicAuth(login_onevizion, pass_onevizion)
        self.message = onevizion.Message

        self.get_unread_messages()

    def get_unread_messages(self):
        self.message('Started integration')

        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        mail.login(self.login_mail, self.pass_mail)
        mail.select(mailbox='INBOX')

        result, messages = mail.search(None, 'UnSeen')
        if result == 'OK':
            for message in messages[0].split():
                ret, data = mail.fetch(message,'(RFC822)')
                msg = email.message_from_bytes(data[0][1])
                subject = msg.get('Subject')
                if re.search(self.subject_mail, subject) is not None:
                    for part in msg.walk():                    
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        file_name = part.get_filename()
                        if re.search('.zip|.csv', file_name) is not None:
                            att_path = os.path.join(file_name)
                            if not os.path.isfile(att_path):
                                fp = open(att_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()

                            if re.search('.zip', file_name):
                                zip_file = zipfile.ZipFile(file_name)
    
                                for extract_file in zip_file.namelist():
                                    if re.search('.csv', extract_file):
                                        zip_file.extract(extract_file)
                                        self.start_import(extract_file)
                                        os.remove(extract_file)

                            elif re.search('.csv', file_name):
                                self.start_import(file_name)
                            os.remove(file_name)

        else: self.message('Failed to retrieve emails')

        self.message('Finished integration')                

    def start_import(self, file_name):
        import_id = self.get_import()
        if import_id != None:
            url = 'https://' + self.url_onevizion + '/api/v3/imports/' + str(import_id) + '/run'
            data = {'action':'INSERT_UPDATE'}
            files = {'file': (file_name, open(file_name, 'rb'))}
            requests.post(url, files=files, params=data, headers={'Accept':'application/json'}, auth=self.auth_onevizion)
            self.message('Import \"' + self.import_name + '\" started')

    def get_import(self):
        url = 'https://' + self.url_onevizion + '/api/v3/imports'
        answer = requests.get(url, headers=self.headers, auth=self.auth_onevizion)
        import_id = None
        if answer.status_code != 200:
            self.message('Problem with getting import - ' + answer.text)
        else:
            response = answer.json()

            for imports in response:
                import_name = imports['name']
                if import_name == self.import_name:
                    import_id = imports['id']
                    break
            
            if import_id == None:
                self.message('Import \"' + self.import_name + '\" not found')

        return import_id

    def url_setting(self, url):
        url_re_start = re.search('^https', url)
        url_re_finish = re.search('/$', url)
        if url_re_start is not None and url_re_finish is not None:
            url_split = re.split('://',url[:-1],2)
            url = url_split[1]  
        elif url_re_start is None and url_re_finish is not None:
            url = url[:-1]
        elif url_re_start is not None and url_re_finish is None:
            url_split = re.split('://',url,2)
            url = url_split[1]
        return url
