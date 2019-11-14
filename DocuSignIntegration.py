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
    def __init__(self, url_onevizion='', login_onevizion='', pass_onevizion='', import_name='', login_mail='', pass_mail=''):
        self.url_onevizion = self.url_setting(url_onevizion)
        self.import_name = import_name
        self.login_mail = login_mail
        self.pass_mail = pass_mail

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
                if re.search('DocuSign', subject) is not None:
                    for part in msg.walk():                    
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        file_name = part.get_filename()
                        if re.search('.zip', file_name) is not None:
                            att_path = os.path.join(file_name)
                            if not os.path.isfile(att_path):
                                fp = open(att_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()

                            zip_file = zipfile.ZipFile(file_name)
 
                            for extract_file in zip_file.namelist():
                                if re.search('.csv', extract_file):
                                    zip_file.extract(extract_file)
                                    self.create_import_file(extract_file)
                                    self.start_import(extract_file)
                                    os.remove(extract_file)
                            os.remove(file_name)

        else: self.message('Failed to retreive emails')

        self.message('Finished integration')                

    def create_import_file(self, file_name):
        with open(file_name, "r") as file:
            reader = csv.reader(file)

            envelope_list = []
            for row in reader:
                field_names = row
                break

            for row in reader:
                result = []
                for i in reversed(range(len(row[1:]))):
                    result.insert(0, row[i+1])

                if re.search('Please DocuSign:', row[0]):
                    row_split = re.split('/', re.split(':',row[0])[1])
                    if re.search(' USCC|USCC| USC|USC', row_split[0]):
                        row_split = re.split(' USCC|USCC| USC|USC', row_split[0], 2)
                        row_sub = re.sub(r'^\s+|\n|[\[\]]|\r|\s+$', '', row_split[1])
                        
                        result.insert(0, row_sub)
                        inner_dict = dict(zip(field_names, result))
                        envelope_list.append(inner_dict)
                    else:                        
                        row_sub = re.sub(r'^\s+|\n|[\[\]]|\r|\s+$', '', row_split[0])

                        result.insert(0, row_sub)
                        inner_dict = dict(zip(field_names, result))
                        envelope_list.append(inner_dict)
                else: 
                    row_split = re.split('/', row[0])
                    row_sub = re.sub(r'^\s+|\n|\r|\s+$', '', row_split[0])

                    result.insert(0, row_sub)
                    inner_dict = dict(zip(field_names, result))
                    envelope_list.append(inner_dict)

        with open(file_name, "w") as file:
            writer = csv.DictWriter(file, delimiter=',', fieldnames=field_names)
            writer.writeheader()
            for row in envelope_list:
                writer.writerow(row)

    def start_import(self, file_name):
        import_id = self.get_import()
        if import_id != None:
            url = 'https://' + self.url_onevizion + '/api/v3/imports/' + str(import_id) + '/run'
            data = {'action':'INSERT_UPDATE'}
            files = {'file': (file_name, open(file_name, 'rb'))}
            requests.post(url, files=files, params=data, headers={'Accept':'application/json'}, auth=self.auth_onevizion)
            self.message('Import \"' + self.import_name + '\" started')
        else: self.message('Import \"' + self.import_name + '\" not found')

    def get_import(self):
        url = 'https://' + self.url_onevizion + '/api/v3/imports'
        answer = requests.get(url, headers=self.headers, auth=self.auth_onevizion)
        response = answer.json()

        import_id = None
        for imports in response:
            import_name = imports['name']
            if import_name == self.import_name:
                import_id = imports['id']
                return import_id

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
