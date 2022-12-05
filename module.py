import re
import os
import email
import requests
from imaplib import IMAP4_SSL
from zipfile import ZipFile
from onevizion import LogLevel, IntegrationLog, OVImport


class Module:
    CSV = '.csv'
    ZIP = '.zip'

    def __init__(self, ov_module_log: IntegrationLog, ov_url: str, settings_data: dict) -> None:
        self._module_log = ov_module_log
        self._mail_settings = settings_data['mail']
        self._import = Import(ov_url, settings_data)

    def start(self) -> None:
        self._module_log.add(LogLevel.INFO, 'Starting Module')

        import_id = self._import.get_import_id(self._import.get_import())
        if import_id is None:
            self._module_log.add(LogLevel.INFO, f'Import \"{self._import.import_name}\" not found')
            raise ModuleError(f'Import \"{self._import.import_name}\" not found',
                              'Before running the Module, make sure that there is an import through which the data will be added')

        self._module_log.add(LogLevel.INFO, f'Import \"{self._import.import_name}\" founded')

        with MailService.connect(self._mail_settings['host'], self._mail_settings['port']) as imap_client:
            mail_service = MailService(imap_client)
            mail_service.login(self._mail_settings['username'], self._mail_settings['password'])
            mail_service.select_mailbox()
            unread_messages = mail_service.get_unread_messages_with_subject(self._mail_settings['subject'])
            files_to_import = self._get_files_from_message_to_import(unread_messages)
            self._import_and_remove_file(import_id, files_to_import)

        self._module_log.add(LogLevel.INFO, 'Module has been completed')

    def _get_files_from_message_to_import(self, messages: list) -> list:
        files_to_import = []
        for message in messages:
            for message_part in message.walk():
                file_name = message_part.get_filename()

                if message_part.get_content_maintype() == 'multipart' or \
                    message_part.get('Content-Disposition') is None or \
                    re.search(f'{Module.ZIP}|{Module.CSV}', file_name) is None:
                    continue

                self._safe_file(file_name, message_part)
                if re.search(Module.ZIP, file_name):
                    files_to_import = self._extract_files_from_archive(file_name)
                elif re.search(Module.CSV, file_name):
                    files_to_import.append(file_name)

        return files_to_import

    def _safe_file(self, file_name: str, message_part) -> None:
        path_to_file = os.path.join(file_name)
        if not os.path.isfile(path_to_file):
            with open(path_to_file, 'wb') as file:
                file.write(message_part.get_payload(decode=True))

    def _extract_files_from_archive(self, file_name: str) -> list:
        extracted_files = []
        with ZipFile(file_name) as zip_file:
            for extract_file in zip_file.namelist():
                if re.search(Module.CSV, extract_file):
                    zip_file.extract(extract_file)
                    extracted_files.append(extract_file)

        return extracted_files

    def _import_and_remove_file(self, import_id: int, files: list) -> None:
        for file_name in files:
            try:
                self._import.start_import(import_id, file_name)
                self._module_log.add(LogLevel.INFO, f'Import \"{self._import.import_name}\" has been running for file "{file_name}"')
            finally:
                try:
                    os.remove(file_name)
                    self._module_log.add(LogLevel.DEBUG, f'File "{file_name}" has been deleted')
                except FileNotFoundError:
                    pass


class Import:

    def __init__(self, ov_url: str, settings_data: dict) -> None:
        self._ov_url = ov_url
        self._ov_access_key = settings_data['ovAccessKey']
        self._ov_secret_key = settings_data['ovSecretKey']
        self.import_name = settings_data['importName']
        self._import_action = settings_data['importAction']

    def get_import(self) -> list:
        url = f'https://{self._ov_url}/api/v3/imports'
        headers = {'Content-type':'application/json', 'Content-Encoding':'utf-8', 'Authorization':f'Bearer {self._ov_access_key}:{self._ov_secret_key}'}
        response = requests.get(url=url, headers=headers)

        if response.ok is False:
            raise ModuleError('Failed to get import',  response.text)

        return response.json()

    def get_import_id(self, import_data: list) -> int:
        import_id = None
        for imp_data in import_data:
            if imp_data['name'] == self.import_name:
                import_id = imp_data['id']
                break

        return import_id

    def start_import(self, import_id: int, file_name: str) -> None:
        ov_import = OVImport(self._ov_url, self._ov_access_key, self._ov_secret_key, import_id, file_name, self._import_action, isTokenAuth=True)
        if len(ov_import.errors) != 0:
            raise ModuleError('Failed to start import', ov_import.request.text)


class MailService:
    SUBJECT_PART = 'Subject'
    UNREAD_MSG = 'UnSeen'
    INBOX_FOLDER = 'INBOX'
    MESSAGE_FORMAT = '(RFC822)'

    def __init__(self, imap_client: IMAP4_SSL) -> None:
        self._imap_client = imap_client

    @staticmethod
    def connect(host: str, port: int) -> IMAP4_SSL:
        try:
            return IMAP4_SSL(host, port)
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

    def login(self, user: str, password: str) -> None:
        try:
            self._imap_client.login(user, password)
        except Exception as exception:
            raise ModuleError('Failed to login', exception) from exception

    def select_mailbox(self, mailbox: str = INBOX_FOLDER) -> None:
        try:
            self._imap_client.select(mailbox)
        except Exception as exception:
            raise ModuleError('Failed to select mailbox', exception) from exception

    def get_unread_mails(self) -> tuple:
        try:
            return self._imap_client.search(None,  MailService.UNREAD_MSG)
        except Exception as exception:
            raise ModuleError('Failed to get unread mails', exception) from exception

    def _get_message(self, mail_data: str) -> tuple:
        try:
            return self._imap_client.fetch(mail_data, MailService.MESSAGE_FORMAT)
        except Exception as exception:
            raise ModuleError('Failed to get message', exception) from exception

    def get_unread_messages_with_subject(self, subject: str) -> list:
        unread_messages = []
        unread_mails = self.get_unread_mails()[1]
        for unread_mail in unread_mails[0].split():
            message_data = self._get_message(unread_mail)[1]
            message = email.message_from_bytes(message_data[0][1])
            unread_mail_subject = message.get(MailService.SUBJECT_PART)
            if re.search(subject, unread_mail_subject) is None:
                continue

            unread_messages.append(message)

        return unread_messages


class ModuleError(Exception):

    def __init__(self, error_message: str, description) -> None:
        self._message = error_message
        self._description = description

    @property
    def message(self) -> str:
        return self._message

    @property
    def description(self) -> str:
        return self._description
