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

    def __init__(self, ov_module_log: IntegrationLog, settings_data: dict) -> None:
        self._module_log = ov_module_log
        self._email_settings = settings_data['email']
        self._import = Import(settings_data['onevizion'])
        self._mail_service = MailService(settings_data['email'])
        self._import_name = settings_data['onevizion']['ovImportName']
        self._import_action = settings_data['onevizion']['ovImportAction']

    def start(self) -> None:
        self._module_log.add(LogLevel.INFO, 'Starting Module')

        import_id = self._import.get_import_id(self._import_name)
        if import_id is None:
            self._module_log.add(LogLevel.WARNING, f'Import \"{self._import_name}\" not found')
            raise ModuleError(f'Import \"{self._import_name}\" not found',
                              'Before running the Module, make sure that there is an import that will be used to add the data')

        self._module_log.add(LogLevel.INFO, f'Import \"{self._import_name}\" founded')

        with self._mail_service.connect() as mail_service:
            unread_messages = self._mail_service.get_unread_messages_by_subject(mail_service)

        attachments = self._mail_service.get_attachments(unread_messages)
        archive_files = list(filter(re.compile(Module.ZIP).search, attachments))
        try:
            attachments = self._extract_files_from_archive(archive_files, attachments)
        finally:
            self._remove_files(archive_files)

        try:
            self._import_files(import_id, attachments)
        finally:
            self._remove_files(attachments)

        self._module_log.add(LogLevel.INFO, 'Module has been completed')

    def _extract_files_from_archive(self, archive_files: list, attachments: list) -> list:
        for file_name in archive_files:
            attachments.remove(file_name)
            with ZipFile(file_name) as zip_file:
                for extract_file in zip_file.namelist():
                    if re.search(Module.CSV, extract_file):
                        zip_file.extract(extract_file)
                        attachments.append(extract_file)

        return attachments

    def _import_files(self, import_id: int, files: list) -> None:
        for file_name in files:
            process_id = self._import.start_import(import_id, self._import_action, file_name)
            self._module_log.add(LogLevel.INFO, f'Import \"{self._import_name}\" has been running for file "{file_name}". Process ID [{process_id}]')

    def _remove_files(self, files: list) -> None:
        for file_name in files:
            try:
                os.remove(file_name)
                self._module_log.add(LogLevel.DEBUG, f'File "{file_name}" has been deleted')
            except FileNotFoundError:
                pass


class Import:
    RUN_COMMENT = 'Imported started by module [import-email-attachment]'

    def __init__(self, onevizion_data: dict) -> None:
        self._ov_url = onevizion_data['ovUrl']
        self._ov_access_key = onevizion_data['ovAccessKey']
        self._ov_secret_key = onevizion_data['ovSecretKey']

    def get_import_id(self, import_name: str) -> int:
        import_id = None
        for imp_data in self._get_import():
            if imp_data['name'] == import_name:
                import_id = imp_data['id']
                break

        return import_id

    def _get_import(self) -> list:
        url = f'{self._ov_url}/api/v3/imports'
        headers = {'Content-type': 'application/json', 'Content-Encoding': 'utf-8', 'Authorization': f'Bearer {self._ov_access_key}:{self._ov_secret_key}'}
        response = requests.get(url=url, headers=headers)

        if response.ok is False:
            raise ModuleError('Failed to get import',  response.text)

        return response.json()

    def start_import(self, import_id: int, import_action: str, file_name: str) -> str:
        ov_import = OVImport(self._ov_url, self._ov_access_key, self._ov_secret_key, import_id, file_name, import_action, Import.RUN_COMMENT, isTokenAuth=True)
        if len(ov_import.errors) != 0:
            raise ModuleError('Failed to start import', ov_import.request.text)

        return ov_import.processId


class MailService:
    SUBJECT_PART = 'Subject'
    UNREAD_MSG = 'UnSeen'
    INBOX_FOLDER = 'INBOX'
    MESSAGE_FORMAT = '(RFC822)'

    def __init__(self, mail_data: dict) -> None:
        self._host = mail_data['host']
        self._port = mail_data['port']
        self._user = mail_data['username']
        self._password = mail_data['password']
        self._subject = mail_data['subject']

    def connect(self) -> IMAP4_SSL:
        try:
            imap_client = IMAP4_SSL(self._host, self._port)
            imap_client.login(self._user, self._password)
            return imap_client
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

    def get_unread_messages_by_subject(self, mail_service: IMAP4_SSL) -> list:
        unread_messages = []
        unread_mails = self._get_unread_messages(mail_service)[1]
        for unread_mail in unread_mails[0].split():
            message_data = self._get_message(mail_service, unread_mail)[1]
            message = email.message_from_bytes(message_data[0][1])
            unread_mail_subject = message.get(MailService.SUBJECT_PART)
            if re.search(self._subject, unread_mail_subject) is not None:
                unread_messages.append(message)

        return unread_messages

    def _get_unread_messages(self, mail_service: IMAP4_SSL, mailbox: str = INBOX_FOLDER) -> tuple:
        try:
            self._select_mailbox(mail_service, mailbox)
            return mail_service.search(None,  MailService.UNREAD_MSG)
        except Exception as exception:
            raise ModuleError('Failed to get unread mails', exception) from exception

    def _select_mailbox(self, mail_service: IMAP4_SSL, mailbox: str) -> None:
        try:
            mail_service.select(mailbox)
        except Exception as exception:
            raise ModuleError('Failed to select mailbox', exception) from exception

    def _get_message(self, mail_service: IMAP4_SSL, mail_data: str) -> tuple:
        try:
            return mail_service.fetch(mail_data, MailService.MESSAGE_FORMAT)
        except Exception as exception:
            raise ModuleError('Failed to get message', exception) from exception

    def get_attachments(self, messages: list) -> list:
        attachments = []
        for message in messages:
            for message_part in message.walk():
                file_name = message_part.get_filename()

                if message_part.get_content_maintype() != 'multipart' and \
                   message_part.get('Content-Disposition') is not None and \
                   re.search(f'{Module.ZIP}|{Module.CSV}', file_name) is not None:
                    self._save_file(file_name, message_part)
                    attachments.append(file_name)

        return attachments

    def _save_file(self, file_name: str, message_part) -> None:
        path_to_file = os.path.join(file_name)
        if not os.path.isfile(path_to_file):
            with open(path_to_file, 'wb') as file:
                file.write(message_part.get_payload(decode=True))


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
