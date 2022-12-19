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

    def __init__(self, process_id: int, module_name: str, ov_module_log: IntegrationLog, settings_data: dict) -> None:
        self._module_log = ov_module_log
        self._email_settings = settings_data['email']
        self._import = Import(settings_data['onevizion'], str(process_id), module_name)
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

        with MailService.connect(self._email_settings['host'], self._email_settings['port'],
                                 self._email_settings['username'], self._email_settings['password']) as imap_client:
            mail_service = MailService(imap_client)
            unread_messages = mail_service.get_unread_messages_by_subject(self._email_settings['subject'])

        attachments = mail_service.get_attachments(unread_messages)
        archive_files = list(filter(re.compile(Module.ZIP).search, attachments))
        try:
            extracted_files = Module._extract_files_from_archive(archive_files)
            attachments = Module._remome_archive_files_from_list(archive_files, attachments)
            attachments.extend(extracted_files)

            self._import_files(import_id, attachments)
        finally:
            self._remove_files(archive_files)
            self._remove_files(attachments)

        self._module_log.add(LogLevel.INFO, 'Module has been completed')

    @staticmethod
    def _remome_archive_files_from_list(archive_files: list, attachments: list) -> list:
        for file_name in archive_files:
            attachments.remove(file_name)

        return attachments

    @staticmethod
    def _extract_files_from_archive(archive_files: list) -> list:
        extracted_files = []
        for file_name in archive_files:
            with ZipFile(file_name) as zip_file:
                for extract_file in zip_file.namelist():
                    if re.search(Module.CSV, extract_file):
                        zip_file.extract(extract_file)
                        extracted_files.append(extract_file)

        return extracted_files

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
    RUN_COMMENT = 'Imported started by module [module_name]. Module Run Process ID: [process_id]'

    def __init__(self, onevizion_data: dict, process_id: str, module_name: str) -> None:
        self._ov_url = onevizion_data['ovUrl']
        self._ov_access_key = onevizion_data['ovAccessKey']
        self._ov_secret_key = onevizion_data['ovSecretKey']
        self._run_comment = Import.RUN_COMMENT.replace('module_name', module_name) \
                                              .replace('process_id', process_id)

    def get_import_id(self, import_name: str) -> int:
        import_id = None
        for imp_data in self._get_all_imports():
            if imp_data['name'] == import_name:
                import_id = imp_data['id']
                break

        return import_id

    def _get_all_imports(self) -> list:
        url = f'{self._ov_url}/api/v3/imports'
        headers = {'Content-type': 'application/json', 'Content-Encoding': 'utf-8', 'Authorization': f'Bearer {self._ov_access_key}:{self._ov_secret_key}'}
        response = requests.get(url=url, headers=headers)

        if response.ok is False:
            raise ModuleError('Failed to get import',  response.text)

        return response.json()

    def start_import(self, import_id: int, import_action: str, file_name: str) -> str:
        ov_import = OVImport(self._ov_url, self._ov_access_key, self._ov_secret_key, import_id, file_name, import_action, self._run_comment, isTokenAuth=True)
        if len(ov_import.errors) != 0:
            raise ModuleError('Failed to start import', ov_import.request.text)

        return ov_import.processId


class MailService:
    SUBJECT_PART = 'Subject'
    UNREAD_MSG = 'UnSeen'
    INBOX_FOLDER = 'INBOX'
    MESSAGE_FORMAT = '(RFC822)'

    def __init__(self, imap_client: IMAP4_SSL) -> None:
        self._imap_client = imap_client

    @staticmethod
    def connect(host: str, port: str, user: str, password: str) -> IMAP4_SSL:
        try:
            imap_client = IMAP4_SSL(host, port)
            imap_client.login(user, password)
            return imap_client
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

    def get_unread_messages_by_subject(self, subject: str) -> list:
        unread_messages = []
        unread_mails = self._get_unread_messages()[1]
        for unread_mail in unread_mails[0].split():
            message_data = self._get_message(unread_mail)[1]
            message = email.message_from_bytes(message_data[0][1])
            unread_mail_subject = message.get(MailService.SUBJECT_PART)
            if re.search(subject, unread_mail_subject) is not None:
                unread_messages.append(message)

        return unread_messages

    def _get_unread_messages(self, mailbox: str = INBOX_FOLDER) -> tuple:
        try:
            self._select_mailbox(mailbox)
            return self._imap_client.search(None, MailService.UNREAD_MSG)
        except Exception as exception:
            raise ModuleError('Failed to get unread messages', exception) from exception

    def _select_mailbox(self, mailbox: str) -> None:
        try:
            self._imap_client.select(mailbox)
        except Exception as exception:
            raise ModuleError('Failed to select mailbox', exception) from exception

    def _get_message(self, mail_data: str) -> tuple:
        try:
            return self._imap_client.fetch(mail_data, MailService.MESSAGE_FORMAT)
        except Exception as exception:
            raise ModuleError('Failed to get message', exception) from exception

    @staticmethod
    def get_attachments(messages: list) -> list:
        attachments = []
        for message in messages:
            for message_part in message.walk():
                file_name = message_part.get_filename()

                if message_part.get_content_maintype() != 'multipart' and \
                   message_part.get('Content-Disposition') is not None and \
                   re.search(f'{Module.ZIP}|{Module.CSV}', file_name) is not None:
                    MailService._save_file(file_name, message_part)
                    attachments.append(file_name)

        return attachments

    @staticmethod
    def _save_file(file_name: str, message_part) -> None:
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
