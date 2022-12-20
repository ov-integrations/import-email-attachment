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

        attachments = MailService.get_attachments_from_unread_messages(self._email_settings)
        csv_files = list(filter(re.compile(Module.CSV).search, attachments))
        zip_files = list(filter(re.compile(Module.ZIP).search, attachments))
        extracted_csv_files = Module._extract_csv_files_from_archive(zip_files)
        csv_files.extend(extracted_csv_files)
        self._module_log.add(LogLevel.INFO, f'{len(csv_files)} CSV files found')

        try:
            self._import_files(import_id, csv_files)
        finally:
            self._remove_files(zip_files)
            self._remove_files(csv_files)

        self._module_log.add(LogLevel.INFO, 'Module has been completed')

    @staticmethod
    def _extract_csv_files_from_archive(zip_files: list) -> list:
        extracted_csv_files = []
        for file_name in zip_files:
            with ZipFile(file_name) as zip_file:
                for extract_file in zip_file.namelist():
                    if re.search(Module.CSV, extract_file):
                        zip_file.extract(extract_file)
                        extracted_csv_files.append(extract_file)

        return extracted_csv_files

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

    @staticmethod
    def get_attachments_from_unread_messages(email_settings: dict) -> list:
        attachments = []
        unread_messages = MailService._get_unread_messages_by_subject(email_settings)
        for message in unread_messages:
            for message_part in message.walk():
                file_name = message_part.get_filename()

                if message_part.get_content_maintype() != 'multipart' and \
                   message_part.get('Content-Disposition') is not None and \
                   re.search(f'{Module.ZIP}|{Module.CSV}', file_name) is not None:
                    MailService._save_file(file_name, message_part)
                    attachments.append(file_name)

        return attachments

    @staticmethod
    def _get_unread_messages_by_subject(email_settings: dict) -> list:
        imap_client = MailService._connect(email_settings['host'], email_settings['port'],
                                          email_settings['username'], email_settings['password'])
        filtered_unread_messages = []
        unread_messages = MailService._get_unread_messages(imap_client)[1]
        for unread_message in unread_messages[0].split():
            message_data = MailService._get_message(imap_client, unread_message)[1]
            message = email.message_from_bytes(message_data[0][1])
            unread_mail_subject = message.get(MailService.SUBJECT_PART)
            if re.search(email_settings['subject'], unread_mail_subject) is not None:
                filtered_unread_messages.append(message)

        MailService._disconnect(imap_client)
        return filtered_unread_messages

    @staticmethod
    def _connect(host: str, port: str, user: str, password: str) -> IMAP4_SSL:
        try:
            imap_client = IMAP4_SSL(host, port)
            imap_client.login(user, password)
            return imap_client
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

    @staticmethod
    def _get_unread_messages(imap_client: IMAP4_SSL, mailbox: str = INBOX_FOLDER) -> tuple:
        try:
            MailService._select_mailbox(imap_client, mailbox)
            return imap_client.search(None, MailService.UNREAD_MSG)
        except Exception as exception:
            raise ModuleError('Failed to get unread messages', exception) from exception

    @staticmethod
    def _select_mailbox(imap_client: IMAP4_SSL, mailbox: str) -> None:
        try:
            imap_client.select(mailbox)
        except Exception as exception:
            raise ModuleError('Failed to select mailbox', exception) from exception

    @staticmethod
    def _get_message(imap_client: IMAP4_SSL, mail_data: str) -> tuple:
        try:
            return imap_client.fetch(mail_data, MailService.MESSAGE_FORMAT)
        except Exception as exception:
            raise ModuleError('Failed to get message', exception) from exception

    @staticmethod
    def _disconnect(imap_client: IMAP4_SSL) -> None:
        try:
            imap_client.close()
            imap_client.logout()
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

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
