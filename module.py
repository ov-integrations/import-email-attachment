import re
import os
import email
import requests
from imaplib import IMAP4_SSL
from zipfile import ZipFile
from onevizion import LogLevel, IntegrationLog, OVImport


class Module:
    FILES_TO_IMPORT_FOLDER = 'files_to_import'
    CSV_REGEXP = '(?i)^.*\\.csv$'
    ZIP_REGEXP = '(?i)^.*\\.zip$'

    def __init__(self, process_id: int, module_name: str, ov_module_log: IntegrationLog, settings_data: dict) -> None:
        _ov_settings = settings_data['onevizion']
        self._module_log = ov_module_log
        self._import = Import(_ov_settings['ovUrl'], _ov_settings['ovAccessKey'], _ov_settings['ovSecretKey'],
                              str(process_id), module_name)
        self._mail_service = MailService(settings_data['email'])
        self._import_name = _ov_settings['ovImportName']
        self._import_action = _ov_settings['ovImportAction']

    def start(self) -> None:
        self._module_log.add(LogLevel.INFO, 'Starting Module')

        import_id = self._import.get_import_id(self._import_name)
        if import_id is None:
            self._module_log.add(LogLevel.ERROR, f'Import [{self._import_name}] not found')
            raise ModuleError(f'Import [{self._import_name}] not found',
                              'Before running the Module, make sure that there is an import that will be used to add the data')

        self._module_log.add(LogLevel.INFO, f'Import ID [{import_id}] found for import [{self._import_name}]')

        unread_messages = self._mail_service.get_unread_messages()
        self._module_log.add(LogLevel.INFO, f'{len(unread_messages)} unread messages found')

        for unread_message in unread_messages:
            attachments = self._get_attachments_from_message(unread_message)
            csv_files = self._get_csv_files_to_import(attachments)
            try:
                self._import_files(import_id, csv_files)
            finally:
                self.remove_files()

        self._module_log.add(LogLevel.INFO, 'Module has been completed')

    def _get_attachments_from_message(self, message) -> list:
        attachments = []
        send_from = message.get('From')
        send_date = message.get('Date')
        for message_part in message.walk():
            file_name = message_part.get_filename()
            if self._file_is_contained_in_part_of_message(message_part) and \
               self._is_file_type_supported(file_name):
                path_to_file = self._save_file(file_name, message_part)
                self._module_log.add(LogLevel.DEBUG,
                                        f'Received file [{file_name}] from a message sent on [{send_date}] from [{send_from}]')
                attachments.append(path_to_file)

        return attachments

    def _file_is_contained_in_part_of_message(self, message_part) -> bool:
        file_is_contained_in_part_of_message = False
        if message_part.get('Content-Disposition') and \
           message_part.get_content_maintype() != 'multipart':
            file_is_contained_in_part_of_message = True

        return file_is_contained_in_part_of_message

    def _is_file_type_supported(self, file_name: str) -> bool:
        is_file_type_supported = False
        if re.search(Module.ZIP_REGEXP, file_name) or re.search(Module.CSV_REGEXP, file_name):
            is_file_type_supported = True

        return is_file_type_supported

    def _save_file(self, file_name: str, message_part) -> str:
        path_to_file = os.path.join(Module.FILES_TO_IMPORT_FOLDER, file_name)
        if not os.path.isfile(path_to_file):
            with open(path_to_file, 'wb') as file:
                file.write(message_part.get_payload(decode=True))

        return path_to_file

    def _get_csv_files_to_import(self, attachments: list) -> list:
        csv_files = list(filter(re.compile(Module.CSV_REGEXP).search, attachments))
        zip_files = list(filter(re.compile(Module.ZIP_REGEXP).search, attachments))
        extracted_csv_files = self._get_csv_files_from_zip(zip_files)
        csv_files.extend(extracted_csv_files)
        self._module_log.add(LogLevel.INFO, f'{len(csv_files)} CSV files found')
        return csv_files

    def _get_csv_files_from_zip(self, zip_files: list) -> list:
        extracted_csv_files = []
        for file_name in zip_files:
            with ZipFile(file_name) as zip_file:
                extracted_files = self._extract_from_zip_by_file_extension(zip_file, Module.CSV_REGEXP)
                extracted_csv_files.extend(extracted_files)

        return extracted_csv_files

    def _extract_from_zip_by_file_extension(self, zip_file: ZipFile, file_extension: str) -> list:
        extracted_files = []
        for extract_file in zip_file.namelist():
            if re.search(file_extension, extract_file):
                path_to_file = zip_file.extract(extract_file, Module.FILES_TO_IMPORT_FOLDER)
                extracted_files.append(path_to_file)
                self._module_log.add(LogLevel.DEBUG, f'Extracted file [{extract_file}] from ZIP [{zip_file}]')

        return extracted_files

    def _import_files(self, import_id: int, files: list) -> None:
        for file_name in files:
            process_id = self._import.start_import(import_id, self._import_action, file_name)
            self._module_log.add(LogLevel.DEBUG, f'Started import for [{file_name}]. Process ID [{process_id}]')

    def create_folder_to_save_files(self):
        if not os.path.exists(Module.FILES_TO_IMPORT_FOLDER):
            os.makedirs(Module.FILES_TO_IMPORT_FOLDER)

    def remove_files(self) -> None:
        for file_name in os.listdir(Module.FILES_TO_IMPORT_FOLDER):
            try:
                os.remove(f'{Module.FILES_TO_IMPORT_FOLDER}/{file_name}')
                self._module_log.add(LogLevel.DEBUG, f'File [{file_name}] has been deleted')
            except FileNotFoundError:
                pass
            except Exception as exception:
                self._module_log.add(LogLevel.ERROR, f'File [{file_name}] can\'t been deleted. Error: [{exception}]')


class Import:
    RUN_COMMENT = 'Imported started by module [module_name]. Module Run Process ID: [process_id]'

    def __init__(self, ov_url: str, ov_access_key: str, ov_secret_key: str, process_id: str, module_name: str) -> None:
        self._ov_url = ov_url
        self._ov_url_without_protocol = re.sub('^http://|^https://', '', ov_url[:-1])
        self._ov_access_key = ov_access_key
        self._ov_secret_key = ov_secret_key
        self._run_comment = Import.RUN_COMMENT.replace('module_name', module_name) \
                                              .replace('process_id', process_id)

    def get_import_id(self, import_name: str) -> int:
        import_id = None
        for imp_data in self._get_imports():
            if imp_data['name'] == import_name:
                import_id = imp_data['id']
                break

        return import_id

    def _get_imports(self) -> list:
        url = f'{self._ov_url}/api/v3/imports'
        headers = {'Content-type': 'application/json', 'Content-Encoding': 'utf-8',
                   'Authorization': f'Bearer {self._ov_access_key}:{self._ov_secret_key}'}
        response = requests.get(url=url, headers=headers)

        if response.ok is False:
            raise ModuleError('Failed to get imports',  response.text)

        return response.json()

    def start_import(self, import_id: int, import_action: str, file_name: str) -> str:
        ov_import = OVImport(self._ov_url_without_protocol, self._ov_access_key, self._ov_secret_key,
                             import_id, file_name, import_action, self._run_comment, isTokenAuth=True)
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

    def get_unread_messages(self) -> list:
        with self._connect() as imap_client:
            unread_messages = self._filter_unread_messages_by_subject(imap_client)

        return unread_messages

    def _connect(self) -> IMAP4_SSL:
        try:
            imap_client = IMAP4_SSL(self._host, self._port)
            imap_client.login(self._user, self._password)
            return imap_client
        except Exception as exception:
            raise ModuleError('Failed to connect', exception) from exception

    def _filter_unread_messages_by_subject(self, imap_client: IMAP4_SSL) -> list:
        filtered_unread_messages = []
        unread_message_numbers = self._get_unread_messages(imap_client)
        unread_messages = unread_message_numbers[0].split()
        for unread_message in unread_messages:
            message = self._get_filtered_message(imap_client, unread_message)
            if re.search(self._subject, message.get(MailService.SUBJECT_PART)):
                filtered_unread_messages.append(message)
        
        return filtered_unread_messages

    def _get_filtered_message(self, imap_client: IMAP4_SSL, unread_message) -> str:
        message_part_with_data, message_part_without_data = self._get_message(imap_client, unread_message)
        message_part_format, message_part_data = message_part_with_data
        return email.message_from_bytes(message_part_data)

    def _get_unread_messages(self, imap_client: IMAP4_SSL, mailbox: str = INBOX_FOLDER) -> list:
        try:
            imap_client.select(mailbox)
            message_status, unread_message_numbers = imap_client.search(None, MailService.UNREAD_MSG)
            return unread_message_numbers
        except Exception as exception:
            raise ModuleError('Failed to get unread messages', exception) from exception

    def _get_message(self, imap_client: IMAP4_SSL, mail_data: str) -> list:
        try:
            message_status, message = imap_client.fetch(mail_data, MailService.MESSAGE_FORMAT)
            return message
        except Exception as exception:
            raise ModuleError('Failed to get message', exception) from exception


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
