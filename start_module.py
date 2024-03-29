import sys
import subprocess

installed_dependencies = subprocess.run(
    [sys.executable, '-m', 'pip', 'install', '-r', 'python_dependencies.ini'],
    check=True, stdout=subprocess.PIPE).stdout.decode().strip()
if 'Successfully installed' in installed_dependencies:
    raise Exception('Some required dependent libraries were installed. ' \
                    'Module execution has to be terminated now to use installed libraries on the next scheduled launch.')

import json
import re
from jsonschema import validate
from onevizion import IntegrationLog, LogLevel
from module import Module, ModuleError

with open('settings.json', "rb") as settings_file:
    settings_data = json.loads(settings_file.read().decode('utf-8'))

with open('settings_schema.json', 'rb') as settings_schema_file:
    settings_schema = json.loads(settings_schema_file.read().decode('utf-8'))

try:
    validate(instance=settings_data, schema=settings_schema)
except Exception as exception:
    raise Exception(f'Incorrect value in the settings file\n{str(exception)}') from exception

onevizion_data = settings_data['onevizion']
ov_url = re.sub('^(https|http)://', '', onevizion_data['ovUrl'])
ov_access_key = onevizion_data['ovAccessKey']
ov_secret_key = onevizion_data['ovSecretKey']

with open('ihub_parameters.json', 'rb') as ihub_parameters_file:
    module_data = json.loads(ihub_parameters_file.read().decode('utf-8'))

process_id = module_data['processId']
log_level = module_data['logLevel']
module_name = module_data['integrationName']

module_log = IntegrationLog(process_id, ov_url, ov_access_key, ov_secret_key, None, True, log_level)
module = Module(process_id, module_name, module_log, settings_data)
module.create_import_files_folder()

try:
    module.start()
except ModuleError as module_error:
    module_log.add(LogLevel.ERROR, str(module_error.message), str(module_error.description))
    raise module_error
finally:
    module.remove_import_files()
