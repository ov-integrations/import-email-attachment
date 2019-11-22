import DocuSignIntegration
import json

with open('settings.json', "rb") as PFile:
    password_data = json.loads(PFile.read().decode('utf-8'))

url_onevizion = password_data["url_OneVizion"]
login_onevizion = password_data["login_OneVizion"]
pass_onevizion = password_data["pass_OneVizion"]

import_name = password_data["import_name"]
login_mail = password_data["login_mail"]
pass_mail = password_data["pass_mail"]
subject_mail = password_data["subject_mail"]

DocuSignIntegration.Integration(url_onevizion=url_onevizion, login_onevizion=login_onevizion, pass_onevizion=pass_onevizion, import_name=import_name, login_mail=login_mail, pass_mail=pass_mail, subject_mail=subject_mail)
