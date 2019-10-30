import DocuSignIntegration
import json

with open('settings.json', "rb") as PFile:
    password_data = json.loads(PFile.read().decode('utf-8'))

url_onevizion = password_data["urlOneVizion"]
login_onevizion = password_data["loginOneVizion"]
pass_onevizion = password_data["passOneVizion"]
import_name = password_data["importname"]
login_mail = password_data["loginmail"]
pass_mail = password_data["passmail"]

DocuSignIntegration.Integration(url_onevizion=url_onevizion, login_onevizion=login_onevizion, pass_onevizion=pass_onevizion, import_name=import_name, login_mail=login_mail, pass_mail=pass_mail)
