# import-email-attachment

Module retrieves unreaded mails using IMAP protocol, filters them by subject and search for CSV or ZIP files in them, and uploads them to OneVizion through import (Module does not track the status of imports).

* onevizion* - dictionary with parameters for onevizion
  * ovImportName* - name of the import from the ADMIN_CONFIG_IMPORT page	
  * ovImportAction* - how to import the data. Valid choices are INSERT, UPDATE and INSERT_UPDATE
* email* - dictionary with parameters for mail
  * host* - IMAP protocol of your email client
  * port* - port for IMAP. By default is 993
  * username* - can be either a mail name or a user name, depending on the mail client
  * subject* - subject by which the emails will be filtered


Example of settings.json

```json
{
    "onevizion": {
        "ovUrl": "https://***.onevizion.com/",
        "ovAccessKey": "******",
        "ovSecretKey": "************",
        "ovImportName": "import name",
        "ovImportAction": "INSERT_UPDATE"
    },
    "email": {
        "host": "imap.gmail.com",
        "port": 993,
        "username": "******",
        "password": "************",
        "subject": "subject"
    }
}
```