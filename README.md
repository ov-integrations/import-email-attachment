# import-email-attachment

Requirements
- python version 3.7.2 or later
- python requests library (pip install requests)
- python onevizion library (pip install onevizion)
- enable IMAP protocol for Gmail, and also allow access for less secure applications

Features
- downloads a file attached to an unread message in Gmail (.zip or .csv)
- extracts .csv from .zip and starts importing with data from .csv

To start integration, you need to fill file settings.json:

For OneVizion:
- URL to onevizion site
- account username and password 
- name of import from onevizion site

For Gmail:
- Gmail username and password
- title of the subject of the letter

