Rocker Reporting for Windows Odoo
---------------
"c:\Odoo 19.0\python\python" pip3.exe install python-pptx
"c:\Odoo 19.0\python\python" pip3.exe install sqlalchemy


Datasource settings:
host: localhost
port:5432
db:Odoo18 or whatever you created
user: openpg (by default)
password: openpgpwd (by default)
if you create a user for reporting, remember to grant access (CONNECT; SELECT) to odoo database

NOTE NOTE!!! Database name is CASE SENSITIVE !!!!


PIP:
postgres drivers OK, no need to install
pip install SQLAlchemy
 

READ ROCKER_INSTALL.PPTX !!!
Slides: Test Excel functionality & Excel error 1: Change Odoo service properties
how to get Excel working


pip import python-pptx
NOTE: Mine went to C:\USERS\ANTTI\APPDATA\ROAMING\PYTHON\SITE-PACKAGES
        Odoo does not find that one. 
        I deleted above directory (Python) 
        and then reinstalled python-pptx
