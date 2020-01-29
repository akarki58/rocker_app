{'name': 'Rocker Reporting Application',
 'summary': 'Collect data from various datasources to Excel',
 'description': 'Collect data from various datasources to Excel',
 'author': 'Antti Kärki',
 'depends': ['base','web'],
 'license': 'AGPL-3',
 'category': 'Reporting',
 'version': '12.0.1.0.0',
 'data': [
    'views/database_view.xml',
    'views/excel_view.xml',
     'views/report_view.xml',
     'views/executor_view.xml',
      'data/rocker_data.xml',
     'views/rocker_menu.xml',
     'security/rocker_security.xml',
     'security/ir.model.access.csv',
     'wizard/rocker_popup_wizard.xml',
     'wizard/rocker_about.xml',
 ],
 'application': True,
 'installable': True,
 'images': ['static/description/main_screenshot.png'],

 }
