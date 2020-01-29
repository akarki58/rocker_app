# -*- coding: utf-8 -*-
#############################################################################
#
#    Copyright (C) 2019-Antti Kärki.
#    Author: Antti Kärki.
#
#    You can modify it under the terms of the GNU AFFERO
#    GENERAL PUBLIC LICENSE (AGPL v3), Version 3.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU AFFERO GENERAL PUBLIC LICENSE (AGPL v3) for more details.
#
#    You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
#    (AGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################

from odoo import api, fields, models
from odoo import exceptions
from odoo import http
import logging
import tempfile
from datetime import datetime
import win32com.client as win32
from win32com.client import constants
import pythoncom
import base64
import os
from . import rocker_connection

_logger = logging.getLogger(__name__)


# _logger.debug('A DEBUG message')
# _logger.info('An INFO message')
# _logger.warning('A WARNING message')
# _logger.error('An ERROR message')
# --log_level=:DEBUG

class Report(models.Model):
    _name = 'rocker.report'
    _description = 'Rocker Reporting'
    name = fields.Char('Name', required=True)
    report_type = fields.Selection(
        [('single', 'Single'), ('collection', 'Collection')], 'Report type', default='single')
    description = fields.Text('Description')
    active = fields.Boolean('Active?', default=True)
    schedule_onoff = fields.Boolean('Scheduling Active', default=False)
    store_history = fields.Boolean('Store Excel to history', default=True)
    date_published = fields.Date(string='Published', default=fields.Date.today())
    database = fields.Many2one('rocker.database', string='Datasource',
                               default=lambda self: self.env['rocker.database'].search([('name', '=', 'Odoo')]))
    report_ids = fields.Many2many('rocker.report', 'rocker_report_collection', 'collection_id', 'report_id',
                                  'Collection reports (check that Sheet names are unique!)',
                                  domain=[('report_type', '=', 'single')])
    collection_ids = fields.Many2many('rocker.report', 'rocker_report_collection', 'report_id', 'collection_id',
                                      'Report in Collections', domain=[('report_type', '=', 'collection')])
    column_headings = fields.Char('Column headings', default='Stage; Count', help="Column headings separated with ;")
    select_clause = fields.Text('Select', default=
    """select ns.name, count(*) 
    from note_note nn
    join note_stage_rel nsr on nsr.note_id = nn.id
    join note_stage ns on ns.id = nsr.stage_id 
    group by ns.name, ns.sequence
    order by ns.sequence""")
    sheet_name = fields.Char('Excel Sheet name', default='Data')
    excel_template = fields.Binary('Excel template', help="")
    excel_report = fields.Binary('Last Excel Report')
    author_id = fields.Many2one('res.users', string='Author', default=lambda self: self.env.user)
    date_executed = fields.Datetime()
    file_name = fields.Char('Excel Filename', size=64, default='report.xlsx')
    template_name = fields.Char('Template Filename', size=64, help="")
    perma_link = fields.Char('Permanent link to latest')
    execute_link = fields.Char('Execute & download link')
    interval_number = fields.Integer(String="Interval", default=1)
    execute_at = fields.Float(String="Execute at (timezone=UTC)")
    firstcall = fields.Date(String='First Execution')
    interval_type = fields.Selection(
        [('min', 'min'), ('hour', 'hour'), ('day', 'day'), ('month', 'month')], 'Execute report every', default='day')
    company_id = fields.Many2one('res.company', string='User belonging this company hierarchy can view report')
    _sqldriver = False

    @api.multi
    #
    def testexcel(self):
        _logger.debug('Starting test')
        mytmpdir = os.environ['TEMP']  # Must be uppercas
        filename = "test_report.xlsx"
        template_filename = "test_template.xlsx"
        try:
            os.remove(os.path.join(mytmpdir, filename))
        except:
            _logger.debug('Test_report does not exist')
        try:
            os.remove(os.path.join(mytmpdir, template_filename))
        except:
            _logger.debug('Test_template does not exist')
        try:
            _logger.debug('Pythoncom')
            pythoncom.CoInitialize()
            # first we create empty excel and store that to template field
            _logger.debug('win32.gencache.Ensuredispatch')
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False  # disable overwrite warning
            wb = excel.Workbooks.Add()
            sheet = wb.Worksheets(1)
            sheet.Name = "Data"
            sheet.Range("A1").Value = "This is a template!"
            _logger.debug('Save as ' + os.path.join(mytmpdir, template_filename) )
            wb.SaveAs(os.path.join(mytmpdir, template_filename))
            wb.Close()
            #
            excel.Application.Quit()
            # now we open that as template
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False  # disable overwrite warning
            wb = excel.Workbooks.Open(os.path.join(mytmpdir, template_filename))
            sheet = wb.Worksheets("Data")
            sheet.Range("A2").Value = "Added some data"
            sheet.Range("B2").Value = "Added some data"
            sheet.ListObjects.Add(1, sheet.Range(sheet.Cells(2, 1), sheet.Cells(2, 2))).Name = "DataTest"
            wb.SaveAs(os.path.join(mytmpdir, filename))
            wb.Close()
            _logger.debug('Excel quit')
            excel.Application.Quit()
            # except:
            #    raise exceptions.ValidationError('Excel: Something went wrong, check odoo.log')
        except Exception as e:
            raise exceptions.ValidationError(
                'Excel test\n\nTried to create files to: ' + mytmpdir + '\n\nCheck folder access rights\n\n' + str(e))
        context = {}
        context['message'] = "Excel worksheet creation seems to work!\nGenerated Excels in " + mytmpdir
        title = 'Success'
        view = self.env.ref('rocker_app.rocker_popup_wizard')
        view_id = False
        return {
            'name': title,
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rocker.popup.wizard',
            'views': [(view.id, 'form')],
            'view_id': view.id,
            'target': 'new',
            'context': context,
        }

    @api.multi
    # multi, otherwise no excels
    def export_xls(self, context=None):
        if self.active != True:
            raise exceptions.ValidationError('Report is not active or you are not allowed to view it!')
        _logger.info('Rocker reporting / Executing report: ' + self.name)
        filename = ''
        if not self.file_name:
            self.file_name = 'report.xlsx'
        if not self.sheet_name:
            self.sheet_name = 'Data'
        odoo_filename = self.file_name.strip()
        filename = self.file_name.strip()
        sheetname = self.sheet_name.strip()
        temp_filename = ''
        if not odoo_filename:
            odoo_filename = 'report.xlsx'
        if not sheetname:
            sheetname = 'Data'
        # remove existing from temp
        try:
            os.unlink(os.path.join(mytmpdir, filename))
        except:
            _logger.debug('File does not exist in TEMP')
        try:
            os.unlink(os.path.join(mytmpdir, template_filename))
        except:
            _logger.debug('Template does not exist in TEMP')

        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False  # disable overwrite warning
        mytmpdir = os.environ['TEMP']  # Must be uppercas

        # take template if exists
        if self.excel_template:
            try:
                file = tempfile.NamedTemporaryFile(mode='w+b', delete=False, suffix='.xlsx')
                filename = file.name
                temp_filename = file.name
                _logger.debug('Using temp file: ' + file.name)
                file2 = base64.b64decode(self.excel_template)
                file.write(file2)
                file.close()
            except:
                _logger.error('Exception: Temp file open')
                raise exceptions.ValidationError('Temp file in use, quit all Excels from task manager')

            # open excel(template) from TEMP
            workbook = excel.Workbooks.Open(filename)
            _logger.debug('Using template')
        else:
            # generate empty excel
            filename = self.file_name
            filename = os.path.join(mytmpdir, filename)
            _logger.debug('Using empty Excel: ' + filename)
            workbook = excel.Workbooks.Add()
            sheet = workbook.Worksheets(1)
            sheet.Name = sheetname
            # create About sheet
            self._about(workbook)
            # then activate data sheet
            sheet.Activate()

        if self.report_type == 'single':
            worksheet = self._worksheet(workbook, self.sheet_name)
            _logger.debug('Single report, using sheet: ' + self.sheet_name)
            con = self._create_connection(self.database)
            # let's fill with data
            if con:
                _logger.debug('Populate single report')
                self._populate_sql(con, worksheet, self.select_clause, self.column_headings)
            else:
                raise exceptions.ValidationError('No DB connection')

        elif self.report_type == 'collection':
            # do loop here
            for report in self.report_ids:
                sheetname = report.sheet_name.strip()
                if not sheetname:
                    raise exceptions.ValidationError('In Collections sheet names are manadatory')
                worksheet = self._worksheet(workbook, sheetname)
                _logger.debug('Collection report, using sheet: ' + report.sheet_name)
                con = self._create_connection(report.database)
                # let's fill with data
                if con:
                    _logger.debug('Collection report populate')
                    self._populate_sql(con, worksheet, report.select_clause, report.column_headings)
                else:
                    raise exceptions.ValidationError('No DB connection')

        # Refresh all pivot tables & graphs
        workbook.RefreshAll()

        # we use temp for saving excel
        file = tempfile.NamedTemporaryFile(mode='w+b', delete=True, suffix='.xlsx')
        filename = file.name
        _logger.debug('Storing Excel to temp file: ' + file.name)
        file.close()

        # save the  workbook
        workbook.SaveAs(filename)
        workbook.Close()

        # workbook ready let's save to database
        datenow = fields.datetime.now()
        _logger.debug('Open file for storing to Odoo: ' + filename)
        file = open(filename, 'rb')
        file.seek(0)
        # save to report log
        if self.store_history:
            _export_id = self.sudo().env['rocker.excel'].create({'name': self.name, 'date_executed': datenow,
                                                                 'excel_file': base64.b64encode(file.read()),
                                                                 'file_name': odoo_filename}, ).id
        # save generated excel
        base_url = self.env['ir.config_parameter'].sudo().get_param('web.base.url')
        permlink = base_url + '/web/login?db=' + self._cr.dbname + '&redirect=' + base_url + '/web/content/rocker.report/%s/excel_report/%s?download=true' % (
        self.id, self.file_name)
        execlink = self.env['ir.config_parameter'].sudo().get_param(
            'web.base.url') + '/web/login?db=' + self._cr.dbname + '&redirect=' + self.env[
                       'ir.config_parameter'].sudo().get_param(
            'web.base.url') + '/web?#model=rocker.report&view_type=list&menu_id=' + str(
            self.env.ref('rocker_app.rocker_menu').id) + '&action=' + str(
            self.env.ref('rocker_app.rocker_report_execute_request').id) + '&id=' + str(self.id)
        file.seek(0)
        self.sudo().write({
            'file_name': odoo_filename,
            'sheet_name': sheetname,
            'excel_report': base64.b64encode(file.read()),
            'date_executed': datenow,
            'perma_link': permlink,
            'execute_link': execlink,
        })

        excel.Application.Quit()
        file.close()  # closes the file, so we can right remove it
        # removing template too
        try:
            os.unlink(filename)
            _logger.debug("% s removed successfully" % filename)
        except OSError as error:
            _logger.debug(error)
            _logger.error("File can not be removed")
        # removing template too if exist
        if temp_filename:
            try:
                os.unlink(temp_filename)
                _logger.debug("% s removed successfully" % temp_filename)
            except OSError as error:
                _logger.debug(error)
                _logger.error("File can not be removed")

        # download
        return {
            'type': 'ir.actions.act_url',
            'name': 'excel',
            'url': '/web/content/rocker.report/%s/excel_report/%s?download=true' % (self.id, self.file_name),
        }

    def _worksheet(self, workbook, sheet, context=None):
        _logger.debug('Trying to find sheet: ' + sheet)
        if not sheet:
            sheet = 'Data'

        # find if sheet name exists in template or in workbook
        try:
            worksheet = workbook.Worksheets(sheet)
            _logger.debug('Found worksheet ' + worksheet.Name)
        except:
            _logger.debug('Exception sheet creation')
            worksheet = workbook.Worksheets.Add()
            worksheet.Name = sheet
            _logger.debug('Created new sheet: ' + sheet)
        return worksheet

    def _about(self, workbook, context=None):
        _logger.debug('Trying to find about sheet: ')
        # we do not create about to template, only to new workbook :-)
        try:
            aboutsheet = workbook.Worksheets('About')
            _logger.debug('Found existing about sheet ')
            aboutsheet.Columns.ClearContents()
        except:
            _logger.debug('About sheet not found, adding new')
            aboutsheet = workbook.Worksheets.Add()
            aboutsheet.Name = 'About'

        aboutsheet.Cells(2, 2).Value = 'Please donate!'
        for xlRow in xrange(2, 3, 1):
            aboutsheet.Hyperlinks.Add(Anchor=aboutsheet.Range('C{}'.format(xlRow)),
                                      Address="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=DGK3E2CC42EJ4&item_name=for+Rocker+Reporting+application+development&currency_code=EUR&source=url",
                                      ScreenTip="Click to Donate",
                                      TextToDisplay="Donate with PayPal")
        aboutsheet.Range('A2:C2').Font.Bold = True
        aboutsheet.Range('A2:C2').Font.Size = 12
        aboutsheet.Cells(4, 2).Value = 'Date executed:'
        aboutsheet.Cells(4, 3).Value = fields.datetime.now()
        aboutsheet.Cells(5, 2).Value = 'Datasource:'
        aboutsheet.Cells(5, 3).Value = self.name
        if self.report_type == 'collection':
            aboutsheet.Cells(6, 2).Value = 'Type:'
            aboutsheet.Cells(6, 3).Value = 'Collection'
        else:
            aboutsheet.Cells(6, 2).Value = 'SQL:'
            aboutsheet.Cells(6, 3).Value = self.select_clause
        aboutsheet.Columns.AutoFit()
        aboutsheet.Rows.AutoFit()

        return True

    def _create_connection(self, database):
        _database_record = self.env['rocker.database'].browse(database.mapped('id'))
        _datasource = _database_record.name
        _driver = _database_record.driver
        _odbcdriver = _database_record.odbcdriver
        _sid = _database_record.database
        _database = _database_record.database
        _host = _database_record.host
        _port = _database_record.port
        _user = _database_record.user
        _password = _database_record.password
        _logger.info('Connecting to ' + _database_record.name)
        con = None
        con = rocker_connection.rocker_connection.create_connection(_database_record)
        if con is not None:
            _logger.info('Database Connect OK')
            return con
        else:
            raise exceptions.ValidationError('Exception, No Database connection')

    def _populate_sql(self, con, worksheet, sql, headings, context=None):

        _headerlist = headings.split(';')
        header = [head.strip() for head in _headerlist]
        cols = len(header)
        # find datarange and clear
        try:
            usedrange = worksheet.Range(worksheet.Name)
            _logger.debug('Range found')
            usedrange.Delete(Shift=constants.xlUp)
        except:
            _logger.debug('Range not found')
            worksheet.ListObjects.Add(1, worksheet.Range(worksheet.Cells(1, 1),
                                                         worksheet.Cells(1, cols))).Name = worksheet.Name
            usedrange = worksheet.Range(worksheet.Name)

        c = 0
        for col in header:
            worksheet.Cells(1, c + 1).Value = header[c]
            c = c + 1

        # select
        cur = con.cursor()
        try:
            cur.execute(sql)
        except Exception as e:
            raise exceptions.ValidationError('Error in Select clause!\n\n' + str(e))

        records = cur.fetchall()
        i = len(records)
        # create data table
        _logger.debug('Creating Range rows')
        usedrange.Rows(i).Insert()

        j = 0
        r = 2

        if not (self._sqldriver == 'sqlserver'):
            # add data rows
            for row in records:
                j = len(row)
                c = 0
                for col in row:
                    worksheet.Cells(r, c + 1).Value = row[c]
                    c = c + 1
                r = r + 1

        elif (self._sqldriver == 'sqlserver'):
            for row in records:
                datalist = []
                datalist = list(row)
                j = len(datalist)
                c = 0
                for col in datalist:
                    worksheet.Cells(r, c + 1).Value = datalist[c]
                    c = c + 1
                r = r + 1

        cur.close()

        worksheet.Columns.AutoFit()

        # end
        if con is not None:
            con.close()
        return True

    @api.model
    def _cron_execute_report(self):
        _process_reports = self.env.cr.execute("""SELECT * FROM rocker_report
                                            WHERE active = True
                                            AND schedule_onoff = True
                                            AND COALESCE(firstcall, CURRENT_DATE - 1)::date <= current_date::date
                                            AND to_timestamp(COALESCE(execute_at,0) * 60 * 60)::time <= current_time::time
                                            AND (CASE WHEN interval_type = 'min' THEN COALESCE(date_executed, CURRENT_DATE - 3650) + interval_number * interval '1' minute
			   WHEN interval_type = 'hour' THEN COALESCE(date_executed, CURRENT_DATE - 3650) + interval_number * interval '1' hour
			   WHEN interval_type = 'day' THEN COALESCE(date_executed, CURRENT_DATE - 3650) + interval_number * interval '1' day
			   WHEN interval_type = 'month' THEN COALESCE(date_executed, CURRENT_DATE - 3650) + interval_number * interval '1' month
			ELSE now() END) <= now()
                                            """)
        _records = self.env.cr.fetchall()
        for _report in _records:
            self = self.env['rocker.report'].search([('id', '=', _report[0])])
            _logger.info('Cron execute report: ' + self.name)
            self.export_xls()

        _logger.debug('Nothing to do...boooring!')

    @api.model
    def _execute_xls(self, context=None):
        report_id = dict(self._context.get('params', {})).get('id')
        self = self.env['rocker.report'].search([('id', '=', report_id)])
        self.export_xls()
        _logger.debug('Base url: ' + self.env['ir.config_parameter'].sudo().get_param('web.base.url'))
        return {
            'type': 'ir.actions.act_url',
            'name': 'excel',
            'url': '/web/content/rocker.report/%s/excel_report/%s?download=true' % (self.id, self.file_name)
        }

    def show_about(self):
        _logger.debug('Open About ')
        context = {}
        context['message'] = "Rocker Reporting is nice"
        title = 'About Rocker Reporting'
        view = self.env.ref('rocker_app.rocker_about')
        view_id = self.env.ref('rocker_app.rocker_about').id
        return {
            'name': title,
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'rocker.about',
            'views': [(view.id, 'form')],
            'view_id': view.id,
            'target': 'new',
            'context': context,
        }
