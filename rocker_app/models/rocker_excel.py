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
import logging

_logger = logging.getLogger(__name__)


class rocker_excel_reports(models.Model):
    _name = "rocker.excel"
    _description = 'Rocker Reporting Excels'
    name = fields.Char('Title', required=True)
    date_executed = fields.Datetime('Execution date')
    excel_file = fields.Binary('Excel size')
    file_name = fields.Char('Excel File', size=64)
