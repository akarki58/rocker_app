# -*- coding: utf-8 -*-
from odoo import api, fields, models, _


class rocker_about(models.TransientModel):
    _name = "rocker.about"
    _description = "Rocker About box"

    name = fields.Html(string="Name", readonly=True, default="<H2>Rocker Reporting</H2")
    paypal = fields.Html(string="PauPal", readonly=True, default='''<p><p>If you like these reports, please click:<p>  <form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_blank">
<input type="hidden" name="cmd" value="_donations" />
<input type="hidden" name="business" value="DGK3E2CC42EJ4" />
<input type="hidden" name="item_name" value="for Rocker Reporting development" />
<input type="hidden" name="currency_code" value="EUR" />
<input type="image" src="https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif" border="0" name="submit" title="PayPal - The safer, easier way to pay online!" alt="Donate with PayPal button" />
<img alt="" border="0" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1" />
</form><p>
''')
    legal = fields.Html(string="Legal", readonly=True, default='''Author: Antti Kärki<p><font size="2">
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published by
    the Free Software Foundation.
<p>
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
<p>
    https://www.gnu.org/licenses.
''')
