# -*- coding: utf-8 -*-
#############################################################################
#
#    Cybrosys Technologies Pvt. Ltd.
#
#    Copyright (C) 2026-TODAY Cybrosys Technologies(<https://www.cybrosys.com>)
#    Author: Cybrosys Techno Solutions(<https://www.cybrosys.com>)
#
#    You can modify it under the terms of the GNU LESSER
#    GENERAL PUBLIC LICENSE (LGPL v3), Version 3.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU LESSER GENERAL PUBLIC LICENSE (LGPL v3) for more details.
#
#    You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENSE
#    (LGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################
import io
import json
from datetime import datetime, timedelta
from odoo import fields, models
from odoo.exceptions import ValidationError
from odoo.tools import json_default

try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter


class RestroReportWizard(models.TransientModel):
    """Pdf and Excel Report for Restaurant / Restro Orders"""

    _name = "restro.report.detail"
    _description = "Restro Report Details"

    checkin = fields.Date(help="Choose the Checkin Date", string="Checkin")
    checkout = fields.Date(help="Choose the Checkout Date", string="Checkout")
    room_id = fields.Many2one("product.template", string="Room",
                              help="Choose The Room")

    def action_restro_report_pdf(self):
        """Button action for creating Restro Report PDF"""
        data = {
            "booking": self.generate_data(),
        }
        return self.env.ref(
            "hotel_management_odoo.action_report_restro_report"
        ).report_action(self, data=data)

    def action_restro_report_excel(self):
        """Button action for creating Restro Report Excel"""
        data = {
            "booking": self.generate_data(),
        }
        return {
            "type": "ir.actions.report",
            "data": {
                "model": "restro.report.detail",
                "options": json.dumps(data, default=json_default),
                "output_format": "xlsx",
                "report_name": "Excel Report",
            },
            "report_type": "xlsx",
        }

    def generate_data(self):
        """Generate data to be printed in the report"""
        domain = []
        restro_list = []
        if self.checkin and self.checkout:
            if self.checkin > self.checkout:
                raise ValidationError(
                    ("Check-in date should be less than Check-out date")
                )
        if self.checkin:
            domain.append(
                ("booking_id.checkin_date", ">=", self.checkin),
            )
        if self.checkout:
            domain.append(
                ("booking_id.checkout_date", "<", self.checkout + timedelta(days=1)),
            )
        if self.room_id:
            domain.append(
                ("booking_id.room_line_ids.room_id", "=", self.room_id.id),
            )

        state_labels = {
            'draft': 'Draft',
            'reserved': 'Reserved',
            'check_in': 'Check In',
            'check_out': 'Check Out',
            'cancel': 'Cancelled',
            'done': 'Done',
        }
        food_booking_lines = self.env["food.booking.line"].search(domain)
        for line in food_booking_lines:
            booking = line.booking_id
            room_names = ", ".join(booking.room_line_ids.mapped("room_id.name")) if booking and booking.room_line_ids else ""
            guest_name = booking.partner_id.name if booking and booking.partner_id else (line.food_id.name if line.food_id else "")
            raw_state = booking.state if booking else ""
            restro_list.append({
                "name": guest_name,
                "room": room_names,
                "state": state_labels.get(raw_state, raw_state),
            })
        return restro_list

    def _format_datetime(self, value):
        """Safely format a datetime value that may be a str or datetime object."""
        if not value:
            return ""
        if isinstance(value, str):
            try:
                value = datetime.strptime(value[:19], '%Y-%m-%d %H:%M:%S')
            except ValueError:
                return value
        return value.strftime('%Y-%m-%d %H:%M:%S')

    def get_xlsx_report(self, data, response):
        """Organizing xlsx report"""
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet()
        cell_format = workbook.add_format(
            {"font_size": "14px", "bold": True, "align": "center",
             "border": True}
        )
        head = workbook.add_format(
            {"align": "center", "bold": True, "font_size": "23px",
             "border": True}
        )
        body = workbook.add_format(
            {"align": "left", "text_wrap": True, "border": True})
        sheet.merge_range("A1:D1", "Restro Report", head)
        sheet.set_column("A2:D2", 20)
        sheet.set_row(0, 30)
        sheet.set_row(1, 20)
        sheet.write("A2", "Sl No.", cell_format)
        sheet.write("B2", "Name", cell_format)
        sheet.write("C2", "Room No", cell_format)
        sheet.write("D2", "Status", cell_format)
        row = 2
        column = 0
        value = 1
        for i in data["booking"]:
            sheet.write(row, column, value, body)
            sheet.write(row, column + 1, i["name"], body)
            sheet.write(row, column + 2, i["room"], body)
            sheet.write(row, column + 3, i["state"], body)
            row = row + 1
            value = value + 1
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()
