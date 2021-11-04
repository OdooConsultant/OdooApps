# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import models, fields, api, _
import xlwt
import io
import base64
from xlwt import easyxf
import datetime


class PrintInvoiceSummary(models.TransientModel):
    _name = "print.invoice.summary"

    @api.model
    def _get_from_date(self):
        company = self.env.user.company_id
        current_date = datetime.date.today()
        from_date = company.compute_fiscalyear_dates(current_date)['date_from']
        return from_date

    from_date = fields.Date(string='From Date', default=_get_from_date)
    to_date = fields.Date(string='To Date', default=datetime.date.today())
    invoice_summary_file = fields.Binary('Invoice Summary Report')
    file_name = fields.Char('File Name')
    invoice_report_printed = fields.Boolean('Invoice Report Printed')
    invoice_status = fields.Selection([('all', 'All'), ('paid', 'Paid'), ('un_paid', 'Unpaid')],
                                      string='Invoice Status')
    partner_id = fields.Many2one('res.partner', 'Customer')
    invoice_objs = fields.Many2many('account.move', string='Invoices')
    amount_total_debit = fields.Float('Amount debit')
    amount_total_credit = fields.Float('Amount credit')

    def action_print_invoice_summary_csv(self):
        new_from_date = self.from_date.strftime('%Y-%m-%d')
        new_to_date = self.to_date.strftime('%Y-%m-%d')

        workbook = xlwt.Workbook()
        amount_tot = 0
        column_heading_style = easyxf('font:height 200;font:bold True;')
        worksheet = workbook.add_sheet('Invoice Summary')
        worksheet.write(2, 3, self.env.user.company_id.name,
                        easyxf('font:height 200;font:bold True;align: horiz center;'))
        worksheet.write(4, 2, new_from_date, easyxf('font:height 200;font:bold True;align: horiz center;'))
        worksheet.write(4, 3, 'To', easyxf('font:height 200;font:bold True;align: horiz center;'))
        worksheet.write(4, 4, new_to_date, easyxf('font:height 200;font:bold True;align: horiz center;'))
        worksheet.write(6, 0, _('Invoice Number'), column_heading_style)
        worksheet.write(6, 1, _('Customer'), column_heading_style)
        worksheet.write(6, 2, _('Invoice Date'), column_heading_style)
        worksheet.write(6, 3, _('Invoice Amount'), column_heading_style)
        worksheet.write(6, 4, _('Invoice Currency'), column_heading_style)
        worksheet.write(6, 5, _('Amount in Company Currency (' + str(self.env.user.company_id.currency_id.name) + ')'),
                        column_heading_style)

        worksheet.col(0).width = 5000
        worksheet.col(1).width = 5000
        worksheet.col(2).width = 5000
        worksheet.col(3).width = 5000
        worksheet.col(4).width = 5000
        worksheet.col(5).width = 8000

        worksheet2 = workbook.add_sheet('Customer wise Invoice Summary')
        worksheet2.write(1, 0, _('Customer'), column_heading_style)
        worksheet2.write(1, 1, _('Paid Amount'), column_heading_style)
        worksheet2.write(1, 2, _('Invoice Currency'), easyxf('font:height 200;font:bold True;align: horiz left;'))
        worksheet2.write(1, 3, _('Amount in Company Currency (' + str(self.env.user.company_id.currency_id.name) + ')'),
                         easyxf('font:height 200;font:bold True;align: horiz left;'))
        worksheet2.col(0).width = 5000
        worksheet2.col(1).width = 5000
        worksheet2.col(2).width = 4000
        worksheet2.col(3).width = 8000

        row = 7
        customer_row = 2
        for wizard in self:
            customer_payment_data = {}
            heading = 'Invoice Summary Report'
            worksheet.write_merge(0, 0, 0, 5, heading, easyxf(
                'font:height 210; align: horiz center;pattern: pattern solid, fore_color black; font: color white; font:bold True;' "borders: top thin,bottom thin"))
            heading = 'Customer wise Invoice Summary'
            worksheet2.write_merge(0, 0, 0, 3, heading, easyxf(
                'font:height 200; align: horiz center;pattern: pattern solid, fore_color black; font: color white; font:bold True;' "borders: top thin,bottom thin"))
            if wizard.invoice_status == 'all':
                domain = [('invoice_date', '>=', wizard.from_date),
                         ('invoice_date', '<=', wizard.to_date),
                         ('state', 'not in', ['draft', 'cancel'])]
            elif wizard.invoice_status == 'paid':
                domain = [('invoice_date', '>=', wizard.from_date),
                          ('invoice_date', '<=', wizard.to_date),
                          ('payment_state', '=', 'paid')]
            else:
                domain = [('invoice_date', '>=', wizard.from_date),
                          ('invoice_date', '<=', wizard.to_date),
                          ('payment_state', '=', 'not_paid')]

            if self.partner_id:
                domain +=[('partner_id', '=', self.partner_id.id)]

            invoice_objs = self.env['account.move'].search(domain)
            for invoice in invoice_objs:
                invoice_date = invoice.invoice_date.strftime('%Y-%m-%d')
                amount = 0
                for journal_item in invoice.line_ids:
                    amount += journal_item.debit
                worksheet.write(row, 0, invoice.display_name)
                worksheet.write(row, 1, invoice.partner_id.name)
                worksheet.write(row, 2, invoice_date)
                worksheet.write(row, 3, invoice.amount_total)
                worksheet.write(row, 4, invoice.currency_id.symbol)
                worksheet.write(row, 5, amount)
                amount_tot += amount
                row += 1
                key = u'_'.join((invoice.partner_id.name, invoice.currency_id.name)).encode('utf-8')
                key = str(key, 'utf-8')
                if key not in customer_payment_data:
                    customer_payment_data.update(
                        {key: {'amount_total': invoice.amount_total, 'amount_company_currency': amount}})
                else:
                    paid_amount_data = customer_payment_data[key]['amount_total'] + invoice.amount_total
                    amount_currency = customer_payment_data[key]['amount_company_currency'] + amount
                    customer_payment_data.update(
                        {key: {'amount_total': paid_amount_data, 'amount_company_currency': amount_currency}})
            worksheet.write(row + 2, 5, amount_tot, column_heading_style)

            for customer in customer_payment_data:
                worksheet2.write(customer_row, 0, customer.split('_')[0])
                worksheet2.write(customer_row, 1, customer_payment_data[customer]['amount_total'])
                worksheet2.write(customer_row, 2, customer.split('_')[1])
                worksheet2.write(customer_row, 3, customer_payment_data[customer]['amount_company_currency'])
                customer_row += 1

            fp = io.BytesIO()
            workbook.save(fp)
            excel_file = base64.b64encode(fp.getvalue())
            wizard.invoice_summary_file = excel_file
            wizard.file_name = 'Invoice Summary Report.xls'
            wizard.invoice_report_printed = True
            fp.close()
            return {
                'view_mode': 'form',
                'res_id': wizard.id,
                'res_model': 'print.invoice.summary',
                'view_type': 'form',
                'type': 'ir.actions.act_window',
                'context': self.env.context,
                'target': 'new',
            }


    def _search_invoices(self):
        if self.invoice_status == 'all':
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('state', 'not in', ['draft', 'cancel'])]
        elif self.invoice_status == 'paid':
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('payment_state', '=', 'paid')]
        else:
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('payment_state', '=', 'not_paid')]

        if self.partner_id:
            domain += [('partner_id', '=', self.partner_id.id)]

        return self.env['account.move'].search(domain)

    def action_print_invoice_summary_pdf(self):

        if self.invoice_status == 'all':
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('state', 'not in', ['draft', 'cancel'])]
        elif self.invoice_status == 'paid':
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('payment_state', '=', 'paid')]
        else:
            domain = [('invoice_date', '>=', self.from_date),
                      ('invoice_date', '<=', self.to_date),
                      ('payment_state', '=', 'not_paid')]

        if self.partner_id:
            domain += [('partner_id', '=', self.partner_id.id)]

        self.invoice_objs = self._search_invoices()
        self.amount_total_debit = 0
        self.amount_total_credit = 0

        for invoice in self.invoice_objs:
            if invoice.move_type != 'out_refund':
                self.amount_total_debit += invoice.amount_total
            else:
                self.amount_total_credit += invoice.amount_total

        xml_id = 'invoice_summary.action_report_invoice_summary'
        res = self.env.ref(xml_id).report_action(self.ids, None)
        res['id'] = self.env.ref(xml_id).id
        return res
