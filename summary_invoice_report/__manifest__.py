# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.
{
    'name': 'Invoice Summary Report',
    'version': '15.0.1',
    'summary': 'Invoice Summary Report',
    'description': """Invoice Summary Report""",
    'license':'AGPL-3',
    'category': 'Accounting',

    'author' : 'Odoo Consultant medconsultantweb@gmail.com',
    'website' : 'http://www.weblemon.org',
    'depends': ['account', 'web'],
    'images': ['static/description/banner.jpg'],
    'data': [
        'security/ir.model.access.csv',

        'wizard/print_invoice_summary_view.xml',
        'reports/invoice_summary_report.xml'
    ],
    'installable': True,
    'auto_install': False,
    'application': True
}
