{
    'name': 'Supplier Payment Receipts',
    'version': '15.0.1.0.0',
    'summary': 'Gestión avanzada de recibos de pago de proveedores con selección de líneas contables',
    'description': '''
        Este módulo permite:
        - Gestionar recibos de pago de proveedores
        - Cargar líneas contables pendientes (cuentas por cobrar)
        - Seleccionar específicamente qué líneas pagar
        - Manejar comisiones, descuentos y otros conceptos
        - Interface JavaScript interactiva para selección y cálculo automático
        - Reconciliación automática de pagos con líneas contables
        - Plantillas personalizadas de cheques por banco (Bancolombia, Davivienda, Occidente)
        - Control de impresión y reimpresión de cheques
    ''',
    'author': 'Fabio - Martinez',
    'website': 'https://www.empaquetadurasyempaques.com',
    'category': 'Accounting',
    'depends': [
        'base',
        'account',
        'mail',
        'multi_store',
        'dev_invoice_multi_payment',
        'consignments_eye',
        'cross_invoice_eye',
        'account_reports_custom',
        'partner_activity',
        'dev_customer_credit_limit',
        'account_check_printing',
        'web',
    ],
    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'views/supplier_payment_receipt_views.xml',
        'views/account_account_views.xml',
        'views/account_payment_term.xml',
        # 'views/account_payment_views.xml',
        # 'views/account_journal_views.xml',
        'wizard/supplier_debt_report_wizard_view.xml',
        'wizard/wizard_discount_view.xml',
        'data/ir_sequence_data.xml',
        'data/actions.xml',
        'reports/reports.xml',
        'reports/template_bancolombia.xml',
        'reports/template_davivienda.xml',
        'reports/template_occidente.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'supplier_payment_receipts/static/src/js/payment_lines_widget.js',
        ],
        'web.assets_qweb': [
            'supplier_payment_receipts/static/src/xml/payment_receipt_templates.xml',
        ],
    },
    'installable': True,
    'auto_install': False,
    'application': True,
    'license': 'LGPL-3',
}