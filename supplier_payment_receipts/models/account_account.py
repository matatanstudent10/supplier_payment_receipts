# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError


# Extensión del modelo account.account para agregar el campo cuentas_por_pagar
class AccountAccount(models.Model):
    _inherit = 'account.account'
    
    cuentas_por_pagar = fields.Boolean(
        string='Es Cuenta por Pagar',
        default=False,
        help='Marcar si esta cuenta se usa para cuentas por Pagar de proveedores.'
    )