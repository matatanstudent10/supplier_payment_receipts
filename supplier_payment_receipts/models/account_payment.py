# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError


class AccountPayment(models.Model):
    _inherit = 'account.payment'
    

    
    def action_draft(self):
        """Override to reset the payment state to draft."""
        for payment in self:
            payment_receipt = self.env['supplier.payment.receipt'].search([
                ('payment_id', '=', payment.id),
                ('state', '!=', 'draft')
            ], limit=1)
            if payment_receipt:
                raise UserError(_("No se puede restablecer el pago al estado de borrador porque está vinculado a un recibo de pago de proveedor que no está en estado de borrador."))
        
            move = self.move_id
            self.env.cr.execute("""
                SELECT 1 
                FROM supplier_account_crossings_lines ccl
                JOIN supplier_account_crossing cc ON ccl.supplier_id = cc.id
                JOIN account_move_line aml ON ccl.note_id = aml.id
                WHERE aml.move_id = %s 
                AND cc.state = 'posted'
                LIMIT 1
            """, (move.id,))
            
            if self.env.cr.fetchone():
                raise UserError(_(
                    'No se puede cancelar el comprobante de egreso porque tiene cruces de cuenta asociados en estado "Publicado".\n\n'
                    'Para cancelar este comprobante:\n'
                    '1. Primero cancele o revierta los cruces de cuenta asociados\n'
                    '2. Luego podrá cancelar este comprobante'
                ))
        return super(AccountPayment, self).action_draft()