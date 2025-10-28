# -*- coding: utf-8 -*-
from odoo import models, fields, api,_


class AccountMoveReversalInherit(models.TransientModel):
    _inherit = 'account.move.reversal'


    receipt_supplier_id = fields.Many2one('supplier.payment.receipt', string='Related Receipt')


    @api.model
    def default_get(self, fields):
        res = super(AccountMoveReversalInherit, self).default_get(fields)
        
        # Capturar los valores del contexto
        if self.env.context.get('from_payment_receipt'):
            res['from_payment_receipt'] = True
            res['payment_id'] = self.env.context.get('payment_id')
            res['receipt_supplier_id'] = self.env.context.get('receipt_supplier_id')
        return res
    
    def reverse_moves(self):
        """Heredar para manejar el campo modelo_libre después de la reversión"""
        # Llamar al método padre
        result = super(AccountMoveReversalInherit, self).reverse_moves()
        
        # Usar los campos del wizard en lugar del contexto
        if self.from_payment_receipt and self.payment_id:
            # Actualizar modelo_libre a False
            self.env.cr.execute("""
                UPDATE account_payment 
                SET modelo_libre = %s 
                WHERE id = %s
            """, (False, self.payment_id.id))
            
            # Actualizar el estado del recibo
            if self.receipt_supplier_id:
                self.receipt_supplier_id.write({
                    'is_reversed': True,
                })
                
                # Mensaje en el chatter
                if self.new_move_ids:
                    self.receipt_supplier_id.message_post(
                        body=_('Recibo revertido. Asiento de reversión: %s') % 
                             ', '.join(self.new_move_ids.mapped('name'))
                    )
            
            self.env.cr.commit()
        return result