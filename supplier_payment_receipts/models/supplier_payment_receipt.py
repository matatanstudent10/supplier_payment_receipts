# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError
import time 
import logging
from datetime import datetime, timedelta
import base64
import io
from xlsxwriter import Workbook


_logger = logging.getLogger(__name__)

class SupplierPaymentReceipt(models.Model):
    _name = 'supplier.payment.receipt'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _description = 'Recibo de Pago de Proveedor'
    _order = 'date desc, id desc'


    def default_store_id(self):
        context = self._context
        current_uid = context.get('uid')
        user_id = self.env['res.users'].browse(current_uid)
        if user_id.store_id.name in ['EMPAQUETADURAS','006 PRODUCCION']:
            store_id = []
        else:
            store_id = user_id.store_id.id
        return store_id   


    name = fields.Char(
        string='Número de Egreso',
        readonly=True,
        tracking=True,
        default=lambda self: _('Nuevo')
    )
    
    partner_id = fields.Many2one(
        'res.partner',
        string='Proveedor',
        required=True,
        domain=[('parent_id', '=', False)],
        states={'draft': [('readonly', False)]},
        tracking=True,
        readonly=True
    )
    
    date = fields.Date(
        string='Fecha',
        required=True,
        default=fields.Date.context_today,
        states={'draft': [('readonly', False)]},
        tracking=True,
        readonly=True
    )
    
    journal_id = fields.Many2one(
        'account.journal',
        string='Diario',
        required=True,
        domain=[('type', 'in', ['bank', 'cash'])],
        states={'draft': [('readonly', False)]},
        tracking=True,
        readonly=True
    )
    
    payment_method_line_id = fields.Many2one(
        'account.payment.method.line',
        string='Método de Pago',
        domain="[('id', 'in', available_payment_method_line_ids)]",
        required=True,
        store=True, 
        copy=False,
        states={'draft': [('readonly', False)]},
        readonly=True
    )
    
    amount_total = fields.Monetary(
        string='Total Pago',
        compute='_compute_amount_total',
        store=True,
        readonly=False,
        tracking=True,
        currency_field='currency_id'
    )
    
    currency_id = fields.Many2one(
        'res.currency',
        string='Moneda',
        required=True,
        default=lambda self: self.env.company.currency_id,
        states={'draft': [('readonly', False)]},
        tracking=True,
        readonly=True
    )
    
    company_id = fields.Many2one(
        'res.company',
        string='Compañía',
        required=True,
        default=lambda self: self.env.company
    )
    
    state = fields.Selection([
        ('draft', 'Borrador'),
        ('posted', 'Publicado'),
        ('cancel', 'Cancelado')
    ], string='Estado', default='draft',tracking=True, readonly=True)
    
    payment_id = fields.Many2one(
        'account.payment',
        string='Pago Creado',
        tracking=True,
        readonly=True,
        copy=False,
    )
    
    line_ids = fields.One2many(
        'supplier.payment.receipt.line',
        'receipt_id',
        string='Líneas de Cuentas por Pagar',
        states={'draft': [('readonly', False)]},
        readonly=True
    )
    
    memo = fields.Char(
        string='Nota',
        states={'draft': [('readonly', False)]},
        tracking=True,
        readonly=True,
    )
    
    destination_account_id = fields.Many2one(
        'account.account',
        string='Cuenta de Destino',
        domain="[('deprecated', '=', False), ('company_id', '=', company_id), ('user_type_id.type', '!=', 'view')]",
        states={'draft': [('readonly', False)]},
        readonly=True,
        tracking=True,
        help="Cuenta donde se registrará el pago"
    )
    
    store_id = fields.Many2one(
        'res.store',
        string='Sucursal',
        domain="[('company_id', '=', company_id)]",
        states={'draft': [('readonly', False)]},
        readonly=True,
        tracking=True,
        default=default_store_id,
        help="Tienda asociada al recibo de pago"
    )
    
    analytic_account_id = fields.Many2one(
        'account.analytic.account',
        string='Centro de Costo',
        domain="[('company_id', '=', company_id)]",
        states={'draft': [('readonly', False)]},
        readonly=True,
        tracking=True,
        help="Cuenta analítica asociada al recibo de pago"
    )
    
    is_multiple_payment = fields.Boolean(
        string='Pago Múltiple',
        tracking=True,
        default=False,
        help="Indica si este recibo es parte de un pago múltiple"
    )
    
    payment_term_text = fields.Char(
        string='Términos de Pago',
        help="Términos de pago asociados al Proveedor"
    )
    
    check_amount_in_words = fields.Char(
        string="Cantidad en Palabras",
        store=True,
        compute='_compute_check_amount_in_words',
    )
    
    is_reconciled = fields.Boolean(
        string='Conciliado',
        related='payment_id.is_reconciled',
        readonly=True,
        tracking=True,
        store=True,
        help="Indica si el comprobante ha sido reconciliado con las líneas contables"
    )
    invoices_count = fields.Integer(
        string='Facturas Asociadas',
        compute='_compute_invoices_count',
        help="Número de facturas asociadas a este comprobante"
    )

    is_reversed = fields.Boolean('Revertido', default=False, readonly=True, tracking=True,
        help="Indica si el comprobante ha sido revertido")

    partner_additional_bank_id = fields.Many2one('res.partner', 
        copy= False, 
        string='Banco', 
        tracking=True)

    reconcilable_accounting_entries = fields.Boolean(
        string='Apuntes Conciliables', 
        default =False,
        copy=False,
        tracking=True,)

    payment_method_code = fields.Char(
        related='payment_method_line_id.code',
        help="Technical field used to adapt the interface to the payment type selected.")


    available_payment_method_line_ids = fields.Many2many('account.payment.method.line',
        compute='_compute_payment_method_line_fields')

    sellos = fields.Selection([('none','Ninguno'),
                               ('crossed','Cruzado'),
                               ('beneficiary','Primer Beneficiario')], 
                              default ='none' , 
                              string="Sellos", 
                              tracking=True)
    
    apply_discount = fields.Boolean(
        string='Aplicar Descuento',
        default=False,
        states={'draft': [('readonly', False)]},    
        readonly=True,
        tracking=True,
        help="Indica si se debe aplicar un descuento al pago"
    )
    
    correria = fields.Boolean(
        string='Gastos de Correría',
        default=False,
        states={'draft': [('readonly', False)]},
        readonly=True,
        tracking=True,
        help="Indica si el pago es de tipo correría"
    )

    discount = fields.Float(
        string='Descuento',
        default=0.0,
        states={'draft': [('readonly', False)]},
        readonly=True,
        tracking=True,
        help="Monto del descuento a aplicar al pago"
    )


    advances_payment_text = fields.Text(
        string="Anticipos", 
        compute='_compute_has_advance_payment', 
        store=True, 
        help="Lista de anticipos disponibles para este proveedor"
    )

    has_advance_payment = fields.Boolean(
        string="Tiene Anticipo de Proveedor", 
        compute='_compute_has_advance_payment', 
        store=True, 
        help="Indica si el proveedor tiene anticipos disponibles"
    )
    
    #?--------------CONTROL DE IMPRESIÓN Y REIMPRESION DE CHEQUES---------------
    impreso = fields.Boolean(
        string='Impreso', 
        tracking=True, 
        default=False,
        help="Indica si el cheque ya fue impreso"
    )

    
    def action_print_checks(self):
        """Método para imprimir cheques usando plantillas personalizadas por banco"""
        if self.journal_id.template_report_id:
            self.write({'impreso': True})
            return self.journal_id.template_report_id.report_action(self)

    
    def action_reprint_checks(self):
        """Método para reimprimir cheques"""
        self.write({'impreso': False})
        self.message_post(
            body=_('Se ha reiniciado el estado de impresión'),
            message_type='notification'
        )

    @api.depends('partner_id')
    def _compute_has_advance_payment(self):
        """Método optimizado para calcular anticipos usando ORM"""
        
        # Limpiar registros sin partner_id
        records_without_partner = self.filtered(lambda r: not r.partner_id)
        records_without_partner.update({
            'has_advance_payment': False,
            'advances_payment_text': False
        })
        
        # Procesar solo registros con partner_id
        records_with_partner = self - records_without_partner
        if not records_with_partner:
            return
        
        # Búsqueda optimizada con search_read (más eficiente que search + acceso a campos)
        domain = [
            ('partner_id', 'in', records_with_partner.mapped('partner_id.id')),
            ('parent_state', '=', 'posted'),
            ('account_id.advance_account_supplier', '=', True),
            ('company_id', '=', 1),
            ('amount_residual', '>', 0.00)
        ]
        
        # Solo cargar los campos que necesitamos
        fields_to_read = [
            'partner_id', 'name', 'amount_residual', 
            'currency_id', 'move_id', 'journal_id'
        ]
        
        advances_data = self.env['account.move.line'].search_read(
            domain, 
            fields_to_read,
            order='partner_id, name'
        )
        
        if not advances_data:
            # Si no hay datos, limpiar todos los registros
            records_with_partner.update({
                'has_advance_payment': False,
                'advances_payment_text': False
            })
            return
        
        # Pre-cargar datos relacionados de una vez (para evitar consultas lazy)
        move_ids = [adv['move_id'][0] for adv in advances_data if adv['move_id']]
        journal_ids = [adv['journal_id'][0] for adv in advances_data if adv['journal_id']]
        currency_ids = [adv['currency_id'][0] for adv in advances_data if adv['currency_id']]
        
        # Obtener nombres de moves en una sola consulta
        moves_data = {}
        if move_ids:
            moves = self.env['account.move'].browse(move_ids).read(['name'])
            moves_data = {move['id']: move['name'] for move in moves}
        
        # Obtener info de journals en una sola consulta
        journals_data = {}
        if journal_ids:
            journals = self.env['account.journal'].browse(journal_ids).read(['saldos_iniciales'])
            journals_data = {j['id']: j['saldos_iniciales'] for j in journals}
        
        # Obtener símbolos de monedas en una sola consulta
        currencies_data = {}
        if currency_ids:
            currencies = self.env['res.currency'].browse(currency_ids).read(['symbol'])
            currencies_data = {c['id']: c['symbol'] for c in currencies}
        
        # Agrupar anticipos por partner_id
        advances_by_partner = {}
        for advance in advances_data:
            partner_id = advance['partner_id'][0]  # search_read devuelve [id, name]
            
            if partner_id not in advances_by_partner:
                advances_by_partner[partner_id] = []
            
            # Enriquecer datos con información pre-cargada
            advance_info = {
                'name': advance['name'],
                'amount_residual': advance['amount_residual'],
                'move_name': moves_data.get(advance['move_id'][0]) if advance['move_id'] else '',
                'is_saldos_iniciales': journals_data.get(advance['journal_id'][0], False) if advance['journal_id'] else False,
                'currency_symbol': currencies_data.get(advance['currency_id'][0], '$') if advance['currency_id'] else '$'
            }
            
            advances_by_partner[partner_id].append(advance_info)
        
        # Asignar resultados a cada registro
        for payment in records_with_partner:
            partner_advances = advances_by_partner.get(payment.partner_id.id, [])
            
            if partner_advances:
                payment.has_advance_payment = True
                
                # Generar texto de anticipos con eliminación de duplicados
                unique_entries = set()
                for advance in partner_advances:
                    # Usar nombre del journal si es saldos_iniciales, sino nombre del move
                    move_name = advance['name'] if advance['is_saldos_iniciales'] else advance['move_name']
                    currency_symbol = advance['currency_symbol']
                    amount = advance['amount_residual']
                    
                    # Formato: "NOMBRE_DOCUMENTO (SIMBOLO_MONEDA MONTO)"
                    unique_entries.add(f"{move_name} ({currency_symbol}{amount:,.2f})")
                
                # Unir todas las entradas únicas y ordenarlas
                payment.advances_payment_text = ', '.join(sorted(unique_entries))
            else:
                payment.has_advance_payment = False
                payment.advances_payment_text = False

    @api.depends('state')
    def _compute_invoices_count(self):
        for record in self:
            record.invoices_count = len(record.line_ids.filtered('account_move_line_id'))

    
    @api.depends('payment_method_line_id', 'currency_id', 'amount_total')
    def _compute_check_amount_in_words(self):
        for pay in self:
            if pay.currency_id:
                pay.check_amount_in_words = pay.currency_id.amount_to_text(pay.amount_total)
            else:
                pay.check_amount_in_words = False

    @api.depends('journal_id', 'currency_id')
    def _compute_payment_method_line_fields(self):
        for pay in self:
            payment_type = 'outbound'  # Since this is for supplier payments
            pay.available_payment_method_line_ids = pay.journal_id._get_available_payment_method_lines(payment_type)
            to_exclude = pay._get_payment_method_codes_to_exclude()
            if to_exclude:
                pay.available_payment_method_line_ids = pay.available_payment_method_line_ids.filtered(lambda x: x.code not in to_exclude)

    @api.onchange('is_multiple_payment')
    def _onchange_is_multiple_payment(self):
        """Manejar activación/desactivación del pago múltiple"""
        if self.is_multiple_payment:

            if self.currency_id.name != 'COP':
                raise UserError(_('El pago múltiple solo está disponible para la moneda COP.'))
            if self.line_ids:
                self.line_ids = [(5, 0, 0)]
        else:
            if self.line_ids:
                self.line_ids = [(5, 0, 0)]


    @api.onchange('discount')
    def _onchange_discount_parent(self):
        """Aplicar descuento del padre a todas las líneas"""
        if self.discount is not False and self.line_ids:
            # Aplicar el descuento a todas las líneas que no sean líneas de descuento
            for line in self.line_ids.filtered(lambda l: not l.is_discount_line):
                line.discount = self.discount
                # Recalcular value_discount automáticamente
                if line.balance:
                    line.value_discount = abs(line.balance * line.discount / 100)    

    @api.onchange('journal_id')
    def _onchange_journal_id(self):
        ''' Compute the 'payment_method_line_id' field.
        This field is not computed in '_compute_payment_method_fields' because it's a stored editable one.
        '''
        for pay in self:
            payment_type = 'outbound'  # Since this is for supplier payments
            available_payment_method_lines = pay.journal_id._get_available_payment_method_lines(payment_type)

            # Select the first available one by default.
            if pay.payment_method_line_id in available_payment_method_lines:
                pay.payment_method_line_id = pay.payment_method_line_id
            elif available_payment_method_lines:
                pay.payment_method_line_id = available_payment_method_lines[0]._origin
            else:
                pay.payment_method_line_id = False    
    
    
    @api.onchange('partner_id')
    def _onchange_partner_id(self):
        """Limpiar líneas cuando cambia el Proveedor"""
        if self.partner_id:
            self.line_ids = [(6, 0, [])]
            self.destination_account_id = self.partner_id.property_account_payable_id.id
            self.payment_term_text = self.partner_id.property_payment_term_id.name
            self.memo = f"Pago de proveedor {self.partner_id.name}"

    def charges_data_partner(self):
        """Obtener fechas de cargos del Proveedor"""
        self.destination_account_id = self.partner_id.property_account_payable_id.id
        self.payment_term_text = self.partner_id.property_payment_term_id.name
        self.memo = f"Pago de proveedor {self.partner_id.name}"
            
    @api.onchange('store_id')
    def _onchange_analytic_account(self):
            for record in self:
                if record.store_id:
                    record.analytic_account_id = self.env['account.analytic.account'].search([('tipo_cuenta','=','venta'),('store_id','=', record.store_id.id)], limit=1).id


    def _get_payment_method_codes_to_exclude(self):
        # can be overriden to exclude payment methods based on the payment characteristics
        self.ensure_one()
        return []

    def action_load_invoices(self):
        """Cargar líneas de cuentas por cobrar pendientes + anticipos CON DESCUENTOS AUTOMÁTICOS"""
        if not self.partner_id:
            raise UserError(_('Debe seleccionar un Proveedor primero.'))
        
        # Buscar cuentas por cobrar (payable accounts)
        payable_accounts = self.env['account.account'].search([
            ('cuentas_por_pagar', '=', True),
            ('reconcile', '=', True),
            ('company_id', '=', self.company_id.id)
        ])
        
        if not payable_accounts:
            raise UserError(_('No se encontraron cuentas por cobrar configuradas.'))
        
        # 🎯 OBTENER PLAZO DE PAGO DEL PROVEEDOR (UNA SOLA VEZ)
        partner_payment_term = self.partner_id.property_supplier_payment_term_id
        
        # OPTIMIZACIÓN 1: Una sola consulta con OR (evita 2 queries)
        if not self.correria:
            domain = [
                ('company_id', '=', self.company_id.id),
                ('partner_id', '=', self.partner_id.id),
                ('parent_state', '=', 'posted'),
                ('amount_residual', '!=', 0),
                ('account_id', 'in', payable_accounts.ids),
                '|',
                ('parent_move_type', 'in', ['in_invoice','in_refund']),
                ('parent_move_type', 'in', ['in_invoice','entry']),
            ]
        else:
            domain = [
                ('company_id', '=', self.company_id.id),
                ('move_id.empleado', '=', self.partner_id.id),
                ('parent_state', '=', 'posted'),
                ('amount_residual', '<', 0),
                ('account_id', 'in', payable_accounts.ids),
                '|',
                ('parent_move_type', 'in', ['in_invoice','in_refund']),
                ('parent_move_type', 'in', ['in_invoice','entry']),
            ]

        # 🎯 CARGAR DATOS DE LÍNEAS
        pending_lines_data = self.env['account.move.line'].search_read(
            domain,
            ['id', 'partner_id', 'account_id', 'move_id', 'date', 'date_maturity', 'amount_residual', 
            'balance', 'currency_id', 'analytic_account_id'],
        )
        
        # CONSULTA ADICIONAL: ANTICIPOS
        domain_anticipos = [
            ('company_id', '=', self.company_id.id),
            ('partner_id', '=', self.partner_id.id),
            ('parent_state', '=', 'posted'),
            ('account_id.advance_account_supplier', '=', True),
            ('amount_residual', '>', 0.00)
        ]

        advance_lines_data = self.env['account.move.line'].search_read(
            domain_anticipos,
            ['id', 'partner_id', 'account_id', 'move_id', 'date', 'date_maturity', 'amount_residual', 
            'balance', 'currency_id', 'analytic_account_id']
        )

        # COMBINAR ambas listas
        pending_lines_data.extend(advance_lines_data)
        
        if not pending_lines_data:
            raise UserError(_('No se encontraron líneas de cuentas por cobrar pendientes ni anticipos para este Proveedor.'))
        
        # 🎯 OBTENER INFORMACIÓN COMPLETA DE LAS FACTURAS
        move_ids = [line['move_id'][0] for line in pending_lines_data if line['move_id']]
        moves_refs = {}
        moves_info = {}  # Para almacenar info completa
        
        if move_ids:
            # Obtener información completa de las facturas
            moves_data = self.env['account.move'].search_read(
                [('id', 'in', move_ids)],
                ['id', 'ref', 'amount_untaxed', 'amount_total', 'move_type']
            )
            
            for move in moves_data:
                move_id = move['id']
                moves_refs[move_id] = move['ref'] or ''
                moves_info[move_id] = {
                    'amount_untaxed': move['amount_untaxed'] or 0.0,
                    'amount_total': move['amount_total'] or 0.0,
                    'move_type': move['move_type']
                }
        
        # DELETE líneas existentes
        self.env.cr.execute("DELETE FROM supplier_payment_receipt_line WHERE receipt_id = %s", (self.id,))
        
        # PREPARAR DATOS para inserción masiva
        lines_data = []
        credit_lines = []
        debit_lines = []
        has_applicable_discount = False
        current_date = self.date
        
        for line_data in pending_lines_data:
            is_advance = line_data in advance_lines_data
            
            # 🔥 CALCULAR PRICE_SUBTOTAL PROPORCIONALMENTE
            move_id = line_data['move_id'][0] if line_data['move_id'] else None
            price_subtotal = 0.0
            
            if move_id and move_id in moves_info:
                move_info = moves_info[move_id]
                
                if move_info['move_type'] in ['in_invoice', 'in_refund']:
                    # Para facturas, calcular proporción del subtotal
                    amount_untaxed = move_info['amount_untaxed']
                    amount_total = move_info['amount_total']
                    amount_paying = abs(line_data['amount_residual'])
                    
                    if amount_total > 0:
                        # Calcular proporción del subtotal según el monto que se está pagando
                        proportion = amount_paying / amount_total
                        price_subtotal = amount_untaxed * proportion
                    else:
                        price_subtotal = amount_untaxed
                else:
                    # Para otros tipos de documento, usar el monto como subtotal
                    price_subtotal = abs(line_data['amount_residual'])
            else:
                # Si no hay información de la factura, usar el monto como subtotal
                price_subtotal = abs(line_data['amount_residual'])
            
            # 🎯 CALCULAR DESCUENTO BASADO EN EL PLAZO DE PAGO DEL PROVEEDOR
            line_discount = 0.0
            line_discount_value = 0.0
            
            if not is_advance:  # Solo aplicar descuento a facturas (no anticipos)
                # 🎯 USAR PLAZO DE PAGO DEL PROVEEDOR EN LUGAR DEL DOCUMENTO
                if partner_payment_term and partner_payment_term.early_discount and partner_payment_term.discount_days > 0:
                    invoice_date = line_data.get('date')
                    
                    if invoice_date:
                        # Convertir string a date si es necesario
                        if isinstance(invoice_date, str):
                            invoice_date = datetime.strptime(invoice_date, '%Y-%m-%d').date()
                        
                        try:
                            # Calcular fecha límite para descuento desde la fecha de la factura
                            discount_date = partner_payment_term._get_last_discount_date(invoice_date)
                            
                            factura_name = line_data['move_id'][1] if line_data['move_id'] else 'Sin nombre'
                            
                            # Si estamos dentro del plazo de descuento
                            if current_date <= discount_date:
                                line_discount = partner_payment_term.discount_percentage
                                # 🔥 APLICAR DESCUENTO SOBRE EL SUBTOTAL PROPORCIONADO
                                line_discount_value = price_subtotal * (partner_payment_term.discount_percentage / 100.0)
                                has_applicable_discount = True
                            else:
                                _logger.info(f"❌ Descuento NO aplicable (fuera de plazo) para {factura_name}")
                                
                        except Exception as e:
                            _logger.error(f"Error calculando descuento para {factura_name}: {e}")
                    else:
                        factura_name = line_data['move_id'][1] if line_data['move_id'] else 'Sin nombre'

            
            processed_line = {
                'receipt_id': self.id,
                'account_move_line_id': line_data['id'],
                'partner_id': line_data['partner_id'][0] if line_data['partner_id'] else False,
                'account_id': line_data['account_id'][0] if line_data['account_id'] else False,
                'move_id': line_data['move_id'][0] if line_data['move_id'] else False,
                'move_name': line_data['move_id'][1] if line_data['move_id'] else '',
                'date': line_data['date'],
                'debit': abs(line_data['amount_residual']) if line_data['amount_residual'] < 0 else 0.0,
                'credit': line_data['amount_residual'] if line_data['amount_residual'] > 0 else 0.0,
                'balance': (abs(line_data['amount_residual']) if line_data['amount_residual'] < 0 else 0.0) - (line_data['amount_residual'] if line_data['amount_residual'] > 0 else 0.0),
                'amount_residual': line_data['amount_residual'],
                'currency_id': line_data['currency_id'][0] if line_data['currency_id'] else self.currency_id.id,
                'cuentas_por_pagar': True,
                'selected': False,
                'description': f"{'APLICAR ANTICIPO' if is_advance else 'PAGO FACTURA PROVEEDOR'} - {line_data['move_id'][1] if line_data['move_id'] else ''}",
                'analytic_account_id': line_data['analytic_account_id'][0] if line_data['analytic_account_id'] else False,
                'sequence': 10,
                # 🎯 CAMPOS DE DESCUENTO CALCULADOS USANDO PLAZO DEL PROVEEDOR
                'discount': line_discount,
                'value_discount': line_discount_value,
                'price_subtotal': price_subtotal,  # 🔥 SUBTOTAL CALCULADO PROPORCIONALMENTE
                'move_ref_for_sort': moves_refs.get(line_data['move_id'][0], '') if line_data['move_id'] else ''
            }
            
            if line_data['amount_residual'] > 0:
                credit_lines.append(processed_line)
            else:
                debit_lines.append(processed_line)
        
        # Ordenar y combinar líneas
        credit_lines_sorted = sorted(credit_lines, key=lambda x: x['move_ref_for_sort'])
        debit_lines_sorted = sorted(debit_lines, key=lambda x: x['move_ref_for_sort'])
        lines_data = credit_lines_sorted + debit_lines_sorted
        
        # Limpiar campos temporales
        for line in lines_data:
            line.pop('move_ref_for_sort', None)

        # INSERCIÓN MASIVA
        if lines_data:
            columns = list(lines_data[0].keys())
            values_placeholder = ', '.join(['%s'] * len(columns))
            columns_str = ', '.join(columns)
            
            insert_query = f"""
                INSERT INTO supplier_payment_receipt_line ({columns_str})
                VALUES ({values_placeholder})
            """
            
            values_list = []
            for line_data in lines_data:
                values_list.append(tuple(line_data.values()))
            
            self.env.cr.executemany(insert_query, values_list)
        
        # 🎯 ACTIVAR DESCUENTO SI HAY AL MENOS UNA LÍNEA CON DESCUENTO
        update_vals = {}
        if has_applicable_discount:
            update_vals['apply_discount'] = True
            # _logger.info(f"✅ Descuentos aplicables encontrados usando plazo del proveedor: {partner_payment_term.name}")
        else:
            update_vals['apply_discount'] = False
            # _logger.info(f"❌ No hay descuentos aplicables con plazo del proveedor: {partner_payment_term.name if partner_payment_term else 'Sin plazo'}")
        
        # Invalidar cache y actualizar
        self.invalidate_cache(['line_ids'])
        
        # Actualizar con contexto especial
        context = {'skip_discount_line': True, 'creating_discount_line': True}
        self.with_context(context).write(update_vals)
        
        # Recalcular total
        self._compute_amount_total()
        if self.apply_discount:
            self._handle_discount_line()

    #!==========================MÉTODO OPTIMIZADO PARA CONFIRMAR RECIBO Y CREAR/ACTUALIZAR PAGO========================
    def action_confirm(self):
        """Confirmar el recibo y crear/actualizar el pago - MÉTODO OPTIMIZADO"""
        selected_lines = self.line_ids
        self._compute_amount_total()
        
        if self.store_id.name == 'EMPAQUETADURAS':
            raise UserError(_('No se puede crear un anticipo de adelanto para la sucursal EMPAQUETADURAS.'))        
        
        if not selected_lines:
            raise UserError(_('Debe seleccionar al menos una línea para pagar.'))
        
        if self.amount_total <= 0:
            raise UserError(_('El monto total debe ser mayor a cero.'))

        if not self.journal_id.default_account_id:
            raise UserError(_('El diario %s no tiene cuenta por defecto configurada.') % self.journal_id.name)
        
        lines_without_account = self.line_ids.filtered(lambda l: not l.account_id)
        if lines_without_account:
            line = lines_without_account[0]
            raise UserError(_('La línea "%s" no tiene cuenta contable definida.') % (line.description or line.move_name or 'Sin descripción'))

        lines_without_partner = self.line_ids.filtered(lambda l: not l.partner_id)
        if lines_without_partner:
            line = lines_without_partner[0]
            raise UserError(_('La línea "%s" no tiene Proveedor definido.') % (line.description or line.move_name or 'Sin descripción'))

        if self.currency_id.name != 'COP' and self.is_multiple_payment:
            raise UserError(_('Los pagos múltiples solo están permitidos en COP.'))

        if self.is_multiple_payment:
            currency_distinct_cop = self.line_ids.mapped('currency_id').filtered(lambda c: c.name != 'COP')
            if currency_distinct_cop:
                raise UserError(_('En pagos múltiples, todas las líneas deben estar en COP.'))

        if not self.payment_id:
            # CREAR NUEVO PAYMENT
            payment = self._create_payment_without_auto_lines()
            
            #  VERIFICAR Y LIMPIAR (por si acaso se crearon líneas)
            self._ensure_clean_move(payment.move_id)
            
            #  INSERTAR LÍNEAS MANUALES CON TU LÓGICA EXISTENTE
            self._insert_custom_move_lines_optimized(payment.move_id)

            optimized_context = {
                'skip_account_move_synchronization': True,
                'no_exchange_difference': True,
                'no_cash_basis': True,
            }
            payment.with_context(**optimized_context).action_post()
            self._reconcile_specific_invoices(payment)
            self.write({
                'payment_id': payment.id,
                'name': payment.move_id.name,
                'state': 'posted'
            })
            payment.message_post(
                body="Pago creado",
                message_type='comment'
            )
        else:
            # ACTUALIZAR PAYMENT EXISTENTE
            if not self.payment_id.move_id:
                raise UserError(_('El pago asociado no tiene un movimiento contable válido.'))
            payment = self.payment_id
            self.env.cr.execute("""
                UPDATE account_payment 
                SET modelo_libre = %s 
                WHERE id = %s
            """, (True, payment.id))    
            self.env.cr.commit()        
            # Actualizar valores del payment
            self._ensure_clean_move(payment.move_id)
            payment.write({
                'amount': self.amount_total,
                'date': self.date,
                'journal_id': self.journal_id.id,
                'destination_account_id': self.destination_account_id.id,
                'analytic_account_id': self.analytic_account_id.id if self.analytic_account_id else None,
                'store_id': self.store_id.id if self.store_id else None,
                'payment_method_line_id': self.payment_method_line_id.id,
                'partner_additional_bank_id': self.partner_additional_bank_id.id if self.partner_additional_bank_id else None,
                'partner_id': self.partner_id.id,
                'currency_id': self.currency_id.id,
                'ref': self.memo,
                'payment_type': 'outbound',
                'partner_type': 'supplier',
                'modelo_libre': True,
                'correria': False,
                'aplica_descuento': False,
                # 'date_payment_receipt': self.date_payment,
                'payment_new_model':True,
            })
            
            #  LIMPIAR LÍNEAS EXISTENTES del move
            #  INSERTAR NUEVAS LÍNEAS PERSONALIZADAS
            self._insert_custom_move_lines_optimized(payment.move_id)
            
            #  POST DEL PAYMENT ACTUALIZADO
            optimized_context = {
                'skip_account_move_synchronization': True,
                'no_exchange_difference': True,
                'no_cash_basis': True,
                'skip_auto_lines': True,
            }
            payment.with_context(**optimized_context).action_post()
            
            #  RECONCILIACIÓN ESPECÍFICA
            self._reconcile_specific_invoices(payment)
            
            # Actualizar estado del recibo
            self.write({
                'name': payment.move_id.name,
                'state': 'posted'
            })

        # Actualizar campo modelo_libre
        self.env.cr.execute("""
            UPDATE account_payment 
            SET modelo_libre = %s 
            WHERE id = %s
        """, (False, payment.id))


        self.env.cr.execute("""
            UPDATE account_move_line
            SET payment_id = %s
            WHERE move_id = %s""",
            (payment.id, payment.move_id.id))


    def _create_payment_without_auto_lines(self):
        """MÉTODO SIMPLE: Usar line_ids=[] para evitar _prepare_move_line_default_vals"""
        
        #  CONTEXTO ultra-optimizado
        no_lines_context = {
            'skip_account_move_synchronization': True,
            'check_move_validity': False,
            'skip_invoice_sync': True,
            'dont_create_taxes': True,
            'skip_payment_sync': True,
            'no_exchange_difference': True,
            'no_cash_basis': True,
            'skip_invoice_line_sync': True,
            'disable_tax_calculation': True,
            'force_delete': True,
            'skip_auto_lines': True,
            'manual_lines_only': True,
            'no_auto_reconcile': True,
            'tracking_disable': True,
            'mail_auto_subscribe_no_notify': True,
            'payment_no_auto_lines': True,
            'force_empty_lines': True,
        }
        
        #  VALORES del payment CON line_ids VACÍO
        payment_vals = {
            'payment_type': 'outbound',
            'partner_type': 'supplier',
            'partner_id': self.partner_id.id,
            'destination_account_id': self.destination_account_id.id,
            'amount': self.amount_total,
            'currency_id': self.currency_id.id,
            'date': self.date,
            'journal_id': self.journal_id.id,
            'analytic_account_id': self.analytic_account_id.id if self.analytic_account_id else None,
            'store_id': self.store_id.id if self.store_id else None,
            'payment_method_line_id': self.payment_method_line_id.id,
            'modelo_libre': True,
            'correria': False,
            'aplica_descuento': False,
            'ref': self.memo,
            'partner_additional_bank_id': self.partner_additional_bank_id.id if self.partner_additional_bank_id else None,
            # 'date_payment_receipt': self.date_payment,
            'payment_new_model':True,
            'line_ids': [],  # 🔥 CLAVE: Esto evita que llame a _prepare_move_line_default_vals
        }
        
        #  CREAR PAYMENT SIN LÍNEAS AUTOMÁTICAS
        payment = self.env['account.payment'].with_context(**no_lines_context).create(payment_vals)
        payment._compute_check_amount_in_words()
        # _logger.info(f"✅ Payment creado sin líneas automáticas: {payment.name}")
        
        return payment

    def _ensure_clean_move(self, move):
        """Verificar y limpiar move por si se crearon líneas automáticas"""
        
        # Refrescar desde BD para ver líneas actuales
        move.invalidate_cache(['line_ids'])
        existing_lines = move.line_ids
        
        if existing_lines:
            
            #  ELIMINACIÓN OPTIMIZADA POR LOTES (más rápida que DELETE masivo)
            line_ids = existing_lines.ids
            batch_size = 200  # Lotes pequeños para evitar locks largos
            
            for i in range(0, len(line_ids), batch_size):
                batch_ids = line_ids[i:i + batch_size]
                
                # Delete directo con ANY (más eficiente)
                self.env.cr.execute(
                    "DELETE FROM account_move_line WHERE id = ANY(%s)",
                    (batch_ids,)
                )
                
                # Commit cada lote para liberar locks inmediatamente
                if len(line_ids) > batch_size:
                    self.env.cr.commit()
            
            # Invalidar cache final
            move.invalidate_cache(['line_ids'])
            # _logger.info(f"✅ {len(existing_lines)} líneas automáticas eliminadas eficientemente")
        else:
            _logger.info("✅ Perfecto: No se crearon líneas automáticas")

    def _insert_custom_move_lines_optimized(self, move):
        """Insertar líneas personalizadas usando tu lógica existente optimizada"""
        
        #  PREPARAR DATOS para inserción masiva (tu lógica)
        move_line_values = []
        line_counter = 1
        
        # 🔥 LÍNEAS DE FACTURAS (cuentas por cobrar)
        for line in self.line_ids:
            move_line_values.append((
                move.id,                                      # move_id
                f'PAGO {line.move_name or line.description}', # name específica por factura
                line.account_id.id,                           # account_id de la factura
                line.partner_id.id,                           # partner_id
                line.debit,                                          # debit = 0 (es un pago)
                line.credit,                            # credit = monto exacto
                line.currency_id.id,                          # currency_id
                self.date,                                    # date
                self.company_id.id,                           # company_id
                line.analytic_account_id.id if line.analytic_account_id else None, # analytic_account_id
                line_counter,                                 # sequence
                line.account_move_line_id.id if line.account_move_line_id else None,
                self.store_id.id,
                line.tax_base_amount
            ))
            line_counter += 1
    
        # 🔥 LÍNEA DEL BANCO (contrapartida total)
        move_line_values.append((
            move.id,
            f'PAGO DE CLIENTE - {self.partner_id.name} - {self.date}',
            self.journal_id.default_account_id.id,          # cuenta del banco
            self.partner_additional_bank_id.id if self.partner_additional_bank_id.id else None,                                           # ✅ partner_id vacío para líneas de banco
            0.0,                                           # credit = 0
            self.amount_total,                              # debit = total del pago
            self.currency_id.id,                           
            self.date,                                     
            self.company_id.id,                            
            self.analytic_account_id.id if self.analytic_account_id else None, 
            line_counter,                                   
            None,
            self.store_id.id,
            0.0 # No tiene referencia para reconciliación
        ))
        
        #  INSERCIÓN MASIVA optimizada (tu lógica existente)
        self._execute_mass_insert_lines(move_line_values)
        
        # _logger.info(f"✅ {len(move_line_values)} líneas insertadas correctamente")

    def _execute_mass_insert_lines(self, move_line_values):
        """Inserción masiva optimizada - TU LÓGICA EXISTENTE"""
        
        account_ids = list(set([val[2] for val in move_line_values]))  # val[2] es account_id
    
        # Consulta para obtener todos los root_id de una vez
        self.env.cr.execute("""
            SELECT id, root_id FROM account_account WHERE id IN %s
        """, (tuple(account_ids),))
        
        account_root_mapping = dict(self.env.cr.fetchall())

        insert_query = """
            INSERT INTO account_move_line (
                create_uid, create_date, write_uid, write_date,
                move_id, journal_id, date, account_id, partner_id, name, 
                debit, credit, balance, amount_currency, amount_residual, amount_residual_currency,
                currency_id, company_id, company_currency_id, reconciled, blocked,
                analytic_account_id, sequence, account_root_id, quantity, 
                centralisation, discount, display_type, store_id, tax_base_amount
            ) VALUES %s
        """


        
        # Preparar valores completos con todos los campos requeridos (tu lógica)
        complete_values = []
        current_time = fields.Datetime.now()
        company_currency_id = self.company_id.currency_id.id
        
        for val in move_line_values:
            move_id, name, account_id, partner_id, debit, credit, currency_id, date, company_id, analytic_account_id, sequence, original_line_id, store_id, tax_base_amount = val
            
            balance = debit - credit
            
            # 🔧 USAR TU FUNCIÓN AUXILIAR para calcular amount_currency correctamente
            amount_currency = self._calculate_amount_currency(debit, credit, currency_id)
            amount_residual_currency = amount_currency
            if account_id not in account_root_mapping:
                raise ValueError(f"No se encontró root_id para account_id {account_id}")
            account_root_id = account_root_mapping[account_id]
            
            complete_val = (
                self.env.uid,                    # create_uid
                current_time,                    # create_date
                self.env.uid,                    # write_uid  
                current_time,                    # write_date
                move_id,                         # move_id
                self.journal_id.id,             # journal_id
                date,                           # date
                account_id,                     # account_id ✅ NUNCA NULL
                partner_id,                     # partner_id ✅ PUEDE SER NULL para banco
                name,                           # name
                debit,                          # debit
                credit,                         # credit
                balance,                        # balance
                amount_currency,                # amount_currency (CORREGIDO)
                balance,                        # amount_residual
                amount_residual_currency,       # amount_residual_currency (CORREGIDO)
                currency_id,                    # currency_id
                company_id,                     # company_id
                company_currency_id,            # company_currency_id
                False,                          # reconciled
                False,                          # blocked
                analytic_account_id,            # analytic_account_id (puede ser NULL)
                sequence,                       # sequence
                account_root_id,               # account_root_id (mismo que account_id)
                1.0,                           # quantity ✅ REQUERIDO
                'normal',                      # centralisation ✅ REQUERIDO
                0.0,                           # discount ✅ REQUERIDO
                None,                           # display_type (puede ser NULL)
                store_id,
                tax_base_amount
            )
            complete_values.append(complete_val)
        
        #  EJECUTAR inserción masiva usando execute_values de psycopg2 (tu método)
        from psycopg2.extras import execute_values
        execute_values(
            self.env.cr, insert_query, complete_values,
            template=None, page_size=1000
        )
        
        # Invalidar cache y recalcular
        move = self.env['account.move'].browse(move_line_values[0][0])  # move_id
        move.invalidate_cache()
        move.write({'invoice_date': self.date})

    def _calculate_amount_currency(self, debit, credit, currency_id):
        """
        Calcular amount_currency según las reglas específicas de Odoo:
        
        REGLA PRINCIPAL:
        - El importe expresado en la moneda secundaria debe ser POSITIVO cuando se CARGA la cuenta (debit > 0)
        - El importe expresado en la moneda secundaria debe ser NEGATIVO cuando se ACREDITA la cuenta (credit > 0)
        - Si la moneda es la MISMA que la de la empresa, este monto debe ser estrictamente IGUAL al balance
        
        Args:
            debit (float): Monto del débito
            credit (float): Monto del crédito  
            currency_id (int): ID de la moneda de la línea contable
            
        Returns:
            float: amount_currency calculado según las reglas de Odoo
        """
        company_currency_id = self.company_id.currency_id.id
        balance = debit - credit
        
        if currency_id == company_currency_id:
            # REGLA: Si la moneda es la misma que la de la empresa, amount_currency = balance
            return balance
        else:
            # REGLA: Si la moneda es diferente
            if debit > 0:
                return debit  # POSITIVO cuando se carga la cuenta
            elif credit > 0:
                return -credit  # NEGATIVO cuando se acredita la cuenta
            else:
                return 0.0

    #  MANTENER TODAS TUS FUNCIONES DE RECONCILIACIÓN EXISTENTES
    def _reconcile_specific_invoices(self, payment):
        """Reconciliación ADAPTATIVA: Directa para pocas facturas, Híbrida para muchas"""
        
        # Obtener líneas a reconciliar (solo cuentas por cobrar)
        if not self.reconcilable_accounting_entries:
            lines_to_reconcile = self.line_ids.filtered(lambda l: l.cuentas_por_pagar and l.account_move_line_id and l.balance != 0)
        else:
            lines_to_reconcile = self.line_ids.filtered(lambda l: l.account_id.reconcile and l.account_move_line_id and l.balance != 0)
        
        if not lines_to_reconcile:
            _logger.warning("No hay líneas de cuentas por cobrar para reconciliar")
            return
        
        total_facturas = len(lines_to_reconcile)
        
        # 🎯 ESTRATEGIA ADAPTATIVA basada en cantidad
        if total_facturas <= 5:
            self._direct_reconcile_small_batch(payment, lines_to_reconcile)
        else:
            self._hybrid_reconcile_large_batch(payment, lines_to_reconcile)

    def _direct_reconcile_small_batch(self, payment, lines_to_reconcile):
        """Reconciliación DIRECTA para ≤5 facturas - SIN OVERHEAD"""
        
        start_time = time.time()
        reconciled_count = 0
        
        # Obtener líneas de pago una sola vez
        payment_lines_dict = {}
        payment_lines = payment.move_id.line_ids.filtered(
            lambda l: l.account_id.user_type_id.type == 'payable' and not l.reconciled
        )
        
        for pline in payment_lines:
            payment_lines_dict[pline.name] = pline
        
        # Contexto mínimo
        direct_context = {
            'no_exchange_difference': True,
            'no_cash_basis': True,
        }
        
        #  RECONCILIACIÓN DIRECTA una por una (máxima simplicidad)
        for receipt_line in lines_to_reconcile:
            original_line = receipt_line.account_move_line_id
            
            if original_line.reconciled:
                continue
            
            # Búsqueda directa por nombre
            matching_payment_line = None
            if receipt_line.move_name:
                for payment_name, payment_line in payment_lines_dict.items():
                    if (receipt_line.move_name in payment_name and 
                        not payment_line.reconciled and
                        payment_line.account_id == original_line.account_id and
                        payment_line.partner_id == original_line.partner_id):
                        matching_payment_line = payment_line
                        break
            
            if matching_payment_line:
                try:
                    # Reconciliación directa sin verificaciones complejas
                    (original_line + matching_payment_line).with_context(**direct_context).reconcile()
                    reconciled_count += 1
                    # _logger.debug(f"✅ {receipt_line.move_name} reconciliada")
                    
                except Exception as e:
                    _logger.error(f"❌ Error {receipt_line.move_name}: {e}")
                    continue
        
        elapsed_time = time.time() - start_time
        _logger.info(f"⚡ DIRECTA completada en {elapsed_time:.2f}s - {reconciled_count}/{len(lines_to_reconcile)} facturas")

    def _hybrid_reconcile_large_batch(self, payment, lines_to_reconcile):
        """Reconciliación HÍBRIDA para >5 facturas - CON OPTIMIZACIONES"""
        
        start_time = time.time()
        
        # 🔥 CLASIFICAR facturas en COMPLETAS vs PARCIALES (solo para muchas facturas)
        full_payment_lines = []    
        partial_payment_lines = [] 
        
        payment_lines_dict = {}
        payment_lines = payment.move_id.line_ids.filtered(
            lambda l: l.account_id.user_type_id.type == 'payable' and not l.reconciled
        )
        
        for pline in payment_lines:
            payment_lines_dict[pline.name] = pline
        
        # Clasificar cada línea del recibo
        for receipt_line in lines_to_reconcile:
            original_line = receipt_line.account_move_line_id
            
            if original_line.reconciled:
                continue
            
            # Encontrar línea de pago correspondiente
            matching_payment_line = None
            if receipt_line.move_name:
                for payment_name, payment_line in payment_lines_dict.items():
                    if (receipt_line.move_name in payment_name and 
                        not payment_line.reconciled and
                        payment_line.account_id == original_line.account_id and
                        payment_line.partner_id == original_line.partner_id):
                        matching_payment_line = payment_line
                        break
            
            if not matching_payment_line:
                continue
            
            # 🎯 CLASIFICAR: ¿Es pago COMPLETO o PARCIAL?
            original_residual = abs(original_line.amount_residual)
            payment_amount = abs(matching_payment_line.amount_residual)
            
            # Tolerancia de 1 centavo para considerar "completo"
            if abs(original_residual - payment_amount) <= 0.01:
                full_payment_lines.append((original_line, matching_payment_line, receipt_line.move_name))
            else:
                partial_payment_lines.append((original_line, matching_payment_line, receipt_line.move_name))
        
        # _logger.info(f"📊 Clasificación: {len(full_payment_lines)} completas, {len(partial_payment_lines)} parciales")
        
        total_reconciled = 0
        
        #  RECONCILIACIÓN MASIVA para pagos completos
        if full_payment_lines:
            mass_reconciled = self._mass_reconcile_complete_payments(full_payment_lines)
            total_reconciled += mass_reconciled
        
        # 🎯 RECONCILIACIÓN INDIVIDUAL para pagos parciales
        if partial_payment_lines:
            individual_reconciled = self._individual_reconcile_partial_payments(partial_payment_lines)
            total_reconciled += individual_reconciled
        

    def _mass_reconcile_complete_payments(self, full_payment_pairs):
        """Reconciliación MASIVA para pagos completos - LOTES BALANCEADOS"""
        
        if not full_payment_pairs:
            return 0
        
        # 🎯 TAMAÑO ÓPTIMO para el caso de uso típico (40-50 facturas)
        total_count = len(full_payment_pairs)
        
        if total_count <= 60:
            BATCH_SIZE = total_count    # Todo junto si son ≤60 (ÓPTIMO para caso típico)
        else:
            BATCH_SIZE = 50            # Lotes de 50 para casos grandes (BALANCEADO)
        
        reconciled_count = 0
        mass_context = {
            'no_exchange_difference': True,
            'no_cash_basis': True,
        }
        
        for i in range(0, len(full_payment_pairs), BATCH_SIZE):
            batch = full_payment_pairs[i:i + BATCH_SIZE]
            batch_num = i // BATCH_SIZE + 1
            
            try:
                #  RECONCILIACIÓN MASIVA
                all_batch_lines = self.env['account.move.line']
                valid_pairs = 0
                
                for original_line, payment_line, move_name in batch:
                    if not original_line.reconciled and not payment_line.reconciled:
                        all_batch_lines += original_line + payment_line
                        valid_pairs += 1
                
                if all_batch_lines:
                    all_batch_lines.with_context(**mass_context).reconcile()
                    reconciled_count += valid_pairs
                    
                    # Commit inmediato para liberar locks
                    self.env.cr.commit()
                    
            except Exception as e:
                _logger.error(f"❌ LOTE {batch_num} FALLÓ: {e}")
                # 🛡️ FALLBACK individual
                for original_line, payment_line, move_name in batch:
                    try:
                        if not original_line.reconciled and not payment_line.reconciled:
                            (original_line + payment_line).with_context(**mass_context).reconcile()
                            reconciled_count += 1
                    except Exception:
                        continue
        
        return reconciled_count

    def _individual_reconcile_partial_payments(self, partial_payment_pairs):
        """Reconciliación INDIVIDUAL para pagos parciales - PRECISA"""
        
        if not partial_payment_pairs:
            return 0
        
        reconciled_count = 0
        individual_context = {
            'no_exchange_difference': True,
            'no_cash_basis': True,
        }
        
        # 🎯 RECONCILIACIÓN UNO POR UNO para pagos parciales
        for original_line, payment_line, move_name in partial_payment_pairs:
            try:
                if original_line.reconciled or payment_line.reconciled:
                    continue
                
                (original_line + payment_line).with_context(**individual_context).reconcile()
                reconciled_count += 1
                
                # Commit cada 10 reconciliaciones parciales
                if reconciled_count % 10 == 0:
                    self.env.cr.commit()
                    
            except Exception as e:
                _logger.error(f"Error reconciliando parcial {move_name}: {e}")
                continue
        
        return reconciled_count
    
    def action_confirm_payment_advance(self):
        """Confirmar el recibo y crear el pago de adelanto"""
        if not self.partner_id:
            raise UserError(_('Debe seleccionar un Proveedor primero.'))
        
        if self.store_id.name == 'EMPAQUETADURAS':
            raise UserError(_('No se puede crear un anticipo de adelanto para la sucursal EMPAQUETADURAS.'))
        
        if not self.journal_id:
            raise UserError(_('Debe seleccionar un diario para el pago.'))
        
        if self.state != 'draft':
            raise UserError(_('El recibo debe estar en estado borrador para confirmar el pago.'))
        
        # Validar que el monto total sea positivo
        if self.amount_total <= 0:
            raise UserError(_('El monto total del recibo debe ser mayor a cero.'))

        if not self.payment_id:
            payment = self._create_payment_advance()
            payment.action_post()
            self.write({
                'payment_id': payment.id,
                'name': payment.move_id.name,
                'state': 'posted'
            })

            self.env.cr.execute("""
                UPDATE account_payment 
                SET modelo_libre = %s 
                WHERE id = %s
            """, (False, payment.id))
        else:
            self.payment_id.write({
                'amount': self.amount_total,
                'date': self.date,
                'journal_id': self.journal_id.id,
                'destination_account_id': self.destination_account_id.id,
                'analytic_account_id': self.analytic_account_id.id if self.analytic_account_id else None,
                'store_id': self.store_id.id if self.store_id else None,
                'payment_method_line_id': self.payment_method_line_id.id,
                'partner_additional_bank_id': self.partner_additional_bank_id.id if self.partner_additional_bank_id else None,
                'partner_id': self.partner_id.id,
                'currency_id': self.currency_id.id,
                'ref': self.memo,
                'payment_type': 'outbound',
                'partner_type': 'supplier',
                'modelo_libre': True,
                'aplica_descuento': False,
                'es_anticipo': True,
                'correria': False,
                'payment_new_model':False,
            })
            self.payment_id.action_post()
            # Actualizar el estado del recibo
            self.write({
                'state': 'posted'
            })
            self.env.cr.execute("""
                UPDATE account_payment 
                SET modelo_libre = %s 
                WHERE id = %s
            """, (False, self.payment_id.id))
        
    def _create_payment_advance(self):
        """Crear un pago de adelanto sin líneas automáticas"""
        payment_vals = {
            'payment_type': 'outbound',
            'partner_type': 'supplier',
            'partner_id': self.partner_id.id,
            'amount': self.amount_total,
            'currency_id': self.currency_id.id,
            'partner_additional_bank_id': self.partner_additional_bank_id.id if self.partner_additional_bank_id else None,
            'date': self.date,
            'journal_id': self.journal_id.id,
            'destination_account_id': self.destination_account_id.id,
            'analytic_account_id': self.analytic_account_id.id if self.analytic_account_id else None,
            'store_id': self.store_id.id if self.store_id else None,
            'payment_method_line_id': self.payment_method_line_id.id,
            'ref': self.memo,
            'es_anticipo': True,
            'modelo_libre': True,
            'aplica_descuento': False,
            'payment_new_model':False,
            'correria': False,
        }
        
        payment = self.env['account.payment'].create(payment_vals)
        
        return payment        

    def unlink(self):
        for record in self:
            if record.state in ['posted','cancel']:
                raise ValidationError(_("No puedes eliminar este registro en estado Publicado o Cancelado"))
            if record.state == 'draft' and record.payment_id.move_id:
                record.payment_id.move_id._unlink_forbid_parts_of_chain()
            return super(SupplierPaymentReceipt, self).unlink()


    def action_cancel(self):
        """Cancelar recibo - OPTIMIZADA PARA MÚLTIPLES REGISTROS"""
        
        for record in self:
            # Validaciones básicas
            if record.state == 'cancel':
                continue
                
            if record.state != 'posted':
                raise UserError(_('El recibo %s no está en estado "Publicado".') % record.name)
            
            correria = self.env['wizard.apply.payment.advances'].search([('payment_id','=',record.payment_id.id),('state','=','posted')], limit=1)
            
            if correria:
                raise UserError(_('No se puede cancelar el recibo %s porque tiene una Correría (%s) asociada en estado "Publicado".') % record.name, correria.name)
            
            if not record.payment_id or not record.payment_id.move_id:
                record.write({'state': 'cancel'})
                continue
            
            move_id = record.payment_id.move_id.id
            
            
            
            #  VERIFICACIÓN MASIVA de cruces (una sola query para todos)
            record.env.cr.execute("""
                SELECT DISTINCT aml.move_id
                FROM supplier_account_crossings_lines ccl
                JOIN supplier_account_crossing cc ON ccl.supplier_id = cc.id
                JOIN account_move_line aml ON ccl.note_id = aml.id
                WHERE aml.move_id = %s 
                AND cc.state = 'posted'
            """, (move_id,))
            
            if record.env.cr.fetchone():
                raise UserError(_(
                    'No se puede cancelar el recibo %s porque tiene cruces de cuenta asociados en estado "Publicado".\n\n'
                    'Para cancelar este recibo:\n'
                    '1. Primero cancele o revierta los cruces de cuenta asociados\n'
                    '2. Luego podrá cancelar este recibo'
                ) % record.name)
            
            # Cancelar movimiento
            try:
                move = record.payment_id.move_id
                move.button_draft()
                move.button_cancel()
                record.write({'state': 'cancel'})
                
            except Exception as e:
                raise UserError(_(
                    'Error cancelando el recibo %s: %s'
                ) % (record.name, str(e)))

    
    def action_set_to_draft(self):
        """Volver a borrador"""
        if self.payment_id:
            self.payment_id.move_id.button_draft()        
        self.state = 'draft'
    
    def action_view_payment(self):
        """Ver el pago creado"""
        if not self.payment_id:
            raise UserError(_('No hay pago asociado a este recibo.'))
        
        return {
            'type': 'ir.actions.act_window',
            'name': _('Pago'),
            'res_model': 'account.payment',
            'res_id': self.payment_id.id,
            'view_mode': 'form',
            'target': 'current',
        }

    #!==========================descargas en excel===========================


    def action_export_excel_detail(self):
        """Exportar detalle del recibo en Excel"""
        self.ensure_one()
        
        if self.state != 'posted':
            raise UserError(_('Solo se pueden exportar recibos en estado "Publicado".'))
        
        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        workbook = Workbook(output, {'in_memory': True})
        
        # Configurar formatos
        formats = self._setup_excel_formats(workbook)
        
        # Crear la hoja principal
        worksheet = workbook.add_worksheet('Comprobante de Egreso')
        
        # Configurar la hoja
        self._setup_worksheet(worksheet)
        
        # Escribir el contenido
        current_row = self._write_excel_content(worksheet, workbook, formats)
        
        # Cerrar el workbook
        workbook.close()
        output.seek(0)
        
        # Crear el archivo adjunto
        filename = f'Comprobante_Egreso_{self.name}_{self.partner_id.name}.xlsx'
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(output.getvalue()),
            'res_model': 'supplier.payment.receipt',
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        
        # Retornar acción para descargar
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'new',
        }

    def _setup_excel_formats(self, workbook):
        """Configurar formatos para Excel"""
        formats = {
            'title': workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#8F5B92',
                'font_color': 'white',
                'border': 1
            }),
            'subtitle': workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#A0658A',
                'font_color': 'white',
                'border': 1
            }),
            'header': workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#E8E8E8',
                'border': 1,
                'text_wrap': True
            }),
            'header_main': workbook.add_format({
                'bold': True,
                'font_size': 12,
                'align': 'left',
                'valign': 'vcenter',
                'bg_color': '#F0F0F0',
                'border': 1
            }),
            'data': workbook.add_format({
                'font_size': 10,
                'align': 'left',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            }),
            'data_center': workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            }),
            'data_right': workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1
            }),
            'money': workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '[$-es-CO] #,##0.00'
            }),
            'money_bold': workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '[$-es-CO] #,##0.00',
                'bg_color': '#F0F0F0'
            }),
            'date': workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'num_format': 'dd/mm/yyyy'
            }),
            'section_title': workbook.add_format({
                'bold': True,
                'font_size': 12,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D0D0D0',
                'border': 1
            }),
            'advance_notice': workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFF3CD',
                'font_color': '#856404',
                'border': 1,
                'text_wrap': True
            })
        }
        return formats

    def _setup_worksheet(self, worksheet):
        """Configurar la hoja de Excel"""
        # Configurar anchos de columnas
        worksheet.set_column('A:A', 15)  # Documento
        worksheet.set_column('B:B', 12)  # Fecha
        worksheet.set_column('C:C', 30)  # Descripción
        worksheet.set_column('D:D', 20)  # Cuenta
        worksheet.set_column('E:E', 8)   # % Descuento
        worksheet.set_column('F:F', 12)  # V. Descuento
        worksheet.set_column('G:G', 12)  # Débito
        worksheet.set_column('H:H', 12)  # Crédito
        worksheet.set_column('I:I', 12)  # Balance
        
        # Configurar altura de filas por defecto
        worksheet.set_default_row(20)

    def _write_excel_content(self, worksheet, workbook, formats):
        """Escribir todo el contenido en Excel"""
        current_row = 0
        
        # Título principal
        current_row = self._write_header_section(worksheet, current_row, formats)
        current_row += 2
        
        # Información principal del pago
        current_row = self._write_payment_info_section(worksheet, current_row, formats)
        current_row += 2
        
        # Facturas pagadas (si las hay)
        invoice_lines = self.line_ids.filtered(lambda l: l.account_move_line_id)
        if invoice_lines:
            current_row = self._write_invoices_section(worksheet, current_row, invoice_lines, formats)
            current_row += 2
        
        # Líneas adicionales (si las hay)
        additional_lines = self.line_ids.filtered(lambda l: not l.account_move_line_id)
        if additional_lines:
            current_row = self._write_additional_lines_section(worksheet, current_row, additional_lines, formats)
            current_row += 2
        
        # Pago anticipado (si no hay líneas con documentos)
        if not invoice_lines:
            current_row = self._write_advance_payment_section(worksheet, current_row, formats)
            current_row += 2
        
        # Total final
        current_row = self._write_total_section(worksheet, current_row, formats)
        
        return current_row

    def _write_header_section(self, worksheet, row, formats):
        """Escribir la sección del encabezado"""
        # Título de la empresa
        worksheet.merge_range(row, 0, row, 8, self.company_id.name, formats['title'])
        row += 1
        
        # Título del documento
        worksheet.merge_range(row, 0, row, 8, 'COMPROBANTE DE EGRESO', formats['subtitle'])
        row += 1
        
        # Número del comprobante
        worksheet.merge_range(row, 0, row, 8, self.name, formats['subtitle'])
        row += 1
        
        return row

    def _write_payment_info_section(self, worksheet, row, formats):
        """Escribir la sección de información del pago"""
        # Encabezados
        headers = ['Proveedor', 'NIT/CC', 'Fecha', 'Número', 'Total Pagado']
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, formats['header'])
        
        # Datos
        row += 1
        worksheet.write(row, 0, self.partner_id.name, formats['data'])
        worksheet.write(row, 1, self.partner_id.vat or self.partner_id.fe_nit or '', formats['data_center'])
        worksheet.write(row, 2, self.date, formats['date'])
        worksheet.write(row, 3, self.name, formats['data_center'])
        worksheet.write(row, 4, self.amount_total, formats['money_bold'])
        
        return row + 1

    def _write_invoices_section(self, worksheet, row, invoice_lines, formats):
        """Escribir la sección de facturas pagadas"""
        # Título de la sección
        worksheet.merge_range(row, 0, row, 8, 'FACTURAS PAGADAS', formats['section_title'])
        row += 1
        
        # Encabezados dinámicos según si hay descuentos
        headers = ['Factura', 'Fecha', 'Descripción', 'Cuenta']
        if self.apply_discount:
            headers.extend(['% Dto', 'V. Descuento'])
        headers.extend(['Débito', 'Crédito', 'Balance'])
        
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, formats['header'])
        row += 1
        
        # Datos de las facturas
        total_facturas = 0
        for line in invoice_lines:
            col = 0
            worksheet.write(row, col, line.move_name or '', formats['data'])
            col += 1
            worksheet.write(row, col, line.date, formats['date'])
            col += 1
            worksheet.write(row, col, line.description or '', formats['data'])
            col += 1
            worksheet.write(row, col, f"{line.account_id.code} - {line.account_id.name}", formats['data'])
            col += 1
            
            if self.apply_discount:
                worksheet.write(row, col, f"{line.discount}%", formats['data_center'])
                col += 1
                worksheet.write(row, col, line.value_discount, formats['money'])
                col += 1
            
            worksheet.write(row, col, line.debit if line.debit > 0 else '', formats['money'])
            col += 1
            worksheet.write(row, col, line.credit if line.credit > 0 else '', formats['money'])
            col += 1
            worksheet.write(row, col, line.balance, formats['money'])
            
            total_facturas += line.balance
            row += 1
        
        # Total de facturas
        total_col = len(headers) - 1
        worksheet.merge_range(row, 0, row, total_col - 1, 'TOTAL FACTURAS:', formats['header_main'])
        worksheet.write(row, total_col, total_facturas, formats['money_bold'])
        
        return row + 1

    def _write_additional_lines_section(self, worksheet, row, additional_lines, formats):
        """Escribir la sección de ajustes y movimientos adicionales"""
        # Título de la sección
        worksheet.merge_range(row, 0, row, 8, 'AJUSTES Y MOVIMIENTOS ADICIONALES', formats['section_title'])
        row += 1
        
        # Encabezados
        headers = ['Descripción', 'Cuenta', 'Débito', 'Crédito', 'Balance']
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, formats['header'])
        row += 1
        
        # Datos
        for line in additional_lines:
            worksheet.write(row, 0, line.description or '', formats['data'])
            worksheet.write(row, 1, f"{line.account_id.code} - {line.account_id.name}", formats['data'])
            worksheet.write(row, 2, line.debit if line.debit > 0 else '', formats['money'])
            worksheet.write(row, 3, line.credit if line.credit > 0 else '', formats['money'])
            worksheet.write(row, 4, line.balance, formats['money'])
            row += 1
        
        return row

    def _write_advance_payment_section(self, worksheet, row, formats):
        """Escribir la sección de pago anticipado"""
        advance_text = f"PAGO ANTICIPADO\n\nEste monto será aplicado a futuras facturas del proveedor\n\nValor: {self.currency_id.symbol} {self.amount_total:,.2f}"
        worksheet.merge_range(row, 0, row + 2, 8, advance_text, formats['advance_notice'])
        return row + 3

    def _write_total_section(self, worksheet, row, formats):
        """Escribir la sección del total final"""
        worksheet.merge_range(row, 0, row, 7, 'TOTAL GENERAL DEL RECIBO:', formats['header_main'])
        worksheet.write(row, 8, self.amount_total, formats['money_bold'])
        
        # Mensaje de agradecimiento
        row += 2
        worksheet.merge_range(row, 0, row, 8, '¡Gracias por confiar en nosotros!', formats['section_title'])
        
        return row + 1

    @api.model
    def remove_lines(self, unselected_lines, selected_lines_remove):
        lines_not_remove = self.env['supplier.payment.receipt.line'].browse(selected_lines_remove)
        payment = lines_not_remove[0]
        if payment.receipt_id.state != 'draft':
            raise ValidationError(_("No puedes eliminar lineas en estado diferente de borrandor!"))
        lines_payments = self.env['supplier.payment.receipt.line'].search([('receipt_id', '=', payment.receipt_id.id), ('id', 'not in', lines_not_remove.ids)])     
        lines_payments.unlink() 
        total_amount = sum(payment.receipt_id.line_ids.mapped('balance'))  
        payment.receipt_id.write({'amount_total':total_amount})


    @api.model
    def update_amount(self, find_selected_lines):
        amount = 0
        lines = self.env['supplier.payment.receipt.line'].browse(find_selected_lines)
        total_amount = sum(lines.mapped('balance'))
        amount += total_amount
        return amount              

    #--------------ACCION PARA ABRIR EL PAGO------------
    def action_open_payment(self):
        self.ensure_one()
        return {
            'name': self.payment_id.name,
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'account.payment',
            'res_id': self.payment_id.id,
            'context': { 'create': False,'edit': False}
        }      


    #--------------ACCION PARA ABRIR EL ASIENTO CONTABLE------------
    def action_open_account_move(self):
        self.ensure_one()
        return {
            'name': self.payment_id.move_id.name,
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'account.move',
            'res_id': self.payment_id.move_id.id,
            'context': { 'create': False,'edit': False}
        }  


    def open_invoices_list(self):
        invoice_ids = self.line_ids.mapped('account_move_line_id.move_id')
        return {
            'name': 'Facturas',
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', invoice_ids.ids)],
            'context': {'create': False,'edit': False}
        }


    @api.constrains('line_ids')
    def _check_duplicate_documents_in_lines(self):
        """Validar que no haya documentos duplicados en las líneas del recibo - VERSIÓN OPTIMIZADA"""
        for receipt in self:
            if not receipt.line_ids:
                continue
            
            #  OPTIMIZACIÓN 1: Solo analizar líneas que tienen documento asignado
            lines_with_docs = receipt.line_ids.filtered('account_move_line_id')
            
            if len(lines_with_docs) <= 1:
                continue  # No puede haber duplicados con 0 o 1 línea
            
            #  OPTIMIZACIÓN 2: Detección de duplicados en O(n) con diccionario
            doc_registry = {}  # {doc_id: línea_original}
            duplicate_groups = []  # [(doc_id, [líneas_duplicadas])]
            
            for line in lines_with_docs:
                doc_id = line.account_move_line_id.id
                
                if doc_id in doc_registry:
                    # Encontramos un duplicado
                    original_line = doc_registry[doc_id]
                    
                    # Buscar si ya tenemos un grupo para este documento
                    existing_group = None
                    for group_doc_id, group_lines in duplicate_groups:
                        if group_doc_id == doc_id:
                            existing_group = group_lines
                            break
                    
                    if existing_group:
                        existing_group.append(line)
                    else:
                        # Crear nuevo grupo de duplicados
                        duplicate_groups.append((doc_id, [original_line, line]))
                else:
                    # Primera aparición del documento
                    doc_registry[doc_id] = line
            
            # 🚨 REPORTAR TODOS LOS DUPLICADOS ENCONTRADOS
            if duplicate_groups:
                error_messages = []
                
                for doc_id, duplicate_lines in duplicate_groups:
                    # Obtener información del documento
                    first_line = duplicate_lines[0]
                    move_name = first_line.account_move_line_id.move_id.name or 'Sin número'
                    partner_name = first_line.account_move_line_id.partner_id.name or 'Sin Proveedor'
                    count = len(duplicate_lines)
                    
                    error_messages.append(
                        f"• Documento '{move_name}' del Proveedor '{partner_name}' "
                        f"aparece {count} veces"
                    )
                
                # 🎯 MENSAJE DE ERROR COMPLETO Y ÚTIL
                raise ValidationError(_(
                    'Error: Documentos duplicados detectados\n\n'
                    'Los siguientes documentos están asignados múltiples veces en este recibo:\n\n'
                    '%s\n\n'
                    'Solución:\n'
                    '• Elimine las líneas duplicadas\n'
                    '• Cada documento puede aparecer solo una vez por recibo\n'
                    '• Revise las líneas antes de confirmar el pago'
                    % '\n'.join(error_messages)
                ))

    def action_reverse_payment(self):
        """Abrir el wizard de reversión estándar de Odoo para el asiento del pago"""
        self.ensure_one()
        
        # Validaciones
        if not self.payment_id:
            raise UserError(_('Este recibo no tiene un pago asociado.'))
        
        if not self.payment_id.move_id:
            raise UserError(_('El pago no tiene un asiento contable asociado.'))
            
        if self.state != 'posted':
            raise UserError(_('Solo se pueden revertir recibos en estado Publicado.'))
            
        if self.payment_id.move_id.state != 'posted':
            raise UserError(_('El asiento contable debe estar publicado para poder revertirlo.'))

        move = self.payment_id.move_id
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
                'No se puede reversar el recibo porque tiene cruces de cuenta asociados en estado "Publicado".\n\n'
                'Para reversar este recibo:\n'
                '1. Primero cancele o revierta los cruces de cuenta asociados\n'
                '2. Luego podrá reversarlo este recibo'
            ))

        self.env.cr.execute("""
            UPDATE account_payment 
            SET modelo_libre = %s 
            WHERE id = %s
        """, (True, self.payment_id.id))
        self.env.cr.commit()

        
        # Preparar el contexto con toda la información necesaria
        context = {
            'active_model': 'account.move',
            'active_ids': [self.payment_id.move_id.id],
            'active_id': self.payment_id.move_id.id,
            'default_move_ids': [(6, 0, [self.payment_id.move_id.id])],
            'default_company_id': self.company_id.id,
            'default_store_id': self.store_id.id if self.store_id else False,
            'default_payment_id': self.payment_id.id if self.payment_id else False,
            'default_receipt_id': self.id,
            'default_from_payment_receipt': True,
            'default_refund_method': 'cancel',  # Para asientos tipo 'entry' por defecto es 'cancel'
            'default_date': fields.Date.context_today(self),
            'default_reason': f'Reversión de recibo {self.name}',
        }
        
        # Invocar el wizard estándar de reversión
        return {
            'name': _('Revertir Asiento Contable'),
            'type': 'ir.actions.act_window',
            'res_model': 'account.move.reversal',
            'view_mode': 'form',
            'view_type': 'form',
            'target': 'new',
            'context': context,
        }


    def action_open_move_reverse(self):
        """Abrir el asiento contable del Comprobante de Egreso Reversado"""
        self.ensure_one()
        reverse_move = self.env['account.move'].search([('reversed_entry_id', '=',self.payment_id.move_id.id),('state','=','posted')],limit=1)
        return {
            'name': reverse_move.name,
            'type': 'ir.actions.act_window',
            'view_type': 'tree',
            'view_mode': 'tree,form',
            'res_model': 'account.move',
            'res_id': reverse_move.id,
            'context': {'create': False,'edit': False}
        }
        
    # En el modelo padre SupplierPaymentReceipt, agrega este método:

    def write(self, vals):
        """Override write para manejar descuentos automáticamente"""
        # Evitar bucle infinito
        if self.env.context.get('creating_discount_line'):
            return super(SupplierPaymentReceipt, self).write(vals)
        
        # Detectar si se está cambiando apply_discount
        apply_discount_changed = 'apply_discount' in vals
        old_apply_discount = self.apply_discount if apply_discount_changed else None
        
        result = super(SupplierPaymentReceipt, self).write(vals)
        
        # Solo procesar si estamos en borrador
        if self.state == 'draft' and not self.env.context.get('skip_discount_line'):
            if apply_discount_changed:
                # Si se cambió apply_discount, manejar según el nuevo valor
                if vals['apply_discount']:
                    # Se activó: crear línea de descuento si hay descuentos
                    self._handle_discount_line()
                else:
                    # Se desactivó: limpiar todo
                    self._clear_all_discounts()
            else:
                # Cambio normal: manejar línea de descuento según estado actual
                self._handle_discount_line()
        
        return result

    @api.model
    def create(self, vals):
        """Override create para crear línea de descuento automáticamente"""
        record = super(SupplierPaymentReceipt, self).create(vals)
        
        # Manejar línea de descuento si es necesario
        if record.state == 'draft':
            record._handle_discount_line()
        
        return record

    def _handle_discount_line(self):
        """Manejar línea de descuento basado en apply_discount"""
        for record in self:
            if record.apply_discount:
                # Si apply_discount está activo, crear/actualizar línea de descuento
                record._create_discount_line()
            else:
                # Si apply_discount está inactivo, eliminar línea de descuento
                record._remove_discount_line()

    def _create_discount_line(self):
        """Crear o actualizar UNA SOLA línea de descuento automáticamente"""
        for record in self:
            # Solo proceder si apply_discount está activo
            if not record.apply_discount:
                return
                
            # Encontrar LA línea de descuento (debe ser única)
            discount_line = record.line_ids.filtered('is_discount_line')
            
            # Calcular total de descuentos (excluyendo la línea de descuento)
            other_lines = record.line_ids.filtered(lambda l: not l.is_discount_line)
            total_discount = sum(other_lines.mapped('value_discount'))
            
            if total_discount <= 0:
                # Si no hay descuento, eliminar línea de descuento existente
                if discount_line:
                    discount_line.unlink()
                return
            
            # Verificar que existe la cuenta 1941
            discount_account = self.env['account.account'].browse(1941)
            if not discount_account.exists():
                continue
            
            discount_line_vals = {
                'receipt_id': record.id,
                'partner_id': record.partner_id.id,
                'account_id': 1941,
                'credit': total_discount,
                'debit': 0.0,
                'currency_id': record.currency_id.id,
                'description': f'DESCUENTO APLICADO - {record.partner_id.name}',
                'analytic_account_id': record.analytic_account_id.id if record.analytic_account_id else None,
                'cuentas_por_pagar': False,
                'selected': False,
                'sequence': 999,
                'is_discount_line': True,
            }
            
            # Usar contexto para evitar bucles
            context = {'creating_discount_line': True}
            
            if len(discount_line) == 1:
                # Actualizar la línea existente
                discount_line.with_context(context).write({
                    'credit': total_discount,
                    'description': f'DESCUENTO APLICADO - {record.partner_id.name}',
                    'partner_id': record.partner_id.id,
                    'currency_id': record.currency_id.id,
                    'analytic_account_id': record.analytic_account_id.id if record.analytic_account_id else None,
                })
            elif len(discount_line) > 1:
                # Si hay múltiples líneas de descuento (error), eliminar todas y crear una nueva
                discount_line.unlink()
                self.env['supplier.payment.receipt.line'].with_context(context).create(discount_line_vals)
            else:
                # No existe línea de descuento, crear una nueva
                self.env['supplier.payment.receipt.line'].with_context(context).create(discount_line_vals)

    def _remove_discount_line(self):
        """Eliminar línea de descuento cuando apply_discount está inactivo"""
        for record in self:
            discount_line = record.line_ids.filtered('is_discount_line')
            if discount_line:
                discount_line.unlink()

    @api.depends('line_ids.balance')
    def _compute_amount_total(self):
        for record in self:
            if not self.env.context.get('skip_auto_compute') and not self.env.context.get('anti_concurrency'):
                # Calcular total incluyendo la línea de descuento
                context = self.env.context.copy()
                context.update({'skip_discount_line': True})
                
                total = sum(record.line_ids.mapped('balance'))
                record.with_context(context).amount_total = total

    @api.onchange('apply_discount')
    def _onchange_apply_discount(self):
        """Manejar cambios en el campo apply_discount"""
        if self.apply_discount:
            # Si se activa, crear línea de descuento si hay descuentos
            self._create_discount_line()
        else:
            # Si se desactiva, limpiar TODOS los descuentos y eliminar línea
            self._clear_all_discounts()

    def _clear_all_discounts(self):
        """Limpiar todos los descuentos y eliminar línea de descuento"""
        for record in self:
            # 1. Eliminar línea de descuento automática
            discount_line = record.line_ids.filtered('is_discount_line')
            if discount_line:
                discount_line.unlink()
            
            # 2. Limpiar valores de descuento en todas las líneas
            other_lines = record.line_ids.filtered(lambda l: not l.is_discount_line)
            if other_lines:
                context = {'creating_discount_line': True}  # Evitar bucles
                other_lines.with_context(context).write({
                    'discount': 0.0,
                    'value_discount': 0.0,
                })
    
                

class CustomerPaymentReceiptLine(models.Model):
    _name = 'supplier.payment.receipt.line'
    _description = 'Línea de Recibo de Pago a Proveedores'
    
    receipt_id = fields.Many2one(
        'supplier.payment.receipt',
        string='Recibo',
        required=True,
        ondelete='cascade'
    )
    
    account_move_line_id = fields.Many2one(
        'account.move.line',
        string='Documento',
        domain="[('partner_id', '=', partner_id), ('amount_residual', '!=', 0), ('reconciled', '=', False)]"
    )
    
    partner_id = fields.Many2one(
        'res.partner',
        string='Cliente',
        required=True
    )
    
    account_id = fields.Many2one(
        'account.account',
        string='Cuenta Contable',
        domain="[('internal_type','!=','view'), ('deprecated', '=', False)]",
        required=True
    )
    
    move_id = fields.Many2one(
        'account.move',
        string='Documento',
        readonly=True
    )
    
    move_name = fields.Char(
        string='Número de Documento',
        readonly=True
    )
    
    date = fields.Date(
        string='Fecha',
        readonly=True
    )
    
    debit = fields.Monetary(
        string='Débito',
        currency_field='currency_id'
    )
    
    credit = fields.Monetary(
        string='Crédito',
        currency_field='currency_id'
    )
    
    balance = fields.Monetary(
        string='Balance',
        currency_field='currency_id',
        compute='_calculate_balance',
        store=True,
        readonly=True
    )
    
    amount_residual = fields.Monetary(
        string='Saldo Pendiente',
        currency_field='currency_id',
        readonly=True
    )
    
    tax_base_amount = fields.Monetary(
        string='Base Imponible',
        currency_field='currency_id',
        readonly=True
    )
    
    currency_id = fields.Many2one(
        'res.currency',
        string='Moneda',
        required=True,
        default=lambda self: self.env.company.currency_id,
    )
    
    cuentas_por_pagar = fields.Boolean(
        string='Cuentas por Cobrar',
        default=False,
        help="Indica si esta línea corresponde a una cuenta por cobrar"
    )
    
    selected = fields.Boolean(
        string='Seleccionado',
        default=False,
        help="Marcar para incluir en el pago"
    )
    
    description = fields.Char(
        string='Descripción',
        help="Descripción adicional para la línea"
    )
    
    is_multiple_payment = fields.Boolean(
        string='Pago Múltiple',
        default=False,
        help="Indica si esta línea es parte de un pago múltiple"
    )
    
    analytic_account_id = fields.Many2one(
        'account.analytic.account',
        string='Cuenta Analítica',
        help="Cuenta analítica asociada a esta línea"
    )
    
    sequence = fields.Integer(
        string='Secuencia',
        default=10,
        help="Orden de visualización de las líneas"
    )
    
    suitable_invoice_line_ids = fields.Many2many(
        'account.move.line',
        string='Líneas de Factura Adecuadas',
        compute='_get_domain',
        help="Líneas de factura adecuadas para este recibo basadas en el Proveedor y la cuenta contable"
    )
    
    company_id = fields.Many2one(
        'res.company',
        string='Compañía',
        default=lambda self: self.env.company,
        readonly=True,
    )
    
    discount = fields.Float(
        string='Descuento',
        help="Descuento aplicado a esta línea"
    )
    
    value_discount = fields.Monetary(
        string='Valor del Descuento',
        currency_field='currency_id',
        help="Valor monetario del descuento aplicado"
    )

    is_discount_line = fields.Boolean(
        string='Línea de Descuento',
        default=False,
        help="Indica si esta línea es la línea automática de descuento"
    )

    price_subtotal = fields.Monetary(
        string='Monto Antes de Impuestos',
        currency_field='currency_id',
        help="Subtotal de la factura original (antes de impuestos) para cálculo correcto de descuentos"
    )


    @api.onchange('account_move_line_id')
    def _onchange_account_move_line_id(self):
        if self.account_move_line_id:
            line = self.account_move_line_id
            self.partner_id = line.partner_id.id
            self.account_id = line.account_id.id
            self.move_id = line.move_id.id
            self.move_name = line.move_id.name
            self.date = line.date
            self.credit = abs(line.amount_residual) if line.amount_residual > 0 else 0.0
            self.debit = abs(line.amount_residual) if line.amount_residual < 0 else 0.0
            self.amount_residual = line.amount_residual
            self.currency_id = line.currency_id.id or self.receipt_id.currency_id.id
            self.analytic_account_id = line.analytic_account_id.id or self.receipt_id.analytic_account_id.id
            
            # 🔥 CALCULAR PRICE_SUBTOTAL PROPORCIONALMENTE
            if line.move_id and line.move_id.move_type in ['in_invoice', 'in_refund']:
                # Para facturas, calcular proporción del subtotal
                amount_untaxed = line.move_id.amount_untaxed or 0.0
                amount_total = line.move_id.amount_total or 0.0
                amount_paying = abs(line.amount_residual)
                
                if amount_total > 0:
                    # Calcular proporción del subtotal según el monto que se está pagando
                    proportion = amount_paying / amount_total
                    self.price_subtotal = amount_untaxed * proportion
                else:
                    self.price_subtotal = amount_untaxed
            else:
                # Para otros tipos de documento, usar el balance como subtotal
                self.price_subtotal = abs(line.amount_residual)
            
            # 🎯 APLICAR LÓGICA DE DESCUENTO AUTOMÁTICO
            self._apply_automatic_discount()
            
            # ============= DESCRIPCIÓN DIFERENCIADA =============
            is_advance = line.account_id.advance_account_supplier if line.account_id else False
            
            if line.move_name:
                self.description = f"{'APLICAR ANTICIPO' if is_advance else 'PAGO FACTURA PROVEEDOR'} - {line.move_name}"
            else:
                self.description = 'APLICAR ANTICIPO' if is_advance else 'PAGO FACTURA PROVEEDOR'
            
            self.cuentas_por_pagar = True if line.account_id.cuentas_por_pagar else False

    def _apply_automatic_discount(self):
        """Aplicar descuento automático basado en el plazo de pago del proveedor - MISMA LÓGICA QUE action_load_invoices"""
        # Solo aplicar si apply_discount está activo en el recibo
        if not self.receipt_id.apply_discount:
            self.discount = 0.0
            self.value_discount = 0.0
            return
        
        # Solo aplicar a facturas, no a anticipos
        if not self.account_move_line_id or not self.account_move_line_id.move_id:
            self.discount = 0.0
            self.value_discount = 0.0
            return
        
        # Verificar si es anticipo
        is_advance = self.account_id.advance_account_supplier if self.account_id else False
        if is_advance:
            self.discount = 0.0
            self.value_discount = 0.0
            return
        
        # 🎯 OBTENER PLAZO DE PAGO DEL PROVEEDOR (IGUAL QUE EN action_load_invoices)
        partner_payment_term = self.partner_id.property_supplier_payment_term_id
        
        if not (partner_payment_term and partner_payment_term.early_discount and partner_payment_term.discount_days > 0):
            self.discount = 0.0
            self.value_discount = 0.0
            return
        
        # Obtener fecha de la factura
        invoice_date = self.account_move_line_id.date
        if not invoice_date:
            self.discount = 0.0
            self.value_discount = 0.0
            return
        
        # Convertir string a date si es necesario
        if isinstance(invoice_date, str):
            from datetime import datetime
            invoice_date = datetime.strptime(invoice_date, '%Y-%m-%d').date()
        
        try:
            # Calcular fecha límite para descuento desde la fecha de la factura
            current_date = self.receipt_id.date or fields.Date.context_today(self)
            discount_date = partner_payment_term._get_last_discount_date(invoice_date)
            
            # Si estamos dentro del plazo de descuento
            if current_date <= discount_date:
                self.discount = partner_payment_term.discount_percentage
                
                # 🔥 CALCULAR DESCUENTO SOBRE EL SUBTOTAL PROPORCIONADO (IGUAL QUE EN action_load_invoices)
                move = self.account_move_line_id.move_id
                if move.move_type in ['in_invoice', 'in_refund'] and move.amount_total > 0:
                    # Calcular proporción del subtotal según el balance de esta línea
                    proportion = abs(self.balance) / move.amount_total
                    base_for_discount = move.amount_untaxed * proportion
                else:
                    # Para otros tipos de documento, usar el price_subtotal calculado
                    base_for_discount = self.price_subtotal
                
                self.value_discount = base_for_discount * (partner_payment_term.discount_percentage / 100.0)
            else:
                # Fuera del plazo de descuento
                self.discount = 0.0
                self.value_discount = 0.0
                
        except Exception as e:
            # Si hay error en el cálculo, no aplicar descuento
            self.discount = 0.0
            self.value_discount = 0.0
    
    # @api.onchange('account_move_line_id')
    # def _onchange_account_move_line_id(self):
    #     if self.account_move_line_id:
    #         line = self.account_move_line_id
    #         self.partner_id = line.partner_id.id
    #         self.account_id = line.account_id.id
    #         self.move_id = line.move_id.id
    #         self.move_name = line.move_id.name
    #         self.date = line.date
    #         self.credit = abs(line.amount_residual) if line.amount_residual > 0 else 0.0
    #         self.debit = abs(line.amount_residual) if line.amount_residual < 0 else 0.0
    #         self.amount_residual = line.amount_residual
    #         self.currency_id = line.currency_id.id or self.receipt_id.currency_id.id
    #         self.analytic_account_id = line.analytic_account_id.id or self.receipt_id.analytic_account_id.id
    #         self.price_subtotal = line.move_id.amount_untaxed or 0.0
    #         # ============= DESCRIPCIÓN DIFERENCIADA =============
    #         is_advance = line.account_id.advance_account_supplier if line.account_id else False
            
    #         if line.move_name:
    #             self.description = f"{'APLICAR ANTICIPO' if is_advance else 'PAGO FACTURA PROVEEDOR'} - {line.move_name}"
    #         else:
    #             self.description = 'APLICAR ANTICIPO' if is_advance else 'PAGO FACTURA PROVEEDOR'
            
    #         self.cuentas_por_pagar = True if line.account_id.cuentas_por_pagar else False

    @api.depends('debit', 'credit')
    def _calculate_balance(self):
        """Calcular balance cuando cambian débito o crédito"""
        for record in self:
            record.balance = record.debit - record.credit 
            
    @api.onchange('account_move_line_id', 'partner_id')
    def _onchange_check_partner(self):
        """Advertencia inmediata si el partner no coincide"""
        if self.partner_id and self.account_move_line_id:
            if self.partner_id != self.account_move_line_id.partner_id:
                raise ValidationError(_(
                    'El Proveedor seleccionado "%s" no coincide con el Proveedor del documento "%s".\n\n'
                    'Por favor, seleccione un Proveedor que coincida con el documento.'
                ) % (self.partner_id.name, self.account_move_line_id.move_id.name))
            
    @api.onchange('debit')
    def _onchange_debit(self):
        """Cuando se ingresa débito, limpiar crédito automáticamente"""
        if self.debit > 0:
            self.credit = 0.0

    @api.onchange('credit') 
    def _onchange_credit(self):
        """Cuando se ingresa crédito, limpiar débito automáticamente"""
        if self.credit > 0:
            self.debit = 0.0

    @api.constrains('debit', 'credit')
    def _check_debit_credit_exclusive(self):
        """Validar que no se ingresen valores en débito y crédito al mismo tiempo"""
        for line in self:
            if line.debit > 0 and line.credit > 0:
                raise ValidationError(_(
                    'Error en línea "%s":\n\n'
                    'No puede ingresar valores en Débito ($%s) y Crédito ($%s) simultáneamente.\n\n'
                    'Solución:\n'
                    '• Deje uno de los campos en $0.00\n'
                    '• Ingrese el valor completo en el campo correspondiente'
                    % (
                        line.description or line.move_name or 'Sin descripción',
                        '{:,.2f}'.format(line.debit), 
                        '{:,.2f}'.format(line.credit)
                    )
                ))        

    @api.onchange('account_id','balance')
    def onchange_tax_base_amount(self):
        for record in self:
            if record.account_id.required_account_vat:
                if record.account_id.percentage == 0.00:
                    raise ValidationError(_('La cuenta %s no tiene el porcentaje configurado en la cuenta'))
                elif record.account_id.percentage < 0.00:
                    raise ValidationError(_('La cuenta %s tiene porcentaje configurado en la cuenta como negativo, no puede ser de esa forma'))
                else:
                    importe = abs(record.balance)
                    record.tax_base_amount = importe / record.account_id.percentage                

    @api.depends('partner_id', 'receipt_id.reconcilable_accounting_entries')
    def _get_domain(self):
        # Limpiar registros sin partner_id
        records_without_partner = self.filtered(lambda r: not r.partner_id)
        for record in records_without_partner:
            record.suitable_invoice_line_ids = [(6, 0, [])]
        
        # Procesar solo registros con partner_id
        records_with_partner = self - records_without_partner
        if not records_with_partner:
            return
        
        # Agrupar por tipo de lógica
        reconcilable_records = records_with_partner.filtered(lambda r: r.receipt_id.reconcilable_accounting_entries)
        normal_records = records_with_partner - reconcilable_records
        
        # Procesar registros reconcilables en lote
        if reconcilable_records:
            partner_ids = reconcilable_records.mapped('partner_id.id')
            
            domain = [
                ('partner_id', 'in', partner_ids),
                ('move_id.move_type', '=', 'entry'),
                ('parent_state', '=', 'posted'),
                ('account_id.reconcile', '=', True),
                # ('payment_id', '=', False),
                ('amount_residual', '!=', 0.00)
            ]
            
            lines = self.env['account.move.line'].search(domain, order="date_maturity asc")
            lines_by_partner = {}
            for line in lines:
                if line.partner_id.id not in lines_by_partner:
                    lines_by_partner[line.partner_id.id] = []
                lines_by_partner[line.partner_id.id].append(line.id)
            
            for record in reconcilable_records:
                suitable_ids = lines_by_partner.get(record.partner_id.id, [])
                record.suitable_invoice_line_ids = [(6, 0, suitable_ids)]
        
        # Procesar registros normales en lote
        if normal_records:
            company_ids = normal_records.mapped('company_id.id')
            partner_ids = normal_records.mapped('partner_id.id')
            
            # Una sola consulta para cuentas por pagar
            payable_accounts = self.env['account.account'].search([
                ('cuentas_por_pagar', '=', True),
                ('reconcile', '=', True),
                ('company_id', 'in', company_ids)
            ])
            
            if payable_accounts:
                saldos_journals = self.env['account.journal'].search([('saldos_iniciales', '=', True)])
                
                # ============= CONSULTA 1: FACTURAS =============
                domain_facturas = [
                    ('partner_id', 'in', partner_ids),
                    ('company_id', 'in', company_ids),
                    ('parent_state', '=', 'posted'),
                    ('amount_residual', '<', 0),
                    ('account_id', 'in', payable_accounts.ids),
                    '|',
                    ('parent_move_type', 'in', ['in_invoice','in_refund','entry']),
                    '&',
                    ('parent_move_type', '=', 'entry'),
                    ('journal_id', 'in', saldos_journals.ids)
                ]
                
                lines_facturas = self.env['account.move.line'].search(domain_facturas, order="date_maturity asc")
                
                # ============= CONSULTA 2: ANTICIPOS =============
                domain_anticipos = [
                    ('partner_id', 'in', partner_ids),
                    ('company_id', 'in', company_ids),
                    ('parent_state', '=', 'posted'),
                    ('account_id.advance_account_supplier', '=', True),
                    ('amount_residual', '>', 0.00)
                ]
                
                lines_anticipos = self.env['account.move.line'].search(domain_anticipos, order="date_maturity asc")
                
                # ============= COMBINAR RESULTADOS =============
                # Unir ambas consultas
                all_lines = lines_facturas | lines_anticipos
                
                lines_by_partner_company = {}
                for line in all_lines:
                    key = (line.partner_id.id, line.company_id.id)
                    if key not in lines_by_partner_company:
                        lines_by_partner_company[key] = []
                    lines_by_partner_company[key].append(line.id)
                
                for record in normal_records:
                    key = (record.partner_id.id, record.company_id.id)
                    suitable_ids = lines_by_partner_company.get(key, [])
                    record.suitable_invoice_line_ids = [(6, 0, suitable_ids)]
            else:
                for record in normal_records:
                    record.suitable_invoice_line_ids = [(6, 0, [])]
                    
    @api.onchange('discount')
    def _onchange_discount(self):
        """Calcular value_discount cuando cambia el porcentaje de descuento - APLICADO SOBRE SUBTOTAL PROPORCIONADO"""
        if not self.receipt_id.apply_discount:
            self.discount = 0.0
            return
            
        for line in self:
            if line.discount:
                # Calcular base para descuento usando proporción del subtotal de la factura
                base_for_discount = 0.0
                
                if line.account_move_line_id and line.account_move_line_id.move_id:
                    move = line.account_move_line_id.move_id
                    
                    if move.move_type in ['in_invoice', 'in_refund'] and move.amount_total > 0:
                        # Calcular proporción del subtotal según el balance de esta línea
                        proportion = abs(line.balance) / move.amount_total
                        base_for_discount = move.amount_untaxed * proportion
                    else:
                        # Para otros tipos de documento, usar el balance
                        base_for_discount = abs(line.balance)
                else:
                    # Si no hay documento asociado, usar el balance
                    base_for_discount = abs(line.balance)
                
                line.value_discount = base_for_discount * line.discount / 100
            elif not line.discount:
                line.value_discount = 0.0

    @api.onchange('value_discount')
    def _onchange_value_discount(self):
        """Calcular discount cuando cambia el valor del descuento - BASADO EN SUBTOTAL PROPORCIONADO"""
        if not self.receipt_id.apply_discount:
            self.value_discount = 0.0
            return
            
        for line in self:
            if line.value_discount:
                # Calcular base para descuento usando proporción del subtotal de la factura
                base_for_discount = 0.0
                
                if line.account_move_line_id and line.account_move_line_id.move_id:
                    move = line.account_move_line_id.move_id
                    
                    if move.move_type in ['in_invoice', 'in_refund'] and move.amount_total > 0:
                        # Calcular proporción del subtotal según el balance de esta línea
                        proportion = abs(line.balance) / move.amount_total
                        base_for_discount = move.amount_untaxed * proportion
                    else:
                        # Para otros tipos de documento, usar el balance
                        base_for_discount = abs(line.balance)
                else:
                    # Si no hay documento asociado, usar el balance
                    base_for_discount = abs(line.balance)
                
                if base_for_discount > 0:
                    line.discount = (line.value_discount * 100) / base_for_discount
                else:
                    line.discount = 0.0
            elif not line.value_discount:
                line.discount = 0.0

    @api.onchange('balance')
    def _onchange_balance(self):
        """Recalcular value_discount cuando cambia el balance (si hay discount establecido)"""
        for line in self:
            if line.discount and line.balance:
                line.value_discount = abs(line.balance * line.discount / 100)             

    @api.constrains('value_discount', 'discount')
    def _check_discount_with_apply_discount(self):
        """Validar que solo se puedan usar descuentos si apply_discount está activo"""
        for line in self:
            if (line.value_discount > 0 or line.discount > 0) and not line.receipt_id.apply_discount:
                raise ValidationError(_(
                    'No puede aplicar descuentos en la línea "%s" porque el campo '
                    '"Aplicar Descuento" no está activo en el recibo.\n\n'
                    'Active primero "Aplicar Descuento" en el recibo.'
                ) % (line.description or line.move_name or 'Sin descripción'))

    @api.onchange('value_discount', 'discount')
    def _onchange_discount_fields(self):
        """Advertir si se intenta poner descuento sin activar apply_discount"""
        if (self.value_discount > 0 or self.discount > 0) and not self.receipt_id.apply_discount:
            return {
                'warning': {
                    'title': 'Descuento no permitido',
                    'message': 'Para aplicar descuentos, primero active "Aplicar Descuento" en el recibo.'
                }
            }
        
    @api.onchange('is_discount_line')
    def _onchange_is_discount_line(self):
        """Cuando se marca como línea de descuento, configurar automáticamente"""
        if self.is_discount_line:
            self.account_id = 1941
            self.debit = 0.0
            self.sequence = 999
            self.selected = False
            self.cuentas_por_pagar = False
            if self.receipt_id.partner_id:
                self.description = f'DESCUENTO APLICADO - {self.receipt_id.partner_id.name}'                