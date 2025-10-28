# -*- coding: utf-8 -*-

import base64
import io
import logging
from collections import defaultdict
from datetime import datetime, date, timedelta
from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError
from odoo.tools import format_date

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

_logger = logging.getLogger(__name__)



class WizardDiscountReport(models.TransientModel):
    _name = 'wizard.discount.report'
    _description = 'Asistente para Informe de Descuentos y Creación Masiva de Pagos'

    # ================== CAMPOS DE FILTRO ÚNICOS ==================
    partner_ids = fields.Many2many(
        'res.partner',
        'wizard_discount_partner_rel',
        'wizard_id',
        'partner_id',
        string='Proveedores',
        domain=[('supplier_rank', '>', 0)],
        help="Dejar vacío para incluir todos los proveedores"
    )
    
    company_id = fields.Many2one(
        'res.company',
        string='Compañía',
        required=True,
        default=lambda self: self.env.company
    )
    
    date_from = fields.Date(
        string='Desde Fecha',
        help="Filtrar facturas desde esta fecha"
    )
    
    date_to = fields.Date(
        string='Hasta Fecha',
        help="Filtrar facturas hasta esta fecha"
    )
    
    only_with_discount = fields.Boolean(
        string='Solo con Descuento Disponible',
        default=True,
        help="Mostrar solo facturas con descuento por pago anticipado vigente"
    )
    
    payment_journal_id = fields.Many2one(
        'account.journal',
        string='Diario de Pago',
        domain=[('type', 'in', ['bank', 'cash'])],
        help="Diario para crear los pagos automáticos"
    )
    
    # ================== CAMPOS DE RESULTADO MÍNIMOS ==================
    excel_file = fields.Binary(
        string='Archivo Excel',
        readonly=True
    )
    
    excel_filename = fields.Char(
        string='Nombre del Archivo',
        readonly=True
    )
    
    payment_ids = fields.Many2many(
        'supplier.payment.receipt',
        'wizard_discount_payment_rel',
        'wizard_id',
        'payment_id',
        string='Pagos Creados',
        readonly=True
    )

    def _get_invoices_data_python(self):
        """Obtener datos de facturas usando la misma lógica Python que action_load_invoices - VERSIÓN OPTIMIZADA"""
        
        # Buscar cuentas por cobrar (payable accounts)
        payable_accounts = self.env['account.account'].search([
            ('cuentas_por_pagar', '=', True),
            ('reconcile', '=', True),
            ('company_id', '=', self.company_id.id)
        ])
        
        if not payable_accounts:
            raise UserError(_('No se encontraron cuentas por cobrar configuradas.'))
        
        # FILTRO DE DIARIOS DE COMPRA CON NOMBRE ESPECÍFICO
        purchase_journals = self.env['account.journal'].search([
            ('type', '=', 'purchase'),
            ('name', 'ilike', 'FACTURAS DE PROVEEDORES'),
            ('company_id', '=', self.company_id.id)
        ])
        
        if not purchase_journals:
            raise UserError(_('No se encontraron diarios de compra que contengan "FACTURAS DE PROVEEDORES" en el nombre.'))
        
        current_date = fields.Date.context_today(self)
        
        # OPTIMIZACIÓN 1: Pre-filtrar proveedores que tengan descuento disponible
        partners_to_process = self.partner_ids if self.partner_ids else self.env['res.partner'].search([('supplier_rank', '>', 0)])
        
        if self.only_with_discount:
            # Solo procesar proveedores que tengan plazo de pago con descuento configurado
            partners_with_discount = []
            for partner in partners_to_process:
                payment_term = partner.property_supplier_payment_term_id
                if payment_term and payment_term.early_discount and payment_term.discount_days > 0:
                    partners_with_discount.append(partner)
            
            if not partners_with_discount:
                raise UserError(_('Ningún proveedor seleccionado tiene configurado descuento por pago anticipado.'))
            
            partners_to_process = partners_with_discount
        
        all_invoices_data = []
        
        # OPTIMIZACIÓN 2: Procesar en lotes para mejor rendimiento
        for partner in partners_to_process:
            # OBTENER PLAZO DE PAGO DEL PROVEEDOR
            partner_payment_term = partner.property_supplier_payment_term_id
            partner_payment_term_name = partner_payment_term.name if partner_payment_term else 'Sin plazo de pago'
            
            # OPTIMIZACIÓN 3: Si solo queremos con descuento y este proveedor no tiene, saltar
            if self.only_with_discount:
                if not (partner_payment_term and partner_payment_term.early_discount and partner_payment_term.discount_days > 0):
                    continue
            
            # DOMINIO PRINCIPAL
            domain = [
                ('company_id', '=', self.company_id.id),
                ('partner_id', '=', partner.id),
                ('parent_state', '=', 'posted'),
                ('amount_residual', '<', 0),
                ('account_id', 'in', payable_accounts.ids),
                ('journal_id', 'in', purchase_journals.ids),
                '|',
                ('parent_move_type', 'in', ['in_invoice','in_refund']),
                ('parent_move_type', 'in', ['in_invoice','entry']),
            ]
            
            # Filtros adicionales por fecha
            if self.date_from:
                domain.append(('move_id.invoice_date', '>=', self.date_from))
            if self.date_to:
                domain.append(('move_id.invoice_date', '<=', self.date_to))
            
            # OPTIMIZACIÓN 4: Si solo queremos con descuento, filtrar por fecha también
            if self.only_with_discount and partner_payment_term:
                # Calcular fecha mínima desde la cual una factura podría tener descuento vigente
                fecha_minima_descuento = current_date - timedelta(days=partner_payment_term.discount_days)
                domain.append(('date', '>=', fecha_minima_descuento))
            
            # CARGAR DATOS
            pending_lines_data = self.env['account.move.line'].search_read(
                domain,
                ['id', 'partner_id', 'account_id', 'move_id', 'date', 'date_maturity', 'amount_residual', 
                'balance', 'currency_id', 'analytic_account_id', 'store_id', 'journal_id'],
            )
            
            if not pending_lines_data:
                continue
            
            # OPTIMIZACIÓN 5: Una sola consulta para obtener todos los moves necesarios
            move_ids = [line['move_id'][0] for line in pending_lines_data if line['move_id']]
            moves_dict = {}
            
            if move_ids:
                moves_data = self.env['account.move'].search_read(
                    [('id', 'in', move_ids)],
                    ['id', 'ref', 'name', 'invoice_date', 'move_type', 'state', 'amount_untaxed', 'amount_total']  # 🔥 AGREGADOS
                )
                moves_dict = {move['id']: move for move in moves_data}
            
            # OPTIMIZACIÓN 6: Pre-calcular datos del proveedor una sola vez
            partner_info = partner
            
            # OPTIMIZACIÓN 7: Cache de objetos relacionados
            store_cache = {}
            journal_cache = {}
            account_cache = {}
            currency_cache = {}
            
            # PROCESAR LÍNEAS EN LOTE
            for line_data in pending_lines_data:
                # CALCULAR DESCUENTO BASADO EN EL PLAZO DE PAGO DEL PROVEEDOR
                line_discount = 0.0
                line_discount_value = 0.0
                estado_descuento = 'Sin descuento'
                fecha_limite_descuento = None
                dias_restantes_descuento = None
                plazo_de_pago = partner_payment_term_name
                
                # USAR PLAZO DE PAGO DEL PROVEEDOR
                if partner_payment_term and partner_payment_term.early_discount and partner_payment_term.discount_days > 0:
                    invoice_date = line_data.get('date')
                    if isinstance(invoice_date, str):
                        from datetime import datetime
                        invoice_date = datetime.strptime(invoice_date, '%Y-%m-%d').date()
                    
                    if invoice_date:
                        # Calcular fecha límite para descuento desde la fecha de la factura
                        fecha_limite_descuento = partner_payment_term._get_last_discount_date(invoice_date)
                        
                        if fecha_limite_descuento:
                            # Calcular días restantes
                            dias_restantes = (fecha_limite_descuento - current_date).days
                            dias_restantes_descuento = max(0, dias_restantes)
                            
                            # Si estamos dentro del plazo de descuento
                            if current_date <= fecha_limite_descuento:
                                line_discount = partner_payment_term.discount_percentage
                                
                                # 🔥 CALCULAR DESCUENTO SOBRE SUBTOTAL PROPORCIONADO
                                move_info = moves_dict.get(line_data['move_id'][0])
                                if move_info and move_info['move_type'] in ['in_invoice', 'in_refund']:
                                    # Para facturas, calcular proporción del subtotal
                                    amount_untaxed = move_info['amount_untaxed'] or 0.0
                                    amount_total = move_info['amount_total'] or 0.0
                                    amount_paying = abs(line_data['amount_residual'])
                                    
                                    if amount_total > 0:
                                        # Calcular proporción del subtotal según el monto que se está pagando
                                        proportion = amount_paying / amount_total
                                        base_for_discount = amount_untaxed * proportion
                                    else:
                                        base_for_discount = amount_untaxed
                                else:
                                    # Para otros tipos de documento, usar el balance como base
                                    base_for_discount = abs(line_data['amount_residual'])
                                
                                line_discount_value = base_for_discount * (partner_payment_term.discount_percentage / 100.0)
                                estado_descuento = 'Descuento disponible'
                            else:
                                estado_descuento = 'Descuento vencido'
                
                # OPTIMIZACIÓN 8: Filtrar temprano y saltar procesamiento innecesario
                if self.only_with_discount and estado_descuento != 'Descuento disponible':
                    continue
                
                # Obtener move info del diccionario
                move_info = moves_dict.get(line_data['move_id'][0])
                if not move_info:
                    continue
                
                # OPTIMIZACIÓN 9: Cache de objetos relacionados para evitar múltiples browse
                account_id = line_data['account_id'][0]
                if account_id not in account_cache:
                    account_cache[account_id] = self.env['account.account'].browse(account_id)
                account_info = account_cache[account_id]
                
                currency_id = line_data['currency_id'][0] if line_data['currency_id'] else self.env.company.currency_id.id
                if currency_id not in currency_cache:
                    currency_cache[currency_id] = self.env['res.currency'].browse(currency_id)
                currency_info = currency_cache[currency_id]
                
                # Store info (con cache)
                store_id = line_data['store_id'][0] if line_data['store_id'] else None
                if store_id and store_id not in store_cache:
                    store_cache[store_id] = self.env['res.store'].browse(store_id)
                store_info = store_cache.get(store_id)
                
                # Journal info (con cache)
                journal_id = line_data['journal_id'][0] if line_data['journal_id'] else None
                if journal_id and journal_id not in journal_cache:
                    journal_cache[journal_id] = self.env['account.journal'].browse(journal_id)
                journal_info = journal_cache.get(journal_id)
                
                # Calcular días vencidos
                dias_vencidos = 0
                if line_data['date_maturity']:
                    fecha_vencimiento = line_data['date_maturity']
                    if isinstance(fecha_vencimiento, str):
                        from datetime import datetime
                        fecha_vencimiento = datetime.strptime(fecha_vencimiento, '%Y-%m-%d').date()
                    
                    if fecha_vencimiento < current_date:
                        dias_vencidos = (current_date - fecha_vencimiento).days
                
                # Calcular valor neto con descuento
                valor_adeudado = abs(line_data['amount_residual'])
                valor_neto_con_descuento = valor_adeudado - line_discount_value if estado_descuento == 'Descuento disponible' else valor_adeudado
                
                # OBTENER FECHA DE FACTURA CORRECTAMENTE
                fecha_factura = line_data['date']
                if 'invoice_date' in locals() and invoice_date:
                    fecha_factura = invoice_date
                
                invoice_data = {
                    'proveedor': partner_info.name,
                    'nit_proveedor': partner_info.vat or partner_info.fe_nit or '',
                    'numero_factura': move_info['name'],
                    'referencia_factura': move_info['ref'] or '',
                    'fecha_factura': fecha_factura,
                    'fecha_vencimiento': line_data['date_maturity'],
                    'valor_adeudado': valor_adeudado,
                    'plazo_de_pago': plazo_de_pago,
                    'tiene_descuento_anticipado': bool(partner_payment_term and partner_payment_term.early_discount),
                    'porcentaje_descuento': line_discount,
                    'dias_descuento': partner_payment_term.discount_days if partner_payment_term else 0,
                    'valor_descuento': line_discount_value,
                    'fecha_limite_descuento': fecha_limite_descuento,
                    'dias_restantes_descuento': dias_restantes_descuento,
                    'estado_descuento': estado_descuento,
                    'cuenta_contable': f"{account_info.code} - {account_info.name}",
                    'tipo_documento': move_info['move_type'],
                    'moneda': currency_info.name,
                    'valor_neto_con_descuento': valor_neto_con_descuento,
                    'move_line_id': line_data['id'],
                    'partner_id': partner_info.id,
                    'move_id': move_info['id'],
                    'tienda': store_info.name if store_info else '',
                    'diario': journal_info.name if journal_info else '',
                    'dias_vencidos': dias_vencidos
                }
                
                all_invoices_data.append(invoice_data)
        
        return all_invoices_data

    def action_generate_excel(self):
        """Generar archivo Excel con el informe"""
        if not xlsxwriter:
            raise UserError(_('La librería xlsxwriter no está instalada. Por favor instálela: pip install xlsxwriter'))
        
        # Obtener datos usando el método Python
        invoices_data = self._get_invoices_data_python()
        
        if not invoices_data:
            raise UserError(_('No se encontraron facturas con los filtros seleccionados.'))
        
        # Crear archivo Excel
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        # Configurar formatos
        formats = self._setup_excel_formats(workbook)
        
        # Crear hoja principal
        worksheet = workbook.add_worksheet('Informe de Descuentos')
        self._write_excel_content(worksheet, invoices_data, formats)
        
        # Crear hoja de resumen
        summary_sheet = workbook.add_worksheet('Resumen por Proveedor')
        self._write_summary_sheet(summary_sheet, invoices_data, formats)
        
        workbook.close()
        output.seek(0)
        
        # Guardar archivo
        filename = f'Informe_Descuentos_Proveedores_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        self.write({
            'excel_file': base64.b64encode(output.getvalue()),
            'excel_filename': filename
        })


        return {
            'type': 'ir.actions.act_url',
            'url': f"web/content/?model={self._name}&id={self.id}&field=excel_file&download=true&filename={filename}",
            'target': 'self',
        }

    def _setup_excel_formats(self, workbook):
        """Configurar formatos para Excel"""
        return {
            'title': workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#1f4e79',
                'font_color': 'white'
            }),
            'header': workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#dbe5f1',
                'border': 1
            }),
            'data': workbook.add_format({
                'font_size': 10,
                'align': 'left',
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
            'date': workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'num_format': 'dd/mm/yyyy'
            }),
            'percent': workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '0.00%'
            }),
            'discount_available': workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#d4edda',
                'font_color': '#155724'
            }),
            'discount_expired': workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#f8d7da',
                'font_color': '#721c24'
            })
        }

    def _write_excel_content(self, worksheet, data, formats):
        """Escribir contenido principal del Excel"""
        # Configurar columnas (agregando 3 columnas más)
        worksheet.set_column('A:A', 25)  # Proveedor
        worksheet.set_column('B:B', 15)  # NIT
        worksheet.set_column('C:C', 15)  # Factura
        worksheet.set_column('D:D', 20)  # Referencia
        worksheet.set_column('E:E', 12)  # Fecha Factura
        worksheet.set_column('F:F', 12)  # Fecha Vencimiento
        worksheet.set_column('G:G', 15)  # Valor Adeudado
        worksheet.set_column('H:H', 20)  # Plazo de Pago
        worksheet.set_column('I:I', 10)  # % Descuento
        worksheet.set_column('J:J', 15)  # Valor Descuento
        worksheet.set_column('K:K', 12)  # Fecha Límite
        worksheet.set_column('L:L', 10)  # Días Restantes
        worksheet.set_column('M:M', 15)  # Estado
        worksheet.set_column('N:N', 15)  # Valor Neto
        worksheet.set_column('O:O', 20)  # Tienda (NUEVO)
        worksheet.set_column('P:P', 20)  # Diario (NUEVO)
        worksheet.set_column('Q:Q', 12)  # Días Vencidos (NUEVO)
        
        # Título (extender el merge range)
        worksheet.merge_range('A1:Q1', 'INFORME DE DESCUENTOS POR PAGO ANTICIPADO', formats['title'])
        
        # Encabezados (agregar los nuevos)
        headers = [
            'Proveedor', 'NIT/CC', 'Nº Factura', 'Referencia', 'Fecha Factura',
            'Fecha Vencimiento', 'Valor Adeudado', 'Plazo de Pago', '% Descuento',
            'Valor Descuento', 'Fecha Límite Descuento', 'Días Restantes',
            'Estado Descuento', 'Valor Neto con Descuento', 'Tienda', 'Diario', 'Días Vencidos'
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(2, col, header, formats['header'])
        
        # Datos (agregar las nuevas columnas)
        row = 3
        for invoice in data:
            worksheet.write(row, 0, invoice['proveedor'] or '', formats['data'])
            worksheet.write(row, 1, invoice['nit_proveedor'] or '', formats['data'])
            worksheet.write(row, 2, invoice['numero_factura'] or '', formats['data'])
            worksheet.write(row, 3, invoice['referencia_factura'] or '', formats['data'])
            worksheet.write(row, 4, invoice['fecha_factura'], formats['date'])
            worksheet.write(row, 5, invoice['fecha_vencimiento'], formats['date'])
            worksheet.write(row, 6, float(invoice['valor_adeudado']), formats['money'])
            worksheet.write(row, 7, invoice['plazo_de_pago'] or '', formats['data'])
            
            if invoice['porcentaje_descuento']:
                worksheet.write(row, 8, float(invoice['porcentaje_descuento'])/100, formats['percent'])
            else:
                worksheet.write(row, 8, 0, formats['percent'])
                
            worksheet.write(row, 9, float(invoice['valor_descuento']), formats['money'])
            worksheet.write(row, 10, invoice['fecha_limite_descuento'], formats['date'])
            worksheet.write(row, 11, invoice['dias_restantes_descuento'] or 0, formats['data'])
            
            # Estado con color
            state_format = formats['discount_available'] if invoice['estado_descuento'] == 'Descuento disponible' else formats['discount_expired']
            worksheet.write(row, 12, invoice['estado_descuento'], state_format)
            
            worksheet.write(row, 13, float(invoice['valor_neto_con_descuento']), formats['money'])
            
            # NUEVAS COLUMNAS:
            worksheet.write(row, 14, invoice['tienda'] or '', formats['data'])  # Tienda
            worksheet.write(row, 15, invoice['diario'] or '', formats['data'])  # Diario
            
            # Días vencidos con formato especial si está vencido
            dias_vencidos = invoice['dias_vencidos']
            if dias_vencidos > 0:
                # Formato rojo para facturas vencidas
                vencido_format = formats.get('discount_expired', formats['data'])
                worksheet.write(row, 16, dias_vencidos, vencido_format)
            else:
                worksheet.write(row, 16, 0, formats['data'])
            
            row += 1
        
    def _write_summary_sheet(self, worksheet, data, formats):
        """Escribir hoja de resumen por proveedor"""
        # Agrupar por proveedor
        summary = defaultdict(lambda: {
            'facturas': 0,
            'total_adeudado': 0,
            'total_descuento': 0,
            'total_neto': 0
        })
        
        for invoice in data:
            partner = invoice['proveedor']
            summary[partner]['facturas'] += 1
            summary[partner]['total_adeudado'] += float(invoice['valor_adeudado'])
            summary[partner]['total_descuento'] += float(invoice['valor_descuento'])
            summary[partner]['total_neto'] += float(invoice['valor_neto_con_descuento'])
        
        # Configurar columnas
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 20)
        
        # Título
        worksheet.merge_range('A1:E1', 'RESUMEN POR PROVEEDOR', formats['title'])
        
        # Encabezados
        headers = ['Proveedor', 'Facturas', 'Total Adeudado', 'Total Descuentos', 'Total Neto']
        for col, header in enumerate(headers):
            worksheet.write(2, col, header, formats['header'])
        
        # Datos
        row = 3
        for partner, data_summary in summary.items():
            worksheet.write(row, 0, partner, formats['data'])
            worksheet.write(row, 1, data_summary['facturas'], formats['data'])
            worksheet.write(row, 2, data_summary['total_adeudado'], formats['money'])
            worksheet.write(row, 3, data_summary['total_descuento'], formats['money'])
            worksheet.write(row, 4, data_summary['total_neto'], formats['money'])
            row += 1

    def action_create_payments(self):
        """Crear pagos masivos por proveedor usando action_load_invoices"""
        if not self.payment_journal_id:
            raise UserError(_('Debe seleccionar un diario de pago para crear los pagos automáticos.'))
        
        # Determinar qué proveedores procesar basado en el filtro only_with_discount
        if self.only_with_discount:
            # Si solo queremos con descuento, usar la lógica existente optimizada
            invoices_data = self._get_invoices_data_python()
            
            # Filtrar solo facturas con descuento disponible
            available_discounts = [inv for inv in invoices_data if inv['estado_descuento'] == 'Descuento disponible']
            
            if not available_discounts:
                raise UserError(_('No se encontraron facturas con descuentos disponibles para crear pagos.'))
            
            # Obtener proveedores únicos que tienen descuentos disponibles
            partners_with_discounts = list(set([inv['partner_id'] for inv in available_discounts]))
            partners_to_process = self.env['res.partner'].browse(partners_with_discounts)
            
            
        else:
            # Si no hay filtro de descuento, procesar todos los proveedores seleccionados o todos los que tengan facturas
            if self.partner_ids:
                partners_to_process = self.partner_ids
            else:
                # Buscar proveedores que tengan facturas pendientes en los diarios configurados
                partners_to_process = self._get_partners_with_pending_invoices()
            
        
        if not partners_to_process:
            raise UserError(_('No se encontraron proveedores para procesar con los filtros seleccionados.'))
        
        created_payments = []
        errors = []
        
        for partner in partners_to_process:
            try:
                
                # Crear recibo de pago básico
                payment_receipt = self._create_payment_receipt_header(partner)
                
                # Usar action_load_invoices para cargar facturas automáticamente
                payment_receipt.action_load_invoices()
                payment_receipt._onchange_analytic_account()
                
                # Verificar si se cargaron líneas
                if not payment_receipt.line_ids:
                    payment_receipt.unlink()
                    continue
                
                created_payments.append(payment_receipt.id)

                
            except Exception as e:
                error_msg = f"Error creando pago para {partner.name}: {str(e)}"
                errors.append(error_msg)
                continue
        
        if not created_payments and errors:
            raise UserError(_('No se pudieron crear pagos. Errores:\n' + '\n'.join(errors)))
        
        # Actualizar wizard con pagos creados
        self.write({
            'payment_ids': [(6, 0, created_payments)]
        })

        return {
            'type': 'ir.actions.act_window',
            'name': 'Pagos Creados',
            'res_model': 'supplier.payment.receipt',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', created_payments)],
            'context': {
                'create': False,
                'edit': True,
            }
        }


    def _get_partners_with_pending_invoices(self):
        """Obtener proveedores que tienen facturas pendientes en los diarios configurados"""
        
        # Buscar cuentas por cobrar (payable accounts)
        payable_accounts = self.env['account.account'].search([
            ('cuentas_por_pagar', '=', True),
            ('reconcile', '=', True),
            ('company_id', '=', self.company_id.id)
        ])
        
        if not payable_accounts:
            return self.env['res.partner']
        
        # FILTRO DE DIARIOS DE COMPRA CON NOMBRE ESPECÍFICO
        purchase_journals = self.env['account.journal'].search([
            ('type', '=', 'purchase'),
            ('name', 'ilike', 'FACTURAS DE PROVEEDORES'),
            ('company_id', '=', self.company_id.id)
        ])
        
        if not purchase_journals:
            return self.env['res.partner']
        
        # DOMINIO PARA BUSCAR FACTURAS PENDIENTES
        domain = [
            ('company_id', '=', self.company_id.id),
            ('parent_state', '=', 'posted'),
            ('amount_residual', '<', 0),
            ('account_id', 'in', payable_accounts.ids),
            ('journal_id', 'in', purchase_journals.ids),
            '|',
            ('parent_move_type', 'in', ['in_invoice','in_refund']),
            ('parent_move_type', 'in', ['in_invoice','entry']),
        ]
        
        # Filtros adicionales por fecha
        if self.date_from:
            domain.append(('move_id.invoice_date', '>=', self.date_from))
        if self.date_to:
            domain.append(('move_id.invoice_date', '<=', self.date_to))
        
        # Filtrar por proveedores seleccionados si los hay
        if self.partner_ids:
            domain.append(('partner_id', 'in', self.partner_ids.ids))
        else:
            domain.append(('partner_id.supplier_rank', '>', 0))
        
        # Buscar proveedores únicos que tienen facturas pendientes
        partner_ids = self.env['account.move.line'].search_read(domain, ['partner_id'])
        unique_partner_ids = list(set([line['partner_id'][0] for line in partner_ids if line['partner_id']]))
        
        return self.env['res.partner'].browse(unique_partner_ids)

    def _create_payment_receipt_header(self, partner):
        """Crear solo el encabezado del recibo de pago - action_load_invoices cargará las líneas"""
        
        # Buscar tienda por defecto
        store = self.env['res.store'].search([('name', '=', '001 MEDELLIN')], limit=1)
        if not store:
            store = self.env['res.store'].search([('company_id', '=', self.company_id.id)], limit=1)
        
        # Crear recibo de pago básico
        receipt_vals = {
            'partner_id': partner.id,
            'journal_id': self.payment_journal_id.id,
            'currency_id': self.company_id.currency_id.id,
            'company_id': self.company_id.id,
            'store_id': store.id if store else False,
            'amount_total': 0.0,  # Se calculará automáticamente
            'is_multiple_payment': True,
            'memo': f'Pago automático desde wizard - {datetime.now().strftime("%Y-%m-%d %H:%M")}',
            'destination_account_id': partner.property_account_payable_id.id,
            'date': fields.Date.context_today(self),
        }
        
        # Determinar método de pago por defecto
        payment_methods = self.payment_journal_id._get_available_payment_method_lines('outbound')
        if payment_methods:
            cheques_method = payment_methods.filtered(lambda x: 'cheques' in x.name.lower())
            receipt_vals['payment_method_line_id'] = cheques_method[0].id if cheques_method else payment_methods[0].id
        
        # Crear el recibo (sin líneas aún)
        receipt = self.env['supplier.payment.receipt'].create(receipt_vals)
    
        return receipt
 
    def action_view_payments(self):
        """Ver pagos creados"""
        if not self.payment_ids:
            raise UserError(_('No hay pagos creados para mostrar.'))
        
        return {
            'type': 'ir.actions.act_window',
            'name': 'Pagos con Descuentos Creados',
            'res_model': 'supplier.payment.receipt',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', self.payment_ids.ids)]
        }