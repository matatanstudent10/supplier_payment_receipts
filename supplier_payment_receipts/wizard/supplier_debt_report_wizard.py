# -*- coding: utf-8 -*-

import base64
import io
from datetime import datetime
from odoo import models, fields, api, _
from odoo.exceptions import UserError
import xlsxwriter


class SupplierDebtReportWizard(models.TransientModel):
    _name = 'supplier.debt.report.wizard'
    _description = 'Reporte de Deudas con Proveedores'

    partner_id = fields.Many2one(
        'res.partner',
        string='Proveedor',
        domain=[('supplier_rank', '>', 0)],
        help='Seleccionar proveedor específico o dejar vacío para todos'
    )
    date_to = fields.Date(
        string='Fecha de corte',
        required=True,
        default=fields.Date.context_today,
        help='Facturas pendientes hasta esta fecha'
    )
    company_id = fields.Many2one(
        'res.company',
        string='Compañía',
        required=True,
        default=lambda self: self.env.company
    )
    excel_file = fields.Binary(
        string='Archivo Excel'
    )
    file_name = fields.Char(
        string='Nombre del archivo'
    )
    state = fields.Selection([
        ('draft', 'Borrador'),
        ('done', 'Completado')
    ], default='draft')

    def generate_report(self):
        """Genera el reporte de deudas pendientes"""
        self.ensure_one()
        
        # Crear el archivo Excel
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('Deudas Proveedores')
        
        # Formatos
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        currency_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'align': 'right'
        })
        
        text_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'align': 'center'
        })
        
        # Encabezados
        headers = [
            'Proveedor',
            'Factura/Documento',
            'Fecha Factura',
            'Fecha Vencimiento',
            'Referencia',
            'Cuenta Contable',
            'Valor Factura',
            'Total Pagado',
            'Saldo Pendiente',
            'Días Pendientes'
        ]
        
        # Escribir encabezados
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Configurar anchos de columna
        worksheet.set_column(0, 0, 25)  # Proveedor
        worksheet.set_column(1, 1, 18)  # Factura
        worksheet.set_column(2, 3, 12)  # Fechas
        worksheet.set_column(4, 4, 15)  # Referencia
        worksheet.set_column(5, 5, 20)  # Cuenta
        worksheet.set_column(6, 9, 12)  # Montos
        
        # Obtener datos
        data = self._get_supplier_debts()
        
        row = 1
        total_debt = 0
        
        for record in data:
            worksheet.write(row, 0, record['partner_name'], text_format)
            worksheet.write(row, 1, record['move_name'], text_format)
            worksheet.write(row, 2, record['invoice_date'], date_format)
            worksheet.write(row, 3, record['date_maturity'], date_format)
            worksheet.write(row, 4, record['ref'] or '', text_format)
            worksheet.write(row, 5, record['account_name'], text_format)
            worksheet.write(row, 6, record['valor_factura'], currency_format)
            worksheet.write(row, 7, record['partial_payments'], currency_format)
            worksheet.write(row, 8, record['balance'], currency_format)
            worksheet.write(row, 9, record['dias_pendientes'], text_format)
            
            total_debt += record['balance']
            row += 1
        
        # Totales
        if data:
            worksheet.write(row + 1, 7, 'TOTAL DEUDA:', header_format)
            worksheet.write(row + 1, 8, total_debt, currency_format)
        
        workbook.close()
        output.seek(0)
        
        # Generar nombre del archivo
        date_str = self.date_to.strftime('%Y%m%d')
        partner_str = self.partner_id.name if self.partner_id else 'Todos_Proveedores'
        file_name = f'Deudas_Proveedores_{partner_str}_{date_str}.xlsx'
        
        # Guardar archivo
        self.write({
            'excel_file': base64.b64encode(output.getvalue()),
            'file_name': file_name,
            'state': 'done'
        })
        
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'supplier.debt.report.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }

    def _get_supplier_debts(self):
        """Obtiene las deudas pendientes con proveedores"""
        
        # Buscar líneas contables en cuentas por pagar (con crédito, como en el SQL)
        domain = [
            ('account_id.cuentas_por_pagar', '=', True),
            ('move_id.state', '=', 'posted'),
            ('date', '<=', self.date_to),  # Facturas hasta la fecha de corte
            ('company_id', '=', self.company_id.id),
            ('credit', '>', 0),  # Líneas de crédito en cuentas por pagar (como debit = 0.00 en SQL)
        ]
        
        if self.partner_id:
            domain.append(('partner_id', '=', self.partner_id.id))
        else:
            domain.append(('partner_id.supplier_rank', '>', 0))
        
        move_lines = self.env['account.move.line'].search(domain, order='partner_id, date, move_id')
        
        result = []
        
        for line in move_lines:
            # Balance original (en SQL usa aml.balance)
            original_balance = line.balance  # Esto será negativo para deudas
            total_pagado = 0
            
            # Buscar pagos realizados hasta la fecha de corte
            # Adaptando la lógica del LEFT JOIN con apr
            reconciles = self.env['account.partial.reconcile'].search([
                ('credit_move_id', '=', line.id),  # Esta línea es la que recibe el pago
                ('max_date', '<=', self.date_to)   # Solo pagos hasta la fecha de corte
            ])
            
            # Sumar los pagos (como SUM(apr.amount) en el SQL)
            for reconcile in reconciles:
                total_pagado += reconcile.amount
            
            # Calcular saldo pendiente (como en SQL: aml.balance + COALESCE(apr.amount,0))
            saldo_pendiente = original_balance + total_pagado
            
            # Solo incluir si tiene saldo pendiente (WHERE "TOTAL PENDIENTE" != 0.00)
            if abs(saldo_pendiente) > 0.01:
                # Calcular días pendientes
                dias_pendientes = (self.date_to - line.date).days
                
                result.append({
                    'partner_name': line.partner_id.name,
                    'move_name': line.move_id.name,
                    'invoice_date': line.date,
                    'date_maturity': line.date_maturity or line.date,
                    'ref': line.move_id.ref or line.ref or '',
                    'account_name': line.account_id.name,
                    'debit': 0.00,  # Como en SQL debit = 0.00
                    'credit': line.credit,
                    'balance': abs(saldo_pendiente),  # Mostrar como positivo
                    'partial_payments': total_pagado,
                    'dias_pendientes': dias_pendientes,
                    'valor_factura': abs(original_balance),
                })
        
        return result

    def download_excel(self):
        """Descargar archivo Excel"""
        self.ensure_one()
        if not self.excel_file:
            raise UserError(_('No hay archivo para descargar. Genere el reporte primero.'))
        
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=supplier.debt.report.wizard&id={self.id}&field=excel_file&download=true&filename={self.file_name}',
            'target': 'self',
        }

    def back_to_draft(self):
        """Volver a borrador"""
        self.write({
            'excel_file': False,
            'file_name': False,
            'state': 'draft'
        })