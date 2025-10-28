# -*- coding: utf-8 -*-

from odoo import api, fields, models, _, Command
from odoo.exceptions import UserError, ValidationError
from odoo.tools import format_date, formatLang, frozendict, date_utils
from odoo.tools.float_utils import float_round

from dateutil.relativedelta import relativedelta


class AccountPaymentTerm(models.Model):
    _inherit = "account.payment.term"

    def _default_line_ids(self):
        return [Command.create({'value': 'percent', 'value_amount': 100.0, 'nb_days': 0})]

    def _default_example_date(self):
        return self._context.get('example_date') or fields.Date.today()

    # Redefine line_ids with new default
    line_ids = fields.One2many('account.payment.term.line', 'payment_id', string='Terms', copy=True, default=_default_line_ids)
    
    # New fields from v17
    fiscal_country_codes = fields.Char(compute='_compute_fiscal_country_codes')
    currency_id = fields.Many2one('res.currency', compute="_compute_currency_id")
    
    display_on_invoice = fields.Boolean(string='Mostrar fechas de cuotas', default=True)
    example_amount = fields.Monetary(currency_field='currency_id', default=1000, store=False, readonly=True)
    example_date = fields.Date(string='Fecha de ejemplo', default=_default_example_date, store=False)
    example_invalid = fields.Boolean(compute='_compute_example_invalid')
    example_preview = fields.Html(compute='_compute_example_preview')
    example_preview_discount = fields.Html(compute='_compute_example_preview')
    
    discount_percentage = fields.Float(string='% Descuento', help='Descuento por pago anticipado otorgado para este plazo de pago', default=2.0)
    discount_days = fields.Integer(string='Días descuento', help='Número de días antes de que expire la propuesta de pago anticipado', default=10)
    early_pay_discount_computation = fields.Selection([
        ('included', 'En pago anticipado'),
        ('excluded', 'Nunca'),
        ('mixed', 'Siempre (en factura)'),
    ], string='Reducción de impuestos por descuento', readonly=False, store=True, compute='_compute_discount_computation')
    early_discount = fields.Boolean(string='Descuento anticipado')

    @api.depends('company_id')
    @api.depends_context('allowed_company_ids')
    def _compute_fiscal_country_codes(self):
        for record in self:
            allowed_companies = record.company_id or self.env.companies
            record.fiscal_country_codes = ",".join(allowed_companies.mapped('account_fiscal_country_id.code'))

    @api.depends_context('company')
    @api.depends('company_id')
    def _compute_currency_id(self):
        for payment_term in self:
            payment_term.currency_id = payment_term.company_id.currency_id or self.env.company.currency_id

    def _get_amount_due_after_discount(self, total_amount, untaxed_amount):
        self.ensure_one()
        if self.early_discount:
            percentage = self.discount_percentage / 100.0
            if self.early_pay_discount_computation in ('excluded', 'mixed'):
                discount_amount_currency = (total_amount - untaxed_amount) * percentage
            else:
                discount_amount_currency = total_amount * percentage
            return self.currency_id.round(total_amount - discount_amount_currency)
        return total_amount

    @api.depends('company_id')
    def _compute_discount_computation(self):
        for pay_term in self:
            country_code = pay_term.company_id.country_code or self.env.company.country_code
            if country_code == 'BE':
                pay_term.early_pay_discount_computation = 'mixed'
            elif country_code == 'NL':
                pay_term.early_pay_discount_computation = 'excluded'
            else:
                pay_term.early_pay_discount_computation = 'included'

    @api.depends('line_ids')
    def _compute_example_invalid(self):
        for payment_term in self:
            payment_term.example_invalid = not payment_term.line_ids

    @api.depends('currency_id', 'example_amount', 'example_date', 'line_ids.value', 'line_ids.value_amount', 'line_ids.nb_days', 'early_discount', 'discount_percentage', 'discount_days')
    def _compute_example_preview(self):
        for record in self:
            example_preview = ""
            record.example_preview_discount = ""
            currency = record.currency_id
            
            # 🎯 DESCUENTO SIMPLE COMO ODOO 17 (sin estilos embebidos)
            if record.early_discount:
                discount_date = record._get_last_discount_date(record.example_date or fields.Date.context_today(record))
                if discount_date:
                    discount_date_formatted = format_date(self.env, discount_date)
                    discount_amount = record.example_amount * (record.discount_percentage / 100.0)
                    
                    record.example_preview_discount = _(
                        "Descuento por pago anticipado: %(amount)s si se paga antes del %(date)s",
                        amount=formatLang(self.env, discount_amount, monetary=True, currency_obj=currency),
                        date=discount_date_formatted,
                    )

            # 🎯 CUOTAS SIMPLES COMO ODOO 17 (sin estilos embebidos)
            if not record.example_invalid:
                terms = record._compute_terms(
                    date_ref=record.example_date or fields.Date.context_today(record),
                    currency=currency,
                    company=self.env.company,
                    tax_amount=0,
                    tax_amount_currency=0,
                    untaxed_amount=record.example_amount,
                    untaxed_amount_currency=record.example_amount,
                    sign=1)
                for i, info_by_dates in enumerate(record._get_amount_by_date(terms).values()):
                    date = info_by_dates['date']
                    amount = info_by_dates['amount']
                    example_preview += _(
                        "%(count)s# Cuota de %(amount)s vence el %(date)s",
                        count=i+1,
                        amount=formatLang(self.env, amount, monetary=True, currency_obj=currency),
                        date=date,
                    )

            record.example_preview = example_preview

    @api.model
    def _get_amount_by_date(self, terms):
        """
        Returns a dictionary with the amount for each date of the payment term
        (grouped by date, discounted percentage and discount last date,
        sorted by date and ignoring null amounts).
        """
        terms_lines = sorted(terms["line_ids"], key=lambda t: t.get('date'))
        amount_by_date = {}
        for term in terms_lines:
            key = frozendict({
                'date': term['date'],
            })
            results = amount_by_date.setdefault(key, {
                'date': format_date(self.env, term['date']),
                'amount': 0.0,
            })
            results['amount'] += term['foreign_amount']
        return amount_by_date

    @api.constrains('line_ids', 'early_discount')
    def _check_lines(self):
        round_precision = self.env['decimal.precision'].precision_get('Payment Terms')
        for terms in self:
            # Check if we have any lines
            if not terms.line_ids:
                continue
                
            total_percent = sum(line.value_amount for line in terms.line_ids if line.value == 'percent')
            if float_round(total_percent, precision_digits=round_precision) != 100:
                raise ValidationError(_('El Plazo de Pago debe tener al menos una línea de porcentaje y la suma de los porcentajes debe ser 100%.'))
            if len(terms.line_ids) > 1 and terms.early_discount:
                raise ValidationError(
                    _("La funcionalidad de Descuento por Pago Anticipado solo se puede usar con plazos de pago que usen una sola línea del 100%. "))
            if terms.early_discount and terms.discount_percentage <= 0.0:
                raise ValidationError(_("El Descuento por Pago Anticipado debe ser estrictamente positivo."))
            if terms.early_discount and terms.discount_days <= 0:
                raise ValidationError(_("Los días de Descuento por Pago Anticipado deben ser estrictamente positivos."))

    def compute(self, value, date_ref=False, currency=None):
        """Método original de Odoo 15 - mantener compatibilidad"""
        self.ensure_one()
        date_ref = date_ref or fields.Date.context_today(self)
        amount = value
        sign = value < 0 and -1 or 1
        result = []
        if not currency and self.env.context.get('currency_id'):
            currency = self.env['res.currency'].browse(self.env.context['currency_id'])
        elif not currency:
            currency = self.env.company.currency_id
            
        for line in self.line_ids:
            if line.value == 'fixed':
                amt = sign * currency.round(line.value_amount)
            elif line.value == 'percent':
                amt = currency.round(value * (line.value_amount / 100.0))
            else:
                # Para compatibilidad con líneas antiguas tipo 'balance'
                amt = currency.round(amount)
                
            # Usar el nuevo método _get_due_date si existe, sino usar lógica antigua
            if hasattr(line, '_get_due_date'):
                next_date = line._get_due_date(date_ref)
            else:
                # Lógica antigua para compatibilidad
                next_date = fields.Date.from_string(date_ref)
                if hasattr(line, 'days'):
                    next_date += relativedelta(days=line.days or 0)
                elif hasattr(line, 'nb_days'):
                    next_date += relativedelta(days=line.nb_days or 0)
                    
            result.append((fields.Date.to_string(next_date), amt))
            amount -= amt
            
        amount = sum(amt for _, amt in result)
        dist = currency.round(value - amount)
        if dist:
            last_date = result and result[-1][0] or fields.Date.context_today(self)
            result.append((last_date, dist))
        return sorted(result, key=lambda k: k[0])

    def _compute_terms(self, date_ref, currency, company, tax_amount, tax_amount_currency, sign, untaxed_amount, untaxed_amount_currency, cash_rounding=None):
        """Get the distribution of this payment term."""
        self.ensure_one()
        company_currency = company.currency_id
        total_amount = tax_amount + untaxed_amount
        total_amount_currency = tax_amount_currency + untaxed_amount_currency
        rate = abs(total_amount_currency / total_amount) if total_amount else 0.0

        pay_term = {
            'total_amount': total_amount,
            'discount_percentage': self.discount_percentage if self.early_discount else 0.0,
            'discount_date': self._get_last_discount_date(date_ref) if self.early_discount else False,
            'discount_balance': 0,
            'line_ids': [],
        }

        if self.early_discount:
            # Early discount is only available on single line, 100% payment terms.
            discount_percentage = self.discount_percentage / 100.0
            if self.early_pay_discount_computation in ('excluded', 'mixed'):
                pay_term['discount_balance'] = company_currency.round(total_amount - untaxed_amount * discount_percentage)
                pay_term['discount_amount_currency'] = currency.round(total_amount_currency - untaxed_amount_currency * discount_percentage)
            else:
                pay_term['discount_balance'] = company_currency.round(total_amount * (1 - discount_percentage))
                pay_term['discount_amount_currency'] = currency.round(total_amount_currency * (1 - discount_percentage))

            if cash_rounding:
                cash_rounding_difference_currency = cash_rounding.compute_difference(currency, pay_term['discount_amount_currency'])
                if not currency.is_zero(cash_rounding_difference_currency):
                    pay_term['discount_amount_currency'] += cash_rounding_difference_currency
                    pay_term['discount_balance'] = company_currency.round(pay_term['discount_amount_currency'] / rate) if rate else 0.0

        residual_amount = total_amount
        residual_amount_currency = total_amount_currency

        for i, line in enumerate(self.line_ids):
            term_vals = {
                'date': line._get_due_date(date_ref),
                'company_amount': 0,
                'foreign_amount': 0,
            }

            # The last line is always the balance, no matter the type
            on_balance_line = i == len(self.line_ids) - 1
            if on_balance_line:
                term_vals['company_amount'] = residual_amount
                term_vals['foreign_amount'] = residual_amount_currency
            elif line.value == 'fixed':
                # Fixed amounts
                term_vals['company_amount'] = sign * company_currency.round(line.value_amount / rate) if rate else 0.0
                term_vals['foreign_amount'] = sign * currency.round(line.value_amount)
            else:
                # Percentage amounts
                line_amount = company_currency.round(total_amount * (line.value_amount / 100.0))
                line_amount_currency = currency.round(total_amount_currency * (line.value_amount / 100.0))
                term_vals['company_amount'] = line_amount
                term_vals['foreign_amount'] = line_amount_currency

            if cash_rounding and not on_balance_line:
                cash_rounding_difference_currency = cash_rounding.compute_difference(currency, term_vals['foreign_amount'])
                if not currency.is_zero(cash_rounding_difference_currency):
                    term_vals['foreign_amount'] += cash_rounding_difference_currency
                    term_vals['company_amount'] = company_currency.round(term_vals['foreign_amount'] / rate) if rate else 0.0

            residual_amount -= term_vals['company_amount']
            residual_amount_currency -= term_vals['foreign_amount']
            pay_term['line_ids'].append(term_vals)

        return pay_term

    def _get_last_discount_date(self, date_ref):
        """
        Calcular fecha límite para descuento por pago anticipado
        
        Args:
            date_ref: Fecha de referencia (normalmente fecha de factura)
            
        Returns:
            date: Fecha límite para aplicar descuento, o False si no aplica
        """
        self.ensure_one()
        if not self.early_discount or not self.discount_days:
            return False
            
        # Convertir a date si viene como string
        if isinstance(date_ref, str):
            from datetime import datetime
            date_ref = datetime.strptime(date_ref, '%Y-%m-%d').date()
        elif not date_ref:
            date_ref = fields.Date.today()
            
        # Calcular fecha límite: fecha_factura + días_descuento
        return date_ref + relativedelta(days=self.discount_days)

    def _get_last_discount_date_formatted(self, date_ref):
        self.ensure_one()
        if not date_ref:
            return None
        return format_date(self.env, self._get_last_discount_date(date_ref))

    def copy(self, default=None):
        default = dict(default or {})
        default['name'] = _('%s (copy)', self.name)
        return super().copy(default)


class AccountPaymentTermLine(models.Model):
    _inherit = "account.payment.term.line"

    # Remove old fields and add new ones
    value = fields.Selection([
        ('percent', 'Porcentaje'),
        ('fixed', 'Monto Fijo')
    ], required=True, default='percent',
       help="Seleccione aquí el tipo de valoración relacionada con esta línea de plazos de pago.")
    
    value_amount = fields.Float(string='Vencimiento', digits='Payment Terms',
                                help="Para porcentaje ingrese una relación entre 0-100.",
                                compute='_compute_value_amount', store=True, readonly=False)
    
    delay_type = fields.Selection([
        ('days_after', 'Días después de fecha de factura'),
        ('days_after_end_of_month', 'Días después del fin de mes'),
        ('days_after_end_of_next_month', 'Días después del fin del mes siguiente'),
    ], required=True, default='days_after')
    
    nb_days = fields.Integer(string='Días', readonly=False, store=True, compute='_compute_days')
    
    # Remove old fields
    days = fields.Integer(string='Número de Días (Obsoleto)', help="Campo obsoleto")
    day_of_the_month = fields.Integer(string='Día del mes (Obsoleto)', help="Campo obsoleto")
    option = fields.Selection([
        ('day_after_invoice_date', "días después de la fecha de factura"),
        ('after_invoice_month', "días después del fin del mes de factura"),
        ('day_following_month', "del mes siguiente"),
        ('day_current_month', "del mes actual"),
    ], string='Opciones (Obsoleto)', help="Campo obsoleto")

    def _get_due_date(self, date_ref):
        self.ensure_one()
        due_date = fields.Date.from_string(date_ref) or fields.Date.today()
        if self.delay_type == 'days_after_end_of_month':
            return date_utils.end_of(due_date, 'month') + relativedelta(days=self.nb_days)
        elif self.delay_type == 'days_after_end_of_next_month':
            return date_utils.end_of(due_date + relativedelta(months=1), 'month') + relativedelta(days=self.nb_days)
        return due_date + relativedelta(days=self.nb_days)

    @api.constrains('value', 'value_amount')
    def _check_percent(self):
        for term_line in self:
            if term_line.value == 'percent' and (term_line.value_amount < 0.0 or term_line.value_amount > 100.0):
                raise ValidationError(_('Los porcentajes en las líneas de Plazos de Pago deben estar entre 0 y 100.'))

    @api.depends('payment_id')
    def _compute_days(self):
        for line in self:
            if not line.nb_days:
                if len(line.payment_id.line_ids) > 1:
                    # Si hay otras líneas, tomar el valor de la línea anterior + 30
                    other_lines = line.payment_id.line_ids.filtered(lambda r: r.id != line.id)
                    if other_lines:
                        line.nb_days = max(other_lines.mapped('nb_days') or [0]) + 30
                    else:
                        line.nb_days = 30
                else:
                    # Si es la primera línea, usar valor por defecto
                    line.nb_days = 0
            # Si ya tiene valor, mantenerlo

    @api.depends('payment_id', 'value')
    def _compute_value_amount(self):
        for line in self:
            if line.value == 'fixed':
                if not line.value_amount:
                    line.value_amount = 0
            else:  # percent
                if not line.value_amount:
                    # Calcular el porcentaje restante
                    used_amount = 0
                    for other_line in line.payment_id.line_ids.filtered(lambda r: r.value == 'percent' and r.id != line.id):
                        used_amount += other_line.value_amount or 0
                    line.value_amount = max(0, 100 - used_amount)