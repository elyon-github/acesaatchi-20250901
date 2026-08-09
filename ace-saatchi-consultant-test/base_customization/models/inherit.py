# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.exceptions import UserError, ValidationError
from dateutil.relativedelta import relativedelta
from odoo import models, _
from datetime import timedelta
from odoo import models, api
from lxml import html
import re


class InheritUsers(models.Model):
    _inherit = 'res.users'

    x_client_assignment_ids = fields.One2many(
        'base_customization.user_client_assignment',
        'user_id',
        string="Client Assignments"
    )

    # Computed fields for backward compatibility and easy filtering
    x_client_ids = fields.Many2many(
        'res.partner',
        compute='_compute_assigned_clients_products',
        string="Assigned Clients"
    )

    x_product_ids = fields.Many2many(
        'product.template',
        compute='_compute_assigned_clients_products',
        string="Assigned Products"
    )

    @api.depends('x_client_assignment_ids.x_partner_id', 'x_client_assignment_ids.x_product_ids')
    def _compute_assigned_clients_products(self):
        """Compute flattened lists for easier domain usage"""
        for user in self:
            user.x_client_ids = user.x_client_assignment_ids.mapped(
                'x_partner_id')
            user.x_product_ids = user.x_client_assignment_ids.mapped(
                'x_product_ids')


class InheritSaleOrder(models.Model):
    _inherit = 'sale.order'

    partner_id = fields.Many2one(
        'res.partner',
        domain=lambda self: [('id', 'in', self.env.user.x_client_ids.ids)]
        if self.env.user.employee_id and
        self.env.user.employee_id.department_id and
        self.env.user.employee_id.department_id.name == 'ACCOUNTS'
        else [('id', '!=', False)]  # This returns no records
    )

    x_client_product_ce_code_domain = fields.Char(
        compute='_compute_x_client_product_ce_code_domain'
    )

    x_client_product_ce_code = fields.Many2one(
        'base_customization.client_product_ce_co_line',
        string="Client - Product CE Code",
        domain="x_client_product_ce_code_domain"
    )

    @api.depends('partner_id')
    def _compute_x_client_product_ce_code_domain(self):
        for record in self:
            
            if self.env.user.id != 2:
                domain = [
                    ('x_client_product_ce_co_id.x_partner_id', '=', record.partner_id.id),
                    ('x_product_id', 'in', self.env.user.x_product_ids.ids)
                ]
            else:
                domain = [
                   ('x_client_product_ce_co_id.x_partner_id', '=', record.partner_id.id),
                ]
            record.x_client_product_ce_code_domain = str(domain)

    x_job_description = fields.Char('Job Description')
    x_ce_code = fields.Char(compute="_compute_x_ce_code",
                            string="Client CE Code", store=True)

    x_alt_currency_amount = fields.Float(
        string="Alt Total Amount",
        digits=(12, 2),
        help="Alternative currency total computed from order lines",
        compute="_compute_alt_currency_amount",
        store=True,  # set to True if you need it stored
    )

    x_alt_currency_id = fields.Many2one(
        'res.currency',
        string="Alt Currency",
        default=lambda self: self._default_alt_currency_id(),
    )

    x_remarks = fields.Char(string="Remarks", tracking=True)

    def write(self, vals):
        res = super().write(vals)
        # Only trigger when the alternate currency changes or always on save
        if 'x_alt_currency_id' in vals:
            self._apply_alt_currency_conversion()
        return res

    def _apply_alt_currency_conversion(self):
        """Update all order lines with fx_currency and converted price."""
        for record in self:
            alt_currency = record.x_alt_currency_id
            if not alt_currency:
                continue

            for line in record.order_line:
                line.fx_currency_id = alt_currency
                line.fx_price_unit = line.currency_id._convert(
                    line.price_unit,
                    alt_currency,
                    record.company_id or record.env.company,
                    record.date_order
                )

    # ========== Financial Tracking Fields ==========
    x_ce_approved_estimate_billing = fields.Monetary(
        string="Approved Estimate | Billing",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Total approved estimate for billing (all lines)"
    )

    x_ce_approved_estimate_revenue = fields.Monetary(
        string="Approved Estimate | Revenue",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Total approved estimate for revenue (Agency Charges only)"
    )

    x_ce_invoiced_billing = fields.Monetary(
        string="Invoiced | Billing",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Total invoiced amount for billing (all lines)"
    )

    x_ce_invoiced_revenue = fields.Monetary(
        string="Invoiced | Revenue",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Total invoiced amount for revenue (Agency Charges only)"
    )

    x_ce_variance_billing = fields.Monetary(
        string="Variance | Billing",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Difference between approved estimate and invoiced (billing)"
    )

    x_ce_variance_revenue = fields.Monetary(
        string="Variance | Revenue",
        currency_field="currency_id",
        store=True,
        compute="_compute_x_ce_amounts",
        help="Difference between approved estimate and invoiced (revenue) - eligible for accrual"
    )

    def _default_alt_currency_id(self):
        """Logic:
        1️⃣ If order currency != company currency → use company currency.
        2️⃣ Else if order lines exist with fx_currency_id → use first fx_currency_id.
        3️⃣ Else fallback to USD.
        """
        # self.ensure_one()

        company_currency = self.company_id.currency_id
        order_currency = self.currency_id

        # case 1
        if order_currency and company_currency and order_currency != company_currency:
            return company_currency.id

        # case 2
        fx_currency = self.order_line.filtered(lambda l: l.fx_currency_id)[
            :1].fx_currency_id
        if fx_currency:
            return fx_currency.id

        # case 3
        try:
            usd_currency = self.env.ref('base.USD')
            return usd_currency.id
        except ValueError:
            return company_currency.id or False

    @api.depends('order_line.fx_price_unit', 'x_alt_currency_id', 'date_order')
    def _compute_alt_currency_amount(self):
        """Compute the total alt amount by converting fx_price_unit to alt_currency."""
        for order in self:
            total = 0.0
            alt_currency = order.x_alt_currency_id or order.company_id.currency_id

            for line in order.order_line:
                fx_currency = line.fx_currency_id or order.currency_id
                fx_price = line.fx_price_unit or 0.0

                total += fx_currency._convert(
                    fx_price * line.product_uom_qty,
                    alt_currency,
                    order.company_id,
                    order.date_order or fields.Date.today()
                )
            order.x_alt_currency_amount = total

    @api.depends('x_client_product_ce_code')
    def _compute_x_ce_code(self):
        for record in self:
            if record.x_client_product_ce_code:
                ce_code = f'{record.x_client_product_ce_code.x_ce_product_code}{record.name}'
                record.x_ce_code = ce_code
            else:
                record.x_ce_code = ''

    @api.onchange('partner_id')
    def _onchange_partner_copy_ce_codes(self):
        """Copy CE codes from master when partner changes"""
        if self.partner_id:
            # Find master CE codes for this partner
            master_lines = self.env['base_customization.client_product_ce_co_line'].search([
                ('x_client_product_ce_co_id.x_partner_id', '=', self.partner_id.id)
            ])

            # Create copies for this sale order (command 5 clears, command 0 creates)
            line_vals = []
            for master_line in master_lines:
                line_vals.append((0, 0, {
                    'x_product_id': master_line.x_product_id.id,
                    'x_ce_product_code': master_line.x_ce_product_code,
                }))

            self.x_client_product_ce_code = [
                (5, 0, 0)] + line_vals  # Clear and add new
        else:
            self.x_client_product_ce_code = [(5, 0, 0)]  # Clear all

    # ========== Compute Methods ==========

    @api.depends(
        'order_line',
        'order_line.price_subtotal',
        'order_line.qty_invoiced',
        'order_line.price_unit',
        'order_line.product_template_id',
    )
    def _compute_x_ce_amounts(self):
        """
        Compute approved estimates, invoiced amounts, and variances

        Logic:
        - BILLING: All lines included
        - REVENUE: Only Agency Charges category lines included
        - VARIANCE: Approved Estimate - Invoiced
        """
        for record in self:
            approved_estimate_billing = 0
            approved_estimate_revenue = 0
            invoiced_billing = 0
            invoiced_revenue = 0

            for line in record.order_line:
                approved_estimate_billing += line.price_subtotal
                invoiced_billing += line.qty_invoiced * line.price_unit

                if record._is_agency_charges_category(line.product_template_id):
                    approved_estimate_revenue += line.price_subtotal
                    invoiced_revenue += line.qty_invoiced * line.price_unit

            record.x_ce_approved_estimate_billing = approved_estimate_billing
            record.x_ce_approved_estimate_revenue = approved_estimate_revenue
            record.x_ce_invoiced_billing = invoiced_billing
            record.x_ce_invoiced_revenue = invoiced_revenue
            record.x_ce_variance_billing = approved_estimate_billing - invoiced_billing
            record.x_ce_variance_revenue = approved_estimate_revenue - invoiced_revenue


class InheritSaleOrderLine(models.Model):
    _inherit = 'sale.order.line'

    fx_currency_id = fields.Many2one(
        'res.currency',
        string="Alt Currency",
        default=lambda self: self.env.ref('base.USD', raise_if_not_found=False)
    )
    fx_price_unit = fields.Float(
        string="Alt Unit Price",
        help="Unit price in foreign exchange currency"
    )

    # Hidden field to track which field was last edited
    _last_edited_price_field = fields.Char(store=False)

    @api.model_create_multi
    def create(self, vals_list):
        """Auto-compute fx_price_unit when creating order lines"""
        lines = super(InheritSaleOrderLine, self).create(vals_list)
        for line in lines:
            if line.price_unit and line.fx_currency_id and not line.fx_price_unit:
                company_currency = line.order_id.company_id.currency_id or self.env.company.currency_id
                if company_currency and company_currency != line.fx_currency_id:
                    line.fx_price_unit = company_currency._convert(
                        line.price_unit,
                        line.fx_currency_id,
                        line.order_id.company_id or self.env.company,
                        line.order_id.date_order or fields.Date.today()
                    )
                elif company_currency == line.fx_currency_id:
                    line.fx_price_unit = line.price_unit
        return lines

    @api.onchange('price_unit', 'fx_currency_id')
    def _onchange_price_unit_to_fx(self):
        """Convert price_unit (order line currency) to fx_price_unit"""
        # Mark that price_unit was edited
        self._last_edited_price_field = 'price_unit'

        if self.price_unit and self.fx_currency_id:
            # Use the sale order line's currency_id (typically pricelist currency)
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from line currency to FX currency
                self.fx_price_unit = line_currency._convert(
                    self.price_unit,
                    self.fx_currency_id,
                    self.order_id.company_id or self.env.company,
                    self.order_id.date_order or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.fx_price_unit = self.price_unit

    @api.onchange('fx_price_unit')
    def _onchange_fx_price_unit_to_price(self):
        """Convert fx_price_unit to price_unit (order line currency)"""
        # Only convert if fx_price_unit was the last field edited by user
        # This prevents the chain reaction when price_unit triggers fx_price_unit change
        if self._last_edited_price_field == 'price_unit':
            # Reset the flag
            self._last_edited_price_field = None
            return

        # Mark that fx_price_unit was edited
        self._last_edited_price_field = 'fx_price_unit'

        if self.fx_price_unit and self.fx_currency_id:
            # Use the sale order line's currency_id (typically pricelist currency)
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from FX currency to line currency
                self.price_unit = self.fx_currency_id._convert(
                    self.fx_price_unit,
                    line_currency,
                    self.order_id.company_id or self.env.company,
                    self.order_id.date_order or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.price_unit = self.fx_price_unit


class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    x_alt_currency_amount = fields.Float(
        string="Alt Total Amount",
        digits=(12, 2),
        help="Alternative currency total computed from order lines",
        compute="_compute_alt_currency_amount",
        store=True,  # set to True if you need it stored
    )

    x_alt_currency_id = fields.Many2one(
        'res.currency',
        string="Alt Currency",
        default=lambda self: self._default_alt_currency_id(),
    )

    def write(self, vals):
        res = super().write(vals)
        # Only trigger when the alternate currency changes or always on save
        if 'x_alt_currency_id' in vals:
            self._apply_alt_currency_conversion()
        return res

    def _apply_alt_currency_conversion(self):
        """Update all order lines with fx_currency and converted price."""
        for record in self:
            alt_currency = record.x_alt_currency_id
            if not alt_currency:
                continue

            for line in record.order_line:
                line.fx_currency_id = alt_currency
                line.fx_price_unit = line.currency_id._convert(
                    line.price_unit,
                    alt_currency,
                    record.company_id or record.env.company,
                    record.date_order
                )

    def _default_alt_currency_id(self):
        """Logic:
        1️⃣ If order currency != company currency → use company currency.
        2️⃣ Else if order lines exist with fx_currency_id → use first fx_currency_id.
        3️⃣ Else fallback to USD.
        """
        # self.ensure_one()

        company_currency = self.company_id.currency_id
        order_currency = self.currency_id

        # case 1
        if order_currency and company_currency and order_currency != company_currency:
            return company_currency.id

        # case 2
        fx_currency = self.order_line.filtered(lambda l: l.fx_currency_id)[
            :1].fx_currency_id
        if fx_currency:
            return fx_currency.id

        # case 3
        try:
            usd_currency = self.env.ref('base.USD')
            return usd_currency.id
        except ValueError:
            return company_currency.id or False

    @api.depends('order_line.fx_price_unit', 'x_alt_currency_id', 'date_order')
    def _compute_alt_currency_amount(self):
        """Compute the total alt amount by converting fx_price_unit to alt_currency."""
        for order in self:
            total = 0.0
            alt_currency = order.x_alt_currency_id or order.company_id.currency_id

            for line in order.order_line:
                fx_currency = line.fx_currency_id or order.currency_id
                fx_price = line.fx_price_unit or 0.0

                total += fx_currency._convert(
                    fx_price * line.product_uom_qty,
                    alt_currency,
                    order.company_id,
                    order.date_order or fields.Date.today()
                )
            order.x_alt_currency_amount = total


class InheritPurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    fx_currency_id = fields.Many2one(
        'res.currency',
        string="Alt Currency",
        default=lambda self: self.env.ref('base.USD', raise_if_not_found=False)
    )
    fx_price_unit = fields.Float(
        string="Alt Unit Price",
        help="Unit price in foreign exchange currency"
    )

    # Hidden field to track which field was last edited
    _last_edited_price_field = fields.Char(store=False)

    @api.model_create_multi
    def create(self, vals_list):
        """Auto-compute fx_price_unit when creating order lines"""
        lines = super(InheritPurchaseOrderLine, self).create(vals_list)
        for line in lines:
            if line.price_unit and line.fx_currency_id and not line.fx_price_unit:
                # Use the purchase order line's currency_id
                line_currency = line.currency_id
                if line_currency and line_currency != line.fx_currency_id:
                    line.fx_price_unit = line_currency._convert(
                        line.price_unit,
                        line.fx_currency_id,
                        line.order_id.company_id or self.env.company,
                        line.order_id.date_order or fields.Date.today()
                    )
                elif line_currency == line.fx_currency_id:
                    line.fx_price_unit = line.price_unit
        return lines

    @api.onchange('price_unit', 'fx_currency_id')
    def _onchange_price_unit_to_fx(self):
        """Convert price_unit (order line currency) to fx_price_unit"""
        # Mark that price_unit was edited
        self._last_edited_price_field = 'price_unit'

        if self.price_unit and self.fx_currency_id:
            # Use the purchase order line's currency_id
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from line currency to FX currency
                self.fx_price_unit = line_currency._convert(
                    self.price_unit,
                    self.fx_currency_id,
                    self.order_id.company_id or self.env.company,
                    self.order_id.date_order or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.fx_price_unit = self.price_unit

    @api.onchange('fx_price_unit')
    def _onchange_fx_price_unit_to_price(self):
        """Convert fx_price_unit to price_unit (order line currency)"""
        # Only convert if fx_price_unit was the last field edited by user
        # This prevents the chain reaction when price_unit triggers fx_price_unit change
        if self._last_edited_price_field == 'price_unit':
            # Reset the flag
            self._last_edited_price_field = None
            return

        # Mark that fx_price_unit was edited
        self._last_edited_price_field = 'fx_price_unit'

        if self.fx_price_unit and self.fx_currency_id:
            # Use the purchase order line's currency_id
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from FX currency to line currency
                self.price_unit = self.fx_currency_id._convert(
                    self.fx_price_unit,
                    line_currency,
                    self.order_id.company_id or self.env.company,
                    self.order_id.date_order or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.price_unit = self.fx_price_unit

    @api.depends('invoice_lines.move_id.state', 'invoice_lines.quantity', 'qty_received', 'product_uom_qty', 'order_id.state')
    def _compute_qty_invoiced(self):
        for line in self:
            # compute qty_invoiced
            qty = 0.0
            for inv_line in line._get_invoice_lines():
                if inv_line.move_id.state not in ['cancel'] or inv_line.move_id.payment_state == 'invoicing_legacy':
                    if inv_line.move_id.move_type == 'in_invoice':
                        qty += inv_line.product_uom_id._compute_quantity(inv_line.quantity, line.product_uom, round=False)
                    elif inv_line.move_id.move_type == 'in_refund':
                        qty -= inv_line.product_uom_id._compute_quantity(inv_line.quantity, line.product_uom)
            line.qty_invoiced = qty

            # compute qty_to_invoice
            if line.order_id.state in ['purchase', 'done']:
                if line.product_id.purchase_method == 'purchase':
                    line.qty_to_invoice = line.product_qty - line.qty_invoiced
                else:
                    line.qty_to_invoice = line.qty_received - line.qty_invoiced
            else:
                line.qty_to_invoice = 0

class AccountMove(models.Model):
    _inherit = 'account.move'

    currency_display = fields.Char(
        related='currency_id.name',
        string='Currency',
        store=True,
        readonly=True
    )

    x_related_so = fields.Many2one(
        'sale.order', compute="_compute_related_so", store=True)

    x_alt_currency_amount = fields.Float(
        string="Alt Total Amount",
        digits=(12, 2),
        help="Alternative currency total computed from order lines",
        compute="_compute_alt_currency_amount",
        store=True,  # set to True if you need it stored
    )

    x_alt_currency_id = fields.Many2one(
        'res.currency',
        store=True,
        compute="_compute_x_alt_currency_id",
        string="Alt Currency"
    )

    x_amount_untaxed_in_company_currency = fields.Monetary(
        string="Tax Excluded in Company Currency",
        currency_field="company_currency_id",
        digits=(12, 2),
        compute="_compute_amount_untaxed_in_company_currency",
        store=True,
        help="Tax excluded amount always in company currency (PHP), converted from invoice currency if needed"
    )

    x_excluded_billed_amount = fields.Monetary(
        string="Excluded Billed Amount",
        currency_field="company_currency_id",
        digits=(12, 2),
        compute="_compute_excluded_billed_amount",
        store=True,
        help="Total amount from journal items posted to excluded billed accounts (as configured in accrual configuration)"
    )

    x_billed_amount_adjustment = fields.Monetary(
        string="Billed Amount (For Adjustment XLSX)",
        currency_field="company_currency_id",
        digits=(12, 2),
        compute="_compute_billed_amount_adjustment",
        store=True,
        help="Adjusted billed amount = Tax Excluded in Company Currency - Excluded Billed Amount"
    )

    company_currency_id = fields.Many2one(
        'res.currency',
        related='company_id.currency_id',
        readonly=True
    )
    
    def get_parsed_table_notes(self):
        """Parse x_studio_table_note HTML to extract tables by line number
        
        Detects variants: Line-1, Line 1, line-1, LINE 1, etc.
        Converts to 0-based indexing (Line 1 → sequence 0)
        Tables without line markers are stored as 'default' and render at end
        
        Returns:
            dict: {'0': '<table>...</table>', '1': '<table>...</table>', 'default': '<table>...</table>'}
        """
        if not self.x_studio_table_note:
            return {}
        
        result = {}
        
        try:
            content = self.x_studio_table_note
            pattern = r'(?i)line\s*[-\s]*(\d+)'
            matches = list(re.finditer(pattern, content))
            
            if matches:
                # Process tables with line markers
                for i, match in enumerate(matches):
                    line_num_raw = int(match.group(1))
                    line_num = str(line_num_raw - 1)
                    start_pos = match.end()
                    
                    if i + 1 < len(matches):
                        end_pos = matches[i + 1].start()
                    else:
                        end_pos = len(content)
                    
                    section_content = content[start_pos:end_pos].strip()
                    
                    if section_content:
                        try:
                            doc = html.fromstring(f'<div>{section_content}</div>')
                            tables = doc.xpath('.//table')
                            
                            if tables:
                                table = tables[0]
                                
                                for cell in table.xpath('.//td | .//th'):
                                    existing_style = cell.get('style', '')
                                    if 'border' not in existing_style.lower():
                                        new_style = existing_style + '; border: 1px solid white;' if existing_style else 'border: 1px solid white;'
                                        cell.set('style', new_style)
                                
                                table_html = html.tostring(table, encoding='unicode')
                                result[line_num] = table_html
                        except Exception as parse_error:
                            if '<table' in section_content and '</table>' in section_content:
                                table_start = section_content.find('<table')
                                table_end = section_content.find('</table>') + 8
                                if table_start != -1 and table_end > table_start:
                                    table_html = section_content[table_start:table_end]
                                    result[line_num] = table_html
                
                # Check if there's content before first marker (unmarked table)
                first_match_pos = matches[0].start()
                if first_match_pos > 0:
                    content_before = content[:first_match_pos].strip()
                    if '<table' in content_before:
                        result['default'] = content_before
            else:
                # No line markers found - entire content is default table
                if '<table' in content:
                    try:
                        doc = html.fromstring(f'<div>{content}</div>')
                        tables = doc.xpath('.//table')
                        
                        if tables:
                            table = tables[0]
                            
                            for cell in table.xpath('.//td | .//th'):
                                existing_style = cell.get('style', '')
                                if 'border' not in existing_style.lower():
                                    new_style = existing_style + '; border: 1px solid white;' if existing_style else 'border: 1px solid white;'
                                    cell.set('style', new_style)
                            
                            table_html = html.tostring(table, encoding='unicode')
                            result['default'] = table_html
                    except:
                        result['default'] = content
            
            return result
            
        except Exception as e:
            import logging
            _logger = logging.getLogger(__name__)
            _logger.error(f"Error parsing table notes: {str(e)}")
            return {}
            
    @api.model
    def create(self, vals):
        record = super().create(vals)
        record._apply_alt_currency_conversion()
        return record

    @api.depends('invoice_line_ids', 'state')
    def _compute_related_so(self):
        for record in self:
            record.x_related_so = False
            # for line in record.invoice_line_ids:
            #     if line.sale_line_ids:
            sale_order = record.invoice_line_ids.mapped('sale_line_ids.order_id')[
                :1] if record.invoice_line_ids else False
            record.x_related_so = sale_order

    def _apply_alt_currency_conversion(self):
        """
        Update invoice lines with fx_currency and converted price.
        Only populate fx_price_unit for invoices in company currency (PHP).
        For foreign currency invoices, alt amount is computed from journal items instead.
        """
        for record in self:
            # Determine which alt currency to use
            alt_currency = record.x_related_so.x_alt_currency_id if record.x_related_so else None
            if not alt_currency:
                alt_currency = record.invoice_line_ids[0].purchase_order_id.x_alt_currency_id if record.invoice_line_ids and record.invoice_line_ids[0].purchase_order_id else None
            
            if not alt_currency:
                continue

            # Only populate fx_price_unit for invoices in company currency (PHP)
            # Foreign currency invoices use journal items (already converted to PHP)
            if record.currency_id == record.company_id.currency_id:
                for line in record.invoice_line_ids:
                    line.fx_currency_id = alt_currency
                    line.fx_price_unit = line.currency_id._convert(
                        line.price_unit,
                        alt_currency,
                        record.company_id or record.env.company,
                        record.x_related_so.date_order if record.x_related_so else (line.purchase_order_id.date_order if line.purchase_order_id else record.date)
                    )
            else:
                # For foreign currency invoices, just set the alt currency on lines
                for line in record.invoice_line_ids:
                    line.fx_currency_id = alt_currency

    @api.depends('name')
    def _compute_x_alt_currency_id(self):
        for record in self:
            if record.invoice_line_ids:
                record.x_alt_currency_id = record.invoice_line_ids[0].fx_currency_id

    @api.depends('invoice_line_ids.fx_price_unit', 'invoice_line_ids.fx_currency_id', 'line_ids.credit', 'state', 'currency_id')
    def _compute_alt_currency_amount(self):
        """
        Compute the total alt amount:
        - If invoice currency != company currency (PHP), sum all journal item credits (already in PHP).
        - If invoice currency == company currency, use invoice lines as before.
        """
        for record in self:
            total = 0.0
            alt_currency = record.x_alt_currency_id or record.company_id.currency_id

            if not record.x_alt_currency_id:
                record.x_alt_currency_amount = 0
                continue

            # Check if invoice currency is different from company currency (PHP)
            if record.currency_id != record.company_id.currency_id:
                # Foreign currency invoice - sum all journal items' credit (already in PHP)
                for line in record.line_ids:
                    total += line.credit
            else:
                # PHP invoice - use invoice lines as before
                for line in record.invoice_line_ids:
                    fx_currency = line.fx_currency_id or record.currency_id
                    fx_price = line.fx_price_unit or 0.0

                    total += fx_currency._convert(
                        fx_price * line.quantity,
                        alt_currency,
                        record.company_id,
                        record.x_related_so.date_order or fields.Date.today()
                    )
            record.x_alt_currency_amount = total

    @api.depends('amount_untaxed', 'line_ids.credit', 'currency_id', 'company_id.currency_id', 'date', 'state', 'move_type')
    def _compute_amount_untaxed_in_company_currency(self):
        """
        Compute tax excluded amount always in company currency (PHP).
        - If invoice currency == company currency (PHP): use amount_untaxed as-is
        - If invoice currency != company currency (USD, etc.): sum journal items' credit (already in PHP)
        For credit memos/refunds, enforce negative sign.
        """
        for record in self:
            if record.currency_id == record.company_id.currency_id:
                value = record.amount_untaxed
            else:
                # Foreign currency invoice - sum journal items' credit (already in PHP)
                total = 0.0
                for line in record.line_ids:
                    total += line.credit
                value = total

            if record.move_type in ['out_refund', 'in_refund']:
                record.x_amount_untaxed_in_company_currency = -abs(value)
            else:
                record.x_amount_untaxed_in_company_currency = value

    def _get_excluded_billed_account_ids(self):
        """Get the set of account IDs to exclude from billed amount calculation.
        
        Reads from the saatchi.accrual_config Many2many field
        'excluded_billed_account_ids' for the current company.
        
        Returns:
            set: Set of account IDs to exclude
        """
        try:
            config = self.env['saatchi.accrual_config'].sudo().search([
                ('company_id', '=', self.company_id.id)
            ], limit=1)
            if config and config.excluded_billed_account_ids:
                return set(config.excluded_billed_account_ids.ids)
        except Exception:
            pass
        return set()

    @api.depends('line_ids.account_id', 'line_ids.credit', 'line_ids.debit', 'state', 'move_type')
    def _compute_excluded_billed_amount(self):
        """
        Compute the total amount from journal items posted to excluded billed accounts.
        
        For each journal item line in an excluded account:
          Total = sum of (credit - debit) for all lines in excluded accounts
        
        For credit memos/refunds, enforce negative sign so it mirrors the move sign.
        """
        for record in self:
            excluded_account_ids = record._get_excluded_billed_account_ids()
            
            if not excluded_account_ids:
                record.x_excluded_billed_amount = 0.0
                continue
            
            total_exclusion = 0.0
            for line in record.line_ids:
                if line.account_id and line.account_id.id in excluded_account_ids:
                    total_exclusion += (line.credit or 0.0) - (line.debit or 0.0)

            if record.move_type in ['out_refund', 'in_refund']:
                record.x_excluded_billed_amount = -abs(total_exclusion)
            else:
                record.x_excluded_billed_amount = total_exclusion

    @api.depends('x_amount_untaxed_in_company_currency', 'x_excluded_billed_amount', 'move_type')
    def _compute_billed_amount_adjustment(self):
        """
        Compute adjusted billed amount in company currency (PHP).
        Formula: Tax Excluded in Company Currency - Excluded Billed Amount
        For refunds, sign should remain negative.
        """
        for record in self:
            adjustment = record.x_amount_untaxed_in_company_currency - record.x_excluded_billed_amount
            if record.move_type in ['out_refund', 'in_refund']:
                record.x_billed_amount_adjustment = -abs(adjustment)
            else:
                record.x_billed_amount_adjustment = adjustment


# BIR Report Customizations
    def _search_account_by_code(self, code_suffix):
        """Search account by last 4 digits of code"""
        # Filter accounts by company using the move's company
        domain = [('code', 'ilike', '%' + code_suffix)]
        
        # Try to filter by company if available
        accounts = self.env['account.account'].with_context(
            allowed_company_ids=[self.company_id.id]
        ).search(domain, limit=1)
        
        return accounts[0] if accounts else False

    def _get_billing_summary_debit(self):
        """Calculate debit side of billing summary"""
        self.ensure_one()
        
        # Define account codes (last 4 digits)
        AR_ADVERTISER_CODE = '1202'  # 111202
        INTER_CO_RECV_CODE = '1206'  # 111206
        
        result = {
            'ar_advertiser': 0.0,
            'inter_co_recv': 0.0,
            'sundries_code': '',
            'sundries_amount': 0.0,
            'total': 0.0
        }
        
        # Get accounts
        ar_account = self._search_account_by_code(AR_ADVERTISER_CODE)
        inter_co_account = self._search_account_by_code(INTER_CO_RECV_CODE)
        
        # Process invoice lines (debit entries)
        for line in self.line_ids.filtered(lambda l: l.debit > 0):
            if ar_account and line.account_id == ar_account:
                result['ar_advertiser'] += line.debit
            elif inter_co_account and line.account_id == inter_co_account:
                result['inter_co_recv'] += line.debit
            else:
                # Sundries - other debit accounts
                result['sundries_amount'] += line.debit
                if not result['sundries_code']:
                    # Use only last 4 digits of account code
                    code = line.account_id.code
                    result['sundries_code'] = code[-4:] if len(code) >= 4 else code
        
        result['total'] = result['ar_advertiser'] + result['inter_co_recv'] + result['sundries_amount']
        
        return result

    def _get_billing_summary_credit(self):
        """Calculate credit side of billing summary"""
        self.ensure_one()
        
        # Define account codes (last 4 digits)
        AP_TRADE_CODE = '2101'  # 112101
        OUTPUT_VAT_CODE = '2507'  # 112507
        
        result = {
            'ap_trade': 0.0,
            'income_code': '',
            'income_amount': 0.0,
            'output_vat': 0.0,
            'sundries_code': '',
            'sundries_amount': 0.0,
            'total': 0.0
        }
        
        # Get accounts
        ap_trade_account = self._search_account_by_code(AP_TRADE_CODE)
        output_vat_account = self._search_account_by_code(OUTPUT_VAT_CODE)
        
        # Process invoice lines (credit entries)
        for line in self.line_ids.filtered(lambda l: l.credit > 0):
            if ap_trade_account and line.account_id == ap_trade_account:
                result['ap_trade'] += line.credit
            elif output_vat_account and line.account_id == output_vat_account:
                result['output_vat'] += line.credit
            elif line.account_id.account_type in ['income', 'income_other']:
                # Income accounts
                result['income_amount'] += line.credit
                if not result['income_code']:
                    # Use only last 4 digits of account code
                    code = line.account_id.code
                    result['income_code'] = code[-4:] if len(code) >= 4 else code
            else:
                # Sundries - other credit accounts
                result['sundries_amount'] += line.credit
                if not result['sundries_code']:
                    # Use only last 4 digits of account code
                    code = line.account_id.code
                    result['sundries_code'] = code[-4:] if len(code) >= 4 else code
        
        result['total'] = result['ap_trade'] + result['income_amount'] + result['output_vat'] + result['sundries_amount']
        
        return result
    
class AccountMoveLine(models.Model):
    _inherit = 'account.move.line'

    currency_display = fields.Char(
        related='currency_id.name',
        string='Currency',
        store=True,
        readonly=True
    )

    fx_currency_id = fields.Many2one(
        'res.currency',
        string="Alt Currency",
        # related='move_id.x_alt_currency_id'
    )
    fx_price_unit = fields.Float(
        string="Alt Unit Price",
        help="Unit price in foreign exchange currency"
    )

    # Hidden field to track which field was last edited
    _last_edited_price_field = fields.Char(store=False)

    @api.model_create_multi
    def create(self, vals_list):
        """Auto-compute fx_price_unit when creating invoice lines"""
        lines = super(AccountMoveLine, self).create(vals_list)
        for line in lines:
            if line.price_unit and line.fx_currency_id and not line.fx_price_unit:
                # Use the invoice line's currency_id
                line_currency = line.currency_id
                if line_currency and line_currency != line.fx_currency_id:
                    line.fx_price_unit = line_currency._convert(
                        line.price_unit,
                        line.fx_currency_id,
                        line.move_id.company_id or self.env.company,
                        line.move_id.date or fields.Date.today()
                    )
                elif line_currency == line.fx_currency_id:
                    line.fx_price_unit = line.price_unit
        return lines

    @api.onchange('price_unit', 'fx_currency_id')
    def _onchange_price_unit_to_fx(self):
        """Convert price_unit (invoice line currency) to fx_price_unit"""
        # Mark that price_unit was edited
        self._last_edited_price_field = 'price_unit'

        if self.price_unit and self.fx_currency_id:
            # Use the invoice line's currency_id
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from line currency to FX currency
                self.fx_price_unit = line_currency._convert(
                    self.price_unit,
                    self.fx_currency_id,
                    self.move_id.company_id or self.env.company,
                    self.move_id.date or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.fx_price_unit = self.price_unit

    @api.onchange('fx_price_unit')
    def _onchange_fx_price_unit_to_price(self):
        """Convert fx_price_unit to price_unit (invoice line currency)"""
        # Only convert if fx_price_unit was the last field edited by user
        # This prevents the chain reaction when price_unit triggers fx_price_unit change
        if self._last_edited_price_field == 'price_unit':
            # Reset the flag
            self._last_edited_price_field = None
            return

        # Mark that fx_price_unit was edited
        self._last_edited_price_field = 'fx_price_unit'

        if self.fx_price_unit and self.fx_currency_id:
            # Use the invoice line's currency_id
            line_currency = self.currency_id
            if line_currency and line_currency != self.fx_currency_id:
                # Convert from FX currency to line currency
                self.price_unit = self.fx_currency_id._convert(
                    self.fx_price_unit,
                    line_currency,
                    self.move_id.company_id or self.env.company,
                    self.move_id.date or fields.Date.today()
                )
            elif line_currency == self.fx_currency_id:
                self.price_unit = self.fx_price_unit


class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    def get_button_approvers(self):
        """
        Returns the approvers configured in Studio approval button
        :return: recordset of res.users
        """
        self.ensure_one()

        approvers = self.env['res.users']

        # Method 1: Get from approval entries (if approval process has started)
        try:
            entries = self.env['studio.approval.entry'].search([
                ('res_id', '=', self.id),
                ('model', '=', 'purchase.order')
            ])
            if entries:
                approvers = entries.mapped('user_id')
                if approvers:
                    return approvers
        except:
            pass

        # Method 2: Get from approval rules configuration
        try:
            # Get the model record for purchase.order
            model_rec = self.env['ir.model'].search([
                ('model', '=', 'purchase.order')
            ], limit=1)

            if model_rec:
                # Find applicable approval rules
                rules = self.env['studio.approval.rule'].search([
                    ('model_id', '=', model_rec.id),
                    ('active', '=', True)
                ])

                for rule in rules:
                    # Check if rule applies to this record based on domain
                    if rule.domain:
                        try:
                            from odoo.tools.safe_eval import safe_eval
                            domain = safe_eval(rule.domain)
                            # Check if this record matches the domain
                            matching = self.search(
                                domain + [('id', '=', self.id)])
                            if not matching:
                                continue
                        except Exception as e:
                            _logger.warning(
                                f"Error evaluating domain for rule {rule.id}: {e}")
                            continue

                    # Get approvers from the rule via studio.approval.rule.approver
                    rule_approvers = self.env['studio.approval.rule.approver'].search([
                        ('rule_id', '=', rule.id)
                    ])

                    for rule_approver in rule_approvers:
                        if rule_approver.user_id:
                            approvers |= rule_approver.user_id
                        if hasattr(rule_approver, 'group_id') and rule_approver.group_id:
                            approvers |= rule_approver.group_id.users

        except Exception as e:
            import logging
            _logger = logging.getLogger(__name__)
            _logger.warning(f"Could not fetch approval rules: {e}")

        return approvers


class HrLeaveType(models.Model):
    _inherit = 'hr.leave.type'

    employer_approver_only_on_days = fields.Integer(
        string="Manager-Only Approval Threshold (Days)",
        help="If leave days are less than or equal to this value, only manager approval is required (HR approval is skipped)."
    )


class HrLeave(models.Model):
    _inherit = 'hr.leave'

    def _get_responsible_for_approval(self):
        self.ensure_one()
        responsible = self.env['res.users']
        leave_type = self.holiday_status_id
        threshold = leave_type.employer_approver_only_on_days
        is_short_leave = threshold and self.number_of_days <= threshold
        is_dual_validation = self.validation_type == 'both'
        
        # Determine the effective validation type
        effective_validation_type = self.validation_type
        # Short leaves: Manager approves (bypass HR)
        if is_dual_validation and is_short_leave and threshold:
            effective_validation_type = 'manager'
        
        # SWAPPED FLOW: HR Officer first, Manager second
        if effective_validation_type == 'hr' or (effective_validation_type == 'both' and self.state == 'confirm'):
            # HR Officer approves first for long leaves
            if self.holiday_status_id.responsible_ids:
                responsible = self.holiday_status_id.responsible_ids
        elif effective_validation_type == 'manager' or (effective_validation_type == 'both' and self.state == 'validate1'):
            # Manager approves: either short leaves directly OR second approval for long leaves
            if self.employee_id.leave_manager_id:
                responsible = self.employee_id.leave_manager_id
            elif self.employee_id.parent_id.user_id:
                responsible = self.employee_id.parent_id.user_id
        
        return responsible

    def action_approve(self, check_state=True):
        # SWAPPED FLOW:
        # - Short leaves: Manager approves directly
        # - Long leaves: HR Officer approves first (confirm → validate1)
        
        if check_state and any(holiday.state != 'confirm' for holiday in self):
            raise UserError(_('Time off request must be confirmed ("To Approve") in order to approve it.'))
        
        current_employee = self.env.user.employee_id
        
        # Separate leaves based on approval flow
        needs_hr_approval = self.env['hr.leave']      # Long leaves needing HR approval first
        needs_manager_approval = self.env['hr.leave']  # Short leaves needing Manager approval
        needs_standard_approval = self.env['hr.leave'] # Other validation types
        
        for holiday in self:
            if holiday.validation_type == 'both':
                threshold = holiday.holiday_status_id.employer_approver_only_on_days
                is_short_leave = threshold and holiday.number_of_days <= threshold
                if is_short_leave:
                    # Short leave: Manager approves directly (bypass HR)
                    needs_manager_approval |= holiday
                else:
                    # Long leave: HR Officer approves first
                    needs_hr_approval |= holiday
            else:
                needs_standard_approval |= holiday
        
        # SWAPPED: Long leaves go to validate1 (HR Officer first approval)
        needs_hr_approval.write({'state': 'validate1', 'first_approver_id': current_employee.id})
        
        # Short leaves and other types go directly to validate
        (needs_manager_approval | needs_standard_approval).action_validate(check_state)
        
        if not self.env.context.get('leave_fast_create'):
            self.activity_update()
        return True



    def _check_approval_update(self, state):
        """ Check if target state is achievable. """
        if self.env.is_superuser():
            return
    
        current_employee = self.env.user.employee_id
        is_officer = self.env.user.has_group('hr_holidays.group_hr_holidays_user')
        is_manager = self.env.user.has_group('hr_holidays.group_hr_holidays_manager')
    
        for holiday in self:
            val_type = holiday.validation_type
            
            # Calculate effective validation type for short leaves
            threshold = holiday.holiday_status_id.employer_approver_only_on_days
            is_short_leave = threshold and holiday.number_of_days <= threshold
            effective_val_type = 'manager' if (val_type == 'both' and is_short_leave) else val_type
    
            if not is_manager:
                if holiday.state == 'cancel' and state != 'confirm':
                    raise UserError(_('A cancelled leave cannot be modified.'))
                if state == 'confirm':
                    if holiday.state == 'refuse':
                        raise UserError(_('Only a Time Off Manager can reset a refused leave.'))
                    if holiday.date_from and holiday.date_from.date() <= fields.Date.today():
                        raise UserError(_('Only a Time Off Manager can reset a started leave.'))
                    if holiday.employee_id != current_employee:
                        raise UserError(_('Only a Time Off Manager can reset other people leaves.'))
                else:
                    if effective_val_type == 'no_validation' and current_employee == holiday.employee_id and (is_officer or is_manager):
                        continue
                        
                    holiday.check_access('write')
    
                    if holiday.employee_id == current_employee\
                            and self.env.user != holiday.employee_id.leave_manager_id\
                            and not is_officer:
                        raise UserError(_('Only a Time Off Officer or Manager can approve/refuse its own requests.'))
    
                    # SWAPPED FLOW PERMISSION CHECKS
                    
                    # validate1 state: HR Officer first approval (for long leaves)
                    if state == 'validate1' and effective_val_type == 'both':
                        # SWAPPED: HR Officer gives first approval
                        if not is_officer:
                            raise UserError(_('You must be a Time off Officer to provide first approval for this leave'))
    
                    # validate state: Final approval
                    elif state == 'validate':
                        if effective_val_type == 'manager':
                            # Manager approval (short leaves OR standard manager validation)
                            if self.env.user != holiday.employee_id.leave_manager_id and not is_officer:
                                raise UserError(_("You must be %s's Manager to approve this leave", holiday.employee_id.name))
                        
                        elif effective_val_type == 'hr':
                            # HR-only approval
                            if not is_officer:
                                raise UserError(_('You must be a Time off Officer to approve this leave'))
                        
                        elif effective_val_type == 'both':
                            # SWAPPED: Manager gives second approval (validate1 → validate)
                            if self.env.user != holiday.employee_id.leave_manager_id and not is_officer:
                                raise UserError(_("You must be %s's Manager to provide second approval", holiday.employee_id.name))


    
    def action_validate(self, check_state=True):
        current_employee = self.env.user.employee_id
        leaves = self._get_leaves_on_public_holiday()
        if leaves:
            raise ValidationError(_('The following employees are not supposed to work during that period:\n %s') % ','.join(leaves.mapped('employee_id.name')))
        if check_state and any(holiday.state not in ['confirm', 'validate1'] and holiday.validation_type != 'no_validation' for holiday in self):
            raise UserError(_('Time off request must be confirmed in order to approve it.'))
    
        self.write({'state': 'validate'})
    
        for leave in self:
            threshold = leave.holiday_status_id.employer_approver_only_on_days
            is_short_leave = threshold and leave.number_of_days <= threshold
            
            if leave.validation_type == 'both':
                if is_short_leave:
                    # Short leave: Manager is the sole approver
                    leave.first_approver_id = current_employee.id
                elif leave.state == 'validate1':
                    # SWAPPED: Long leave coming from validate1 - Manager is second approver
                    leave.second_approver_id = current_employee.id
                else:
                    # Coming directly from confirm (shouldn't happen in swapped flow for long leaves)
                    leave.first_approver_id = current_employee.id
            elif leave.validation_type == 'manager':
                leave.first_approver_id = current_employee.id
            else:  # 'hr' or 'no_validation'
                leave.first_approver_id = current_employee.id
    
        self._validate_leave_request()
        if not self.env.context.get('leave_fast_create'):
            self.filtered(lambda holiday: holiday.validation_type != 'no_validation').activity_update()
        return True
    
    def _check_double_validation_rules(self, employees, state):
        if self.env.user.has_group('hr_holidays.group_hr_holidays_manager'):
            return
    
        is_leave_user = self.env.user.has_group('hr_holidays.group_hr_holidays_user')
        
        # Skip double validation check for short leaves
        for holiday in self:
            threshold = holiday.holiday_status_id.employer_approver_only_on_days
            is_short_leave = threshold and holiday.number_of_days <= threshold
            if holiday.validation_type == 'both' and is_short_leave:
                # Short leaves: Manager can approve directly
                if state in ['validate', 'validate1']:
                    if self.env.user != holiday.employee_id.leave_manager_id and not is_leave_user:
                        raise AccessError(_('You cannot approve a time off for %s, because you are not their time off manager', holiday.employee_id.name))
                return
        
        # SWAPPED FLOW: For long leaves with 'both' validation
        if state == 'validate1':
            # SWAPPED: First approval is HR Officer
            if not is_leave_user:
                raise AccessError(_('You must be a Time Off Officer to provide first approval'))
        elif state == 'validate':
            # Second approval or direct validation
            for holiday in self:
                if holiday.validation_type == 'both' and holiday.state == 'validate1':
                    # SWAPPED: Manager gives second approval
                    if self.env.user != holiday.employee_id.leave_manager_id and not is_leave_user:
                        raise AccessError(_('You cannot approve a time off for %s, because you are not their time off manager', holiday.employee_id.name))
                elif holiday.validation_type == 'both' and holiday.state == 'confirm':
                    # Direct validation from confirm requires HR Officer (shouldn't happen in normal swapped flow)
                    if not is_leave_user:
                        raise AccessError(_('You don\'t have the rights to apply approval on a time off request'))
    
    
    def activity_update(self):
        if self.env.context.get('mail_activity_automation_skip'):
            return False
    
        to_clean, to_do, to_do_confirm_activity = self.env['hr.leave'], self.env['hr.leave'], self.env['hr.leave']
        activity_vals = []
        today = fields.Date.today()
        model_id = self.env['ir.model']._get_id('hr.leave')
        confirm_activity = self.env.ref('hr_holidays.mail_act_leave_approval')
        approval_activity = self.env.ref('hr_holidays.mail_act_leave_second_approval')
        
        for holiday in self:
            if holiday.state in ['confirm', 'validate1']:
                if holiday.holiday_status_id.leave_validation_type != 'no_validation':
                    if holiday.state == 'confirm':
                        activity_type = confirm_activity
                        
                        # Check if short leave for custom activity note
                        threshold = holiday.holiday_status_id.employer_approver_only_on_days
                        is_short_leave = threshold and holiday.number_of_days <= threshold
                        
                        if holiday.validation_type == 'both' and is_short_leave:
                            # Short leave: Manager approves
                            note = _(
                                'New %(leave_type)s Request (%(days)s days - Manager Approval) created by %(user)s',
                                leave_type=holiday.holiday_status_id.name,
                                days=holiday.number_of_days,
                                user=holiday.create_uid.name,
                            )
                        elif holiday.validation_type == 'both':
                            # SWAPPED: Long leave - HR Officer approves first
                            note = _(
                                'New %(leave_type)s Request (HR First Approval Required) created by %(user)s',
                                leave_type=holiday.holiday_status_id.name,
                                user=holiday.create_uid.name,
                            )
                        else:
                            note = _(
                                'New %(leave_type)s Request created by %(user)s',
                                leave_type=holiday.holiday_status_id.name,
                                user=holiday.create_uid.name,
                            )
                    else:
                        # SWAPPED: validate1 state - Manager gives second approval
                        activity_type = approval_activity
                        note = _(
                            'Manager (Second Approval) Required for %(leave_type)s',
                            leave_type=holiday.holiday_status_id.name,
                        )
                        to_do_confirm_activity += holiday
                        
                    user_ids = holiday.sudo()._get_responsible_for_approval().ids
                    for user_id in user_ids:
                        date_deadline = (
                            (holiday.date_from -
                             relativedelta(**{activity_type.delay_unit or 'days': activity_type.delay_count or 0})).date()
                            if holiday.date_from else today)
                        if date_deadline < today:
                            date_deadline = today
                        activity_vals.append({
                            'activity_type_id': activity_type.id,
                            'automated': True,
                            'date_deadline': date_deadline,
                            'note': note,
                            'user_id': user_id,
                            'res_id': holiday.id,
                            'res_model_id': model_id,
                        })
            elif holiday.state == 'validate':
                to_do |= holiday
            elif holiday.state in ['refuse', 'cancel']:
                to_clean |= holiday
                
        if to_clean:
            to_clean.activity_unlink(['hr_holidays.mail_act_leave_approval', 'hr_holidays.mail_act_leave_second_approval'])
        if to_do_confirm_activity:
            to_do_confirm_activity.activity_feedback(['hr_holidays.mail_act_leave_approval'])
        if to_do:
            to_do.activity_feedback(['hr_holidays.mail_act_leave_approval', 'hr_holidays.mail_act_leave_second_approval'])
        self.env['mail.activity'].with_context(short_name=False).create(activity_vals)


class HrExpenseSheet(models.Model):
    _inherit = 'hr.expense.sheet'

    x_has_cash_advance_line_ids = fields.Boolean(
        string="Has Cash Advance Lines", compute="_compute_x_has_cash_advance_line_ids", store=True)

    @api.depends('expense_line_ids.x_studio_cash_advance_series')
    def _compute_x_has_cash_advance_line_ids(self):
        for record in self:
            record.x_has_cash_advance_line_ids = any(
                line.x_studio_cash_advance_series for line in record.expense_line_ids
            )

class JournalReportCustomHandler(models.AbstractModel):
    _inherit = "account.journal.report.handler"
    
    
    def _get_columns_for_journal(self, journal, export_type='pdf'):
        """
        Creates a columns list that will be used in this journal for the pdf report

        :return: A list of the columns as dict each having:
            - name (mandatory):     A string that will be displayed
            - label (mandatory):    A string used to link lines with the column
            - class (optional):     A string with css classes that need to be applied to all that column
        """
        columns = [
            {'name': _('Document'), 'label': 'document'},
        ]

        # We have different columns regarding we are exporting to a PDF file or an XLSX document
        if export_type == 'pdf':
            columns.append({'name': _('Account'), 'label': 'account_label'})
        else:
            columns.extend([
                {'name': _('Account Code'), 'label': 'account_code'},
                {'name': _('Account Label'), 'label': 'account_label'}
            ])

        columns.extend([
            {'name': _('Name'), 'label': 'name'},
            {'name': _('Debit'), 'label': 'debit', 'class': 'o_right_alignment '},
            {'name': _('Credit'), 'label': 'credit', 'class': 'o_right_alignment '},
        ])
        # HERE TAXES COLUMN ARE DISABLED FOR NOW MARKIE
        # if journal.get('tax_summary'):
        #     columns.append(
        #         {'name': _('Taxes'), 'label': 'taxes'},
        #     )
        #     if journal['tax_summary'].get('tax_grid_summary_lines'):
        #         columns.append({'name': _('Tax Grids'), 'label': 'tax_grids'})

        if self._should_use_bank_journal_export(journal):
            columns.append({
                'name': _('Balance'),
                'label': 'balance',
                'class': 'o_right_alignment '
            })

            if journal.get('multicurrency_column'):
                columns.append({
                    'name': _('Amount Currency'),
                    'label': 'amount_currency',
                    'class': 'o_right_alignment '
                })

        return columns
