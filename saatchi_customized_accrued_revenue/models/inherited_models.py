# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.exceptions import UserError



class SaleOrder(models.Model):
    _inherit="sale.order"

    sub_currency_id = fields.Many2one(
        'res.currency',
        string="Foreign Currency",
        default=lambda self: self.env.ref('base.USD')
    )
    
    custom_exchange_rate = fields.Float(
        string="Foreign Currency Exchange Rate",
        digits=(12, 2),
        help="1 unit of document currency = X sub currency units",
        compute="_compute_custom_exchange_rate",
        # store=True
    )
    
    # @api.depends('currency_id', 'sub_currency_id', 'date')
    def _compute_custom_exchange_rate(self):
        for record in self:
            # If no sub_currency_id is set, default to 1.0
            if not record.sub_currency_id:
                record.custom_exchange_rate = 1.0
                continue
                
            # If document currency is the same as sub currency, rate is 1.0
            if record.currency_id == record.sub_currency_id:
                record.custom_exchange_rate = 2.0
                continue
                
            # Compute exchange rate from document currency to sub_currency_id
            try:
                date = record.date_order or fields.Date.today()
                company = record.company_id or self.env.company
                
                # Convert 1 unit of document currency to sub_currency_id
                rate = record.currency_id._convert(
                    record.amount_total,
                    record.sub_currency_id,
                    company,
                    date
                )
                record.custom_exchange_rate = rate
                
            except Exception as e:
                # Fallback: try manual rate lookup using the rates from currency table
                try:
                    
                    # Get the rate for document currency (JPY in your case)
                    doc_currency_rate = self.env['res.currency.rate'].search([
                        ('currency_id', '=', record.currency_id.id),
                        ('name', '<=', search_date),
                        ('company_id', 'in', [company.id, False])
                    ], order='name desc', limit=1)
                    
                    # Get the rate for sub currency (PHP in your case)  
                    sub_currency_rate = self.env['res.currency.rate'].search([
                        ('currency_id', '=', record.sub_currency_id.id),
                        ('name', '<=', search_date),
                        ('company_id', 'in', [company.id, False])
                    ], order='name desc', limit=1)
                    
                    if doc_currency_rate and sub_currency_rate:
                        # If both currencies have rates, calculate cross rate
                        # rate = (1 / doc_rate) * sub_rate
                        record.custom_exchange_rate = sub_currency_rate.rate / doc_currency_rate.rate
                    elif not doc_currency_rate and sub_currency_rate:
                        # Document currency is base currency (rate = 1.0)
                        record.custom_exchange_rate = sub_currency_rate.rate
                    elif doc_currency_rate and not sub_currency_rate:
                        # Sub currency is base currency (rate = 1.0)
                        record.custom_exchange_rate = 1 / doc_currency_rate.rate
                    else:
                        record.custom_exchange_rate = 1
                        
                except Exception:
                    record.custom_exchange_rate = 1
                    
    def action_create_custom_accrued_revenue(self, is_override=False):
        """Create a new accrued revenue entry for this sale order"""
        if not is_override:
            if self.state != 'sale' or self.x_studio_ce_status not in ['Signed', 'Billable']:
                return False
        
        # Create the accrued revenue record
        accrued_revenue = self.env['saatchi.accrued_revenue'].create({
            'related_ce_id': self.id,
            'currency_id': self.currency_id.id,
        })
        
        # Create lines for each sale order line, but only for Agency Charges
        total_eligible_for_accrue = 0
        process_lines = False  # Flag to track if we're in the Agency section
        
        for line in self.order_line:
            # Check if this is a section line
            if line.display_type == 'line_section':
                # Check if this is the Agency Charges section
                if line.name == 'Agency Charges':
                    process_lines = True
                else:
                    # This is a different section (Non-Agency, etc.), stop processing
                    process_lines = False
                continue  # Skip the section line itself
                
            # Skip line notes and other display types
            if line.display_type:
                continue
                
            # Only process lines if we're in the Agency section
            if process_lines:
                # Calculate accrued quantity (delivered but not invoiced)
                accrued_qty = line.product_uom_qty - line.qty_invoiced
                
                # Skip if nothing to accrue
                if accrued_qty <= 0:
                    continue
                    
                # Calculate accrued amount
                accrued_amount = accrued_qty * line.price_unit
                
                # Get analytic distribution from sale order line
                analytic_distribution = line.analytic_distribution or {}
                
                # If no analytic distribution on line, try to get from sale order
                if not analytic_distribution and hasattr(self, 'analytic_distribution') and self.analytic_distribution:
                    # Use sale order's analytic distribution if available
                    analytic_distribution = self.analytic_distribution
                
                self.env['saatchi.accrued_revenue_lines'].create({
                    'accrued_revenue_id': accrued_revenue.id,
                    'ce_line_id': line.id,
                    'account_id': line.product_id.property_account_income_id.id or line.product_id.categ_id.property_account_income_categ_id.id,
                    'label': f'{self.name} - {line.name}' or 'Accrued Revenue Line',
                    'credit': accrued_amount,
                    'currency_id': line.currency_id.id,
                    'analytic_distribution': analytic_distribution,  # Add analytic distribution
                })
                total_eligible_for_accrue += accrued_amount
                
        # Create the total line (usually doesn't need analytic distribution)
        self.env['saatchi.accrued_revenue_lines'].create({
            'accrued_revenue_id': accrued_revenue.id,
            'label': 'Total Accrued',
            'currency_id': self.currency_id.id,
            # Note: Total Accrued line typically doesn't need analytic_distribution
            # as it's a balancing entry, but you can add it if needed:
            # 'analytic_distribution': {},
        })
        
        accrued_revenue.write({'ce_original_total_amount': total_eligible_for_accrue})
        
        return True


class AccountMove(models.Model):
    _inherit = "account.move"
    
    related_custom_accrued_record = fields.Many2one(
        'saatchi.accrued_revenue', 
        store=True, 
        readonly=True
    )
    
    sub_currency_id = fields.Many2one(
        'res.currency',
        string="Foreign Currency",
        default=lambda self: self.env.ref('base.USD')
    )
    
    custom_exchange_rate = fields.Float(
        string="Foreign Currency Exchange Rate",
        digits=(12, 2),
        help="1 unit of document currency = X sub currency units",
        compute="_compute_custom_exchange_rate",
        # store=True
    )
    
    # @api.depends('currency_id', 'sub_currency_id', 'date')
    def _compute_custom_exchange_rate(self):
        for record in self:
            # If no sub_currency_id is set, default to 1.0
            if not record.sub_currency_id:
                record.custom_exchange_rate = 1.0
                continue
                
            # If document currency is the same as sub currency, rate is 1.0
            if record.currency_id == record.sub_currency_id:
                record.custom_exchange_rate = 1.0
                continue
                
            # Compute exchange rate from document currency to sub_currency_id
            try:
                date = record.date or fields.Date.today()
                company = record.company_id or self.env.company
                
                # Convert 1 unit of document currency to sub_currency_id
                rate = record.currency_id._convert(
                    record.amount_total,
                    record.sub_currency_id,
                    company,
                    date
                )
                record.custom_exchange_rate = rate
                
            except Exception as e:
                # Fallback: try manual rate lookup using the rates from currency table
                try:
                    # In Odoo, base currency (usually USD) has rate = 1.0
                    # In Odoo, base currency (usually USD) has rate = 1.0
                    # Other currencies have rates relative to base currency
                    
                    # Get the rate for document currency (JPY in your case)
                    doc_currency_rate = self.env['res.currency.rate'].search([
                        ('currency_id', '=', record.currency_id.id),
                        ('name', '<=', search_date),
                        ('company_id', 'in', [company.id, False])
                    ], order='name desc', limit=1)
                    
                    # Get the rate for sub currency (PHP in your case)  
                    sub_currency_rate = self.env['res.currency.rate'].search([
                        ('currency_id', '=', record.sub_currency_id.id),
                        ('name', '<=', search_date),
                        ('company_id', 'in', [company.id, False])
                    ], order='name desc', limit=1)
                    
                    if doc_currency_rate and sub_currency_rate:
                        # If both currencies have rates, calculate cross rate
                        # rate = (1 / doc_rate) * sub_rate
                        record.custom_exchange_rate = sub_currency_rate.rate / doc_currency_rate.rate
                    elif not doc_currency_rate and sub_currency_rate:
                        # Document currency is base currency (rate = 1.0)
                        record.custom_exchange_rate = sub_currency_rate.rate
                    elif doc_currency_rate and not sub_currency_rate:
                        # Sub currency is base currency (rate = 1.0)
                        record.custom_exchange_rate = 1.0 / doc_currency_rate.rate
                    else:
                        record.custom_exchange_rate = 1.0
                        
                except Exception:
                    record.custom_exchange_rate = 1.0
            