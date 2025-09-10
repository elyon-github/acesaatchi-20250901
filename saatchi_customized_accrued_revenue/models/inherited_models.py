# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.exceptions import UserError



class SaleOrder(models.Model):
    _inherit="sale.order"
    
    def action_create_custom_accrued_revenue(self):
        """Create a new accrued revenue entry for this sale order"""
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
                
                self.env['saatchi.accrued_revenue_lines'].create({
                    'accrued_revenue_id': accrued_revenue.id,
                    'ce_line_id': line.id,
                    'account_id': line.product_id.property_account_income_id.id or line.product_id.categ_id.property_account_income_categ_id.id,
                    'label': f'{self.name} - {line.name}' or 'Accrued Revenue Line',
                    'credit': accrued_amount,
                    'currency_id': line.currency_id.id
                })
                total_eligible_for_accrue += accrued_amount
                
        # Create the total line
        self.env['saatchi.accrued_revenue_lines'].create({
            'accrued_revenue_id': accrued_revenue.id,
            'label': 'Total Accrued',
            'currency_id': self.currency_id.id
        })
        
        accrued_revenue.write({'ce_original_total_amount': total_eligible_for_accrue})
        
        return True



class AccountMove(models.Model):
    _inherit="account.move"

    related_custom_accrued_record = fields.Many2one('saatchi.accrued_revenue', store=True, readonly=True)
        
            