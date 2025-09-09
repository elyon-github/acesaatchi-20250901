# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.exceptions import UserError


class SaatchiCustomizedAccruedRevenue(models.Model):
    _name = 'saatchi.accrued_revenue'
    _description = 'Saatchi Customized Accrued Revenue'

    related_ce_id = fields.Many2one('sale.order', string="CE#", readonly=True)

    ce_partner_id = fields.Many2one(
        'res.partner',
        string="Customer",
        compute="_compute_ce_fields",
        store=True,
        readonly=True
    )

    ce_status = fields.Selection(
        
        selection=[('Signed', 'Signed'),('Billable', 'Billable'),('Closed', 'Closed'),('Cancelled', 'Cancelled'), ('For Client Signature', 'For Client Signature')],  # placeholder, Odoo will infer from compute if needed
        string="Status",
        compute="_compute_ce_fields",
        store=True,
        readonly=True
    )

    ce_job_description = fields.Char(
        string="Job Description",
        compute="_compute_ce_fields",
        store=True,
        readonly=True
    )

    ce_original_total_amount = fields.Monetary(
        string="Total CE Maximum Amount for Accrue",
        currency_field="currency_id",
        store=True
    )

    line_ids = fields.One2many(
        'saatchi.accrued_revenue_lines',
        'accrued_revenue_id',
        string="Revenue Lines"
    )


    # --- Accounting Information ---
    journal_id = fields.Many2one(
        'account.journal',
        string="Journal",
    )

    accrual_account_id = fields.Many2one(
        'account.account',
        string="Accrual Account",

    )

    date = fields.Date(
        string="Accrual Date",
        default=fields.Date.context_today,
        required=True
    )

    reversal_date = fields.Date(
        string="Reversal Date"
    )

    currency_id = fields.Many2one(
        'res.currency',
        string="Currency",
        required=True,
    )
    
    total_debit_in_accrue_account = fields.Monetary(
        string="Total Debit for Accrue Account",
        compute="_compute_total_debit_in_accrue_account",
        currency_field="currency_id",
        store=True
    )
    
    @api.depends('line_ids.credit')
    def _compute_total_debit_in_accrue_account(self):
        for record in self:
            # Calculate total excluding the "Total Accrued" line to avoid circular dependency
            credit_lines = record.line_ids.filtered(lambda l: l.label != 'Total Accrued')
            total = sum(credit_lines.mapped('credit'))
            record.total_debit_in_accrue_account = total
    
    company_id = fields.Many2one(
        'res.company',
        string="Company",
        required=True,
        default=lambda self: self.env.company
    )
    
    @api.depends("related_ce_id")
    def _compute_ce_fields(self):
        for rec in self:
            if rec.related_ce_id:
                rec.ce_partner_id = rec.related_ce_id.partner_id.id or False
                rec.ce_status = rec.related_ce_id.x_studio_ce_status or False
                rec.ce_job_description = rec.related_ce_id.x_studio_job_description or False
            else:
                rec.ce_partner_id = False
                rec.ce_status = False
                rec.ce_job_description = False

    

    def update_total_accrued_line(self):
        """Update or create the Total Accrued line with the computed total"""
        for record in self:
            credit_lines = record.line_ids.filtered(lambda l: l.label != 'Total Accrued')
            total = sum(credit_lines.mapped('credit'))
            
            # Find existing "Total Accrued" line
            accrued_total_line = record.line_ids.filtered(lambda l: l.label == 'Total Accrued')
            
            if total > 0:  # Only create/update if there's a total
                if accrued_total_line:
                    # Update existing line
                    if record.ce_original_total_amount and total > record.ce_original_total_amount:
                        raise UserError("Total accrued amount cannot exceed the original CE amount.")
                    accrued_total_line.write({
                        'debit': total,
                        'account_id': record.accrual_account_id.id if record.accrual_account_id else accrued_total_line.account_id.id,
                        'currency_id': record.currency_id.id,
                    })
                # else:
                #     # Create new "Total Accrued" line
                #     self.env['saatchi.accrued_revenue_lines'].create({
                #         'accrued_revenue_id': record.id,
                #         'label': 'Total Accrued',
                #         'debit': total,
                #         'credit': 0.0,
                #         'account_id': record.accrual_account_id.id,
                #         'sequence': 9999,  # Put it at the top,
                #         'currency_id': record.currency_id.id,
                #     })
            elif accrued_total_line and total == 0:
                # Remove the line if total is 0
                accrued_total_line.unlink()



class SaatchiCustomizedAccruedRevenueLines(models.Model):
    _name = 'saatchi.accrued_revenue_lines'
    _description = 'Saatchi Customized Accrued Revenue Lines'
    _order = 'sequence desc'

    sequence = fields.Integer(
        string="Sequence",
        default=10
    )
    accrued_revenue_id = fields.Many2one(
        'saatchi.accrued_revenue',
        string="Accrued Revenue",
        ondelete='cascade',
        required=True
    )

    ce_line_id = fields.Many2one('sale.order.line', string="Sale Order Line", ondelete='cascade')

    account_id = fields.Many2one(
        'account.account',
        string="Account",
        domain=[('deprecated', '=', False)]
    )

    label = fields.Char(
        string="Label",
        required=True
    )

    debit = fields.Monetary(
        string="Debit",
        currency_field='currency_id',
        default=0.0
    )

    credit = fields.Monetary(
        string="Credit",
        currency_field='currency_id',
        default=0.0
    )

    currency_id = fields.Many2one(
        'res.currency',
        string="Currency",
        required=True,
        default=lambda self: self.account_id.currency_id
    )

    company_id = fields.Many2one(
        'res.company',
        string="Company",
        required=True,
        default=lambda self: self.env.company
    )


    @api.model_create_multi
    def create(self, vals_list):
        lines = super().create(vals_list)
        for line in lines:
            if line.accrued_revenue_id:
                line.accrued_revenue_id.update_total_accrued_line()
        return lines
    
    def write(self, vals):
        result = super().write(vals)
        if 'credit' in vals:
            for line in self:
                if line.accrued_revenue_id:
                    line.accrued_revenue_id.update_total_accrued_line()
                    
        return result
    
    def unlink(self):
        accrued_revenues = self.mapped('accrued_revenue_id')
        result = super().unlink()
        for revenue in accrued_revenues:
            revenue.update_total_accrued_line()
        return result

class SaleOrder(models.Model):
    _inherit="sale.order"

    def action_create_custom_accrued_revenue(self):
        """Create a new accrued revenue entry for this sale order"""
        
        # Create the accrued revenue record
        accrued_revenue = self.env['saatchi.accrued_revenue'].create({
            'related_ce_id': self.id,
            'currency_id': self.currency_id.id,
        })
        
        # Create lines for each sale order line
        total_eligible_for_accrue = 0
        for line in self.order_line:
            self.env['saatchi.accrued_revenue_lines'].create({
                'accrued_revenue_id': accrued_revenue.id,
                'ce_line_id': line.id,
                'account_id': line.product_id.property_account_income_id.id or line.product_id.categ_id.property_account_income_categ_id.id,
                'label': line.name or 'Accrued Revenue Line',
                'credit': line.price_subtotal, 'currency_id': line.currency_id.id
            })
            total_eligible_for_accrue += line.price_subtotal

            
        self.env['saatchi.accrued_revenue_lines'].create(
                {'accrued_revenue_id': accrued_revenue.id,
                 'label': 'Total Accrued', 'currency_id': self.currency_id.id})

        accrued_revenue.write({'ce_original_total_amount': total_eligible_for_accrue})
        # Return action to open the created record
        return {
            'type': 'ir.actions.act_window',
            'name': 'Accrued Revenue',
            'res_model': 'saatchi.accrued_revenue',
            'res_id': accrued_revenue.id,
            'view_mode': 'form',
            'target': 'current',
        }
        
