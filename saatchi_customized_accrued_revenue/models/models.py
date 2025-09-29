# -*- coding: utf-8 -*-
from odoo import models, fields, api, _, Command
from odoo.exceptions import UserError
from dateutil.relativedelta import relativedelta

class SaatchiCustomizedAccruedRevenue(models.Model):
    _name = 'saatchi.accrued_revenue'
    _description = 'Saatchi Customized Accrued Revenue'
    _rec_name = 'display_name'
    _inherit = ['mail.thread', 'mail.activity.mixin']  # Add this line for chatter
    display_name = fields.Char(compute="_compute_display_name")
    
    related_ce_id = fields.Many2one('sale.order', string="SO#", readonly=True)

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

    ce_code = fields.Char(string='CE Code', store=True, compute="_compute_ce_code")

    ce_original_total_amount = fields.Monetary(
        string="Total CE Maximum Amount Subject for Accrual",
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
        default=lambda self: self.env['account.journal'].browse(11),
    )

    accrual_account_id = fields.Many2one(
        'account.account',
        string="Accrual Account",
        default=lambda self: self.env['account.account'].search([('code', '=', '1210')], limit=1),
    )


    date = fields.Date(
        string="Accrual Date",
        default=fields.Date.context_today,
        required=True
    )

    reversal_date = fields.Date(
        string="Reversal Date", compute="_compute_reversal_date", store=True, readonly=False
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
    
    state = fields.Selection(
        [
            ('draft', 'Draft'),
            ('accrued', 'Accrued'),
            ('reversed', 'Reversed'),
            ('cancel', 'cancelled')
        ],
        string="Status",
        default='draft',
        required=True,
        store=True,
        compute="_compute_state"
    )

    related_accrued_entry = fields.Many2one('account.move', readonly=True,string="Accrued Entry")
    related_reverse_accrued_entry = fields.Many2one('account.move', readonly=True,string="Reverse Acrrue Entry")


    @api.depends('related_ce_id')
    def _compute_ce_code(self):
        for record in self:
            record.ce_code = f'{record.related_ce_id.x_studio_ce_code.name}{record.related_ce_id.name[1:]}'
    @api.depends('related_accrued_entry.state', 'related_reverse_accrued_entry.state')
    def _compute_state(self):
        for record in self:
            if not record.related_accrued_entry and not record.related_reverse_accrued_entry:
                record.state = 'draft'
            elif record.related_accrued_entry and record.related_reverse_accrued_entry:
                if record.related_accrued_entry.state == 'posted' and record.related_reverse_accrued_entry.state == 'draft':
                    record.state = 'accrued'
                elif (record.related_accrued_entry.state == 'cancel' or
                      record.related_reverse_accrued_entry.state == 'cancel'):
                    record.state = 'cancel'
                elif (record.related_accrued_entry.state == 'posted' and
                      record.related_reverse_accrued_entry.state == 'posted'):
                    record.state = 'reversed'
                else:
                    record.state = 'draft'
            else:
                record.state = 'draft'


            
        
    def _compute_display_name(self):
        for record in self:
            record.display_name = f'{record.related_ce_id.name} - {record.id}'
        
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

    def write(self, vals):
        result = super().write(vals)
        if 'accrual_account_id' in vals:
            for record in self:
                if record.accrual_account_id:
                    record.update_total_accrued_account_id()

    @api.depends('date')
    def _compute_reversal_date(self):
        for record in self:
            if not record.reversal_date or record.reversal_date <= record.date:
                record.reversal_date = record.date + relativedelta(days=1)
            else:
                record.reversal_date = record.reversal_date
    
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

    

    def update_total_accrued_account_id(self):
        """Update or create the Total Accrued line with the computed total"""
        for record in self:
            accrued_total_line = record.line_ids.filtered(lambda l: l.label == 'Total Accrued')
            accrued_total_line.write({'account_id': record.accrual_account_id})

        
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
                    
                    # Calculate weighted analytic distribution based on credit amounts
                    analytic_distribution = {}
                    total_credit = sum(credit_lines.mapped('credit'))
                    
                    if total_credit > 0:
                        # Aggregate analytic distributions weighted by credit amounts
                        analytic_totals = {}
                        
                        for line in credit_lines:
                            if line.analytic_distribution and line.credit > 0:
                                line_weight = line.credit / total_credit
                                for analytic_id, percentage in line.analytic_distribution.items():
                                    if analytic_id not in analytic_totals:
                                        analytic_totals[analytic_id] = 0
                                    analytic_totals[analytic_id] += (percentage * line_weight)
                        
                        # Round percentages and ensure they add up to 100%
                        if analytic_totals:
                            analytic_distribution = {k: round(v, 2) for k, v in analytic_totals.items()}
                            
                            # Adjust for rounding differences to ensure total = 100%
                            total_percentage = sum(analytic_distribution.values())
                            if total_percentage != 100.0 and analytic_distribution:
                                # Add the difference to the largest percentage
                                largest_key = max(analytic_distribution.keys(), key=lambda k: analytic_distribution[k])
                                analytic_distribution[largest_key] += (100.0 - total_percentage)
                    
                    # Fallback to sale order's analytic distribution if no line distributions found
                    if not analytic_distribution and record.related_ce_id:
                        if hasattr(record.related_ce_id, 'analytic_distribution') and record.related_ce_id.analytic_distribution:
                            analytic_distribution = record.related_ce_id.analytic_distribution
                    
                    accrued_total_line.write({
                        'debit': total,
                        'account_id': record.accrual_account_id.id if record.accrual_account_id else accrued_total_line.account_id.id,
                        'currency_id': record.currency_id.id,
                        'analytic_distribution': analytic_distribution,
                    })
            elif accrued_total_line and total == 0:
                # Remove the line if total is 0
                accrued_total_line.unlink()


    def sync_new_records_for_accrual(self):
        ce_records = self.env['sale.order'].search([
            ('state', '=', 'sale'), 
            ('x_studio_ce_status', 'in', ['Signed', 'Billable'])
        ])
        
        total_success = 0
        
        for ce in ce_records:
            state = ce.action_create_custom_accrued_revenue()
            if state:
                total_success += 1
        
        # Simply return the count
        return total_success

    def create_multiple_entries(self):
        for record in self:
            record.create_entries()
    
    def create_entries(self):
        """Create accrual journal entries and their automatic reversals"""
        self.ensure_one()
        
        # Validation checks
        if self.state != 'draft':
            raise UserError(_('Entries can only be created for records in "Draft" status.'))
            
        if self.reversal_date <= self.date:
            raise UserError(_('Reversal date must be posterior to date.'))
            
        if not self.line_ids:
            raise UserError(_('Cannot create entries without any revenue lines.'))
            
        if not self.journal_id:
            raise UserError(_('Please specify a journal for the accrual entries.'))
        
        # Prepare move values
        move_vals = self._prepare_move_vals()
        
        # Create and post the accrual move
        move = self.env['account.move'].create(move_vals)
        move._post()
        self.related_accrued_entry = move.id
        
        # Create automatic reversal
        reverse_move = move._reverse_moves(default_values_list=[{
            'ref': _('Reversal of: %s', move.ref),
            'name': '/',
            'date': self.reversal_date,
            'related_custom_accrued_record': self.id
        }])
        reverse_move._post()
        self.related_reverse_accrued_entry = reverse_move.id
        # Update state
        self.state = 'accrued'
        
        # Post message to related sale order
        if self.related_ce_id:
            body = _(
                'Accrual entry created on %(date)s: %(accrual_entry)s. '
                'And its reverse entry: %(reverse_entry)s.',
                date=self.date,
                accrual_entry=move._get_html_link(),
                reverse_entry=reverse_move._get_html_link(),
            )
            self.related_ce_id.message_post(body=body)
        
        return {
            'name': _('Accrual Moves'),
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('id', 'in', (move.id, reverse_move.id))],
        }
        
    def _prepare_move_vals(self):
        """Prepare the accounting move values from the accrued revenue lines"""
        self.ensure_one()
        
        # Filter out zero amount lines
        valid_lines = self.line_ids.filtered(lambda l: l.credit != 0 or l.debit != 0)
        
        if not valid_lines:
            raise UserError(_('No valid lines found to create journal entries.'))
        
        # Prepare move line values
        move_line_vals = []
        
        for line in valid_lines:
            if not line.account_id:
                raise UserError(_('Account is required for line: %s') % line.label)
            
            # Determine currency for move line - use line currency or fallback to record/company currency
            line_currency_id = line.currency_id.id if line.currency_id else (self.currency_id.id if self.currency_id else self.company_id.currency_id.id)
            
            # Prepare move line data with analytic distribution
            move_line_data = {
                'name': line.label,
                'account_id': line.account_id.id,
                'debit': line.debit,
                'credit': line.credit,
                'partner_id': self.ce_partner_id.id if self.ce_partner_id else False,
                'currency_id': line_currency_id,
            }
            
            # Add analytic distribution if it exists
            if hasattr(line, 'analytic_distribution') and line.analytic_distribution:
                move_line_data['analytic_distribution'] = line.analytic_distribution
            
            move_line_vals.append(move_line_data)
        
        # Determine currency for the move
        move_currency_id = False
        if self.currency_id:
            move_currency_id = self.currency_id.id
        else:
            move_currency_id = self.company_id.currency_id.id
        
        # Prepare the move values
        move_vals = {
            'ref': f'Accrual - {self.related_ce_id.name if self.related_ce_id else self.display_name}',
            'journal_id': self.journal_id.id,
            'date': self.date,
            'company_id': self.company_id.id,
            'currency_id': move_currency_id,
            'line_ids': [(0, 0, line_vals) for line_vals in move_line_vals],
            'related_custom_accrued_record': self.id
        }
        
        return move_vals
    
    def reverse_entries(self):
        """Manual reversal of accrued entries"""
        self.ensure_one()
        
        if self.state != 'accrued':
            raise UserError(_('Only accrued entries can be reversed.'))
        
        # Find the original accrual move
        domain = [
            ('ref', 'like', f'Accrual - {self.related_ce_id.name if self.related_ce_id else self.display_name}'),
            ('journal_id', '=', self.journal_id.id),
            ('date', '=', self.date),
            ('state', '=', 'posted'),
        ]
        
        original_moves = self.env['account.move'].search(domain)
        
        if not original_moves:
            raise UserError(_('No posted accrual entries found to reverse.'))
        
        # Update state
        self.state = 'reversed'
        
        return {
            'name': _('Accrual Entries'),
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('id', 'in', original_moves.ids)],
        }

    def action_open_journal_entries(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': 'Journal Entries',
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('related_custom_accrued_record', '=', self.id)],
        }

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
        required=True,
        readonly=True
    )

    ce_line_id = fields.Many2one('sale.order.line', string="Sale Order Line", ondelete='cascade',readonly=True)

    account_id = fields.Many2one(
        'account.account',
        string="Account",
        domain=[('deprecated', '=', False)],
        readonly=True
    )

    label = fields.Char(
        string="Label",
        required=True,
        readonly=True
    )

    debit = fields.Float(
        string="Debit",
        # currency_field='currency_id',
        default=0.0,
        readonly=True
    )

    credit = fields.Float(
        string="Credit",
        # currency_field='currency_id',
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

    # Analytic Distribution field for Odoo 18
    analytic_distribution = fields.Json(
        string="Analytic Distribution",
        help="Analytic distribution for this line"
    )
   # Required for analytic distribution widget
    analytic_precision = fields.Integer(
        string="Analytic Precision",
        compute="_compute_analytic_precision",
        readonly=True
    )

    @api.model_create_multi
    def create(self, vals_list):
        lines = super().create(vals_list)
        for line in lines:
            if line.accrued_revenue_id:
                line.accrued_revenue_id.update_total_accrued_line()
        return lines
    
    def write(self, vals):
        # Skip validation during bulk updates, validate at the end
        if len(self) > 1 and 'credit' in vals:
            result = super().write(vals)
            # Validate only once after all changes
            accrued_revenues = self.mapped('accrued_revenue_id')
            for revenue in accrued_revenues:
                revenue.update_total_accrued_line()
            return result
        else:
            # Normal single-line update
            result = super().write(vals)
            if 'credit' in vals:
                for line in self:
                    if line.accrued_revenue_id:
                        line.accrued_revenue_id.update_total_accrued_line()
                    if line.label == 'Total Accrued':
                        raise UserError("You cannot set credit amount on this line!")
            return result
    
    def unlink(self):
        accrued_revenues = self.mapped('accrued_revenue_id')
        result = super().unlink()
        for revenue in accrued_revenues:
            revenue.update_total_accrued_line()
        return result

    def _compute_analytic_precision(self):
        """Compute analytic precision for distribution calculations"""
        for record in self:
            # Default to 2 decimal places for percentage calculations
            # You can adjust this based on your company's requirements
            record.analytic_precision = 2
        
