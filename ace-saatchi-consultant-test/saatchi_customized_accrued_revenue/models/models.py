# -*- coding: utf-8 -*-
"""
Saatchi Customized Accrued Revenue Module
==========================================
This module manages accrued revenue entries for sale orders, creating journal entries
for revenue recognition and their automatic reversals.

Main Features:
- Automatic accrual generation from sale orders
- Manual override for non-standard accruals
- Adjustment entries for reducing accruals (no auto-reversal)
- Journal entry creation with proper accounting entries
- Automatic reversal entries (for normal accruals only)
- Multi-currency support
- Analytic distribution tracking
- ✓ MULTI-COMPANY SUPPORT (via saatchi.accrual_config model)

Business Logic:
- Normal accruals: Dr. Accrued Revenue (1210) | Cr. Revenue Accounts (with auto-reversal)
- Adjustment entries: Dr. Digital Income (5787) | Cr. Accrued Revenue (1210) (NO auto-reversal)
"""

from odoo import models, fields, api, _, Command
from odoo.exceptions import UserError
from dateutil.relativedelta import relativedelta
import logging

_logger = logging.getLogger(__name__)


class SaatchiAccrualConfig(models.Model):
    """
    Multi-Company Accrual Configuration
    
    Stores company-specific settings for accrued revenue management:
    - Journal for accrual entries
    - Accrued revenue account (1210)
    - Digital income account for adjustments (5787)
    """
    _name = 'saatchi.accrual_config'
    _description = 'Saatchi Accrual Configuration'
    _rec_name = 'company_id'
    
    company_id = fields.Many2one(
        'res.company',
        string='Company',
        required=True,
        unique=True,
        ondelete='cascade',
        index=True,
        default=lambda self: self.env.company,
        help='Company for which this accrual configuration applies'
    )
    
    accrued_journal_id = fields.Many2one(
        'account.journal',
        string='Accrual Journal',
        required=True,
        domain=[('type', '=', 'general')],
        help='General journal for posting accrual entries'
    )
    
    accrued_revenue_account_id = fields.Many2one(
        'account.account',
        string='Accrued Revenue Account',
        required=True,
        domain=[('deprecated', '=', False)],
        help='Account 1210 - Accrued Revenue (debited for normal accruals, credited for adjustments)'
    )
    
    digital_income_account_id = fields.Many2one(
        'account.account',
        string='Digital Income Account',
        required=True,
        domain=[('deprecated', '=', False)],
        help='Account 5787 - Digital Income (used for adjustment entries)'
    )

    opening_balance_cutoff_date = fields.Date(
        string='Opening Balance Cutoff Date',
        help='Month-end date for which opening balances should be used. '
             'When generating the accrued revenue report, if the report month '
             'immediately follows this date (i.e. previous month end = this date), '
             'the system will look up opening balances from the Opening Balance model '
             'for any CE# that has no prior DB transaction history.\n\n'
             'Example: Set to 2025-12-31 so that when generating January 2026, '
             'the Dec 31 balance column will use imported opening balances.'
    )

    @api.constrains('accrued_journal_id')
    def _check_journal_company(self):
        """Ensure journal belongs to the configured company"""
        for record in self:
            if record.accrued_journal_id.company_id != record.company_id:
                raise UserError(
                    _('The accrual journal must belong to the selected company.')
                )
    
    @api.constrains('accrued_revenue_account_id')
    def _check_accrued_revenue_account_company(self):
        """Ensure accrued revenue account is available for the configured company"""
        for record in self:
            if record.company_id not in record.accrued_revenue_account_id.company_ids:
                raise UserError(
                    _('The accrued revenue account must be available for the selected company.')
                )
    
    @api.constrains('digital_income_account_id')
    def _check_digital_income_account_company(self):
        """Ensure digital income account is available for the configured company"""
        for record in self:
            if record.company_id not in record.digital_income_account_id.company_ids:
                raise UserError(
                    _('The digital income account must be available for the selected company.')
                )
    
    @api.model_create_multi
    def create(self, vals_list):
        """Override create to ensure one config per company"""
        for vals in vals_list:
            existing = self.search([('company_id', '=', vals.get('company_id'))], limit=1)
            if existing:
                raise UserError(
                    _('Configuration for this company already exists. Please edit the existing record.')
                )
        return super().create(vals_list)


class SaatchiCustomizedAccruedRevenue(models.Model):
    """
    Accrued Revenue Management
    
    Handles creation and tracking of accrued revenue entries including:
    - Normal accruals (Dr. Accrued Revenue | Cr. Revenue)
    - Adjustment entries (Dr. Digital Income | Cr. Accrued Revenue) - NO auto-reversal
    """
    _name = 'saatchi.accrued_revenue'
    _description = 'Saatchi Customized Accrued Revenue'
    _rec_name = 'display_name'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    
    # ========== Display & Identification ==========
    display_name = fields.Char(
        compute="_compute_display_name",
        string="Display Name"
    )
    
    # ========== Related Sale Order ==========
    x_related_ce_id = fields.Many2one(
        'sale.order',
        string="Sale Order",
        readonly=True,
        index=True
    )

    ce_partner_id = fields.Many2one(
        'res.partner',
        string="Customer",
        compute="_compute_ce_fields",
        store=True,
        readonly=True
    )

    ce_status = fields.Selection(
        selection=[
            ('for_client_signature', 'For Client Signature'),
            ('signed', 'Signed'),
            ('billable', 'Billable'),
            ('closed', 'Closed'),
            ('cancelled', 'Cancelled')
        ],
        string="CE Status",
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

    ce_code = fields.Char(
        string='CE Code',
        compute="_compute_ce_code",
        store=True
    )

    old_ce_code = fields.Char(
        string='Old CE Code',
        compute="_compute_old_ce_code",
        store=True,
        help="Legacy CE code from Studio field (x_studio_old_ce) on the related sale order"
    )

    old_ce_date = fields.Date(
        string='Old CE Date',
        compute="_compute_old_ce_date",
        store=True,
        help="Legacy CE date from Studio field (x_studio_old_ce_date) on the related sale order"
    )

    effective_ce_code = fields.Char(
        string='Effective CE Code',
        compute="_compute_effective_ce_code",
        store=True,
        help="Old CE code if available, otherwise the standard CE code. Used for display and matching."
    )

    ce_original_total_amount = fields.Monetary(
        string="Total CE Maximum Amount Subject for Accrual",
        currency_field="currency_id",
        store=True,
        help="Original total amount eligible for accrual from the sale order"
    )

    remarks = fields.Char(
        string='Remarks',
        store=True,
        tracking=True
    )

    # ========== Revenue Lines ==========
    line_ids = fields.One2many(
        'saatchi.accrued_revenue_lines',
        'accrued_revenue_id',
        string="Revenue Lines"
    )

    # ========== Accounting Information ==========
    journal_id = fields.Many2one(
        'account.journal',
        string="Journal",
        default=lambda self: self._get_accrued_journal_id(),
        required=True
    )

    accrual_account_id = fields.Many2one(
        'account.account',
        string="Accrual Account",
        default=lambda self: self._get_accrued_revenue_account_id(),
        required=True,
        help="Account 1210 - Accrued Revenue. Debited for normal accruals, credited for adjustments"
    )
    
    digital_income_account_id = fields.Many2one(
        'account.account',
        string="Digital Income Account",
        default=lambda self: self._get_adjustment_accrued_revenue_account_id(),
        help="Account used for adjustment entries (Dr. side)"
    )

    date = fields.Date(
        string="Accrual Date",
        default=lambda self: self._default_accrual_date(),
        required=True,
        tracking=True
    )
    
    reversal_date = fields.Date(
        string="Reversal Date",
        default=lambda self: self._default_reversal_date(),
        # required=True,
        tracking=True
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
        store=True,
        help="Sum of all credit lines (excluding Total Accrued line)"
    )
    
    # ========== Accrual Type ==========
    is_adjustment_entry = fields.Boolean(
        string="Is Adjustment Entry",
        default=False,
        readonly=True,
        help="True if this is an adjustment entry (Dr. Digital Income | Cr. Accrued Revenue) - NO auto-reversal"
    )
    
    # ========== State & Journal Entries ==========
    state = fields.Selection(
        [
            ('draft', 'Draft'),
            ('accrued', 'Accrued'),
            ('reversed', 'Reversed'),
            ('cancel', 'Cancelled')
        ],
        string="Status",
        default='draft',
        required=True,
        store=True,
        compute="_compute_state",
        tracking=True
    )
    
    x_accrual_system_generated = fields.Boolean(
        string="Is System Generated?",
        default=True,
        help="Indicates if this accrual was generated automatically by the system"
    )
    
    related_accrued_entry = fields.Many2one(
        'account.move',
        readonly=True,
        string="Accrued Entry"
    )
    
    related_reverse_accrued_entry = fields.Many2one(
        'account.move',
        readonly=True,
        string="Reverse Accrue Entry"
    )

    company_id = fields.Many2one(
        'res.company',
        string="Company",
        required=True,
        related="x_related_ce_id.company_id"
    )

    total_amount_accrued = fields.Monetary(
        string="Total Accrued Amount",
        compute="_compute_total_amount_accrued",
        currency_field="currency_id",
        store=True,
        help="Sum of all credit lines (excluding Total Accrued line)"
    )
    # ========== Compute Methods ==========
                
    @api.depends('x_related_ce_id', 'x_related_ce_id.x_ce_code')
    def _compute_ce_code(self):
        """Compute CE code from related sale order"""
        for record in self:
            record.ce_code = record.x_related_ce_id.x_ce_code if record.x_related_ce_id else False

    @api.depends('x_related_ce_id')
    def _compute_old_ce_code(self):
        """Compute old CE code from Studio field on related sale order"""
        for record in self:
            if record.x_related_ce_id:
                record.old_ce_code = getattr(record.x_related_ce_id, 'x_studio_old_ce', False)
            else:
                record.old_ce_code = False

    # @api.depends('x_related_ce_id', 'x_related_ce_id.x_studio_old_ce_date')
    @api.depends('x_related_ce_id')
    def _compute_old_ce_date(self):
        """Compute old CE date from Studio field on related sale order"""
        for record in self:
            if record.x_related_ce_id:
                record.old_ce_date = getattr(record.x_related_ce_id, 'x_studio_old_ce_date', False)
            else:
                record.old_ce_date = False

    @api.depends('old_ce_code', 'ce_code')
    def _compute_effective_ce_code(self):
        """Compute effective CE code: old CE code takes priority over new CE code"""
        for record in self:
            record.effective_ce_code = record.old_ce_code or record.ce_code or False

    @api.depends('related_accrued_entry.state', 'related_reverse_accrued_entry.state')
    def _compute_state(self):
        """
        Compute state based on journal entry states
        
        State Logic:
        - draft: No entries OR entries cleared after cancellation
        - accrued: Accrual entry posted (and reversal in draft if exists)
        - reversed: Both entries posted (only for normal accruals with reversal)
        - cancel: Any entry cancelled (links still exist)
        """
        # raise UserError('huh')
        for record in self:
            # Adjustment entries don't have reversal entries
            if record.is_adjustment_entry:
                if not record.related_accrued_entry:
                    record.state = 'draft'
                elif record.related_accrued_entry.state == 'cancel':
                    record.state = 'cancel'
                elif record.related_accrued_entry.state == 'posted':
                    record.state = 'accrued'
                else:
                    record.state = 'draft'
            else:
                # Normal accruals (with reversal)
                if not record.related_accrued_entry and not record.related_reverse_accrued_entry:
                    record.state = 'draft'
                elif record.related_accrued_entry and record.related_reverse_accrued_entry:
                    accrual_state = record.related_accrued_entry.state
                    reversal_state = record.related_reverse_accrued_entry.state
                    
                    if accrual_state == 'cancel' or reversal_state == 'cancel':
                        record.state = 'cancel'
                    elif accrual_state == 'posted' and reversal_state == 'posted':
                        record.state = 'reversed'
                    elif accrual_state == 'posted' and reversal_state == 'draft':
                        record.state = 'accrued'
                    elif accrual_state == 'posted' and not reversal_state:
                        record.state = 'accrued'
                    else:
                        record.state = 'draft'

                else:
                    record.state = 'draft'
            
    def _compute_display_name(self):
        """Generate display name from sale order and record ID"""
        for record in self:
            so_name = record.x_related_ce_id.name if record.x_related_ce_id else 'New'
            suffix = ' | [ADJ] ' if record.is_adjustment_entry else ''
            ce_display = record.effective_ce_code or record.ce_code or ''
            record.display_name = f'{so_name} | {ce_display}{suffix}'

    @api.depends('line_ids.debit')
    def _compute_total_amount_accrued(self):
        """Calculate total debit amount from all credit lines (excluding Total Accrued)"""
        for record in self:
            debit_lines = record.line_ids.filtered(lambda l: l.label == 'Total Accrued')
            record.total_amount_accrued = sum(debit_lines.mapped('debit'))
            
    @api.depends('line_ids.credit')
    def _compute_total_debit_in_accrue_account(self):
        """Calculate total debit amount from all credit lines (excluding Total Accrued)"""
        for record in self:
            credit_lines = record.line_ids.filtered(lambda l: l.label != 'Total Accrued')
            record.total_debit_in_accrue_account = sum(credit_lines.mapped('credit'))
    
    @api.depends("x_related_ce_id")
    def _compute_ce_fields(self):
        """Compute fields from related sale order"""
        for rec in self:
            if rec.x_related_ce_id:
                rec.ce_partner_id = rec.x_related_ce_id.partner_id.id or False
                rec.ce_status = rec.x_related_ce_id.x_ce_status or False
                rec.ce_job_description = rec.x_related_ce_id.x_job_description or False
            else:
                rec.ce_partner_id = False
                rec.ce_status = False
                rec.ce_job_description = False

    # ========== Default Methods ==========
    
    def _default_accrual_date(self):
        """Default to last day of previous month"""
        today = fields.Date.context_today(self)
        first_of_current_month = today.replace(day=1)
        return first_of_current_month - relativedelta(days=1)
    
    def _default_reversal_date(self):
        """Default to first day of current month"""
        today = fields.Date.context_today(self)
        return today.replace(day=1)

        
    def _get_accrued_revenue_account_id(self):
        """Get accrued revenue account ID from company-specific configuration"""
        ce_company_id = self.env.context.get('default_company_id')
        target_company = self.env['res.company'].browse(ce_company_id) if ce_company_id else self.env.company
        
        try:
            config = self.env['saatchi.accrual_config'].search(
                [('company_id', '=', target_company.id)],
                limit=1
            )
            
            if config and config.accrued_revenue_account_id:
                return config.accrued_revenue_account_id.id
            
            _logger.warning(
                f'No accrual configuration found for company {target_company.name}. '
                'Please configure accrual settings in Accounting > Configuration > Accrual Configuration.'
            )
            return 0
            
        except Exception as e:
            _logger.error(f'Error retrieving accrued revenue account: {str(e)}')
            return 0

    def _get_accrued_journal_id(self):
        """Get accrued journal ID from company-specific configuration"""
        ce_company_id = self.env.context.get('default_company_id')
        target_company = self.env['res.company'].browse(ce_company_id) if ce_company_id else self.env.company
        
        try:
            config = self.env['saatchi.accrual_config'].search(
                [('company_id', '=', target_company.id)],
                limit=1
            )
            
            if config and config.accrued_journal_id:
                return config.accrued_journal_id.id
            
            _logger.warning(
                f'No accrual configuration found for company {target_company.name}. '
                'Please configure accrual settings in Accounting > Configuration > Accrual Configuration.'
            )
            return 0
            
        except Exception as e:
            _logger.error(f'Error retrieving accrued journal: {str(e)}')
            return 0
        
    def _get_adjustment_accrued_revenue_account_id(self):
        """Get digital income account ID from company-specific configuration"""
        ce_company_id = self.env.context.get('default_company_id')
        target_company = self.env['res.company'].browse(ce_company_id) if ce_company_id else self.env.company
        
        try:
            config = self.env['saatchi.accrual_config'].search(
                [('company_id', '=', target_company.id)],
                limit=1
            )
            
            if config and config.digital_income_account_id:
                return config.digital_income_account_id.id
            
            _logger.warning(
                f'No accrual configuration found for company {target_company.name}. '
                'Please configure accrual settings in Accounting > Configuration > Accrual Configuration.'
            )
            return 0
            
        except Exception as e:
            _logger.error(f'Error retrieving digital income account: {str(e)}')
            return 0

    # ========== CRUD Operations ==========
    
    def write(self, vals):
        """Override write to update Total Accrued line when accrual account changes"""
        result = super().write(vals)
        if 'accrual_account_id' in vals:
            for record in self:
                if record.accrual_account_id:
                    record.update_total_accrued_account_id()
        return result

    # ========== Business Logic Methods ==========

    def update_total_accrued_account_id(self):
        """Update the account_id of Total Accrued line to match accrual_account_id"""
        for record in self:
            accrued_total_line = record.line_ids.filtered(lambda l: l.label == 'Total Accrued')
            if accrued_total_line:
                # Use with_context to prevent recursion
                accrued_total_line.with_context(skip_total_update=True).write({
                    'account_id': record.accrual_account_id.id
                })
            
    def update_total_accrued_line(self):
        """
        Update or create the Total Accrued line with computed total
        
        For normal accruals: Balance debit/credit lines properly
        For adjustments: Flexible based on which side has amounts
            - If debit lines exist: Cr. Total Accrued (reducing accrual)
            - If credit lines exist: Dr. Total Accrued (increasing accrual)
        
        Handles negative amounts (reversals/returns) by properly balancing entries
        """
        for record in self:
            # Refresh to get latest lines (prevents duplicate detection issues)
            record.line_ids.invalidate_recordset(['label'])
            
            # Find the Total Accrued line first
            accrued_total_line = record.line_ids.filtered(lambda l: l.label == 'Total Accrued')
            
            if record.is_adjustment_entry:
                # Adjustment: support both directions
                other_lines = record.line_ids.filtered(lambda l: l.label != 'Total Accrued')
                total_debit = sum(other_lines.mapped('debit'))
                total_credit = sum(other_lines.mapped('credit'))
                
                # Determine direction based on which side has amounts
                if total_debit > 0 and total_credit == 0:
                    # Dr. Digital Income | Cr. Accrued Revenue (reducing accrual)
                    total = total_debit
                    debit_amount = 0.0
                    credit_amount = total
                    analytic_distribution = self._calculate_weighted_analytic_distribution(other_lines)
                elif total_credit > 0 and total_debit == 0:
                    # Dr. Accrued Revenue | Cr. Digital Income (increasing accrual)
                    total = total_credit
                    debit_amount = total
                    credit_amount = 0.0
                    analytic_distribution = self._calculate_weighted_analytic_distribution(other_lines)
                elif total_debit > 0 and total_credit > 0:
                    # Mixed entries - net the amounts
                    net_amount = abs(total_debit - total_credit)
                    if total_debit > total_credit:
                        # Net debit - so credit Total Accrued
                        debit_amount = 0.0
                        credit_amount = net_amount
                    else:
                        # Net credit - so debit Total Accrued
                        debit_amount = net_amount
                        credit_amount = 0.0
                    total = net_amount
                    analytic_distribution = self._calculate_weighted_analytic_distribution(other_lines)
                else:
                    # No amounts - delete Total Accrued line if it exists
                    if accrued_total_line:
                        accrued_total_line.with_context(skip_total_update=True).unlink()
                    continue
                
                # VALIDATION: Check against CE original amount
                if record.ce_original_total_amount and abs(total) > abs(record.ce_original_total_amount):
                    raise UserError(_("Total accrued amount cannot exceed the original CE amount."))
                
                if not analytic_distribution and record.x_related_ce_id:
                    if hasattr(record.x_related_ce_id, 'analytic_distribution') and record.x_related_ce_id.analytic_distribution:
                        analytic_distribution = record.x_related_ce_id.analytic_distribution
                
                if total != 0:
                    if accrued_total_line:
                        # Ensure we only have ONE Total Accrued line
                        if len(accrued_total_line) > 1:
                            (accrued_total_line[1:]).with_context(skip_total_update=True).unlink()
                            accrued_total_line = accrued_total_line[0]
                        
                        # Update existing line
                        accrued_total_line.with_context(skip_total_update=True).write({
                            'debit': debit_amount,
                            'credit': credit_amount,
                            'account_id': record.accrual_account_id.id if record.accrual_account_id else accrued_total_line.account_id.id,
                            'currency_id': record.currency_id.id,
                            'analytic_distribution': analytic_distribution,
                        })
                    else:
                        # Double-check one more time before creating (in case of race condition)
                        existing = self.env['saatchi.accrued_revenue_lines'].search([
                            ('accrued_revenue_id', '=', record.id),
                            ('label', '=', 'Total Accrued')
                        ])
                        if existing:
                            # Update existing instead of creating
                            existing.with_context(skip_total_update=True).write({
                                'debit': debit_amount,
                                'credit': credit_amount,
                                'account_id': record.accrual_account_id.id,
                                'currency_id': record.currency_id.id,
                                'analytic_distribution': analytic_distribution,
                            })
                        else:
                            # Create new Total Accrued line
                            self.env['saatchi.accrued_revenue_lines'].with_context(skip_total_update=True).create({
                                'accrued_revenue_id': record.id,
                                'label': 'Total Accrued',
                                'account_id': record.accrual_account_id.id,
                                'debit': debit_amount,
                                'credit': credit_amount,
                                'currency_id': record.currency_id.id,
                                'analytic_distribution': analytic_distribution,
                                'sequence': 999,
                            })
                elif accrued_total_line:
                    # Delete if total is 0
                    accrued_total_line.with_context(skip_total_update=True).unlink()
                    
            else:
                # Normal accrual: Balance the debit/credit lines
                other_lines = record.line_ids.filtered(lambda l: l.label != 'Total Accrued')
                total_debit = sum(other_lines.mapped('debit'))
                total_credit = sum(other_lines.mapped('credit'))
                net_amount = total_credit - total_debit  # Net credit amount (positive = more credits, negative = more debits)
                
                if net_amount != 0:
                    # VALIDATION: Check against CE original amount
                    if record.ce_original_total_amount and abs(net_amount) > abs(record.ce_original_total_amount):
                        raise UserError(_("Total accrued amount cannot exceed the original CE amount."))
                    
                    analytic_distribution = self._calculate_weighted_analytic_distribution(other_lines)
                    
                    if not analytic_distribution and record.x_related_ce_id:
                        if hasattr(record.x_related_ce_id, 'analytic_distribution') and record.x_related_ce_id.analytic_distribution:
                            analytic_distribution = record.x_related_ce_id.analytic_distribution
                    
                    # Balance the entry properly
                    # If net_amount > 0: More credits than debits -> Debit Total Accrued to balance
                    # If net_amount < 0: More debits than credits -> Credit Total Accrued to balance
                    if net_amount > 0:
                        debit_amount = net_amount
                        credit_amount = 0.0
                    else:
                        debit_amount = 0.0
                        credit_amount = abs(net_amount)
                    
                    if accrued_total_line:
                        # Ensure we only have ONE Total Accrued line
                        if len(accrued_total_line) > 1:
                            (accrued_total_line[1:]).with_context(skip_total_update=True).unlink()
                            accrued_total_line = accrued_total_line[0]
                        
                        # Update existing line
                        accrued_total_line.with_context(skip_total_update=True).write({
                            'debit': debit_amount,
                            'credit': credit_amount,
                            'account_id': record.accrual_account_id.id if record.accrual_account_id else accrued_total_line.account_id.id,
                            'currency_id': record.currency_id.id,
                            'analytic_distribution': analytic_distribution,
                        })
                    else:
                        # Double-check before creating
                        existing = self.env['saatchi.accrued_revenue_lines'].search([
                            ('accrued_revenue_id', '=', record.id),
                            ('label', '=', 'Total Accrued')
                        ])
                        if existing:
                            existing.with_context(skip_total_update=True).write({
                                'debit': debit_amount,
                                'credit': credit_amount,
                                'account_id': record.accrual_account_id.id,
                                'currency_id': record.currency_id.id,
                                'analytic_distribution': analytic_distribution,
                            })
                        else:
                            # Create new Total Accrued line
                            self.env['saatchi.accrued_revenue_lines'].with_context(skip_total_update=True).create({
                                'accrued_revenue_id': record.id,
                                'label': 'Total Accrued',
                                'account_id': record.accrual_account_id.id,
                                'debit': debit_amount,
                                'credit': credit_amount,
                                'currency_id': record.currency_id.id,
                                'analytic_distribution': analytic_distribution,
                                'sequence': 999,
                            })
                elif accrued_total_line:
                    # Delete if total is 0
                    accrued_total_line.with_context(skip_total_update=True).unlink()
        
    def _calculate_weighted_analytic_distribution(self, lines):
        """
        Calculate weighted analytic distribution from lines
        
        Args:
            lines: Recordset of accrual lines
            
        Returns:
            dict: Weighted analytic distribution with percentages summing to 100.0
        """
        analytic_distribution = {}
        
        # For adjustment entries, use debit; for normal, use credit
        if self.is_adjustment_entry:
            total_amount = sum(lines.mapped('debit'))
        else:
            total_amount = sum(lines.mapped('credit'))
        
        if total_amount > 0:
            analytic_totals = {}
            
            for line in lines:
                if line.analytic_distribution:
                    line_amount = line.debit if self.is_adjustment_entry else line.credit
                    if line_amount > 0:
                        line_weight = line_amount / total_amount
                        for analytic_id, percentage in line.analytic_distribution.items():
                            if analytic_id not in analytic_totals:
                                analytic_totals[analytic_id] = 0
                            analytic_totals[analytic_id] += (percentage * line_weight)
            
            if analytic_totals:
                analytic_distribution = {k: round(v, 2) for k, v in analytic_totals.items()}
                
                # Adjust rounding errors
                total_percentage = sum(analytic_distribution.values())
                if total_percentage != 100.0 and analytic_distribution:
                    largest_key = max(analytic_distribution.keys(), key=lambda k: analytic_distribution[k])
                    analytic_distribution[largest_key] += (100.0 - total_percentage)
        
        return analytic_distribution
    
    def sync_new_records_for_accrual(self):
        """
        Collect potential accruals and show wizard for selection
        
        Returns:
            dict: Action to open wizard or notification
        """
        accrual_date = self._default_accrual_date()
        reversal_date = self._default_reversal_date()
        
        potential_sos, duplicate_sos, client_sig_so_ids = self.env['sale.order'].collect_potential_accruals(
            accrual_date=accrual_date,
            reversal_date=reversal_date
        )
        
        if not potential_sos:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('No Sale Orders Found'),
                    'message': _('No sale orders found that meet accrual criteria (state=sale, status=signed/billable).'),
                    'type': 'warning',
                    'sticky': False,
                }
            }
        
        wizard = self.env['saatchi.accrued_revenue.wizard'].create({
            'accrual_date': accrual_date,
            'reversal_date': reversal_date,
        })
        
        for so in potential_sos:
            amount_total = so._calculate_accrual_amount(accrual_date=accrual_date)
            if not amount_total:
                continue
            has_duplicate = so in duplicate_sos
            is_client_sig = so.id in client_sig_so_ids
            self.env['saatchi.accrued_revenue.wizard.line'].create({
                'wizard_id': wizard.id,
                'sale_order_id': so.id,
                'has_existing_accrual': has_duplicate,
                'amount_total': amount_total,
                'create_accrual': not has_duplicate,
                'is_from_reversal_ob': is_client_sig,
            })
        
        wizard_name = _('Generate Accrued Revenues - Duplicates Found') if duplicate_sos else _('Generate Accrued Revenues')
        
        return {
            'type': 'ir.actions.act_window',
            'name': wizard_name,
            'res_model': 'saatchi.accrued_revenue.wizard',
            'view_mode': 'form',
            'res_id': wizard.id,
            'views': [(False, 'form')],
            'target': 'new',
            'context': self.env.context,
        }

    def create_multiple_entries(self):
        """Create journal entries for multiple accrual records"""
        for record in self:
            record.create_entries()
    
    def create_entries(self):
        """
        Create accrual journal entries
        
        For normal accruals: Creates entry + automatic reversal
        For adjustment entries: Creates single entry (NO reversal)
        
        Returns:
            dict: Action to open created journal entries
        """
        self.ensure_one()
        
        if self.state != 'draft':
            raise UserError(_('Entries can only be created for records in "Draft" status.'))
            
        if not self.is_adjustment_entry and self.reversal_date <= self.date:
            raise UserError(_('Reversal date must be after accrual date.'))
            
        if not self.line_ids:
            raise UserError(_('Cannot create entries without any revenue lines.'))
            
        if not self.journal_id:
            raise UserError(_('Please specify a journal for the accrual entries.'))
        
        # Create accrual entry
        move_vals = self._prepare_move_vals()
        move = self.env['account.move'].create(move_vals)
        move._post()
        self.related_accrued_entry = move.id
        
        # Only create reversal for NORMAL accruals (not adjustment entries)
        if not self.is_adjustment_entry:
            reverse_move = move._reverse_moves(default_values_list=[{
                'ref': _('Reversal of: %s', move.ref),
                'name': '/',
                'date': self.reversal_date,
                'x_related_custom_accrued_record': self.id,
                'x_accrual_system_generated': self.x_accrual_system_generated,  # Pass through system flag
            }])
            reverse_move._post()
            self.related_reverse_accrued_entry = reverse_move.id
        
        # self.state = 'accrued'
        
        # Post message to sale order
        if self.x_related_ce_id:
            if self.is_adjustment_entry:
                body = _(
                    'Adjustment entry created on %(date)s: %(accrual_entry)s',
                    date=self.date,
                    accrual_entry=move._get_html_link(),
                )
            else:
                body = _(
                    'Accrual entry created on %(date)s: %(accrual_entry)s. '
                    'And its reverse entry: %(reverse_entry)s.',
                    date=self.date,
                    accrual_entry=move._get_html_link(),
                    reverse_entry=self.related_reverse_accrued_entry._get_html_link(),
                )
            self.x_related_ce_id.message_post(body=body)
        
        move_ids = [move.id]
        if self.related_reverse_accrued_entry:
            move_ids.append(self.related_reverse_accrued_entry.id)
        
        return {
            'name': _('Accrual Moves'),
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('id', 'in', move_ids)],
        }
            
    def _prepare_move_vals(self):
        """
        Prepare accounting move values from accrued revenue lines
        
        Returns:
            dict: Values for creating account.move
        """
        self.ensure_one()
        
        valid_lines = self.line_ids.filtered(lambda l: l.credit != 0 or l.debit != 0)
        
        if not valid_lines:
            raise UserError(_('No valid lines found to create journal entries.'))
        
        move_line_vals = []
        company_currency = self.company_id.currency_id
        
        for line in valid_lines:
            if not line.account_id:
                raise UserError(_('Account is required for line: %s') % line.label)
            
            line_currency = line.currency_id if line.currency_id else self.currency_id
            
            # Use old CE code if available, otherwise use CE code
            effective_ce = self.old_ce_code or self.ce_code

            move_line_data = {
                'name': line.label,
                'account_id': line.account_id.id,
                'partner_id': self.ce_partner_id.id if self.ce_partner_id else False,
                'x_ce_code': effective_ce,
                'x_ce_date': self.x_related_ce_id.date_order if self.x_related_ce_id else False,
                'x_remarks': self.remarks,
            }
            
            # Handle foreign currency
            if line_currency and line_currency != company_currency:
                debit_company = line_currency._convert(
                    line.debit,
                    company_currency,
                    self.company_id,
                    self.date
                )
                credit_company = line_currency._convert(
                    line.credit,
                    company_currency,
                    self.company_id,
                    self.date
                )
                
                move_line_data.update({
                    'debit': debit_company,
                    'credit': credit_company,
                    'currency_id': line_currency.id,
                    'amount_currency': line.debit - line.credit,
                })
            else:
                move_line_data.update({
                    'debit': line.debit,
                    'credit': line.credit,
                    'currency_id': company_currency.id,
                })
            
            if hasattr(line, 'analytic_distribution') and line.analytic_distribution:
                move_line_data['analytic_distribution'] = line.analytic_distribution
            
            move_line_vals.append((move_line_data, line.label))

        # ── Fix currency conversion rounding discrepancy ──
        # When multiple lines are individually converted from foreign currency,
        # the sum of rounded converted amounts may differ by a few cents.
        # Adjust the "Total Accrued" balancing line to absorb the difference.
        total_debit = sum(vals['debit'] for vals, _ in move_line_vals)
        total_credit = sum(vals['credit'] for vals, _ in move_line_vals)
        rounding_diff = company_currency.round(total_debit - total_credit)

        if rounding_diff != 0:
            # Find the "Total Accrued" line (the balancing entry)
            total_accrued_idx = None
            for idx, (vals, label) in enumerate(move_line_vals):
                if label == 'Total Accrued':
                    total_accrued_idx = idx
                    break

            if total_accrued_idx is not None:
                ta_vals = move_line_vals[total_accrued_idx][0]
                if rounding_diff > 0:
                    # More debit than credit → increase credit on Total Accrued
                    ta_vals['credit'] = company_currency.round(
                        ta_vals['credit'] + rounding_diff)
                else:
                    # More credit than debit → increase debit on Total Accrued
                    ta_vals['debit'] = company_currency.round(
                        ta_vals['debit'] + abs(rounding_diff))

        # Strip the label helper, keep only the vals dicts
        move_line_vals = [vals for vals, _ in move_line_vals]

        move_currency_id = self.currency_id.id if self.currency_id else company_currency.id
        
        ref_prefix = 'Adjustment - ' if self.is_adjustment_entry else 'Accrual - '
        effective_ce = self.old_ce_code or self.ce_code
    
        move_vals = {
            'ref': f'{ref_prefix}{effective_ce if self.x_related_ce_id else self.display_name}',
            'journal_id': self.journal_id.id,
            'partner_id': self.ce_partner_id.id if self.ce_partner_id else False,
            'date': self.date,
            'company_id': self.company_id.id,
            'currency_id': move_currency_id,
            'line_ids': [(0, 0, line_vals) for line_vals in move_line_vals],
            'x_related_custom_accrued_record': self.id,
            'x_remarks': self.remarks,
            'x_accrual_system_generated': self.x_accrual_system_generated,  # This will be used in compute
        }
        
        return move_vals

    def action_open_journal_entries(self):
        """Open journal entries related to this accrual"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Journal Entries'),
            'res_model': 'account.move',
            'view_mode': 'list,form',
            'domain': [('x_related_custom_accrued_record', '=', self.id)],
        }

    def action_reset_and_cancel(self):
        """
        Reset accrual to draft and cancel related journal entries
        
        Returns:
            bool: True if successful
        """
        for record in self:
            if record.related_reverse_accrued_entry:
                try:
                    if record.related_reverse_accrued_entry.state == 'posted':
                        record.related_reverse_accrued_entry.button_draft()
                    if record.related_reverse_accrued_entry.state == 'draft':
                        record.related_reverse_accrued_entry.button_cancel()
                except Exception as e:
                    raise UserError(_(
                        'Failed to cancel reversal entry: %s\n'
                        'You may need to manually unreconcile or remove constraints.'
                    ) % str(e))
            
            if record.related_accrued_entry:
                try:
                    if record.related_accrued_entry.state == 'posted':
                        record.related_accrued_entry.button_draft()
                    if record.related_accrued_entry.state == 'draft':
                        record.related_accrued_entry.button_cancel()
                except Exception as e:
                    raise UserError(_(
                        'Failed to cancel accrual entry: %s\n'
                        'You may need to manually unreconcile or remove constraints.'
                    ) % str(e))
            
            record.message_post(
                body=_('Accrual reset. Journal entries cancelled. State will update to "Cancelled".'),
                subject=_('Accrual Reset')
            )
        
        return True
    
    def action_reset_and_clear_links(self):
        """
        Reset accrual, cancel entries, and clear all links (allows deletion)
        
        Returns:
            bool: True if successful
        """
        self.action_reset_and_cancel()
        
        for record in self:
            record.write({
                'related_accrued_entry': False,
                'related_reverse_accrued_entry': False,
            })
            
            record.message_post(
                body=_('Accrual links cleared. Record can now be deleted.'),
                subject=_('Links Cleared')
            )
        
        return True
    
    def action_cancel_and_replace(self):
        """
        Cancel existing accrual and create a new one to replace it
        
        Returns:
            dict: Action to open the new accrual record
        """
        self.ensure_one()
        
        if not self.x_related_ce_id:
            raise UserError(_('No related sale order found. Cannot replace accrual.'))
        
        so_id = self.x_related_ce_id.id
        accrual_date = self.date
        reversal_date = self.reversal_date
        old_display_name = self.display_name
        
        self.action_reset_and_cancel()
        
        new_accrual_id = self.x_related_ce_id.action_create_custom_accrued_revenue(
            is_override=True,
            accrual_date=accrual_date,
            reversal_date=reversal_date
        )
        
        if not new_accrual_id:
            raise UserError(_(
                'Failed to create replacement accrual. '
                'The sale order may not have eligible lines.'
            ))
        
        self.message_post(
            body=_('Accrual replaced with new record: %s') % new_accrual_id,
            subject=_('Accrual Replaced')
        )
        
        new_accrual = self.env['saatchi.accrued_revenue'].browse(new_accrual_id)
        new_accrual.message_post(
            body=_('This accrual replaces cancelled record: %s') % old_display_name,
            subject=_('Replacement Accrual')
        )
        
        return {
            'type': 'ir.actions.act_window',
            'name': _('Replacement Accrual'),
            'res_model': 'saatchi.accrued_revenue',
            'view_mode': 'form',
            'res_id': new_accrual_id,
            'target': 'current',
        }


class SaatchiCustomizedAccruedRevenueLines(models.Model):
    """
    Accrued Revenue Lines
    
    Individual line items for accrued revenue:
    - Normal accruals: Credit lines (Revenue) + Debit line (Total Accrued)
    - Adjustment entries: Debit line (Digital Income) + Credit line (Total Accrued)
    """
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
        readonly=True,
        index=True
    )

    ce_line_id = fields.Many2one(
        'sale.order.line',
        string="Sale Order Line",
        ondelete='cascade',
        readonly=True
    )

    account_id = fields.Many2one(
        'account.account',
        string="Account",
        domain=[('deprecated', '=', False)],
        readonly=True,
        required=True
    )

    label = fields.Char(
        string="Label",
        required=True,
        readonly=True
    )

    debit = fields.Float(
        string="Debit",
        default=0.0
    )

    credit = fields.Float(
        string="Credit",
        default=0.0
    )

    currency_id = fields.Many2one(
        'res.currency',
        string="Currency",
        required=True,
        default=lambda self: self.env.company.currency_id
    )

    company_id = fields.Many2one(
        'res.company',
        string="Company",
        required=True,
        related="accrued_revenue_id.x_related_ce_id.company_id"
    )

    analytic_distribution = fields.Json(
        string="Analytic Distribution",
        help="Analytic distribution for this line"
    )
    
    analytic_precision = fields.Integer(
        string="Analytic Precision",
        compute="_compute_analytic_precision",
        readonly=True
    )

    state = fields.Selection(
        [
            ('draft', 'Draft'),
            ('accrued', 'Accrued'),
            ('reversed', 'Reversed'),
            ('cancel', 'Cancelled')
        ],
        related="accrued_revenue_id.state",
        string="Status",
        store=True
    )

    @api.depends('analytic_distribution')
    def _compute_analytic_precision(self):
        """Set analytic precision to 2 decimal places"""
        for record in self:
            record.analytic_precision = 2

    @api.model_create_multi
    def create(self, vals_list):
        """Override create to update Total Accrued line after creation"""
        lines = super().create(vals_list)
        
        # Only update if not creating the "Total Accrued" line itself
        # Group by accrued_revenue_id to batch process
        accrued_revenues = lines.mapped('accrued_revenue_id')
        for revenue in accrued_revenues:
            # Only update if we're not in the middle of updating
            if not self.env.context.get('skip_total_update'):
                revenue.update_total_accrued_line()
        
        return lines
    
    def write(self, vals):
        """Override write to update Total Accrued line when amounts change"""
        # Skip recursion if we're updating from update_total_accrued_line
        if self.env.context.get('skip_total_update'):
            return super().write(vals)
        
        result = super().write(vals)
        
        # Update Total Accrued line if debit or credit changed
        if 'credit' in vals or 'debit' in vals or 'analytic_distribution' in vals:
            accrued_revenues = self.mapped('accrued_revenue_id')
            for revenue in accrued_revenues:
                revenue.update_total_accrued_line()
        
        return result
    
    def unlink(self):
        """Override unlink to update Total Accrued line after deletion"""
        # Skip if we're in the middle of updating
        if self.env.context.get('skip_total_update'):
            return super().unlink()
        
        accrued_revenues = self.mapped('accrued_revenue_id')
        result = super().unlink()
        
        for revenue in accrued_revenues:
            if revenue.exists():  # Check if record still exists
                revenue.update_total_accrued_line()
        
        return result