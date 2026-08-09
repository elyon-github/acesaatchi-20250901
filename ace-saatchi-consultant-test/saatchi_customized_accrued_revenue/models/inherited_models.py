# -*- coding: utf-8 -*-
"""
Inherited Model Extensions for Accrued Revenue
===============================================
Extends sale.order, account.move, account.move.line, and
account.general.ledger.report.handler to support accrued revenue functionality.
"""

from odoo import models, fields, api, _, Command
from odoo.exceptions import UserError
from odoo.tools import SQL
from dateutil.relativedelta import relativedelta
import logging

_logger = logging.getLogger(__name__)


class SaleOrder(models.Model):
    """
    Sale Order Extension

    Adds custom fields and methods for:
    - CE status tracking
    - Approved estimates and invoiced amounts (billing & revenue)
    - Variance calculations
    - Accrual creation logic (normal, override, adjustment)
    """
    _inherit = "sale.order"

    # ========== Custom CE Fields ==========
    x_ce_status = fields.Selection(
        [
            ('for_client_signature', 'For Client Signature'),
            ('signed', 'Signed'),
            ('billable', 'Billable'),
            ('closed', 'Closed'),
            ('cancelled', 'Cancelled')
        ],
        default='for_client_signature',
        required=True,
        string='C.E. Status',
        tracking=True
    )

    related_accrued_revenue_count = fields.Integer(
        string="Related Accrued Revenue Count", compute="_compute_related_accrued_revenue_count")
    related_accrued_revenue_journal_items_count = fields.Integer(
        string="Related Accrued Revenue Journal Items Count", compute="_compute_related_accrued_revenue_journal_items_count")

    # ========== Accrual Collection Methods ==========

    @api.model
    def _normalize_ce_code_for_match(self, ce_code):
        """Normalize CE code: remove whitespace, uppercase."""
        if not ce_code:
            return ''
        import re
        return re.sub(r'\s+', '', ce_code.strip().upper())

    @api.model
    def collect_potential_accruals(self, accrual_date, reversal_date):
        """
        Collect sale orders eligible for accrual creation

        Criteria:
        - state = 'sale'
        - x_ce_status in ['signed', 'billable']
        - ALSO: x_ce_status = 'for_client_signature' IF their old CE code
          matches a CE# in the reversal opening balance model

        Args:
            accrual_date: Accrual period start date
            reversal_date: Accrual period end date

        Returns:
            tuple: (potential_sos, duplicate_sos, client_sig_so_ids)
                client_sig_so_ids: set of SO IDs that are Client Signature from reversal OB
        """
        potential = self.env['sale.order']
        duplicates = self.env['sale.order']
        client_sig_so_ids = set()

        # ── Standard: signed/billable SOs ──
        # Only include SOs whose effective date <= accrual_date.
        # Priority: old_ce_date > create_date (when old CE# exists) > date_order
        eligible_sos = self.search([
            ('state', '=', 'sale'),
            ('x_ce_status', 'in', ['signed', 'billable']),
            ('x_ce_code', '!=', False),
            '|',
            # Case 1: old CE# AND old CE date exist → use old CE date
            '&', ('x_studio_old_ce', '!=', False),
            '&', ('x_studio_old_ce_date', '!=', False),
                 ('x_studio_old_ce_date', '<=', accrual_date),
            '|',
            # Case 2: old CE# exists but no old CE date → use create_date
            '&', ('x_studio_old_ce', '!=', False),
            '&', ('x_studio_old_ce_date', '=', False),
                 ('create_date', '<=', accrual_date),
            # Case 3: no old CE# → use date_order AND create_date (ignore old CE date)
            '&', ('x_studio_old_ce', '=', False),
            '&', ('date_order', '<=', accrual_date),
                 ('create_date', '<=', accrual_date),
        ])

        for so in eligible_sos:
            potential |= so

            existing = self.env['saatchi.accrued_revenue'].search([
                ('x_related_ce_id', '=', so.id),
                ('date', '>=', accrual_date),
                ('date', '<=', reversal_date),
                ('state', 'in', ['draft', 'accrued', 'reversed'])
            ], limit=1)

            if existing:
                duplicates |= so

        # ── Special: Client Signature SOs with old CE in reversal OB ──
        client_sig_sos = self.search([
            ('state', '=', 'sale'),
            ('x_ce_status', '=', 'for_client_signature'),
            ('x_ce_code', '!=', False),
            '|',
            # Case 1: old CE# AND old CE date exist → use old CE date
            '&', ('x_studio_old_ce', '!=', False),
            '&', ('x_studio_old_ce_date', '!=', False),
                 ('x_studio_old_ce_date', '<=', accrual_date),
            '|',
            # Case 2: old CE# exists but no old CE date → use create_date
            '&', ('x_studio_old_ce', '!=', False),
            '&', ('x_studio_old_ce_date', '=', False),
                 ('create_date', '<=', accrual_date),
            # Case 3: no old CE# → use date_order AND create_date (ignore old CE date)
            '&', ('x_studio_old_ce', '=', False),
            '&', ('date_order', '<=', accrual_date),
                 ('create_date', '<=', accrual_date),
        ])

        if client_sig_sos:
            # ── Tier 1: Reversal OB match — ONLY for the cutoff month ──
            # This tier fires exactly once: when generating accruals for the
            # month immediately after the configured cutoff date.
            # e.g. cutoff = 2025-12-31 → Tier 1 only fires for Jan 2026
            #      (accrual_date = 2026-01-31, prev_month_end = 2025-12-31)
            # For Feb 2026 and beyond, Tier 2 (continuation chain) takes over.
            first_of_accrual_month = accrual_date.replace(day=1)
            prev_month_end = first_of_accrual_month - relativedelta(days=1)

            accrual_config = self.env['saatchi.accrual_config'].sudo().search([
                ('company_id', '=', self.env.company.id),
            ], limit=1)
            cutoff_date = accrual_config.opening_balance_cutoff_date if accrual_config else False

            is_cutoff_month = cutoff_date and prev_month_end == cutoff_date

            if is_cutoff_month:
                reversal_ob_records = self.env[
                    'saatchi.accrued_revenue_reversal_opening_balance'
                ].sudo().search([
                    ('company_id', '=', self.env.company.id),
                    ('balance_date', '=', prev_month_end),
                ])
                reversal_ob_ce_codes = set()
                for rec in reversal_ob_records:
                    if rec.ce_code:
                        reversal_ob_ce_codes.add(
                            self._normalize_ce_code_for_match(rec.ce_code)
                        )

                # Fetch opening balance data to check balance > 0
                ob_balances = self.env[
                    'saatchi.accrued_revenue_opening_balance'
                ].sudo().get_opening_balances_for_month(
                    balance_date=prev_month_end,
                    company_id=self.env.company.id,
                )

                for so in client_sig_sos:
                    old_ce = getattr(so, 'x_studio_old_ce', '')
                    if old_ce and self._normalize_ce_code_for_match(old_ce) in reversal_ob_ce_codes:
                        # Only include if opening balance for this CE > 0
                        norm_old_ce = self._normalize_ce_code_for_match(old_ce)
                        ob_balance = ob_balances.get(norm_old_ce, 0.0)
                        if ob_balance <= 0:
                            _logger.debug(
                                'Skipping Client Signature SO %s (CE %s): '
                                'opening balance %.2f <= 0',
                                so.name, old_ce, ob_balance,
                            )
                            continue

                        potential |= so
                        client_sig_so_ids.add(so.id)

                        existing = self.env['saatchi.accrued_revenue'].search([
                            ('x_related_ce_id', '=', so.id),
                            ('date', '>=', accrual_date),
                            ('date', '<=', reversal_date),
                            ('state', 'in', ['draft', 'accrued', 'reversed'])
                        ], limit=1)

                        if existing:
                            duplicates |= so

        # ── Tier 2: Continuation chain — Client Signature SOs with previous month accrual ──
        # Picks up for_client_signature SOs that already have an accrual record
        # from the previous month. This creates a chain: Tier 1 seeds the first
        # month (cutoff), then Tier 2 carries them forward month-by-month.
        # Chain breaks when previous month's ending balance drops to 0 or below.
        if client_sig_sos:
            # Reuse first_of_accrual_month/prev_month_end from Tier 1 if available,
            # otherwise compute them (in case Tier 1 was skipped)
            if 'first_of_accrual_month' not in dir():
                first_of_accrual_month = accrual_date.replace(day=1)
                prev_month_end = first_of_accrual_month - relativedelta(days=1)
            prev_month_start = (first_of_accrual_month -
                                relativedelta(months=1))

            # Reuse accrual_config from Tier 1 if available
            if 'accrual_config' not in dir():
                accrual_config = self.env['saatchi.accrual_config'].sudo().search([
                    ('company_id', '=', self.env.company.id),
                ], limit=1)
                cutoff_date = accrual_config.opening_balance_cutoff_date if accrual_config else False

            # Pre-compute previous month ending balances from DB for
            # Client Signature CEs (accrued revenue account, cumulative
            # debit - credit up to prev_month_end).
            accrued_account_ids = []
            if accrual_config and accrual_config.accrued_revenue_account_id:
                accrued_account_ids = accrual_config.accrued_revenue_account_id.ids

            # Build a map of normalized CE code -> cumulative balance
            prev_month_ce_balances = {}
            if accrued_account_ids:
                prev_lines = self.env['account.move.line'].sudo().search([
                    ('account_id', 'in', accrued_account_ids),
                    ('date', '<=', prev_month_end),
                    ('parent_state', '=', 'posted'),
                ])
                for line in prev_lines:
                    ce_code = line.x_ce_code or ''
                    if ce_code:
                        norm = self._normalize_ce_code_for_match(ce_code)
                        prev_month_ce_balances[norm] = prev_month_ce_balances.get(
                            norm, 0.0) + (line.debit or 0) - (line.credit or 0)

            # Also check opening balance as fallback for CEs with no DB history
            cutoff_date = accrual_config.opening_balance_cutoff_date if accrual_config else False
            ob_fallback_balances = {}
            if cutoff_date:
                ob_fallback_balances = self.env[
                    'saatchi.accrued_revenue_opening_balance'
                ].sudo().get_opening_balances_for_month(
                    balance_date=cutoff_date,
                    company_id=self.env.company.id,
                )

            for so in client_sig_sos:
                if so in potential:
                    # Already included (e.g., from reversal OB tier above)
                    continue

                # Check if this SO had an accrual in the previous month
                prev_accrual = self.env['saatchi.accrued_revenue'].search([
                    ('x_related_ce_id', '=', so.id),
                    ('date', '>=', prev_month_start),
                    ('date', '<=', prev_month_end),
                    ('state', 'in', ['draft', 'accrued', 'reversed'])
                ], limit=1)

                if prev_accrual:
                    # Check previous month ending balance > 0
                    old_ce = getattr(so, 'x_studio_old_ce', '') or ''
                    ce_code_for_balance = so.x_ce_code or ''
                    norm_old_ce = self._normalize_ce_code_for_match(
                        old_ce) if old_ce else ''
                    norm_ce = self._normalize_ce_code_for_match(
                        ce_code_for_balance) if ce_code_for_balance else ''

                    # Try DB balance first (old CE or regular CE)
                    balance = prev_month_ce_balances.get(norm_old_ce, None)
                    if balance is None and norm_ce:
                        balance = prev_month_ce_balances.get(norm_ce, None)

                    # Fallback to opening balance if no DB history
                    if balance is None:
                        balance = ob_fallback_balances.get(norm_old_ce, None)
                        if balance is None and norm_ce:
                            balance = ob_fallback_balances.get(norm_ce, None)

                    if balance is None or balance <= 0:
                        _logger.debug(
                            'Skipping continuation Client Signature SO %s (CE %s): '
                            'prev month balance %.2f <= 0',
                            so.name, old_ce or ce_code_for_balance,
                            balance if balance is not None else 0.0,
                        )
                        continue

                    potential |= so
                    client_sig_so_ids.add(so.id)

                    # Check for duplicates in current period
                    existing = self.env['saatchi.accrued_revenue'].search([
                        ('x_related_ce_id', '=', so.id),
                        ('date', '>=', accrual_date),
                        ('date', '<=', reversal_date),
                        ('state', 'in', ['draft', 'accrued', 'reversed'])
                    ], limit=1)

                    if existing:
                        duplicates |= so

        return potential, duplicates, client_sig_so_ids

    # ========== Wizard Action Methods ==========

    def action_open_wizard_create_accrued_revenue(self, records, special_case=False):
        """
        Open wizard to create accrued revenue for selected sale orders.

        Amounts are converted from SO currency to company currency (PHP) at population time.
        The Select toggle defaults to False for SOs with existing accruals, True otherwise.

        Args:
            records: Sale order recordset
            special_case: If True, enables scenario selection (Manual Accrue/Adjustment)

        Returns:
            dict: Action to open wizard
        """
        accrual_date = self.env['saatchi.accrued_revenue']._default_accrual_date(
        )
        reversal_date = self.env['saatchi.accrued_revenue']._default_reversal_date(
        )

        wizard = self.env['saatchi.accrued_revenue.wizard'].create({
            'accrual_date': accrual_date,
            'reversal_date': reversal_date,
            'special_case_mode': special_case,
        })

        company_currency = self.env.company.currency_id
        for record in records:
            if (record.x_ce_variance_revenue > 0 or special_case) and record.state == 'sale':
                existing = self.env['saatchi.accrued_revenue'].search([
                    ('x_related_ce_id', '=', record.id),
                    ('date', '>=', accrual_date),
                    ('date', '<=', reversal_date),
                    ('state', 'in', ['draft', 'accrued', 'reversed'])
                ])

                # Use _calculate_accrual_amount for consistency with other paths
                amount_total = record._calculate_accrual_amount(
                    accrual_date=accrual_date)

                # Convert to company currency (PHP) if SO is in foreign currency
                if record.currency_id != company_currency and amount_total:
                    date_order = record.date_order.date() if record.date_order else accrual_date
                    amount_total = record.currency_id._convert(
                        amount_total, company_currency, record.company_id, date_order)

                self.env['saatchi.accrued_revenue.wizard.line'].create({
                    'wizard_id': wizard.id,
                    'sale_order_id': record.id,
                    'amount_total': amount_total,
                    'has_existing_accrual': bool(existing),
                    'existing_accrual_ids': [(6, 0, existing.ids)],
                    'create_accrual': not bool(existing),
                })

        return {
            'type': 'ir.actions.act_window',
            'name': _('Create Accrued Revenue'),
            'res_model': 'saatchi.accrued_revenue.wizard',
            'view_mode': 'form',
            'res_id': wizard.id,
            'views': [(False, 'form')],
            'target': 'new',
            'context': self.env.context,
        }

    # ========== Accrual Creation Methods ==========

    def _calculate_accrual_amount(self, accrual_date=None):
        """
        Calculate total accrual amount for this sale order (in SO currency).

        Only processes Agency Charges category lines that have:
        - Delivered but not invoiced quantity (accrued_qty > 0)

        If accrual_date is provided, only counts invoices with
        invoice_date <= accrual_date (posted, non-cancelled).
        This ensures accurate historical accrual
        (e.g., Jan accrual won't be affected by Feb invoices).

        Note: Returns amount in the SO's own currency. Callers are responsible
        for converting to company currency (PHP) if needed.

        Args:
            accrual_date: Optional date to filter invoices by.
                         If None, uses current qty_invoiced (backwards compatible).

        Returns:
            float: Total accrual amount in SO currency
        """
        self.ensure_one()
        amount_total = 0

        for line in self.order_line:
            if line.display_type:
                continue

            if not self._is_agency_charges_category(line.product_template_id):
                continue

            if accrual_date:
                # Manually compute qty_invoiced considering only invoices
                # with invoice_date <= accrual_date (mirrors Odoo's _compute_qty_invoiced)
                qty_invoiced = 0.0
                for inv_line in line.invoice_lines:
                    move = inv_line.move_id
                    if move.state == 'cancel' and move.payment_state != 'invoicing_legacy':
                        continue
                    if not move.invoice_date or move.invoice_date > accrual_date:
                        continue
                    if move.move_type == 'out_invoice':
                        qty_invoiced += inv_line.product_uom_id._compute_quantity(
                            inv_line.quantity, line.product_uom, round=False)
                    elif move.move_type == 'out_refund':
                        qty_invoiced -= inv_line.product_uom_id._compute_quantity(
                            inv_line.quantity, line.product_uom, round=False)
                accrued_qty = line.product_uom_qty - qty_invoiced
            else:
                accrued_qty = line.product_uom_qty - line.qty_invoiced

            if accrued_qty <= 0:
                continue

            accrued_amount = accrued_qty * line.price_unit
            amount_total += accrued_amount

        return amount_total

    def action_create_custom_accrued_revenue(self, is_override=False, accrual_date=False, reversal_date=False, is_adjustment=False, is_system_generated=True, keep_foreign_currency=False, adjustment_amount=0):
        """
        Create accrued revenue entry for this sale order

        Process:
        1. Validate SO state (skip if override=True)
        2. Create accrued revenue record
        3. Create lines based on type:
           - Normal/Override: Individual SO lines + Total Accrued (with auto-reversal)
           - Adjustment: Digital Income line + Total Accrued (NO auto-reversal)

        Args:
            is_override: If True, skip validation (Scenario 1: Manual Accrue)
            accrual_date: Custom accrual date
            reversal_date: Custom reversal date
            is_adjustment: If True, create adjustment entry
            is_system_generated: If True, marks accrual as system generated
            keep_foreign_currency: If False (default), convert foreign currency to company currency
            adjustment_amount: Signed amount for adjustment entries (negative=reduce, positive=increase)

        Returns:
            int: Accrual record ID if successful, False otherwise
        """
        self.ensure_one()

        # Validate sale order state (skip if override or adjustment)
        if not is_override and not is_adjustment:
            if self.state != 'sale' or self.x_ce_status not in ['signed', 'billable', 'for_client_signature']:
                _logger.warning(
                    f"SO {self.name} does not meet accrual criteria")
                return False

        if not accrual_date:
            accrual_date = self.env['saatchi.accrued_revenue']._default_accrual_date(
            )
        # if not reversal_date:
        #     reversal_date = self.env['saatchi.accrued_revenue']._default_reversal_date()

        # Determine if we should convert to company currency
        company_currency = self.company_id.currency_id
        so_currency = self.currency_id
        should_convert = not keep_foreign_currency and so_currency != company_currency

        # Determine conversion date: Old CE Date > Create Date (if old CE# exists) > Order Date > Accrual Date
        conversion_date = accrual_date
        if should_convert:
            old_ce_code = getattr(self, 'x_studio_old_ce', False)
            old_ce_date = getattr(self, 'x_studio_old_ce_date', False) if old_ce_code else False
            if old_ce_code and old_ce_date:
                conversion_date = old_ce_date
            elif old_ce_code and self.create_date:
                # Old CE# exists but no old CE date → use create_date
                conversion_date = self.create_date.date() if hasattr(self.create_date, 'date') else self.create_date
            elif self.date_order:
                conversion_date = self.date_order

        # Use company currency if converting, otherwise SO currency
        record_currency = company_currency if should_convert else so_currency

        # Create accrued revenue record
        accrued_revenue = self.env['saatchi.accrued_revenue'].with_context(default_company_id=self.company_id.id
                                                                           ).create({
                                                                               'x_related_ce_id': self.id,
                                                                               'currency_id': record_currency.id,
                                                                               'date': accrual_date,
                                                                               'reversal_date': reversal_date,
                                                                               'is_adjustment_entry': is_adjustment,
                                                                               'x_accrual_system_generated': is_system_generated,
                                                                               'keep_foreign_currency': keep_foreign_currency,
                                                                           })

        if is_adjustment:
            # Scenario 2/3: Create adjustment entry from wizard-provided amount
            result = self._create_adjustment_entry_lines(
                accrued_revenue, should_convert, conversion_date,
                adjustment_amount=adjustment_amount)
        else:
            # Normal or Override: Create lines from SO (with auto-reversal)
            result = self._create_normal_accrual_lines(
                accrued_revenue, should_convert, conversion_date)

        # Return the result (accrual ID or False)
        return result

    def _create_normal_accrual_lines(self, accrued_revenue, should_convert=False, conversion_date=False):
        """
        Create normal accrual lines from SO lines

        Structure:
        - For positive amounts: Cr. Revenue | Dr. Accrued Revenue
        - For negative amounts (returns): Dr. Revenue | Cr. Accrued Revenue
        - Creates automatic reversal entry

        Args:
            accrued_revenue: The accrued revenue record
            should_convert: If True, convert foreign currency amounts to company currency
            conversion_date: Date to use for currency conversion

        Returns:
            int: Accrual record ID if successful, False otherwise
        """
        total_eligible_for_accrue = 0
        lines_created = 0

        # ── Performance (2026-08-09) ────────────────────────────────────────
        # Line values are collected here and written in ONE create() after the
        # loop.  Two per-call caches avoid re-running searches that can only
        # ever return the same answer within a single call.  See the comments
        # at each use site, and the create() call at the end of this method.
        # Behaviour is unchanged: same lines, same values, same totals.
        # ────────────────────────────────────────────────────────────────────
        line_vals_list = []
        agencies_distribution = None   # None = not looked up yet; {} = none found
        income_account_cache = {}      # {template_account_id: account record}

        _logger.info(f"Starting accrual creation for SO {self.name}")

        # Get the target company from accrued_revenue
        target_company = accrued_revenue.company_id

        # Use accrual date to filter invoices for accurate qty_invoiced
        accrual_date = accrued_revenue.date

        for line in self.order_line:
            if line.display_type:
                continue

            if not self._is_agency_charges_category(line.product_template_id):
                _logger.debug(
                    f"Skipping line {line.name} - not Agency Charges category")
                continue

            # Compute qty_invoiced filtered by accrual date
            if accrual_date:
                qty_invoiced = 0.0
                for inv_line in line.invoice_lines:
                    move = inv_line.move_id
                    if move.state == 'cancel' and move.payment_state != 'invoicing_legacy':
                        continue
                    if not move.invoice_date or move.invoice_date > accrual_date:
                        continue
                    if move.move_type == 'out_invoice':
                        qty_invoiced += inv_line.product_uom_id._compute_quantity(
                            inv_line.quantity, line.product_uom, round=False)
                    elif move.move_type == 'out_refund':
                        qty_invoiced -= inv_line.product_uom_id._compute_quantity(
                            inv_line.quantity, line.product_uom, round=False)
            else:
                qty_invoiced = line.qty_invoiced

            accrued_qty = line.product_uom_qty - qty_invoiced
            if accrued_qty == 0:  # Changed from <= to == to allow negative
                _logger.debug(
                    f"Skipping line {line.name} - no accrued qty (qty: {line.product_uom_qty}, invoiced: {qty_invoiced})")
                continue

            accrued_amount = accrued_qty * line.price_unit

            # Convert to company currency if needed
            if should_convert:
                accrued_amount = self.currency_id._convert(
                    accrued_amount,
                    self.company_id.currency_id,
                    self.company_id,
                    conversion_date or accrued_revenue.date
                )
                line_currency = self.company_id.currency_id
            else:
                line_currency = line.currency_id

            # Determine analytic distribution with fallback logic
            analytic_distribution = line.analytic_distribution or {}

            # Fallback to SO level analytic distribution
            if not analytic_distribution and hasattr(self, 'analytic_distribution') and self.analytic_distribution:
                analytic_distribution = self.analytic_distribution

            # Default to "AGENCIES" analytic account for target company if still no distribution found.
            # The lookup depends only on target_company, which is constant for the
            # whole call, so it is resolved once and reused.  It previously ran one
            # search() per order line.  A copy is handed out each time so a caller
            # mutating one line's distribution cannot affect the others.
            if not analytic_distribution:
                if agencies_distribution is None:
                    agencies_account = self.env['account.analytic.account'].sudo().search([
                        ('name', '=', 'AGENCIES'),
                        ('company_id', '=', target_company.id)
                    ], limit=1)
                    if agencies_account:
                        agencies_distribution = {agencies_account.id: 100}
                        _logger.debug(
                            f"Using default AGENCIES analytic account (ID: {agencies_account.id}) for SO {self.name}")
                    else:
                        agencies_distribution = {}
                        _logger.warning(
                            f"No AGENCIES analytic account found for company {target_company.name}. "
                            f"Skipping analytic distribution for SO {self.name}")
                analytic_distribution = dict(agencies_distribution)

            # Get income account from product
            template_income_account = line.product_id.property_account_income_id or \
                line.product_id.categ_id.property_account_income_categ_id

            if not template_income_account:
                _logger.warning(
                    f"No income account found for line {line.name} in SO {self.name}")
                continue

            # Resolve the equivalent account in the target company.
            # Cached per template account: order lines routinely repeat the same
            # product (and therefore the same income account), and this used to
            # re-run the search for every one of them.
            if template_income_account.id in income_account_cache:
                income_account = income_account_cache[template_income_account.id]
            elif target_company in template_income_account.company_ids:
                # Account is already valid for the target company
                income_account = template_income_account
                income_account_cache[template_income_account.id] = income_account
            else:
                # Find the equivalent account in the target company by name.
                # NOTE: a second "fallback: try by name" search used to sit here,
                # byte-for-byte identical to this one (same domain, same limit), so
                # it could never return anything this search had not already found.
                # It was removed -- this is the only lookup that ever did anything.
                income_account = self.env['account.account'].sudo().search([
                    ('name', '=', template_income_account.name),
                    ('company_ids', 'in', target_company.id),
                    ('deprecated', '=', False)
                ], limit=1)
                income_account_cache[template_income_account.id] = income_account

            if not income_account:
                _logger.warning(
                    f"No equivalent income account found for line {line.name} in company {target_company.name}. "
                    f"Template account: {template_income_account.code} - {template_income_account.name}"
                )
                continue

            # Handle negative amounts (returns/adjustments).
            # Values are collected here and created in one batch after the loop.
            effective_ce = self.x_studio_old_ce or self.x_ce_code
            line_vals = {
                'accrued_revenue_id': accrued_revenue.id,
                'ce_line_id': line.id,
                'account_id': income_account.id,
                'label': f'{effective_ce} - {line.name}',
                'currency_id': line_currency.id,
                'analytic_distribution': analytic_distribution,
            }
            if accrued_amount < 0:
                # Negative accrual: Dr. Revenue (reverse the credit)
                line_vals.update({'debit': abs(accrued_amount), 'credit': 0.0})
            else:
                # Positive accrual: Cr. Revenue (normal)
                line_vals.update({'credit': accrued_amount, 'debit': 0.0})
            line_vals_list.append(line_vals)

            total_eligible_for_accrue += accrued_amount  # Keep the sign
            lines_created += 1
            _logger.debug(
                f"Prepared line for {line.name}, amount: {accrued_amount}")

        if lines_created == 0:
            accrued_revenue.unlink()
            _logger.warning(
                f"No eligible lines found for accrual in SO {self.name}")
            return False

        # Total Accrued line -- balanced by update_total_accrued_line().
        #
        # sequence=999 is REQUIRED here and must not be dropped.  The model is
        # ordered '_order = sequence desc', so 999 keeps Total Accrued at the top
        # of the line list.  Before batching, this line was created without a
        # sequence (defaulting to 10), but update_total_accrued_line() had already
        # created its own Total Accrued line with sequence=999 on the first
        # per-line create(); the next recompute then deleted the surplus line and
        # kept the 999 one.  Batching removes that accidental duplicate, so the
        # sequence now has to be set explicitly to preserve the same end state.
        line_vals_list.append({
            'accrued_revenue_id': accrued_revenue.id,
            'label': 'Total Accrued',
            'currency_id': accrued_revenue.currency_id.id,
            'account_id': accrued_revenue.accrual_account_id.id,
            'debit': 0.0,
            'credit': 0.0,
            'sequence': 999,
        })

        # ── Single batched create (2026-08-09) ──────────────────────────────
        # saatchi.accrued_revenue_lines.create() is an @api.model_create_multi
        # override that calls update_total_accrued_line() ONCE PER DISTINCT
        # accrued_revenue_id in the batch.  Creating the lines one at a time
        # therefore triggered a full Total-Accrued recompute *and* write for
        # every single line; one create() call triggers exactly one, with an
        # identical end state.  Do not split this back into per-line creates.
        # ────────────────────────────────────────────────────────────────────
        self.env['saatchi.accrued_revenue_lines'].create(line_vals_list)

        # Store original total (already converted if should_convert was True)
        accrued_revenue.write(
            {'ce_original_total_amount': total_eligible_for_accrue})

        _logger.info(
            f"✓ Created accrual ID {accrued_revenue.id} for SO {self.name} with {lines_created} lines, total: {total_eligible_for_accrue}")

        return accrued_revenue.id

    def _create_adjustment_entry_lines(self, accrued_revenue, should_convert=False, conversion_date=False, adjustment_amount=0):
        """
        Create adjustment entry lines (Scenario 2 & 3)

        Sign convention:
        - Negative amount (reduce accrual):
            Dr. Digital Income  |  Cr. Accrued Revenue
        - Positive amount (increase accrual):
            Dr. Accrued Revenue  |  Cr. Digital Income

        Args:
            accrued_revenue: The accrued revenue record
            should_convert: If True, convert foreign currency amounts to company currency
            conversion_date: Date to use for currency conversion
            adjustment_amount: Signed adjustment amount from wizard
                               (negative = reduce accrual, positive = increase accrual)

        Returns:
            int: Accrual record ID if successful, False otherwise
        """
        # Use wizard-provided amount; fall back to calculated amount if zero
        if not adjustment_amount:
            adjustment_amount = -(self._calculate_accrual_amount(
                accrual_date=accrued_revenue.date) or 0)
            if not adjustment_amount:
                _logger.info(
                    f"Creating adjustment entry for SO {self.name} with 0 default amount")

            # Convert fallback amount (in SO currency) to company currency if needed
            if should_convert and adjustment_amount:
                adjustment_amount = self.currency_id._convert(
                    adjustment_amount,
                    self.company_id.currency_id,
                    self.company_id,
                    conversion_date or accrued_revenue.date
                )
        # else: wizard-provided adjustment_amount is already in company currency (PHP)

        line_currency = self.company_id.currency_id if should_convert else self.currency_id
        abs_amount = abs(adjustment_amount)

        # Get Digital Income account
        digital_income_account = accrued_revenue.digital_income_account_id
        if not digital_income_account:
            accrued_revenue.unlink()
            raise UserError(
                _("Digital Income account not configured. Please set it in the accrual settings."))

        # Get analytic distribution from SO
        analytic_distribution = {}
        if hasattr(self, 'analytic_distribution') and self.analytic_distribution:
            analytic_distribution = self.analytic_distribution

        # Determine debit/credit based on sign
        if adjustment_amount < 0:
            # REDUCE accrual: Dr. Digital Income | Cr. Accrued Revenue
            income_debit = abs_amount
            income_credit = 0.0
        else:
            # INCREASE accrual: Dr. Accrued Revenue | Cr. Digital Income
            income_debit = 0.0
            income_credit = abs_amount

        # Line 1: Digital Income line
        self.env['saatchi.accrued_revenue_lines'].create({
            'accrued_revenue_id': accrued_revenue.id,
            'account_id': digital_income_account.id,
            'label': 'Digital Income - Adjustment',
            'debit': income_debit,
            'credit': income_credit,
            'currency_id': line_currency.id,
            'analytic_distribution': analytic_distribution,
        })

        # Line 2: Total Accrued (auto-balanced by update_total_accrued_line via create trigger)
        self.env['saatchi.accrued_revenue_lines'].create({
            'accrued_revenue_id': accrued_revenue.id,
            'label': 'Total Accrued',
            'currency_id': line_currency.id,
            'account_id': accrued_revenue.accrual_account_id.id,
            'debit': 0.0,
            'credit': 0.0,
        })

        accrued_revenue.write(
            {'ce_original_total_amount': abs_amount})

        _logger.info(
            f"✓ Created adjustment entry for SO {self.name}, amount: {adjustment_amount} "
            f"({'reduce' if adjustment_amount < 0 else 'increase'} accrual)")

        return accrued_revenue.id

    def _is_agency_charges_category(self, product):
        """
        Check if product belongs to Agency Charges category or its children

        Args:
            product: product.template recordset

        Returns:
            bool: True if product is in Agency Charges category
        """
        if not product or not product.categ_id:
            return False

        current_categ = product.categ_id
        while current_categ:
            if current_categ.name.lower() == 'agency charges':
                return True
            current_categ = current_categ.parent_id

        return False

    def action_view_accrued_revenues(self):
        self.ensure_one()
        return {
            'name': 'Accrued Revenues',
            'type': 'ir.actions.act_window',
            'res_model': 'saatchi.accrued_revenue',
            'view_mode': 'list,form',
            'domain': [('x_related_ce_id', '=', self.id)],
            'context': {'default_x_related_ce_id': self.id},
        }

    def action_view_accrued_revenues_journal_items(self):
        self.ensure_one()
        effective_ce = self.x_studio_old_ce or self.x_ce_code
        list_view_id = self.env.ref(
            'saatchi_customized_accrued_revenue.view_accrued_revenue_journal_items_list').id
        search_view_id = self.env.ref(
            'saatchi_customized_accrued_revenue.view_account_move_line_accrued_revenue_filter').id
        return {
            'name': 'Accrued Revenues Journal Items',
            'type': 'ir.actions.act_window',
            'res_model': 'account.move.line',
            'view_mode': 'list,form',
            'views': [(list_view_id, 'list')],
            'search_view_id': search_view_id,  # Add comma here!
            'domain': [('x_ce_code', '=', effective_ce)],
            'context': {'journal_type': 'general', 'search_default_posted': 1},
        }

    def _compute_related_accrued_revenue_count(self):
        for record in self:
            record.related_accrued_revenue_count = self.env['saatchi.accrued_revenue'].search_count([
                ('x_related_ce_id', '=', record.id)
            ])

    def _compute_related_accrued_revenue_journal_items_count(self):
        for record in self:
            effective_ce = record.x_studio_old_ce or record.x_ce_code
            record.related_accrued_revenue_journal_items_count = self.env['account.move.line'].search_count([
                ('x_ce_code', '=', effective_ce), ('x_type_of_entry', '!=', False)
            ])


class AccountMove(models.Model):
    """
    Account Move Extension

    Adds fields to track accrued revenue-related journal entries.
    """
    _inherit = "account.move"

    x_related_custom_accrued_record = fields.Many2one(
        'saatchi.accrued_revenue',
        store=True,
        readonly=True,
        string="Related Accrued Revenue",
        help="Link to the accrued revenue record that generated this move"
    )

    x_remarks = fields.Char(
        string="Remarks",
        help="Additional remarks for this journal entry"
    )

    x_accrual_system_generated = fields.Boolean(
        string="System Generated",
        help="Indicates if this accrual was generated automatically by the system"
    )
    x_is_reversal = fields.Boolean(string="Is Reversal Entry?")

    x_is_accrued_entry = fields.Boolean(string="Is Accrued Entry?")

    x_type_of_entry = fields.Selection(
        selection=[
            ('reversal_system', 'Reversal Entry - System'),
            ('reversal_manual', 'Reversal Entry - Manual'),
            ('accrued_system', 'Accrued Entry - System'),
            ('accrued_manual', 'Accrued Entry - Manual'),
            ('adjustment_system', 'Adjustment Entry - System'),
            ('adjustment_manual', 'Adjustment Entry - Manual'),
        ],
        string="Entry Type",
        compute="_compute_entry_type",
        store=True,
        readonly=True
    )

    @api.depends('x_related_custom_accrued_record',
                 'x_related_custom_accrued_record.is_adjustment_entry',
                 'x_accrual_system_generated',
                 'ref')
    def _compute_entry_type(self):
        """Determine entry type based on accrued revenue record and generation method"""
        for move in self:
            if not move.x_related_custom_accrued_record:
                move.x_type_of_entry = False
                continue

            # Determine if system or manual
            suffix = 'system' if move.x_accrual_system_generated else 'manual'

            # Determine entry type
            if move.x_related_custom_accrued_record.is_adjustment_entry:
                move.x_type_of_entry = f'adjustment_{suffix}'
            elif move.ref and 'Reversal' in move.ref:
                move.x_type_of_entry = f'reversal_{suffix}'
            else:
                move.x_type_of_entry = f'accrued_{suffix}'

    x_sales_order = fields.Many2one(
        'sale.order',
        string="Sales Order",
        compute="_compute_sales_order",
        store=True
    )

    @api.depends('invoice_line_ids.sale_line_ids',
                 'invoice_line_ids.sale_line_ids.order_id',
                 'x_related_custom_accrued_record',
                 'x_related_custom_accrued_record.x_related_ce_id')
    def _compute_sales_order(self):
        """Compute linked sale order"""
        for move in self:
            so = False

            # Priority 1: Get from accrued record
            if move.x_related_custom_accrued_record and move.x_related_custom_accrued_record.x_related_ce_id:
                so = move.x_related_custom_accrued_record.x_related_ce_id

            # Priority 2: Get from invoice lines
            elif move.invoice_line_ids:
                # Try to get from sale lines first
                sale_lines = move.invoice_line_ids.mapped('sale_line_ids')
                if sale_lines:
                    so = sale_lines[0].order_id
                else:
                    # Fallback to purchase lines
                    po_sale_lines = move.invoice_line_ids.mapped(
                        'purchase_line_id')

                    if po_sale_lines:
                        # Safely check if x_studio_sales_order exists and has a value
                        first_po_line = po_sale_lines[0]
                        if first_po_line.order_id and first_po_line.order_id.x_studio_sales_order:
                            so = first_po_line.order_id.x_studio_sales_order
            move.x_sales_order = so


class AccountMoveLine(models.Model):
    _inherit = "account.move.line"

    x_ce_code = fields.Char(
        string="CE Code",
        compute="_compute_ce_fields",
        store=True,
        help="Contract Estimate code from sale order",
        readonly=False
    )

    x_sales_order = fields.Many2one(
        'sale.order',
        string="Sales Order",
        related='move_id.x_sales_order',  # Now just use related!
        store=True
    )
    x_salesperson_id = fields.Many2one(
        'res.users',
        string="Salesperson",
        store=True,
        related='x_sales_order.user_id'
    )

    x_salesteam_id = fields.Many2one(
        'crm.team',
        string="Sales Team",
        store=True,
        related='x_sales_order.team_id')

    x_client_product_ce_code = fields.Many2one(
        'product.template',
        store=True,
        string="Client Product",
    )
    x_ce_date = fields.Date(
        string="CE Date",
        compute="_compute_ce_fields",
        store=True,
        help="Contract Estimate date from sale order",
        readonly=False
    )

    x_ce_status = fields.Selection(
        [
            ('for_client_signature', 'For Client Signature'),
            ('signed', 'Signed'),
            ('billable', 'Billable'),
            ('closed', 'Closed'),
            ('cancelled', 'Cancelled')
        ],
        string='C.E. Status',
        compute="_compute_ce_fields",
        store=True
    )

    x_remarks = fields.Char(
        string="Remarks",
        help="Additional remarks for this journal entry line"
    )

    x_reference = fields.Char(
        string="Reference",
        related='move_id.ref'
    )

    x_is_reversal = fields.Boolean(
        string="Is Reversal Entry?",
        related='move_id.x_is_reversal'
    )

    x_is_accrued_entry = fields.Boolean(
        string="Is Accrued Entry?",
        related='move_id.x_is_accrued_entry'
    )

    x_is_adjustment_entry = fields.Boolean(
        string="Is Adjustment Entry?",
        related='move_id.x_is_accrued_entry'
    )

    x_type_of_entry = fields.Selection(
        selection=[
            ('reversal_system', 'Reversal Entry - System'),
            ('reversal_manual', 'Reversal Entry - Manual'),
            ('accrued_system', 'Accrued Entry - System'),
            ('accrued_manual', 'Accrued Entry - Manual'),
            ('adjustment_system', 'Adjustment Entry - System'),
            ('adjustment_manual', 'Adjustment Entry - Manual'),
        ],
        string="Entry Type",
        related='move_id.x_type_of_entry',
        store=True,
        readonly=True
    )

    @api.depends('x_sales_order',
                 'x_sales_order.x_ce_code',
                 'x_sales_order.date_order',
                 'x_sales_order.x_ce_status')
    def _compute_ce_fields(self):
        """Compute CE-related fields from linked sale order. Prioritizes old CE code and date over new."""
        for line in self:
            so = line.x_sales_order
            if so:
                line.x_ce_code = so.x_studio_old_ce or so.x_ce_code
                # Use old CE date if it exists; leave blank if old CE# exists but date is missing
                old_ce_code = getattr(so, 'x_studio_old_ce', False)
                old_ce_date = getattr(so, 'x_studio_old_ce_date', False) if old_ce_code else False
                if old_ce_code and old_ce_date:
                    line.x_ce_date = old_ce_date
                elif old_ce_code:
                    # Old CE# exists but no old CE date → leave blank
                    line.x_ce_date = False
                else:
                    line.x_ce_date = so.date_order
                line.x_ce_status = so.x_ce_status
                line.x_client_product_ce_code = so.x_client_product_ce_code.x_product_id.id if so.x_client_product_ce_code and so.x_client_product_ce_code.x_product_id else False
            else:
                line.x_ce_code = False
                line.x_ce_date = False
                line.x_ce_status = False
                line.x_client_product_ce_code = False


class GeneralLedgerCustomHandler(models.AbstractModel):
    """
    General Ledger Report Handler Extension

    Overrides the query to include currency_name in the General Ledger report.
    """
    _inherit = 'account.general.ledger.report.handler'

    def _get_query_amls(self, report, options, expanded_account_ids, offset=0, limit=None):
        """Override to add currency_name field to General Ledger query"""
        additional_domain = [('account_id', 'in', expanded_account_ids)
                             ] if expanded_account_ids is not None else None
        queries = []
        journal_name = self.env['account.journal']._field_to_sql(
            'journal', 'name')

        for column_group_key, group_options in report._split_options_per_column_group(options).items():
            query = report._get_report_query(
                group_options, domain=additional_domain, date_scope='strict_range')
            account_alias = query.left_join(lhs_alias='account_move_line', lhs_column='account_id',
                                            rhs_table='account_account', rhs_column='id', link='account_id')
            account_code = self.env['account.account']._field_to_sql(
                account_alias, 'code', query)
            account_name = self.env['account.account']._field_to_sql(
                account_alias, 'name')
            account_type = self.env['account.account']._field_to_sql(
                account_alias, 'account_type')

            query = SQL(
                '''
                SELECT
                    account_move_line.id,
                    account_move_line.date,
                    MIN(account_move_line.date_maturity)    AS date_maturity,
                    MIN(account_move_line.name)             AS name,
                    MIN(account_move_line.ref)              AS ref,
                    MIN(account_move_line.company_id)       AS company_id,
                    MIN(account_move_line.account_id)       AS account_id,
                    MIN(account_move_line.payment_id)       AS payment_id,
                    MIN(account_move_line.partner_id)       AS partner_id,
                    MIN(account_move_line.currency_id)      AS currency_id,
                    MIN(currency.name)                      AS currency_name,
                    SUM(account_move_line.amount_currency)  AS amount_currency,
                    MIN(COALESCE(account_move_line.invoice_date, account_move_line.date)) AS invoice_date,
                    account_move_line.date                  AS date,
                    SUM(%(debit_select)s)                   AS debit,
                    SUM(%(credit_select)s)                  AS credit,
                    SUM(%(balance_select)s)                 AS balance,
                    MIN(move.name)                          AS move_name,
                    MIN(company.currency_id)                AS company_currency_id,
                    MIN(partner.name)                       AS partner_name,
                    MIN(move.move_type)                     AS move_type,
                    MIN(%(account_code)s)                   AS account_code,
                    MIN(%(account_name)s)                   AS account_name,
                    MIN(%(account_type)s)                   AS account_type,
                    MIN(journal.code)                       AS journal_code,
                    MIN(%(journal_name)s)                   AS journal_name,
                    MIN(full_rec.id)                        AS full_rec_name,
                    %(column_group_key)s                    AS column_group_key
                FROM %(table_references)s
                JOIN account_move move                      ON move.id = account_move_line.move_id
                %(currency_table_join)s
                LEFT JOIN res_company company               ON company.id = account_move_line.company_id
                LEFT JOIN res_partner partner               ON partner.id = account_move_line.partner_id
                LEFT JOIN res_currency currency             ON currency.id = account_move_line.currency_id
                LEFT JOIN account_journal journal           ON journal.id = account_move_line.journal_id
                LEFT JOIN account_full_reconcile full_rec   ON full_rec.id = account_move_line.full_reconcile_id
                WHERE %(search_condition)s
                GROUP BY account_move_line.id, account_move_line.date
                ORDER BY account_move_line.date, move_name, account_move_line.id
                ''',
                account_code=account_code,
                account_name=account_name,
                account_type=account_type,
                journal_name=journal_name,
                column_group_key=column_group_key,
                table_references=query.from_clause,
                currency_table_join=report._currency_table_aml_join(
                    group_options),
                debit_select=report._currency_table_apply_rate(
                    SQL("account_move_line.debit")),
                credit_select=report._currency_table_apply_rate(
                    SQL("account_move_line.credit")),
                balance_select=report._currency_table_apply_rate(
                    SQL("account_move_line.balance")),
                search_condition=query.where_clause,
            )
            queries.append(query)

        full_query = SQL(" UNION ALL ").join(SQL("(%s)", query)
                                             for query in queries)

        if offset:
            full_query = SQL('%s OFFSET %s ', full_query, offset)
        if limit:
            full_query = SQL('%s LIMIT %s ', full_query, limit)

        return full_query


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    accrued_revenue_account_id = fields.Many2one(
        'account.account',
        string='Accrued Revenue Account',
        config_parameter='account.accrued_revenue_account_id',
        domain=[('deprecated', '=', False)],
        help='Default account for accrued revenues'
    )

    accrued_journal_id = fields.Many2one(
        'account.journal',
        string='Accrued Revenue Journal',
        config_parameter='account.accrued_journal_id',
        domain=[('type', '=', 'general')],
        help='Default journal for accrued revenue entries'
    )

    accrued_default_adjustment_account_id = fields.Many2one(
        'account.account',
        string='Default Accrual Adjustment Account',
        config_parameter='account.accrued_default_adjustment_account_id',
        domain=[('deprecated', '=', False)],
        help='Default account for accrual adjustments, Account used for adjustment entries (Dr. side)'
    )
