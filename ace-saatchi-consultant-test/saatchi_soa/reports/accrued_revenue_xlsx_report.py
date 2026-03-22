from odoo import models
import datetime
import re
from xlsxwriter.workbook import Workbook
from odoo.exceptions import ValidationError, UserError
from dateutil.relativedelta import relativedelta
import logging
from collections import defaultdict

_logger = logging.getLogger(__name__)


class AccruedRevenueXLSX(models.AbstractModel):
    _name = 'report.accrued_revenue_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    _description = 'Accrued Revenue XLSX Report'
    
    def _get_accrued_revenue_account_id(self):
        """Get accrued revenue account IDs with fallback for multiple companies"""
        # Use user's allowed companies as target companies
        target_companies = self.env.companies
        
        try:
            account_id = int(self.env['ir.config_parameter'].sudo().get_param(
                'account.accrued_revenue_account_id',
                default='0'
            ) or 0)
            
            if account_id:
                template_account = self.env['account.account'].sudo().browse(account_id)
                if template_account.exists() and not template_account.deprecated:
                    # Check if account is accessible by user's allowed companies
                    if any(company in template_account.company_ids for company in target_companies):
                        # Account is accessible and valid for at least one target company
                        return template_account.ids
                    
                    # Find the equivalent accounts in target companies
                    equivalent_accounts = self.env['account.account'].sudo().search([
                        ('name', '=', template_account.name),
                        ('company_ids', 'in', target_companies.ids),
                        ('deprecated', '=', False)
                    ])
                    
                    if not equivalent_accounts:
                        # Fallback: try by name
                        equivalent_accounts = self.env['account.account'].sudo().search([
                            ('name', '=', template_account.name),
                            ('company_ids', 'in', target_companies.ids),
                            ('deprecated', '=', False)
                        ])
                    
                    if equivalent_accounts:
                        return equivalent_accounts.ids
            
            # Fallback: Find miscellaneous income accounts
            misc_accounts = self.env['account.account'].sudo().search([
                ('account_type', '=', 'income_other'),
                ('deprecated', '=', False),
                ('company_ids', 'in', target_companies.ids)
            ])
            
            return misc_accounts.ids if misc_accounts else []
            
        except (ValueError, TypeError):
            return []

    def _define_formats(self, workbook):
        """Define and return format objects."""
        base_font = {'font_name': 'Calibri', 'font_size': 10}

        # Title formats
        title_format = workbook.add_format({
            **base_font,
            'bold': True,
            'font_size': 11
        })

        # Section header format with borders
        section_header_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 2
        })

        # Section header format without borders (for empty cells)
        section_header_no_border = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Column header format with thick borders
        column_header_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 2
        })

        # Normal cell format with borders
        normal_format = workbook.add_format({
            **base_font,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })

        # Normal format with text wrapping (for CLIENT column)
        normal_wrap_format = workbook.add_format({
            **base_font,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })

        # Centered cell format with borders
        centered_format = workbook.add_format({
            **base_font,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Date cell format with borders
        date_format = workbook.add_format({
            **base_font,
            'num_format': 'mm/dd/yyyy',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Currency format with borders (with dash for zero)
        currency_format = workbook.add_format({
            **base_font,
            'num_format': '#,##0.00;-#,##0.00;"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

        # Currency format with parentheses for negatives with borders (with dash for zero)
        currency_negative_format = workbook.add_format({
            **base_font,
            'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

        # Red text format for OB-only rows (CE# column)
        centered_red_format = workbook.add_format({
            **base_font,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_color': 'red'
        })

        # Blue text format for new CE#s (not in reversal opening balances)
        centered_blue_format = workbook.add_format({
            **base_font,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_color': 'blue'
        })

        return {
            'title': title_format,
            'section_header': section_header_format,
            'section_header_no_border': section_header_no_border,
            'column_header': column_header_format,
            'normal': normal_format,
            'normal_wrap': normal_wrap_format,
            'centered': centered_format,
            'date': date_format,
            'currency': currency_format,
            'currency_negative': currency_negative_format,
            'centered_red': centered_red_format,
            'centered_blue': centered_blue_format
        }

    def _determine_accrual_months(self, lines, start_date, end_date):
        """
        Determine accrual months based on the wizard's date range.
        Returns all months between start_date and end_date (inclusive).
        """
        accrual_months = []

        # Start from the first day of start_date month
        current_month = start_date.replace(day=1)

        # End at the first day of end_date month
        end_month = end_date.replace(day=1)

        # Generate all months in the range
        while current_month <= end_month:
            accrual_months.append(current_month)
            current_month = current_month + relativedelta(months=1)

        return accrual_months

    def _group_lines_by_ce(self, lines, accrual_month):
        """Group account.move.line records by partner and CE code for a specific month"""
        grouped = defaultdict(lambda: defaultdict(lambda: {
            'ce_date': None,
            'description': '',
            'year': None,
            'ce_status': '',
            'so_reference': '',
            'lines': []
        }))

        # Filter lines for this specific accrual month
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        # Include lines that are:
        # 1. Reversals dated on the 1st of accrual_month
        # 2. Accruals/Adjustments dated within accrual_month
        month_lines = []
        for line in lines:
            if not line.date:
                continue

            # Reversal entries dated in accrual month (typically 1st)
            if line.x_type_of_entry in ['reversal_system', 'reversal_manual']:
                if line.date.month == accrual_month.month and line.date.year == accrual_month.year:
                    month_lines.append(line)
            # Accrual/Adjustment entries dated within accrual month
            elif line.x_type_of_entry in ['accrued_system', 'accrued_manual', 'adjustment_system', 'adjustment_manual']:
                if accrual_month <= line.date <= accrual_month_end:
                    month_lines.append(line)

        for line in month_lines:
            partner_name = line.partner_id.name.upper(
            ) if line.partner_id and line.partner_id.name else 'UNKNOWN'
            ce_code = line.x_ce_code.upper() if line.x_ce_code else 'NO_CE'

            # Store line for processing
            grouped[partner_name][ce_code]['lines'].append(line)

            # Capture CE-level fields (use first non-empty value found)
            if not grouped[partner_name][ce_code]['ce_date']:
                # Prioritize old CE date from accrued revenue record, fallback to x_ce_date
                if line.move_id and line.move_id.x_related_custom_accrued_record:
                    accrued_record = line.move_id.x_related_custom_accrued_record
                    if accrued_record.old_ce_date:
                        grouped[partner_name][ce_code]['ce_date'] = accrued_record.old_ce_date
                    elif line.x_ce_date:
                        grouped[partner_name][ce_code]['ce_date'] = line.x_ce_date
                elif line.x_ce_date:
                    grouped[partner_name][ce_code]['ce_date'] = line.x_ce_date
            
            if grouped[partner_name][ce_code]['ce_date']:
                grouped[partner_name][ce_code]['year'] = grouped[partner_name][ce_code]['ce_date'].year

            if line.move_id and line.move_id.x_related_custom_accrued_record and not grouped[partner_name][ce_code]['description']:
                desc = line.move_id.x_related_custom_accrued_record.ce_job_description or ''
                grouped[partner_name][ce_code]['description'] = desc.upper()

            if line.move_id and line.move_id.x_related_custom_accrued_record and not grouped[partner_name][ce_code]['ce_status']:
                accrued_record = line.move_id.x_related_custom_accrued_record
                if accrued_record.ce_status:
                    # Get CE Status from the accrued revenue record
                    selection_dict = dict(accrued_record._fields['ce_status'].selection)
                    ce_status = selection_dict.get(accrued_record.ce_status, '')
                    grouped[partner_name][ce_code]['ce_status'] = ce_status.upper()

            if line.move_id and line.move_id.x_related_custom_accrued_record and not grouped[partner_name][ce_code]['so_reference']:
                accrued_record = line.move_id.x_related_custom_accrued_record
                so_ref = ''
                if accrued_record.x_related_ce_id:
                    so_ref = accrued_record.x_related_ce_id.name or ''
                grouped[partner_name][ce_code]['so_reference'] = so_ref.upper() if so_ref else ''

        return grouped

    def _normalize_ce_code(self, ce_code):
        """
        Normalize a CE code for consistent matching.
        Removes all whitespace and converts to uppercase.

        Args:
            ce_code (str): Raw CE code string

        Returns:
            str: Normalized CE code (e.g., 'CE 001' -> 'CE001')
        """
        if not ce_code:
            return ''
        return re.sub(r'\s+', '', ce_code.strip().upper())

    def _normalize_ce_status(self, status):
        """
        Normalize CE Status values for consistent display.

        Mapping:
            BILLABLE (APPROVED) -> BILLABLE
            #NA / FOR ENCODING -> CLIENT SIGNATURE
            SIGNED (RELEASED)  -> SIGNED
        """
        if not status:
            return ''
        s = status.strip().upper()
        if s in ('BILLABLE (APPROVED)',):
            return 'BILLABLE'
        if s in ('#NA', 'FOR ENCODING'):
            return 'FOR CLIENT SIGNATURE'
        if s in ('SIGNED (RELEASED)',):
            return 'SIGNED'
        return s

    def _get_opening_balance_cutoff_date(self):
        """
        Get the opening balance cutoff date from the accrual configuration
        for the current company.

        Returns:
            date or False: The cutoff month-end date, or False if not configured
        """
        try:
            config = self.env['saatchi.accrual_config'].sudo().search([
                ('company_id', '=', self.env.company.id)
            ], limit=1)
            if config and config.opening_balance_cutoff_date:
                return config.opening_balance_cutoff_date
        except Exception as e:
            _logger.warning('Could not retrieve opening balance cutoff date: %s', str(e))
        return False

    def _is_opening_balance_month(self, accrual_month):
        """
        Check if the given accrual month is the opening balance month.
        This is true when the previous month end equals the configured cutoff date.
        """
        cutoff_date = self._get_opening_balance_cutoff_date()
        if not cutoff_date:
            return False
        prev_month_end = accrual_month - relativedelta(days=1)
        return prev_month_end == cutoff_date

    def _find_sale_order_by_ce_code(self, ce_code):
        """
        Find a sale.order matching the given CE code.
        Searches both x_ce_code and x_studio_old_ce (Studio field) on sale.order.
        Returns the first matching sale.order or empty recordset.
        """
        if not ce_code:
            return self.env['sale.order']

        SaleOrder = self.env['sale.order'].sudo()
        company_domain = [('company_id', 'in', self.env.companies.ids)]

        # 1. Exact match on x_ce_code
        so = SaleOrder.search([('x_ce_code', '=', ce_code)] + company_domain, limit=1)
        if so:
            return so

        # 2. Exact match on x_studio_old_ce (Studio field)
        try:
            so = SaleOrder.search([('x_studio_old_ce', '=', ce_code)] + company_domain, limit=1)
            if so:
                return so
        except Exception:
            pass

        # 3. Fuzzy match on x_ce_code (whitespace/case variations)
        so = SaleOrder.search([('x_ce_code', 'ilike', ce_code.strip())] + company_domain, limit=1)
        if so:
            return so

        # 4. Fuzzy match on x_studio_old_ce
        try:
            so = SaleOrder.search([('x_studio_old_ce', 'ilike', ce_code.strip())] + company_domain, limit=1)
            if so:
                return so
        except Exception:
            pass

        # 5. Normalized match (remove all spaces and compare)
        # Handles cases like XRN00004 vs XRN 00004
        norm_ce = self._normalize_ce_code(ce_code)
        all_sos = SaleOrder.search(company_domain)
        for so in all_sos:
            if so.x_ce_code and self._normalize_ce_code(so.x_ce_code) == norm_ce:
                return so
            old_ce = getattr(so, 'x_studio_old_ce', '')
            if old_ce and self._normalize_ce_code(old_ce) == norm_ce:
                return so

        return SaleOrder

    def _merge_opening_balance_rows(self, grouped_data):
        """
        Merge opening-balance-only rows into grouped_data.

        For CE codes that exist in the opening balance model but have NO accrual
        records (not already in grouped_data), create a row by pulling data from
        the matching sale.order (for live/current info) or falling back to the
        opening balance record's descriptive fields.

        These rows will have empty movement columns; only the first balance column
        (previous month ending balance) will be populated via _calculate_prev_month_balances().
        """
        cutoff_date = self._get_opening_balance_cutoff_date()
        if not cutoff_date:
            return

        try:
            ob_records = self.env[
                'saatchi.accrued_revenue_opening_balance'
            ].get_opening_balance_records_for_month(
                balance_date=cutoff_date,
                company_id=self.env.company.id
            )
        except Exception as e:
            _logger.warning('Error fetching opening balance records for OB rows: %s', str(e))
            return

        if not ob_records:
            return

        # Collect normalized CE codes already present in grouped_data
        existing_normalized_ces = set()
        for partner_name, ces in grouped_data.items():
            for ce_code_key in ces.keys():
                existing_normalized_ces.add(self._normalize_ce_code(ce_code_key))

        # Add OB-only rows for CEs not already present
        for norm_ce, ob_data in ob_records.items():
            if norm_ce in existing_normalized_ces:
                continue

            # Try to find matching sale.order for live data
            so = self._find_sale_order_by_ce_code(ob_data.get('ce_code_display', ''))

            if so:
                partner_name = (so.partner_id.name or ob_data.get('partner_name') or 'UNKNOWN').upper()
                ce_code_display = (
                    getattr(so, 'x_studio_old_ce', '') or
                    so.x_ce_code or
                    ob_data.get('ce_code_display', 'NO_CE')
                ).upper()
                old_ce_date = getattr(so, 'x_studio_old_ce_date', False)
                ce_date = old_ce_date or so.date_order or ob_data.get('ce_date')
                description = (getattr(so, 'x_job_description', '') or ob_data.get('job_description', '')).upper()

                # Get CE status from SO
                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_selection = dict(so._fields['x_ce_status'].selection)
                        ce_status = status_selection.get(so.x_ce_status, '').upper()
                    except Exception:
                        pass
                
                # Fallback to OB record's ce_status if available
                if not ce_status and ob_data.get('ce_status'):
                    ce_status = ob_data.get('ce_status', '').upper()
            else:
                partner_name = (ob_data.get('partner_name') or 'UNKNOWN').upper()
                ce_code_display = (ob_data.get('ce_code_display') or 'NO_CE').upper()
                ce_date = ob_data.get('ce_date')
                description = (ob_data.get('job_description') or '').upper()
                # Use ce_status directly from OB record (now a Char field)
                ce_status = (ob_data.get('ce_status') or '').upper()

            # Add to grouped_data with empty lines (no movement)
            grouped_data[partner_name][ce_code_display] = {
                'ce_date': ce_date,
                'description': description,
                'year': ce_date.year if ce_date else None,
                'ce_status': ce_status,
                'so_reference': '',
                'lines': [],  # No move lines - OB-only row
            }

            _logger.debug(
                'Added OB-only row for CE# %s (partner: %s)',
                ce_code_display, partner_name
            )

    def _merge_reversal_opening_balance_rows(self, grouped_data):
        """
        Merge reversal-opening-balance-only rows into grouped_data.

        For CE codes that exist in the reversal opening balance model but have NO
        accrual records (not already in grouped_data), create a row by pulling data
        from the matching sale.order (for live/current info) or falling back to the
        reversal opening balance record's descriptive fields.

        These rows will have empty movement columns; only balance columns
        (F from first OB, G & I from reversal OB) will be populated.

        This method is called BEFORE _merge_opening_balance_rows() so that
        reversal OB rows take priority over first OB rows.
        """
        cutoff_date = self._get_opening_balance_cutoff_date()
        if not cutoff_date:
            return

        try:
            reversal_ob_records = self.env[
                'saatchi.accrued_revenue_reversal_opening_balance'
            ].get_reversal_opening_balances_for_month(
                balance_date=cutoff_date,
                company_id=self.env.company.id
            )
        except Exception as e:
            _logger.warning('Error fetching reversal OB records for row merge: %s', str(e))
            return

        if not reversal_ob_records:
            return

        # Collect normalized CE codes already present in grouped_data
        existing_normalized_ces = set()
        for partner_name, ces in grouped_data.items():
            for ce_code_key in ces.keys():
                existing_normalized_ces.add(self._normalize_ce_code(ce_code_key))

        # Add reversal-OB-only rows for CEs not already present
        for norm_ce, rob_data in reversal_ob_records.items():
            if norm_ce in existing_normalized_ces:
                continue

            # Try to find matching sale.order for live data
            so = self._find_sale_order_by_ce_code(rob_data.get('ce_code_display', ''))

            if so:
                partner_name = (so.partner_id.name or rob_data.get('partner_name') or 'UNKNOWN').upper()
                ce_code_display = (
                    getattr(so, 'x_studio_old_ce', '') or
                    so.x_ce_code or
                    rob_data.get('ce_code_display', 'NO_CE')
                ).upper()
                old_ce_date = getattr(so, 'x_studio_old_ce_date', False)
                ce_date = old_ce_date or so.date_order or rob_data.get('ce_date')
                description = (getattr(so, 'x_job_description', '') or rob_data.get('job_description', '')).upper()

                # Get CE status from SO
                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_selection = dict(so._fields['x_ce_status'].selection)
                        ce_status = status_selection.get(so.x_ce_status, '').upper()
                    except Exception:
                        pass

                # Fallback to reversal OB record's ce_status if available
                if not ce_status and rob_data.get('ce_status'):
                    ce_status = rob_data.get('ce_status', '').upper()
            else:
                partner_name = (rob_data.get('partner_name') or 'UNKNOWN').upper()
                ce_code_display = (rob_data.get('ce_code_display') or 'NO_CE').upper()
                ce_date = rob_data.get('ce_date')
                description = (rob_data.get('job_description') or '').upper()
                ce_status = (rob_data.get('ce_status') or '').upper()

            # Add to grouped_data with empty lines (no movement)
            grouped_data[partner_name][ce_code_display] = {
                'ce_date': ce_date,
                'description': description,
                'year': ce_date.year if ce_date else None,
                'ce_status': ce_status,
                'so_reference': '',
                'lines': [],  # No move lines - reversal OB-only row
            }

            _logger.debug(
                'Added reversal-OB-only row for CE# %s (partner: %s)',
                ce_code_display, partner_name
            )

    def _calculate_prev_month_balances(self, accrued_account_ids, accrual_month):
        """
        Calculate the ending balance of the previous month for each partner/CE code.

        Uses a two-tier approach:
        1. PRIMARY: Query all posted accrued revenue account lines from the DB
           up to the last day of the previous month (cumulative balance).
        2. FALLBACK: For any CE# that has NO DB history, check the Opening Balance
           model if the previous month end matches the configured cutoff date.

        The DB balance always takes priority over the opening balance.

        Args:
            accrued_account_ids (list): Account IDs for the accrued revenue account
            accrual_month (date): First day of the current accrual month

        Returns:
            dict: {(PARTNER_NAME, CE_CODE): balance} where balance = debit - credit
        """
        prev_month_end = accrual_month - relativedelta(days=1)  # last day of previous month

        # ── TIER 1: Database transaction history ──
        prev_lines = self.env['account.move.line'].sudo().search([
            ('account_id', 'in', accrued_account_ids),
            ('date', '<=', prev_month_end),
            ('parent_state', '=', 'posted'),
        ])

        balances = defaultdict(float)
        db_ce_codes_normalized = set()  # Track which CE codes have DB history
        cutoff_date = self._get_opening_balance_cutoff_date()
        db_ce_codes_pre_cutoff = set()  # CE codes with DB history on or before cutoff

        for line in prev_lines:
            partner_name = (
                line.partner_id.name.upper()
                if line.partner_id and line.partner_id.name
                else 'UNKNOWN'
            )
            ce_code = line.x_ce_code.upper() if line.x_ce_code else 'NO_CE'
            net_amount = (line.debit or 0) - (line.credit or 0)
            balances[(partner_name, ce_code)] += net_amount

            # Track normalized CE code so we know it has DB history
            norm = self._normalize_ce_code(ce_code)
            db_ce_codes_normalized.add(norm)

            # Track CE codes with DB history on or before the cutoff date
            if cutoff_date and line.date <= cutoff_date:
                db_ce_codes_pre_cutoff.add(norm)

        # ── TIER 2: Opening Balance + Reversal OB ──
        # For the cutoff month (prev_month_end == cutoff_date):
        #   Add OB for CE#s with no DB history at all.
        # For subsequent months (prev_month_end > cutoff_date):
        #   Add OB for CE#s with no pre-cutoff DB history.
        #   Also add Reversal OB — these were used in January's reversal columns
        #   but are NOT stored in the DB. Without them the balance chain breaks.
        if cutoff_date and prev_month_end >= cutoff_date:
            is_cutoff_month = (prev_month_end == cutoff_date)
            # Decide which exclusion set to use:
            # - Cutoff month: exclude CE#s with ANY DB history
            # - Later months: exclude CE#s with pre-cutoff DB history only
            exclude_set = db_ce_codes_normalized if is_cutoff_month else db_ce_codes_pre_cutoff

            # Fetch regular OB
            try:
                opening_records = self.env[
                    'saatchi.accrued_revenue_opening_balance'
                ].get_opening_balance_records_for_month(
                    balance_date=cutoff_date,
                    company_id=self.env.company.id
                )
            except Exception as e:
                _logger.warning('Error fetching opening balances: %s', str(e))
                opening_records = {}

            # Fetch reversal OB (only needed for months AFTER cutoff)
            reversal_ob_records = {}
            if not is_cutoff_month:
                try:
                    reversal_ob_records = self.env[
                        'saatchi.accrued_revenue_reversal_opening_balance'
                    ].get_reversal_opening_balances_for_month(
                        balance_date=cutoff_date,
                        company_id=self.env.company.id
                    )
                except Exception as e:
                    _logger.warning('Error fetching reversal OB: %s', str(e))

            # Add regular OB amounts
            for norm_ce_code, ob_data in opening_records.items():
                if norm_ce_code not in exclude_set:
                    partner_name = (ob_data.get('partner_name') or 'UNKNOWN').upper()
                    ce_code_display = (ob_data.get('ce_code_display') or 'NO_CE').upper()
                    balances[(partner_name, ce_code_display)] += ob_data.get('balance', 0)

                    _logger.debug(
                        'OB for CE# %s: %.2f (prev_month_end=%s)',
                        ce_code_display, ob_data.get('balance', 0), prev_month_end
                    )

            # Add reversal OB amounts (for months after cutoff only)
            # These represent the reversal impact from the cutoff month that is NOT in DB.
            # Signs: system_reversal and manual_reversal are typically negative (credits).
            # manual_reversal_adjustment is subtracted in the report (makes reversal more negative).
            for norm_ce_code, rev_data in reversal_ob_records.items():
                if norm_ce_code not in exclude_set:
                    rev_system = rev_data.get('system_reversal', 0)
                    rev_manual = rev_data.get('manual_reversal', 0)
                    rev_adj = rev_data.get('manual_reversal_adjustment', 0)
                    # Match report formula: reversal columns add to balance,
                    # adjustment is subtracted from manual_reversal
                    rev_total = rev_system + rev_manual - rev_adj

                    if rev_total != 0:
                        # Find matching partner/CE from regular OB or reversal OB
                        partner_name = (rev_data.get('partner_name') or 'UNKNOWN').upper()
                        ce_code_display = (rev_data.get('ce_code_display') or 'NO_CE').upper()
                        balances[(partner_name, ce_code_display)] += rev_total

                        _logger.debug(
                            'Reversal OB for CE# %s: %.2f (prev_month_end=%s)',
                            ce_code_display, rev_total, prev_month_end
                        )

        return balances

    def _calculate_reversal_opening_balances(self, accrual_month):
        """
        Get reversal opening balances for the opening balance month.

        Uses the same 2-tier approach as _calculate_prev_month_balances:
        1. PRIMARY: Check for DB transaction history (reversal entries) for the
           previous month. If a CE# has DB reversal transactions, its computed
           amounts take priority.
        2. FALLBACK: For CE# with NO DB reversal history, use the
           reversal_opening_balance model.

        Only applies when the previous month end matches the configured cutoff date.

        Args:
            accrual_month (date): First day of the current accrual month

        Returns:
            dict: {normalized_ce_code: {
                'system_reversal': float,
                'manual_reversal': float
            }}
            Empty dict if not the opening balance month.
        """
        cutoff_date = self._get_opening_balance_cutoff_date()
        if not cutoff_date:
            return {}

        prev_month_end = accrual_month - relativedelta(days=1)
        if prev_month_end != cutoff_date:
            return {}

        # Identify CE codes that already have DB reversal history for this month
        # (these will use their computed amounts, not the OB)
        db_reversal_ce_codes = set()
        accrued_account_ids = self._get_accrued_revenue_account_id()
        if accrued_account_ids:
            accrual_month_end = (accrual_month + relativedelta(months=1)) - relativedelta(days=1)
            reversal_lines = self.env['account.move.line'].sudo().search([
                ('account_id', 'in', accrued_account_ids),
                ('date', '>=', accrual_month),
                ('date', '<=', accrual_month_end),
                ('parent_state', '=', 'posted'),
                ('x_type_of_entry', 'in', ['reversal_system', 'reversal_manual']),
            ])
            for line in reversal_lines:
                if line.x_ce_code:
                    db_reversal_ce_codes.add(self._normalize_ce_code(line.x_ce_code))

        # Fetch reversal OB records
        try:
            reversal_ob = self.env[
                'saatchi.accrued_revenue_reversal_opening_balance'
            ].get_reversal_opening_balances_for_month(
                balance_date=cutoff_date,
                company_id=self.env.company.id
            )
        except Exception as e:
            _logger.warning(
                'Error fetching reversal opening balances for %s: %s',
                cutoff_date, str(e)
            )
            return {}

        # Filter out CE codes that have DB history
        result = {}
        for norm_ce, ob_data in reversal_ob.items():
            if norm_ce not in db_reversal_ce_codes:
                result[norm_ce] = {
                    'system_reversal': ob_data.get('system_reversal', 0),
                    'manual_reversal': ob_data.get('manual_reversal', 0),
                    'manual_reversal_adjustment': ob_data.get('manual_reversal_adjustment', 0),
                }
                _logger.debug(
                    'Using reversal OB for CE# %s: sys=%.2f, manual=%.2f',
                    ob_data.get('ce_code_display', norm_ce),
                    ob_data.get('system_reversal', 0),
                    ob_data.get('manual_reversal', 0)
                )

        return result

    def _calculate_amounts_by_type(self, lines, accrual_month):
        """Calculate amounts for each entry type category"""
        amounts = {
            'system_reversal': 0,
            'system_accrual': 0,
            'manual_reversal': 0,
            'manual_reaccrual': 0,
            'manual_adjustment': 0
        }

        # Accrual month date range
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        for line in lines:
            if not line.date:
                continue

            # Calculate net amount (debit - credit)
            net_amount = (line.debit or 0) - (line.credit or 0)

            # Categorize based on type and date
            if line.x_type_of_entry == 'reversal_system':
                # Reversals dated on the 1st of accrual month
                if line.date.month == accrual_month.month and line.date.year == accrual_month.year:
                    amounts['system_reversal'] += net_amount

            elif line.x_type_of_entry == 'accrued_system':
                # Accruals dated within the accrual month
                if accrual_month <= line.date <= accrual_month_end:
                    amounts['system_accrual'] += net_amount

            elif line.x_type_of_entry == 'reversal_manual':
                # Manual reversals dated on the 1st of accrual month
                if line.date.month == accrual_month.month and line.date.year == accrual_month.year:
                    amounts['manual_reversal'] += net_amount

            elif line.x_type_of_entry == 'accrued_manual':
                # Manual re-accruals dated within the accrual month
                if accrual_month <= line.date <= accrual_month_end:
                    amounts['manual_reaccrual'] += net_amount

            elif line.x_type_of_entry in ['adjustment_system', 'adjustment_manual']:
                # Adjustments & their reversals within the accrual month
                # (adjustment and reversal land in different months by design)
                if accrual_month <= line.date <= accrual_month_end:
                    amounts['manual_adjustment'] += net_amount

        return amounts

    def generate_xlsx_report(self, workbook, data, lines):
        """
        Main report generation method.
        Now uses start_date and end_date from wizard data instead of relying on selected records.
        """
        # Get accrued revenue account
        accrued_account_ids = self._get_accrued_revenue_account_id()
        if not accrued_account_ids:
            raise UserError(
                "Accrued Revenue account not configured. Please set it in system parameters.")

        # Filter lines to only accrued revenue account
        filtered_lines = lines.filtered(
            lambda l: l.account_id.id in accrued_account_ids)

        # Get date range from wizard data
        if 'start_date' not in data or 'end_date' not in data:
            raise UserError(
                "Date range parameters are missing. Please use the Accrued Revenue wizard to generate the report with start and end dates.")
        
        start_date = datetime.datetime.strptime(
            data['start_date'], '%Y-%m-%d').date()
        end_date = datetime.datetime.strptime(
            data['end_date'], '%Y-%m-%d').date()

        # Define formats
        formats = self._define_formats(workbook)

        # Determine accrual months based on wizard date range
        accrual_months = self._determine_accrual_months(
            filtered_lines, start_date, end_date)

        # Allow empty filtered_lines if an opening balance month is in range
        if not filtered_lines:
            has_ob_month = False
            cutoff_date = self._get_opening_balance_cutoff_date()
            if cutoff_date:
                for m in accrual_months:
                    if (m - relativedelta(days=1)) == cutoff_date:
                        has_ob_month = True
                        break
            if not has_ob_month:
                raise UserError(
                    "No accrued revenue entries found for the selected date range.")

        # Generate sheets for each month
        for accrual_month in accrual_months:
            self._generate_month_sheets(
                workbook, formats, filtered_lines, lines, accrual_month, accrued_account_ids[0])

        return True

    def _generate_month_sheets(self, workbook, formats, filtered_lines, all_lines, accrual_month, accrued_account_id):
        """Generate all sheets for a specific accrual month"""
        prev_month = accrual_month - relativedelta(months=1)

        # Get last day of previous month and current month
        prev_month_last_day = (accrual_month - relativedelta(days=1)).day
        accrual_month_last_day = (
            (accrual_month + relativedelta(months=1)) - relativedelta(days=1)).day

        # Format month names
        prev_month_abbr = prev_month.strftime('%b').upper()
        accrual_month_abbr = accrual_month.strftime('%b').upper()

        # Format full month name and year for title
        accrual_month_full = accrual_month.strftime('%B %Y').upper()

        # Report date
        report_date = accrual_month.strftime('%m/%d/%Y')

        # Calculate previous month ending balances from the database
        prev_month_balances = self._calculate_prev_month_balances(
            [accrued_account_id], accrual_month)

        # Calculate reversal opening balances (only populated for OB month)
        reversal_ob_balances = self._calculate_reversal_opening_balances(accrual_month)

        # Group lines by client and CE# for this specific month
        grouped_data = self._group_lines_by_ce(filtered_lines, accrual_month)

        # If this is the opening balance month OR any month after cutoff,
        # merge OB-only rows so CE#s from the OB continue to appear.
        # Reversal OB rows are merged FIRST (higher priority), then first OB rows.
        cutoff_date = self._get_opening_balance_cutoff_date()
        if cutoff_date:
            prev_month_end = accrual_month - relativedelta(days=1)
            if prev_month_end >= cutoff_date:
                self._merge_reversal_opening_balance_rows(grouped_data)
                self._merge_opening_balance_rows(grouped_data)

        # Skip this month if no data (even after OB merge)
        if not grouped_data:
            return

        # Create main sheet with formatted name (e.g., "November 2025")
        sheet_name_main = accrual_month.strftime('%B %Y')
        sheet = workbook.add_worksheet(sheet_name_main)

        # Set column widths
        sheet.set_column(0, 0, 30)   # CLIENT
        sheet.set_column(1, 1, 20)   # SO REFERENCE
        sheet.set_column(2, 2, 15)   # CE#
        sheet.set_column(3, 3, 12)   # CE DATE
        sheet.set_column(4, 4, 40)   # DESCRIPTION
        sheet.set_column(5, 5, 8)    # Year
        sheet.set_column(6, 6, 15)   # Balance (prev month)
        sheet.set_column(7, 7, 18)   # System Reversal
        sheet.set_column(8, 8, 18)   # System Accrual
        sheet.set_column(9, 9, 18)   # Manual Reversal
        sheet.set_column(10, 10, 18)  # Manual Re-accrual
        sheet.set_column(11, 11, 18)  # Manual Adjustment
        sheet.set_column(12, 12, 15)  # Balance (current month)
        sheet.set_column(13, 13, 20)  # CE Status

        # Write report header
        row = 0
        sheet.write(
            row, 0, self.env.company.name, formats['title'])
        row += 1
        sheet.write(
            row, 0, f'ACCRUED REVENUE - {accrual_month_full}', formats['title'])
        row += 1
        sheet.write(row, 0, report_date, formats['title'])
        row += 2

        # Write section headers row
        sheet.write(
            row, 6, f'Balance', formats['section_header'])
        sheet.merge_range(
            row, 7, row, 8, 'REVENUE ACCRUAL - SYSTEM', formats['section_header'])
        sheet.merge_range(row, 9, row, 11, 'MANUAL ADJUSTMENT',
                          formats['section_header'])
        sheet.write(
            row, 12, f'Balance', formats['section_header'])

        row += 1

        # Write column headers
        headers = [
            'CLIENT', 'SO REFERENCE', 'CE#', 'CE DATE', 'DESCRIPTION', 'Year',
            f'{prev_month_last_day}-{prev_month_abbr}',
            f'{prev_month_abbr} REVERSAL',
            f'{accrual_month_abbr} ACCRUAL',
            f'{prev_month_abbr} REVERSAL',
            f'{accrual_month_abbr} RE-ACCRUAL',
            f'{accrual_month_abbr} ADDL ADJ',
            f'{accrual_month_last_day}-{accrual_month_abbr}',
            'CE STATUS'
        ]

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Write data rows
        for partner_name in sorted(grouped_data.keys()):
            ces = grouped_data[partner_name]

            for ce_code in sorted(ces.keys()):
                ce_data = ces[ce_code]
                amounts = self._calculate_amounts_by_type(
                    ce_data['lines'], accrual_month)

                # Check if this is an OB-only row (no accrued lines)
                is_ob_only_row = not ce_data['lines']
                
                # If OB-only, verify it also doesn't exist in sale.order
                if is_ob_only_row:
                    # Use the existing method that checks x_ce_code and x_studio_old_ce
                    so_match = self._find_sale_order_by_ce_code(ce_code)
                    # Only mark red if CE doesn't exist in any sale order
                    is_ob_only_row = not so_match

                sheet.write(row, 0, partner_name, formats['normal'])
                
                # Determine CE# format based on OB and reversal OB presence
                sheet.write(row, 1, ce_data['so_reference'], formats['centered'])
                sheet.write(row, 2, ce_code, formats['centered'])

                if ce_data['ce_date']:
                    sheet.write(row, 3, ce_data['ce_date'], formats['date'])
                else:
                    sheet.write(row, 3, '', formats['centered'])

                sheet.write(row, 4, ce_data['description'], formats['normal'])

                if ce_data['year']:
                    sheet.write(row, 5, ce_data['year'], formats['centered'])
                else:
                    sheet.write(row, 5, '', formats['centered'])

                # Get previous month ending balance.
                # First try exact match, then fall back to normalized CE# matching.
                # SUM all matching entries (OB and DB may have different partner keys).
                prev_balance = prev_month_balances.get(
                    (partner_name, ce_code), None)

                # Always do normalized CE# matching to catch entries with different
                # partner names (e.g., OB partner vs DB journal partner).
                # Sum ALL entries with the same normalized CE code.
                norm_ce = self._normalize_ce_code(ce_code)
                total_from_all_keys = 0
                found_any = False
                for (bal_partner, bal_ce), bal_amount in prev_month_balances.items():
                    if self._normalize_ce_code(bal_ce) == norm_ce:
                        total_from_all_keys += bal_amount
                        found_any = True

                if found_any:
                    prev_balance = total_from_all_keys

                prev_balance = prev_balance or 0
                sheet.write(row, 6, prev_balance, formats['currency'])

                # Check for reversal opening balance overrides (OB month only)
                norm_ce_for_rev = self._normalize_ce_code(ce_code)
                rev_ob = reversal_ob_balances.get(norm_ce_for_rev, {})

                # Column H (7): System Reversal
                system_reversal_val = amounts['system_reversal']
                if system_reversal_val == 0 and rev_ob.get('system_reversal', 0) != 0:
                    system_reversal_val = rev_ob['system_reversal']
                sheet.write(
                    row, 7, system_reversal_val, formats['currency_negative'])

                sheet.write(
                    row, 8, amounts['system_accrual'], formats['currency_negative'])

                # Column J (9): Manual Reversal (includes adjustment from reversal OB)
                manual_reversal_val = amounts['manual_reversal']
                if manual_reversal_val == 0 and rev_ob.get('manual_reversal', 0) != 0:
                    manual_reversal_val = rev_ob['manual_reversal']
                # Subtract manual reversal adjustment from reversal OB (combined into col I)
                manual_reversal_val -= rev_ob.get('manual_reversal_adjustment', 0)
                sheet.write(
                    row, 9, manual_reversal_val, formats['currency_negative'])
                sheet.write(
                    row, 10, amounts['manual_reaccrual'], formats['currency_negative'])
                sheet.write(
                    row, 11, amounts['manual_adjustment'], formats['currency_negative'])

                excel_row = row + 1
                sheet.write_formula(
                    row, 12, f'=G{excel_row}+H{excel_row}+I{excel_row}+J{excel_row}+K{excel_row}+L{excel_row}', formats['currency'])

                sheet.write(row, 13, self._normalize_ce_status(ce_data['ce_status']), formats['centered'])

                row += 1

        # Add totals row
        total_row = row

        # Create bold formats for totals
        base_font = {'font_name': 'Calibri', 'font_size': 10}
        currency_bold_format = workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': '#,##0.00;-#,##0.00;"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        currency_negative_bold_format = workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        bold_with_border = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Empty cells before TOTAL label
        for col in range(0, 5):
            sheet.write(total_row, col, '',
                        formats['section_header_no_border'])

        # TOTAL label
        sheet.write(total_row, 5, 'TOTAL', bold_with_border)

        # Sum formulas for monetary columns
        sheet.write_formula(
            total_row, 6, f'=SUM(G{data_start_row + 1}:G{total_row})', currency_bold_format)
        sheet.write_formula(
            total_row, 7, f'=SUM(H{data_start_row + 1}:H{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 8, f'=SUM(I{data_start_row + 1}:I{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 9, f'=SUM(J{data_start_row + 1}:J{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 10, f'=SUM(K{data_start_row + 1}:K{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 11, f'=SUM(L{data_start_row + 1}:L{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 12, f'=SUM(M{data_start_row + 1}:M{total_row})', currency_bold_format)

        # Empty cell after totals
        sheet.write(total_row, 13, '', formats['section_header_no_border'])

        # Generate additional sheets
        self._generate_accrual_breakdown_sheet(
            workbook, formats, all_lines, accrual_month, accrued_account_id)
        self._generate_gl_sheet(
            workbook, formats, all_lines, accrual_month, accrued_account_id)

    def _generate_accrual_breakdown_sheet(self, workbook, formats, lines, accrual_month, accrued_account_id):
        """Generate the breakdown sheet for accrual entries"""
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        # Filter accrual lines (excluding accrued revenue account) within the accrual month
        accrual_lines = lines.filtered(lambda l:
                                       l.x_type_of_entry in ['accrued_system', 'accrued_manual'] and
                                       l.account_id.id != accrued_account_id and
                                       l.date and
                                       accrual_month <= l.date <= accrual_month_end
                                       )

        if not accrual_lines:
            return

        sheet_name = accrual_month.strftime('%B %Y Accruals')
        sheet = workbook.add_worksheet(sheet_name)

        # Set column widths
        sheet.set_column(0, 0, 12)   # Date
        sheet.set_column(1, 1, 15)   # Entry Type
        sheet.set_column(2, 2, 20)   # Journal Entry
        sheet.set_column(3, 3, 25)   # Account
        sheet.set_column(4, 4, 30)   # Client Name
        sheet.set_column(5, 5, 15)   # CE Code
        sheet.set_column(6, 6, 12)   # CE Date
        sheet.set_column(7, 7, 35)   # Label
        sheet.set_column(8, 8, 20)   # Reference
        sheet.set_column(9, 9, 15)   # C.E. Status
        sheet.set_column(10, 10, 15)  # Debit
        sheet.set_column(11, 11, 15)  # Credit
        sheet.set_column(12, 12, 30)  # Remarks

        row = 0
        headers = ['Date', 'Entry Type', 'Journal Entry', 'Account', 'Client Name', 'CE Code',
                   'CE Date', 'Label', 'Reference', 'C.E. Status', 'Debit', 'Credit', 'Remarks']

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Sort lines by date and partner name
        sorted_lines = accrual_lines.sorted(key=lambda l: (
            l.date or datetime.date.min, l.partner_id.name or ''))

        # Write data rows
        for line in sorted_lines:
            # Date
            if line.date:
                sheet.write(row, 0, line.date, formats['date'])
            else:
                sheet.write(row, 0, '', formats['centered'])

            # Entry Type
            entry_type_display = 'Accrued - System' if line.x_type_of_entry == 'accrued_system' else 'Accrued - Manual'
            sheet.write(row, 1, entry_type_display, formats['normal'])

            # Journal Entry
            sheet.write(
                row, 2, line.move_id.name if line.move_id else '', formats['normal'])

            # Account
            sheet.write(
                row, 3, line.account_id.display_name if line.account_id else '', formats['normal'])

            # Client Name
            sheet.write(
                row, 4, line.partner_id.name if line.partner_id else '', formats['normal'])

            # CE Code
            sheet.write(row, 5, line.x_ce_code or '', formats['centered'])

            # CE Date
            if line.x_ce_date:
                sheet.write(row, 6, line.x_ce_date, formats['date'])
            else:
                sheet.write(row, 6, '', formats['centered'])

            # Label
            sheet.write(row, 7, line.name or '', formats['normal'])

            # Reference
            sheet.write(row, 8, line.x_reference or '', formats['normal'])

            # CE Status
            selection_dict = dict(line._fields['x_ce_status'].selection)
            ce_status = selection_dict.get(line.x_ce_status, '')
            sheet.write(row, 9, ce_status, formats['centered'])

            # Debit
            sheet.write(row, 10, line.debit or 0, formats['currency'])

            # Credit
            sheet.write(row, 11, line.credit or 0, formats['currency'])

            # Remarks
            sheet.write(row, 12, line.x_studio_remarks or '',
                        formats['normal'])

            row += 1

        # Add totals row
        total_row = row

        # Create bold formats for totals
        base_font = {'font_name': 'Calibri', 'font_size': 10}
        currency_bold_format = workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': '#,##0.00;-#,##0.00;"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        bold_with_border = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Empty cells before TOTAL label
        for col in range(0, 9):
            sheet.write(total_row, col, '',
                        formats['section_header_no_border'])

        # TOTAL label
        sheet.write(total_row, 9, 'TOTAL', bold_with_border)

        # Sum formulas for Debit and Credit columns
        sheet.write_formula(
            total_row, 10, f'=SUM(K{data_start_row + 1}:K{total_row})', currency_bold_format)
        sheet.write_formula(
            total_row, 11, f'=SUM(L{data_start_row + 1}:L{total_row})', currency_bold_format)

        # Empty cell for Remarks
        sheet.write(total_row, 12, '', formats['section_header_no_border'])

        return True

    def _generate_gl_sheet(self, workbook, formats, lines, accrual_month, accrued_account_id):
        """Generate the GL sheet for all accrued revenue account entries"""
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        # Filter GL lines (accrued revenue account only) within the accrual month
        gl_lines = lines.filtered(lambda l:
                                  l.account_id.id == accrued_account_id and
                                  l.date and
                                  accrual_month <= l.date <= accrual_month_end
                                  )

        if not gl_lines:
            return

        month_name = accrual_month.strftime('%B')
        sheet_name = f'GL_{month_name}'
        sheet = workbook.add_worksheet(sheet_name)

        # Create bold formats for totals
        base_font = {'font_name': 'Calibri', 'font_size': 10}
        currency_bold_format = workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': '#,##0.00;-#,##0.00;"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        currency_negative_bold_format = workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

        # Set column widths
        sheet.set_column(0, 0, 12)   # Date
        sheet.set_column(1, 1, 15)   # Entry Type
        sheet.set_column(2, 2, 20)   # Journal Entry
        sheet.set_column(3, 3, 25)   # Account
        sheet.set_column(4, 4, 30)   # Client Name
        sheet.set_column(5, 5, 15)   # CE Code
        sheet.set_column(6, 6, 12)   # CE Date
        sheet.set_column(7, 7, 35)   # Label
        sheet.set_column(8, 8, 20)   # Reference
        sheet.set_column(9, 9, 15)  # Debit
        sheet.set_column(10, 10, 15)  # Credit
        sheet.set_column(11, 11, 15)  # DR Less CR
        sheet.set_column(12, 12, 30)  # Remarks

        row = 0
        headers = ['Date', 'Entry Type', 'Journal Entry', 'Account', 'Client Name', 'CE Code',
                   'CE Date', 'Label', 'Reference', 'Debit', 'Credit', 'DR Less CR', 'Remarks']

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Get accrued account for opening balance row
        accrued_account = self.env['account.account'].browse(accrued_account_id)
        
        # Calculate opening balance (previous month's ending balance)
        prev_month_end = accrual_month - relativedelta(days=1)
        prev_balance_lines = lines.filtered(lambda l:
                                           l.account_id.id == accrued_account_id and
                                           l.date and
                                           l.date <= prev_month_end and
                                           l.parent_state == 'posted'
                                           )
        
        opening_balance = sum(prev_balance_lines.mapped(lambda l: l.debit - l.credit))

        # Write opening balance row
        sheet.write(row, 0, accrual_month, formats['date'])  # Date: first day of month
        sheet.write(row, 1, 'Opening Balance', formats['normal'])  # Entry Type
        sheet.write(row, 2, '', formats['normal'])  # Journal Entry (blank)
        sheet.write(row, 3, accrued_account.display_name if accrued_account else '', formats['normal'])  # Account
        sheet.write(row, 4, '', formats['normal'])  # Client Name (blank)
        sheet.write(row, 5, '', formats['centered'])  # CE Code (blank)
        sheet.write(row, 6, '', formats['centered'])  # CE Date (blank)
        sheet.write(row, 7, 'Opening Balance', formats['normal'])  # Label
        sheet.write(row, 8, '', formats['normal'])  # Reference (blank)
        sheet.write(row, 9, 0, formats['currency'])  # Debit (0)
        sheet.write(row, 10, 0, formats['currency'])  # Credit (0)
        sheet.write(row, 11, opening_balance, formats['currency'])  # DR Less CR (opening balance)
        sheet.write(row, 12, '', formats['normal'])  # Remarks (blank)
        row += 1

        # Sort lines by date and partner name
        sorted_lines = gl_lines.sorted(key=lambda l: (
            l.date or datetime.date.min, l.partner_id.name or ''))

        # Entry type mapping
        entry_type_map = {
            'accrued_system': 'Accrued - System',
            'accrued_manual': 'Accrued - Manual',
            'reversal_system': 'Reversal - System',
            'reversal_manual': 'Reversal - Manual',
            'adjustment_system': 'Adjustment - System',
            'adjustment_manual': 'Adjustment - Manual'
        }

        # Write data rows
        for line in sorted_lines:
            # Date
            if line.date:
                sheet.write(row, 0, line.date, formats['date'])
            else:
                sheet.write(row, 0, '', formats['centered'])

            # Entry Type
            sheet.write(row, 1, entry_type_map.get(
                line.x_type_of_entry, ''), formats['normal'])

            # Journal Entry
            sheet.write(
                row, 2, line.move_id.name if line.move_id else '', formats['normal'])

            # Account
            sheet.write(
                row, 3, line.account_id.display_name if line.account_id else '', formats['normal'])

            # Client Name
            sheet.write(
                row, 4, line.partner_id.name if line.partner_id else '', formats['normal'])

            # CE Code
            sheet.write(row, 5, line.x_ce_code or '', formats['centered'])

            # CE Date
            if line.x_ce_date:
                sheet.write(row, 6, line.x_ce_date, formats['date'])
            else:
                sheet.write(row, 6, '', formats['centered'])

            # Label
            sheet.write(row, 7, line.name or '', formats['normal'])

            # Reference
            sheet.write(row, 8, line.x_reference or '', formats['normal'])

            # Debit
            sheet.write(row, 9, line.debit or 0, formats['currency'])

            # Credit
            sheet.write(row, 10, line.credit or 0, formats['currency'])

            # DR Less CR (formula)
            excel_row = row + 1
            sheet.write_formula(
                row, 11, f'=J{excel_row}-K{excel_row}', formats['currency_negative'])

            # Remarks
            sheet.write(row, 12, line.x_studio_remarks or '',
                        formats['normal'])

            row += 1

        # Add totals row
        total_row = row

        # Empty cells before TOTAL label
        for col in range(0, 8):
            sheet.write(total_row, col, '',
                        formats['section_header_no_border'])

        # TOTAL label
        bold_with_border = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 10,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        sheet.write(total_row, 8, 'TOTAL', bold_with_border)

        # Sum formulas for monetary columns
        sheet.write_formula(
            total_row, 9, f'=SUM(J{data_start_row + 1}:J{total_row})', currency_bold_format)
        sheet.write_formula(
            total_row, 10, f'=SUM(K{data_start_row + 1}:K{total_row})', currency_bold_format)
        sheet.write_formula(
            total_row, 11, f'=SUM(L{data_start_row + 1}:L{total_row})', currency_negative_bold_format)
        sheet.write(total_row, 12, '', formats['section_header_no_border'])

        return True
