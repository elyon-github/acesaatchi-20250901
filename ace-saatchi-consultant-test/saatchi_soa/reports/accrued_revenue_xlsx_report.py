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
                template_account = self.env['account.account'].sudo().browse(
                    account_id)
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
            # Every sale order contributing to this CE row.  Joined into
            # so_reference at the end of this method -- see the comment there.
            'so_names': set(),
            'lines': []
        }))

        # Filter lines for this specific accrual month
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        month_lines = []
        for line in lines:
            if not line.date:
                continue
            if line.x_type_of_entry in ['reversal_system', 'reversal_manual']:
                if line.date.month == accrual_month.month and line.date.year == accrual_month.year:
                    month_lines.append(line)
            elif line.x_type_of_entry in ['accrued_system', 'accrued_manual', 'adjustment_system', 'adjustment_manual']:
                if accrual_month <= line.date <= accrual_month_end:
                    month_lines.append(line)
            elif not line.x_type_of_entry:
                if line.x_ce_code:
                    if accrual_month <= line.date <= accrual_month_end:
                        month_lines.append(line)

        def _ce_canon(c):
            """Canonical form for zero-padding comparison only (used for orphan JE matching)."""
            c = re.sub(r'\s+', '', (c or '').strip().upper())
            m = re.match(r'^([A-Z]*)(\d+)(.*)$', c)
            if m:
                return m.group(1) + str(int(m.group(2))) + m.group(3)
            return c

        # Process regular typed lines first so their CE rows exist
        # before orphan JEs try to match against them.
        regular_lines = [l for l in month_lines if l.x_type_of_entry]
        orphan_lines = [l for l in month_lines if not l.x_type_of_entry]

        for line in regular_lines + orphan_lines:
            partner_name = line.partner_id.name.upper(
            ) if line.partner_id and line.partner_id.name else 'UNKNOWN'
            ce_code = line.x_ce_code.upper() if line.x_ce_code else 'NO_CE'

            # For orphan JEs only: if a regular-accrual row already exists for this
            # partner with the same CE code after stripping leading zeros, fold the
            # orphan into that row instead of creating a separate "BA00003" row.
            if not line.x_type_of_entry:
                existing_ces = dict.get(grouped, partner_name)
                if existing_ces is not None:
                    ce_canon = _ce_canon(ce_code)
                    for existing_key in existing_ces.keys():
                        if _ce_canon(existing_key) == ce_canon:
                            ce_code = existing_key
                            break

            grouped[partner_name][ce_code]['lines'].append(line)

            if not grouped[partner_name][ce_code]['ce_date']:
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
                    selection_dict = dict(
                        accrued_record._fields['ce_status'].selection)
                    ce_status = selection_dict.get(
                        accrued_record.ce_status, '')
                    grouped[partner_name][ce_code]['ce_status'] = ce_status.upper()
            # Collect EVERY sale order feeding this CE row, not just the first.
            # Several sale orders can legitimately share one old CE# -- the
            # client does this deliberately for recurring monthly work (e.g.
            # 'SSM 00025' covers six separate monthly jobs).  Because the row is
            # keyed on CE, showing only the first SO made the others look like
            # they were missing from the report entirely.
            if line.move_id and line.move_id.x_related_custom_accrued_record:
                accrued_record = line.move_id.x_related_custom_accrued_record
                so_ref = accrued_record.x_related_ce_id.name if accrued_record.x_related_ce_id else ''
                if so_ref:
                    grouped[partner_name][ce_code]['so_names'].add(so_ref.upper())

        # Render the collected sale orders as the SO REFERENCE cell, pipe
        # separated.  Sorted so the output is stable between runs, and
        # de-duplicated by the set above.
        # Display only -- no amount, grouping or total depends on this value.
        for ces in grouped.values():
            for ce_data in ces.values():
                if ce_data.get('so_names'):
                    ce_data['so_reference'] = ' | '.join(sorted(ce_data['so_names']))

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

    def _normalize_ce_zero_collapse(self, ce_code):
        """
        Last-resort normalization that collapses any run of 2+ consecutive zeros
        to a single zero, after removing whitespace and uppercasing.

        This handles the pattern where a JE CE code has one extra zero inserted
        into an existing zero-run compared to the system CE code, e.g.:
            SSY0001  <->  SSY00001   (both -> SSY01)
            LBC0036  <->  LBC00036   (both -> LBC036)
            PHO0249A <->  PHO00249A  (both -> PHO0249A)
            B450009  <->  B4500009   (both -> B45009)

        Only used as Step 4 fallback after exact, spaces-only, and full-normalize
        all fail, to minimise false-positive matches.
        """
        if not ce_code:
            return ''
        s = re.sub(r'\s+', '', ce_code.strip().upper())
        return re.sub(r'0{2,}', '0', s)

    def _consolidate_ce_rows(self, grouped_data):
        """
        Post-process grouped_data to merge rows whose CE codes are identical
        after stripping whitespace and leading zeros (e.g. BA0003 == BA00003).
        The row that has actual transaction lines is kept as the canonical key;
        other rows' lines and metadata are folded into it.
        Called right after _group_lines_by_ce, before OB merges.
        """
        def _canon(c):
            c = re.sub(r'\s+', '', (c or '').strip().upper())
            m = re.match(r'^([A-Z]*)(\d+)(.*)$', c)
            if m:
                return m.group(1) + str(int(m.group(2))) + m.group(3)
            return c

        for partner_name in list(grouped_data.keys()):
            ces = grouped_data[partner_name]
            # Group CE keys by canonical form
            canon_map = defaultdict(list)
            for key in list(ces.keys()):
                canon_map[_canon(key)].append(key)

            for keys in canon_map.values():
                if len(keys) <= 1:
                    continue
                # Primary: prefer the row that already has description/ce_status
                # (regular accrual rows have these; orphan JEs never do).
                # This ensures orphan JEs always fold INTO the regular row, not vice versa.
                has_meta = [k for k in keys if ces[k].get(
                    'description') or ces[k].get('ce_status')]
                has_lines = [k for k in keys if ces[k]['lines']]
                if has_meta:
                    primary = next(
                        (k for k in has_meta if ces[k]['lines']), has_meta[0])
                elif has_lines:
                    primary = has_lines[0]
                else:
                    primary = min(keys)
                for key in keys:
                    if key == primary:
                        continue
                    src = ces[key]
                    dst = ces[primary]
                    dst['lines'].extend(src['lines'])
                    if not dst['ce_date'] and src['ce_date']:
                        dst['ce_date'] = src['ce_date']
                        dst['year'] = src['year']
                    if not dst['description'] and src['description']:
                        dst['description'] = src['description']
                    if not dst['ce_status'] and src['ce_status']:
                        dst['ce_status'] = src['ce_status']
                    if not dst['so_reference'] and src['so_reference']:
                        dst['so_reference'] = src['so_reference']
                    del ces[key]

    def _assign_prev_balances(self, grouped_data, prev_month_balances, prev_month_ce_dates):
        """Assign each previous-month balance to exactly ONE sheet row.

        Rows are grouped by (partner, CE) but the balance lookup matches on CE
        alone.  So a CE appearing under two partners -- which happens when the
        customer on a sale order is changed mid-year, leaving already-posted
        lines under the old partner -- had its opening balance claimed by BOTH
        rows and counted twice.  (Observed: CE 'PVI 00002', June 2026,
        40,910.82 double-counted.)

        Claiming is two-tier so existing behaviour is preserved:
          1. a row whose partner matches the balance's partner claims it first;
          2. anything still unclaimed goes to the first row matching on CE
             alone, which keeps partner-name variants working as they do today
             (e.g. "SM PRIME HOLDINGS, INC." vs "SM PRIME HOLDINGS INC.").

        A balance claimed once today is still claimed exactly once, so a total
        can only shed a genuine duplicate -- it can never gain an omission.

        Returns:
            dict: {(partner_name, ce_code): balance}
        """
        row_keys = [
            (partner_name, ce_code)
            for partner_name in sorted(grouped_data.keys())
            for ce_code in sorted(grouped_data[partner_name].keys())
        ]

        def _matching_balance_keys(row_key):
            """Balance keys matching this row -- mirrors the original lookup."""
            partner_name, ce_code = row_key
            norm_ce = self._normalize_ce_code(ce_code)
            norm_ce_collapsed = self._normalize_ce_zero_collapse(ce_code)
            row_ce_date = grouped_data[partner_name][ce_code].get('ce_date')
            matches = []
            for bal_key in prev_month_balances:
                bal_ce = bal_key[1]
                if self._normalize_ce_code(bal_ce) == norm_ce:
                    matches.append(bal_key)
                elif self._normalize_ce_zero_collapse(bal_ce) == norm_ce_collapsed:
                    # Zero-collapse match only: if both sides carry a CE date and
                    # they differ, this is a genuinely different CE -- skip it.
                    bal_ce_date = prev_month_ce_dates.get(bal_key)
                    if not row_ce_date or not bal_ce_date or row_ce_date == bal_ce_date:
                        matches.append(bal_key)
            return matches

        candidates = {rk: _matching_balance_keys(rk) for rk in row_keys}
        claimed = {}

        # Tier 1: a row with the same partner as the balance claims it.
        for row_key in row_keys:
            for bal_key in candidates[row_key]:
                if bal_key not in claimed and bal_key[0] == row_key[0]:
                    claimed[bal_key] = row_key

        # Tier 2: remaining balances go to the first row matching on CE alone.
        for row_key in row_keys:
            for bal_key in candidates[row_key]:
                if bal_key not in claimed:
                    claimed[bal_key] = row_key

        totals = defaultdict(float)
        for bal_key, row_key in claimed.items():
            totals[row_key] += prev_month_balances[bal_key]
        return totals

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
            _logger.warning(
                'Could not retrieve opening balance cutoff date: %s', str(e))
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
        so = SaleOrder.search(
            [('x_ce_code', '=', ce_code)] + company_domain, limit=1)
        if so:
            return so

        # 2. Exact match on x_studio_old_ce (Studio field)
        try:
            so = SaleOrder.search(
                [('x_studio_old_ce', '=', ce_code)] + company_domain, limit=1)
            if so:
                return so
        except Exception:
            pass

        # 3. Fuzzy match on x_ce_code (whitespace/case variations)
        so = SaleOrder.search(
            [('x_ce_code', 'ilike', ce_code.strip())] + company_domain, limit=1)
        if so:
            return so

        # 4. Fuzzy match on x_studio_old_ce
        try:
            so = SaleOrder.search(
                [('x_studio_old_ce', 'ilike', ce_code.strip())] + company_domain, limit=1)
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
            _logger.warning(
                'Error fetching opening balance records for OB rows: %s', str(e))
            return

        if not ob_records:
            return

        # Map norm_ce -> (partner_name, ce_code_key) so we can patch existing rows
        existing_ces_map = {}
        for partner_name, ces in grouped_data.items():
            for ce_code_key in ces.keys():
                existing_ces_map[self._normalize_ce_code(
                    ce_code_key)] = (partner_name, ce_code_key)

        # Add OB-only rows for CEs not already present.
        # For existing rows, patch blank description/ce_status from OB data.
        for norm_ce, ob_data in ob_records.items():
            if norm_ce in existing_ces_map:
                p_name, ce_key = existing_ces_map[norm_ce]
                self._patch_metadata_from_ob(
                    grouped_data[p_name][ce_key], ob_data)
                continue

            # Try to find matching sale.order for live data
            so = self._find_sale_order_by_ce_code(
                ob_data.get('ce_code_display', ''))

            if so:
                partner_name = (so.partner_id.name or ob_data.get(
                    'partner_name') or 'UNKNOWN').upper()
                ce_code_display = (
                    getattr(so, 'x_studio_old_ce', '') or
                    so.x_ce_code or
                    ob_data.get('ce_code_display', 'NO_CE')
                ).upper()
                old_ce_date = getattr(so, 'x_studio_old_ce_date', False)
                ce_date = old_ce_date or so.date_order or ob_data.get(
                    'ce_date')
                description = (getattr(so, 'x_job_description', '')
                               or ob_data.get('job_description', '')).upper()

                # Get CE status from SO
                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_selection = dict(
                            so._fields['x_ce_status'].selection)
                        ce_status = status_selection.get(
                            so.x_ce_status, '').upper()
                    except Exception:
                        pass

                # Fallback to OB record's ce_status if available
                if not ce_status and ob_data.get('ce_status'):
                    ce_status = ob_data.get('ce_status', '').upper()
            else:
                partner_name = (ob_data.get('partner_name')
                                or 'UNKNOWN').upper()
                ce_code_display = (ob_data.get(
                    'ce_code_display') or 'NO_CE').upper()
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
            _logger.warning(
                'Error fetching reversal OB records for row merge: %s', str(e))
            return

        if not reversal_ob_records:
            return

        # Map norm_ce -> (partner_name, ce_code_key) so we can patch existing rows
        existing_ces_map = {}
        for partner_name, ces in grouped_data.items():
            for ce_code_key in ces.keys():
                existing_ces_map[self._normalize_ce_code(
                    ce_code_key)] = (partner_name, ce_code_key)

        # Add reversal-OB-only rows for CEs not already present.
        # For existing rows, patch blank description/ce_status from OB data.
        for norm_ce, rob_data in reversal_ob_records.items():
            if norm_ce in existing_ces_map:
                p_name, ce_key = existing_ces_map[norm_ce]
                self._patch_metadata_from_ob(
                    grouped_data[p_name][ce_key], rob_data)
                continue

            # Try to find matching sale.order for live data
            so = self._find_sale_order_by_ce_code(
                rob_data.get('ce_code_display', ''))

            if so:
                partner_name = (so.partner_id.name or rob_data.get(
                    'partner_name') or 'UNKNOWN').upper()
                ce_code_display = (
                    getattr(so, 'x_studio_old_ce', '') or
                    so.x_ce_code or
                    rob_data.get('ce_code_display', 'NO_CE')
                ).upper()
                old_ce_date = getattr(so, 'x_studio_old_ce_date', False)
                ce_date = old_ce_date or so.date_order or rob_data.get(
                    'ce_date')
                description = (getattr(so, 'x_job_description', '')
                               or rob_data.get('job_description', '')).upper()

                # Get CE status from SO
                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_selection = dict(
                            so._fields['x_ce_status'].selection)
                        ce_status = status_selection.get(
                            so.x_ce_status, '').upper()
                    except Exception:
                        pass

                # Fallback to reversal OB record's ce_status if available
                if not ce_status and rob_data.get('ce_status'):
                    ce_status = rob_data.get('ce_status', '').upper()
            else:
                partner_name = (rob_data.get('partner_name')
                                or 'UNKNOWN').upper()
                ce_code_display = (rob_data.get(
                    'ce_code_display') or 'NO_CE').upper()
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

    def _patch_metadata_from_ob(self, row, ob_data):
        """
        Fill blank description/ce_status on an EXISTING grouped_data row using
        the same SO lookup + OB record fallback used when creating new OB rows.
        Called by the OB merge methods when the CE row already exists.
        """
        needs_desc = not row.get('description')
        needs_status = not row.get('ce_status')
        if not needs_desc and not needs_status:
            return
        so = self._find_sale_order_by_ce_code(
            ob_data.get('ce_code_display', ''))
        if so:
            if needs_desc:
                desc = (getattr(so, 'x_job_description', '')
                        or ob_data.get('job_description', '')).upper()
                if desc:
                    row['description'] = desc
            if needs_status:
                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_sel = dict(so._fields['x_ce_status'].selection)
                        ce_status = status_sel.get(so.x_ce_status, '').upper()
                    except Exception:
                        pass
                if not ce_status:
                    ce_status = (ob_data.get('ce_status') or '').upper()
                if ce_status:
                    row['ce_status'] = ce_status
        else:
            if needs_desc:
                desc = (ob_data.get('job_description') or '').upper()
                if desc:
                    row['description'] = desc
            if needs_status:
                cs = (ob_data.get('ce_status') or '').upper()
                if cs:
                    row['ce_status'] = cs

    def _fill_missing_metadata(self, grouped_data):
        """
        After all grouping and OB merges, any CE row that still has blank
        description or ce_status gets a final SO lookup attempt.

        This handles the case where the accrued record has no job description
        set, but the sale.order does — previously the OB merge's SO lookup
        would have provided it, but the OB merge skips rows that already
        exist in grouped_data.
        """
        for ces in grouped_data.values():
            for ce_code, ce_data in ces.items():
                if ce_code == 'NO_CE':
                    continue
                needs_desc = not ce_data.get('description')
                needs_status = not ce_data.get('ce_status')
                if not needs_desc and not needs_status:
                    continue
                try:
                    so = self._find_sale_order_by_ce_code(ce_code)
                    if not so:
                        continue
                    if needs_desc:
                        desc = getattr(so, 'x_job_description', '') or ''
                        if desc:
                            ce_data['description'] = desc.upper()
                    if needs_status and hasattr(so, 'x_ce_status') and so.x_ce_status:
                        status_sel = dict(so._fields['x_ce_status'].selection)
                        ce_status = status_sel.get(so.x_ce_status, '').upper()
                        if ce_status:
                            ce_data['ce_status'] = ce_status
                except Exception as e:
                    _logger.debug(
                        'Could not fill metadata for CE# %s: %s', ce_code, e)

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
            tuple: (
                {(PARTNER_NAME, CE_CODE): balance},          # debit - credit per CE
                {(PARTNER_NAME, CE_CODE): ce_date or None}   # first CE date seen per key
            )
        """
        prev_month_end = accrual_month - \
            relativedelta(days=1)  # last day of previous month

        # ── TIER 1: Database transaction history ──
        prev_lines = self.env['account.move.line'].sudo().search([
            ('account_id', 'in', accrued_account_ids),
            ('date', '<=', prev_month_end),
            ('parent_state', '=', 'posted'),
        ])

        balances = defaultdict(float)
        ce_date_index = {}  # (partner_name, ce_code) → first non-None CE date seen
        db_ce_codes_normalized = set()  # Track which CE codes have DB history
        cutoff_date = self._get_opening_balance_cutoff_date()
        db_ce_codes_pre_cutoff = set()  # CE codes with DB history on or before cutoff
        db_ce_codes_strictly_pre_cutoff = set()  # CE codes with DB history strictly BEFORE cutoff

        for line in prev_lines:
            partner_name = (
                line.partner_id.name.upper()
                if line.partner_id and line.partner_id.name
                else 'UNKNOWN'
            )
            ce_code = line.x_ce_code.upper() if line.x_ce_code else 'NO_CE'
            net_amount = (line.debit or 0) - (line.credit or 0)
            balances[(partner_name, ce_code)] += net_amount

            # Record the first CE date seen for this key (used for zero-collapse CE date check)
            key = (partner_name, ce_code)
            if line.x_ce_date and key not in ce_date_index:
                ce_date_index[key] = line.x_ce_date

            # Track normalized CE code so we know it has DB history
            norm = self._normalize_ce_code(ce_code)
            db_ce_codes_normalized.add(norm)

            # Track CE codes with DB history on or before the cutoff date
            if cutoff_date and line.date <= cutoff_date:
                db_ce_codes_pre_cutoff.add(norm)

            # Track CE codes with DB history strictly BEFORE the cutoff date.
            # Same-day cutoff orphan JEs (e.g. 12/31 adjustments) are intentionally
            # excluded here so that their OB balance is preserved and the JE is additive.
            if cutoff_date and line.date < cutoff_date:
                db_ce_codes_strictly_pre_cutoff.add(norm)

        # ── TIER 2: Opening Balance + Reversal OB ──
        # For the cutoff month (prev_month_end == cutoff_date):
        #   Add OB for CE#s with no pre-cutoff DB history.
        #   CE#s that only have same-day cutoff JEs keep their OB (additive).
        # For subsequent months (prev_month_end > cutoff_date):
        #   Add OB for CE#s with no pre-cutoff DB history.
        #   Also add Reversal OB — these were used in January's reversal columns
        #   but are NOT stored in the DB. Without them the balance chain breaks.
        if cutoff_date and prev_month_end >= cutoff_date:
            is_cutoff_month = (prev_month_end == cutoff_date)
            # Exclude CE#s that have DB history strictly BEFORE the cutoff date.
            # Same-day cutoff orphan JEs (dated exactly on the cutoff) are intentionally
            # kept out of this set so that their OB balance remains additive for ALL
            # months — both the cutoff month (January) and every rolling month thereafter.
            # CE#s with genuine pre-cutoff transactions are in both sets, so their OB
            # is still correctly excluded (no double-counting).
            exclude_set = db_ce_codes_strictly_pre_cutoff

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
                    partner_name = (ob_data.get('partner_name')
                                    or 'UNKNOWN').upper()
                    ce_code_display = (ob_data.get(
                        'ce_code_display') or 'NO_CE').upper()
                    balances[(partner_name, ce_code_display)
                             ] += ob_data.get('balance', 0)

                    ob_key = (partner_name, ce_code_display)
                    if ob_data.get('ce_date') and ob_key not in ce_date_index:
                        ce_date_index[ob_key] = ob_data['ce_date']

                    _logger.debug(
                        'OB for CE# %s: %.2f (prev_month_end=%s)',
                        ce_code_display, ob_data.get(
                            'balance', 0), prev_month_end
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
                        partner_name = (rev_data.get(
                            'partner_name') or 'UNKNOWN').upper()
                        ce_code_display = (rev_data.get(
                            'ce_code_display') or 'NO_CE').upper()
                        balances[(partner_name, ce_code_display)] += rev_total

                        rev_key = (partner_name, ce_code_display)
                        if rev_data.get('ce_date') and rev_key not in ce_date_index:
                            ce_date_index[rev_key] = rev_data['ce_date']

                        _logger.debug(
                            'Reversal OB for CE# %s: %.2f (prev_month_end=%s)',
                            ce_code_display, rev_total, prev_month_end
                        )

        return balances, ce_date_index

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
            accrual_month_end = (
                accrual_month + relativedelta(months=1)) - relativedelta(days=1)
            reversal_lines = self.env['account.move.line'].sudo().search([
                ('account_id', 'in', accrued_account_ids),
                ('date', '>=', accrual_month),
                ('date', '<=', accrual_month_end),
                ('parent_state', '=', 'posted'),
                ('x_type_of_entry', 'in', [
                 'reversal_system', 'reversal_manual']),
            ])
            for line in reversal_lines:
                if line.x_ce_code:
                    db_reversal_ce_codes.add(
                        self._normalize_ce_code(line.x_ce_code))

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

    def _get_cutoff_day_orphan_lines(self, accrued_account_ids, cutoff_date):
        """
        Return posted orphan JEs on the accrued revenue account dated exactly on
        the cutoff date that have a CE code but no x_type_of_entry.
        These are manual adjustment JEs (e.g. 12/31 reversals) whose net should
        be added on top of the Opening Balance in Column G of the OB month.
        """
        return self.env['account.move.line'].sudo().search([
            ('account_id', 'in', accrued_account_ids),
            ('date', '=', cutoff_date),
            ('parent_state', '=', 'posted'),
            ('x_type_of_entry', 'in', [False, '']),
            ('x_ce_code', '!=', False),
        ])

    def _find_grouped_data_match_priority(self, ce_code, ce_date, grouped_data):
        """
        Find an existing row in grouped_data for the given CE code using a
        4-step priority to minimise false-positive merges:

          1. Exact match          (strip + uppercase, case-insensitive)
          2. Spaces-only removal  (remove all spaces, case-insensitive)
          3. Full normalisation   (_normalize_ce_code: all whitespace + uppercase)
          4. Zero-collapse        (collapses 00+ runs to 0, last resort)

        CE Date tie-breaking applies at every step: if both the JE line and the
        candidate row have a CE Date and they differ, the candidate is skipped.
        If either side has no CE Date the tie-break is not applied (allow match).

        Returns (partner_name, ce_key) for the first valid match, or None.
        """
        if not ce_code:
            return None

        ce_stripped = ce_code.strip().upper()
        ce_no_spaces = ce_code.replace(' ', '').upper()
        ce_normalized = self._normalize_ce_code(ce_code)
        ce_collapsed = self._normalize_ce_zero_collapse(ce_code)

        def dates_ok(row_data):
            """Return True if CE dates are compatible (no conflict)."""
            row_date = row_data.get('ce_date')
            if not ce_date or not row_date:
                return True  # either blank → no conflict
            return ce_date == row_date

        # Step 1: exact (strip + uppercase)
        for p_name, ces in grouped_data.items():
            for existing_ce, row_data in ces.items():
                if existing_ce.strip().upper() == ce_stripped and dates_ok(row_data):
                    return (p_name, existing_ce)

        # Step 2: spaces-only removal (preserves other characters, case-insensitive)
        for p_name, ces in grouped_data.items():
            for existing_ce, row_data in ces.items():
                if existing_ce.replace(' ', '').upper() == ce_no_spaces and dates_ok(row_data):
                    return (p_name, existing_ce)

        # Step 3: full normalisation (all whitespace removed + uppercase)
        for p_name, ces in grouped_data.items():
            for existing_ce, row_data in ces.items():
                if self._normalize_ce_code(existing_ce) == ce_normalized and dates_ok(row_data):
                    return (p_name, existing_ce)

        # Step 4: zero-collapse normalisation (last resort — one extra zero in a zero run)
        # e.g. SSY0001 matches SSY00001, B450009 matches B4500009
        for p_name, ces in grouped_data.items():
            for existing_ce, row_data in ces.items():
                if self._normalize_ce_zero_collapse(existing_ce) == ce_collapsed and dates_ok(row_data):
                    return (p_name, existing_ce)

        return None

    def _merge_cutoff_day_unmatched_rows(self, grouped_data, cutoff_day_lines):
        """
        For each orphan JE line dated on the cutoff date, try to find a matching
        row in grouped_data using priority normalisation.  Lines that match an
        existing row are skipped (their net amount flows into Column G via
        _calculate_prev_month_balances Tier 1 + OB additive fix).

        Lines whose CE code cannot be matched to ANY existing row get their own
        new row flagged with `is_unmatched_cutoff_adj: True`.  The row will show:
          - CE# in red
          - CLIENT populated from the JE line
          - All descriptive columns (description, date, status) left blank
          - Column G = JE net amount (picked up from prev_month_balances)
          - Column M = G (H–L are all zero)
        """
        seen_unmatched = set()  # avoid duplicate rows for same (CE code, CE date) pair

        for line in cutoff_day_lines:
            ce_code = line.x_ce_code.strip().upper() if line.x_ce_code else 'NO_CE'
            ce_date = line.x_ce_date or None
            partner_name = (
                line.partner_id.name.upper()
                if line.partner_id and line.partner_id.name
                else 'UNKNOWN'
            )

            match = self._find_grouped_data_match_priority(ce_code, ce_date, grouped_data)
            if match:
                continue  # already has a row; OB+JE balance handled by Tier 1/2 fix

            dedup_key = (ce_code, ce_date)
            if dedup_key in seen_unmatched:
                continue
            seen_unmatched.add(dedup_key)

            grouped_data[partner_name][ce_code] = {
                'ce_date': ce_date,
                'year': ce_date.year if ce_date else None,
                'description': '',
                'ce_status': '',
                'so_reference': '',
                'lines': [],
                'is_unmatched_cutoff_adj': True,
            }

            _logger.debug(
                'Added unmatched cutoff-day adjustment row for CE# %s (partner: %s)',
                ce_code, partner_name
            )

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

            elif not line.x_type_of_entry:
                # Orphan JE: no entry type linkage, identified by 'manual adjustment' in label
                if line.x_ce_code:
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
        prev_month_balances, prev_month_ce_dates = self._calculate_prev_month_balances(
            [accrued_account_id], accrual_month)

        # Calculate reversal opening balances (only populated for OB month)
        reversal_ob_balances = self._calculate_reversal_opening_balances(
            accrual_month)

        # Group lines by client and CE# for this specific month.
        # Orphan JE lines are folded into matching regular-accrual CE rows
        # (zero-padding variants like BA00003 → BA0003) inside _group_lines_by_ce.
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

                # For the OB month only: find orphan JEs posted on the cutoff date
                # (e.g. 12/31 manual adjustment JEs) and add rows for any CE codes
                # that have no matching row yet.  Their balance flows into Column G
                # via the Tier 1 + strictly-pre-cutoff exclusion fix.
                if prev_month_end == cutoff_date:
                    cutoff_orphan_lines = self._get_cutoff_day_orphan_lines(
                        [accrued_account_id], cutoff_date)
                    if cutoff_orphan_lines:
                        self._merge_cutoff_day_unmatched_rows(
                            grouped_data, cutoff_orphan_lines)

        # Final pass: fill blank description/ce_status from SO for any remaining gaps
        self._fill_missing_metadata(grouped_data)

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

        # Assign each previous-month balance to exactly one row before writing,
        # so a CE spanning two partners cannot have its balance counted twice.
        prev_balance_by_row = self._assign_prev_balances(
            grouped_data, prev_month_balances, prev_month_ce_dates)

        # Write data rows
        for partner_name in sorted(grouped_data.keys()):
            ces = grouped_data[partner_name]

            for ce_code in sorted(ces.keys()):
                ce_data = ces[ce_code]
                amounts = self._calculate_amounts_by_type(
                    ce_data['lines'], accrual_month)

                # ── Compute all monetary values before writing ──

                # Previous month ending balance (column G).
                # Matching (standard normalisation, then zero-collapse with a CE
                # Date tie-break) happens in _assign_prev_balances, which also
                # guarantees each balance is claimed by exactly one row.
                prev_balance = prev_balance_by_row.get(
                    (partner_name, ce_code), 0.0) or 0

                # Reversal opening balance overrides (OB month only)
                norm_ce_for_rev = self._normalize_ce_code(ce_code)
                rev_ob = reversal_ob_balances.get(norm_ce_for_rev, {})

                # Column H (7): System Reversal
                system_reversal_val = amounts['system_reversal']
                if system_reversal_val == 0 and rev_ob.get('system_reversal', 0) != 0:
                    system_reversal_val = rev_ob['system_reversal']

                # Column J (9): Manual Reversal (includes adjustment from reversal OB)
                manual_reversal_val = amounts['manual_reversal']
                if manual_reversal_val == 0 and rev_ob.get('manual_reversal', 0) != 0:
                    manual_reversal_val = rev_ob['manual_reversal']
                manual_reversal_val -= rev_ob.get(
                    'manual_reversal_adjustment', 0)

                # Skip row if ALL monetary columns (G through M) are zero
                all_monetary = [
                    prev_balance,
                    system_reversal_val,
                    amounts['system_accrual'],
                    manual_reversal_val,
                    amounts['manual_reaccrual'],
                    amounts['manual_adjustment'],
                ]
                if all(v == 0 for v in all_monetary):
                    continue

                # Check if this is an OB-only row (no accrued lines)
                is_ob_only_row = not ce_data['lines']

                # If OB-only, verify it also doesn't exist in sale.order
                if is_ob_only_row:
                    so_match = self._find_sale_order_by_ce_code(ce_code)
                    is_ob_only_row = not so_match

                # Unmatched cutoff-day adjustment rows get a red CE# to signal they
                # come from a 12/31 orphan JE with no corresponding accrual row.
                is_unmatched_cutoff_adj = ce_data.get('is_unmatched_cutoff_adj', False)

                # ── Write the row ──

                sheet.write(row, 0, partner_name, formats['normal'])
                sheet.write(
                    row, 1, ce_data['so_reference'], formats['centered'])
                ce_code_format = formats['centered_red'] if is_unmatched_cutoff_adj else formats['centered']
                sheet.write(row, 2, ce_code, ce_code_format)

                if ce_data['ce_date']:
                    sheet.write(row, 3, ce_data['ce_date'], formats['date'])
                else:
                    sheet.write(row, 3, '', formats['centered'])

                sheet.write(row, 4, ce_data['description'], formats['normal'])

                if ce_data['year']:
                    sheet.write(row, 5, ce_data['year'], formats['centered'])
                else:
                    sheet.write(row, 5, '', formats['centered'])

                sheet.write(row, 6, prev_balance, formats['currency'])
                sheet.write(
                    row, 7, system_reversal_val, formats['currency_negative'])
                sheet.write(
                    row, 8, amounts['system_accrual'], formats['currency_negative'])
                sheet.write(
                    row, 9, manual_reversal_val, formats['currency_negative'])
                sheet.write(
                    row, 10, amounts['manual_reaccrual'], formats['currency_negative'])
                sheet.write(
                    row, 11, amounts['manual_adjustment'], formats['currency_negative'])

                excel_row = row + 1
                sheet.write_formula(
                    row, 12, f'=G{excel_row}+H{excel_row}+I{excel_row}+J{excel_row}+K{excel_row}+L{excel_row}', formats['currency'])

                sheet.write(row, 13, self._normalize_ce_status(
                    ce_data['ce_status']), formats['centered'])

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
        """Generate the GL sheet: a 1:1 replication of Odoo's General Ledger for
        this account and month.

        Queries account.move.line directly instead of filtering the ``lines``
        recordset handed in.  That recordset is date-bounded by the wizard to
        ``start_month - 1 month`` (accrued_revenue_wizard.py), so filtering it
        made the opening balance cover only the single preceding month rather
        than all prior history -- the same month produced different opening
        balances depending on the report's start date.

        Deliberately does NOT consult saatchi.accrued_revenue_opening_balance:
        Odoo's General Ledger knows nothing about that model, so including it
        would break the 1:1 tie-out.  The month sheets do apply it (via
        _calculate_prev_month_balances), which is why the two can legitimately
        differ for CE#s that exist only as an opening balance.

        ``lines`` is retained for signature stability and is intentionally
        unused for the ledger figures -- do not "optimise" back to filtering it.
        """
        accrual_month_end = (
            accrual_month + relativedelta(months=1)) - relativedelta(days=1)

        AccountMoveLine = self.env['account.move.line'].sudo()
        # Posted-only mirrors Odoo's GL default ("Include unposted entries" off).
        base_domain = [
            ('account_id', '=', accrued_account_id),
            ('parent_state', '=', 'posted'),
        ]

        # Movement rows for the month, straight from the ledger.
        gl_lines = AccountMoveLine.search(base_domain + [
            ('date', '>=', accrual_month),
            ('date', '<=', accrual_month_end),
        ])

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
        accrued_account = self.env['account.account'].browse(
            accrued_account_id)

        # Opening balance = cumulative debit - credit over ALL prior history,
        # matching Odoo's "Initial Balance" row.  Queried from the ledger, not
        # from the date-bounded recordset passed in (see docstring).
        prev_month_end = accrual_month - relativedelta(days=1)
        prev_balance_lines = AccountMoveLine.search(
            base_domain + [('date', '<=', prev_month_end)])

        opening_balance = sum(prev_balance_lines.mapped(
            lambda l: l.debit - l.credit))

        # Write opening balance row
        # Date: first day of month
        sheet.write(row, 0, accrual_month, formats['date'])
        sheet.write(row, 1, 'Opening Balance', formats['normal'])  # Entry Type
        sheet.write(row, 2, '', formats['normal'])  # Journal Entry (blank)
        sheet.write(row, 3, accrued_account.display_name if accrued_account else '',
                    formats['normal'])  # Account
        sheet.write(row, 4, '', formats['normal'])  # Client Name (blank)
        sheet.write(row, 5, '', formats['centered'])  # CE Code (blank)
        sheet.write(row, 6, '', formats['centered'])  # CE Date (blank)
        sheet.write(row, 7, 'Opening Balance', formats['normal'])  # Label
        sheet.write(row, 8, '', formats['normal'])  # Reference (blank)
        sheet.write(row, 9, 0, formats['currency'])  # Debit (0)
        sheet.write(row, 10, 0, formats['currency'])  # Credit (0)
        # DR Less CR (opening balance)
        sheet.write(row, 11, opening_balance, formats['currency'])
        sheet.write(row, 12, '', formats['normal'])  # Remarks (blank)
        row += 1

        # Match Odoo's General Ledger ordering: date, journal entry name, id.
        sorted_lines = gl_lines.sorted(key=lambda l: (
            l.date or datetime.date.min,
            (l.move_id.name or '') if l.move_id else '',
            l.id))

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
            entry_type = entry_type_map.get(line.x_type_of_entry, '')
            if not entry_type and not line.x_type_of_entry:
                if line.x_ce_code:
                    entry_type = 'Manual Adjustment'
            sheet.write(row, 1, entry_type, formats['normal'])

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
