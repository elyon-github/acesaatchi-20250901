from odoo import models
import datetime
import re
from xlsxwriter.workbook import Workbook
from odoo.exceptions import ValidationError, UserError
from dateutil.relativedelta import relativedelta
import logging
from collections import defaultdict

_logger = logging.getLogger(__name__)


class SalesOrderRevenueXLSX(models.AbstractModel):
    _name = 'report.sales_order_revenue_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    _description = 'Sales Order Revenue XLSX Report'

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

        return {
            'title': title_format,
            'section_header': section_header_format,
            'section_header_no_border': section_header_no_border,
            'column_header': column_header_format,
            'normal': normal_format,
            'centered': centered_format,
            'date': date_format,
            'currency': currency_format,
            'currency_negative': currency_negative_format
        }

    def _normalize_ce_code(self, ce_code):
        """Normalize a CE code for consistent matching.
        Removes all whitespace and converts to uppercase."""
        if not ce_code:
            return ''
        return re.sub(r'\s+', '', ce_code.strip().upper())

    def _get_opening_balance_cutoff_date(self):
        """Get the opening balance cutoff date from the accrual configuration."""
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

    def _find_sale_order_by_ce_code(self, ce_code):
        """Find a sale.order matching the given CE code."""
        if not ce_code:
            return self.env['sale.order']

        SaleOrder = self.env['sale.order'].sudo()
        company_domain = [('company_id', 'in', self.env.companies.ids)]

        so = SaleOrder.search(
            [('x_ce_code', '=', ce_code)] + company_domain, limit=1)
        if so:
            return so

        try:
            so = SaleOrder.search(
                [('x_studio_old_ce', '=', ce_code)] + company_domain, limit=1)
            if so:
                return so
        except Exception:
            pass

        so = SaleOrder.search(
            [('x_ce_code', 'ilike', ce_code.strip())] + company_domain, limit=1)
        if so:
            return so

        try:
            so = SaleOrder.search(
                [('x_studio_old_ce', 'ilike', ce_code.strip())] + company_domain, limit=1)
            if so:
                return so
        except Exception:
            pass

        norm_ce = self._normalize_ce_code(ce_code)
        all_sos = SaleOrder.search(company_domain)
        for so in all_sos:
            if so.x_ce_code and self._normalize_ce_code(so.x_ce_code) == norm_ce:
                return so
            old_ce = getattr(so, 'x_studio_old_ce', '')
            if old_ce and self._normalize_ce_code(old_ce) == norm_ce:
                return so

        return SaleOrder

    def _calculate_reversal_opening_balances(self, report_month):
        """Get reversal opening balances for the opening balance month.

        Uses the same 2-tier approach as the original accrued revenue report:
        1. PRIMARY: Check for DB reversal entries. If a CE# has DB reversal
           transactions, its computed amounts take priority.
        2. FALLBACK: For CE#s with NO DB reversal history, use the
           reversal_opening_balance model.
        """
        cutoff_date = self._get_opening_balance_cutoff_date()
        if not cutoff_date:
            return {}

        # The accrual month is the report month itself
        accrual_month = report_month.replace(day=1)
        prev_month_end = accrual_month - relativedelta(days=1)
        if prev_month_end != cutoff_date:
            return {}

        # Identify CE codes that already have DB reversal history
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

        return result

    def _merge_reversal_opening_balance_rows(self, grouped_data):
        """Merge reversal-opening-balance-only rows into grouped_data.

        For CE codes that exist in the reversal opening balance model but have NO
        accrual records, create a row from matching sale.order or reversal OB fields.
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

        # Collect normalized CE codes already present.
        # Also include normalized x_studio_old_ce from matching sale orders so
        # that OB rows keyed by the old CE code (e.g. "BLF 00004") are
        # recognised as duplicates of existing SO rows (e.g. BLFSO000211
        # whose x_studio_old_ce == "BLF 00004").
        existing_normalized_ces = set()
        for partner_name, ces in grouped_data.items():
            for ce_code_key in ces.keys():
                existing_normalized_ces.add(
                    self._normalize_ce_code(ce_code_key))
                # Look up SO by CE code and add its x_studio_old_ce
                so = self._find_sale_order_by_ce_code(ce_code_key)
                if so:
                    old_ce = getattr(so, 'x_studio_old_ce', '') or ''
                    if old_ce:
                        existing_normalized_ces.add(
                            self._normalize_ce_code(old_ce))

        # Add reversal-OB-only rows for CEs not already present
        for norm_ce, rob_data in reversal_ob_records.items():
            if norm_ce in existing_normalized_ces:
                continue

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

                ce_status = ''
                if hasattr(so, 'x_ce_status') and so.x_ce_status:
                    try:
                        status_selection = dict(
                            so._fields['x_ce_status'].selection)
                        ce_status = status_selection.get(
                            so.x_ce_status, '').upper()
                    except Exception:
                        pass
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

            grouped_data[partner_name][ce_code_display] = {
                'ce_date': ce_date,
                'description': description,
                'year': ce_date.year if ce_date else None,
                'month': ce_date.strftime('%B').upper() if ce_date else None,
                'ce_status': ce_status,
                'so_reference': '',
                'lines': [],
                'sales_orders': set(),
            }

            _logger.debug(
                'Added reversal-OB-only row for CE# %s (partner: %s)',
                ce_code_display, partner_name
            )

    def _fill_missing_ce_metadata(self, grouped_data):
        """Fill in missing CE metadata (ce_status, description, ce_date) from
        sale orders and reversal opening balance records."""
        # Get reversal OB records for fallback
        cutoff_date = self._get_opening_balance_cutoff_date()
        reversal_ob_records = {}
        if cutoff_date:
            try:
                reversal_ob_records = self.env[
                    'saatchi.accrued_revenue_reversal_opening_balance'
                ].get_reversal_opening_balances_for_month(
                    balance_date=cutoff_date,
                    company_id=self.env.company.id
                )
            except Exception:
                pass

        for partner_name, ces in grouped_data.items():
            for ce_code, ce_data in ces.items():
                needs_status = not ce_data.get('ce_status')
                needs_description = not ce_data.get('description')
                needs_date = not ce_data.get('ce_date')

                if not (needs_status or needs_description or needs_date):
                    continue

                # Try sale order first
                so = self._find_sale_order_by_ce_code(ce_code)
                if so:
                    if needs_status and hasattr(so, 'x_ce_status') and so.x_ce_status:
                        try:
                            status_selection = dict(
                                so._fields['x_ce_status'].selection)
                            ce_data['ce_status'] = status_selection.get(
                                so.x_ce_status, '').upper()
                            needs_status = not ce_data['ce_status']
                        except Exception:
                            pass

                    if needs_description and hasattr(so, 'x_job_description') and so.x_job_description:
                        ce_data['description'] = so.x_job_description.upper()
                        needs_description = False

                    if needs_date and so.date_order:
                        ce_data['ce_date'] = so.date_order
                        ce_data['year'] = so.date_order.year
                        ce_data['month'] = so.date_order.strftime('%B').upper()
                        needs_date = False

                # Try reversal OB records as second fallback
                if needs_status or needs_description or needs_date:
                    norm_ce = self._normalize_ce_code(ce_code)
                    rob_data = reversal_ob_records.get(norm_ce, {})

                    if needs_status and rob_data.get('ce_status'):
                        ce_data['ce_status'] = rob_data['ce_status'].upper()

                    if needs_description and rob_data.get('job_description'):
                        ce_data['description'] = rob_data['job_description'].upper()

                    if needs_date and rob_data.get('ce_date'):
                        ce_data['ce_date'] = rob_data['ce_date']
                        ce_data['year'] = rob_data['ce_date'].year
                        ce_data['month'] = rob_data['ce_date'].strftime(
                            '%B').upper()

    def _sanitize_sheet_name(self, name):
        """Sanitize sheet name to comply with Excel rules

        Excel sheet name rules:
        - Max 31 characters
        - Cannot contain: \ / ? * [ ]
        - Cannot be empty
        """
        if not name:
            return 'UNNAMED'

        # Remove invalid characters
        invalid_chars = ['\\', '/', '?', '*', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '')

        # Truncate to 31 characters
        if len(name) > 31:
            name = name[:31]

        return name if name else 'UNNAMED'

    def _group_lines_by_ce(self, lines, report_month, all_billed_so_ids=None):
        """Group account.move.line records by partner and CE code

        Includes:
        - Lines from accrued revenue entries
        - ALL sales orders that were billed in the report month (whether they have accrued entries or not)
        """
        grouped = defaultdict(lambda: defaultdict(lambda: {
            'ce_date': None,
            'description': '',
            'year': None,
            'month': None,
            'ce_status': '',
            'so_reference': '',
            'lines': [],
            'sales_orders': set()
        }))

        # Calculate date ranges
        # Reversal month: current report month
        reversal_start = report_month.replace(day=1)
        reversal_end = (reversal_start + relativedelta(months=1)
                        ) - relativedelta(days=1)

        # Accrual month: 1 month before report month
        accrual_month = report_month - relativedelta(months=1)
        accrual_start = accrual_month.replace(day=1)
        accrual_end = (accrual_start + relativedelta(months=1)
                       ) - relativedelta(days=1)

        # Process accrued revenue lines
        relevant_lines = []
        for line in lines:
            if not line.date:
                continue

            # Include reversals from current month or accruals from last month
            if (reversal_start <= line.date <= reversal_end) or (accrual_start <= line.date <= accrual_end):
                relevant_lines.append(line)

        for line in relevant_lines:
            partner_name = line.partner_id.name.upper(
            ) if line.partner_id and line.partner_id.name else 'UNKNOWN'
            ce_code = line.x_ce_code.upper() if line.x_ce_code else 'NO_CE'

            # Store line for processing
            grouped[partner_name][ce_code]['lines'].append(line)

            # Track sales order for billing calculation
            if line.x_sales_order:
                grouped[partner_name][ce_code]['sales_orders'].add(
                    line.x_sales_order.id)

            # Capture CE-level fields (use first non-empty value found)
            if line.x_ce_date and not grouped[partner_name][ce_code]['ce_date']:
                grouped[partner_name][ce_code]['ce_date'] = line.x_ce_date
                grouped[partner_name][ce_code]['year'] = line.x_ce_date.year
                grouped[partner_name][ce_code]['month'] = line.x_ce_date.strftime(
                    '%B').upper()

            if line.move_id and line.move_id.x_related_custom_accrued_record and not grouped[partner_name][ce_code]['description']:
                desc = line.move_id.x_related_custom_accrued_record.ce_job_description or ''
                grouped[partner_name][ce_code]['description'] = desc.upper()

            if line.x_ce_status and not grouped[partner_name][ce_code]['ce_status']:
                selection_dict = dict(line._fields['x_ce_status'].selection)
                ce_status = selection_dict.get(line.x_ce_status, '')
                grouped[partner_name][ce_code]['ce_status'] = ce_status.upper(
                ) if ce_status else ''

            if line.move_id and line.move_id.x_related_custom_accrued_record and not grouped[partner_name][ce_code]['so_reference']:
                accrued_record = line.move_id.x_related_custom_accrued_record
                so_ref = ''
                if accrued_record.x_related_ce_id:
                    so_ref = accrued_record.x_related_ce_id.name or ''
                grouped[partner_name][ce_code]['so_reference'] = so_ref.upper(
                ) if so_ref else ''

        # Process ALL billed sales orders (including those already in grouped data)
        if all_billed_so_ids:
            all_billed_sos = self.env['sale.order'].browse(all_billed_so_ids)

            for so in all_billed_sos:
                partner_name = so.partner_id.name.upper(
                ) if so.partner_id and so.partner_id.name else 'UNKNOWN'
                ce_code = so.x_ce_code.upper() if so.x_ce_code else 'NO_CE'

                # If the SO has x_studio_old_ce that already exists as a key
                # in grouped_data for this partner, merge into that row instead
                # of creating a duplicate (e.g. SO "BLFSO000211" with
                # x_studio_old_ce "BLF 00004" should merge into "BLF 00004").
                old_ce = (getattr(so, 'x_studio_old_ce', '')
                          or '').upper().strip()
                if old_ce and partner_name in grouped and old_ce in grouped[partner_name]:
                    ce_code = old_ce

                # Add SO to sales_orders set (will merge with existing if already present)
                grouped[partner_name][ce_code]['sales_orders'].add(so.id)

                # Populate CE-level fields from SO if not already set
                if so.date_order and not grouped[partner_name][ce_code]['ce_date']:
                    grouped[partner_name][ce_code]['ce_date'] = so.date_order
                    grouped[partner_name][ce_code]['year'] = so.date_order.year
                    grouped[partner_name][ce_code]['month'] = so.date_order.strftime(
                        '%B').upper()

                # Get description from SO if available and not already set
                if hasattr(so, 'x_job_description') and so.x_job_description and not grouped[partner_name][ce_code]['description']:
                    grouped[partner_name][ce_code]['description'] = so.x_job_description.upper(
                    )

                if so.x_ce_status and not grouped[partner_name][ce_code]['ce_status']:
                    selection_dict = dict(so._fields['x_ce_status'].selection)
                    ce_status = selection_dict.get(so.x_ce_status, '')
                    grouped[partner_name][ce_code]['ce_status'] = ce_status.upper(
                    ) if ce_status else ''

        return grouped

    def _calculate_amounts_by_type(self, lines, report_month):
        """Calculate amounts for each entry type category

        Logic:
        - Reversals: from current report month (e.g., if report_month is July, reversals from July)
        - Accruals: from 1 month ago (e.g., if report_month is July, accruals from June)
        """
        amounts = {
            'system_accrual': 0,
            'system_reversal': 0,
            'manual_accrual': 0,
            'manual_reversal': 0,
            'addl_adj': 0
        }

        # Calculate date range: all entries use the report month
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - \
            relativedelta(days=1)

        for line in lines:
            if not line.date:
                continue

            # Only include lines within the report month
            if not (month_start <= line.date <= month_end):
                continue

            # Calculate net amount (debit - credit)
            net_amount = (line.debit or 0) - (line.credit or 0)

            # Categorize based on type of entry
            if line.x_type_of_entry == 'reversal_system':
                amounts['system_reversal'] += net_amount
            elif line.x_type_of_entry == 'accrued_system':
                amounts['system_accrual'] += net_amount
            elif line.x_type_of_entry == 'reversal_manual':
                amounts['manual_reversal'] += net_amount
            elif line.x_type_of_entry == 'accrued_manual':
                amounts['manual_accrual'] += net_amount
            elif line.x_type_of_entry in ('adjustment_system', 'adjustment_manual'):
                # Adjustments & their reversals both appear here
                # (they land in different months by design, so no netting)
                amounts['addl_adj'] += net_amount

        return amounts

    def _get_excluded_billed_account_ids(self):
        """Get the set of account IDs to exclude from billed amount calculation.

        Reads from the saatchi.accrual_config Many2many field
        'excluded_billed_account_ids' for the current company.
        """
        try:
            config = self.env['saatchi.accrual_config'].sudo().search([
                ('company_id', '=', self.env.company.id)
            ], limit=1)
            if config and config.excluded_billed_account_ids:
                return set(config.excluded_billed_account_ids.ids)
        except Exception as e:
            _logger.warning(
                'Could not retrieve excluded billed accounts: %s', str(e))
        return set()

    def _get_amount_untaxed_in_company_currency(self, move):
        """Get the untaxed invoice amount in company currency (PHP).

        When the invoice is in a foreign currency (e.g. USD), the stored
        ``amount_untaxed`` is in that foreign currency.  Instead of
        converting via exchange rates, this method reads the PHP amounts
        directly from the journal item balances which are *always* stored
        in company currency.

        Only product / rounding lines (excluding tax-repartition rounding)
        are summed — the same set Odoo uses internally.

        Args:
            move: account.move record (invoice / credit memo)

        Returns:
            float: absolute untaxed amount in company currency
        """
        company_currency = move.company_id.currency_id or self.env.company.currency_id
        if move.currency_id == company_currency:
            return move.amount_untaxed

        # Sum the balance (debit − credit) of product / rounding lines.
        # This mirrors Odoo's own _compute_amount logic for amount_untaxed_signed.
        total_balance = 0.0
        for line in move.line_ids:
            if line.display_type in ('product', 'rounding') \
                    and not line.tax_repartition_line_id:
                total_balance += (line.debit or 0.0) - (line.credit or 0.0)

        # Return the absolute value; callers handle the sign
        # (positive for invoices, negated for credit memos).
        return abs(total_balance)

    def _get_cm_deduction_for_invoice(self, invoice):
        """Get the total credit memo deduction applied against an invoice.
    
        Returns the untaxed portion of the CM in company currency, to match
        how the invoice billed amount is calculated (untaxed only).
        """
        cm_total = 0.0
        company_currency = invoice.company_id.currency_id or self.env.company.currency_id
    
        receivable_lines = invoice.line_ids.filtered(
            lambda l: l.account_id.account_type == 'asset_receivable'
        )
    
        processed_cm_ids = set()
    
        for line in receivable_lines:
            for partial in line.matched_credit_ids:
                counterpart = partial.credit_move_id
                cm_move = counterpart.move_id
                if cm_move.move_type == 'out_refund' and cm_move.id not in processed_cm_ids:
                    processed_cm_ids.add(cm_move.id)
                    cm_total += self._get_amount_untaxed_in_company_currency(cm_move)
    
            for partial in line.matched_debit_ids:
                counterpart = partial.debit_move_id
                cm_move = counterpart.move_id
                if cm_move.move_type == 'out_refund' and cm_move.id not in processed_cm_ids:
                    processed_cm_ids.add(cm_move.id)
                    cm_total += self._get_amount_untaxed_in_company_currency(cm_move)
    
        return cm_total

    def _calculate_billed_amount(self, sales_order_ids, report_month):
        """Calculate total billed amount for given sales orders in the report month.

        Logic:
            1. Find all posted Customer Invoices (out_invoice) with
               x_studio_invoice_type == 'CE Related' in the report month.
            2. For each invoice, get the untaxed amount in company currency.
            3. Check if any credit memos (out_refund) have been reconciled
               against the invoice as payment — if so, subtract that amount.
            4. Subtract excluded-account journal item amounts.

        Additionally, this method searches for "orphan" invoices — moves NOT
        linked to any sales order but whose x_studio_old_ce_1 field (on
        account.move) matches the x_studio_old_ce_ field on one of the given
        sales orders (normalized comparison).

        Args:
            sales_order_ids: set of sale.order IDs
            report_month: date object representing the report month

        Returns:
            float: Adjusted billed amount for posted invoices in the report month
        """
        if not sales_order_ids:
            return 0.0

        # Get the start and end of the report month
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - \
            relativedelta(days=1)

        # Get excluded account IDs from accrual configuration
        excluded_account_ids = self._get_excluded_billed_account_ids()

        total_billed = 0.0
        total_exclusion = 0.0
        processed_invoice_ids = set()

        def _apply_invoice(inv):
            """Calculate billed and exclusion for a single invoice.

            Billed = untaxed amount in company currency
                     minus any credit memos reconciled as payment
            Exclusion = net amount on excluded accounts

            Returns (billed_delta, exclusion_delta).
            """
            billed_delta = self._get_amount_untaxed_in_company_currency(inv)

            # Subtract credit memos that were applied as payment
            cm_deduction = self._get_cm_deduction_for_invoice(inv)
            if cm_deduction:
                _logger.info(
                    'Invoice %s: subtracting CM deduction %.2f from billed %.2f',
                    inv.name, cm_deduction, billed_delta,
                )
                billed_delta -= cm_deduction

            exclusion_delta = 0.0
            if excluded_account_ids:
                for line in inv.line_ids:
                    if line.account_id and line.account_id.id in excluded_account_ids:
                        exclusion_delta += (line.credit or 0.0)
                        exclusion_delta -= (line.debit or 0.0)
            return billed_delta, exclusion_delta

        # Browse sales orders
        sales_orders = self.env['sale.order'].browse(list(sales_order_ids))

        # ------------------------------------------------------------------
        # STEP 1: Process SO-linked invoices (out_invoice only)
        # ------------------------------------------------------------------
        old_ce_codes = set()  # collect old CE codes for orphan lookup later

        for so in sales_orders:
            # Collect the SO's old CE code for orphan invoice matching
            so_old_ce = getattr(so, 'x_studio_old_ce_', None) or getattr(
                so, 'x_studio_old_ce', '') or ''
            if so_old_ce:
                old_ce_codes.add(self._normalize_ce_code(so_old_ce))

            # Process SO-linked invoices (out_invoice only)
            moves = so.invoice_ids.filtered(
                lambda inv: inv.state == 'posted'
                and inv.move_type == 'out_invoice'
                and inv.invoice_date
                and month_start <= inv.invoice_date <= month_end
                and getattr(inv, 'x_studio_invoice_type', '') == 'CE Related'
            )
            for move in moves:
                processed_invoice_ids.add(move.id)
                b_delta, e_delta = _apply_invoice(move)
                total_billed += b_delta
                total_exclusion += e_delta

        # ------------------------------------------------------------------
        # STEP 2: Find orphan invoices (not linked to ANY SO) whose
        #         x_studio_old_ce_1 matches one of the SO old CE codes.
        #         Only out_invoice — credit memos are handled via
        #         reconciliation against the invoices above.
        # ------------------------------------------------------------------

        if old_ce_codes:

            orphan_domain = [
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
                ('invoice_date', '>=', month_start),
                ('invoice_date', '<=', month_end),
                ('company_id', 'in', self.env.companies.ids),
                ('x_studio_old_ce_1', '!=', False),
            ]

            try:
                candidate_moves = self.env['account.move'].sudo().search(
                    orphan_domain)
            except Exception:
                candidate_moves = self.env['account.move']

            for inv in candidate_moves:
                # Skip already processed (SO-linked) invoices
                if inv.id in processed_invoice_ids:
                    continue

                # Must be CE Related
                if getattr(inv, 'x_studio_invoice_type', '') != 'CE Related':
                    continue

                # Check if invoice's x_studio_old_ce_1 matches any SO old CE
                inv_old_ce = getattr(inv, 'x_studio_old_ce_1', '') or ''
                if not inv_old_ce:
                    continue
                if self._normalize_ce_code(inv_old_ce) not in old_ce_codes:
                    continue

                # Verify this invoice is truly orphan (not linked to ANY SO)
                has_sale_link = any(
                    aml.sale_line_ids for aml in inv.line_ids
                )
                if has_sale_link:
                    continue

                # Include this orphan invoice
                processed_invoice_ids.add(inv.id)
                b_delta, e_delta = _apply_invoice(inv)
                total_billed += b_delta
                total_exclusion += e_delta
                if inv.name:
                    _logger.info(
                        'Included orphan invoice %s (old CE: %s) in billed calculation '
                        'with billed_delta %.2f and exclusion_delta %.2f. Total billed so far: %.2f',
                        inv.name, inv.x_studio_old_ce_1,
                        b_delta, e_delta, total_billed,
                    )

        return total_billed - total_exclusion

    def generate_xlsx_report(self, workbook, data, docids):
        """
        Main report generation method.

        Args:
            workbook: xlsxwriter workbook object
            data: dictionary containing report_date, partner_ids, and move_line_ids from wizard
            docids: not used (we use move_line_ids from data instead)
        """
        # Extract report date from wizard data
        if data and 'report_date' in data:
            report_date_str = data['report_date']
            report_month = datetime.datetime.strptime(
                report_date_str, '%Y-%m-%d').date()
        else:
            # Fallback to today if no date provided
            report_month = datetime.date.today()

        # Get move lines from data
        if data and 'move_line_ids' in data:
            move_line_ids = data['move_line_ids']
            move_lines = self.env['account.move.line'].browse(move_line_ids)
        else:
            raise UserError("No move line IDs provided in report data.")

        # Validate move_lines
        if not move_lines:
            raise UserError(
                "No accrued revenue entries found for the selected criteria.")

        # Define formats
        formats = self._define_formats(workbook)

        # Get all billed SO IDs from data
        all_billed_so_ids = data.get('all_billed_so_ids', []) if data else []

        # Group lines by partner and CE (includes both accrued and billed)
        grouped_data = self._group_lines_by_ce(
            move_lines, report_month, all_billed_so_ids)

        # Merge reversal opening balance rows (for CEs only in OB, no journal entries)
        self._merge_reversal_opening_balance_rows(grouped_data)

        # Fill in missing CE metadata (ce_status, description, ce_date) from SO / reversal OB
        self._fill_missing_ce_metadata(grouped_data)

        if not grouped_data:
            raise UserError("No data found for the report month.")

        # Calculate reversal opening balances for fallback
        reversal_ob_balances = self._calculate_reversal_opening_balances(
            report_month)

        # Generate summary sheet first
        self._generate_summary_sheet(
            workbook, formats, grouped_data, report_month, reversal_ob_balances)

        # Generate individual customer sheets
        for partner_name in sorted(grouped_data.keys()):
            self._generate_customer_sheet(
                workbook, formats, partner_name, grouped_data[partner_name], report_month, reversal_ob_balances)

        return True

    def _generate_summary_sheet(self, workbook, formats, grouped_data, report_month, reversal_ob_balances=None):
        """Generate summary sheet with customer totals (no CE breakdown)"""
        if reversal_ob_balances is None:
            reversal_ob_balances = {}

        sheet_name = 'SUMMARY'
        sheet = workbook.add_worksheet(sheet_name)

        # Format month names
        month_full = report_month.strftime('%B %Y').upper()
        report_date = report_month.strftime('%m/%d/%Y')

        # Set column widths
        sheet.set_column(0, 0, 30)   # CLIENT
        sheet.set_column(1, 1, 40)   # DESCRIPTION
        sheet.set_column(2, 2, 8)    # Year
        sheet.set_column(3, 3, 12)   # Month
        sheet.set_column(4, 4, 18)   # BILLED
        sheet.set_column(5, 5, 18)   # System Accrual
        sheet.set_column(6, 6, 18)   # System Reversal
        sheet.set_column(7, 7, 18)   # Manual Accrual
        sheet.set_column(8, 8, 18)   # Manual Reversal
        sheet.set_column(9, 9, 18)   # ADDL ADJ
        sheet.set_column(10, 10, 15)  # Total

        # Write report header
        row = 0
        sheet.write(row, 0, self.env.company.name, formats['title'])
        row += 1
        sheet.write(
            row, 0, f'REVENUE REPORT SUMMARY - {month_full}', formats['title'])
        row += 1
        sheet.write(row, 0, report_date, formats['title'])
        row += 2

        # Write column headers
        headers = [
            'CLIENT', 'DESCRIPTION', 'Year', 'Month', 'BILLED',
            'SYSTEM ACCRUAL', 'SYSTEM REVERSAL', 'MANUAL ACCRUAL', 'MANUAL REVERSAL',
            'ADDL ADJ', 'TOTAL'
        ]

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Write summary data rows (aggregate by customer)
        for partner_name in sorted(grouped_data.keys()):
            ces_data = grouped_data[partner_name]

            # Aggregate all amounts for this customer
            total_amounts = {
                'system_accrual': 0,
                'system_reversal': 0,
                'manual_accrual': 0,
                'manual_reversal': 0,
                'addl_adj': 0
            }

            # Collect all descriptions, years, months, and sales orders
            descriptions = set()
            years = set()
            months = set()
            all_sales_orders = set()

            for ce_code, ce_data in ces_data.items():
                amounts = self._calculate_amounts_by_type(
                    ce_data['lines'], report_month)

                # Apply reversal OB fallback for this CE.
                # Try CE code first, then x_studio_old_ce from linked SO
                # (e.g. SO "BLFSO000211" with x_studio_old_ce "BLF 00004").
                norm_ce = self._normalize_ce_code(ce_code)
                rev_ob = reversal_ob_balances.get(norm_ce, {})
                if not rev_ob:
                    so_match = self._find_sale_order_by_ce_code(ce_code)
                    if so_match:
                        old_ce = getattr(so_match, 'x_studio_old_ce', '') or ''
                        if old_ce:
                            rev_ob = reversal_ob_balances.get(
                                self._normalize_ce_code(old_ce), {})

                system_reversal_val = amounts['system_reversal']
                if system_reversal_val == 0 and rev_ob.get('system_reversal', 0) != 0:
                    system_reversal_val = rev_ob['system_reversal']

                manual_reversal_val = amounts['manual_reversal']
                if manual_reversal_val == 0 and rev_ob.get('manual_reversal', 0) != 0:
                    manual_reversal_val = rev_ob['manual_reversal']
                manual_reversal_val -= rev_ob.get(
                    'manual_reversal_adjustment', 0)

                total_amounts['system_accrual'] += amounts['system_accrual']
                total_amounts['system_reversal'] += system_reversal_val
                total_amounts['manual_accrual'] += amounts['manual_accrual']
                total_amounts['manual_reversal'] += manual_reversal_val
                total_amounts['addl_adj'] += amounts['addl_adj']

                if ce_data['description']:
                    descriptions.add(ce_data['description'])
                if ce_data['year']:
                    years.add(str(ce_data['year']))
                if ce_data['month']:
                    months.add(ce_data['month'])

                # Collect all sales orders for this customer
                all_sales_orders.update(ce_data['sales_orders'])

            # Calculate billed amount for all sales orders of this customer
            billed_amount = self._calculate_billed_amount(
                all_sales_orders, report_month)

            # Write customer row
            sheet.write(row, 0, partner_name, formats['normal'])

            # Concatenate multiple descriptions if any
            desc_str = ', '.join(sorted(descriptions)) if descriptions else ''
            sheet.write(row, 1, desc_str, formats['normal'])

            # Year is always the report month's year
            year_str = str(report_month.year)
            sheet.write(row, 2, year_str, formats['centered'])

            # Concatenate multiple months
            month_str = report_month.strftime('%B').upper()
            sheet.write(row, 3, month_str, formats['centered'])

            # Write BILLED amount
            sheet.write(row, 4, billed_amount, formats['currency_negative'])

            sheet.write(
                row, 5, total_amounts['system_accrual'], formats['currency_negative'])
            sheet.write(
                row, 6, total_amounts['system_reversal'], formats['currency_negative'])
            sheet.write(
                row, 7, total_amounts['manual_accrual'], formats['currency_negative'])
            sheet.write(
                row, 8, total_amounts['manual_reversal'], formats['currency_negative'])
            sheet.write(
                row, 9, total_amounts['addl_adj'], formats['currency_negative'])

            # Total formula
            excel_row = row + 1
            sheet.write_formula(
                row, 10, f'=E{excel_row}+F{excel_row}+G{excel_row}+H{excel_row}+I{excel_row}+J{excel_row}', formats['currency'])

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
        for col in range(0, 3):
            sheet.write(total_row, col, '',
                        formats['section_header_no_border'])

        # TOTAL label
        sheet.write(total_row, 3, 'TOTAL', bold_with_border)

        # Sum formulas for monetary columns
        sheet.write_formula(
            total_row, 4, f'=SUM(E{data_start_row + 1}:E{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 5, f'=SUM(F{data_start_row + 1}:F{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 6, f'=SUM(G{data_start_row + 1}:G{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 7, f'=SUM(H{data_start_row + 1}:H{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 8, f'=SUM(I{data_start_row + 1}:I{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 9, f'=SUM(J{data_start_row + 1}:J{total_row})', currency_negative_bold_format)
        sheet.write_formula(
            total_row, 10, f'=SUM(K{data_start_row + 1}:K{total_row})', currency_bold_format)

        return True

    def _generate_customer_sheet(self, workbook, formats, partner_name, ces_data, report_month, reversal_ob_balances=None):
        """Generate individual customer sheet with CE breakdown"""
        if reversal_ob_balances is None:
            reversal_ob_balances = {}

        # Sanitize sheet name
        sheet_name = self._sanitize_sheet_name(partner_name)
        sheet = workbook.add_worksheet(sheet_name)

        # Format month names
        month_full = report_month.strftime('%B %Y').upper()
        report_date = report_month.strftime('%m/%d/%Y')
        cost_to_client_month = report_month.strftime('%B').upper()

        # Set column widths
        sheet.set_column(0, 0, 15)   # CE#
        sheet.set_column(1, 1, 18)   # SO Reference
        sheet.set_column(2, 2, 12)   # CE DATE
        sheet.set_column(3, 3, 40)   # DESCRIPTION
        sheet.set_column(4, 4, 8)    # Year
        sheet.set_column(5, 5, 12)   # Month
        sheet.set_column(6, 6, 18)   # BILLED
        sheet.set_column(7, 7, 18)   # System Accrual
        sheet.set_column(8, 8, 18)   # System Reversal
        sheet.set_column(9, 9, 18)   # Manual Accrual
        sheet.set_column(10, 10, 18)  # Manual Reversal
        sheet.set_column(11, 11, 18)  # ADDL ADJ
        sheet.set_column(12, 12, 15)  # Total
        sheet.set_column(13, 13, 20)  # CE Status
        sheet.set_column(14, 14, 15)  # Per CSD
        sheet.set_column(15, 15, 15)  # Variance
        sheet.set_column(16, 16, 20)  # Cost to Client
        sheet.set_column(17, 17, 20)  # For Revenue Adjustment
        sheet.set_column(18, 18, 30)  # Remarks

        # Write report header
        row = 0
        sheet.write(row, 0, self.env.company.name, formats['title'])
        row += 1
        sheet.write(row, 0, f'REVENUE REPORT - {month_full}', formats['title'])
        row += 1
        sheet.write(row, 0, report_date, formats['title'])
        row += 1
        sheet.write(row, 0, f'CLIENT: {partner_name}', formats['title'])
        row += 2

        # Write column headers
        headers = [
            'CE#', 'SO REFERENCE', 'CE DATE', 'DESCRIPTION', 'Year', 'Month', 'BILLED',
            'SYSTEM ACCRUAL', 'SYSTEM REVERSAL', 'MANUAL ACCRUAL', 'MANUAL REVERSAL',
            'ADDL ADJ', 'TOTAL', 'CE STATUS', 'PER CSD', 'VARIANCE',
            f'COST TO CLIENT - {cost_to_client_month}', 'FOR REVENUE ADJUSTMENT', 'REMARKS'
        ]

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Write data rows
        for ce_code in sorted(ces_data.keys()):
            ce_data = ces_data[ce_code]
            amounts = self._calculate_amounts_by_type(
                ce_data['lines'], report_month)

            # Calculate billed amount for this CE's sales orders
            billed_amount = self._calculate_billed_amount(
                ce_data['sales_orders'], report_month)

            sheet.write(row, 0, ce_code, formats['centered'])

            # Write SO Reference
            sheet.write(row, 1, ce_data['so_reference'], formats['centered'])

            if ce_data['ce_date']:
                sheet.write(row, 2, ce_data['ce_date'], formats['date'])
            else:
                sheet.write(row, 2, '', formats['centered'])

            sheet.write(row, 3, ce_data['description'], formats['normal'])

            if ce_data['year']:
                sheet.write(row, 4, ce_data['year'], formats['centered'])
            else:
                sheet.write(row, 4, '', formats['centered'])

            if ce_data['month']:
                sheet.write(row, 5, ce_data['month'], formats['centered'])
            else:
                sheet.write(row, 5, '', formats['centered'])

            # Write BILLED amount
            sheet.write(row, 6, billed_amount, formats['currency_negative'])

            # Apply reversal OB fallback for this CE.
            # Try the CE code first, then fall back to x_studio_old_ce
            # from the matching sale order (e.g. SO "BLFSO000211" has
            # x_studio_old_ce "BLF 00004" which matches the OB key).
            norm_ce = self._normalize_ce_code(ce_code)
            rev_ob = reversal_ob_balances.get(norm_ce, {})
            if not rev_ob:
                so_match = self._find_sale_order_by_ce_code(ce_code)
                if so_match:
                    old_ce = getattr(so_match, 'x_studio_old_ce', '') or ''
                    if old_ce:
                        rev_ob = reversal_ob_balances.get(
                            self._normalize_ce_code(old_ce), {})

            system_reversal_val = amounts['system_reversal']
            if system_reversal_val == 0 and rev_ob.get('system_reversal', 0) != 0:
                system_reversal_val = rev_ob['system_reversal']

            manual_reversal_val = amounts['manual_reversal']
            if manual_reversal_val == 0 and rev_ob.get('manual_reversal', 0) != 0:
                manual_reversal_val = rev_ob['manual_reversal']
            manual_reversal_val -= rev_ob.get('manual_reversal_adjustment', 0)

            sheet.write(row, 7, amounts['system_accrual'],
                        formats['currency_negative'])
            sheet.write(row, 8, system_reversal_val,
                        formats['currency_negative'])
            sheet.write(row, 9, amounts['manual_accrual'],
                        formats['currency_negative'])
            sheet.write(row, 10, manual_reversal_val,
                        formats['currency_negative'])
            sheet.write(row, 11, amounts['addl_adj'],
                        formats['currency_negative'])

            # Total formula (includes BILLED + accruals/reversals + ADDL ADJ: G+H+I+J+K+L which is columns 6-11)
            excel_row = row + 1
            sheet.write_formula(
                row, 12, f'=G{excel_row}+H{excel_row}+I{excel_row}+J{excel_row}+K{excel_row}+L{excel_row}', formats['currency'])

            sheet.write(row, 13, ce_data['ce_status'], formats['centered'])

            # PER CSD - empty for user input
            sheet.write(row, 14, '', formats['currency'])

            # VARIANCE formula: Total - Per CSD (M - O which is 12 - 14)
            sheet.write_formula(
                row, 15, f'=M{excel_row}-O{excel_row}', formats['currency'])

            # COST TO CLIENT - empty for user input
            sheet.write(row, 16, '', formats['currency'])

            # FOR REVENUE ADJUSTMENT formula: Variance - Cost to Client (P - Q which is 15 - 16)
            sheet.write_formula(
                row, 17, f'=P{excel_row}-Q{excel_row}', formats['currency'])

            # REMARKS - empty for user input
            sheet.write(row, 18, '', formats['normal'])

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
            total_row, 6, f'=SUM(G{data_start_row + 1}:G{total_row})', currency_negative_bold_format)
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

        # Empty CE Status cell
        sheet.write(total_row, 13, '', formats['section_header_no_border'])

        # Sum for PER CSD
        sheet.write_formula(
            total_row, 14, f'=SUM(O{data_start_row + 1}:O{total_row})', currency_bold_format)

        # Sum for VARIANCE
        sheet.write_formula(
            total_row, 15, f'=SUM(P{data_start_row + 1}:P{total_row})', currency_bold_format)

        # Sum for COST TO CLIENT
        sheet.write_formula(
            total_row, 16, f'=SUM(Q{data_start_row + 1}:Q{total_row})', currency_bold_format)

        # Sum for FOR REVENUE ADJUSTMENT
        sheet.write_formula(
            total_row, 17, f'=SUM(R{data_start_row + 1}:R{total_row})', currency_bold_format)

        # Empty REMARKS cell
        sheet.write(total_row, 18, '', formats['section_header_no_border'])

        return True
