from odoo import models
import datetime
import re
from xlsxwriter.workbook import Workbook
from odoo.exceptions import ValidationError, UserError
from dateutil.relativedelta import relativedelta
import logging
from collections import defaultdict

_logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# FEATURE SWITCH: excluded-account handling in the BILLED column.
#
# When True (the original behaviour), journal items posted to the accounts
# listed in saatchi.accrual_config.excluded_billed_account_ids are netted out
# of the BILLED figure.  In this database those are, per company:
#     1218  Unbilled Charges_WIP             (asset_current)
#     4310  Rental Income                    (income_other)
#     4112  Cost of Services - Unbilled WIP  (expense_direct_cost)
#
# Set to False on 2026-08-09 at the client's request -- BILLED is now reported
# GROSS, with no account exclusions applied anywhere in this report.
#
# To re-enable, flip this back to True.  That is the ONLY change required:
#   * the saatchi.accrual_config records were deliberately left untouched, and
#   * every consumer of _get_excluded_billed_account_ids() guards its exclusion
#     branch with `if excluded_account_ids:` -- namely _calculate_billed_amount
#     (both _apply_invoice and the STEP 3 credit-memo branch) and
#     _build_pnl_billed_map -- so the whole mechanism switches on and off here.
#
# NOTE for future devs: base_customization/models/inherit.py has its own,
# separate _get_excluded_billed_account_ids() feeding the stored fields
# x_excluded_billed_amount / x_billed_amount_adjustment on account.move.  Those
# fields are NOT read by this report and are unaffected by this switch.
# ---------------------------------------------------------------------------
EXCLUDE_BILLED_ACCOUNTS = False


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

    def _normalize_ce_code_loose(self, ce_code):
        """Last-resort CE normalisation: strip EVERY non-alphanumeric character.

        `_normalize_ce_code` removes whitespace only, so it keeps separators --
        which means an invoice written 'PJB 0178-1' and a credit note written
        'PJB0178 1' normalise to 'PJB0178-1' and 'PJB01781' and never match.
        That silently dropped the credit note from the BILLED column
        (P&G Philippines, July 2026: billed overstated by 49,407.25).

        Used ONLY as a fallback for credit notes that matched nothing under the
        strict rule -- never as the primary comparison -- because collapsing
        separators is deliberately loose and could otherwise merge genuinely
        different CE codes (e.g. 'ABC-1' and 'ABC1').

        Args:
            ce_code (str): raw CE code

        Returns:
            str: uppercased, alphanumerics only ('PJB 0178-1' -> 'PJB01781')
        """
        if not ce_code:
            return ''
        return re.sub(r'[^A-Za-z0-9]', '', ce_code).upper()

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
                'so_reference': (so.name.upper() if so and so.name else ''),
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

    def _group_lines_by_ce(self, lines, report_month, all_billed_so_ids=None, standalone_invoice_ids=None):
        """Group account.move.line records by partner and CE code

        Includes:
        - Lines from accrued revenue entries
        - ALL sales orders that were billed in the report month (whether they have accrued entries or not)
        - Standalone invoices: posted invoices with x_studio_old_ce_1 set but no matching SO
        """
        grouped = defaultdict(lambda: defaultdict(lambda: {
            'ce_date': None,
            'description': '',
            'year': None,
            'month': None,
            'ce_status': '',
            'so_reference': '',
            'lines': [],
            'sales_orders': set(),
            'direct_invoices': set(),
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

                # Set SO Reference from SO name if not already set
                if so.name and not grouped[partner_name][ce_code]['so_reference']:
                    grouped[partner_name][ce_code]['so_reference'] = so.name.upper()

        # Fallback: fill so_reference from tracked sales_orders if still empty
        for partner_name, ces in grouped.items():
            for ce_code, ce_data in ces.items():
                if not ce_data['so_reference'] and ce_data['sales_orders']:
                    so_records = self.env['sale.order'].browse(
                        list(ce_data['sales_orders']))
                    so_names = [s.name for s in so_records if s.name]
                    if so_names:
                        ce_data['so_reference'] = ', '.join(
                            so_names).upper()

        # Standalone invoices: posted out_invoices with x_studio_old_ce_1 set
        # but no matching SO. Create CE rows directly from them.
        if standalone_invoice_ids:
            standalone_invoices = self.env['account.move'].browse(standalone_invoice_ids)
            for inv in standalone_invoices:
                if not inv.partner_id or not inv.x_studio_old_ce_1:
                    continue
                partner_name = inv.partner_id.name.upper()
                ce_code = inv.x_studio_old_ce_1.upper()
                grouped[partner_name][ce_code]['direct_invoices'].add(inv.id)
                if inv.invoice_date and not grouped[partner_name][ce_code]['ce_date']:
                    grouped[partner_name][ce_code]['ce_date'] = inv.invoice_date
                    grouped[partner_name][ce_code]['year'] = inv.invoice_date.year
                    grouped[partner_name][ce_code]['month'] = inv.invoice_date.strftime('%B').upper()

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

        Currently DISABLED -- returns an empty set unless the module-level
        EXCLUDE_BILLED_ACCOUNTS switch is turned back on.  See the comment block
        at the top of this file for the full rationale and how to re-enable.

        Returns:
            set: account IDs to exclude (empty while the switch is off)
        """
        if not EXCLUDE_BILLED_ACCOUNTS:
            # Switch is off: returning an empty set makes every caller skip its
            # `if excluded_account_ids:` branch, so BILLED is reported gross.
            # The config records below are still read when the switch is on.
            return set()

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

    def _get_cm_deduction_for_invoice(self, invoice, report_month):
        """Get the total credit memo deduction applied against an invoice.

        Only considers credit memos that fall within the same report month
        (based on invoice_date) to ensure month-specific revenue adjustments.

        Uses the actual reconciled amount (partial.amount) from
        account.partial.reconcile — NOT the full CM total.  This ensures
        that when only ₱0.01 of a ₱30k CM is reconciled against the
        invoice, we deduct ₱0.01, not ₱30k.

        For each partial reconcile involving a credit memo:
        - If the CM was *fully* reconciled against this invoice
          (partial.amount >= CM amount_residual before matching, i.e. the
          CM amount_residual is now 0), we use the CM's untaxed amount
          in company currency so the deduction matches the billed
          calculation basis (untaxed).
        - If the CM was only *partially* reconciled, we use
          partial.amount (the actual reconciled amount in company
          currency) as the deduction.

        Args:
            invoice: account.move record (the invoice)
            report_month: date object representing the report month

        Returns:
            float: total CM deduction amount (positive value)
        """
        # Get the start and end of the report month
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - \
            relativedelta(days=1)

        cm_total = 0.0

        receivable_lines = invoice.line_ids.filtered(
            lambda l: l.account_id.account_type == 'asset_receivable'
        )

        processed_cm_ids = set()

        for line in receivable_lines:
            for partial in line.matched_credit_ids:
                counterpart = partial.credit_move_id
                cm_move = counterpart.move_id
                if (cm_move.move_type == 'out_refund' and
                    cm_move.id not in processed_cm_ids and
                    cm_move.invoice_date and
                        month_start <= cm_move.invoice_date <= month_end):
                    processed_cm_ids.add(cm_move.id)
                    # Check if the CM was fully applied to this invoice
                    cm_total_amount = cm_move.amount_total_in_currency_signed
                    if cm_total_amount and abs(partial.amount - abs(cm_total_amount)) < 0.02:
                        # Fully reconciled — use untaxed amount for consistency
                        cm_total += self._get_amount_untaxed_in_company_currency(
                            cm_move)
                    else:
                        # Partially reconciled — only deduct what was actually applied
                        cm_total += partial.amount

            for partial in line.matched_debit_ids:
                counterpart = partial.debit_move_id
                cm_move = counterpart.move_id
                if (cm_move.move_type == 'out_refund' and
                    cm_move.id not in processed_cm_ids and
                    cm_move.invoice_date and
                        month_start <= cm_move.invoice_date <= month_end):
                    processed_cm_ids.add(cm_move.id)
                    cm_total_amount = cm_move.amount_total_in_currency_signed
                    if cm_total_amount and abs(partial.amount - abs(cm_total_amount)) < 0.02:
                        cm_total += self._get_amount_untaxed_in_company_currency(
                            cm_move)
                    else:
                        cm_total += partial.amount

        return cm_total, processed_cm_ids

    def _calculate_billed_amount(self, sales_order_ids, report_month, direct_invoice_ids=None,
                                 consumed_move_ids=None):
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
            direct_invoice_ids: set of account.move IDs for standalone invoices
                                 (have x_studio_old_ce_1 but no matching SO)

        Returns:
            float: Adjusted billed amount for posted invoices in the report month
        """
        if not sales_order_ids and not direct_invoice_ids:
            return 0.0, 0.0

        # Get the start and end of the report month
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - \
            relativedelta(days=1)

        # Get excluded account IDs from accrual configuration
        excluded_account_ids = self._get_excluded_billed_account_ids()

        total_billed = 0.0
        total_exclusion = 0.0
        total_cm_deduction = 0.0
        processed_invoice_ids = set()
        processed_cm_ids = set()

        def _apply_invoice(inv):
            """Calculate billed, exclusion, and CM deduction for a single invoice.

            Billed = untaxed amount in company currency minus CM deductions
            Exclusion = net amount on excluded accounts
            CM deduction = credit memos reconciled as payment

            Returns (net_billed_delta, exclusion_delta, cm_deduction).
            """
            gross_billed = self._get_amount_untaxed_in_company_currency(inv)

            # Get credit memos that were applied as payment
            cm_deduction, cm_ids = self._get_cm_deduction_for_invoice(
                inv, report_month)
            processed_cm_ids.update(cm_ids)
            net_billed = gross_billed
            if cm_deduction:
                _logger.info(
                    'Invoice %s: subtracting CM deduction %.2f from billed %.2f',
                    inv.name, cm_deduction, gross_billed,
                )
                net_billed -= cm_deduction

            exclusion_delta = 0.0
            if excluded_account_ids:
                for line in inv.line_ids:
                    if line.account_id and line.account_id.id in excluded_account_ids:
                        exclusion_delta += (line.credit or 0.0)
                        exclusion_delta -= (line.debit or 0.0)
            return net_billed, exclusion_delta, cm_deduction

        # old_ce_codes is shared across STEP 0-3 so CMs are always picked up by STEP 3
        old_ce_codes = set()

        # ------------------------------------------------------------------
        # STEP 0: Process standalone invoices — posted out_invoices that have
        #         x_studio_old_ce_1 but no matching SO anywhere.
        #         Also seed old_ce_codes with their x_studio_old_ce_1 values so
        #         STEP 3 can deduct matching credit memos in the same month.
        # ------------------------------------------------------------------
        if direct_invoice_ids:
            for inv in self.env['account.move'].browse(list(direct_invoice_ids)):
                if inv.state != 'posted' or inv.move_type != 'out_invoice':
                    continue
                if not (inv.invoice_date and month_start <= inv.invoice_date <= month_end):
                    continue
                if inv.id in processed_invoice_ids:
                    continue
                processed_invoice_ids.add(inv.id)
                b_delta, e_delta, cm_delta = _apply_invoice(inv)
                total_billed += b_delta
                total_exclusion += e_delta
                total_cm_deduction += cm_delta
                inv_ce1 = getattr(inv, 'x_studio_old_ce_1', '') or ''
                if inv_ce1:
                    old_ce_codes.add(self._normalize_ce_code(inv_ce1))
                _logger.info(
                    'Included standalone invoice %s (old CE: %s) in billed with billed_delta %.2f',
                    inv.name, inv_ce1, b_delta,
                )

        # Browse sales orders
        sales_orders = self.env['sale.order'].browse(list(sales_order_ids)) if sales_order_ids else self.env['sale.order']

        # ------------------------------------------------------------------
        # STEP 1: Process SO-linked invoices (out_invoice only)
        # ------------------------------------------------------------------
        # old_ce_codes continues to be populated from SO x_studio_old_ce values

        for so in sales_orders:
            # Collect the SO's old CE code so we can match unlinked invoices below
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
                b_delta, e_delta, cm_delta = _apply_invoice(move)
                total_billed += b_delta
                total_exclusion += e_delta
                total_cm_deduction += cm_delta

        # ------------------------------------------------------------------
        # STEP 2: Find invoices whose x_studio_old_ce_1 matches x_studio_old_ce
        #         on any SO in this CE row.  No invoice-type or sale-link
        #         restriction — CE code match is the sole criterion.
        # ------------------------------------------------------------------
        _logger.info(sales_orders.mapped('name'))
        if old_ce_codes:
            ce_inv_domain = [
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
                ('invoice_date', '>=', month_start),
                ('invoice_date', '<=', month_end),
                ('company_id', 'in', self.env.companies.ids),
                ('x_studio_old_ce_1', '!=', False),
            ]

            try:
                ce_inv_candidates = self.env['account.move'].sudo().search(
                    ce_inv_domain)
            except Exception:
                ce_inv_candidates = self.env['account.move']

            for inv in ce_inv_candidates:
                if inv.id in processed_invoice_ids:
                    continue
                inv_old_ce = getattr(inv, 'x_studio_old_ce_1', '') or ''
                if not inv_old_ce:
                    continue
                if self._normalize_ce_code(inv_old_ce) not in old_ce_codes:
                    continue

                processed_invoice_ids.add(inv.id)
                b_delta, e_delta, cm_delta = _apply_invoice(inv)
                total_billed += b_delta
                total_exclusion += e_delta
                total_cm_deduction += cm_delta
                _logger.info(
                    'Included CE-matched invoice %s (old CE: %s) in billed '
                    'with billed_delta %.2f. Total billed so far: %.2f',
                    inv.name, inv_old_ce, b_delta, total_billed,
                )

        # ------------------------------------------------------------------
        # STEP 3: Find credit memos whose x_studio_old_ce_1 matches
        #         x_studio_old_ce on any SO in this CE row.
        # ------------------------------------------------------------------
        if old_ce_codes:
            ce_cm_domain = [
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_refund'),
                ('invoice_date', '>=', month_start),
                ('invoice_date', '<=', month_end),
                ('company_id', 'in', self.env.companies.ids),
                ('x_studio_old_ce_1', '!=', False),
            ]
            try:
                ce_cm_candidates = self.env['account.move'].sudo().search(
                    ce_cm_domain)
            except Exception:
                ce_cm_candidates = self.env['account.move']

            # Loose variants of the CE codes we are looking for, used only as a
            # last resort below.  Built once rather than per candidate.
            old_ce_codes_loose = {
                self._normalize_ce_code_loose(c) for c in old_ce_codes if c
            }

            for cm in ce_cm_candidates:
                if cm.id in processed_cm_ids:
                    continue
                cm_old_ce = getattr(cm, 'x_studio_old_ce_1', '') or ''
                if not cm_old_ce:
                    continue

                # Tier 1: strict match (whitespace stripped, uppercased).
                if self._normalize_ce_code(cm_old_ce) not in old_ce_codes:
                    # Tier 2 (last resort): separators stripped too.  Catches
                    # 'PJB0178 1' vs 'PJB 0178-1', where a hyphen on one side
                    # and a space on the other defeated the strict compare.
                    if self._normalize_ce_code_loose(cm_old_ce) in old_ce_codes_loose:
                        _logger.warning(
                            'Credit memo %s (old CE %r) matched only after stripping '
                            'separators -- the CE code formatting differs from the '
                            'invoice/SO it belongs to. Deducting %.2f; consider '
                            'correcting the CE code on the credit memo.',
                            cm.name, cm_old_ce, cm.amount_untaxed or 0.0,
                        )
                    else:
                        # Nothing matched.  Log it: an unmatched credit memo is
                        # silently absent from BILLED, which is exactly how the
                        # July 2026 P&G discrepancy went unnoticed.
                        _logger.warning(
                            'Credit memo %s (old CE %r, %.2f) matched no sale order '
                            'or invoice in this CE row and was NOT deducted from '
                            'BILLED. Check its Old CE# against the related invoice.',
                            cm.name, cm_old_ce, cm.amount_untaxed or 0.0,
                        )
                        continue

                cm_gross = self._get_amount_untaxed_in_company_currency(cm)
                cm_exclusion = 0.0
                if excluded_account_ids:
                    for line in cm.line_ids:
                        if line.account_id and line.account_id.id in excluded_account_ids:
                            cm_exclusion += (line.credit or 0.0)
                            cm_exclusion -= (line.debit or 0.0)

                cm_amount = cm_gross - abs(cm_exclusion)
                total_billed -= cm_amount
                total_cm_deduction += cm_amount
                processed_cm_ids.add(cm.id)
                _logger.info(
                    'Included CE-matched CM %s (old CE: %s): gross %.2f, excluded %.2f, net deduction %.2f',
                    cm.name, cm_old_ce, cm_gross, cm_exclusion, cm_amount,
                )

        # Report back which documents actually made it into BILLED, so the
        # reconciliation block on the Summary / customer sheets can diagnose the
        # ones that did not.  Optional -- callers that pass nothing are unaffected.
        if consumed_move_ids is not None:
            consumed_move_ids.update(processed_invoice_ids)
            consumed_move_ids.update(processed_cm_ids)

        return total_billed - total_exclusion, total_cm_deduction

    # ------------------------------------------------------------------
    # PnL billed helpers (Option 2)
    # ------------------------------------------------------------------

    def _resolve_pnl_ce_from_so_links(self, move):
        """Return normalized CE code from the first SO found on move's invoice lines."""
        try:
            for aml in move.line_ids:
                if not aml.sale_line_ids:
                    continue
                so = aml.sale_line_ids[0].order_id
                if not so:
                    continue
                old_ce = getattr(so, 'x_studio_old_ce', '') or ''
                if old_ce:
                    return self._normalize_ce_code(old_ce)
                if so.x_ce_code:
                    return self._normalize_ce_code(so.x_ce_code)
        except Exception as e:
            _logger.debug('_resolve_pnl_ce_from_so_links error on %s: %s', getattr(move, 'name', '?'), e)
        return None

    def _resolve_pnl_ce_code(self, move, inv_ce_cache=None):
        """Resolve normalized CE code for any posted move for PnL billed calculation.

        Priority chain:
        1. move.x_studio_old_ce_1 directly
        2. CMs only: walk matched_debit_ids reconciliation → reconciled invoice →
           invoice.x_studio_old_ce_1, then invoice SO links
        3. SO links on the move itself (invoice or CM)
        """
        try:
            # 1. Direct field
            ce1 = getattr(move, 'x_studio_old_ce_1', '') or ''
            if ce1 and ce1.strip():
                return self._normalize_ce_code(ce1)

            # 2. CMs: walk reconciliation
            if move.move_type == 'out_refund':
                receivable = move.line_ids.filtered(
                    lambda l: l.account_id.account_type == 'asset_receivable'
                )
                for line in receivable:
                    for partial in line.matched_debit_ids:
                        try:
                            inv = partial.debit_move_id.move_id
                        except Exception:
                            continue
                        if not inv or inv.move_type != 'out_invoice':
                            continue
                        # Use pre-built cache when available
                        if inv_ce_cache is not None and inv.id in inv_ce_cache:
                            cached = inv_ce_cache[inv.id]
                            if cached:
                                return cached
                        # Direct field on reconciled invoice
                        inv_ce1 = getattr(inv, 'x_studio_old_ce_1', '') or ''
                        if inv_ce1 and inv_ce1.strip():
                            return self._normalize_ce_code(inv_ce1)
                        # SO links on reconciled invoice
                        ce = self._resolve_pnl_ce_from_so_links(inv)
                        if ce:
                            return ce

            # 3. SO links on the move itself
            return self._resolve_pnl_ce_from_so_links(move)

        except Exception as e:
            _logger.debug('_resolve_pnl_ce_code error on %s: %s', getattr(move, 'name', '?'), e)
            return None

    def _build_pnl_billed_map(self, report_month, partner_ids=None):
        """Build {norm_ce_code: net_billed_float} for all posted invoices and CMs
        in the report month.  Called once per report for PnL mode.

        Invoices contribute positively; CMs negatively.
        Excluded-account line amounts are stripped.
        Two-pass approach: invoices are resolved first to seed a cache used
        by CMs during their reconciliation walk.
        """
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - relativedelta(days=1)

        domain = [
            ('state', '=', 'posted'),
            ('move_type', 'in', ['out_invoice', 'out_refund']),
            ('invoice_date', '>=', month_start),
            ('invoice_date', '<=', month_end),
            ('company_id', 'in', self.env.companies.ids),
        ]
        if partner_ids:
            domain.append(('partner_id', 'in', partner_ids))

        try:
            moves = self.env['account.move'].sudo().search(domain)
        except Exception as e:
            _logger.warning('_build_pnl_billed_map: search failed: %s', e)
            return {}

        if not moves:
            return {}

        excluded_account_ids = self._get_excluded_billed_account_ids()

        # Pass 1 — resolve CE codes for all invoices and build cache
        inv_ce_cache = {}  # {inv_id: norm_ce or None}
        invoices = moves.filtered(lambda m: m.move_type == 'out_invoice')
        for inv in invoices:
            ce1 = getattr(inv, 'x_studio_old_ce_1', '') or ''
            if ce1 and ce1.strip():
                inv_ce_cache[inv.id] = self._normalize_ce_code(ce1)
            else:
                inv_ce_cache[inv.id] = self._resolve_pnl_ce_from_so_links(inv)

        # Pass 2 — accumulate net billed per CE code
        ce_billed_map = defaultdict(float)

        for move in moves:
            try:
                norm_ce = self._resolve_pnl_ce_code(move, inv_ce_cache)
            except Exception as e:
                _logger.debug('PnL CE resolve error for %s: %s', getattr(move, 'name', '?'), e)
                continue

            if not norm_ce:
                continue

            try:
                amount = self._get_amount_untaxed_in_company_currency(move)
            except Exception:
                amount = move.amount_untaxed or 0.0

            exclusion = 0.0
            if excluded_account_ids:
                try:
                    for line in move.line_ids:
                        if line.account_id and line.account_id.id in excluded_account_ids:
                            exclusion += (line.credit or 0.0)
                            exclusion -= (line.debit or 0.0)
                except Exception:
                    pass

            net = amount - abs(exclusion)
            if move.move_type == 'out_invoice':
                ce_billed_map[norm_ce] += net
            else:
                ce_billed_map[norm_ce] -= net

        return dict(ce_billed_map)

    def _get_pnl_billed_for_row(self, ce_code, ce_data, pnl_billed_map, consumed_ce_codes=None):
        """Look up the PnL net billed for a CE row from the pre-built map.

        Checks the row's CE code key plus all CE variants from its SOs and
        direct invoices.  consumed_ce_codes (set) prevents double-counting when
        the same norm CE appears in two rows (rare data-quality issue).
        """
        if not pnl_billed_map:
            return 0.0

        variants = set()
        norm_key = self._normalize_ce_code(ce_code)
        if norm_key:
            variants.add(norm_key)

        # Variants from SOs
        so_ids = ce_data.get('sales_orders') or set()
        if so_ids:
            for so in self.env['sale.order'].browse(list(so_ids)):
                try:
                    if so.x_ce_code:
                        variants.add(self._normalize_ce_code(so.x_ce_code))
                    old_ce = getattr(so, 'x_studio_old_ce', '') or ''
                    if old_ce:
                        variants.add(self._normalize_ce_code(old_ce))
                except Exception:
                    pass

        # Variants from direct (standalone) invoices
        inv_ids = ce_data.get('direct_invoices') or set()
        if inv_ids:
            for inv in self.env['account.move'].browse(list(inv_ids)):
                try:
                    ce1 = getattr(inv, 'x_studio_old_ce_1', '') or ''
                    if ce1:
                        variants.add(self._normalize_ce_code(ce1))
                except Exception:
                    pass

        total = 0.0
        for v in variants:
            if v not in pnl_billed_map:
                continue
            if consumed_ce_codes is not None and v in consumed_ce_codes:
                continue
            total += pnl_billed_map[v]
            if consumed_ce_codes is not None:
                consumed_ce_codes.add(v)

        return total

    # ------------------------------------------------------------------
    # BILLED reconciliation (Summary + per-customer sheets)
    # ------------------------------------------------------------------
    #
    # Finance reconciles this report by opening Accounting > Invoices, filtering
    # Invoices-or-Credit-Notes / Posted / <report month> / Invoice Type =
    # "CE Related", and reading the "Billed Amount (w/TP Cost)" total.  The
    # helpers below reproduce exactly that control total at generation time,
    # compare it with the BILLED the report produced, and name every document
    # behind any difference -- so a discrepancy is explained in the workbook
    # instead of being hunted invoice by invoice.
    # ------------------------------------------------------------------

    def _get_billed_control_moves(self, report_month, partner_ids=None):
        """Posted customer documents that finance's Odoo filter would show.

        Mirrors: Invoices or Credit Notes + Posted + invoice_date in the report
        month + Invoice Type = 'CE Related'.

        Returns:
            account.move recordset
        """
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - relativedelta(days=1)
        domain = [
            ('state', '=', 'posted'),
            ('move_type', 'in', ['out_invoice', 'out_refund']),
            ('invoice_date', '>=', month_start),
            ('invoice_date', '<=', month_end),
            ('company_id', 'in', self.env.companies.ids),
            ('x_studio_invoice_type', '=', 'CE Related'),
        ]
        if partner_ids:
            domain.append(('partner_id', 'in', list(partner_ids)))
        try:
            return self.env['account.move'].sudo().search(domain)
        except Exception as e:
            _logger.warning('Reconciliation: control-set search failed: %s', e)
            return self.env['account.move']

    def _diagnose_unbilled_moves(self, moves, report_month):
        """Explain, per document, why it never reached the BILLED column.

        Args:
            moves: account.move recordset (already known to be excluded)
            report_month: date in the report month

        Returns:
            list of dicts: name, doc_type, partner, old_ce, amount, reason
        """
        month_start = report_month.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - relativedelta(days=1)

        # CE codes the report can currently attach a document to.
        so_all = self.env['sale.order'].sudo().search(
            [('x_studio_old_ce', '!=', False)])
        inv_month = self.env['account.move'].sudo().search([
            ('state', '=', 'posted'), ('move_type', '=', 'out_invoice'),
            ('invoice_date', '>=', month_start), ('invoice_date', '<=', month_end),
            ('x_studio_old_ce_1', '!=', False),
        ])
        known_strict, known_loose = set(), set()
        for val in list(so_all.mapped('x_studio_old_ce')) + list(inv_month.mapped('x_studio_old_ce_1')):
            if val:
                known_strict.add(self._normalize_ce_code(val))
                known_loose.add(self._normalize_ce_code_loose(val))

        rows = []
        for mv in moves:
            old_ce = (getattr(mv, 'x_studio_old_ce_1', '') or '').strip()
            has_so = bool(mv.line_ids.mapped('sale_line_ids'))
            is_cm = mv.move_type == 'out_refund'

            if not old_ce and not has_so:
                reason = ('No Old CE# and no sale order link, so the report has no CE row '
                          'to attach it to. Set the Old CE# on this document.')
            elif not old_ce and has_so:
                reason = ('Has a sale order but no Old CE#; its sale order produced no '
                          'accrual or billing row this month.')
            elif self._normalize_ce_code(old_ce) not in known_strict:
                if self._normalize_ce_code_loose(old_ce) in known_loose:
                    reason = (f"Old CE# '{old_ce}' is formatted differently from the sale "
                              f"order/invoice it belongs to (spaces vs separators). "
                              f"Align the CE code formatting.")
                elif is_cm:
                    reason = (f"Standalone credit note: no sale order and no invoice this "
                              f"month carries Old CE# '{old_ce}', so no CE row exists for it.")
                else:
                    reason = (f"Old CE# '{old_ce}' matches no sale order and no other "
                              f"invoice this month.")
            else:
                reason = ('Its CE code is known, but no CE row in this report claimed it '
                          '(no accrual activity and no billed sale order this month).')

            rows.append({
                'name': mv.name or '',
                'doc_type': 'Credit Note' if is_cm else 'Invoice',
                'partner': mv.partner_id.name or '',
                'old_ce': old_ce or '(none)',
                'amount': mv.x_amount_untaxed_in_company_currency or 0.0,
                'reason': reason,
            })
        rows.sort(key=lambda r: (-abs(r['amount']), r['name']))
        return rows

    def _write_reconciliation_block(self, workbook, sheet, start_row, report_month,
                                    report_billed_total, consumed_move_ids,
                                    partner_ids=None, billed_mode='standard',
                                    show_customer=True):
        """Render the BILLED reconciliation block. Returns the next free row.

        Laid out below the sheet's TOTAL row with breathing space above it, so it
        reads as a separate section rather than more data.
        """
        base = {'font_name': 'Calibri', 'font_size': 10}
        f_banner = workbook.add_format({
            **base, 'bold': True, 'font_size': 12, 'font_color': '#FFFFFF',
            'bg_color': '#1F3864', 'align': 'left', 'valign': 'vcenter', 'border': 1})
        f_label = workbook.add_format({**base, 'align': 'left', 'valign': 'vcenter'})
        f_label_b = workbook.add_format({**base, 'bold': True, 'align': 'left', 'valign': 'vcenter'})
        f_money = workbook.add_format({
            **base, 'num_format': '#,##0.00;(#,##0.00);"-"', 'align': 'right', 'valign': 'vcenter'})
        f_money_b = workbook.add_format({
            **base, 'bold': True, 'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right', 'valign': 'vcenter', 'top': 1})
        f_diff_bad = workbook.add_format({
            **base, 'bold': True, 'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right', 'valign': 'vcenter', 'font_color': '#9C0006',
            'bg_color': '#FFC7CE', 'border': 1})
        f_diff_ok = workbook.add_format({
            **base, 'bold': True, 'num_format': '#,##0.00;(#,##0.00);"-"',
            'align': 'right', 'valign': 'vcenter', 'font_color': '#006100',
            'bg_color': '#C6EFCE', 'border': 1})
        f_hdr = workbook.add_format({
            **base, 'bold': True, 'bg_color': '#D9E1F2', 'align': 'center',
            'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        f_cell = workbook.add_format({**base, 'align': 'left', 'valign': 'top', 'border': 1})
        f_cell_c = workbook.add_format({**base, 'align': 'center', 'valign': 'top', 'border': 1})
        f_cell_m = workbook.add_format({
            **base, 'num_format': '#,##0.00;(#,##0.00);"-"', 'align': 'right',
            'valign': 'top', 'border': 1})
        # No text_wrap: the reason column sits in a narrow numeric column that
        # cannot be widened without distorting the data table above, so the text
        # is left to overflow into the empty cells to its right, where it stays
        # readable on one line.
        f_reason = workbook.add_format({
            **base, 'align': 'left', 'valign': 'vcenter'})
        f_note_ok = workbook.add_format({**base, 'italic': True, 'font_color': '#006100'})
        f_note_bad = workbook.add_format({**base, 'italic': True, 'font_color': '#9C0006'})

        control_moves = self._get_billed_control_moves(report_month, partner_ids)
        control_total = sum(
            m.x_amount_untaxed_in_company_currency or 0.0 for m in control_moves)
        difference = report_billed_total - control_total
        month_label = report_month.strftime('%B %Y').upper()

        row = start_row + 2  # breathing space under the TOTAL row

        # ASCII only in every user-visible string here: these land in a workbook
        # opened on Windows Excel, and a stray em-dash gains nothing.
        sheet.merge_range(row, 0, row, 5,
                          f'BILLED RECONCILIATION - {month_label}', f_banner)
        sheet.set_row(row, 20)
        row += 2

        sheet.write(row, 0,
                    'Odoo control total  (Invoices + Credit Notes / Posted / CE Related)',
                    f_label)
        sheet.write(row, 4, control_total, f_money)
        row += 1
        sheet.write(row, 0, 'This report - BILLED total', f_label)
        sheet.write(row, 4, report_billed_total, f_money_b)
        row += 1
        sheet.write(row, 0, 'DIFFERENCE', f_label_b)
        sheet.write(row, 4, difference,
                    f_diff_ok if abs(difference) < 0.01 else f_diff_bad)
        row += 2

        if abs(difference) < 0.01:
            sheet.write(row, 0,
                        'No discrepancy - every posted CE Related document for this month is '
                        'accounted for in the BILLED column.', f_note_ok)
            return row + 1

        if billed_mode != 'standard':
            sheet.write(row, 0,
                        'Document-level diagnosis is available in Standard billed mode only; '
                        'this report was generated in PnL mode.', f_note_bad)
            return row + 1

        excluded = control_moves.filtered(lambda m: m.id not in (consumed_move_ids or set()))
        details = self._diagnose_unbilled_moves(excluded, report_month)

        sheet.write(row, 0, 'Documents not included in the BILLED column', f_label_b)
        row += 1

        headers = ['Document', 'Type', 'Customer', 'Old CE#', 'Amount', 'Why it is not included']
        if not show_customer:
            headers.pop(2)
        for col, head in enumerate(headers):
            sheet.write(row, col, head, f_hdr)
        sheet.set_row(row, 28)
        row += 1

        listed = 0.0
        for d in details:
            col = 0
            sheet.write(row, col, d['name'], f_cell_c); col += 1
            sheet.write(row, col, d['doc_type'], f_cell_c); col += 1
            if show_customer:
                sheet.write(row, col, d['partner'], f_cell); col += 1
            sheet.write(row, col, d['old_ce'], f_cell_c); col += 1
            sheet.write(row, col, d['amount'], f_cell_m); col += 1
            sheet.write(row, col, d['reason'], f_reason)
            listed += d['amount']
            row += 1

        row += 1
        # -listed because an excluded document's own sign is the opposite of the
        # effect its absence has on the report total.
        unexplained = difference - (-listed)
        if abs(unexplained) < 0.01:
            sheet.write(row, 0,
                        f'The {len(details)} document(s) above fully account for the difference.',
                        f_note_ok)
        else:
            sheet.write(row, 0,
                        f'The {len(details)} document(s) above account for '
                        f'{-listed:,.2f} of the {difference:,.2f} difference - '
                        f'{unexplained:,.2f} is still unexplained and needs review.',
                        f_note_bad)
        return row + 1

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
        standalone_invoice_ids = data.get('standalone_invoice_ids', []) if data else []
        billed_mode = data.get('billed_mode', 'standard') if data else 'standard'

        # Group lines by partner and CE (includes both accrued and billed)
        grouped_data = self._group_lines_by_ce(
            move_lines, report_month, all_billed_so_ids, standalone_invoice_ids)

        # Merge reversal opening balance rows (for CEs only in OB, no journal entries)
        self._merge_reversal_opening_balance_rows(grouped_data)

        # Fill in missing CE metadata (ce_status, description, ce_date) from SO / reversal OB
        self._fill_missing_ce_metadata(grouped_data)

        if not grouped_data:
            raise UserError("No data found for the report month.")

        # Calculate reversal opening balances for fallback
        reversal_ob_balances = self._calculate_reversal_opening_balances(
            report_month)

        # For PnL mode: build the billed map once for the entire report
        pnl_billed_map = {}
        if billed_mode == 'pnl':
            partner_ids = data.get('partner_ids') if data else None
            pnl_billed_map = self._build_pnl_billed_map(report_month, partner_ids or None)

        # Generate summary sheet first
        self._generate_summary_sheet(
            workbook, formats, grouped_data, report_month, reversal_ob_balances,
            billed_mode=billed_mode, pnl_billed_map=pnl_billed_map)

        # Generate individual customer sheets
        for partner_name in sorted(grouped_data.keys()):
            self._generate_customer_sheet(
                workbook, formats, partner_name, grouped_data[partner_name], report_month,
                reversal_ob_balances, billed_mode=billed_mode, pnl_billed_map=pnl_billed_map)

        return True

    def _generate_summary_sheet(self, workbook, formats, grouped_data, report_month,
                                reversal_ob_balances=None, billed_mode='standard', pnl_billed_map=None):
        """Generate summary sheet with customer totals (no CE breakdown)"""
        if reversal_ob_balances is None:
            reversal_ob_balances = {}
        if pnl_billed_map is None:
            pnl_billed_map = {}

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

        # Accumulated while writing the rows below, then used by the BILLED
        # reconciliation block appended under the TOTAL row.
        recon_consumed_ids = set()
        recon_billed_total = 0.0

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

            # Collect all descriptions, years, months, sales orders and direct invoices
            descriptions = set()
            years = set()
            months = set()
            all_sales_orders = set()
            all_direct_invoices = set()

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

                # Collect all sales orders and direct invoices for this customer
                all_sales_orders.update(ce_data['sales_orders'])
                all_direct_invoices.update(ce_data.get('direct_invoices', set()))

            # Calculate billed amount for all sales orders of this customer
            if billed_mode == 'pnl':
                consumed = set()
                billed_amount = sum(
                    self._get_pnl_billed_for_row(ce_code, ces_data[ce_code], pnl_billed_map, consumed)
                    for ce_code in ces_data
                )
            else:
                billed_amount, _ = self._calculate_billed_amount(
                    all_sales_orders, report_month, all_direct_invoices,
                    consumed_move_ids=recon_consumed_ids)
            recon_billed_total += billed_amount

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

            # Write BILLED amount (net, after CM deduction)
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

            # Total formula (E+F+G+H+I+J = BILLED+accruals/reversals+ADDL ADJ)
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

        # Sum formulas for monetary columns (E=BILLED, F-J=accruals/reversals/adj, K=TOTAL)
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

        # BILLED reconciliation for every customer in this report.
        self._write_reconciliation_block(
            workbook, sheet, total_row, report_month,
            report_billed_total=recon_billed_total,
            consumed_move_ids=recon_consumed_ids,
            partner_ids=None,
            billed_mode=billed_mode,
            show_customer=True,
        )

        return True

    def _generate_customer_sheet(self, workbook, formats, partner_name, ces_data, report_month,
                                 reversal_ob_balances=None, billed_mode='standard', pnl_billed_map=None):
        """Generate individual customer sheet with CE breakdown"""
        if reversal_ob_balances is None:
            reversal_ob_balances = {}
        if pnl_billed_map is None:
            pnl_billed_map = {}

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
        sheet.set_column(9, 9, 18)  # Manual Accrual
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
            'CE#', 'SO REFERENCE', 'CE DATE', 'DESCRIPTION', 'Year', 'Month',
            'BILLED', 'SYSTEM ACCRUAL', 'SYSTEM REVERSAL', 'MANUAL ACCRUAL', 'MANUAL REVERSAL',
            'ADDL ADJ', 'TOTAL', 'CE STATUS', 'PER CSD', 'VARIANCE',
            f'COST TO CLIENT - {cost_to_client_month}', 'FOR REVENUE ADJUSTMENT', 'REMARKS'
        ]

        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['column_header'])

        sheet.autofilter(row, 0, row, len(headers) - 1)
        row += 1
        data_start_row = row

        # Accumulated while writing the rows below, then used by the BILLED
        # reconciliation block appended under this sheet's TOTAL row.
        recon_consumed_ids = set()
        recon_billed_total = 0.0

        # Write data rows
        for ce_code in sorted(ces_data.keys()):
            ce_data = ces_data[ce_code]
            amounts = self._calculate_amounts_by_type(
                ce_data['lines'], report_month)

            # Calculate billed amount for this CE row
            if billed_mode == 'pnl':
                billed_amount = self._get_pnl_billed_for_row(ce_code, ce_data, pnl_billed_map)
            else:
                billed_amount, _ = self._calculate_billed_amount(
                    ce_data['sales_orders'], report_month, ce_data.get('direct_invoices', set()),
                    consumed_move_ids=recon_consumed_ids)
            recon_billed_total += billed_amount

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

            # Write BILLED amount (col G, index 6) - net, after CM deduction
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

            # Total formula (G+H+I+J+K+L = BILLED+accruals/reversals+ADDL ADJ)
            excel_row = row + 1
            sheet.write_formula(
                row, 12, f'=G{excel_row}+H{excel_row}+I{excel_row}+J{excel_row}+K{excel_row}+L{excel_row}', formats['currency'])

            sheet.write(row, 13, ce_data['ce_status'], formats['centered'])

            # PER CSD - empty for user input
            sheet.write(row, 14, '', formats['currency'])

            # VARIANCE formula: Total - Per CSD (M - O)
            sheet.write_formula(
                row, 15, f'=M{excel_row}-O{excel_row}', formats['currency'])

            # COST TO CLIENT - empty for user input
            sheet.write(row, 16, '', formats['currency'])

            # FOR REVENUE ADJUSTMENT formula: Variance - Cost to Client (P - Q)
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

        # Sum formulas for monetary columns (G=BILLED, H-M=accruals/reversals/adj, M=TOTAL)
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

        # Resolve this sheet's customer to real partner ids so the reconciliation
        # control set is scoped to them.  Taken from the underlying records rather
        # than matching on the display name, which is uppercased and not unique.
        recon_partner_ids = set()
        for ce_data in ces_data.values():
            for line in ce_data.get('lines', []):
                if line.partner_id:
                    recon_partner_ids.add(line.partner_id.id)
            for so in self.env['sale.order'].sudo().browse(list(ce_data.get('sales_orders') or [])):
                if so.partner_id:
                    recon_partner_ids.add(so.partner_id.id)
            for inv in self.env['account.move'].sudo().browse(list(ce_data.get('direct_invoices') or [])):
                if inv.partner_id:
                    recon_partner_ids.add(inv.partner_id.id)

        # BILLED reconciliation for this customer only.
        self._write_reconciliation_block(
            workbook, sheet, total_row, report_month,
            report_billed_total=recon_billed_total,
            consumed_move_ids=recon_consumed_ids,
            partner_ids=recon_partner_ids or None,
            billed_mode=billed_mode,
            show_customer=False,
        )

        return True
