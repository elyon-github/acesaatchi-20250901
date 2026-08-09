from odoo import models, fields, api
from odoo.exceptions import UserError
from dateutil.relativedelta import relativedelta


class SalesOrderRevenueWizard(models.TransientModel):
    _name = 'sales.order.revenue_report.wizard'
    _description = 'Sales Order Revenue Report Wizard'

    partner_ids = fields.Many2many(
        'res.partner',
        string='Customers',
        help='Select specific customers. Leave empty to include all customers.'
    )

    report_date = fields.Date(
        string='Report Month',
        required=True,
        default=lambda self: fields.Date.context_today(
            self) - relativedelta(months=1),
        help='Select any date in the month you want to report on'
    )

    billed_mode = fields.Selection(
        selection=[
            ('standard', 'Standard'),
            ('pnl', 'PnL Style (includes CMs linked to the month)'),
        ],
        help="Select the mode for calculating billed column.",
        string='Billed Calculation Mode',
        required=True,
        default='standard',
    )

    def action_print_report(self):
        """Generate the Sales Order Revenue XLSX report"""
        self.ensure_one()

        # Get the accrued revenue account ID
        accrued_account_ids = self._get_accrued_revenue_account_id()
        if not accrued_account_ids:
            raise UserError(
                "Accrued Revenue account not configured. "
                "Please set it in system parameters (account.accrued_revenue_account_id)."
            )

        # Calculate month range
        month_start = self.report_date.replace(day=1)
        month_end = (month_start + relativedelta(months=1)) - \
            relativedelta(days=1)

        # Build domain for account.move.line search (accrued entries)
        domain = [
            ('account_id', 'in', accrued_account_ids),
            ('parent_state', '=', 'posted'),
            ('x_sales_order', '!=', False),  # Must have sales order link
        ]

        # Add partner filter if specific customers selected
        if self.partner_ids:
            domain.append(('partner_id', 'in', self.partner_ids.ids))

        # Fetch all relevant accrued revenue move lines
        accrued_lines = self.env['account.move.line'].search(domain)

        # Get Sales Orders that were billed in the report month
        invoice_domain = [
            ('move_type', '=', 'out_invoice'),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', month_start),
            ('invoice_date', '<=', month_end),
        ]

        if self.partner_ids:
            invoice_domain.append(('partner_id', 'in', self.partner_ids.ids))

        invoices = self.env['account.move'].search(invoice_domain)

        # Get ALL sales orders from these invoices (whether they have accrued entries or not)
        all_billed_sales_orders = invoices.mapped(
            'invoice_line_ids.sale_line_ids.order_id')

        # Also pick up SOs that are not directly linked to any invoice but whose
        # x_studio_old_ce matches x_studio_old_ce_1 on one of the month's invoices.
        old_ce_values = [
            v for v in invoices.mapped('x_studio_old_ce_1') if v
        ]
        matched_so_old_ces = set()
        if old_ce_values:
            try:
                unlinked_sos = self.env['sale.order'].sudo().search([
                    ('x_studio_old_ce', 'in', old_ce_values),
                    ('company_id', 'in', self.env.companies.ids),
                ])
                all_billed_sales_orders |= unlinked_sos
                matched_so_old_ces = set(
                    unlinked_sos.mapped('x_studio_old_ce'))
            except Exception:
                pass

        # Standalone invoices: have x_studio_old_ce_1 set but no matching SO at all.
        # These still need to appear in the billed column even without a SO.
        so_linked_inv_ids = set(
            all_billed_sales_orders.mapped('invoice_ids').ids
        )
        standalone_invoice_ids = []
        for inv in invoices:
            inv_old_ce = getattr(inv, 'x_studio_old_ce_1', '') or ''
            if not inv_old_ce:
                continue
            if inv_old_ce in matched_so_old_ces:
                continue
            if inv.id in so_linked_inv_ids:
                continue
            standalone_invoice_ids.append(inv.id)
        # PnL mode: also collect SOs linked to CMs in the month so CE rows exist
        # even when the related invoice falls outside the current month range.
        if self.billed_mode == 'pnl':
            cm_domain = [
                ('move_type', '=', 'out_refund'),
                ('state', '=', 'posted'),
                ('invoice_date', '>=', month_start),
                ('invoice_date', '<=', month_end),
            ]
            if self.partner_ids:
                cm_domain.append(('partner_id', 'in', self.partner_ids.ids))
            try:
                cms = self.env['account.move'].sudo().search(cm_domain)
            except Exception:
                cms = self.env['account.move']

            # 1. CMs with x_studio_old_ce_1 → find matching SO by old CE or x_ce_code
            cm_ce_values = [v for v in cms.mapped('x_studio_old_ce_1') if v]
            if cm_ce_values:
                try:
                    cm_sos = self.env['sale.order'].sudo().search([
                        '|',
                        ('x_studio_old_ce', 'in', cm_ce_values),
                        ('x_ce_code', 'in', cm_ce_values),
                        ('company_id', 'in', self.env.companies.ids),
                    ])
                    all_billed_sales_orders |= cm_sos
                except Exception:
                    pass

            # 2. CMs without x_studio_old_ce_1 → walk reconciliation to the invoice,
            #    then collect that invoice's SOs (handles prior-month invoices).
            for cm in cms:
                try:
                    receivable = cm.line_ids.filtered(
                        lambda l: l.account_id.account_type == 'asset_receivable'
                    )
                    for line in receivable:
                        for partial in line.matched_debit_ids:
                            try:
                                inv = partial.debit_move_id.move_id
                                if not inv or inv.move_type != 'out_invoice':
                                    continue
                                inv_sos = inv.invoice_line_ids.mapped(
                                    'sale_line_ids.order_id')
                                all_billed_sales_orders |= inv_sos
                                # Also match via invoice's x_studio_old_ce_1
                                inv_ce1 = getattr(
                                    inv, 'x_studio_old_ce_1', '') or ''
                                if inv_ce1:
                                    ce1_sos = self.env['sale.order'].sudo().search([
                                        '|',
                                        ('x_studio_old_ce', '=', inv_ce1),
                                        ('x_ce_code', '=', inv_ce1),
                                        ('company_id', 'in',
                                         self.env.companies.ids),
                                    ])
                                    all_billed_sales_orders |= ce1_sos
                            except Exception:
                                continue
                except Exception:
                    continue

        # raise UserError(accrued_account_ids)
        if not accrued_lines and not all_billed_sales_orders:
            if self.partner_ids:
                customer_names = ', '.join(self.partner_ids.mapped('name'))
                raise UserError(
                    f"No accrued revenue entries or billed sales orders found for selected customer(s): {customer_names}"
                )
            else:
                raise UserError(
                    "No accrued revenue entries or billed sales orders found.")

        # Prepare data to pass to report
        report_data = {
            'report_date': self.report_date.isoformat(),
            'partner_ids': self.partner_ids.ids if self.partner_ids else [],
            'move_line_ids': accrued_lines.ids,
            'all_billed_so_ids': all_billed_sales_orders.ids,
            'standalone_invoice_ids': standalone_invoice_ids,
            'billed_mode': self.billed_mode,
        }

        # Return the report action
        return self.env.ref('saatchi_soa.action_report_sales_order_revenue_xlsx').report_action(
            self.env['account.move.line'],
            data=report_data
        )

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
