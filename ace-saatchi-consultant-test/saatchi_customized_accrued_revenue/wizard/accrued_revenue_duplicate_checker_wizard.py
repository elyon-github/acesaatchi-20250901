# -*- coding: utf-8 -*-
"""
Accrued Revenue Wizard
======================
Wizard for creating accrued revenue entries with scenario support.

Features:
- Default mode: Auto-generate accruals for signed/billable SOs
- Special case mode: Three scenarios for flexible accrual creation
  * Scenario 1: Manual Accrue (override validation)
  * Scenario 2: Cancel & Replace existing accruals
  * Scenario 3: Create adjustment entries (NO auto-reversal)

Scenarios Explained:
- Scenario 1: Creates accruals bypassing CE status validation
- Scenario 2: Cancels existing accruals and replaces with new ones
- Scenario 3: Creates adjustment entries to reduce previous accruals
              (Dr. Digital Income | Cr. Accrued Revenue) - NO auto-reversal
"""

from odoo import models, fields, api, _
from odoo.exceptions import UserError
from dateutil.relativedelta import relativedelta
import logging

_logger = logging.getLogger(__name__)


class SaatchiAccruedRevenueWizard(models.TransientModel):
    """
    Accrued Revenue Wizard

    Main wizard model for batch accrual creation with scenario support.
    """
    _name = 'saatchi.accrued_revenue.wizard'
    _description = 'Accrued Revenue Creation Wizard'

    accrual_date = fields.Date(
        string="Accrual Date",
        required=True,
        default=lambda self: self._default_accrual_date(),
        help="Last day of the period for accrual (e.g., Oct 31, 2025)"
    )

    reversal_date = fields.Date(
        string="Reversal Date",
        required=True,
        default=lambda self: self._default_reversal_date(),
        help="First day of next period for reversal (e.g., Nov 1, 2025) - only used for normal accruals"
    )

    so_line_ids = fields.One2many(
        'saatchi.accrued_revenue.wizard.line',
        'wizard_id',
        string="Sale Orders"
    )

    has_existing_accruals = fields.Boolean(
        compute="_compute_has_existing_accruals",
        string="Has Duplicates",
        help="True if any sale order has existing accruals in this period"
    )

    special_case_mode = fields.Boolean(
        string="Special Case Mode",
        default=False,
        help="Enable scenario selection for advanced accrual handling"
    )

    accrual_scenario = fields.Selection(
        [
            ('scenario_1', 'Scenario 1: Manual Accrue (Override Validation)'),
            ('scenario_2', 'Scenario 2: Cancel & Replace Existing Accrue (Accrued/Reversed State Only)'),
            ('scenario_3', 'Scenario 3: Create Adjustment Entry (NO Auto-Reversal)')
        ],
        string="Accrual Scenario",
        default='scenario_1',
        help="""
        Scenario 1: Creates accruals bypassing CE status validation (allows any status)
        Scenario 2: Cancels existing accruals in 'Accrued' state and replaces them with new ones (Draft/Cancelled accruals are ignored)
        Scenario 3: Creates adjustment entries (Dr. Digital Income | Cr. Accrued Revenue) - PERMANENT entry with NO auto-reversal
        """
    )

    keep_foreign_currency = fields.Boolean(
        string="Keep Foreign Currency",
        default=False,
        help="When unchecked (default), foreign currency amounts (e.g. USD) are automatically converted "
             "to the company currency (PHP) using the Old CE Date or Order Date as the conversion date. "
             "When checked, amounts remain in the original foreign currency on the accrual record."
    )

    # ========== Onchange Methods ==========

    @api.onchange('accrual_date')
    def onchange_reversal_date(self):
        for record in self:
            if record.accrual_date:
                record.reversal_date = record.accrual_date + \
                    relativedelta(months=1, day=1)

    # ========== Compute Methods ==========

    @api.depends('so_line_ids.existing_accrual_ids')
    def _compute_has_existing_accruals(self):
        """Check if any wizard line has existing accruals"""
        for wizard in self:
            wizard.has_existing_accruals = any(
                line.existing_accrual_ids for line in wizard.so_line_ids
            )

    # ========== Default Methods ==========

    def action_refresh_so_lines(self):
        """
        Button action: Re-collect potential accruals and rebuild wizard lines.
        Uses direct DB operations (unlink + create) to avoid onchange/One2many conflicts.
        Returns an action to reload the wizard form.
        """
        self.ensure_one()
        if not self.accrual_date or not self.reversal_date:
            return

        # Update reversal_date based on new accrual_date
        self.reversal_date = self.accrual_date + relativedelta(months=1, day=1)

        potential_sos, duplicate_sos, client_sig_so_ids = self.env['sale.order'].collect_potential_accruals(
            accrual_date=self.accrual_date,
            reversal_date=self.reversal_date
        )

        # Delete existing lines from DB
        self.so_line_ids.unlink()

        # Create new lines directly in DB
        for so in potential_sos:
            amount_total = so._calculate_accrual_amount(
                accrual_date=self.accrual_date)
            if not amount_total:
                continue
            has_duplicate = so in duplicate_sos
            is_client_sig = so.id in client_sig_so_ids
            self.env['saatchi.accrued_revenue.wizard.line'].create({
                'wizard_id': self.id,
                'sale_order_id': so.id,
                'has_existing_accrual': has_duplicate,
                'amount_total': amount_total,
                'create_accrual': not has_duplicate,
                'is_from_reversal_ob': is_client_sig,
            })

        # Reload the wizard form to reflect changes
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'saatchi.accrued_revenue.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'views': [(False, 'form')],
            'target': 'new',
            'context': self.env.context,
        }

    def _default_accrual_date(self):
        """Default to last day of previous month"""
        today = fields.Date.context_today(self)
        first_of_current_month = today.replace(day=1)
        return first_of_current_month - relativedelta(days=1)

    def _default_reversal_date(self):
        """Default to first day of current month"""
        today = fields.Date.context_today(self)
        return today.replace(day=1)

    # ========== Action Methods ==========

    def action_create_accruals(self):
        """
        Create accruals for selected sale orders based on chosen scenario

        Process:
        1. Validate scenario compatibility
        2. Filter selected lines
        3. Execute scenario-specific logic
        4. Track success/failures
        5. Show notification

        Returns:
            dict: Client action to show notification
        """
        self.ensure_one()

        selected_lines = self.so_line_ids.filtered(lambda l: l.create_accrual)

        if not selected_lines:
            raise UserError(
                _('Please select at least one sale order to create accruals.'))

        # Validate scenario compatibility BEFORE processing
        if self.special_case_mode:
            validation_errors = self._validate_scenario_compatibility(
                selected_lines)
            if validation_errors:
                raise UserError('\n\n'.join(validation_errors))

        # Execute scenario with proper error handling
        try:
            if not self.special_case_mode:
                return self._execute_default_scenario(selected_lines)
            elif self.accrual_scenario == 'scenario_1':
                return self._execute_scenario_1(selected_lines)
            elif self.accrual_scenario == 'scenario_2':
                return self._execute_scenario_2(selected_lines)
            elif self.accrual_scenario == 'scenario_3':
                return self._execute_scenario_3(selected_lines)
        except Exception as e:
            _logger.exception("Fatal error in action_create_accruals")
            raise UserError(_(
                'An unexpected error occurred during accrual creation:\n%s\n\n'
                'Please check the logs for more details.'
            ) % str(e))

    def _validate_scenario_compatibility(self, selected_lines):
        """
        Validate that selected SOs are compatible with chosen scenario

        Args:
            selected_lines: Selected wizard lines

        Returns:
            list: Error messages (empty if valid)
        """
        errors = []

        # ========== UNIVERSAL VALIDATION: CE Code Required ==========
        sos_without_ce_code = []
        for line in selected_lines:
            if not line.sale_order_id.x_ce_code:
                sos_without_ce_code.append(line.sale_order_id.name)

        if sos_without_ce_code:
            so_list = '\n'.join(f'  • {so}' for so in sos_without_ce_code)
            errors.append(_(
                'CE Code Missing\n\n'
                'The following Sale Orders do not have a CE Code:\n\n'
                '%s\n\n'
                'Please make sure all SO have CE Code before proceeding.'
            ) % so_list)
            # Return early - no point checking other validations if CE codes are missing
            return errors

        # ========== SCENARIO-SPECIFIC VALIDATIONS ==========

        if self.accrual_scenario == 'scenario_1':
            # Filter lines where ANY existing accrual is draft or accrued
            with_existing = selected_lines.filtered(
                lambda l: any(a.state in ['accrued', 'draft', 'reversed']
                              for a in l.existing_accrual_ids)
            )

            if with_existing:
                so_list = '\n'.join(
                    f'  • {so}' for so in with_existing.mapped('sale_order_id.name'))
                errors.append(_(
                    'Scenario 1 Error\n\n'
                    'The following Sale Orders already have existing accruals:\n\n'
                    '%s\n\n'
                    'Scenario 1 can only be used for Sale Orders without any existing accruals.'
                ) % so_list)

        if self.accrual_scenario == 'scenario_2':
            # Scenario 2: Only SOs with existing accruals in 'accrued' state
            sos_without_accrued = []
            for line in selected_lines:
                accrued_records = line.existing_accrual_ids.filtered(
                    lambda a: a.state == 'reversed' or a.state == 'accrued')
                if not accrued_records:
                    sos_without_accrued.append(line.sale_order_id.name)

            if sos_without_accrued:
                so_list = '\n'.join(f'  • {so}' for so in sos_without_accrued)
                errors.append(_(
                    'Scenario 2 Error\n\n'
                    'The following Sale Orders have no posted accruals to replace:\n\n'
                    '%s\n\n'
                    'Scenario 2 requires existing accruals in "Accrued" state. Draft or cancelled accruals cannot be replaced.'
                ) % so_list)

        elif self.accrual_scenario == 'scenario_3':
            # Scenario 3: Only SOs with existing accruals (must have at least one "accrued" entry)
            sos_without_existing = []
            sos_without_accrued = []

            for line in selected_lines:
                # 1. No existing accruals at all
                if not line.existing_accrual_ids:
                    sos_without_existing.append(line.sale_order_id.name)
                    continue

                # 2. Has accruals but none are in "accrued" state
                accrued_records = line.existing_accrual_ids.filtered(
                    lambda a: a.state == 'reversed' or a.state == 'accrued')
                if not accrued_records:
                    sos_without_accrued.append(line.sale_order_id.name)

            # --- Error 1: SOs with NO accruals at all ---
            if sos_without_existing:
                so_list = '\n'.join(f'  • {so}' for so in sos_without_existing)
                errors.append(_(
                    'Scenario 3 Error\n\n'
                    'The following Sale Orders have no existing accruals:\n\n'
                    '%s\n\n'
                    'Scenario 3 creates adjustment entries and requires existing accruals to adjust.'
                ) % so_list)

            # --- Error 2: SOs with accruals but none in "accrued" state ---
            if sos_without_accrued:
                so_list = '\n'.join(f'  • {so}' for so in sos_without_accrued)
                errors.append(_(
                    'Scenario 3 Error\n\n'
                    'The following Sale Orders have no posted accruals:\n\n'
                    '%s\n\n'
                    'Scenario 3 requires accruals in "Accrued" state. Draft or cancelled accruals cannot be adjusted.'
                ) % so_list)

        return errors

    def _show_success_and_open_records(self, accrual_ids, title, summary_message):
        """
        Show success notification and open the list view of created records

        Args:
            accrual_ids: List of created accrual record IDs
            title: Notification title
            summary_message: Summary message for notification

        Returns:
            dict: Action to open list view with created records
        """
        if not accrual_ids:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': title,
                    'message': summary_message,
                    'type': 'warning',
                    'sticky': True,
                }
            }

        # Show notification and open records
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': title,
                'message': summary_message,
                'type': 'success',
                'sticky': False,
                'next': {
                    'type': 'ir.actions.act_window',
                    'name': _('Created Accrued Revenues'),
                    'res_model': 'saatchi.accrued_revenue',
                    'view_mode': 'list,form',
                    'views': [(False, 'list'), (False, 'form')],
                    'domain': [('id', 'in', accrual_ids)],
                    'target': 'current',
                }
            }
        }

    def _execute_default_scenario(self, selected_lines):
        """
        Default scenario: Create accruals for signed/billable SOs only.
        Also handles Client Signature SOs that matched reversal OB (marked as manual).

        Args:
            selected_lines: Wizard lines to process

        Returns:
            dict: Notification and navigation to created records
        """
        created_count = 0
        created_accrual_ids = []
        skipped_invalid_status = []
        skipped_no_lines = []
        failed_sos = []

        for line in selected_lines:
            so = line.sale_order_id

            # Check CE status - allow for_client_signature only if from reversal OB
            if so.state != 'sale' or so.x_ce_status not in ['signed', 'billable', 'for_client_signature']:
                skipped_invalid_status.append(
                    f"{so.name} (Status: {dict(so._fields['x_ce_status'].selection).get(so.x_ce_status, 'Unknown')})")
                continue

            # Client Signature SOs must be from reversal OB
            if so.x_ce_status == 'for_client_signature' and not line.is_from_reversal_ob:
                skipped_invalid_status.append(
                    f"{so.name} (Status: For Client Signature - not in Reversal OB)")
                continue

            # Client Signature from reversal OB -> mark as manual (not system generated)
            is_system = not line.is_from_reversal_ob

            try:
                result = so.action_create_custom_accrued_revenue(
                    is_override=False,
                    accrual_date=self.accrual_date,
                    reversal_date=self.reversal_date,
                    is_adjustment=False,
                    is_system_generated=is_system,
                    keep_foreign_currency=self.keep_foreign_currency
                )

                if result:
                    created_count += 1
                    created_accrual_ids.append(result)
                    _logger.info(
                        f"✓ Default: Created accrual {result} for SO {so.name}")
                else:
                    skipped_no_lines.append(so.name)
                    _logger.warning(
                        f"⚠ Default: No eligible lines for SO {so.name}")

            except Exception as e:
                error_msg = str(e)
                _logger.error(
                    f"❌ Default: Failed for SO {so.name}: {error_msg}", exc_info=True)
                failed_sos.append(f"{so.name} ({error_msg[:100]})")
                continue

        # Build summary message
        message_parts = [
            f'✅ Successfully created {created_count} accrual record(s)']

        if skipped_invalid_status:
            message_parts.append(
                f'\n⚠️ Skipped {len(skipped_invalid_status)} SO(s) due to invalid status'
            )

        if skipped_no_lines:
            message_parts.append(
                f'\n⚠️ Skipped {len(skipped_no_lines)} SO(s) - no eligible lines'
            )

        if failed_sos:
            message_parts.append(
                f'\n❌ Failed to create {len(failed_sos)} Accrual(s)'
            )

        summary = '\n'.join(
            message_parts) if created_count > 0 else 'No accruals were created. Please check the criteria. or Check the default accounts and journals in the configuration settings if set properly.'

        return self._show_success_and_open_records(
            created_accrual_ids,
            _('Accrual Creation Complete'),
            summary
        )

    def _execute_scenario_1(self, selected_lines):
        """
        Scenario 1: Manual Accrue (Override Validation)

        Creates accruals bypassing CE status validation.
        Allows accruals for any SO regardless of status.

        Args:
            selected_lines: Wizard lines to process

        Returns:
            dict: Notification and navigation to created records
        """
        created_count = 0
        created_accrual_ids = []
        failed_sos = []

        for line in selected_lines:
            try:
                result = line.sale_order_id.action_create_custom_accrued_revenue(
                    is_override=True,
                    accrual_date=self.accrual_date,
                    reversal_date=self.reversal_date,
                    is_adjustment=False,
                    is_system_generated=False,
                    keep_foreign_currency=self.keep_foreign_currency,
                )

                if result:
                    created_count += 1
                    created_accrual_ids.append(result)
                    _logger.info(
                        f"✓ Scenario 1: Created accrual {result} for SO {line.sale_order_id.name}")
                else:
                    error_msg = "No eligible lines found"
                    _logger.warning(
                        f"⚠ Scenario 1: {error_msg} for SO {line.sale_order_id.name}")
                    failed_sos.append(
                        f"{line.sale_order_id.name} ({error_msg})")

            except Exception as e:
                error_msg = str(e)
                _logger.error(
                    f"❌ Scenario 1: Failed for SO {line.sale_order_id.name}: {error_msg}", exc_info=True)
                failed_sos.append(
                    f"{line.sale_order_id.name} ({error_msg[:100]})")
                continue

        # Build summary message
        summary_parts = [
            f'✅ Created {created_count} Accrual(s) with override (Scenario 1)']

        if failed_sos:
            summary_parts.append(f'\n❌ Failed: {len(failed_sos)} SO(s)')

        summary = '\n'.join(
            summary_parts) if created_count > 0 else f'❌ Failed to create accruals for {len(failed_sos)} SO(s)'

        return self._show_success_and_open_records(
            created_accrual_ids,
            _('Scenario 1 Complete'),
            summary
        )

    def _execute_scenario_2(self, selected_lines):
        """
        Scenario 2: Cancel & Replace Existing

        Cancels existing accruals (ONLY in 'accrued' state) and creates new ones.
        Draft or cancelled accruals are skipped.

        Args:
            selected_lines: Wizard lines to process

        Returns:
            dict: Notification and navigation to created records
        """
        created_count = 0
        created_accrual_ids = []
        replaced_count = 0
        skipped_not_accrued = []
        failed_sos = []

        for line in selected_lines:
            try:
                # Filter to only 'accrued' state records
                accrued_records = line.existing_accrual_ids.filtered(
                    lambda a: a.state == 'accrued' or a.state == 'reversed')

                if not accrued_records:
                    skipped_not_accrued.append(
                        f"{line.sale_order_id.name} (No posted accruals found)")
                    _logger.warning(
                        f"⚠ Scenario 2: Skipped SO {line.sale_order_id.name} - no accrued state records to replace")
                    continue

                # Cancel only the 'accrued' state records
                for existing_accrual in accrued_records:
                    existing_accrual.action_reset_and_cancel()

                replaced_count += len(accrued_records)
                _logger.info(
                    f"✓ Scenario 2: Cancelled {len(accrued_records)} accrued record(s) for SO {line.sale_order_id.name}")

                # Create new accrual
                result = line.sale_order_id.action_create_custom_accrued_revenue(
                    is_override=False,
                    accrual_date=self.accrual_date,
                    reversal_date=self.reversal_date,
                    is_adjustment=False,
                    is_system_generated=False,
                    keep_foreign_currency=self.keep_foreign_currency
                )

                if result:
                    created_count += 1
                    created_accrual_ids.append(result)
                    _logger.info(
                        f"✓ Scenario 2: Created replacement accrual {result} for SO {line.sale_order_id.name}")
                else:
                    error_msg = "No eligible lines found"
                    _logger.warning(
                        f"⚠ Scenario 2: {error_msg} for SO {line.sale_order_id.name}")
                    failed_sos.append(
                        f"{line.sale_order_id.name} ({error_msg})")

            except Exception as e:
                error_msg = str(e)
                _logger.error(
                    f"❌ Scenario 2: Failed for SO {line.sale_order_id.name}: {error_msg}", exc_info=True)
                failed_sos.append(
                    f"{line.sale_order_id.name} ({error_msg[:100]})")
                continue

        # Build summary message
        summary_parts = [
            f'✅ Scenario 2: Cancelled {replaced_count} and created {created_count} Accrual(s)'
        ]

        if skipped_not_accrued:
            summary_parts.append(
                f'\nℹ️ Skipped {len(skipped_not_accrued)} SO(s) - no posted accruals')

        if failed_sos:
            summary_parts.append(f'\n❌ Failed: {len(failed_sos)} SO(s)')

        summary = '\n'.join(
            summary_parts) if created_count > 0 else 'No records were created'

        return self._show_success_and_open_records(
            created_accrual_ids,
            _('Scenario 2 Complete'),
            summary
        )

    def _execute_scenario_3(self, selected_lines):
        """
        Scenario 3: Create Adjustment Entry (NO Auto-Reversal)

        Creates adjustment entries to reduce previous accruals.
        Entry structure: Dr. Digital Income | Cr. Accrued Revenue

        This is a PERMANENT adjustment - NO automatic reversal is created.
        User can edit the amount in draft before posting.

        Only processes SOs with existing accruals.

        Args:
            selected_lines: Wizard lines to process

        Returns:
            dict: Notification and navigation to created records
        """
        created_count = 0
        created_accrual_ids = []
        skipped_count = 0
        failed_sos = []

        for line in selected_lines:
            try:
                # Only create adjustment if existing accruals exist
                if not line.existing_accrual_ids:
                    skipped_count += 1
                    _logger.warning(
                        f"⚠ Scenario 3: Skipped SO {line.sale_order_id.name} - no existing accruals")
                    continue

                # Create adjustment entry (NO auto-reversal)
                result = line.sale_order_id.action_create_custom_accrued_revenue(
                    is_override=False,
                    accrual_date=self.accrual_date,
                    reversal_date=False,
                    is_adjustment=True,
                    is_system_generated=False,
                    keep_foreign_currency=self.keep_foreign_currency
                )

                if result:
                    created_count += 1
                    created_accrual_ids.append(result)
                    _logger.info(
                        f"✓ Scenario 3: Created adjustment entry {result} for SO {line.sale_order_id.name}")
                else:
                    error_msg = "No accrual amount calculated"
                    _logger.warning(
                        f"⚠ Scenario 3: {error_msg} for SO {line.sale_order_id.name}")
                    failed_sos.append(
                        f"{line.sale_order_id.name} ({error_msg})")

            except Exception as e:
                error_msg = str(e)
                _logger.error(
                    f"❌ Scenario 3: Failed for SO {line.sale_order_id.name}: {error_msg}", exc_info=True)
                failed_sos.append(
                    f"{line.sale_order_id.name} ({error_msg[:100]})")
                continue

        # Build summary message
        summary_parts = [
            f'✅ Created {created_count} adjustment entr(y/ies) in DRAFT',
            '⚠️ Remember: Edit "Digital Income" amount before posting (NO auto-reversal)'
        ]

        if skipped_count > 0:
            summary_parts.append(
                f'\nℹ️ Skipped {skipped_count} SO(s) - no existing accruals')

        if failed_sos:
            summary_parts.append(f'\n❌ Failed: {len(failed_sos)} SO(s)')

        summary = '\n'.join(
            summary_parts) if created_count > 0 else 'No adjustment entries were created'

        return self._show_success_and_open_records(
            created_accrual_ids,
            _('Scenario 3 Complete'),
            summary
        )


class SaatchiAccruedRevenueWizardLine(models.TransientModel):
    """
    Accrued Revenue Wizard Line

    Individual line for each sale order in the wizard.
    """
    _name = 'saatchi.accrued_revenue.wizard.line'
    _description = 'Accrued Revenue Wizard Line'
    _order = 'create_accrual desc, has_existing_accrual desc, sale_order_id'

    wizard_id = fields.Many2one(
        'saatchi.accrued_revenue.wizard',
        string="Wizard",
        required=True,
        ondelete='cascade',
        index=True
    )

    sale_order_id = fields.Many2one(
        'sale.order',
        string="Sale Order",
        required=True,
        readonly=True,
        index=True
    )

    partner_id = fields.Many2one(
        'res.partner',
        string="Customer",
        related='sale_order_id.partner_id',
        readonly=True
    )

    ce_code = fields.Char(
        string="CE Code",
        related='sale_order_id.x_ce_code',
        readonly=True,
    )

    old_ce_code = fields.Char(
        string="Old CE #",
        related='sale_order_id.x_studio_old_ce',
        readonly=True,
    )

    old_ce_date = fields.Date(
        string="Old CE Date",
        related='sale_order_id.x_studio_old_ce_date',
        readonly=True,
    )

    create_date = fields.Datetime(
        string="Created",
        related='sale_order_id.create_date',
        readonly=True,
    )

    ce_status = fields.Selection(
        string="CE Status",
        related='sale_order_id.x_ce_status',
        readonly=True,
    )

    amount_total = fields.Monetary(
        string="Accrual Amount",
        readonly=True,
        help="Total amount to be accrued for this sale order"
    )

    currency_id = fields.Many2one(
        'res.currency',
        related='sale_order_id.currency_id',
        readonly=True
    )

    has_existing_accrual = fields.Boolean(
        string="Has Existing Accrual",
        default=False,
        readonly=True,
        help="This sale order already has an accrual for this period",
        compute="_compute_existing_accrual_total",
        store=True
    )

    is_from_reversal_ob = fields.Boolean(
        string="From Reversal OB",
        default=False,
        readonly=True,
        help="This SO is a Client Signature SO included because its old CE code "
             "was found in the Accrual Reversal Opening Balances. "
             "Will be marked as manual (not system generated)."
    )

    existing_accrual_ids = fields.Many2many(
        'saatchi.accrued_revenue',
        string="Existing Accruals",
        compute="_compute_existing_accruals",
        store=True,
        help="Existing accrual records for this sale order in the current period"
    )

    existing_accrual_total = fields.Monetary(
        string="Existing Accrual Total",
        compute="_compute_existing_accrual_total",
        currency_field="currency_id",
        help="Total amount from existing accruals"
    )

    create_accrual = fields.Boolean(
        string="Select",
        default=False,
        help="Check to create accrual for this sale order",
        store=True,
        compute="_compute_create_accrual"
    )

    # ========== Compute Methods ==========

    @api.depends('sale_order_id', 'wizard_id.accrual_date', 'wizard_id.reversal_date')
    def _compute_existing_accruals(self):
        """Find existing accruals for this sale order in the current period"""
        for line in self:
            if line.sale_order_id and line.wizard_id.accrual_date and line.wizard_id.reversal_date:
                existing = self.env['saatchi.accrued_revenue'].search([
                    ('x_related_ce_id', '=', line.sale_order_id.id),
                    ('date', '>=', line.wizard_id.accrual_date),
                    ('date', '<=', line.wizard_id.reversal_date),
                    ('state', 'in', ['draft', 'accrued', 'reversed'])
                ])
                line.existing_accrual_ids = [(6, 0, existing.ids)]
            else:
                line.existing_accrual_ids = [(5, 0, 0)]

    @api.depends('existing_accrual_ids', 'existing_accrual_ids.total_debit_in_accrue_account')
    def _compute_existing_accrual_total(self):
        """Calculate total from existing accruals"""
        for line in self:
            line.existing_accrual_total = sum(
                line.existing_accrual_ids.mapped(
                    'total_debit_in_accrue_account')
            )

            if line.existing_accrual_total:
                # line.create_accrual = False
                line.has_existing_accrual = True
            else:
                # line.create_accrual = True
                line.has_existing_accrual = False

    @api.depends('wizard_id.accrual_date')
    def _compute_create_accrual(self):
        for record in self:
            if record.existing_accrual_ids:
                record.create_accrual = False
            else:
                record.create_accrual = True
