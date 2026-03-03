from odoo import models
import datetime
from xlsxwriter.workbook import Workbook
from odoo.exceptions import ValidationError, UserError
from dateutil.relativedelta import relativedelta
import pytz


class AgedReceivablesXLSX(models.AbstractModel):
    _name = 'report.aged_receivables_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    _description = 'Aged Receivables XLSX Report'

    def _define_formats(self, workbook):
        """Define and return format objects."""
        base_font = {'font_name': 'Calibri', 'font_size': 10}

        # Yellow header format
        yellow_header_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#FFFF00',  # Yellow
            'border': 1
        })

        # Normal cell format
        normal_format = workbook.add_format({
            **base_font,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })

        # Centered cell format
        centered_format = workbook.add_format({
            **base_font,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Right-aligned format for dash
        right_aligned_format = workbook.add_format({
            **base_font,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'indent': 1
        })

        # Date cell format
        date_format = workbook.add_format({
            **base_font,
            'num_format': 'mm/dd/yyyy',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Subtotal row format (no bold, just regular with border)
        subtotal_format = workbook.add_format({
            **base_font,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

        # Subtotal label format (lighter green background, bold, left-aligned)
        subtotal_label_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#CCFF66',  # Lighter green
            'border': 1
        })

        # Subtotal cell format (lighter green background)
        subtotal_cell_format = workbook.add_format({
            **base_font,
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#CCFF66',  # Lighter green
            'border': 1
        })

        # Grand total row format (darker green background, bold)
        grand_total_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#99CC33',  # Darker green
            'border': 1
        })

        # Grand total label format
        grand_total_label_format = workbook.add_format({
            **base_font,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#99CC33',  # Darker green
            'border': 1
        })

        return {
            'yellow_header': yellow_header_format,
            'normal': normal_format,
            'centered': centered_format,
            'right_aligned': right_aligned_format,
            'date': date_format,
            'subtotal': subtotal_format,
            'subtotal_label': subtotal_label_format,
            'subtotal_cell': subtotal_cell_format,
            'grand_total': grand_total_format,
            'grand_total_label': grand_total_label_format
        }

    def _get_aging_bucket(self, invoice_date, reference_date):
        """Determine which aging bucket (0-30, 31-60, 61-90, 91-120, OVER 120) the invoice falls into."""
        if not invoice_date:
            return None
        
        # Ensure we're working with date objects, not datetime
        if isinstance(invoice_date, datetime.datetime):
            invoice_date = invoice_date.date()
        if isinstance(reference_date, datetime.datetime):
            reference_date = reference_date.date()
        
        days_old = (reference_date - invoice_date).days
        
        # If not yet due (negative days_old), put in 0-30 bucket
        if days_old <= 0:
            return 0  # Not yet due - goes to 0-30
        elif days_old <= 30:
            return 0  # 0-30
        elif days_old <= 60:
            return 1  # 31-60
        elif days_old <= 90:
            return 2  # 61-90
        elif days_old <= 120:
            return 3  # 91-120
        else:
            return 4  # OVER 120

    def _convert_to_php(self, amount, from_currency, to_currency, date):
        """Convert amount from one currency to another (typically USD to PHP)."""
        if from_currency == to_currency:
            return amount
        
        # Use Odoo's currency conversion with the invoice date
        company = self.env.company
        converted_amount = from_currency._convert(
            amount, 
            to_currency, 
            company, 
            date
        )
        return converted_amount

    def _get_conversion_date(self, move):
        """Get the date to use for currency conversion (sales order date)."""
        # Get sale order
        sale_order = move.invoice_line_ids.mapped('sale_line_ids.order_id')[:1] if move.invoice_line_ids else False
        
        # Priority: x_studio_sales_order.order_date > sale_order.date_order > invoice_date > move.date
        if sale_order:
            # Try to get x_studio_sales_order if it exists
            if hasattr(sale_order, 'x_studio_sales_order') and sale_order.x_studio_sales_order:
                if hasattr(sale_order.x_studio_sales_order, 'order_date'):
                    return sale_order.x_studio_sales_order.order_date
            # Fallback to regular sale order date
            if sale_order.date_order:
                return sale_order.date_order.date() if isinstance(sale_order.date_order, datetime.datetime) else sale_order.date_order
        
        # Final fallback to invoice date or move date
        inv_date = move.invoice_date or move.date
        return inv_date.date() if isinstance(inv_date, datetime.datetime) else inv_date

    def _create_currency_format(self, workbook, currency_symbol, base_font):
        """Create currency format with symbol."""
        return workbook.add_format({
            **base_font,
            'num_format': f'{currency_symbol}#,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })

    def _create_subtotal_currency_format(self, workbook, currency_symbol, base_font):
        """Create subtotal currency format with lighter green background."""
        return workbook.add_format({
            **base_font,
            'num_format': f'{currency_symbol}#,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#CCFF66',  # Lighter green
            'border': 1
        })

    def _create_grand_total_currency_format(self, workbook, currency_symbol, base_font):
        """Create grand total currency format with darker green background."""
        return workbook.add_format({
            **base_font,
            'bold': True,
            'num_format': f'{currency_symbol}#,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#99CC33',  # Darker green
            'border': 1
        })

    def generate_xlsx_report(self, workbook, data, lines):
        """Main report generation method."""
        formats = self._define_formats(workbook)
        base_font = {'font_name': 'Calibri', 'font_size': 10}

        # Get reference date (today)
        reference_date = datetime.date.today()

        # Get PHP currency for grand total conversion
        php_currency = self.env['res.currency'].search([('name', '=', 'PHP')], limit=1)
        if not php_currency:
            php_currency = self.env.company.currency_id  # Fallback to company currency

        # Group invoices by partner only (not by currency)
        by_partner = {}

        for move in lines:
            if move.move_type in ['out_invoice', 'out_refund'] and move.state == 'posted' and move.amount_residual != 0:
                partner_name = move.partner_id.name or 'Unknown'

                if partner_name not in by_partner:
                    by_partner[partner_name] = []

                by_partner[partner_name].append(move)

        # Create single sheet for all partners
        sheet_name = f"AR Aging as of {reference_date.strftime('%B %d, %Y')}"
        sheet = workbook.add_worksheet(sheet_name[:31])

        # Set column widths
        sheet.set_column(0, 0, 12)   # PO#
        sheet.set_column(1, 1, 12)   # CE#
        sheet.set_column(2, 2, 40)   # CLIENT
        sheet.set_column(3, 3, 15)   # INVOICE #
        sheet.set_column(4, 4, 12)   # DATE
        sheet.set_column(5, 5, 15)   # FOREIGN AMOUNT
        sheet.set_column(6, 6, 15)   # AMOUNT (PHP)
        sheet.set_column(7, 7, 15)   # 0-30
        sheet.set_column(8, 8, 15)   # 31-60
        sheet.set_column(9, 9, 15)   # 61-90
        sheet.set_column(10, 10, 15) # 91-120
        sheet.set_column(11, 11, 15) # OVER 120

        # Write header row
        row = 0
        aging_labels = ['0-30', '31-60', '61-90', '91-120', 'OVER 120']
        
        sheet.write(row, 0, 'PO#', formats['yellow_header'])
        sheet.write(row, 1, 'CE#', formats['yellow_header'])
        sheet.write(row, 2, 'CLIENT', formats['yellow_header'])
        sheet.write(row, 3, 'INVOICE #', formats['yellow_header'])
        sheet.write(row, 4, 'DATE', formats['yellow_header'])
        sheet.write(row, 5, 'AMOUNT', formats['yellow_header'])
        sheet.write(row, 6, 'FOREIGN AMOUNT', formats['yellow_header'])
        
        for i, label in enumerate(aging_labels):
            sheet.write(row, 7 + i, label, formats['yellow_header'])

        sheet.set_row(row, 20)
        row += 1

        # Freeze the header row
        sheet.freeze_panes(1, 0)

        # Sort partners alphabetically
        sorted_partners = sorted(by_partner.items())

        # Initialize grand totals across all clients (in PHP)
        grand_totals = {
            'total': 0,
            'buckets': [0, 0, 0, 0, 0]
        }

        # Get PHP symbol for formatting
        php_symbol = php_currency.symbol if php_currency else '₱'
        
        # Create PHP currency formats
        currency_format = self._create_currency_format(workbook, php_symbol, base_font)
        subtotal_currency_format = self._create_subtotal_currency_format(workbook, php_symbol, base_font)
        grand_total_currency_format = self._create_grand_total_currency_format(workbook, php_symbol, base_font)
        
        # Dictionary to cache currency formats for foreign currencies
        foreign_currency_formats = {}

        # Process each partner
        for partner_name, moves in sorted_partners:
            # Sort moves by invoice name, then by invoice date
            moves_sorted = sorted(moves, key=lambda x: (x.name or '', x.invoice_date or x.date))

            # Initialize subtotals for this partner (in PHP)
            subtotals = {
                'total': 0,
                'buckets': [0, 0, 0, 0, 0]  # 0-30, 31-60, 61-90, 91-120, OVER 120
            }

            # Write data rows for this partner
            for move in moves_sorted:
                # Get original amount
                original_amount = move.amount_residual if move.move_type == 'out_invoice' else -move.amount_residual
                
                # Determine foreign amount (non-PHP) and PHP amount
                foreign_amount = None
                if move.currency_id != php_currency:
                    foreign_amount = original_amount
                    amount = move.x_alt_currency_amount if hasattr(move, 'x_alt_currency_amount') else original_amount
                else:
                    amount = original_amount
                
                subtotals['total'] += amount
                
                # Use invoice date for display
                inv_date = move.invoice_date or move.date or reference_date
                
                # Use due date for aging calculation
                due_date = move.invoice_date_due or inv_date
                
                # Add to grand total (already in PHP)
                grand_totals['total'] += amount

                # Determine aging bucket based on DUE date
                bucket_index = self._get_aging_bucket(due_date, reference_date)

                if bucket_index is not None:
                    subtotals['buckets'][bucket_index] += amount
                    grand_totals['buckets'][bucket_index] += amount

                # Get sale order for CE# and other fields
                sale_order = move.invoice_line_ids.mapped('sale_line_ids.order_id')[:1] if move.invoice_line_ids else False

                # Get CE# with fallback to old CE
                ce_code = ''
                if sale_order:
                    ce_code = sale_order.x_studio_old_ce or sale_order.x_ce_code or move.x_studio_old_ce_1 or ''
                else:
                    ce_code = move.x_studio_old_ce_1 or ''
                
                # Use old CE date if available, else fallback to invoice/move date
                # ce_date = inv_date
                # if sale_order and sale_order.x_studio_old_ce_date:
                    # ce_date = sale_order.x_studio_old_ce_date
                # elif sale_order:
                    # ce_date = sale_order.date_order or ce_date

                # Write row data
                sheet.write(row, 0, move.ref or '', formats['normal'])  # PO#
                sheet.write(row, 1, ce_code, formats['centered'])  # CE#
                sheet.write(row, 2, partner_name, formats['normal'])  # CLIENT
                sheet.write(row, 3, move.name or '', formats['centered'])  # INVOICE #
                # sheet.write(row, 4, ce_date, formats['date'])  # DATE
                sheet.write(row, 4, inv_date, formats['date'])  # DATE - change to actual invoice date
                
                sheet.write(row, 5, amount, currency_format)  # AMOUNT (in PHP)
                
                # Write foreign amount (only if not PHP)
                if foreign_amount is not None:
                    # Get currency symbol from the invoice's currency
                    currency_symbol = move.currency_id.symbol or move.currency_id.name
                    
                    # Create or retrieve cached format for this currency
                    if currency_symbol not in foreign_currency_formats:
                        foreign_currency_formats[currency_symbol] = self._create_currency_format(
                            workbook, currency_symbol, base_font)
                    
                    sheet.write(row, 6, foreign_amount, foreign_currency_formats[currency_symbol])  # FOREIGN AMOUNT
                else:
                    sheet.write(row, 6, '-', formats['right_aligned'])  # FOREIGN AMOUNT (dash if PHP)

                # Aging buckets - only populate the matching bucket
                for i in range(5):
                    if i == bucket_index:
                        sheet.write(row, 7 + i, amount, currency_format)
                    else:
                        sheet.write(row, 7 + i, '-', formats['right_aligned'])

                sheet.set_row(row, 18)
                row += 1

            # Write subtotal row for this partner
            sheet.merge_range(row, 0, row, 4, f'Total | {partner_name}', formats['subtotal_label'])
            sheet.write(row, 5, subtotals['total'], subtotal_currency_format)
            sheet.write(row, 6, '-', formats['subtotal_cell'])  # FOREIGN AMOUNT (dash)

            for i in range(5):
                sheet.write(row, 7 + i, subtotals['buckets'][i], subtotal_currency_format)

            sheet.set_row(row, 18)
            row += 1

        # GRAND TOTAL ROW
        # Write GRAND TOTAL row at the bottom (always in PHP)
        sheet.write(row, 0, 'Total Aged Receivable', formats['grand_total_label'])
        sheet.write(row, 1, '', formats['grand_total'])
        sheet.write(row, 2, '', formats['grand_total'])
        sheet.write(row, 3, '', formats['grand_total'])
        sheet.write(row, 4, '', formats['grand_total'])
        sheet.write(row, 5, grand_totals['total'], grand_total_currency_format)
        sheet.write(row, 6, '-', formats['grand_total'])  # FOREIGN AMOUNT (dash)

        for i in range(5):
            sheet.write(row, 7 + i, grand_totals['buckets'][i], grand_total_currency_format)

        sheet.set_row(row, 20)

        return True