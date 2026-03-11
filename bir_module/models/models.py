# -*- coding: utf-8 -*-

import string
from odoo import models, fields, api
from datetime import datetime
import calendar
import psycopg2
import json
import os
import xlwt
import csv
from odoo.exceptions import UserError


class bir_module(models.Model):

    _name = 'bir_module.print_history'
    _description = 'bir_module.print_history'

    form_type = fields.Char(string='BIR Form Type')
    report_type = fields.Char()
    print_date = fields.Char()
    quarter_scope = fields.Char()


class print_history_line(models.Model):
    _name = 'bir_module.print_history_line'
    _description = 'bir_module.print_history_line'

    print_id = fields.Many2one('bir_module.print_history')
    move_id = fields.Many2one('account.move')
    scope = fields.Char()
    form_type = fields.Char(string='BIR Form Type')


class atc_setup(models.Model):
    _name = 'bir_module.atc_setup'
    _description = 'bir_module.atc_setup'

    name = fields.Char(required=True)
    tax_id = fields.Many2one('account.tax', required=True)
    company_id = fields.Many2one('res.company', string='Company', required=True, 
                                  default=lambda self: self.env.company, readonly=True)
    description = fields.Char()
    scope = fields.Selection(
        [('sales', 'Sales'), ('purchase', 'Purchases')], required=True)
    remarks = fields.Char()
    atc_code_company = fields.Char(string='ATC Code - Company (WC)', help='ATC code for corporate/company vendors')
    atc_code_individual = fields.Char(string='ATC Code - Individual (WI)', help='ATC code for individual vendors')

    @api.onchange('tax_id')
    def _onchange_tax_id(self):
        """Auto-fill company based on the selected tax"""
        if self.tax_id:
            self.company_id = self.tax_id.company_id


class bir_add_partner_field(models.Model):
    _name = 'account.move.line'
    _inherit = 'account.move.line'

    service_provider = fields.Many2one(
        'res.partner', string="Service Provider")


class AccountMoveLine(models.Model):
    _inherit = 'account.move.line'

    exclude_from_invoice_tab = fields.Boolean(
        string='Exclude from Invoice Tab',
        default=False
    )


class bir_reports(models.Model):
    _name = 'account.move'
    _inherit = 'account.move'

    test_field = fields.Char()

    @api.model
    def test(self):
        pass

    @api.model
    def _parse_checked_ids(self, checked_ids_param):
        """Parse checked IDs from JSON string and ensure list of ints
        
        Args:
            checked_ids_param: JSON string or list of selected invoice IDs from frontend
            
        Returns:
            List of integer move IDs, or empty list if parsing fails
        """
        import json
        try:
            if isinstance(checked_ids_param, str):
                if checked_ids_param == '[]' or not checked_ids_param:
                    return []
                loaded = json.loads(checked_ids_param)
            elif isinstance(checked_ids_param, list):
                loaded = checked_ids_param
            else:
                return []
            # Ensure all are ints and valid
            return [int(i) for i in loaded if str(i).isdigit()]
        except Exception:
            return []

##############################################################################################################################################################################
################################################################ 2307 ########################################################################################################
##############################################################################################################################################################################


    def x_2307_forms(self, args):
        import json
        from urllib.parse import urlencode
        import logging
        logger = logging.getLogger(__name__)
        
        search = args.get('search', '')
        checked_ids = args.get('checked_ids', [])
        signee_id = args.get('signee_id', 0)
        # Get company_id from args (passed from frontend) or use user's current company
        company_id = args.get('company_id', self.env.user.company_id.id)
        
        # Debug logging
        logger.warning(f"DEBUG x_2307_forms: args company_id={args.get('company_id', 'NOT PROVIDED')}, fallback company_id={self.env.user.company_id.id}, final company_id={company_id}")
        
        # Support both old 'month' and new 'from_date'/'to_date' parameters
        from_date = args.get('from_date', '')
        to_date = args.get('to_date', '')
        month = args.get('month', '')
        
        # Build URL with search, checked_ids, and company_id parameters
        url_params = {
            'id': args['id'],
            'trigger': args['trigger'],
            'tranid': args['tranid'],
            'search': search,
            'checked_ids': json.dumps(checked_ids) if checked_ids else '[]',
            'signee_id': signee_id,
            'company_id': company_id
        }
        
        # Add date parameters (either date range or legacy month)
        if from_date and to_date:
            url_params['from_date'] = from_date
            url_params['to_date'] = to_date
        elif month:
            url_params['month'] = month
        
        url = f"/report/pdf/bir_module.form_2307/?{urlencode(url_params)}"
        
        logger.warning(f"DEBUG URL: {url}")
        
        return {
            'type': 'ir.actions.act_url',
            'url': url,
            'target': 'new',
        }

    def x_get_2307_data(self, args):
        data = []
        search = args[5] if len(args) > 5 else ""
        checked_ids = args[6] if len(args) > 6 else []
        # Extract company_id if provided in args (7th element), otherwise use env.company
        company_id = args[7] if len(args) > 7 else self.env.company.id

        if args[2] == 'reprint':
            transactional = self._2307_query_reprint(args, company_id)
            data.append(transactional)
        # elif args[2] == 'ammend' or args[2] == 'ammend_view':
        #     transactional = self._2307_query_ammend(args)
        #     data.append(transactional)
        else:
            transactional = self._2307_query_normal(args, search, checked_ids, company_id)
            data.append(transactional)

        if args[2] == 'table':
            init = self.process_2307_ammend(data)

            data = init

        if args[2] == 'ammend':
            self.record_bir_form_print(data[0], '2307', args[3], args[0][1])

        return data

    def _2307_query_reprint(self, args, company_id=None):
        if company_id is None:
            company_id = self.env.company.id
            
        query = """ SELECT Abs(T1.price_subtotal)*(Abs(T3.amount)/100), T1.price_subtotal, T5.name, T5.vat, T4.name, T3.name,
            T0.id, T0.move_type, T0.name, amount_total, amount_untaxed, T0.invoice_date, T5.id  
            FROM bir_module_print_history_line T6 
            JOIN account_move T0 ON T0.id = T6.move_id  
            JOIN account_move_line T1 ON T0.id = T1.move_id  
            JOIN account_move_line_account_tax_rel T2 ON T1.id = T2.account_move_line_id 
            JOIN account_tax T3 ON T2.account_tax_id = T3.id 
            JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id 
            JOIN res_partner T5 ON T0.partner_id = T5.id 
            WHERE T0.state='posted' AND T6.print_id = '{0}' AND T0.company_id = {1}"""

        self._cr.execute(query.format(args[4], company_id))
        val = self._cr.fetchall()

        return val

    # def _2307_query_ammend(self, args):
    #     ids = args[5]
    #     if type(args[5]) == str:
    #         ids = json.loads(args[5])

    #     query = """SELECT Abs(T1.price_total)*(Abs(T3.amount)/100), T1.price_total, T5.name, T5.vat, T4.name, T3.name,
    #         T0.id, T0.move_type, T0.name, T0.amount_total, T0.invoice_date, T6.id
    #         FROM account_move T0
    #         JOIN account_move_line T1 ON T0.id = T1.move_id
    #         JOIN account_move_line_account_tax_rel T2 ON T1.id = T2.account_move_line_id
    #         JOIN account_tax T3 ON T2.account_tax_id = T3.id
    #         JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id
    #         JOIN res_partner T5 ON T0.partner_id = T5.id
    #         LEFT JOIN bir_module_print_history_line T6 ON T6.move_id = T0.id AND T6.form_type = '2307'
    #         WHERE """

    #     ctr = 1
    #     for id in ids:
    #         if len(ids) == ctr:
    #             query += "T0.id = " + str(id)
    #         else:
    #             query += "T0.id = " + str(id) + " OR "

    #         ctr += 1

    #     self._cr.execute(query)
    #     val = self._cr.fetchall()

    #     return val

    def _2307_query_normal(self, args, search="", checked_ids=[], company_id=None):
        """Build and execute query for BIR 2307 form data
        
        Filters by:
        - Posted vendor bills/invoices
        - Selected partner
        - Search term (bill name/number)
        - Checkbox-selected records (if any provided)
        - Company ID (if provided)
        
        Args:
            args: [partner_id, month] from form
            search: Search term to filter bills by name
            checked_ids: List of selected move IDs from checkboxes (empty = all records)
            company_id: Company ID to filter invoices by
            
        Returns:
            List of tuples containing invoice data for processing
        """
        if company_id is None:
            company_id = self.env.company.id
            
        query = """ SELECT T0.id, T0.amount_untaxed, T5.name, T5.vat, 
            CASE WHEN T5.is_company THEN T4.atc_code_company ELSE T4.atc_code_individual END, T4.description,
            T0.id, T0.move_type, T0.name, T0.amount_total, T0.amount_untaxed, T0.invoice_date, T0.invoice_date_due, T0.payment_state, T3.id
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id  
            JOIN account_move_line_account_tax_rel T2 ON T1.id = T2.account_move_line_id 
            JOIN account_tax T3 ON T2.account_tax_id = T3.id 
            JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id 
            JOIN res_partner T5 ON T0.partner_id = T5.id 
            {3} 
            WHERE T0.state='posted' AND T0.company_id = {0} AND T0.move_type = 'in_invoice' {1}
            GROUP BY T0.id, T0.amount_untaxed, T5.name, T5.vat, T5.is_company, T4.atc_code_company, T4.atc_code_individual, T4.description, T0.move_type, T0.name, T0.amount_total, T0.invoice_date, T0.invoice_date_due, T0.payment_state, T3.id"""

        end_parameter = self._2307_params(trans=args[1], id=args[0], search=search, checked_ids=checked_ids)

        self._cr.execute(query.format(company_id,
                         end_parameter[0], end_parameter[1], end_parameter[2]))
        val = self._cr.fetchall()

        return val

    def _2307_params(self, **kwargs):
        """Build WHERE clause parameters for BIR 2307 query
        
        Handles:
        - Partner filtering
        - Date range filtering (new) or quarter filtering (legacy)
        - Search term filtering
        - Checkbox-selected records filtering
        
        Returns:
            [param_string, field_string, join_string] for query construction
        """
        param = ""
        field = ""
        join = ""
        search = kwargs.get('search', '')
        checked_ids = kwargs.get('checked_ids', [])

        if kwargs['trans'] == "transactional":
            param = " AND T0.id = " + str(kwargs['id'][0])
        else:
            # Extract parameters - could be [partner_id, from_date, to_date] or [partner_id, month, ""]
            partner_id = kwargs['id'][0]
            param_1 = kwargs['id'][1]  # from_date or month
            param_2 = kwargs['id'][2] if len(kwargs['id']) > 2 else ""  # to_date or empty
            
            param = " AND T0.partner_id = " + str(partner_id) + " AND T6.id IS NULL AND "
            
            # Check if we have date range (from_date and to_date)
            if param_1 and param_2 and '-' in param_1 and '-' in param_2:
                # New date range filtering
                param += f" T0.invoice_date >= '{param_1}' AND T0.invoice_date <= '{param_2}'"
            elif param_1:
                # Legacy month filtering
                parameter = param_1.replace("-", " ").split()
                span = self.check_quarter_2307(int(parameter[1]))
                param += self.sawt_map_params(span, parameter[0])

            field = ", T0.invoice_date, T6.id "
            join = "LEFT JOIN bir_module_print_history_line T6 ON T6.move_id = T0.id AND T6.form_type = '2307'"

        # Add search filter for bill name if provided
        if search:
            param += " AND T0.name ILIKE '%" + search + "%'"

        # Add filter for checked IDs if any are selected
        # When user selects checkboxes, only those records are included in the report
        # When no checkboxes are selected, all records are included (default behavior)
        if checked_ids and len(checked_ids) > 0:
            # Ensure all are ints and valid
            valid_ids = [str(int(id)) for id in checked_ids if str(id).isdigit()]
            if valid_ids:
                checked_ids_str = ",".join(valid_ids)
                param += f" AND T0.id IN ({checked_ids_str})"

        return [param, field, join]

    def process_2307_quarterly(self, data, from_date=None, to_date=None):
        """
        Process 2307 quarterly data. 
        Fetches tax_totals via ORM to find WHT amount for the specific tax linked in ATC setup.
        """
        import json
        
        # Group data by ATC code and move_id (vendor bill)
        bill_groups = {}
        move_ids = set()
        
        for dat in data:
            move_ids.add(dat[6])  # T0.id - invoice/bill ID
        
        # Fetch all move records via ORM to get tax_totals field
        moves = self.env['account.move'].browse(list(move_ids))
        move_tax_totals = {move.id: move.tax_totals for move in moves}
        
        for dat in data:
            atc = dat[4]  # ATC code
            move_id = dat[6]  # Invoice/bill ID (T0.id)
            untaxed_amount = dat[1]  # amount_untaxed
            tax_id_in_atc = dat[14]  # T3.id - the tax_id from ATC setup
            invoice_date = dat[11]
            desc = dat[5]
            code = dat[4]
            
            # Get tax_totals from ORM
            tax_totals = move_tax_totals.get(move_id)
            wht_amount = 0
            
            if tax_totals:
                try:
                    if isinstance(tax_totals, str):
                        tax_totals = json.loads(tax_totals)
                    
                    # Search through tax groups to find the one matching our tax_id
                    if 'subtotals' in tax_totals and len(tax_totals['subtotals']) > 0:
                        for subtotal in tax_totals['subtotals']:
                            if 'tax_groups' in subtotal:
                                for tax_group in subtotal['tax_groups']:
                                    # Check if this tax group's involved_tax_ids contains our tax_id
                                    if 'involved_tax_ids' in tax_group:
                                        if tax_id_in_atc in tax_group['involved_tax_ids']:
                                            wht_amount = abs(tax_group.get('tax_amount', 0))
                                            break
                except Exception as e:
                    wht_amount = 0
            
            if atc not in bill_groups:
                bill_groups[atc] = {}
            
            if move_id not in bill_groups[atc]:
                bill_groups[atc][move_id] = {
                    'wht': wht_amount,
                    'untaxed': untaxed_amount,
                    'date': invoice_date,
                    'desc': desc,
                    'code': code
                }
        
        # Now process by ATC and sum WHT amounts
        final = []
        
        for atc, bills in bill_groups.items():
            dict = {'desc': '', 'code': '', 'm1': 0,
                    'm2': 0, 'm3': 0, 'taxed': 0, 'm_total': 0}
            
            for move_id, bill_data in bills.items():
                # Use date range mapping if provided, otherwise use legacy method
                if from_date and to_date:
                    month_num = self.get_month_position_for_date(bill_data['date'], from_date, to_date)
                else:
                    month_num = self.get_bir_month_num(bill_data['date'])
                
                # Only include invoices that are within the selected date range
                if month_num != '0':  # '0' means outside selected range
                    if month_num == '1':
                        dict['m1'] += bill_data['untaxed']
                    elif month_num == '2':
                        dict['m2'] += bill_data['untaxed']
                    elif month_num == '3':
                        dict['m3'] += bill_data['untaxed']
                    
                    # Use actual WHT from tax_totals
                    dict['taxed'] += bill_data['wht']
                    dict['m_total'] += bill_data['untaxed']
            
            dict['desc'] = bill_data['desc']
            dict['code'] = bill_data['code']
            final.append(dict)
        
        return final

    def process_2307_transactional(self, data):
        """
        Process 2307 transactional data.
        Fetches tax_totals via ORM to find WHT amount for the specific tax linked in ATC setup.
        """
        import json
        
        # Group data by ATC code and move_id (vendor bill)
        bill_groups = {}
        move_ids = set()
        
        for dat in data[0]:
            move_ids.add(dat[6])  # T0.id - invoice/bill ID
        
        # Fetch all move records via ORM to get tax_totals field
        moves = self.env['account.move'].browse(list(move_ids))
        move_tax_totals = {move.id: move.tax_totals for move in moves}
        
        for dat in data[0]:
            atc = dat[4]  # ATC code
            move_id = dat[6]  # Invoice/bill ID (T0.id)
            untaxed_amount = dat[1]  # amount_untaxed
            tax_id_in_atc = dat[14]  # T3.id - the tax_id from ATC setup
            desc = dat[5]
            code = dat[4]
            
            # Get tax_totals from ORM
            tax_totals = move_tax_totals.get(move_id)
            wht_amount = 0
            
            if tax_totals:
                try:
                    if isinstance(tax_totals, str):
                        tax_totals = json.loads(tax_totals)
                    
                    # Search through tax groups to find the one matching our tax_id
                    if 'subtotals' in tax_totals and len(tax_totals['subtotals']) > 0:
                        for subtotal in tax_totals['subtotals']:
                            if 'tax_groups' in subtotal:
                                for tax_group in subtotal['tax_groups']:
                                    # Check if this tax group's involved_tax_ids contains our tax_id
                                    if 'involved_tax_ids' in tax_group:
                                        if tax_id_in_atc in tax_group['involved_tax_ids']:
                                            wht_amount = abs(tax_group.get('tax_amount', 0))
                                            break
                except Exception as e:
                    wht_amount = 0
            
            if atc not in bill_groups:
                bill_groups[atc] = {}
            
            if move_id not in bill_groups[atc]:
                bill_groups[atc][move_id] = {
                    'wht': wht_amount,
                    'untaxed': untaxed_amount,
                    'desc': desc,
                    'code': code
                }
        
        # Process by ATC using actual WHT from tax_totals
        final = []
        month_num = self.get_bir_month_num(data[1])
        
        for atc, bills in bill_groups.items():
            dict = {'desc': '', 'code': '', 'm1': 0,
                    'm2': 0, 'm3': 0, 'taxed': 0, 'm_total': 0}
            
            for move_id, bill_data in bills.items():
                if month_num == '1':
                    dict['m1'] += bill_data['untaxed']
                elif month_num == '2':
                    dict['m2'] += bill_data['untaxed']
                elif month_num == '3':
                    dict['m3'] += bill_data['untaxed']
                
                # Use actual WHT from tax_totals
                dict['taxed'] += bill_data['wht']
                dict['m_total'] += bill_data['untaxed']
                dict['desc'] = bill_data['desc']
                dict['code'] = bill_data['code']
            
            final.append(dict)
        
        return final

    def process_2307_ammend(self, data):
        cont = []
        temp_out = ()
        final = []

        for val in data[0]:
            cont.append(val[6])

        temp_in = set(cont)
        temp_out = tuple(temp_in)

        for bp in temp_out:
            dict = []
            total = 0
            untaxed = 0
            bill_date = None
            due_date = None
            payment_state = "Unpaid"
            for dat in data[0]:
                if bp == dat[6]:
                    total = dat[9]
                    untaxed = dat[10]  # amount_untaxed
                    bill_date = dat[11]  # invoice_date
                    due_date = dat[12]   # invoice_date_due
                    payment_state = dat[13]  # payment_state
                    # Normalize payment_state display
                    if payment_state == 'paid':
                        payment_state = 'Paid'
                    elif payment_state == 'not_paid':
                        payment_state = 'Unpaid'
                    elif payment_state == 'in_payment':
                        payment_state = 'In Payment'
                    else:
                        payment_state = 'Unpaid'
                    dict = [dat[6], dat[8], dat[7], untaxed, total, bill_date, due_date, payment_state]
            final.append(dict)

        return final

##############################################################################################################################################################################
################################################################ 2550M & 2550Q ###############################################################################################
##############################################################################################################################################################################

    def x_2550_print_action(self, args):
        if args['trans'] == '2550M':
            return self.env.ref('bir_module.bir_form_2550M').report_action(self, data={'name': 'BIR Form 2550M', 'month': args['month'], 'trans': args['trans'], 'trigger': args['trigger'], 'tranid': 'none', 'ids': args['ids']})
        else:
            return self.env.ref('bir_module.bir_form_2550Q').report_action(self, data={'name': 'BIR Form 2550Q', 'month': args['month'], 'trans': args['trans'], 'trigger': args['trigger'], 'tranid': 'none', 'ids': 'none'})

    def x_2550_forms(self, args):
        val = []

        if args[2] == 'reprint':
            val = self.fetch_2550_data_reprint(args)
        elif args[2] == 'exclude-view' or args[2] == 'exclude-print':
            val = self.fetch_2550_exclude_data(args)
        else:
            val = self.fetch_2550_data_normal(args)

        if args[2] == 'print' or args[2] == 'exclude-print':
            self.record_bir_form_print(val, '2550', args[3], args[0])

        processed = self.x_2550_process_data(val)
        return processed

    def fetch_2550_table_docs_data(self, args):
        processed = []
        val = self.fetch_2550_data_normal(args)

        if args[2] == 'table':
            processed = self.process_2550_ammend(val)

        return processed

    def fetch_2550_data_reprint(self, args):
        query = """SELECT T1.move_type, T2.price_total, T2.tax_base_amount, T3.name, T3.amount, T3.tax_scope, T4.name, T5.name, T1.id, T1.name, T1.amount_total 
            FROM bir_module_print_history_line T0 
            JOIN account_move T1 ON T1.id = T0.move_id 
            JOIN account_move_line T2 ON T1.id = T2.move_id AND T2.exclude_from_invoice_tab = 'true' 
            JOIN account_tax T3 ON T3.id = T2.tax_line_id AND T3.amount >= 0 
            JOIN res_partner T4 ON T4.id = T1.partner_id 
            LEFT JOIN res_partner_industry T5 ON T5.id = T4.industry_id 
            WHERE T1.state='posted' AND T0.print_id = {0}"""

        self._cr.execute(query.format(args[4]))
        val = self._cr.fetchall()

        return val

    def fetch_2550_exclude_data(self, args):
        ids = args[5]
        if type(args[5]) == str:
            ids = json.loads(args[5])

        query = """SELECT T0.move_type, T1.price_total, T1.tax_base_amount, T3.name, T3.amount, T3.tax_scope, T4.name, T5.name, T0.id, T0.name, T0.amount_total
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id AND T1.exclude_from_invoice_tab = 'true' 
            JOIN account_tax T3 ON T3.id = T1.tax_line_id AND T3.amount >= 0 
            JOIN res_partner T4 ON T4.id = T0.partner_id 
            LEFT JOIN res_partner_industry T5 ON T5.id = T4.industry_id 
            LEFT JOIN stock_landed_cost T6 ON T0.id = T6.vendor_bill_id 
            WHERE """
        ctr = 1
        for id in ids:
            if len(ids) == ctr:
                query += "T0.id = " + str(id)
            else:
                query += "T0.id = " + str(id) + " OR "

            ctr += 1

        self._cr.execute(query)
        val = self._cr.fetchall()

        return val

    def fetch_2550_data_normal(self, args):
        param = args[0].replace("-", " ").split()
        #                   AR or AP    TAX total per line  line total      tax Name    tax Amount       ven/cust name  industry    has landed cost?
        #                   0           1               2                   3           4           5           6       7       8       9
        query = """ SELECT T0.move_type, T1.price_total, T1.tax_base_amount, T3.name, T3.amount, T3.tax_scope, T4.name, T5.name, T0.id, T0.name, T0.amount_total {3}
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id AND T1.exclude_from_invoice_tab = 'true' 
            {2}
            JOIN account_tax T3 ON T3.id = T1.tax_line_id AND T3.amount >= 0 
            JOIN res_partner T4 ON T4.id = T0.partner_id 
            LEFT JOIN res_partner_industry T5 ON T5.id = T4.industry_id 
            LEFT JOIN stock_landed_cost T6 ON T0.id = T6.vendor_bill_id 
            WHERE T0.state='posted' AND T0.company_id = {0} AND {1}"""

        quarter = {'month': param[1], 'year': param[0], 'trans': 'month'}
        if args[1] == '2550Q':
            quarter = {'month': self.x_2550_qrtrs(
                param), 'year': param[0], 'trans': 'qrtr'}

        end_param = self.x_2550_param(quarter)

        self._cr.execute(query.format(self.env.company.id,
                         end_param[0], end_param[1], end_param[2]))
        val = self._cr.fetchall()

        return val

    def x_2550_process_data(self, data):
        sales = {'12A': 0, '12B': 0, '13A': 0, '13B': 0,
                 '14': 0, '15': 0, '16A': 0, '16B': 0}
        purchase = {'E': 0, 'F': 0, 'G': 0, 'H': 0, 'I': 0,
                    'J': 0, 'K': 0, 'L': 0, 'M': 0, 'P1': 0, 'P2': 0}

        for val in data:
            flag = False
            if val[0] == 'out_invoice':
                if str(val[7]) == 'Government':
                    sales['13A'] += abs(float(val[2]))
                    sales['13B'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 12:
                    sales['12A'] += abs(float(val[2]))
                    sales['12B'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 0:
                    sales['14'] += abs(float(val[2]))
                elif 'excempt' in str(val[3]).lower():
                    sales['15'] += abs(float(val[2]))
                else:
                    pass

                if flag == True:
                    sales['16A'] += abs(float(val[2]))
                    sales['16B'] += abs(float(val[1]))
                else:
                    sales['16A'] += abs(float(val[2]))
            else:
                if int(val[4]) == 12 and str(val[5]) == 'consu':
                    purchase['E'] += abs(float(val[2]))
                    purchase['F'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 12 and str(val[5]) == 'consu':
                    purchase['G'] += abs(float(val[2]))
                    purchase['H'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 12 and str(val[5]) == 'service':
                    purchase['I'] += abs(float(val[2]))
                    purchase['J'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 12 and str(val[5]) == 'service':
                    purchase['K'] += abs(float(val[2]))
                    purchase['L'] += abs(float(val[1]))
                    flag = True
                elif int(val[4]) == 0:
                    purchase['M'] += abs(float(val[2]))
                else:
                    pass

                if flag == True:
                    purchase['P1'] += abs(float(val[2]))
                    purchase['P2'] += abs(float(val[1]))
                else:
                    purchase['P1'] += abs(float(val[2]))

        value = [sales, purchase, self.env.company.id]

        return value

    def x_2550_param(self, data):
        query = ""
        join = ""
        select = ""

        if (data['trans']) == 'month':
            init = "{0} = EXTRACT(MONTH FROM T0.date) AND {1} = EXTRACT(YEAR FROM T0.date)"
            query = init.format(data['month'], data['year'])
        else:
            init = "EXTRACT(MONTH FROM T0.date) >= {0} AND EXTRACT(MONTH FROM T0.date) <= {1} AND EXTRACT(YEAR FROM T0.date) = {2} AND T2.id IS NULL"
            query = init.format(
                data['month'][0], data['month'][1], data['year'])

            join = "LEFT JOIN bir_module_print_history_line T2 ON T2.move_id = T0.id AND T2.form_type = '2550' "
            select = ", T2.id"

        return query, join, select

    def process_2550_ammend(self, data):
        cont = []
        temp_out = ()
        final = []

        for val in data:
            cont.append(val[8])

        temp_in = set(cont)
        temp_out = tuple(temp_in)

        for move in temp_out:
            dict = []
            total = 0
            for dat in data:
                if move == dat[8]:
                    # total = dat[0]
                    dict = [dat[8], dat[9], dat[0], dat[10]]
            final.append(dict)

        return final


##############################################################################################################################################################################
################################################################ SAWT AND MAP ################################################################################################
##############################################################################################################################################################################


    def SAWT_report(self, month):
        param = month.replace("-", " ").split()

        query = """ SELECT SUM(Abs(T1.price_total)), SUM(Abs(T1.tax_base_amount)), T5.vat, MAX(T5.name), T4.name, MAX(Abs(T3.amount)) 
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id AND T1.exclude_from_invoice_tab = 'true' 
            JOIN account_tax T3 ON T3.id = T1.tax_line_id 
            JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id 
            JOIN res_partner T5 ON T5.id = T0.partner_id AND T5.vat IS NOT NULL
            WHERE T0.state='posted' AND T0.company_id = {0} AND T0.move_type = 'out_invoice' AND {1} GROUP BY T5.vat, T4.name, T3.amount"""

        quarter_iden = self.check_quarter(int(param[1]))
        end_parameter = self.sawt_map_params(quarter_iden, int(param[0]))

        self._cr.execute(query.format(self.env.company.id, end_parameter))
        val = self._cr.fetchall()

        return val

    def MAP_report(self, month):
        param = month.replace("-", " ").split()

        # return self.env.company.id
        query = """ SELECT SUM(Abs(T1.price_total)), SUM(Abs(T1.tax_base_amount)), T5.vat, MAX(T5.name), T4.name, MAX(Abs(T3.amount)) 
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id AND T1.exclude_from_invoice_tab = 'true' 
            JOIN account_tax T3 ON T3.id = T1.tax_line_id 
            JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id 
            JOIN res_partner T5 ON T5.id = T0.partner_id AND T5.vat IS NOT NULL
            WHERE T0.state='posted' AND T0.company_id = {0} AND T0.move_type = 'in_invoice' AND {1} GROUP BY T5.vat, T4.name, T3.amount"""

        quarter_iden = self.check_quarter(int(param[1]))
        end_parameter = self.sawt_map_params(quarter_iden, int(param[0]))

        self._cr.execute(query.format(self.env.company.id, end_parameter))
        val = self._cr.fetchall()

        return val

    def sawt_map_params(self, param, year):
        append = ""
        if param[0] == "monthly":
            append = "EXTRACT(MONTH FROM T0.invoice_date) = " + \
                str(param[1]) + \
                " AND EXTRACT(YEAR FROM T0.invoice_date) = " + str(year)
        else:
            append = "EXTRACT(MONTH FROM T0.invoice_date) >= " + str(param[0]) + "  AND EXTRACT(MONTH FROM T0.invoice_date) <= " + str(
                param[1]) + " AND EXTRACT(YEAR FROM T0.invoice_date) = " + str(year)
        return append

    def export_sawt_map(self, month, report):
        try:
            ### Data Preparation ###
            data = []
            title = ""
            fname = ""
            if report == 'sawt':
                data = self.SAWT_report(month)  # Fetch data
                title = "SUMMARY ALPHALIST OF WITHHOLDING TAXES"
                fname = "SAWT report.xls"
            else:
                data = self.MAP_report(month)  # Fetch data
                title = "Monthly Alphalist of Payees"
                fname = "MAP report.xls"

            param = month.replace("-", " ").split()
            scope_init = self.check_quarter(
                int(param[1]))  # Get Month or Quarter
            scope = ""
            if len(scope_init) == 2:
                scope = self.get_string_month(int(param[1]))
            else:
                scope = scope_init[2]

             # self.env.company.vat
            wb = xlwt.Workbook()
            sheet = wb.add_sheet('SAWT report')

            sheet.write(0, 0, "BIR Form 1702")
            sheet.write(1, 0, title)
            sheet.write(2, 0, "FOR THE MONTH/S OF " +
                        str(scope) + ", " + str(param[0]))
            sheet.write(4, 0, "TIN: " + str(self.env.company.vat))
            sheet.write(5, 0, "Payee's Name: " + str(self.env.company.name))

            sheet.write(7, 0, "Seq Number")
            sheet.write(8, 0, "(1)")
            sheet.write(7, 1, "Taxpayer Identification Number")
            sheet.write(8, 1, "(2)")
            sheet.write(7, 2, "Corporation (Registered Name)")
            sheet.write(8, 2, "(3)")
            sheet.write(7, 3, "ATC Code")
            sheet.write(8, 3, "(4)")
            sheet.write(7, 4, "Amount of Income Payment")
            sheet.write(8, 4, "(5)")
            sheet.write(7, 5, "Tax Rate")
            sheet.write(8, 5, "(6)")
            sheet.write(7, 6, "Amount of Tax Withheld")
            sheet.write(8, 6, "(7)")

            ctr = 9
            seq = 1
            for val in data:
                sheet.write(ctr, 0, seq)
                sheet.write(ctr, 1, val[2])
                sheet.write(ctr, 2, val[3])
                sheet.write(ctr, 3, val[4])
                sheet.write(ctr, 4, val[1])
                sheet.write(ctr, 5, val[5])
                sheet.write(ctr, 6, val[0])

                ctr += 1
                seq += 1

            if os.path.exists("C:/Odoo BIR Export/"):
                wb.save("C:/Odoo BIR Export/" + fname)
            else:
                os.makedirs("C:/Odoo BIR Export/")
                wb.save("C:/Odoo BIR Export/" + fname)

            # return Path.home()
        except Exception as ex:
            return str(ex)

    def export_sawt_map_csv(self, month, report):
        name = ["BIR Form 1702"]
        header = ["Seq Number", "Taxpayer Identification Number",
                  "Corporation (Registered Name)", "ATC Code", "Amount of Income Payment", "Tax Rate", "Amount of Tax Withheld"]
        number = ["1", "2", "3", "4", "5", "6", "7"]
        tin = ["TIN:"+str(self.env.company.vat)]
        company = ["Payee's Name:"+str(self.env.company.name)]
        data = []
        title = []
        fname = []
        vals = []

        if report == 'sawt':
            data = self.SAWT_report(month)  # Fetch data
            title = ["SUMMARY ALPHALIST OF WITHHOLDING TAXES"]
            fname = "SAWT report.csv"
        else:
            data = self.MAP_report(month)  # Fetch data
            title = ["Monthly Alphalist of Payees"]
            fname = "MAP report.csv"

        seq = 1
        for val in data:
            vals.append([seq, val[2], val[3], val[4], val[1], val[5], val[0]])
            seq += 1

        if os.path.exists("C:/Odoo BIR Export/") == False:
            os.makedirs("C:/Odoo BIR Export/")

        with open('C:/Odoo BIR Export/'+fname, 'w') as f:
            write = csv.writer(f)
            write.writerow(name)  # Form Name
            write.writerow(title)  # TITLE
            write.writerow(tin)  # TIN
            write.writerow(company)  # COMPANY NAME
            write.writerow(header)  # HEADER
            write.writerow(number)  # COL NUMBER
            write.writerows(vals)

##############################################################################################################################################################################
################################################################ SLS AND SLP #################################################################################################
##############################################################################################################################################################################

    def SLS_SLP_report(self, month, trans):
        param = month.replace("-", " ").split()

        quarter_iden = self.check_quarter(int(param[1]))
        end_parameter = self.sawt_map_params(quarter_iden, int(param[0]))

        contacts = self.get_contacts(end_parameter, trans)
        numbers = self.get_numbers(end_parameter, trans)

        vals = self.get_sls_slp_values(contacts, numbers)

        return vals

    def get_contacts(self, end_parameter, trans):
        query = """ SELECT DISTINCT(T1.vat), T1.name
            FROM account_move T0 
            JOIN res_partner T1 ON T1.id = T0.partner_id 
            WHERE T0.company_id = {0} AND T0.state = 'posted' AND T0.move_type = '{1}' AND {2} """

        self._cr.execute(query.format(
            self.env.company.id, trans, end_parameter))
        val = self._cr.fetchall()

        return val

    def get_numbers(self, end_parameter, trans):

        query = """ SELECT
            T4.vat as vat_name, T4.name as company_name, T0.name, T3.amount, T5.name, T3.tax_scope,
            CASE WHEN Abs(T3.amount) = 12 THEN Abs(price_subtotal) ELSE 0 END as VAT, 
            CASE WHEN LOWER(T5.name) LIKE '%zero%' THEN Abs(price_subtotal) ELSE 0 END as Zero_Rated, 
            CASE WHEN LOWER(T5.name) LIKE '%exempt%' THEN Abs(price_subtotal) ELSE 0 END as Exempt 
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id 
            JOIN account_move_line_account_tax_rel T2 ON T1.id = T2.account_move_line_id 
            JOIN account_tax T3 ON T3.id = T2.account_tax_id AND T3.amount >= 0 
            JOIN res_partner T4 ON T4.id = T0.partner_id 
            LEFT JOIN account_tax_group T5 ON T5.id = T3.tax_group_id 
            WHERE T0.company_id = {0} AND T0.state = 'posted' AND T0.move_type = '{1}' AND {2} """

        self._cr.execute(query.format(
            self.env.company.id, trans, end_parameter))
        val = self._cr.fetchall()

        return val

    def get_sls_slp_values(self, bp, data):
        vals = []

        for x in bp:
            vals.append({'vat': x[0], 'name': x[1],
                         'gross_sales_po': 0,
                         'exempt': 0,
                         'zero_rated': 0,
                         'taxable': 0,
                         'po_services': 0,
                         'po_capital_goods': 0,
                         'po_other': 0,
                         'tax': 0,
                         'gross_tax': 0,
                         })

        for y in vals:
            for z in data:
                if str(y['vat']) == str(z[0]):
                    tax = float(z[6]) * (float(z[3]) / 100)

                    y['gross_sales_po'] += float(z[6])
                    y['exempt'] += float(z[7])
                    y['zero_rated'] += float(z[8])
                    y['taxable'] += float(z[6])
                    y['po_other'] += float(z[6])
                    y['tax'] += round(tax, 2)
                    y['gross_tax'] += round(float(z[6]) + tax, 2)

        return vals


##############################################################################################################################################################################
################################################################ 1601e  ######################################################################################################
##############################################################################################################################################################################


    def x_1601e_print_action(self, month):
        company_id = self.env.company.id
        return self.env.ref('bir_module.1601e_report_action_id').report_action(self, data={'name': 'BIR Form 1601e', 'month': month, 'company_id:': company_id})

    def x_1601e_data(self, month):
        param = month.replace("-", " ").split()

        query = """ SELECT SUM(Abs(T1.price_total)), SUM(Abs(T1.tax_base_amount)), T4.name, T4.description, MAX(Abs(T3.amount)) 
            FROM account_move T0 
            JOIN account_move_line T1 ON T0.id = T1.move_id AND T1.exclude_from_invoice_tab = 'true' 
            JOIN account_tax T3 ON T3.id = T1.tax_line_id 
            JOIN bir_module_atc_setup T4 ON T3.id = T4.tax_id 
            WHERE T0.state='posted' AND T0.company_id = {0} AND T0.move_type = 'out_invoice' AND EXTRACT(MONTH FROM T0.date) = {1} AND EXTRACT(YEAR FROM T0.date) = {2} GROUP BY T4.name, T4.description"""

        self._cr.execute(query.format(self.env.company.id, param[1], param[0]))
        val = self._cr.fetchall()

        return val

##############################################################################################################################################################################
################################################################ PRINT HISTORY  ##############################################################################################
##############################################################################################################################################################################

    def fetch_print_types(self):
        query = "SELECT DISTINCT report_type FROM bir_module_print_history"

        self._cr.execute(query)
        val = self._cr.fetchall()

        return val

    def fetch_print_history(self, type):
        const = ""
        query = """SELECT T0.id, T0.report_type, T0.print_date, T2.name 
            FROM bir_module_print_history T0 
            JOIN res_users T1 ON T1.id = T0.create_uid 
            JOIN res_partner T2 ON T2.id = T1.partner_id 
            {0}"""

        if str(type) != 'all':
            const = "WHERE T0.report_type = '"+str(type)+"'"

        self._cr.execute(query.format(const))
        val = self._cr.fetchall()

        return val

    def fetch_print_history_details(self, id):
        query = """SELECT move_id, T1.name, scope 
            FROM bir_module_print_history_line T0 
            JOIN account_move T1 ON T1.id = T0.move_id 
            WHERE T0.print_id = {0}"""

        self._cr.execute(query.format(id))
        val = self._cr.fetchall()

        return val

    def record_bir_form_print(self, arr, process, report_type, coverage):
        data = self.process_array(arr, process)

        header_query = ""
        header_query = "INSERT INTO bir_module_print_history (form_type, report_type, create_uid, print_date, quarter_scope, create_date) VALUES('{0}', '{3}', '{1}', '{2}', '{4}', current_timestamp) RETURNING id"

        if len(data) > 0:
            self._cr.execute(header_query.format(
                data[0][1], data[0][2], data[0][3], report_type, coverage))
            print_id = self._cr.fetchone()[0]

        for val in data:
            base_query = "INSERT INTO bir_module_print_history_line (print_id, move_id, scope, form_type, create_uid, create_date) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', current_timestamp)"

            self._cr.execute(base_query.format(
                print_id, val[0], val[4], val[1], val[2]))

    def process_array(self, arr, process):
        data = []
        curr_datetime = datetime.today().strftime("%d/%m/%Y")
        if process == '2550':
            for val in arr:
                data.append([val[8], '2550', self._uid, curr_datetime, val[0]])
        elif process == '2307':
            temp = []
            set_temp = {}
            for val in arr:
                temp.append(val[6])
                set_temp = set(temp)

            for ids in list(set_temp):
                data.append([ids, '2307', self._uid, curr_datetime, arr[0][7]])

        return data

    def get_reprint_trans(self, id):
        query = "SELECT report_type, quarter_scope FROM bir_module_print_history WHERE id = '{0}'"

        self._cr.execute(query.format(id))
        val = self._cr.fetchone()

        return val

##############################################################################################################################################################################
################################################################ GENERAL FUNCTIONS  ##########################################################################################
##############################################################################################################################################################################

    def check_quarter(self, month):
        iden = ["monthly", month]
        if month == 3:
            iden = [1, 3, "January - March"]
        elif month == 6:
            iden = [4, 6, "April - June"]
        elif month == 9:
            iden = [7, 9, "July - September"]
        elif month == 12:
            iden = [10, 12, "October - December"]

        return iden

    def check_quarter_2307(self, month):
        iden = ["monthly", month]
        if month <= 3 and month >= 1:
            iden = [1, 3, "January - March"]
        elif month <= 6 and month >= 4:
            iden = [4, 6, "April - June"]
        elif month <= 9 and month >= 5:
            iden = [7, 9, "July - September"]
        elif month <= 12 and month >= 10:
            iden = [10, 12, "October - December"]

        return iden

    def get_string_month(self, num):
        months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December']

        return months[num-1]

    def x_2550_qrtrs(self, param):
        month = int(param[1])
        quarter = []
        if month >= 1 and month <= 3:
            quarter = [1, 3, "January", "March", 1]
        elif month >= 4 and month <= 6:
            quarter = [4, 6, "April", "June", 2]
        elif month >= 7 and month <= 9:
            quarter = [7, 9, "July", "September", 3]
        else:
            quarter = [10, 12, "October", "December", 4]

        return quarter

    def fetch_BP(self):
        query = """ SELECT id, name FROM res_partner WHERE supplier_rank > 0 AND active = true AND vat IS NOT NULL"""

        self._cr.execute(query)
        val = self._cr.fetchall()
        
        # Fetch tags for each partner using sudo to bypass security checks
        result = []
        for partner_id, partner_name in val:
            try:
                partner = self.env['res.partner'].sudo().browse(partner_id)
                tags = [tag.name for tag in partner.category_id]
            except Exception:
                tags = []
            result.append([partner_id, partner_name, tags])

        return result

    def get_bir_quarter(self, value):
        quarter = 0
        month = int(value.month)

        if month == 4 or month == 1 or month == 7 or month == 10:
            quarter = 1
        elif month == 2 or month == 5 or month == 8 or month == 11:
            quarter = 2
        else:
            quarter = 3

        return str(quarter)

    def get_bir_month_num(self, value):
        if not value:
            return '1'  # Default to first quarter month if no value
        
        import datetime
        import logging
        _logger = logging.getLogger(__name__)
        
        # Extract month number from the value
        month_int = None
        
        # If it's a date object, get month directly
        if hasattr(value, 'month'):
            month_int = value.month
        else:
            # Convert to string and try to parse
            value_str = str(value).strip()
            
            # Try to extract month from various date formats
            try:
                # Format: YYYY-MM-DD
                if '-' in value_str:
                    parts = value_str.split('-')
                    if len(parts) >= 2:
                        month_int = int(parts[1])
                # Format: YYYY/MM/DD
                elif '/' in value_str:
                    parts = value_str.split('/')
                    if len(parts) >= 2:
                        month_int = int(parts[1])
                # Format: MM/DD/YYYY
                elif len(value_str) == 10 and '/' in value_str:
                    parts = value_str.split('/')
                    month_int = int(parts[0])
            except (ValueError, IndexError, AttributeError) as e:
                _logger.warning(f"Failed to extract month from {value}: {e}")
                pass
        
        # If we couldn't extract month, default to 1
        if month_int is None:
            _logger.warning(f"Could not determine month from value: {value}, defaulting to 1")
            return '1'
        
        _logger.info(f"Processing date {value}, extracted month: {month_int}")
        
        # Determine quarter position based on month
        if month_int == 1 or month_int == 4 or month_int == 7 or month_int == 10:
            quarter = 1
        elif month_int == 2 or month_int == 5 or month_int == 8 or month_int == 11:
            quarter = 2
        else:  # months 3, 6, 9, 12
            quarter = 3

        _logger.info(f"Calculated quarter {quarter} for month {month_int}")
        return str(quarter)

    def splice_month(self, month):
        if not month:
            return ['0', '1']  # Return default year and month
        
        # Handle date objects by converting to string
        month_str = str(month)
        
        # Replace common date separators and split
        parts = month_str.replace("-", " ").replace("/", " ").split()
        
        # If we got date parts, return first 2 elements (year and month)
        if len(parts) >= 2:
            return parts[:2]
        elif len(parts) == 1:
            return ['0', parts[0]]  # If only one element, assume it's the month
        else:
            return ['0', '1']  # Return defaults if we couldn't parse

    def get_marker_quarter(self, month):
        return self.x_2550_qrtrs(self.splice_month(month))

    def fetch_period_dates(self, value):
        vals = self.splice_month(value)
        
        # Ensure we have the required elements in vals
        if len(vals) < 2:
            raise ValueError(f"Invalid month format. Expected 'YYYY-MM' but got '{value}'")

        month = self.x_2550_qrtrs(vals)

        range1 = calendar.monthrange(int(vals[0]), int(month[0]))
        range2 = calendar.monthrange(int(vals[0]), int(month[1]))

        month1 = str(month[0])
        month2 = str(month[1])

        if int(month[0]) < 10:
            month1 = "0" + month1

        if int(month[1]) < 10:
            month2 = "0" + month2

        return [str(vals[0][2:]), month1, month2, str(range1[1]), str(range2[1]), str(vals[0]), str(vals[1])]

    def fetch_month_coverage(self, value):
        year = value.strftime("%Y")
        month = value.strftime("%m")

        range = calendar.monthrange(int(year), int(month))

        return {'month': str(month), 'year': str(year), 'date': str(range[1])}

    # def x_format_vat(self, vat):
    #     val = ""
    #     if vat != False:
    #         val = vat[:3] + "-" + vat[3:6] + "-" + vat[6:]

    #     return str(val)

    # def x_slice_vat(self, vat):
    #     val = []
    #     if vat != False:
    #         val = [vat[:3], vat[3:6], vat[6:], '000']

    #     return val

    def x_format_vat(self, vat):
        val = ""
        if vat != False:
            if vat[3] == "-":
                val = vat
            else:
                val = vat[:3] + "-" + vat[4:6] + "-" + vat[7:]

        return str(val)

    def x_slice_vat(self, vat):
        val = []
        if vat != False:
            if vat[3] == "-":
                val = [vat[:3], vat[4:7], vat[8:11], vat[12:15]]  # Extract actual last segment
            else:
                val = [vat[:3], vat[3:6], vat[6:9], vat[9:12]]  # Extract actual last segment
        return val

    def x_fetch_company_id(self):
        """Return the user's currently active company ID from their session"""
        return self.env.user.company_id.id
    
    def _log_company_debug(self, template_name, raw_company_id, context_company_ids, company_id):
        """Debug helper to log company context information"""
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"BIR 2307 {template_name}: raw_company_id={raw_company_id}, context_company_ids={context_company_ids}, final_company_id={company_id}, env.company={self.env.company.id}, env.user.company_id={self.env.user.company_id.id}")
        return True
    
    def _debug_2307_template(self, raw_company_id, context_company_ids, company_id):
        """Debug helper for template to log values"""
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"DEBUG TEMPLATE: raw_company_id={raw_company_id}, context_company_ids={context_company_ids}, final_company_id={company_id}, env.company={self.env.company.id}")
        return True
    
    # ============================= BIR 2307 DATE RANGE METHODS =============================
    
    def parse_date_string(self, date_str):
        """Parse date string (YYYY-MM-DD format) to date object"""
        if hasattr(date_str, 'year'):  # Already a date object
            return date_str
        from datetime import datetime
        try:
            return datetime.strptime(str(date_str), '%Y-%m-%d').date()
        except:
            return None
    
    def get_quarter_from_date(self, date_obj):
        """Get quarter (1-4) from a date object"""
        if not date_obj:
            return None
        month = date_obj.month
        if month in [1, 2, 3]:
            return 1
        elif month in [4, 5, 6]:
            return 2
        elif month in [7, 8, 9]:
            return 3
        else:
            return 4
    
    def get_quarter_month_range(self, quarter_num):
        """Get the month range for a given quarter (1-4)"""
        quarter_ranges = {
            1: (1, 3),
            2: (4, 6),
            3: (7, 9),
            4: (10, 12)
        }
        return quarter_ranges.get(quarter_num, (1, 3))
    
    def validate_date_range_quarterly(self, from_date, to_date):
        """
        Validate that the date range is within one quarter.
        Both dates should fall within the same quarter and be normalized to month boundaries.
        Returns: tuple (is_valid, error_message, normalized_from_date, normalized_to_date)
        """
        from datetime import datetime, timedelta
        
        from_date_obj = self.parse_date_string(from_date)
        to_date_obj = self.parse_date_string(to_date)
        
        if not from_date_obj or not to_date_obj:
            return False, "Invalid date format. Please use YYYY-MM-DD format.", None, None
        
        if from_date_obj > to_date_obj:
            return False, "From date cannot be after To date.", None, None
        
        # Get quarters for both dates
        from_quarter = self.get_quarter_from_date(from_date_obj)
        to_quarter = self.get_quarter_from_date(to_date_obj)
        
        # Check if quarters match
        if from_quarter != to_quarter:
            return False, f"Both dates must fall within the same quarter. From date is Q{from_quarter}, To date is Q{to_quarter}.", None, None
        
        # Get the year from both dates
        if from_date_obj.year != to_date_obj.year:
            return False, "Both dates must be in the same calendar year.", None, None
        
        # Normalize dates: from_date to 1st, to_date to last day of month
        normalized_from = from_date_obj.replace(day=1)
        
        # Get last day of to_date month
        next_month = to_date_obj.replace(day=28) + timedelta(days=4)
        last_day = (next_month - timedelta(days=next_month.day)).day
        normalized_to = to_date_obj.replace(day=last_day)
        
        return True, "", normalized_from, normalized_to
    
    def get_selected_months_in_quarter(self, from_date, to_date):
        """
        Get which months within the quarter are selected (1st, 2nd, 3rd).
        Returns list of month positions selected and mapping info.
        Example: Feb 1 - Feb 28 in Q1 returns [2] (only 2nd month)
                 Jan 1 - Mar 31 in Q1 returns [1, 2, 3] (all months)
        """
        from datetime import datetime
        
        from_date_obj = self.parse_date_string(from_date)
        to_date_obj = self.parse_date_string(to_date)
        
        if not from_date_obj or not to_date_obj:
            return []
        
        quarter = self.get_quarter_from_date(from_date_obj)
        quarter_start_month, quarter_end_month = self.get_quarter_month_range(quarter)
        
        # Get months in the selected range
        selected_months = []
        current_month = from_date_obj.month
        
        while current_month <= to_date_obj.month:
            # Map calendar month to quarter position (1, 2, or 3)
            month_position = current_month - quarter_start_month + 1
            if month_position not in selected_months:
                selected_months.append(month_position)
            current_month += 1
        
        return sorted(selected_months)
    
    def get_month_position_for_date(self, date_obj, selected_from_date, selected_to_date):
        """
        Determine which month position (1, 2, or 3) within the quarter an invoice date falls into,
        based on the selected date range.
        """
        from_date_obj = self.parse_date_string(selected_from_date)
        to_date_obj = self.parse_date_string(selected_to_date)
        invoice_date = self.parse_date_string(date_obj)
        
        if not from_date_obj or not to_date_obj or not invoice_date:
            return '1'
        
        # Check if invoice is outside selected range
        if invoice_date < from_date_obj or invoice_date > to_date_obj:
            return '0'  # Outside range
        
        quarter = self.get_quarter_from_date(invoice_date)
        quarter_start_month, _ = self.get_quarter_month_range(quarter)
        
        # Get month position within quarter (1, 2, or 3)
        month_position = invoice_date.month - quarter_start_month + 1
        
        return str(max(1, min(3, month_position)))  # Clamp to 1-3
    
    def fetch_period_dates_range(self, from_date, to_date):
        """
        Fetch formatted period dates for the selected range.
        Returns array: [from_year_2digit, from_month_padded, to_month_padded, from_day, to_day, full_year, from_month_num]
        """
        from datetime import datetime, timedelta
        
        from_date_obj = self.parse_date_string(from_date)
        to_date_obj = self.parse_date_string(to_date)
        
        if not from_date_obj or not to_date_obj:
            return []
        
        # Normalize dates
        from_normalized = from_date_obj.replace(day=1)
        next_month = to_date_obj.replace(day=28) + timedelta(days=4)
        last_day = (next_month - timedelta(days=next_month.day)).day
        to_normalized = to_date_obj.replace(day=last_day)
        
        from_month = str(from_normalized.month).zfill(2)
        to_month = str(to_normalized.month).zfill(2)
        
        return [
            str(from_normalized.year)[2:],  # 2-digit year
            from_month,                      # from month (padded)
            to_month,                        # to month (padded)
            '01',                            # from day (always 1st)
            str(last_day).zfill(2),          # to day (last day of month)
            str(from_normalized.year),       # full year
            str(from_normalized.month)       # from month number
        ]

    
class bir_signee_setup(models.Model):
    _name = 'bir_module.signee_setup'
    _description = 'BIR Signee / Authorized Representative Setup'
    _order = 'sequence asc'

    name = fields.Char(string='Name', required=True)
    tax_id = fields.Char(string='Tax ID / TIN')
    position = fields.Char(string='Position', required=True)
    sequence = fields.Integer(string='Sequence', default=1, help='Sequence to determine default signee (lower number = higher priority)')

    @api.model
    def get_default_signee(self):
        """Get the signee with the lowest sequence (highest priority)"""
        signee = self.search([], order='sequence asc', limit=1)
        if signee:
            return {
                'id': signee.id,
                'name': signee.name,
                'tax_id': signee.tax_id or '',
                'position': signee.position
            }
        return {'id': 0, 'name': '', 'tax_id': '', 'position': ''}

    @api.model
    def get_signee_by_id(self, signee_id):
        """Get signee details by ID"""
        if not signee_id:
            return self.get_default_signee()
        signee = self.browse(signee_id)
        if signee.exists():
            return {
                'id': signee.id,
                'name': signee.name,
                'tax_id': signee.tax_id or '',
                'position': signee.position
            }
        return self.get_default_signee()
