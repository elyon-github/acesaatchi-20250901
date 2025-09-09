# -*- coding: utf-8 -*-
# from odoo import http


# class SaatchiCustomizedAccruedRevenue(http.Controller):
#     @http.route('/saatchi_customized_accrued_revenue/saatchi_customized_accrued_revenue', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/saatchi_customized_accrued_revenue/saatchi_customized_accrued_revenue/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('saatchi_customized_accrued_revenue.listing', {
#             'root': '/saatchi_customized_accrued_revenue/saatchi_customized_accrued_revenue',
#             'objects': http.request.env['saatchi_customized_accrued_revenue.saatchi_customized_accrued_revenue'].search([]),
#         })

#     @http.route('/saatchi_customized_accrued_revenue/saatchi_customized_accrued_revenue/objects/<model("saatchi_customized_accrued_revenue.saatchi_customized_accrued_revenue"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('saatchi_customized_accrued_revenue.object', {
#             'object': obj
#         })

