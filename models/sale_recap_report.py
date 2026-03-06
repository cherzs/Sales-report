# -*- coding: utf-8 -*-

from odoo import models, fields, api, tools
from odoo.exceptions import UserError
import logging
import io
import base64

_logger = logging.getLogger(__name__)


class SaleOrder(models.Model):
    """Inherit Sale Order to add missing field"""
    _inherit = 'sale.order'
    
    is_tax_computed_externally = fields.Boolean(
        string='Tax Computed Externally',
        default=False,
        help='Technical field to prevent external tax computation errors'
    )


class AccountMove(models.Model):
    """Inherit Account Move to add missing field"""
    _inherit = 'account.move'
    
    is_tax_computed_externally = fields.Boolean(
        string='Is Tax Computed Externally',
        default=False,
        help='Technical field to prevent external tax computation errors'
    )


class AccountBankStatementLine(models.Model):
    """Inherit Bank Statement Line to add missing field"""
    _inherit = 'account.bank.statement.line'
    
    is_tax_computed_externally = fields.Boolean(
        string='Is Tax Computed Externally',
        default=False,
        help='Technical field to prevent external tax computation errors'
    )


class GrossProfit(models.Model):
    """1. Gross Profit Report - Fixed with proper COGS calculation"""
    
    _name = 'x_gross.profit'
    _description = 'Gross Profit'
    _auto = False
    _order = 'category_items'
    
    category_items = fields.Char(string='Category Items', readonly=True)
    qty = fields.Float(string='Qty', readonly=True)
    amount = fields.Float(string='Amount', readonly=True)
    cogs = fields.Float(string='COGS', readonly=True)
    gp_percent = fields.Float(string='GP %', readonly=True)
    total_gross_profit = fields.Float(string='Total Gross Profit', readonly=True)
    # Date field for filtering
    so_date = fields.Date(string='SO Date', readonly=True)
    # Clickable reference
    categ_id = fields.Many2one('product.category', string='Category Ref', readonly=True)
    
    @api.model
    def init(self):
        _logger.info("[GROSS_PROFIT] Creating view %s", self._table)
        try:
            tools.drop_view_if_exists(self.env.cr, self._table)
            self.env.cr.execute("""
                CREATE OR REPLACE VIEW %s AS (
                    SELECT
                        ROW_NUMBER() OVER () AS id,
                        COALESCE(NULLIF(TRIM(pc.name), ''), 'Uncategorized') AS category_items,
                        SUM(sol.product_uom_qty) AS qty,
                        SUM(sol.price_subtotal) AS amount,
                        SUM(sol.product_uom_qty * COALESCE((SELECT value::numeric FROM jsonb_each_text(pp.standard_price) LIMIT 1), 0)) AS cogs,
                        CASE 
                            WHEN SUM(sol.price_subtotal) > 0 
                            THEN ((SUM(sol.price_subtotal) - SUM(sol.product_uom_qty * COALESCE((SELECT value::numeric FROM jsonb_each_text(pp.standard_price) LIMIT 1), 0))) / SUM(sol.price_subtotal)) * 100
                            ELSE 0 
                        END AS gp_percent,
                        SUM(sol.price_subtotal) - SUM(sol.product_uom_qty * COALESCE((SELECT value::numeric FROM jsonb_each_text(pp.standard_price) LIMIT 1), 0)) AS total_gross_profit,
                        DATE_TRUNC('month', MIN(so.date_order))::date AS so_date,
                        COALESCE(MAX(pc.id), 0) AS categ_id
                    FROM sale_order_line sol
                    INNER JOIN sale_order so ON so.id = sol.order_id
                    LEFT JOIN product_product pp ON pp.id = sol.product_id
                    LEFT JOIN product_template pt ON pt.id = pp.product_tmpl_id
                    LEFT JOIN product_category pc ON pc.id = pt.categ_id
                    WHERE so.state IN ('sale', 'done')
                    GROUP BY pc.id, pc.name
                    ORDER BY category_items
                )
            """ % self._table)
            _logger.info("[GROSS_PROFIT] View %s created successfully", self._table)
        except Exception as e:
            _logger.error("[GROSS_PROFIT] Error creating view %s: %s", self._table, str(e))
            raise


class RekapSOPayment(models.Model):
    """2. Rekap SO sampai Payment - Complete with all fields"""
    
    _name = 'x_rekap.so.payment'
    _description = 'Rekap SO sampai Payment'
    _auto = False
    _order = 'po_date desc, so_number'
    
    # SO Info
    so_number = fields.Char(string='SO Number', readonly=True)
    po_date = fields.Date(string='PO Date', readonly=True)
    customer = fields.Char(string='Customer', readonly=True)
    company_name = fields.Char(string='Company Name', readonly=True)
    customer_po_number = fields.Char(string='Customer PO Number', readonly=True)
    salesperson = fields.Char(string='Salesperson', readonly=True)
    # Clickable references
    order_id = fields.Many2one('sale.order', string='SO Ref', readonly=True)
    partner_id = fields.Many2one('res.partner', string='Customer Ref', readonly=True)
    picking_id = fields.Many2one('stock.picking', string='Delivery Ref', readonly=True)
    invoice_id = fields.Many2one('account.move', string='Invoice Ref', readonly=True)
    
    # Product
    product_name = fields.Char(string='Product', readonly=True)
    qty = fields.Float(string='Qty', readonly=True)
    price_unit = fields.Float(string='Price Unit', readonly=True)
    subtotal = fields.Float(string='Subtotal (Before Tax)', readonly=True)
    tax_amount = fields.Float(string='Tax', readonly=True)
    total_amount = fields.Float(string='Total Amount', readonly=True)
    
    # Delivery
    delivery_number = fields.Char(string='Delivery Number', readonly=True)
    delivery_date = fields.Date(string='Delivery Date', readonly=True)
    delivery_status = fields.Char(string='Delivery Status', readonly=True)
    delivered_qty = fields.Float(string='Delivered Qty', readonly=True)
    branch_delivery = fields.Char(string='Branch Delivery', readonly=True)
    receiver = fields.Char(string='Receiver', readonly=True)
    shipping_note = fields.Text(string='Shipping Note', readonly=True)
    
    # Invoice & Payment
    invoice_status = fields.Char(string='Invoice Status', readonly=True)
    invoice_number = fields.Char(string='Invoice Number', readonly=True)
    invoice_date = fields.Date(string='Invoice Date', readonly=True)
    payment_date = fields.Date(string='Payment Date', readonly=True)
    payment_state = fields.Char(string='Payment State', readonly=True)
    
    @api.model
    def init(self):
        _logger.info("[REKAP_SO] Creating view %s", self._table)
        try:
            tools.drop_view_if_exists(self.env.cr, self._table)
            sql = """
                CREATE OR REPLACE VIEW %s AS (
                    WITH product_names AS (
                        -- Resolve multilang product name (jsonb or plain text)
                        SELECT
                            pt.id AS tmpl_id,
                            COALESCE(
                                NULLIF(TRIM(
                                    CASE
                                        WHEN LEFT(pt.name::text, 1) = '{'
                                        THEN pt.name->>'en_US'
                                        ELSE pt.name::text
                                    END
                                ), ''),
                                ''
                            ) AS product_name
                        FROM product_template pt
                    ),
                    payment_data AS (
                        -- Earliest payment date per invoice via partial reconcile
                        SELECT
                            aml_inv.move_id                     AS invoice_id,
                            MIN(ap.date)                        AS first_payment_date,
                            STRING_AGG(DISTINCT ap.state, ', ') AS payment_states
                        FROM account_move_line aml_inv
                        INNER JOIN account_partial_reconcile apr
                            ON apr.debit_move_id  = aml_inv.id
                            OR apr.credit_move_id = aml_inv.id
                        INNER JOIN account_move_line aml_pay
                            ON (apr.credit_move_id = aml_pay.id AND aml_inv.id != aml_pay.id)
                            OR (apr.debit_move_id  = aml_pay.id AND aml_inv.id != aml_pay.id)
                        INNER JOIN account_payment ap ON ap.move_id = aml_pay.move_id
                        WHERE ap.state = 'posted'
                        GROUP BY aml_inv.move_id
                    ),
                    delivery_data AS (
                        -- HYBRID: handles both new data (sale_line_id populated)
                        -- and legacy data (sale_line_id = NULL, use origin fallback)

                        -- METHOD A: New data — linked via stock_move.sale_line_id
                        SELECT
                            sm.sale_line_id                    AS so_line_id,
                            sp.id                              AS picking_id,
                            sp.name                            AS delivery_number,
                            sp.date_done::date                 AS delivery_date,
                            sp.state                           AS delivery_status,
                            rp.name                            AS receiver,
                            sp.note                            AS shipping_note,
                            sw.name                            AS branch_name,
                            SUM(COALESCE(sml.quantity, 0))     AS total_delivered_qty
                        FROM stock_move sm
                        INNER JOIN sale_order_line sol_chk ON sol_chk.id = sm.sale_line_id
                            AND sol_chk.product_id = sm.product_id
                        INNER JOIN stock_picking sp ON sp.id = sm.picking_id
                        INNER JOIN stock_move_line sml ON sml.move_id = sm.id
                        LEFT JOIN res_partner rp ON rp.id = sp.partner_id
                        LEFT JOIN stock_picking_type spt ON spt.id = sp.picking_type_id
                        LEFT JOIN stock_warehouse sw ON sw.id = spt.warehouse_id
                        WHERE sm.sale_line_id IS NOT NULL
                          AND sm.state = 'done'
                          AND sp.state = 'done'
                          AND spt.code = 'outgoing'
                        GROUP BY
                            sm.sale_line_id, sp.id, sp.name, sp.date_done,
                            sp.state, rp.name, sp.note, sw.name

                        UNION

                        -- METHOD B: Legacy data — sale_line_id IS NULL on stock_move
                        -- Fallback to origin = so.name, matched by product
                        SELECT
                            sol_leg.id                         AS so_line_id,
                            sp.id                              AS picking_id,
                            sp.name                            AS delivery_number,
                            sp.date_done::date                 AS delivery_date,
                            sp.state                           AS delivery_status,
                            rp.name                            AS receiver,
                            sp.note                            AS shipping_note,
                            sw.name                            AS branch_name,
                            SUM(COALESCE(sml.quantity, 0))     AS total_delivered_qty
                        FROM stock_picking sp
                        INNER JOIN sale_order so_leg ON so_leg.name = sp.origin
                        INNER JOIN sale_order_line sol_leg ON sol_leg.order_id = so_leg.id
                        INNER JOIN stock_move sm ON sm.picking_id = sp.id
                            AND sm.product_id = sol_leg.product_id
                            AND sm.sale_line_id IS NULL
                        INNER JOIN stock_move_line sml ON sml.move_id = sm.id
                        LEFT JOIN res_partner rp ON rp.id = sp.partner_id
                        LEFT JOIN stock_picking_type spt ON spt.id = sp.picking_type_id
                        LEFT JOIN stock_warehouse sw ON sw.id = spt.warehouse_id
                        WHERE sm.state = 'done'
                          AND sp.state = 'done'
                          AND spt.code = 'outgoing'
                          AND so_leg.state IN ('sale', 'done')
                        GROUP BY
                            sol_leg.id, sp.id, sp.name, sp.date_done,
                            sp.state, rp.name, sp.note, sw.name
                    ),
                    invoice_data AS (
                        -- Invoice linked via the junction table sale_order_line_invoice_rel
                        -- Odoo 18: column1='order_line_id', column2='invoice_line_id'
                        -- Supports: 1 Delivery → 1 Invoice AND N Deliveries → 1 Invoice
                        SELECT DISTINCT
                            rel.order_line_id                  AS so_line_id,
                            am.id                              AS invoice_id,
                            am.name                            AS invoice_number,
                            am.invoice_date                    AS invoice_date
                        FROM sale_order_line_invoice_rel rel
                        INNER JOIN account_move_line aml ON aml.id = rel.invoice_line_id
                        INNER JOIN account_move am ON am.id = aml.move_id
                        WHERE am.move_type = 'out_invoice'
                          AND am.state     = 'posted'
                    )
                    SELECT
                        -- Stable unique ID: sol.id * 10,000,000 + sp.id
                        (sol.id::bigint * 10000000 + dd.picking_id::bigint) AS id,
                        so.id AS order_id,
                        so.name AS so_number,
                        so.date_order::date AS po_date,
                        rp.id AS partner_id,
                        COALESCE(NULLIF(TRIM(rp.name), ''), '') AS customer,
                        COALESCE(
                            NULLIF(TRIM(rp.commercial_company_name), ''),
                            NULLIF(TRIM(rp.name), ''),
                            ''
                        ) AS company_name,
                        COALESCE(NULLIF(TRIM(so.client_order_ref), ''), '') AS customer_po_number,
                        COALESCE(NULLIF(TRIM(rp_sales.name), ''), '') AS salesperson,
                        COALESCE(NULLIF(TRIM(pn.product_name), ''), '') AS product_name,
                        COALESCE(sol.product_uom_qty, 0.0) AS qty,
                        COALESCE(sol.price_unit, 0.0) AS price_unit,
                        COALESCE(sol.price_subtotal, 0.0) AS subtotal,
                        COALESCE(sol.price_tax, 0.0) AS tax_amount,
                        COALESCE(sol.price_total, 0.0) AS total_amount,
                        dd.picking_id,
                        COALESCE(dd.delivery_number, '') AS delivery_number,
                        dd.delivery_date AS delivery_date,
                        COALESCE(dd.delivery_status, '') AS delivery_status,
                        COALESCE(dd.total_delivered_qty, 0.0) AS delivered_qty,
                        COALESCE(NULLIF(TRIM(dd.branch_name), ''), '-') AS branch_delivery,
                        COALESCE(NULLIF(TRIM(dd.receiver), ''), '') AS receiver,
                        COALESCE(dd.shipping_note, '') AS shipping_note,
                        COALESCE(so.invoice_status, '') AS invoice_status,
                        idata.invoice_id,
                        COALESCE(idata.invoice_number, '') AS invoice_number,
                        idata.invoice_date AS invoice_date,
                        pd.first_payment_date AS payment_date,
                        COALESCE(pd.payment_states, '') AS payment_state
                    FROM sale_order_line sol
                    -- SO
                    INNER JOIN sale_order so ON so.id = sol.order_id
                    -- Multi-delivery: INNER JOIN ensures one row per (sol, picking)
                    INNER JOIN delivery_data dd ON dd.so_line_id = sol.id
                    -- Customer & Salesperson
                    LEFT JOIN res_partner rp ON rp.id = so.partner_id
                    LEFT JOIN res_users ru ON ru.id = so.user_id
                    LEFT JOIN res_partner rp_sales ON rp_sales.id = ru.partner_id
                    -- Product name
                    LEFT JOIN product_product pp ON pp.id = sol.product_id
                    LEFT JOIN product_names pn ON pn.tmpl_id = pp.product_tmpl_id
                    -- Invoice via junction table
                    LEFT JOIN invoice_data idata ON idata.so_line_id = sol.id
                    -- Payment via partial reconcile
                    LEFT JOIN payment_data pd ON pd.invoice_id = idata.invoice_id
                    WHERE so.state IN ('sale', 'done')
                    ORDER BY so.date_order DESC, so.name, sol.sequence, dd.delivery_number
                )
            """ % self._table
            _logger.info("[REKAP_SO] SQL Query prepared")
            self.env.cr.execute(sql)
            _logger.info("[REKAP_SO] View %s created successfully", self._table)
        except Exception as e:
            _logger.error("[REKAP_SO] Error creating view %s: %s", self._table, str(e))
            _logger.error("[REKAP_SO] SQL: %s", sql)
            raise


class SalesContribution(models.Model):
    """3. Sales Contribution Report - Fixed with proper COGS calculation"""
    
    _name = 'x_sales.contribution'
    _description = 'Sales Contribution'
    _auto = False
    _order = 'category'
    
    category = fields.Char(string='Category', readonly=True)
    sales_amount = fields.Float(string='Sales Amount', readonly=True)
    cogs = fields.Float(string='COGS', readonly=True)
    gross_profit = fields.Float(string='Gross Profit', readonly=True)
    margin_percent = fields.Float(string='Margin %', readonly=True)
    sales_contribution_percent = fields.Float(string='Sales Contribution %', readonly=True)
    # Date field for filtering
    so_date = fields.Date(string='SO Date', readonly=True)
    # Clickable reference
    categ_id = fields.Many2one('product.category', string='Category Ref', readonly=True)
    
    @api.model
    def init(self):
        _logger.info("[SALES_CONTRIBUTION] Creating view %s", self._table)
        try:
            tools.drop_view_if_exists(self.env.cr, self._table)
            sql = """
                CREATE OR REPLACE VIEW %s AS (
                    WITH category_sales AS (
                        SELECT
                            COALESCE(NULLIF(TRIM(pc.name), ''), 'Uncategorized') AS category,
                            COALESCE(MAX(pc.id), 0) AS categ_id,
                            SUM(sol.price_subtotal) AS sales_amount,
                            SUM(sol.product_uom_qty * COALESCE((SELECT value::numeric FROM jsonb_each_text(pp.standard_price) LIMIT 1), 0)) AS cogs,
                            SUM(sol.price_subtotal) - SUM(sol.product_uom_qty * COALESCE((SELECT value::numeric FROM jsonb_each_text(pp.standard_price) LIMIT 1), 0)) AS gross_profit,
                            DATE_TRUNC('month', MIN(so.date_order))::date AS so_date
                        FROM sale_order_line sol
                        INNER JOIN sale_order so ON so.id = sol.order_id
                        LEFT JOIN product_product pp ON pp.id = sol.product_id
                        LEFT JOIN product_template pt ON pt.id = pp.product_tmpl_id
                        LEFT JOIN product_category pc ON pc.id = pt.categ_id
                        WHERE so.state IN ('sale', 'done')
                        GROUP BY pc.id, pc.name
                    ),
                    total_sales AS (
                        SELECT SUM(sales_amount) AS total FROM category_sales
                    )
                    SELECT
                        ROW_NUMBER() OVER () AS id,
                        cs.category,
                        cs.categ_id,
                        cs.sales_amount,
                        cs.cogs,
                        cs.gross_profit,
                        CASE 
                            WHEN cs.sales_amount > 0 THEN (cs.gross_profit / cs.sales_amount) * 100 
                            ELSE 0 
                        END AS margin_percent,
                        CASE 
                            WHEN ts.total > 0 THEN (cs.sales_amount / ts.total) * 100 
                            ELSE 0 
                        END AS sales_contribution_percent,
                        cs.so_date
                    FROM category_sales cs
                    CROSS JOIN total_sales ts
                    ORDER BY cs.sales_amount DESC
                )
            """ % self._table
            _logger.info("[SALES_CONTRIBUTION] SQL Query prepared")
            self.env.cr.execute(sql)
            _logger.info("[SALES_CONTRIBUTION] View %s created successfully", self._table)
        except Exception as e:
            _logger.error("[SALES_CONTRIBUTION] Error creating view %s: %s", self._table, str(e))
            _logger.error("[SALES_CONTRIBUTION] SQL: %s", sql)
            raise


class SaleRecapExportExcel(models.TransientModel):
    """Wizard Export Excel untuk semua report"""
    
    _name = 'x_sale.recap.export.excel'
    _description = 'Export Sales Recap to Excel'
    
    report_type = fields.Selection([
        ('gross_profit', 'Gross Profit'),
        ('rekap_so', 'Rekap SO sampai Payment'),
        ('sales_contribution', 'Sales Contribution'),
        ('all', 'All Combined'),
    ], string='Report Type', required=True, default='all')
    
    date_from = fields.Date(string='Date From', required=True)
    date_to = fields.Date(string='Date To', required=True)
    
    @api.model
    def default_get(self, fields_list):
        res = super(SaleRecapExportExcel, self).default_get(fields_list)
        today = fields.Date.context_today(self)
        res['date_from'] = today.replace(day=1)
        res['date_to'] = today
        return res
    
    def action_open_wizard(self):
        """Open wizard form"""
        _logger.info("[EXPORT] Opening wizard for report_type: %s", self.env.context.get('default_report_type'))
        return {
            'name': 'Export to Excel',
            'type': 'ir.actions.act_window',
            'res_model': 'x_sale.recap.export.excel',
            'view_mode': 'form',
            'target': 'new',
            'context': self.env.context,
        }
    
    def action_export_xlsx(self):
        """Generate XLSX report"""
        self.ensure_one()
        _logger.info("[EXPORT] Starting export for report_type: %s, date_from: %s, date_to: %s", 
                     self.report_type, self.date_from, self.date_to)
        
        try:
            import xlsxwriter
        except ImportError as e:
            _logger.error("[EXPORT] xlsxwriter not installed: %s", str(e))
            raise UserError('Library xlsxwriter tidak ditemukan. Install dengan: pip install xlsxwriter')
        
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        try:
            if self.report_type == 'gross_profit' or self.report_type == 'all':
                _logger.info("[EXPORT] Exporting Gross Profit")
                self._export_gross_profit(workbook)
            
            if self.report_type == 'rekap_so' or self.report_type == 'all':
                _logger.info("[EXPORT] Exporting Rekap SO")
                self._export_rekap_so(workbook)
            
            if self.report_type == 'sales_contribution' or self.report_type == 'all':
                _logger.info("[EXPORT] Exporting Sales Contribution")
                self._export_sales_contribution(workbook)
            
            workbook.close()
            output.seek(0)
            
            filename = 'SALES_RECAP_%s.xlsx' % fields.Date.today().strftime('%Y%m%d')
            attachment = self.env['ir.attachment'].create({
                'name': filename,
                'type': 'binary',
                'datas': base64.b64encode(output.read()).decode('utf-8'),
                'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            })
            
            _logger.info("[EXPORT] Export completed successfully: %s", filename)
            return {
                'type': 'ir.actions.act_url',
                'url': '/web/content/%s?download=true' % attachment.id,
                'target': 'self',
            }
        except Exception as e:
            _logger.error("[EXPORT] Error during export: %s", str(e))
            raise
    
    def _write_excel_header(self, sheet, workbook, title, company_name):
        """Write standardized header to Excel"""
        company_format = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
        })
        title_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'
        })
        date_format = workbook.add_format({
            'italic': True, 'align': 'center', 'valign': 'vcenter'
        })
        
        # Merge cells for header
        sheet.merge_range('A1:F1', company_name, company_format)
        sheet.merge_range('A2:F2', title, title_format)
        sheet.merge_range('A3:F3', 'AS OF %s' % fields.Date.today().strftime('%B %Y'), date_format)
        
        # Set row heights
        sheet.set_row(0, 25)
        sheet.set_row(1, 22)
        sheet.set_row(2, 20)
    
    def _export_gross_profit(self, workbook):
        sheet = workbook.add_worksheet('Gross Profit')
        
        # Write header
        self._write_excel_header(sheet, workbook, 'GROSS PROFIT REPORT', self.env.company.name)
        
        # Formats
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        num_format = workbook.add_format({'border': 1, 'num_format': '#,##0', 'valign': 'vcenter'})
        percent_format = workbook.add_format({'border': 1, 'num_format': '0.00%', 'valign': 'vcenter'})
        
        # Headers - starting from row 5 (after header section)
        headers = ['Category Items', 'Qty', 'Amount', 'COGS', 'GP %', 'Total Gross Profit']
        for col, header in enumerate(headers):
            sheet.write(5, col, header, header_format)
            sheet.set_column(col, col, 18)
        sheet.set_column(0, 0, 25)  # Category column wider
        
        # Data
        records = self.env['x_gross.profit'].search([])
        _logger.info("[EXPORT] Gross Profit records: %s", len(records))
        for row, rec in enumerate(records, 6):
            sheet.write(row, 0, rec.category_items or '', cell_format)
            sheet.write(row, 1, rec.qty or 0, num_format)
            sheet.write(row, 2, rec.amount or 0, num_format)
            sheet.write(row, 3, rec.cogs or 0, num_format)
            sheet.write(row, 4, (rec.gp_percent or 0) / 100, percent_format)
            sheet.write(row, 5, rec.total_gross_profit or 0, num_format)
    
    def _export_rekap_so(self, workbook):
        sheet = workbook.add_worksheet('Rekap SO Payment')
        
        # Write header
        self._write_excel_header(sheet, workbook, 'REKAP SO (SALES ORDER TO PAYMENT)', self.env.company.name)
        
        # Formats
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#70AD47', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        date_format = workbook.add_format({'border': 1, 'num_format': 'DD/MM/YYYY', 'valign': 'vcenter'})
        num_format = workbook.add_format({'border': 1, 'num_format': '#,##0', 'valign': 'vcenter'})
        
        # Headers - starting from row 5
        headers = [
            'SO Number', 'PO Date', 'Customer', 'Company', 'Customer PO', 'Salesperson',
            'Product', 'Qty', 'Price Unit', 'Subtotal (Before Tax)', 'Tax', 'Total Amount',
            'Delivery No', 'Delivery Date', 'Delivery Status', 'Delivered Qty',
            'Branch', 'Receiver', 'Shipping Note',
            'Invoice Status', 'Invoice No', 'Invoice Date',
            'Payment Date', 'Payment State'
        ]
        for col, header in enumerate(headers):
            sheet.write(5, col, header, header_format)
            sheet.set_column(col, col, 15)
        sheet.set_column(2, 2, 20)  # Customer
        sheet.set_column(4, 4, 18)  # Customer PO
        sheet.set_column(6, 6, 25)  # Product
        sheet.set_column(18, 18, 20)  # Shipping Note
        
        # Build domain for date filtering
        domain = []
        if self.date_from:
            domain.append(('po_date', '>=', self.date_from))
        if self.date_to:
            domain.append(('po_date', '<=', self.date_to))
        
        records = self.env['x_rekap.so.payment'].search(domain)
        _logger.info("[EXPORT] Rekap SO records: %s", len(records))
        
        for row, rec in enumerate(records, 6):
            sheet.write(row, 0, rec.so_number or '', cell_format)
            sheet.write_datetime(row, 1, rec.po_date, date_format) if rec.po_date else sheet.write(row, 1, '', cell_format)
            sheet.write(row, 2, rec.customer or '', cell_format)
            sheet.write(row, 3, rec.company_name or '', cell_format)
            sheet.write(row, 4, rec.customer_po_number or '', cell_format)
            sheet.write(row, 5, rec.salesperson or '', cell_format)
            sheet.write(row, 6, rec.product_name or '', cell_format)
            sheet.write(row, 7, rec.qty or 0, num_format)
            sheet.write(row, 8, rec.price_unit or 0, num_format)
            sheet.write(row, 9, rec.subtotal or 0, num_format)
            sheet.write(row, 10, rec.tax_amount or 0, num_format)
            sheet.write(row, 11, rec.total_amount or 0, num_format)
            sheet.write(row, 12, rec.delivery_number or '', cell_format)
            sheet.write_datetime(row, 13, rec.delivery_date, date_format) if rec.delivery_date else sheet.write(row, 13, '', cell_format)
            sheet.write(row, 14, rec.delivery_status or '', cell_format)
            sheet.write(row, 15, rec.delivered_qty or 0, num_format)
            sheet.write(row, 16, rec.branch_delivery or '', cell_format)
            sheet.write(row, 17, rec.receiver or '', cell_format)
            sheet.write(row, 18, rec.shipping_note or '', cell_format)
            sheet.write(row, 19, rec.invoice_status or '', cell_format)
            sheet.write(row, 20, rec.invoice_number or '', cell_format)
            sheet.write_datetime(row, 21, rec.invoice_date, date_format) if rec.invoice_date else sheet.write(row, 21, '', cell_format)
            sheet.write_datetime(row, 22, rec.payment_date, date_format) if rec.payment_date else sheet.write(row, 22, '', cell_format)
            sheet.write(row, 23, rec.payment_state or '', cell_format)
    
    def _export_sales_contribution(self, workbook):
        sheet = workbook.add_worksheet('Sales Contribution')
        
        # Write header
        self._write_excel_header(sheet, workbook, 'SALES CONTRIBUTION REPORT', self.env.company.name)
        
        # Formats
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#ED7D31', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        num_format = workbook.add_format({'border': 1, 'num_format': '#,##0', 'valign': 'vcenter'})
        percent_format = workbook.add_format({'border': 1, 'num_format': '0.00%', 'valign': 'vcenter'})
        total_format = workbook.add_format({
            'bold': True, 'border': 1, 'num_format': '#,##0', 'bg_color': '#F2F2F2'
        })
        
        # Headers - starting from row 5
        headers = ['Category', 'Sales Amount', 'COGS', 'Gross Profit', 'Margin %', 'Sales Contribution %']
        for col, header in enumerate(headers):
            sheet.write(5, col, header, header_format)
            sheet.set_column(col, col, 18)
        sheet.set_column(0, 0, 25)  # Category column wider
        
        # Data
        records = self.env['x_sales.contribution'].search([])
        _logger.info("[EXPORT] Sales Contribution records: %s", len(records))
        
        total_sales = 0
        total_cogs = 0
        total_gp = 0
        
        for row, rec in enumerate(records, 6):
            sheet.write(row, 0, rec.category or '', cell_format)
            sheet.write(row, 1, rec.sales_amount or 0, num_format)
            sheet.write(row, 2, rec.cogs or 0, num_format)
            sheet.write(row, 3, rec.gross_profit or 0, num_format)
            sheet.write(row, 4, (rec.margin_percent or 0) / 100, percent_format)
            sheet.write(row, 5, (rec.sales_contribution_percent or 0) / 100, percent_format)
            
            total_sales += rec.sales_amount or 0
            total_cogs += rec.cogs or 0
            total_gp += rec.gross_profit or 0
        
        # Summary row
        summary_row = 6 + len(records) + 1
        sheet.write(summary_row, 0, 'TOTAL', total_format)
        sheet.write(summary_row, 1, total_sales, total_format)
        sheet.write(summary_row, 2, total_cogs, total_format)
        sheet.write(summary_row, 3, total_gp, total_format)
        sheet.write(summary_row, 4, (total_gp / total_sales) if total_sales else 0, percent_format)
        sheet.write(summary_row, 5, 1.0, percent_format)
