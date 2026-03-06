# -*- coding: utf-8 -*-
{
    'name': 'Sales Recap Report',
    'version': '18.0.4.1.0',
    'category': 'Sales',
    'summary': '3 Report: Gross Profit, Rekap SO (sampai Payment), Sales Contribution',
    'description': """
Sales Recap Report
==================
Complete Sales Reporting Module with 3 Reports:

1. GROSS PROFIT REPORT
   - Category Items, Qty, Amount, COGS, GP %, Total Gross Profit
   - Proper COGS calculation from product.standard_price
   - Pivot and Graph views for analysis
   
2. REKAP SO (Sales Order to Payment)
   - SO Number, PO Date, Customer, Company, Customer PO, Salesperson
   - Product, Qty, Price Unit, Subtotal, Tax, Total Amount
   - Delivery No, Delivery Date, Delivery Status, Delivered Qty
   - Branch, Receiver, Shipping Note
   - Invoice Status, Invoice No, Invoice Date
   - Payment Date, Payment State
   - Complete flow tracking from SO to Payment

3. SALES CONTRIBUTION REPORT
   - Category, Sales Amount, COGS, Gross Profit
   - Margin %, Sales Contribution %
   - Proper COGS and Profit calculations
   - Category-level summary

Features:
---------
- Export Excel with professional headers
- Filter by Date, Customer, Salesperson, Product Category
- List View, Pivot View, and Graph View for all reports
- Date range filtering for all reports
- Enhanced search with grouping options
- Payment tracking with actual payment dates
- Proper Gross Profit calculations using product cost

This module fully replaces Excel-based reporting workflows.
    """,
    'author': 'PT Injani',
    'website': 'https://www.injani.com/id/',
    'depends': [
        'base',
        'sale',
        'sale_management',
        'sale_stock',       # Required: adds sale_line_id column to stock_move
        'stock',
        'account',
    ],
    'data': [
        'security/ir.model.access.csv',
        'views/sale_recap_views.xml',
        'views/menu.xml',
        'reports/export_excel.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'LGPL-3',
}
