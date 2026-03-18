# -*- coding: utf-8 -*-

from . import models
from . import reports

def pre_init_hook(env):
    """Drop existing views before install to avoid schema conflicts"""
    # Drop SQL views
    env.cr.execute("DROP VIEW IF EXISTS x_gross_profit CASCADE")
    env.cr.execute("DROP VIEW IF EXISTS x_rekap_so_payment CASCADE")
    env.cr.execute("DROP VIEW IF EXISTS x_sales_contribution CASCADE")
    env.cr.execute("DROP TABLE IF EXISTS x_sale_recap_export_excel CASCADE")
    
    # Delete existing ir.ui.view records to avoid validation errors during upgrade
    # This ensures views will be recreated after models are properly initialized
    env.cr.execute("""
        DELETE FROM ir_ui_view 
        WHERE name IN (
            'x.gross.profit.tree',
            'x.gross.profit.pivot', 
            'x.gross.profit.graph',
            'x.gross.profit.search',
            'x.rekap.so.payment.tree',
            'x.rekap.so.payment.pivot',
            'x.rekap.so.payment.graph',
            'x.rekap.so.payment.search',
            'x.sales.contribution.tree',
            'x.sales.contribution.pivot',
            'x.sales.contribution.graph',
            'x.sales.contribution.search'
        )
        AND model IN (
            'x_gross.profit',
            'x_rekap.so.payment', 
            'x_sales.contribution'
        )
    """)
    env.cr.commit()
