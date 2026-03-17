# -*- coding: utf-8 -*-

from . import models
from . import reports

def pre_init_hook(env):
    """Drop existing views before install to avoid schema conflicts"""
    env.cr.execute("DROP VIEW IF EXISTS x_gross_profit CASCADE")
    env.cr.execute("DROP VIEW IF EXISTS x_sales_contribution CASCADE")
    env.cr.execute("DROP TABLE IF EXISTS x_sale_recap_export_excel CASCADE")
    env.cr.commit()
