# -*- coding: utf-8 -*-
"""
Demo Data Script - Sales Order sampai Payment Report
Odoo 18 Compatible

Cara penggunaan:
    python odoo-bin shell -d SalesReport --no-http < demo_data.py

Atau dari direktori odoo:
    ./odoo-bin shell -d SalesReport --no-http < /Users/mac/Documents/Dev/work/sale_recap_report/demo_data.py
"""

from odoo import fields
from odoo.exceptions import UserError
import logging

_logger = logging.getLogger(__name__)

# ============================================================
# UTILS
# ============================================================

def commit(msg=''):
    env.cr.commit()
    print(f"  ✓ Committed: {msg}")


def get_or_create_partner(name):
    partner = env['res.partner'].search([('name', '=', name)], limit=1)
    if not partner:
        partner = env['res.partner'].create({'name': name, 'is_company': True})
        print(f"  Created customer: {name}")
    else:
        print(f"  Found customer:   {name}")
    return partner


def get_or_create_product(name, price):
    tmpl = env['product.template'].search([('name', '=', name)], limit=1)
    if not tmpl:
        tmpl = env['product.template'].create({
            'name': name,
            'type': 'consu',           # Odoo 18: storable
            'is_storable': True,       # Odoo 18 flag for storable
            'list_price': price,
            'uom_id': env.ref('uom.product_uom_unit').id,
            'uom_po_id': env.ref('uom.product_uom_unit').id,
        })
        print(f"  Created product:  {name} @ {price:,.0f}")
    else:
        print(f"  Found product:    {name}")
    return tmpl.product_variant_id


def add_stock(product, qty):
    """Add stock to WH/Stock location"""
    location = env.ref('stock.stock_location_stock')
    env['stock.quant']._update_available_quantity(product, location, qty)
    print(f"  Stock added:      {product.display_name} +{qty} units")


def get_bank_journal():
    journal = env['account.journal'].search([
        ('type', '=', 'bank'),
        ('company_id', '=', env.company.id)
    ], limit=1)
    if not journal:
        raise UserError("Tidak ada bank journal! Buat dulu di Accounting → Configuration → Journals.")
    return journal


def create_so(customer, lines, ref=None):
    """Create and return a Sale Order"""
    order_lines = []
    for product, qty in lines:
        order_lines.append((0, 0, {
            'product_id': product.id,
            'product_uom_qty': qty,
            'price_unit': product.lst_price,
        }))
    so = env['sale.order'].create({
        'partner_id': customer.id,
        'client_order_ref': ref or '',
        'order_line': order_lines,
    })
    return so


def validate_picking(picking, qty_per_move=None):
    """
    Validate a picking. Odoo 18: set move.quantity directly.
    qty_per_move: dict {move_id: qty} for partial delivery.
    """
    picking.action_assign()  # reserve stock

    for move in picking.move_ids.filtered(lambda m: m.state not in ('done', 'cancel')):
        qty = (qty_per_move or {}).get(move.id, move.product_uom_qty)
        move.quantity = qty  # Odoo 18: sets done quantity directly

    result = picking.button_validate()

    # Handle backorder wizard
    if isinstance(result, dict) and result.get('res_model') == 'stock.backorder.confirmation':
        ctx = result.get('context', {})
        wiz = env['stock.backorder.confirmation'].with_context(**ctx).create({
            'pick_ids': [(4, picking.id)],
        })
        wiz.process()  # create backorder

    return picking


def create_invoice_and_post(so):
    """Create invoice from SO and post it"""
    so._create_invoices()
    inv = so.invoice_ids.filtered(lambda i: i.state == 'draft')
    if not inv:
        inv = so.invoice_ids[0]
    inv.action_post()
    return inv


def register_payment(invoice, amount=None):
    """Register full or partial payment"""
    ctx = {
        'active_model': 'account.move',
        'active_ids': invoice.ids,
        'active_id': invoice.id,
    }
    vals = {
        'journal_id': get_bank_journal().id,
        'payment_date': fields.Date.today(),
    }
    if amount:
        vals['amount'] = amount

    wizard = env['account.payment.register'].with_context(**ctx).create(vals)
    wizard.action_create_payments()
    invoice.invalidate_recordset()
    print(f"  Payment state:    {invoice.payment_state}")


# ============================================================
# 1. MASTER DATA
# ============================================================
print("\n" + "="*60)
print("STEP 1: Creating Master Data")
print("="*60)

# Customers
customer1 = get_or_create_partner('PT Demo Palm Industry')
customer2 = get_or_create_partner('PT Agro Sawit Nusantara')
customer3 = get_or_create_partner('PT Mitra Perkasa')

# Products
p_desk     = get_or_create_product('Desk Combination', 2_500_000)
p_hook     = get_or_create_product('Harvesting Hook',    450_000)
p_knife    = get_or_create_product('Oil Palm Knife',     250_000)
p_sarung   = get_or_create_product('Sarung Tojok',        80_000)

# Add stock (1000 units each)
add_stock(p_desk,   1000)
add_stock(p_hook,   1000)
add_stock(p_knife,  1000)
add_stock(p_sarung, 1000)

commit("Master data")


# ============================================================
# SCENARIO 1: Simple Flow
# SO → 1 DO (full) → Invoice → PAID
# ============================================================
print("\n" + "="*60)
print("SCENARIO 1: Simple Flow (Paid)")
print("="*60)

so1 = create_so(customer1, [(p_desk, 40)], ref='PO-DEMO-001')
so1.action_confirm()
print(f"  SO Created: {so1.name}")

picking1 = so1.picking_ids[0]
validate_picking(picking1)
print(f"  DO Validated: {picking1.name} — state: {picking1.state}")

inv1 = create_invoice_and_post(so1)
print(f"  Invoice: {inv1.name}")
register_payment(inv1)  # Full payment

commit("Scenario 1")


# ============================================================
# SCENARIO 2: Partial Delivery (Multi DO)
# SO (100 units) → DO1 (40) + backorder DO2 (60) → Invoice → PARTIAL PAYMENT
# ============================================================
print("\n" + "="*60)
print("SCENARIO 2: Partial Delivery — Multi DO (Partial Payment)")
print("="*60)

so2 = create_so(customer2, [(p_hook, 100)], ref='PO-DEMO-002')
so2.action_confirm()
print(f"  SO Created: {so2.name}")

picking2 = so2.picking_ids[0]
move2_id = picking2.move_ids[0].id
validate_picking(picking2, qty_per_move={move2_id: 40})
print(f"  DO1 Validated: {picking2.name} (40 units) — state: {picking2.state}")

# Backorder picking
picking3 = so2.picking_ids.filtered(lambda p: p.state not in ('done', 'cancel'))[0]
validate_picking(picking3)
print(f"  DO2 Validated: {picking3.name} (60 units) — state: {picking3.state}")

inv2 = create_invoice_and_post(so2)
print(f"  Invoice: {inv2.name}")
# Partial payment: 50% of total
register_payment(inv2, amount=round(inv2.amount_total * 0.5))

commit("Scenario 2")


# ============================================================
# SCENARIO 3: Multi Product — 1 DO → Invoice → NOT PAID
# SO (3 products) → 1 DO → Invoice (no payment)
# ============================================================
print("\n" + "="*60)
print("SCENARIO 3: Multi Product (Not Paid)")
print("="*60)

so3 = create_so(customer3, [
    (p_desk,   20),
    (p_knife,  50),
    (p_sarung, 30),
], ref='PO-DEMO-003')
so3.action_confirm()
print(f"  SO Created: {so3.name}")

picking4 = so3.picking_ids[0]
validate_picking(picking4)
print(f"  DO Validated: {picking4.name} (3 products) — state: {picking4.state}")

inv3 = create_invoice_and_post(so3)
print(f"  Invoice: {inv3.name} (NO payment — status: {inv3.payment_state})")

commit("Scenario 3")


# ============================================================
# SCENARIO 4: Multi Delivery → 1 Invoice → PAID
# SO (80 units) → DO1 (30) + DO2 (50) → 1 invoice → full payment
# ============================================================
print("\n" + "="*60)
print("SCENARIO 4: Multi Delivery → 1 Invoice (Paid)")
print("="*60)

so4 = create_so(customer1, [(p_knife, 80)], ref='PO-DEMO-004')
so4.action_confirm()
print(f"  SO Created: {so4.name}")

picking5 = so4.picking_ids[0]
move5_id = picking5.move_ids[0].id
validate_picking(picking5, qty_per_move={move5_id: 30})
print(f"  DO1 Validated: {picking5.name} (30 units) — state: {picking5.state}")

picking6 = so4.picking_ids.filtered(lambda p: p.state not in ('done', 'cancel'))[0]
validate_picking(picking6)
print(f"  DO2 Validated: {picking6.name} (50 units) — state: {picking6.state}")

# 1 invoice for both deliveries
inv4 = create_invoice_and_post(so4)
print(f"  Invoice: {inv4.name} (covers both DOs)")
register_payment(inv4)  # Full payment

commit("Scenario 4")


# ============================================================
# SUMMARY
# ============================================================
print("\n" + "="*60)
print("DEMO DATA SUMMARY")
print("="*60)
print(f"""
  Sales Orders:
    {so1.name} — PT Demo Palm Industry    — Desk Combination x40      — PAID
    {so2.name} — PT Agro Sawit Nusantara  — Harvesting Hook x100      — PARTIAL
    {so3.name} — PT Mitra Perkasa         — 3 Products                — NOT PAID
    {so4.name} — PT Demo Palm Industry    — Oil Palm Knife x80        — PAID

  Delivery Orders:
    {picking1.name} → {so1.name} (40 units, done)
    {picking2.name} → {so2.name} (40 units, done)
    {picking3.name} → {so2.name} (60 units, done)
    {picking4.name} → {so3.name} (3 products, done)
    {picking5.name} → {so4.name} (30 units, done)
    {picking6.name} → {so4.name} (50 units, done)

  Invoices:
    {inv1.name} — Paid (full)
    {inv2.name} — Partial payment
    {inv3.name} — Not paid
    {inv4.name} — Paid (covers 2 DOs)

  Expected Report Rows:
    {so1.name} | Desk Combination  | {picking1.name} | 40  | {inv1.name}
    {so2.name} | Harvesting Hook   | {picking2.name} | 40  | {inv2.name}
    {so2.name} | Harvesting Hook   | {picking3.name} | 60  | {inv2.name}
    {so3.name} | Desk Combination  | {picking4.name} | 20  | {inv3.name}
    {so3.name} | Oil Palm Knife    | {picking4.name} | 50  | {inv3.name}
    {so3.name} | Sarung Tojok      | {picking4.name} | 30  | {inv3.name}
    {so4.name} | Oil Palm Knife    | {picking5.name} | 30  | {inv4.name}
    {so4.name} | Oil Palm Knife    | {picking6.name} | 50  | {inv4.name}
""")

print("✅ All demo data created successfully!")
