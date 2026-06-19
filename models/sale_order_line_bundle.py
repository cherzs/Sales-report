from odoo import api, models


class SaleOrderLineBundle(models.Model):
    _inherit = "sale.order.line.bundle"

    @api.depends(
        "product_id",
        "product_id.default_code",
        "product_id.name",
        "quantity",
        "uom_id",
        "uom_id.name",
    )
    def _compute_display_name(self):
        for rec in self:
            product = rec.product_id
            code = product.default_code or ""
            name = product.name or ""
            qty = rec.quantity or 0.0
            uom = rec.uom_id.name or ""

            if code:
                rec.display_name = "[%s] %s x%s %s" % (code, name, qty, uom)
            elif name:
                rec.display_name = "%s x%s %s" % (name, qty, uom)
            else:
                rec.display_name = "%s, %s" % (self._description, rec.id)
