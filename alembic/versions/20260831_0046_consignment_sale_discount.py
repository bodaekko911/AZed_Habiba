"""Consignment sale discount — gross subtotal and client discount on each payment.

Revision ID: 20260831_0046_consignment_sale_discount
Revises: 20260813_0045_rename_pcs_to_piece
Create Date: 2026-08-31

A consignment client's invoice total is already net of the client's agreed
discount, but the payment screen totalled the reported sold items at their
gross consignment unit price. Recording a payment therefore over-collected
against the invoice.

Two columns are added to ``consignment_sales`` so each payment keeps the full
picture:

  • subtotal — gross sum of the line items (qty x unit_price)
  • discount — the client's discount applied to that subtotal

``amount`` keeps its meaning of "what the client actually paid", which is now
subtotal - discount. Existing rows carry no discount, so their subtotal is
backfilled from amount and discount stays 0. Added defensively (skipped if
present) so it coexists with the runtime schema guard in ``app/app_factory.py``.
"""
from alembic import op
import sqlalchemy as sa

revision = "20260831_0046_consignment_sale_discount"
down_revision = "20260813_0045_rename_pcs_to_piece"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    if "consignment_sales" not in set(inspector.get_table_names()):
        return
    cols = {c["name"] for c in inspector.get_columns("consignment_sales")}

    if "subtotal" not in cols:
        op.add_column("consignment_sales", sa.Column("subtotal", sa.Numeric(14, 2), server_default="0"))
    if "discount" not in cols:
        op.add_column("consignment_sales", sa.Column("discount", sa.Numeric(14, 2), server_default="0"))

    # Payments recorded before the discount was applied were collected gross,
    # so their subtotal equals the amount already stored.
    op.execute(
        "UPDATE consignment_sales SET subtotal = COALESCE(amount, 0) "
        "WHERE subtotal IS NULL OR subtotal = 0"
    )


def downgrade() -> None:
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    if "consignment_sales" not in set(inspector.get_table_names()):
        return
    cols = {c["name"] for c in inspector.get_columns("consignment_sales")}
    if "discount" in cols:
        op.drop_column("consignment_sales", "discount")
    if "subtotal" in cols:
        op.drop_column("consignment_sales", "subtotal")
