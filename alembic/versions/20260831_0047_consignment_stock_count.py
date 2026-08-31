"""Consignment stock counts - the physical shelf, recorded.

Revision ID: 20260831_0047_consignment_stock_count
Revises: 20260831_0046_consignment_sale_discount
Create Date: 2026-08-31

Stock on hand for a consignment client is derived from what was invoiced out
and how much of it has been paid for. That cannot know about a sale nobody
recorded, or goods paid for that never left the client's shelf, so the derived
figure drifts from the real one in both directions.

Two tables let an operator record what is actually there:

  * consignment_stock_counts       - who counted, for which client, when
  * consignment_stock_count_items  - the counted quantity per product

From a count's date forward it is the baseline: deliveries, reported sales and
returns after it are applied on top; everything before it is superseded. Added
defensively (skipped if present) so it coexists with the runtime schema guard
in ``app/app_factory.py``.
"""
from alembic import op
import sqlalchemy as sa

revision = "20260831_0047_consignment_stock_count"
down_revision = "20260831_0046_consignment_sale_discount"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    tables = set(sa.inspect(bind).get_table_names())

    if "consignment_stock_counts" not in tables:
        op.create_table(
            "consignment_stock_counts",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("client_id", sa.Integer(), sa.ForeignKey("b2b_clients.id"), nullable=False),
            sa.Column("user_id", sa.Integer(), sa.ForeignKey("users.id"), nullable=True),
            sa.Column("counted_at", sa.DateTime(timezone=True), nullable=False),
            sa.Column("notes", sa.Text(), nullable=True),
            sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now()),
        )
        op.create_index("ix_consignment_stock_counts_id", "consignment_stock_counts", ["id"])
        op.create_index("ix_consignment_stock_counts_client_id",
                        "consignment_stock_counts", ["client_id"])

    if "consignment_stock_count_items" not in tables:
        op.create_table(
            "consignment_stock_count_items",
            sa.Column("id", sa.Integer(), primary_key=True),
            sa.Column("count_id", sa.Integer(),
                      sa.ForeignKey("consignment_stock_counts.id"), nullable=False),
            sa.Column("product_id", sa.Integer(), sa.ForeignKey("products.id"), nullable=False),
            sa.Column("qty", sa.Numeric(12, 3), nullable=False),
        )
        op.create_index("ix_consignment_stock_count_items_id",
                        "consignment_stock_count_items", ["id"])
        op.create_index("ix_consignment_stock_count_items_count_id",
                        "consignment_stock_count_items", ["count_id"])


def downgrade() -> None:
    bind = op.get_bind()
    tables = set(sa.inspect(bind).get_table_names())
    if "consignment_stock_count_items" in tables:
        op.drop_table("consignment_stock_count_items")
    if "consignment_stock_counts" in tables:
        op.drop_table("consignment_stock_counts")
