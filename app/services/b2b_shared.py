"""
Shared helpers used by both app/routers/b2b.py and the B2B sales import service.
Extracted to avoid duplication and to allow the import service to pass
created_at / ref_id that the router doesn't need.
"""
from decimal import Decimal
from datetime import datetime
from typing import Optional

from sqlalchemy import select
from sqlalchemy.sql import func as sa_func
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.b2b import B2BInvoice, Consignment
from app.models.accounting import Account, Journal, JournalEntry


async def post_journal(
    db: AsyncSession,
    description: str,
    ref_type: str,
    entries: list,
    user_id: Optional[int] = None,
    created_at: Optional[datetime] = None,
    ref_id: Optional[int] = None,
) -> None:
    journal = Journal(ref_type=ref_type, description=description, user_id=user_id)
    if created_at is not None:
        journal.created_at = created_at
    if ref_id is not None:
        journal.ref_id = ref_id
    db.add(journal)
    await db.flush()
    for code, debit, credit in entries:
        _r = await db.execute(select(Account).where(Account.code == code))
        acc = _r.scalar_one_or_none()
        if acc:
            db.add(JournalEntry(
                journal_id=journal.id, account_id=acc.id,
                debit=debit, credit=credit,
            ))
            acc.balance += Decimal(str(debit)) - Decimal(str(credit))


async def seed_deferred_revenue(db: AsyncSession) -> None:
    """Ensure account 2200 Deferred Revenue exists."""
    _r = await db.execute(select(Account).where(Account.code == "2200"))
    if _r.scalar_one_or_none() is None:
        db.add(Account(
            code="2200", name="Deferred Revenue",
            type="liability", balance=Decimal("0"),
        ))
        await db.commit()


async def next_b2b_number(db: AsyncSession) -> str:
    _r = await db.execute(select(sa_func.max(B2BInvoice.id)))
    max_id = _r.scalar() or 0
    return f"B2B-{str(max_id + 1).zfill(5)}"


async def next_cons_number(db: AsyncSession) -> str:
    _r = await db.execute(select(sa_func.max(Consignment.id)))
    max_id = _r.scalar() or 0
    return f"CONS-{str(max_id + 1).zfill(4)}"


async def get_b2b_client_top_products(db: AsyncSession) -> dict:
    """
    Returns the top 5 products purchased by each client.
    Returns: {client_id: [{"product_id": int, "name": str, "total_qty": float}, ...]}
    """
    from app.models.b2b import B2BInvoice, B2BInvoiceItem
    from app.models.product import Product

    query = (
        select(
            B2BInvoice.client_id,
            Product.id.label("product_id"),
            Product.name.label("product_name"),
            sa_func.sum(B2BInvoiceItem.qty).label("total_qty")
        )
        .select_from(B2BInvoice)
        .join(B2BInvoiceItem, B2BInvoice.id == B2BInvoiceItem.invoice_id)
        .join(Product, B2BInvoiceItem.product_id == Product.id)
        .group_by(B2BInvoice.client_id, Product.id, Product.name)
    )
    result = await db.execute(query)

    client_products = {}
    for client_id, product_id, product_name, total_qty in result.all():
        if client_id not in client_products:
            client_products[client_id] = []
        client_products[client_id].append({
            "product_id": product_id,
            "name": product_name,
            "total_qty": float(total_qty)
        })

    # Sort by total_qty descending and take top 5 for each client
    for client_id in client_products:
        client_products[client_id].sort(key=lambda x: x["total_qty"], reverse=True)
        client_products[client_id] = client_products[client_id][:5]

    return client_products


# ── Client outstanding balance ───────────────────────────────────────────────
# One definition, used by every screen that shows what a client owes. It used
# to be re-derived in each place and they disagreed: the B2B clients list and
# the accounting clients list both counted unpaid invoices only, so a refund
# never showed up; the accounting invoices list read the stored
# B2BClient.outstanding field, which drifts because not every path maintains
# it. Refunds are credits against the account and must come off the balance.
#
#   outstanding = unpaid/partial invoice balances
#                 − refunds raised since the oldest still-open invoice   (min 0)
#
# That last qualifier matters. A refund credits the account once, against
# whatever was open when it was raised. Nothing links a refund row to an
# invoice, so subtracting every refund ever meant a credit from months back
# kept discounting today's balance long after the invoices it belonged to had
# been settled — the client appeared to owe less than the sum of their open
# invoices, forever. A refund raised before every currently-open invoice was
# issued belongs to closed history: its credit was consumed when those
# invoices were settled, so it no longer moves the current balance. It still
# appears on the statement, where the history belongs.

# Every status that is not fully settled. "consignment" belongs here: those
# invoices track AR like any other, their amount_paid is maintained as the
# client reports sales, and their unpaid balance is money owed — leaving them
# out under-reported what consignment clients owe on every screen.
UNPAID_INVOICE_STATUSES = ("unpaid", "partial", "consignment")


def open_invoice_since_subquery():
    """Per-client date of the oldest invoice that still owes something.

    Refunds older than this belong to invoices that have since been settled,
    so they no longer count against the current balance.
    """
    return (
        select(
            B2BInvoice.client_id.label("client_id"),
            sa_func.min(B2BInvoice.created_at).label("since"),
        )
        .where(
            B2BInvoice.status.in_(UNPAID_INVOICE_STATUSES),
            B2BInvoice.total > B2BInvoice.amount_paid,
        )
        .group_by(B2BInvoice.client_id)
        .subquery()
    )


def client_invoice_balance_subquery():
    """Per-client sum of what is still owed on unpaid/partial invoices."""
    return (
        select(
            B2BInvoice.client_id,
            sa_func.coalesce(
                sa_func.sum(B2BInvoice.total - B2BInvoice.amount_paid), 0
            ).label("outstanding"),
        )
        .where(B2BInvoice.status.in_(UNPAID_INVOICE_STATUSES))
        .group_by(B2BInvoice.client_id)
        .subquery()
    )


def client_refund_subquery():
    """Per-client total of refunds that still credit the current balance."""
    from app.models.b2b import B2BRefund

    since = open_invoice_since_subquery()
    return (
        select(
            B2BRefund.client_id,
            sa_func.coalesce(sa_func.sum(B2BRefund.total), 0).label("refunded"),
        )
        # An inner join drops clients with nothing open — their balance is
        # zero regardless, so a stale credit cannot push it negative.
        .join(since, since.c.client_id == B2BRefund.client_id)
        .where(B2BRefund.created_at >= since.c.since)
        .group_by(B2BRefund.client_id)
        .subquery()
    )


async def client_outstanding_value(db: AsyncSession, client_id: int) -> float:
    """What one client actually owes, by the definition above."""
    from app.models.b2b import B2BRefund

    invoiced = await db.execute(
        select(sa_func.coalesce(sa_func.sum(B2BInvoice.total - B2BInvoice.amount_paid), 0))
        .where(
            B2BInvoice.client_id == client_id,
            B2BInvoice.status.in_(UNPAID_INVOICE_STATUSES),
        )
    )
    since = await db.execute(
        select(sa_func.min(B2BInvoice.created_at))
        .where(
            B2BInvoice.client_id == client_id,
            B2BInvoice.status.in_(UNPAID_INVOICE_STATUSES),
            B2BInvoice.total > B2BInvoice.amount_paid,
        )
    )
    open_since = since.scalar()
    if open_since is None:
        return 0.0
    refunded = await db.execute(
        select(sa_func.coalesce(sa_func.sum(B2BRefund.total), 0))
        .where(
            B2BRefund.client_id == client_id,
            B2BRefund.created_at >= open_since,
        )
    )
    return max(float(invoiced.scalar() or 0) - float(refunded.scalar() or 0), 0.0)


async def client_outstanding_map(db: AsyncSession) -> dict[int, float]:
    """{client_id: outstanding} for every client that has invoices or refunds.
    Avoids a per-row query when rendering a list."""
    from app.models.b2b import B2BRefund

    invoiced = await db.execute(
        select(
            B2BInvoice.client_id,
            sa_func.coalesce(sa_func.sum(B2BInvoice.total - B2BInvoice.amount_paid), 0),
        )
        .where(B2BInvoice.status.in_(UNPAID_INVOICE_STATUSES))
        .group_by(B2BInvoice.client_id)
    )
    balances: dict[int, float] = {cid: float(amt or 0) for cid, amt in invoiced.all()}

    since = client_refund_subquery()
    refunded = await db.execute(select(since.c.client_id, since.c.refunded))
    for cid, amt in refunded.all():
        balances[cid] = balances.get(cid, 0.0) - float(amt or 0)

    return {cid: max(value, 0.0) for cid, value in balances.items()}
