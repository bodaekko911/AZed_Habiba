"""B2B client refunds — the balance owed, and parity between the two screens.

Two defects are pinned here:

1. The clients list computed outstanding from unpaid invoices alone, so a
   refund left the figure unchanged on screen — even though the statement and
   the client analysis both already netted refunds off. The three disagreed.

2. The Accounting page raised a refund by writing a journal and decrementing a
   stored balance field. No refund record, no line items, no stock returned,
   and nothing visible in the B2B refunds list.
"""

import asyncio
from datetime import date, datetime, timezone
from decimal import Decimal
from types import SimpleNamespace

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from tests.env_defaults import apply_test_environment_defaults

apply_test_environment_defaults()

import app.routers.accounting as accounting
import app.routers.b2b as b2b
from app.core.log import ActivityLog
from app.database import Base
from app.models.accounting import Account, Journal, JournalEntry
from app.models.b2b import (
    B2BClient, B2BClientPrice, B2BInvoice, B2BInvoiceItem, B2BRefund, B2BRefundItem,
)
from app.models.inventory import StockMove
from app.models.product import Product
from app.models.user import User
from app.schemas.invoice import B2BPaymentRequest


class AsyncSessionAdapter:
    def __init__(self, session):
        self.session = session

    async def execute(self, statement, params=None):
        return self.session.execute(statement, params or {})

    async def commit(self):
        self.session.commit()

    async def rollback(self):
        self.session.rollback()

    async def refresh(self, obj):
        self.session.refresh(obj)

    async def flush(self):
        self.session.flush()

    def add(self, obj):
        self.session.add(obj)


def run(coro):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(coro)
    finally:
        loop.close()
        asyncio.set_event_loop(asyncio.new_event_loop())


def make_session():
    engine = create_engine("sqlite:///:memory:")
    Base.metadata.create_all(engine, tables=[
        User.__table__, Product.__table__, ActivityLog.__table__,
        Account.__table__, Journal.__table__, JournalEntry.__table__,
        StockMove.__table__,
        B2BClient.__table__, B2BClientPrice.__table__,
        B2BInvoice.__table__, B2BInvoiceItem.__table__,
        B2BRefund.__table__, B2BRefundItem.__table__,
    ])
    Session = sessionmaker(bind=engine, expire_on_commit=False)
    return Session()


USER = SimpleNamespace(id=1, name="Admin", role="admin")


def seed(session):
    """Green Cafe owes 1,000: a 1,000 unpaid invoice for 50 kg of tomato."""
    session.add_all([
        User(id=1, name="Admin", email="a@x.example", password="x", role="admin"),
        Product(id=1, sku="TOM", name="Tomato", price=Decimal("20"), unit="kg",
                stock=Decimal("100"), item_type="finished"),
        B2BClient(id=1, name="Green Cafe", is_active=True,
                  outstanding=Decimal("1000"), discount_pct=Decimal("0")),
        Account(id=1, code="2200", name="AR", type="asset"),
        Account(id=2, code="1100", name="Sales", type="revenue"),
    ])
    session.flush()
    session.add(B2BInvoice(id=1, invoice_number="B2B-00001", client_id=1,
                           invoice_type="credit", status="unpaid",
                           subtotal=Decimal("1000"), total=Decimal("1000"),
                           amount_paid=Decimal("0"),
                           created_at=datetime(2026, 8, 2, tzinfo=timezone.utc)))
    session.flush()
    session.add(B2BInvoiceItem(invoice_id=1, product_id=1, qty=Decimal("50"),
                               unit_price=Decimal("20"), total=Decimal("1000")))
    session.commit()


def listed_outstanding(session, client_id=1):
    rows = run(b2b.get_clients(q="", db=AsyncSessionAdapter(session)))
    return next(r["outstanding"] for r in rows if r["id"] == client_id)


def b2b_refund(session, qty=10, price=20, notes=None):
    payload = b2b.ClientRefundCreate(
        client_id=1, notes=notes,
        items=[b2b.RefundItemIn(product_id=1, qty=qty, unit_price=price)],
    )
    return run(b2b.create_client_refund_core(AsyncSessionAdapter(session), USER, payload))


def accounting_refund(session, qty=10, price=20, reason=None):
    payload = accounting.B2BRefundIn(
        items=[accounting.B2BRefundItemIn(product_id=1, qty=qty, unit_price=price)],
        reason=reason,
    )
    return run(accounting.refund_b2b_client_account(
        1, payload, AsyncSessionAdapter(session), USER,
    ))


# ── 1. Outstanding reflects refunds ──────────────────────────────────────────

def test_refund_reduces_the_outstanding_shown_on_the_clients_list():
    with make_session() as session:
        seed(session)
        assert listed_outstanding(session) == 1000.0
        b2b_refund(session, qty=10, price=20)          # 200 returned
        after = listed_outstanding(session)

    assert after == 800.0


def test_outstanding_never_goes_negative():
    with make_session() as session:
        seed(session)
        # Pay the invoice off, then refund goods against it
        invoice = session.get(B2BInvoice, 1)
        invoice.amount_paid = Decimal("1000")
        invoice.status = "paid"
        session.commit()
        session.add(B2BRefund(id=9, refund_number="BRF-9", client_id=1,
                              subtotal=Decimal("300"), total=Decimal("300"),
                              created_at=datetime(2026, 8, 9, tzinfo=timezone.utc)))
        session.commit()

        assert listed_outstanding(session) == 0.0


def settle(session, invoice_id=1):
    inv = session.get(B2BInvoice, invoice_id)
    inv.amount_paid = inv.total
    inv.status = "paid"
    session.commit()


def later_invoice(session, when, total=500):
    session.add(B2BInvoice(id=2, invoice_number="B2B-00002", client_id=1,
                           invoice_type="credit", status="unpaid",
                           subtotal=Decimal(str(total)), total=Decimal(str(total)),
                           amount_paid=Decimal("0"), created_at=when))
    session.flush()
    session.add(B2BInvoiceItem(invoice_id=2, product_id=1, qty=Decimal("25"),
                               unit_price=Decimal("20"), total=Decimal(str(total))))
    session.commit()


def test_a_refund_from_closed_history_no_longer_discounts_todays_balance():
    """A credit is spent once, against whatever was open when it was raised.

    Nothing links a refund row to an invoice, so subtracting every refund ever
    meant a months-old credit kept discounting the balance long after the
    invoices it belonged to were settled - the client permanently appeared to
    owe less than the sum of their open invoices.
    """
    with make_session() as session:
        seed(session)
        b2b_refund(session, qty=10, price=20)      # 200 credit, Aug
        settle(session)                            # its invoice is paid off
        later_invoice(session, datetime(2026, 10, 1, tzinfo=timezone.utc), total=500)

        # The October invoice is the only thing open; the August credit belongs
        # to history and must not come off it.
        assert listed_outstanding(session) == 500.0


def test_a_refund_against_a_still_open_invoice_does_credit_the_balance():
    with make_session() as session:
        seed(session)
        later_invoice(session, datetime(2026, 8, 1, tzinfo=timezone.utc), total=500)
        # The August invoice from seed() is still open, and the refund is
        # raised after it, so the credit is live.
        b2b_refund(session, qty=10, price=20)      # 200

        assert listed_outstanding(session) == 1300.0   # 1000 + 500 - 200


def test_a_fully_settled_account_owes_nothing_even_with_old_credits():
    with make_session() as session:
        seed(session)
        b2b_refund(session, qty=10, price=20)
        settle(session)

        assert listed_outstanding(session) == 0.0


def test_refund_guard_uses_the_live_balance_not_the_stored_field():
    """The stored client.outstanding drifts. A refund must be judged against
    the same balance the screen shows."""
    from fastapi import HTTPException
    with make_session() as session:
        seed(session)
        client = session.get(B2BClient, 1)
        client.outstanding = Decimal("0")      # stale — invoice still unpaid
        session.commit()

        result = b2b_refund(session, qty=5, price=20)   # would have been rejected
        assert result["amount"] == 100.0

    with make_session() as session:
        seed(session)
        try:
            b2b_refund(session, qty=100, price=20)      # 2,000 against 1,000 owed
        except HTTPException as exc:
            assert exc.status_code == 400
            assert "1000.00" in exc.detail
        else:
            raise AssertionError("a refund larger than the balance was accepted")


def test_reported_outstanding_matches_the_clients_list():
    with make_session() as session:
        seed(session)
        result = b2b_refund(session, qty=10, price=20)
        assert result["outstanding"] == listed_outstanding(session)


# ── 2. Accounting raises the same refund as the B2B page ─────────────────────

def test_accounting_refund_creates_a_real_refund_record():
    with make_session() as session:
        seed(session)
        result = accounting_refund(session, qty=10, price=20)

        refunds = session.query(B2BRefund).all()
        items = session.query(B2BRefundItem).all()

    assert len(refunds) == 1
    assert refunds[0].refund_number == result["refund_number"]
    assert float(refunds[0].total) == 200.0
    assert len(items) == 1
    assert float(items[0].qty) == 10.0


def test_accounting_refund_returns_the_goods_to_stock():
    with make_session() as session:
        seed(session)
        accounting_refund(session, qty=10, price=20)
        product = session.get(Product, 1)
        session.refresh(product)
        moves = session.query(StockMove).filter_by(ref_type="b2b_refund").all()

    assert float(product.stock) == 110.0            # 100 + 10 returned
    assert len(moves) == 1
    assert moves[0].type == "in"


def test_accounting_refund_reduces_the_outstanding():
    with make_session() as session:
        seed(session)
        result = accounting_refund(session, qty=10, price=20)
        after = listed_outstanding(session)

    assert after == 800.0
    assert result["client_outstanding"] == 800.0


def test_both_screens_produce_identical_refunds():
    def snapshot(fn):
        with make_session() as session:
            seed(session)
            result = fn(session)
            refund = session.query(B2BRefund).one()
            item = session.query(B2BRefundItem).one()
            product = session.get(Product, 1)
            session.refresh(product)
            return {
                "total": float(refund.total),
                "subtotal": float(refund.subtotal),
                "qty": float(item.qty),
                "stock": float(product.stock),
                "outstanding": result["outstanding"],
                "journals": session.query(Journal).filter_by(ref_type="b2b_refund").count(),
            }

    assert snapshot(b2b_refund) == snapshot(accounting_refund)


def test_accounting_refund_requires_items():
    from fastapi import HTTPException
    with make_session() as session:
        seed(session)
        payload = accounting.B2BRefundIn(items=[], reason=None)
        try:
            run(accounting.refund_b2b_client_account(
                1, payload, AsyncSessionAdapter(session), USER,
            ))
        except HTTPException as exc:
            assert exc.status_code == 400
        else:
            raise AssertionError("an empty refund was accepted")


# ── 3. Every screen agrees on what a client owes ─────────────────────────────

def accounting_listed_outstanding(session, client_id=1):
    rows = run(accounting.get_accounting_b2b_clients(q=None, db=AsyncSessionAdapter(session)))
    return next(r["outstanding"] for r in rows if r["id"] == client_id)


def accounting_invoice_outstanding(session, invoice_id=1):
    # Query(None) defaults are not resolved when calling the handler
    # directly, so pass the filter arguments explicitly.
    rows = run(accounting.get_b2b_invoices(
        invoice_type=None, status=None, search=None,
        from_date=None, to_date=None, db=AsyncSessionAdapter(session),
    ))
    return next(r["client_outstanding"] for r in rows if r["id"] == invoice_id)


def test_accounting_clients_list_reflects_refunds():
    with make_session() as session:
        seed(session)
        assert accounting_listed_outstanding(session) == 1000.0
        b2b_refund(session, qty=10, price=20)
        after = accounting_listed_outstanding(session)

    assert after == 800.0


def test_accounting_invoices_list_reflects_refunds():
    """This row read the stored client.outstanding, a third source that drifts
    independently of the two lists."""
    with make_session() as session:
        seed(session)
        client = session.get(B2BClient, 1)
        client.outstanding = Decimal("4321")        # deliberately wrong
        session.commit()

        assert accounting_invoice_outstanding(session) == 1000.0   # not 4321
        b2b_refund(session, qty=10, price=20)
        after = accounting_invoice_outstanding(session)

    assert after == 800.0


def test_all_four_screens_report_the_same_balance():
    with make_session() as session:
        seed(session)
        session.get(B2BClient, 1).outstanding = Decimal("999")     # stale field
        session.commit()
        b2b_refund(session, qty=10, price=20)

        values = {
            "b2b clients":          listed_outstanding(session),
            "accounting clients":   accounting_listed_outstanding(session),
            "accounting invoices":  accounting_invoice_outstanding(session),
            "single lookup":        run(b2b._client_outstanding_value(
                                        AsyncSessionAdapter(session), 1)),
        }

    assert len(set(values.values())) == 1, values
    assert values["b2b clients"] == 800.0


def test_consignment_payment_guard_uses_the_live_balance():
    """It checked the stored field, so a stale zero blocked legitimate
    payments and a stale high value let through more than was owed."""
    from fastapi import HTTPException
    with make_session() as session:
        seed(session)
        session.get(B2BClient, 1).outstanding = Decimal("0")       # stale
        session.commit()
        invoice = session.get(B2BInvoice, 1)
        invoice.invoice_type = "consignment"
        session.commit()

        payload = B2BPaymentRequest(amount=100.0)
        result = run(accounting.accounting_client_consignment_payment(
            1, payload, AsyncSessionAdapter(session), USER,
        ))
        assert result["ok"] is True


def test_consignment_balances_count_toward_outstanding():
    """Consignment invoices track AR like any other and their unpaid balance is
    money owed. Excluding the 'consignment' status under-reported what those
    clients owed on every screen."""
    with make_session() as session:
        seed(session)
        invoice = session.get(B2BInvoice, 1)
        invoice.invoice_type = "consignment"
        invoice.status = "consignment"
        session.commit()

        assert listed_outstanding(session) == 1000.0
        assert accounting_listed_outstanding(session) == 1000.0


def test_settled_consignment_invoice_leaves_nothing_outstanding():
    with make_session() as session:
        seed(session)
        invoice = session.get(B2BInvoice, 1)
        invoice.invoice_type = "consignment"
        invoice.status = "consignment"
        invoice.amount_paid = Decimal("1000")
        session.commit()

        assert listed_outstanding(session) == 0.0
