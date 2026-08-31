import asyncio
from decimal import Decimal

import pytest
from fastapi import HTTPException
from sqlalchemy import create_engine, select
from sqlalchemy.orm import sessionmaker

from tests.env_defaults import apply_test_environment_defaults

apply_test_environment_defaults()

from app.core.log import ActivityLog
from app.database import Base
from app.models.accounting import Account, Journal, JournalEntry
from app.models.b2b import (
    B2BClient, B2BInvoice, B2BInvoiceItem, B2BRefund,
    Consignment, ConsignmentItem, ConsignmentSale, ConsignmentSaleItem,
)
from app.models.product import Product
from app.models.user import User
from app.routers.accounting import _record_consignment_client_payment
from app.schemas.invoice import ConsignmentSaleItemIn


class AsyncSessionAdapter:
    """Minimal async-shaped wrapper over a synchronous SQLite session — enough
    for _record_consignment_client_payment (execute / add / flush)."""

    def __init__(self, session):
        self.session = session

    async def execute(self, statement, params=None):
        return self.session.execute(statement, params or {})

    def add(self, obj):
        self.session.add(obj)

    async def flush(self):
        self.session.flush()

    async def commit(self):
        self.session.commit()


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
    Base.metadata.create_all(
        engine,
        tables=[
            User.__table__,
            Product.__table__,
            B2BClient.__table__,
            B2BInvoice.__table__,
            B2BInvoiceItem.__table__,
            # The outstanding balance nets off refunds, so this table is part
            # of the schema any payment path touches.
            B2BRefund.__table__,
            Consignment.__table__,
            ConsignmentItem.__table__,
            ConsignmentSale.__table__,
            ConsignmentSaleItem.__table__,
            Account.__table__,
            Journal.__table__,
            JournalEntry.__table__,
            ActivityLog.__table__,
        ],
    )
    Session = sessionmaker(bind=engine, expire_on_commit=False)
    return Session()


def _seed(session, discount_pct=0):
    for cid, code, name in [(1, "1000", "Cash"), (2, "1100", "AR"), (3, "2200", "Deferred"), (4, "4000", "Revenue")]:
        session.add(Account(id=cid, code=code, name=name, type="asset", balance=0))
    user = User(id=1, name="Op", role="admin", email="op@x.com", password="x")
    client = B2BClient(
        id=1, name="Cons Client", payment_terms="consignment",
        outstanding=Decimal("500.00"), discount_pct=Decimal(str(discount_pct)),
    )
    product = Product(id=1, sku="SKU-1", name="Dates", price=Decimal("100.00"), unit="kg", stock=100)
    invoice = B2BInvoice(
        id=1, client_id=1, invoice_number="HB2B-C-1", invoice_type="consignment",
        status="consignment", total=Decimal("500.00"), amount_paid=Decimal("0.00"),
    )
    session.add_all([user, client, product, invoice])
    session.commit()
    return user, client


def test_matching_amount_records_sale_and_items_and_allocates():
    with make_session() as session:
        user, client = _seed(session)
        db = AsyncSessionAdapter(session)
        items = [ConsignmentSaleItemIn(product_id=1, qty=3, unit_price=100.0)]  # 300.00

        payload = run(_record_consignment_client_payment(
            db, client=client, amount=300.0, month_label="July 2026",
            current_user=user, sold_items=items,
        ))
        session.commit()

        assert payload["amount"] == 300.0
        assert payload["sale_id"] is not None
        # client outstanding reduced 500 -> 200
        assert round(float(client.outstanding), 2) == 200.0
        # invoice partially paid
        inv = session.execute(select(B2BInvoice).where(B2BInvoice.id == 1)).scalar_one()
        assert round(float(inv.amount_paid), 2) == 300.0
        assert inv.status == "partial"
        # sale + line item persisted
        sale = session.execute(select(ConsignmentSale)).scalar_one()
        assert sale.month_label == "July 2026"
        assert round(float(sale.amount), 2) == 300.0
        line = session.execute(select(ConsignmentSaleItem)).scalar_one()
        assert line.product_id == 1
        assert round(float(line.qty), 3) == 3.0
        assert round(float(line.total), 2) == 300.0


def test_mismatched_amount_is_rejected():
    with make_session() as session:
        user, client = _seed(session)
        db = AsyncSessionAdapter(session)
        items = [ConsignmentSaleItemIn(product_id=1, qty=3, unit_price=100.0)]  # 300.00

        with pytest.raises(HTTPException) as excinfo:
            run(_record_consignment_client_payment(
                db, client=client, amount=250.0, month_label="July 2026",
                current_user=user, sold_items=items,
            ))
        assert excinfo.value.status_code == 400
        assert "does not match" in excinfo.value.detail
        # nothing recorded
        assert session.execute(select(ConsignmentSale)).first() is None


def test_no_items_behaves_as_amount_only_payment():
    with make_session() as session:
        user, client = _seed(session)
        db = AsyncSessionAdapter(session)

        payload = run(_record_consignment_client_payment(
            db, client=client, amount=200.0, month_label="July 2026",
            current_user=user, sold_items=None,
        ))
        session.commit()

        assert payload["sale_id"] is None
        assert session.execute(select(ConsignmentSale)).first() is None
        assert round(float(client.outstanding), 2) == 300.0


def test_discount_is_applied_to_the_sold_items_total():
    # The consignment invoice was booked net of the client's discount, so the
    # payment must collect the sold items net of that same discount.
    with make_session() as session:
        user, client = _seed(session, discount_pct=10)
        db = AsyncSessionAdapter(session)
        items = [ConsignmentSaleItemIn(product_id=1, qty=3, unit_price=100.0)]  # 300.00 gross

        payload = run(_record_consignment_client_payment(
            db, client=client, amount=270.0, month_label="July 2026",
            current_user=user, sold_items=items,
        ))
        session.commit()

        assert payload["subtotal"] == 300.0
        assert payload["discount"] == 30.0
        assert payload["discount_pct"] == 10.0
        assert payload["amount"] == 270.0
        # only the net amount comes off the balances
        assert round(float(client.outstanding), 2) == 230.0
        inv = session.execute(select(B2BInvoice).where(B2BInvoice.id == 1)).scalar_one()
        assert round(float(inv.amount_paid), 2) == 270.0
        # the sale record keeps gross, discount and net
        sale = session.execute(select(ConsignmentSale)).scalar_one()
        assert round(float(sale.subtotal), 2) == 300.0
        assert round(float(sale.discount), 2) == 30.0
        assert round(float(sale.amount), 2) == 270.0
        # line items stay at their gross consignment price
        line = session.execute(select(ConsignmentSaleItem)).scalar_one()
        assert round(float(line.total), 2) == 300.0


def test_gross_amount_is_rejected_for_a_discounted_client():
    with make_session() as session:
        user, client = _seed(session, discount_pct=10)
        db = AsyncSessionAdapter(session)
        items = [ConsignmentSaleItemIn(product_id=1, qty=3, unit_price=100.0)]

        with pytest.raises(HTTPException) as excinfo:
            run(_record_consignment_client_payment(
                db, client=client, amount=300.0, month_label="July 2026",
                current_user=user, sold_items=items,
            ))
        assert excinfo.value.status_code == 400
        assert "270.00" in excinfo.value.detail
        assert "10% client discount" in excinfo.value.detail
        assert session.execute(select(ConsignmentSale)).first() is None


def test_zero_discount_client_pays_the_gross_items_total():
    with make_session() as session:
        user, client = _seed(session)
        db = AsyncSessionAdapter(session)
        items = [ConsignmentSaleItemIn(product_id=1, qty=2, unit_price=100.0)]

        payload = run(_record_consignment_client_payment(
            db, client=client, amount=200.0, month_label=None,
            current_user=user, sold_items=items,
        ))
        session.commit()

        assert payload["discount"] == 0.0
        assert payload["amount"] == 200.0
        sale = session.execute(select(ConsignmentSale)).scalar_one()
        assert round(float(sale.subtotal), 2) == 200.0
        assert round(float(sale.discount), 2) == 0.0
