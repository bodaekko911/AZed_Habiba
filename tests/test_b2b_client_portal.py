"""B2B client portal — the shareable, login-free account link.

Two things are worth pinning here:

1. **Access control.** The token is the only credential, so every way a link can
   be dead (unknown, revoked, disabled, inactive client) must 404, and the
   portal must be reachable with no session at all.

2. **The netting rules** behind "products received". A consignment invoice
   writes the same lines onto both the invoice and the consignment, so counting
   both would double every consignment delivery; returns come from two
   unrelated places (refunds and the settle flow) and must not overlap.
"""

import asyncio
from collections.abc import AsyncGenerator
from datetime import datetime, timezone
from decimal import Decimal

from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from tests.env_defaults import apply_test_environment_defaults

apply_test_environment_defaults()

import app.app_factory as app_factory
import app.routers.b2b as b2b
from app.app_factory import create_app
from app.database import Base, get_async_session
from app.core.log import ActivityLog
from app.models.accounting import Account, Journal, JournalEntry
from app.models.b2b import (
    B2BClient,
    B2BInvoice,
    B2BInvoiceItem,
    B2BRefund,
    B2BRefundItem,
    Consignment,
    ConsignmentItem,
    ConsignmentSale,
    ConsignmentSaleItem,
    ConsignmentStockCount,
    ConsignmentStockCountItem,
)
from app.models.product import Product
from app.models.user import User


class AsyncSessionAdapter:
    """Sync SQLAlchemy session behind the async API the routers expect."""

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


TABLES = [
    User.__table__,
    ActivityLog.__table__,
    Product.__table__,
    Account.__table__,
    Journal.__table__,
    JournalEntry.__table__,
    B2BClient.__table__,
    B2BInvoice.__table__,
    B2BInvoiceItem.__table__,
    Consignment.__table__,
    ConsignmentItem.__table__,
    ConsignmentSale.__table__,
    ConsignmentSaleItem.__table__,
    ConsignmentStockCount.__table__,
    ConsignmentStockCountItem.__table__,
    B2BRefund.__table__,
    B2BRefundItem.__table__,
]


def make_session():
    # StaticPool + check_same_thread=False: TestClient runs the app on its own
    # thread, and a default in-memory SQLite connection is bound to the thread
    # that created it.
    engine = create_engine(
        "sqlite:///:memory:",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    Base.metadata.create_all(engine, tables=TABLES)
    Session = sessionmaker(bind=engine, expire_on_commit=False)
    return Session()


def seed(session, *, portal_token="tok_" + "a" * 40, portal_enabled=True, client_active=True):
    client = B2BClient(
        id=1, name="Green Cafe", phone="0100", email="hi@green.example",
        payment_terms="full_payment", credit_limit=Decimal("10000"),
        is_active=client_active,
        portal_token=portal_token, portal_enabled=portal_enabled,
        portal_created_at=datetime(2026, 8, 1, tzinfo=timezone.utc), portal_view_count=0,
    )
    other = B2BClient(id=2, name="Rival Bakery", is_active=True,
                      portal_token="tok_" + "b" * 40, portal_enabled=True)
    tomato = Product(id=1, sku="TOM", name="Tomato", price=Decimal("20"), unit="kg")
    herb = Product(id=2, sku="HRB", name="Herb bunch", price=Decimal("5"), unit="pcs")
    session.add_all([client, other, tomato, herb])
    session.flush()

    # 1) A plain credit invoice — straightforward delivery
    inv = B2BInvoice(
        id=1, invoice_number="B2B-00001", client_id=1, invoice_type="credit",
        status="partial", subtotal=Decimal("800"), total=Decimal("800"),
        amount_paid=Decimal("300"), created_at=datetime(2026, 8, 2, tzinfo=timezone.utc),
    )
    session.add(inv)
    session.flush()
    session.add_all([
        B2BInvoiceItem(invoice_id=1, product_id=1, qty=Decimal("30"), unit_price=Decimal("20"), total=Decimal("600")),
        B2BInvoiceItem(invoice_id=1, product_id=2, qty=Decimal("40"), unit_price=Decimal("5"), total=Decimal("200")),
    ])

    # 2) A consignment invoice — the SAME lines land on the invoice and the
    #    consignment, so only one of the two may be counted.
    cons_inv = B2BInvoice(
        id=2, invoice_number="B2B-00002", client_id=1, invoice_type="consignment",
        status="consignment", subtotal=Decimal("200"), total=Decimal("200"),
        amount_paid=Decimal("0"), created_at=datetime(2026, 8, 4, tzinfo=timezone.utc),
    )
    session.add(cons_inv)
    session.flush()
    session.add(B2BInvoiceItem(invoice_id=2, product_id=1, qty=Decimal("10"), unit_price=Decimal("20"), total=Decimal("200")))
    cons = Consignment(id=1, ref_number="CONS-0001", client_id=1, invoice_id=2,
                       status="active", created_at=datetime(2026, 8, 4, tzinfo=timezone.utc))
    session.add(cons)
    session.flush()
    session.add(ConsignmentItem(consignment_id=1, product_id=1, qty_sent=Decimal("10"),
                                qty_sold=Decimal("6"), qty_returned=Decimal("4"),
                                unit_price=Decimal("20")))

    # 3) A standalone consignment with no invoice — must be counted
    standalone = Consignment(id=2, ref_number="CONS-0002", client_id=1, invoice_id=None,
                             status="active", created_at=datetime(2026, 8, 6, tzinfo=timezone.utc))
    session.add(standalone)
    session.flush()
    session.add(ConsignmentItem(consignment_id=2, product_id=2, qty_sent=Decimal("20"),
                                qty_sold=Decimal("0"), qty_returned=Decimal("0"),
                                unit_price=Decimal("5")))

    # 4) A refund — returned goods
    refund = B2BRefund(id=1, refund_number="BRF-0001", client_id=1,
                       subtotal=Decimal("100"), total=Decimal("100"),
                       created_at=datetime(2026, 8, 8, tzinfo=timezone.utc))
    session.add(refund)
    session.flush()
    session.add(B2BRefundItem(refund_id=1, product_id=1, qty=Decimal("5"),
                              unit_price=Decimal("20"), total=Decimal("100")))

    # 5) Another client's invoice — must never leak into client 1's portal
    rival = B2BInvoice(id=3, invoice_number="B2B-00003", client_id=2, invoice_type="credit",
                       status="unpaid", subtotal=Decimal("999"), total=Decimal("999"),
                       amount_paid=Decimal("0"), created_at=datetime(2026, 8, 3, tzinfo=timezone.utc))
    session.add(rival)
    session.flush()
    session.add(B2BInvoiceItem(invoice_id=3, product_id=1, qty=Decimal("50"),
                               unit_price=Decimal("20"), total=Decimal("999")))

    session.commit()
    return client


def build_products(session, client_id=1):
    return run(b2b._build_client_products_payload(client_id, AsyncSessionAdapter(session)))


def build_stock(session, client_id=1):
    return run(b2b._build_client_consignment_stock_payload(client_id, AsyncSessionAdapter(session)))


def add_reported_sale(session, *, product_id, qty, unit_price, client_id=1, when=None):
    """A consignment payment recorded through Accounting: the client reported
    these goods sold, but the flow deliberately leaves qty_sold alone."""
    sale = ConsignmentSale(client_id=client_id, month_label="August 2026",
                           subtotal=Decimal(str(qty * unit_price)),
                           discount=Decimal("0"),
                           amount=Decimal(str(qty * unit_price)),
                           created_at=when or datetime(2026, 8, 31, tzinfo=timezone.utc))
    session.add(sale)
    session.flush()
    session.add(ConsignmentSaleItem(sale_id=sale.id, product_id=product_id,
                                    qty=Decimal(str(qty)), unit_price=Decimal(str(unit_price)),
                                    total=Decimal(str(qty * unit_price))))
    session.commit()


# ── Netting rules ────────────────────────────────────────────────────────────

def test_consignment_invoice_lines_are_not_counted_twice():
    with make_session() as session:
        seed(session)
        data = build_products(session)

    tomato = next(p for p in data["products"] if p["name"] == "Tomato")
    # 30 from the credit invoice + 10 from the consignment invoice — the
    # matching consignment lines must NOT add another 10.
    assert tomato["qty_received"] == 40.0
    assert tomato["value_received"] == 800.0


def test_standalone_consignment_is_counted():
    with make_session() as session:
        seed(session)
        data = build_products(session)

    herb = next(p for p in data["products"] if p["name"] == "Herb bunch")
    # 40 invoiced + 20 sent on a consignment that has no invoice behind it
    assert herb["qty_received"] == 60.0
    assert herb["value_received"] == 300.0
    assert herb["qty_returned"] == 0.0
    assert herb["qty_net"] == 60.0


def test_returns_come_from_refunds_and_settled_consignments():
    with make_session() as session:
        seed(session)
        data = build_products(session)

    tomato = next(p for p in data["products"] if p["name"] == "Tomato")
    # 5 returned on the refund + 4 returned through the consignment settle flow
    assert tomato["qty_returned"] == 9.0
    assert tomato["value_returned"] == 180.0     # 5×20 + 4×20
    assert tomato["qty_net"] == 31.0
    assert tomato["value_net"] == 620.0
    assert tomato["last_received"] == "04-Aug-2026"


def test_totals_and_delivery_log():
    with make_session() as session:
        seed(session)
        data = build_products(session)

    totals = data["totals"]
    assert totals["product_lines"] == 2
    assert totals["qty_net"] == 91.0             # 31 tomato + 60 herb
    assert totals["value_net"] == 920.0          # 620 + 300
    assert totals["value_received"] == 1100.0
    assert totals["value_returned"] == 180.0

    refs = [d["ref"] for d in data["deliveries"]]
    assert refs == ["BRF-0001", "CONS-0002", "B2B-00002", "B2B-00001"]   # newest first
    assert [d["kind"] for d in data["deliveries"]][0] == "return"
    # The consignment invoice appears once, as the invoice — not again as CONS-0001
    assert "CONS-0001" not in refs


def test_other_clients_data_never_appears():
    with make_session() as session:
        seed(session)
        data = build_products(session)

    assert all(d["ref"] != "B2B-00003" for d in data["deliveries"])
    tomato = next(p for p in data["products"] if p["name"] == "Tomato")
    assert tomato["qty_received"] == 40.0        # not 90


# ── Stock still held by the client ───────────────────────────────────────────

def test_stock_on_hand_is_what_was_sent_less_sold_and_returned():
    with make_session() as session:
        seed(session)
        data = build_stock(session)

    herb = next(r for r in data["items"] if r["name"] == "Herb bunch")
    # 20 sent on the standalone consignment, nothing sold or returned
    assert herb["qty_on_hand"] == 20.0
    assert herb["unit_price"] == 5.0
    assert herb["value_on_hand"] == 100.0

    tomato = next(r for r in data["items"] if r["name"] == "Tomato")
    # 10 sent, 6 settled as sold, 4 returned — nothing left with the client
    assert tomato["qty_on_hand"] == 0.0
    assert tomato["value_on_hand"] == 0.0

    # Goods bought outright are not consignment stock and never appear here
    assert data["totals"]["qty_on_hand"] == 20.0
    assert data["totals"]["value_on_hand"] == 100.0
    assert data["totals"]["product_lines"] == 1      # counts only what is in stock


def test_reported_sales_reduce_stock_even_though_settle_never_ran():
    # The Accounting payment flow is bookkeeping only — it does not touch
    # qty_sold — so the reported sold quantities have to be netted off here.
    with make_session() as session:
        seed(session)
        add_reported_sale(session, product_id=2, qty=8, unit_price=5)
        data = build_stock(session)

    herb = next(r for r in data["items"] if r["name"] == "Herb bunch")
    assert herb["qty_sold"] == 8.0
    assert herb["qty_on_hand"] == 12.0
    assert herb["value_on_hand"] == 60.0


def test_stock_never_goes_negative():
    with make_session() as session:
        seed(session)
        add_reported_sale(session, product_id=2, qty=50, unit_price=5)
        data = build_stock(session)

    herb = next(r for r in data["items"] if r["name"] == "Herb bunch")
    assert herb["qty_on_hand"] == 0.0


def test_another_clients_consignment_is_not_our_stock():
    with make_session() as session:
        seed(session)
        add_reported_sale(session, product_id=2, qty=5, unit_price=5, client_id=2)
        data = build_stock(session)

    # Client 2's reported sale must not eat into client 1's stock
    herb = next(r for r in data["items"] if r["name"] == "Herb bunch")
    assert herb["qty_on_hand"] == 20.0
    assert build_stock(session, client_id=2)["items"] == []


def seed_drifted_consignment(session):
    """A consignment invoice whose Consignment mirror lost its quantities.

    Both records are written at invoice time, and they demonstrably drift in
    the wild (edits, part-built records, imports). The invoice is the record
    of what physically went out, so stock must follow it.
    """
    client = B2BClient(id=1, name="Yo Studio", payment_terms="consignment",
                       is_active=True, portal_token=TOKEN, portal_enabled=True)
    beans = Product(id=1, sku="BNS", name="Coffee beans", price=Decimal("180"), unit="kg")
    session.add_all([client, beans])
    session.flush()

    inv = B2BInvoice(id=1, invoice_number="B2B-00001", client_id=1,
                     invoice_type="consignment", status="consignment",
                     subtotal=Decimal("900"), total=Decimal("900"),
                     amount_paid=Decimal("0"),
                     created_at=datetime(2026, 8, 4, tzinfo=timezone.utc))
    session.add(inv)
    session.flush()
    session.add(B2BInvoiceItem(invoice_id=1, product_id=1, qty=Decimal("5"),
                               unit_price=Decimal("180"), total=Decimal("900")))
    cons = Consignment(id=1, ref_number="CONS-0001", client_id=1, invoice_id=1,
                       status="active",
                       created_at=datetime(2026, 8, 4, tzinfo=timezone.utc))
    session.add(cons)
    session.flush()
    # The drift: the mirror carries the price but no quantity.
    session.add(ConsignmentItem(consignment_id=1, product_id=1, qty_sent=Decimal("0"),
                                qty_sold=Decimal("0"), qty_returned=Decimal("0"),
                                unit_price=Decimal("180")))
    session.commit()
    return client


def test_stock_follows_the_invoice_when_the_consignment_mirror_drifted():
    with make_session() as session:
        seed_drifted_consignment(session)
        stock = build_stock(session)
        received = build_products(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    # 5 on the invoice, nothing sold or returned - NOT the mirror's zero
    assert beans["qty_sent"] == 5.0
    assert beans["qty_on_hand"] == 5.0
    assert beans["unit_price"] == 180.0
    assert beans["value_on_hand"] == 900.0
    # ... and the two portal tabs agree, which is the whole point
    assert received["products"][0]["qty_net"] == beans["qty_on_hand"]


def test_reported_sale_against_a_drifted_mirror_still_reduces_stock():
    with make_session() as session:
        seed_drifted_consignment(session)
        add_reported_sale(session, product_id=1, qty=2, unit_price=180)
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_sold"] == 2.0
    assert beans["qty_on_hand"] == 3.0
    assert beans["value_on_hand"] == 540.0


def test_client_refunds_take_consignment_goods_back_off_the_shelf():
    with make_session() as session:
        seed_drifted_consignment(session)
        refund = B2BRefund(id=1, refund_number="BRF-0001", client_id=1,
                           subtotal=Decimal("360"), total=Decimal("360"),
                           created_at=datetime(2026, 8, 9, tzinfo=timezone.utc))
        session.add(refund)
        session.flush()
        session.add(B2BRefundItem(refund_id=1, product_id=1, qty=Decimal("2"),
                                  unit_price=Decimal("180"), total=Decimal("360")))
        session.commit()
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_returned"] == 2.0
    assert beans["qty_on_hand"] == 3.0


def pay(session, amount, *, invoice_id=1, subtotal=None, total=None):
    """Mark money against the consignment invoice, the way collecting does."""
    inv = session.get(B2BInvoice, invoice_id)
    if subtotal is not None:
        inv.subtotal = Decimal(str(subtotal))
    if total is not None:
        inv.total = Decimal(str(total))
    inv.amount_paid = Decimal(str(amount))
    inv.status = "paid" if float(inv.amount_paid) >= float(inv.total) else "partial"
    session.commit()


def test_a_fully_paid_consignment_leaves_nothing_on_the_shelf():
    # The client pays for what they sell. A consignment invoice paid in full
    # means the goods are gone, even when nobody recorded which ones - the
    # case for every payment taken before sold-item detail existed.
    with make_session() as session:
        seed_drifted_consignment(session)
        pay(session, 900)
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_sent"] == 5.0
    assert beans["qty_on_hand"] == 0.0
    assert stock["totals"]["value_on_hand"] == 0.0


def test_a_part_paid_consignment_leaves_the_unpaid_share():
    with make_session() as session:
        seed_drifted_consignment(session)
        pay(session, 360)          # 40% of 900
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_on_hand"] == 3.0        # 5 x (1 - 0.4)
    assert beans["value_on_hand"] == 540.0


def test_the_discount_is_taken_out_before_money_is_matched_to_goods():
    # Line prices are gross, payments are net. Valuing the shelf through the
    # invoice's own net ratio is what makes a fully paid discounted invoice
    # come out at zero instead of leaving a phantom 20% on the shelf.
    with make_session() as session:
        seed_drifted_consignment(session)
        pay(session, 720, subtotal=900, total=720)     # 20% client discount
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_on_hand"] == 0.0


def test_an_itemised_payment_is_not_counted_twice():
    # A payment recorded with its sold items lands in BOTH places: the
    # ConsignmentSale lines and the invoice's amount_paid. Retiring the named
    # goods and then spreading the same money again would sell them twice.
    with make_session() as session:
        seed_drifted_consignment(session)
        add_reported_sale(session, product_id=1, qty=2, unit_price=180)   # 360
        pay(session, 360)
        stock = build_stock(session)

    beans = next(r for r in stock["items"] if r["name"] == "Coffee beans")
    assert beans["qty_sold"] == 2.0
    assert beans["qty_on_hand"] == 3.0        # not 1.8


def test_shelf_value_reconciles_to_what_the_client_still_owes():
    # The invariant the whole model rests on: a consignment client is invoiced
    # for everything sent and pays it down as they sell, so with no unrecorded
    # sales the shelf is worth exactly the outstanding balance.
    with make_session() as session:
        seed_drifted_consignment(session)
        pay(session, 360, subtotal=900, total=720)
        stock = build_stock(session)
        inv = session.get(B2BInvoice, 1)
        owed = float(inv.total) - float(inv.amount_paid)

    assert abs(stock["totals"]["net_value_on_hand"] - owed) < 0.01


def seed_two_consignment_invoices(session):
    """The Yostudio shape: an older delivery settled in full, then a newer one
    part paid, with a single itemised payment spanning both."""
    client = B2BClient(id=1, name="Yostudio", payment_terms="consignment",
                       discount_pct=Decimal("20"), is_active=True,
                       portal_token=TOKEN, portal_enabled=True)
    balls = Product(id=1, sku="DB", name="Date Balls", price=Decimal("40"), unit="pcs")
    mint = Product(id=2, sku="WM", name="Wild Mint", price=Decimal("75"), unit="pcs")
    session.add_all([client, balls, mint])
    session.flush()

    def delivery(inv_id, cons_id, number, ref, when, lines, paid):
        gross = sum(q * p for _pid, q, p in lines)
        net = gross * 0.8
        session.add(B2BInvoice(
            id=inv_id, invoice_number=number, client_id=1,
            invoice_type="consignment",
            status="paid" if paid >= net - 0.005 else "partial",
            subtotal=Decimal(str(gross)), discount=Decimal(str(gross - net)),
            total=Decimal(str(net)), amount_paid=Decimal(str(paid)), created_at=when))
        session.flush()
        session.add(Consignment(id=cons_id, ref_number=ref, client_id=1,
                                invoice_id=inv_id, status="active", created_at=when))
        session.flush()
        for pid, qty, price in lines:
            session.add(B2BInvoiceItem(invoice_id=inv_id, product_id=pid,
                                       qty=Decimal(str(qty)), unit_price=Decimal(str(price)),
                                       total=Decimal(str(qty * price))))
            # The mirror drifts to zero, as it does in production.
            session.add(ConsignmentItem(consignment_id=cons_id, product_id=pid,
                                        qty_sent=Decimal("0"), qty_sold=Decimal("0"),
                                        qty_returned=Decimal("0"),
                                        unit_price=Decimal(str(price))))

    # 21-Aug: 10 Date Balls, paid in full (400 gross -> 320 net)
    delivery(1, 1, "B2B-00324", "CONS-0048",
             datetime(2026, 8, 21, tzinfo=timezone.utc), [(1, 10, 40.0)], 320.0)
    # 31-Aug: 10 more Date Balls + 3 Wild Mint (625 gross -> 500 net), 180 paid
    delivery(2, 2, "B2B-00326", "CONS-0050",
             datetime(2026, 8, 31, tzinfo=timezone.utc),
             [(1, 10, 40.0), (2, 3, 75.0)], 180.0)
    session.commit()
    return client


def test_payment_settles_the_invoice_it_was_for_not_the_newest_goods():
    # One payment of 500 net: 320 cleared B2B-00324 (its 10 Date Balls) and
    # 180 went to B2B-00326 (its 3 Wild Mint, 225 gross x 0.8). The 10 Date
    # Balls delivered on the newer invoice are still on the shelf - they must
    # not be retired by money that belonged to the older one.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        add_reported_sale(session, product_id=1, qty=10, unit_price=40)   # 400 -> inv 324
        add_reported_sale(session, product_id=2, qty=3, unit_price=75)    # 225 -> inv 326
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    mint = next(r for r in stock["items"] if r["name"] == "Wild Mint")
    assert balls["qty_sent"] == 20.0
    assert balls["qty_sold"] == 10.0
    assert balls["qty_on_hand"] == 10.0
    assert mint["qty_on_hand"] == 0.0
    # What is left is exactly the unpaid part of the newer invoice
    assert stock["totals"]["net_value_on_hand"] == 320.0     # 400 gross x 0.8


def test_a_settled_older_invoice_does_not_swallow_the_named_items():
    # Named items are the newest money. Matching them against the oldest
    # fully-paid invoice first would consume them there and leave the goods
    # they were actually reported against sitting on the shelf.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        add_reported_sale(session, product_id=2, qty=3, unit_price=75)
        stock = build_stock(session)

    mint = next(r for r in stock["items"] if r["name"] == "Wild Mint")
    assert mint["qty_sold"] == 3.0
    assert mint["qty_on_hand"] == 0.0


def test_a_return_is_not_also_counted_as_sold():
    # The refunded goods came off the older invoice, which was paid in full.
    # Counting them as sold there AND as returned would retire them twice and
    # eat into the newer delivery.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        refund = B2BRefund(id=1, refund_number="BRF-0001", client_id=1,
                           subtotal=Decimal("80"), total=Decimal("80"),
                           created_at=datetime(2026, 8, 25, tzinfo=timezone.utc))
        session.add(refund)
        session.flush()
        session.add(B2BRefundItem(refund_id=1, product_id=1, qty=Decimal("2"),
                                  unit_price=Decimal("40"), total=Decimal("80")))
        session.commit()
        add_reported_sale(session, product_id=1, qty=10, unit_price=40)
        add_reported_sale(session, product_id=2, qty=3, unit_price=75)
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    # 20 sent, 10 reported sold, 2 handed back -> 8 left, not 6: the two that
    # came back must not also be retired as sold by the invoice that paid.
    assert balls["qty_sent"] == 20.0
    assert balls["qty_returned"] == 2.0
    assert balls["qty_sold"] == 10.0
    assert balls["qty_on_hand"] == 8.0


def count_stock(session, when, lines, client_id=1):
    """Record a physical count: {product_id: qty}."""
    c = ConsignmentStockCount(client_id=client_id, counted_at=when)
    session.add(c)
    session.flush()
    for pid, qty in lines.items():
        session.add(ConsignmentStockCountItem(count_id=c.id, product_id=pid,
                                              qty=Decimal(str(qty))))
    session.commit()
    return c


def test_a_count_overrides_what_the_invoices_imply():
    # The derived figure cannot know about a sale nobody recorded, nor about
    # goods that were paid for but never left the shelf. A count can say both.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        # Derivation would leave 10 Date Balls; the shelf actually has 4, and
        # 3 Wild Mint the payment implied were gone are still there.
        count_stock(session, datetime(2026, 9, 1, tzinfo=timezone.utc),
                    {1: 4, 2: 3})
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    mint = next(r for r in stock["items"] if r["name"] == "Wild Mint")
    assert balls["qty_on_hand"] == 4.0
    assert mint["qty_on_hand"] == 3.0
    assert stock["counted_at"] == "01-Sep-2026"


def test_deliveries_after_a_count_are_added_to_it():
    with make_session() as session:
        seed_two_consignment_invoices(session)
        count_stock(session, datetime(2026, 8, 25, tzinfo=timezone.utc), {1: 4})
        # B2B-00326 (31-Aug) lands after the count: 10 more Date Balls, and
        # its 3 Wild Mint. Only 180 of its 500 net is paid.
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    assert balls["qty_counted"] == 4.0
    assert balls["qty_sent"] == 10.0          # only the post-count delivery
    assert balls["qty_on_hand"] > 4.0         # the count plus what is unpaid


def test_a_sale_reported_after_a_count_comes_off_it():
    with make_session() as session:
        seed_two_consignment_invoices(session)
        count_stock(session, datetime(2026, 9, 1, tzinfo=timezone.utc), {1: 4})
        add_reported_sale(session, product_id=1, qty=3, unit_price=40,
                          when=datetime(2026, 9, 5, tzinfo=timezone.utc))
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    assert balls["qty_on_hand"] == 1.0


def test_history_before_a_count_is_superseded_not_re_subtracted():
    # Everything the older invoices and their payments imply is represented by
    # the count itself, so none of it may be applied on top of it again.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        add_reported_sale(session, product_id=1, qty=10, unit_price=40)
        count_stock(session, datetime(2026, 9, 1, tzinfo=timezone.utc), {1: 6})
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    assert balls["qty_on_hand"] == 6.0        # not 6 - 10 clamped to 0


def test_the_latest_count_is_the_one_that_counts():
    with make_session() as session:
        seed_two_consignment_invoices(session)
        count_stock(session, datetime(2026, 9, 1, tzinfo=timezone.utc), {1: 4})
        count_stock(session, datetime(2026, 9, 20, tzinfo=timezone.utc), {1: 7})
        stock = build_stock(session)

    balls = next(r for r in stock["items"] if r["name"] == "Date Balls")
    assert balls["qty_on_hand"] == 7.0
    assert stock["counted_at"] == "20-Sep-2026"


def test_a_product_left_off_the_count_is_counted_as_zero():
    # A count is a statement about the whole shelf, so anything not on it is
    # not there - otherwise stale stock would linger with no way to clear it.
    with make_session() as session:
        seed_two_consignment_invoices(session)
        count_stock(session, datetime(2026, 9, 1, tzinfo=timezone.utc), {1: 4})
        stock = build_stock(session)

    mint = next((r for r in stock["items"] if r["name"] == "Wild Mint"), None)
    assert mint is None or mint["qty_on_hand"] == 0.0
    assert stock["totals"]["qty_on_hand"] == 4.0


def test_a_counted_shelf_is_still_valued_at_the_clients_billed_prices():
    # With every invoice superseded by the count there is no invoice left to
    # read the discount rate from, so it has to come from the client instead -
    # otherwise a counted shelf would be valued gross while every other one is
    # valued net.
    with make_session() as session:
        seed_two_consignment_invoices(session)      # client discount is 20%
        count_stock(session, datetime(2026, 9, 30, tzinfo=timezone.utc), {1: 5})
        stock = build_stock(session)

    assert stock["totals"]["value_on_hand"] == 200.0            # 5 x 40 gross
    assert stock["totals"]["net_value_on_hand"] == 160.0        # x 0.8


def test_a_client_with_no_consignments_has_no_stock():
    with make_session() as session:
        seed(session)
        assert build_stock(session, client_id=2)["items"] == []
        assert build_stock(session, client_id=2)["totals"]["value_on_hand"] == 0.0


# ── Portal access control ────────────────────────────────────────────────────

def make_client(session):
    async def override_session() -> AsyncGenerator[AsyncSessionAdapter, None]:
        yield AsyncSessionAdapter(session)

    async def noop() -> None:
        return None

    app_factory.configure_logging = lambda: None
    app_factory.configure_monitoring = lambda: None
    app_factory.verify_migration_status = noop
    # Lifespan opens a Redis pool and every startup schema guard; none of that
    # is under test here and the Redis dial-out costs seconds per test.
    for guard in (
        "ensure_payroll_columns", "ensure_price_precision",
        "ensure_delivery_transport_columns", "ensure_product_categories_table",
        "ensure_consignment_sales_tables", "ensure_b2b_portal_columns",
        "ensure_carbon_methodology", "sync_livestock_emissions_on_boot",
        "seed_chart_of_accounts",
    ):
        setattr(app_factory, guard, noop)
    import app.core.cache as cache
    cache.init_redis_pool = noop
    cache.close_redis_pool = noop

    app = create_app()
    app.dependency_overrides[get_async_session] = override_session
    return TestClient(app)


TOKEN = "tok_" + "a" * 40


def test_portal_opens_with_no_session_at_all():
    with make_session() as session:
        seed(session)
        with make_client(session) as client:
            res = client.get(f"/portal/c/{TOKEN}")

    assert res.status_code == 200
    assert "Green Cafe" in res.text
    # The URL alone unlocks the data — keep it out of indexes and shared caches
    assert "noindex" in res.headers["X-Robots-Tag"]
    assert "no-store" in res.headers["Cache-Control"]


def test_portal_data_returns_statement_and_products():
    with make_session() as session:
        seed(session)
        with make_client(session) as client:
            res = client.get(f"/portal/c/{TOKEN}/data")

    assert res.status_code == 200
    body = res.json()
    assert body["client"]["name"] == "Green Cafe"
    assert body["client"]["code"] == "C0001"
    assert body["total_invoiced"] == 1000.0          # 800 + 200
    assert body["balance_due"] == 600.0              # 1000 charged − 300 paid − 100 refund
    assert body["product_totals"]["value_net"] == 920.0
    assert {p["name"] for p in body["products"]} == {"Tomato", "Herb bunch"}
    # Stock on hand shows only what the client still holds — the fully
    # settled tomato consignment is dropped, the untouched herbs remain
    assert [r["name"] for r in body["stock"]] == ["Herb bunch"]
    assert body["stock"][0]["qty_on_hand"] == 20.0
    assert body["stock_totals"]["value_on_hand"] == 100.0
    # Nothing that could identify another client
    assert "Rival Bakery" not in res.text


def test_unknown_revoked_and_disabled_tokens_all_404():
    for label, kwargs, token in [
        ("unknown", {}, "tok_" + "z" * 40),
        ("revoked", {"portal_token": None}, TOKEN),
        ("disabled", {"portal_enabled": False}, TOKEN),
        ("inactive client", {"client_active": False}, TOKEN),
        ("too short", {}, "abc"),
    ]:
        with make_session() as session:
            seed(session, **kwargs)
            with make_client(session) as client:
                res = client.get(f"/portal/c/{token}")
                data_res = client.get(f"/portal/c/{token}/data")
        assert res.status_code == 404, label
        assert data_res.status_code == 404, label
        assert "Green Cafe" not in res.text, label


def seed_payment(session, *, amount="300", invoice_id=1, description="Bank transfer for B2B-00001"):
    """A recorded client payment — journal + cash-account entry, the shape
    _load_client_payment_activity looks for."""
    session.add(Account(id=1, code="1000", name="Cash", type="asset"))
    session.flush()
    session.add(Journal(id=1, ref_type="b2b_payment", ref_id=invoice_id, description=description,
                        created_at=datetime(2026, 8, 5, tzinfo=timezone.utc)))
    session.flush()
    session.add(JournalEntry(journal_id=1, account_id=1, debit=Decimal(amount), credit=Decimal("0")))
    session.commit()


def test_portal_data_serialises_when_the_client_has_payments():
    """Regression: payment records carry a raw datetime under "date", and
    building the JSONResponse ourselves skips FastAPI's encoder — so this
    endpoint used to 500 for every client who had ever paid anything."""
    with make_session() as session:
        seed(session)
        seed_payment(session)
        with make_client(session) as api:
            res = api.get(f"/portal/c/{TOKEN}/data")

    assert res.status_code == 200
    payments = res.json()["payment_activity"]
    assert len(payments) == 1
    # The four fields the page actually renders, under the names it reads
    assert payments[0] == {
        "date": "05-Aug-2026",
        "ref": "B2B-00001",
        "desc": "Bank transfer for B2B-00001",
        "amount": 300.0,
    }


def test_portal_never_exposes_our_staff_names():
    """Payment records carry the employee who booked the payment. That is
    internal — it must not ride along in a client-facing payload."""
    with make_session() as session:
        seed(session)
        session.add(User(id=7, name="Sara Bookkeeper", email="sara@farm.example",
                         password="x", role="accountant"))
        session.flush()
        seed_payment(session)
        session.execute(Journal.__table__.update().values(user_id=7))
        session.commit()
        with make_client(session) as api:
            res = api.get(f"/portal/c/{TOKEN}/data")

    assert res.status_code == 200
    assert "Sara Bookkeeper" not in res.text
    assert "user_name" not in res.text


def fake_request(host="farm.example.com"):
    from starlette.requests import Request
    return Request({
        "type": "http", "method": "GET", "path": "/", "headers": [(b"host", host.encode())],
        "query_string": b"", "scheme": "https", "server": (host, 443), "root_path": "",
    })


def test_issue_rotate_and_revoke_lifecycle():
    with make_session() as session:
        client_row = seed(session, portal_token=None, portal_enabled=False)
        db = AsyncSessionAdapter(session)
        user = User(id=1, name="Admin", email="a@x.example", password="x", role="admin")
        session.add(user)
        session.commit()
        req = fake_request()

        # Nothing issued yet
        assert run(b2b.get_client_portal_link(1, req, db))["enabled"] is False

        issued = run(b2b.create_client_portal_link(1, req, rotate=False, db=db, current_user=user))
        assert issued["enabled"] is True
        assert issued["rotated"] is True
        first_url = issued["url"]
        assert first_url.startswith("https://farm.example.com/portal/c/")
        first_token = first_url.rsplit("/", 1)[-1]
        assert len(first_token) >= 32          # secrets.token_urlsafe(32)

        # Re-opening the dialog must hand back the SAME link the client bookmarked
        again = run(b2b.create_client_portal_link(1, req, rotate=False, db=db, current_user=user))
        assert again["url"] == first_url
        assert again["rotated"] is False

        # Rotating mints a new token and kills the old one
        rotated = run(b2b.create_client_portal_link(1, req, rotate=True, db=db, current_user=user))
        assert rotated["rotated"] is True
        assert rotated["url"] != first_url

        # Revoking clears the token as well as the flag
        run(b2b.revoke_client_portal_link(1, db=db, current_user=user))
        session.refresh(client_row)
        assert client_row.portal_enabled is False
        assert client_row.portal_token is None
        assert run(b2b.get_client_portal_link(1, req, db))["url"] is None


def test_rotated_link_kills_the_previous_url():
    with make_session() as session:
        seed(session)
        db = AsyncSessionAdapter(session)
        user = User(id=1, name="Admin", email="a@x.example", password="x", role="admin")
        session.add(user)
        session.commit()
        run(b2b.create_client_portal_link(1, fake_request(), rotate=True, db=db, current_user=user))

        with make_client(session) as api:
            dead = api.get(f"/portal/c/{TOKEN}")

    assert dead.status_code == 404


def test_html_view_counts_opens_but_polling_does_not():
    with make_session() as session:
        client_row = seed(session)
        with make_client(session) as api:
            api.get(f"/portal/c/{TOKEN}")
            api.get(f"/portal/c/{TOKEN}")
            api.get(f"/portal/c/{TOKEN}/data")
            api.get(f"/portal/c/{TOKEN}/data")
        session.refresh(client_row)
        views = client_row.portal_view_count
        last = client_row.portal_last_viewed_at

    assert views == 2          # the two polls must not inflate this
    assert last is not None
