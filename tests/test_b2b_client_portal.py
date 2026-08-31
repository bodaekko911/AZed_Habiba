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


def add_reported_sale(session, *, product_id, qty, unit_price, client_id=1):
    """A consignment payment recorded through Accounting: the client reported
    these goods sold, but the flow deliberately leaves qty_sold alone."""
    sale = ConsignmentSale(client_id=client_id, month_label="August 2026",
                           subtotal=Decimal(str(qty * unit_price)),
                           discount=Decimal("0"),
                           amount=Decimal(str(qty * unit_price)))
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
