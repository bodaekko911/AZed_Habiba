from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import HTMLResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload
from sqlalchemy import func, select
from typing import Optional, List
from pydantic import BaseModel
from decimal import Decimal
from datetime import date, datetime, time, timedelta, timezone
import re

from app.database import get_async_session
from app.core.permissions import get_current_user, require_action, require_admin, require_permission
from app.core.log import record
from app.core.navigation import render_app_header
from app.core.product_types import is_stock_tracked_product, normalize_item_type
from app.core.templates import templates
from app.models.b2b import (
    B2BClient, B2BInvoice, B2BInvoiceItem, Consignment, ConsignmentItem,
    ConsignmentSale, ConsignmentSaleItem, B2BRefund, B2BRefundItem, B2BClientPrice,
)
from app.models.product import Product
from app.models.inventory import StockMove
from app.models.accounting import Journal, JournalEntry
from app.models.user import User

router = APIRouter(
    prefix="/b2b",
    tags=["B2B"],
    dependencies=[Depends(require_permission("page_b2b"))],
)


# ── Schemas ────────────────────────────────────────────
class ClientCreate(BaseModel):
    name:           str
    contact_person: Optional[str] = None
    phone:          Optional[str] = None
    email:          Optional[str] = None
    address:        Optional[str] = None
    payment_terms:  str = "cash"
    discount_pct:   float = 0
    credit_limit:   float = 0
    notes:          Optional[str] = None

class ClientUpdate(BaseModel):
    name:           Optional[str] = None
    contact_person: Optional[str] = None
    phone:          Optional[str] = None
    email:          Optional[str] = None
    address:        Optional[str] = None
    payment_terms:  Optional[str] = None
    discount_pct:   Optional[float] = None
    credit_limit:   Optional[float] = None
    notes:          Optional[str] = None

class InvoiceItemIn(BaseModel):
    product_id: int
    qty:        float
    unit_price: float

class InvoiceCreate(BaseModel):
    client_id:      int
    invoice_type:   Optional[str] = None
    payment_method: Optional[str] = None
    discount_pct:   float = 0
    notes:          Optional[str] = None
    items:          List[InvoiceItemIn]

class PaymentRecord(BaseModel):
    amount: float
    method: str = "transfer"

class PaymentReversal(BaseModel):
    # Omit amount to reverse the whole payment; supply one to reverse part of it.
    amount: Optional[float] = None
    reason: Optional[str] = None

class ConsignmentSettle(BaseModel):
    items: List[dict]

class RefundItemIn(BaseModel):
    product_id: int
    qty:        float
    unit_price: float

class ClientRefundCreate(BaseModel):
    client_id: int
    notes:     Optional[str] = None
    items:     List[RefundItemIn]


# ── HELPERS ────────────────────────────────────────────
from app.services.b2b_shared import (
    post_journal      as _post_journal,
    seed_deferred_revenue as _seed_deferred_revenue,
    next_b2b_number   as _next_b2b_number,
    next_cons_number  as _next_cons_number,
    get_b2b_client_top_products as _get_b2b_client_top_products,
)

async def _next_refund_number(db: AsyncSession) -> str:
    _r = await db.execute(select(func.max(B2BRefund.id)))
    max_id = _r.scalar() or 0
    return f"RFD-{str(max_id + 1).zfill(5)}"

def _normalized_client_terms(client: B2BClient) -> str:
    terms = (client.payment_terms or "cash").strip().lower()
    if terms in ("cash", "full_payment", "consignment"):
        return terms
    if terms in ("immediate", "pay_now", "cod"):
        return "cash"
    if terms in ("credit", "net15", "net30", "net60"):
        return "full_payment"
    return "cash"

def _client_discount_pct(client: B2BClient) -> float:
    return float(client.discount_pct or 0)


def _b2b_num(value) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


async def _validate_b2b_invoice_stock(
    db: AsyncSession,
    items: List[InvoiceItemIn],
) -> dict[int, Product]:
    products_by_id: dict[int, Product] = {}
    requested_by_product: dict[int, float] = {}

    for item in items:
        product = products_by_id.get(item.product_id)
        if product is None:
            _r = await db.execute(select(Product).where(Product.id == item.product_id))
            product = _r.scalar_one_or_none()
            if not product:
                raise HTTPException(status_code=404, detail=f"Product not found: {item.product_id}")
            products_by_id[item.product_id] = product

        requested_by_product[item.product_id] = requested_by_product.get(item.product_id, 0.0) + float(item.qty)

    for product_id, requested_qty in requested_by_product.items():
        product = products_by_id[product_id]
        if is_stock_tracked_product(product) and float(product.stock) < requested_qty:
            raise HTTPException(
                status_code=400,
                detail=f"Not enough stock for '{product.name}'. Available: {float(product.stock)}",
            )

    return products_by_id


def _date_key(value) -> datetime:
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone(timezone.utc).replace(tzinfo=None)
        return value
    return datetime.min


async def _load_client_payment_activity(
    db: AsyncSession,
    *,
    client_id: int,
    as_of: Optional[date] = None,
):
    payment_ref_types = ("consignment_client_payment", "consignment_payment", "b2b_payment", "b2b_collection")
    stmt = (
        select(Journal)
        .where(Journal.ref_type.in_(payment_ref_types))
        .options(selectinload(Journal.entries).selectinload(JournalEntry.account), selectinload(Journal.user))
        .order_by(Journal.created_at)
    )
    if as_of:
        stmt = stmt.where(Journal.created_at < datetime.combine(as_of + timedelta(days=1), time.min, tzinfo=timezone.utc))
    payment_result = await db.execute(stmt)
    journals = payment_result.scalars().all()

    invoice_result = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.client_id == client_id)
        .options(selectinload(B2BInvoice.client))
    )
    invoices = invoice_result.scalars().all()
    invoice_by_id = {invoice.id: invoice for invoice in invoices}
    invoice_by_number = {str(invoice.invoice_number or "").upper(): invoice for invoice in invoices}
    invoice_pattern = re.compile(r"\b([A-Z]*B2B-\d{5,})\b", re.IGNORECASE)

    records = []
    for journal in journals:
        matched_invoice = None
        if journal.ref_type == "consignment_client_payment":
            if journal.ref_id != client_id:
                continue
        else:
            if journal.ref_id and journal.ref_id in invoice_by_id:
                matched_invoice = invoice_by_id[journal.ref_id]
            else:
                match = invoice_pattern.search(journal.description or "")
                if match:
                    matched_invoice = invoice_by_number.get(match.group(1).upper())
            if not matched_invoice or matched_invoice.client_id != client_id:
                continue

        amount = 0.0
        for entry in journal.entries:
            if entry.account and entry.account.code == "1000" and float(entry.debit or 0) > 0:
                amount = float(entry.debit or 0)
                break
        if amount <= 0:
            amount = max((float(entry.debit or 0) for entry in journal.entries), default=0.0)
        if amount <= 0:
            continue

        reference = f"PAY-{journal.id}"
        if matched_invoice and matched_invoice.invoice_number:
            reference = matched_invoice.invoice_number

        records.append({
            "date": journal.created_at,
            "date_str": journal.created_at.strftime("%d-%b-%Y") if journal.created_at else "—",
            "ref": reference,
            "type": "payment",
            "desc": journal.description or "Client payment",
            "amount": round(amount, 2),
            "ref_type": journal.ref_type or "payment",
            "user_name": journal.user.name if journal.user else "—",
        })
    return records

async def _reverse_invoice_stock(invoice, db: AsyncSession):
    for item in invoice.items:
        _r = await db.execute(select(Product).where(Product.id == item.product_id))
        product = _r.scalar_one_or_none()
        if product and is_stock_tracked_product(product):
            before = float(product.stock)
            after  = before + float(item.qty)
            product.stock = after
            db.add(StockMove(
                product_id=product.id, type="in",
                qty=float(item.qty), qty_before=before, qty_after=after,
                ref_type="b2b_reversal", ref_id=invoice.id,
                note=f"Edit/Delete reversal — {invoice.invoice_number}",
            ))

async def _reverse_invoice_journal(invoice, db: AsyncSession):
    total = float(invoice.total)
    if invoice.invoice_type == "cash":
        # Cash invoices now use AR at creation (Dr AR / Cr Revenue)
        # Reverse: Dr Revenue / Cr AR
        await _post_journal(db, f"Reversal — {invoice.invoice_number}", "b2b_reversal", [
            ("1100", 0, total),
            ("4000", total, 0),
        ])
        client = invoice.client
        # Reverse outstanding: subtract unpaid portion
        unpaid = max(0.0, total - float(invoice.amount_paid))
        if unpaid > 0:
            client.outstanding = Decimal(str(max(0, float(client.outstanding) - unpaid)))
    elif invoice.invoice_type in ("full_payment", "consignment"):
        # Reverse: debit Deferred Revenue, credit AR
        await _post_journal(db, f"Reversal — {invoice.invoice_number}", "b2b_reversal", [
            ("2200", total, 0),   # Debit Deferred Revenue (reversal)
            ("1100", 0, total),   # Credit AR
        ])
        client = invoice.client
        # Only subtract the UNPAID portion — the paid portion was already removed
        # from client.outstanding when payment was collected, so subtracting total
        # again would double-reverse it.
        unpaid = max(0.0, total - float(invoice.amount_paid))
        if unpaid > 0:
            client.outstanding = Decimal(str(max(0, float(client.outstanding) - unpaid)))


async def _reverse_refund_effects(refund, db: AsyncSession, current_user: User):
    for item in refund.items:
        product = item.product
        if not product:
            continue
        if not is_stock_tracked_product(product):
            continue
        before = float(product.stock)
        after = before - float(item.qty)
        if after < -0.001:
            raise HTTPException(
                status_code=400,
                detail=f"Cannot delete refund {refund.refund_number}: stock for {product.name} would become negative",
            )
        after = max(0.0, after)
        product.stock = after
        db.add(StockMove(
            product_id=product.id,
            type="out",
            qty=float(item.qty),
            user_id=current_user.id,
            qty_before=before,
            qty_after=after,
            ref_type="b2b_refund_delete",
            ref_id=refund.id,
            note=f"Delete refund â€” {refund.refund_number}",
        ))

    if refund.client:
        refund.client.outstanding = Decimal(str(float(refund.client.outstanding) + float(refund.total)))

    await _post_journal(
        db,
        f"Delete refund â€” {refund.refund_number}",
        "b2b_refund_delete",
        [
            ("2200", 0, float(refund.total)),
            ("1100", float(refund.total), 0),
        ],
        user_id=current_user.id,
    )


# ── SEED DEFERRED REVENUE ──────────────────────────────
@router.post("/api/seed-accounts")
async def seed_accounts(db: AsyncSession = Depends(get_async_session), _: User = Depends(require_admin)):
    await _seed_deferred_revenue(db)
    return {"ok": True}


# ── CLIENT API ─────────────────────────────────────────
from app.services.b2b_shared import (
    client_invoice_balance_subquery as _client_invoice_balance_subquery,
    client_outstanding_value as _client_outstanding_value,
    client_refund_subquery as _client_refund_subquery,
)


@router.get("/api/clients")
async def get_clients(q: str = "", db: AsyncSession = Depends(get_async_session)):
    # Compute outstanding live from invoice data so it always matches the
    # invoices tab — less refunds, which are credits against the account. The
    # refund leg used to be missing here, so issuing a refund left the balance
    # on this screen unchanged even though the statement and the client
    # analysis both already netted it off.
    outstanding_sub = _client_invoice_balance_subquery()
    refund_sub = _client_refund_subquery()
    # Clamped to zero in Python, not SQL: a two-argument MAX is an aggregate in
    # Postgres (GREATEST is the scalar there) but a scalar in SQLite, so doing
    # it in the query would only work on one of them.
    computed_outstanding_expr = (
        func.coalesce(outstanding_sub.c.outstanding, 0)
        - func.coalesce(refund_sub.c.refunded, 0)
    )
    stmt = (
        select(B2BClient, computed_outstanding_expr.label("computed_outstanding"))
        .outerjoin(outstanding_sub, outstanding_sub.c.client_id == B2BClient.id)
        .outerjoin(refund_sub, refund_sub.c.client_id == B2BClient.id)
        .where(B2BClient.is_active == True)
        .options(selectinload(B2BClient.invoices))
        .order_by(B2BClient.name)
    )
    if q:
        stmt = stmt.where(
            B2BClient.name.ilike(f"%{q}%") |
            B2BClient.phone.ilike(f"%{q}%")
        )
    _r = await db.execute(stmt)
    rows = _r.all()
    return [
        {
            "id":             c.id,
            "name":           c.name,
            "contact_person": c.contact_person or "—",
            "phone":          c.phone or "—",
            "email":          c.email or "—",
            "address":        c.address or "—",
            "payment_terms":  c.payment_terms,
            "discount_pct":   float(c.discount_pct or 0),
            "credit_limit":   float(c.credit_limit or 0),
            "outstanding":    max(float(computed_outstanding or 0), 0.0),
            "notes":          c.notes or "",
            "invoice_count":  len(c.invoices),
            "portal_enabled": bool(c.portal_enabled and c.portal_token),
            "portal_views":   int(c.portal_view_count or 0),
        }
        for c, computed_outstanding in rows
    ]

@router.post("/api/clients", dependencies=[Depends(require_action("b2b", "clients", "create_client"))])
async def create_client(data: ClientCreate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    c = B2BClient(
        name=data.name, contact_person=data.contact_person,
        phone=data.phone, email=data.email, address=data.address,
        payment_terms=data.payment_terms,
        discount_pct=data.discount_pct,
        credit_limit=data.credit_limit,
        notes=data.notes,
    )
    db.add(c); await db.commit(); await db.refresh(c)
    return {"id": c.id, "name": c.name}

@router.put("/api/clients/{client_id}", dependencies=[Depends(require_action("b2b", "clients", "update_client"))])
async def update_client(client_id: int, data: ClientUpdate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    c = _r.scalar_one_or_none()
    if not c:
        raise HTTPException(status_code=404, detail="Client not found")
    if data.name is not None:           c.name           = data.name
    if data.contact_person is not None: c.contact_person = data.contact_person
    if data.phone is not None:          c.phone          = data.phone
    if data.email is not None:          c.email          = data.email
    if data.address is not None:        c.address        = data.address
    if data.payment_terms is not None:  c.payment_terms  = data.payment_terms
    if data.discount_pct is not None:   c.discount_pct   = data.discount_pct
    if data.credit_limit is not None:   c.credit_limit   = data.credit_limit
    if data.notes is not None:          c.notes          = data.notes
    await db.commit()
    return {"ok": True}

@router.delete("/api/clients/{client_id}", dependencies=[Depends(require_action("b2b", "clients", "delete_client"))])
async def delete_client(client_id: int, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    c = _r.scalar_one_or_none()
    if not c:
        raise HTTPException(status_code=404, detail="Client not found")
    c.is_active = False
    await db.commit()
    return {"ok": True}


# ── INVOICE API ────────────────────────────────────────
@router.get("/api/invoices")
async def get_invoices(client_id: int = None, skip: int = 0, limit: int = 100, db: AsyncSession = Depends(get_async_session)):
    where = []
    if client_id:
        where.append(B2BInvoice.client_id == client_id)
    cnt_r = await db.execute(select(func.count()).select_from(B2BInvoice).where(*where))
    total = cnt_r.scalar()
    inv_r = await db.execute(
        select(B2BInvoice)
        .where(*where)
        .options(
            selectinload(B2BInvoice.client),
            selectinload(B2BInvoice.items).selectinload(B2BInvoiceItem.product),
        )
        .order_by(B2BInvoice.created_at.desc()).offset(skip).limit(limit)
    )
    invoices = inv_r.scalars().all()
    return {
        "total": total,
        "invoices": [
            {
                "id":             i.id,
                "invoice_number": i.invoice_number,
                "client":         i.client.name if i.client else "—",
                "client_id":      i.client_id,
                "invoice_type":   i.invoice_type,
                "status":         i.status,
                "payment_method": i.payment_method or "—",
                "subtotal":       float(i.subtotal),
                "discount":       float(i.discount),
                "total":          float(i.total),
                "amount_paid":    float(i.amount_paid),
                "balance_due":    float(i.total) - float(i.amount_paid),
                "discount_pct":   round(float(i.discount) / float(i.subtotal) * 100, 1) if float(i.subtotal) > 0 else 0,
                "notes":          i.notes or "",
                "created_at":     i.created_at.strftime("%Y-%m-%d %H:%M") if i.created_at else "—",
                "items": [
                    {
                        "product":    item.product.name if item.product else "—",
                        "product_id": item.product_id,
                        "qty":        float(item.qty),
                        "unit_price": float(item.unit_price),
                        "total":      float(item.total),
                    }
                    for item in i.items
                ],
            }
            for i in invoices
        ],
    }

@router.post("/api/invoices", dependencies=[Depends(require_action("b2b", "invoices", "create"))])
async def create_invoice(data: InvoiceCreate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    await _seed_deferred_revenue(db)

    _r = await db.execute(select(B2BClient).where(B2BClient.id == data.client_id))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    if not data.items:
        raise HTTPException(status_code=400, detail="Invoice must have at least one item")

    products_by_id = await _validate_b2b_invoice_stock(db, data.items)

    invoice_type    = _normalized_client_terms(client)
    discount_pct    = _client_discount_pct(client)
    subtotal        = sum(i.qty * i.unit_price for i in data.items)
    discount_amount = round(subtotal * (discount_pct / 100), 2)
    total           = round(subtotal - discount_amount, 2)
    invoice_number  = await _next_b2b_number(db)

    # 100% discount (or any invoice that nets to zero) has nothing to collect,
    # so it is booked as fully paid on creation with no AR / outstanding impact.
    fully_discounted = discount_pct >= 100 or total <= 0.005
    status      = "paid" if fully_discounted else "unpaid"
    amount_paid = total if fully_discounted else 0

    invoice = B2BInvoice(
        invoice_number=invoice_number, client_id=data.client_id,
        user_id=current_user.id,
        invoice_type=invoice_type,
        status=status,
        payment_method=invoice_type,
        subtotal=round(subtotal, 2), discount=discount_amount,
        total=total,
        amount_paid=amount_paid,
        notes=data.notes,
    )
    db.add(invoice); await db.flush()

    for item in data.items:
        product = products_by_id[item.product_id]
        db.add(B2BInvoiceItem(
            invoice_id=invoice.id, product_id=product.id,
            qty=item.qty, unit_price=item.unit_price,
            total=round(item.qty * item.unit_price, 2),
        ))
        if is_stock_tracked_product(product):
            before = float(product.stock); after = before - item.qty
            product.stock = after
            db.add(StockMove(
                product_id=product.id, type="out", qty=-item.qty,
                user_id=current_user.id,
                qty_before=before, qty_after=after,
                ref_type="b2b", ref_id=invoice.id,
                note=f"B2B {invoice_number} ({invoice_type})",
            ))

    # ── ACCOUNTING ──────────────────────────────────────
    # Cash invoices: journal is posted when payment is collected (not at creation)
    # AR is always tracked for all types so outstanding balance works
    if invoice_type == "cash":
        if not fully_discounted:
            await _post_journal(db, f"B2B Cash Invoice - {invoice_number}", "b2b", [
                ("1100", total, 0),
                ("4000", 0, total),
            ], user_id=current_user.id)
            client.outstanding = Decimal(str(float(client.outstanding) + total))

    elif invoice_type == "full_payment":
        if not fully_discounted:
            await _post_journal(db, f"B2B Full Payment Invoice - {invoice_number}", "b2b", [
                ("1100", total, 0),
                ("2200", 0, total),
            ], user_id=current_user.id)
            client.outstanding = Decimal(str(float(client.outstanding) + total))

    elif invoice_type == "consignment":
        if not fully_discounted:
            await _post_journal(db, f"B2B Consignment Invoice - {invoice_number}", "b2b", [
                ("1100", total, 0),
                ("2200", 0, total),
            ], user_id=current_user.id)
            client.outstanding = Decimal(str(float(client.outstanding) + total))

        cons_ref = await _next_cons_number(db)
        consignment = Consignment(
            ref_number=cons_ref, client_id=data.client_id,
            invoice_id=invoice.id, user_id=current_user.id, status="active", notes=data.notes,
        )
        db.add(consignment); await db.flush()
        for item in data.items:
            db.add(ConsignmentItem(
                consignment_id=consignment.id, product_id=item.product_id,
                qty_sent=item.qty, qty_sold=0, qty_returned=0,
                unit_price=item.unit_price,
            ))

    record(db, "B2B", "create_invoice",
           f"B2B invoice {invoice_number} — {client.name} — {total:.2f} — {invoice_type}",
           user=current_user, ref_type="b2b_invoice", ref_id=invoice.id)
    await db.commit(); await db.refresh(invoice)
    return {"id": invoice.id, "invoice_number": invoice_number, "total": total}


@router.put("/api/invoices/{invoice_id}", dependencies=[Depends(require_action("b2b", "invoices", "update"))])
async def edit_invoice(invoice_id: int, data: InvoiceCreate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    _r = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(
            selectinload(B2BInvoice.items),
            selectinload(B2BInvoice.client),
        )
    )
    invoice = _r.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if invoice.status == "paid" and float(invoice.amount_paid) > 0:
        raise HTTPException(status_code=400, detail="Cannot edit a paid invoice. Refund first if needed.")
    if invoice.invoice_type == "consignment":
        cons_r = await db.execute(
            select(Consignment).where(Consignment.invoice_id == invoice_id)
            .options(selectinload(Consignment.items))
        )
        cons_chk = cons_r.scalar_one_or_none()
        if cons_chk and any(float(ci.qty_sold) > 0 for ci in cons_chk.items):
            raise HTTPException(status_code=400, detail="Cannot edit a consignment that has sales recorded.")

    await _reverse_invoice_stock(invoice, db)
    await _reverse_invoice_journal(invoice, db)

    for item in invoice.items:
        await db.delete(item)
    old_cons_r = await db.execute(
        select(Consignment).where(Consignment.invoice_id == invoice_id)
        .options(selectinload(Consignment.items))
    )
    old_cons = old_cons_r.scalar_one_or_none()
    if old_cons:
        for ci in old_cons.items:
            await db.delete(ci)
        await db.delete(old_cons)

    products_by_id = await _validate_b2b_invoice_stock(db, data.items)

    _r = await db.execute(select(B2BClient).where(B2BClient.id == data.client_id))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    invoice_type    = _normalized_client_terms(client)
    discount_pct    = _client_discount_pct(client)
    subtotal        = sum(i.qty * i.unit_price for i in data.items)
    discount_amount = round(subtotal * (discount_pct / 100), 2)
    total           = round(subtotal - discount_amount, 2)

    # 100% discount (or any invoice that nets to zero) has nothing to collect,
    # so it is booked as fully paid with no AR / outstanding impact.
    fully_discounted = discount_pct >= 100 or total <= 0.005

    invoice.client_id      = data.client_id
    invoice.user_id        = current_user.id
    invoice.invoice_type   = invoice_type
    invoice.payment_method = invoice_type
    invoice.subtotal       = round(subtotal, 2)
    invoice.discount       = discount_amount
    invoice.total          = total
    invoice.amount_paid    = total if fully_discounted else 0
    invoice.status         = "paid" if fully_discounted else "unpaid"
    invoice.notes          = data.notes

    for item in data.items:
        product = products_by_id[item.product_id]
        db.add(B2BInvoiceItem(
            invoice_id=invoice.id, product_id=product.id,
            qty=item.qty, unit_price=item.unit_price,
            total=round(item.qty * item.unit_price, 2),
        ))
        if is_stock_tracked_product(product):
            before = float(product.stock); after = before - item.qty
            product.stock = after
            db.add(StockMove(
                product_id=product.id, type="out", qty=-item.qty,
                user_id=current_user.id,
                qty_before=before, qty_after=after,
                ref_type="b2b", ref_id=invoice.id,
                note=f"B2B {invoice.invoice_number} (edited)",
            ))

    if invoice_type == "cash":
        if not fully_discounted:
            await _post_journal(db, f"B2B Cash Invoice (edited) - {invoice.invoice_number}", "b2b", [
                ("1100", total, 0),
                ("4000", 0, total),
            ], user_id=current_user.id)
            if client:
                client.outstanding = Decimal(str(float(client.outstanding) + total))
    elif invoice_type in ("full_payment", "consignment"):
        if not fully_discounted:
            await _post_journal(db, f"B2B {invoice_type} Invoice (edited) - {invoice.invoice_number}", "b2b", [
                ("1100", total, 0),
                ("2200", 0, total),
            ], user_id=current_user.id)
            if client:
                client.outstanding = Decimal(str(float(client.outstanding) + total))
        if invoice_type == "consignment":
            cons_ref = await _next_cons_number(db)
            consignment = Consignment(
                ref_number=cons_ref, client_id=data.client_id,
                invoice_id=invoice.id, user_id=current_user.id, status="active", notes=data.notes,
            )
            db.add(consignment); await db.flush()
            for item in data.items:
                db.add(ConsignmentItem(
                    consignment_id=consignment.id, product_id=item.product_id,
                    qty_sent=item.qty, qty_sold=0, qty_returned=0,
                    unit_price=item.unit_price,
                ))

    record(db, "B2B", "edit_invoice",
           f"Edited B2B invoice {invoice.invoice_number} — {total:.2f}",
           user=current_user, ref_type="b2b_invoice", ref_id=invoice_id)
    await db.commit()
    return {"ok": True, "invoice_number": invoice.invoice_number, "total": total}


@router.delete("/api/invoices/{invoice_id}", dependencies=[Depends(require_action("b2b", "invoices", "delete"))])
async def delete_invoice(invoice_id: int, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    _r = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(
            selectinload(B2BInvoice.items),
            selectinload(B2BInvoice.client),
        )
    )
    invoice = _r.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    inv_num = invoice.invoice_number
    await _reverse_invoice_stock(invoice, db)
    await _reverse_invoice_journal(invoice, db)
    cons_r = await db.execute(
        select(Consignment).where(Consignment.invoice_id == invoice_id)
        .options(selectinload(Consignment.items))
    )
    cons = cons_r.scalar_one_or_none()
    if cons:
        for ci in cons.items:
            await db.delete(ci)
        await db.delete(cons)
    await db.delete(invoice)
    record(db, "B2B", "delete_invoice",
           f"Deleted B2B invoice {inv_num} — stock and journal reversed",
           ref_type="b2b_invoice", ref_id=invoice_id)
    await db.commit()
    return {"ok": True}


@router.post("/api/invoices/{invoice_id}/pay", dependencies=[Depends(require_action("b2b", "invoices", "approve"))])
async def record_payment(invoice_id: int, data: PaymentRecord, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    """
    Collect payment on an invoice (cash, full_payment, or consignment).
    For cash: Dr Cash / Cr AR (revenue was already recognised at creation).
    For full_payment: Dr Cash / Cr AR, and Dr Deferred Revenue / Cr Sales Revenue.
    """
    _r = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(selectinload(B2BInvoice.client))
    )
    invoice = _r.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    balance = float(invoice.total) - float(invoice.amount_paid)
    if data.amount > balance + 0.01:
        raise HTTPException(status_code=400, detail=f"Amount exceeds balance: {balance:.2f}")

    amount = round(data.amount, 2)
    invoice.amount_paid = Decimal(str(float(invoice.amount_paid) + amount))
    invoice.status = "paid" if float(invoice.amount_paid) >= float(invoice.total) else "partial"

    client = invoice.client
    client.outstanding = Decimal(str(max(0, float(client.outstanding) - amount)))

    if invoice.invoice_type == "cash":
        # Revenue already recognised at creation (Dr AR / Cr Revenue)
        # Now just collect cash: Dr Cash / Cr AR
        await _post_journal(db, f"Cash collected - {invoice.invoice_number}", "b2b_collection", [
            ("1000", amount, 0),
            ("1100", 0, amount),
        ], user_id=current_user.id, ref_id=invoice.id)
    else:
        # full_payment / consignment: Dr Cash / Cr AR, Dr Deferred Revenue / Cr Revenue
        await _post_journal(db, f"Payment received - {invoice.invoice_number}", "b2b_payment", [
            ("1000", amount, 0),
            ("1100", 0, amount),
            ("2200", amount, 0),
            ("4000", 0, amount),
        ], user_id=current_user.id, ref_id=invoice.id)

    await db.commit()
    return {"ok": True, "status": invoice.status}


@router.post("/api/invoices/{invoice_id}/reverse-payment", dependencies=[Depends(require_action("b2b", "invoices", "approve"))])
async def reverse_payment(
    invoice_id: int,
    data: Optional[PaymentReversal] = None,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    """Undo a payment collected against an invoice, in full or in part.

    Posts a CONTRA journal rather than deleting the original. The original
    entry stays on the ledger and the reversal sits beside it, so the history
    shows what actually happened instead of pretending the payment never
    existed — and the trial balance stays consistent either way.

    The legs mirror record_payment exactly, including the deferred-revenue pair
    that non-cash invoices carry, so reversing a payment in full leaves every
    account exactly where it was before it was collected.
    """
    result = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(selectinload(B2BInvoice.client))
    )
    invoice = result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")

    already_paid = round(float(invoice.amount_paid or 0), 2)
    if already_paid <= 0:
        raise HTTPException(status_code=400, detail="This invoice has no payment to reverse")

    requested = (data.amount if data and data.amount is not None else already_paid)
    amount = round(float(requested), 2)
    if amount <= 0:
        raise HTTPException(status_code=400, detail="Reversal amount must be greater than 0")
    if amount > already_paid + 0.01:
        raise HTTPException(
            status_code=400,
            detail=f"Cannot reverse more than was paid: {already_paid:.2f}",
        )

    remaining = round(already_paid - amount, 2)
    invoice.amount_paid = Decimal(str(remaining))
    if remaining >= float(invoice.total) - 0.005:
        invoice.status = "paid"
    elif remaining > 0:
        invoice.status = "partial"
    else:
        # Back to whatever it was before any payment. A status that is not a
        # payment state (an imported "consignment", say) is left alone rather
        # than being flattened to "unpaid".
        invoice.status = invoice.status if invoice.status not in ("paid", "partial") else "unpaid"

    client = invoice.client
    if client is not None:
        client.outstanding = Decimal(str(round(float(client.outstanding or 0) + amount, 2)))

    reason = (data.reason if data else None) or ""
    reason = reason.strip()
    description = f"Payment reversed - {invoice.invoice_number}"
    if reason:
        description += f" - {reason}"

    # Mirror image of the legs record_payment posted for this invoice type.
    if invoice.invoice_type == "cash":
        entries = [("1100", amount, 0), ("1000", 0, amount)]
        ref_type = "b2b_collection_reversal"
    else:
        entries = [
            ("1100", amount, 0), ("1000", 0, amount),
            ("4000", amount, 0), ("2200", 0, amount),
        ]
        ref_type = "b2b_payment_reversal"
    await _post_journal(db, description, ref_type, entries,
                        user_id=current_user.id, ref_id=invoice.id)

    record(db, "B2B", "reverse_payment",
           f"Reversed {amount:.2f} of payment on {invoice.invoice_number}"
           + (f" - {reason}" if reason else ""),
           user=current_user, ref_type="b2b_invoice", ref_id=invoice.id)
    await db.commit()

    return {
        "ok": True,
        "invoice_number": invoice.invoice_number,
        "reversed": amount,
        "amount_paid": remaining,
        "balance_due": round(float(invoice.total) - remaining, 2),
        "status": invoice.status,
        "client_outstanding": await _client_outstanding_value(db, invoice.client_id),
    }


# ── CONSIGNMENT API ────────────────────────────────────
@router.get("/api/consignments")
async def get_consignments(db: AsyncSession = Depends(get_async_session)):
    _r = await db.execute(
        select(Consignment)
        .options(
            selectinload(Consignment.client),
            selectinload(Consignment.items).selectinload(ConsignmentItem.product),
        )
        .order_by(Consignment.created_at.desc())
    )
    conses = _r.scalars().all()
    return [
        {
            "id":         c.id,
            "ref_number": c.ref_number,
            "client":     c.client.name if c.client else "—",
            "client_id":  c.client_id,
            "status":     c.status,
            "created_at": c.created_at.strftime("%Y-%m-%d") if c.created_at else "—",
            "notes":      c.notes or "",
            "items": [
                {
                    "id":           ci.id,
                    "product":      ci.product.name if ci.product else "—",
                    "product_id":   ci.product_id,
                    "unit_price":   float(ci.unit_price),
                    "qty_sent":     float(ci.qty_sent),
                    "qty_sold":     float(ci.qty_sold),
                    "qty_returned": float(ci.qty_returned),
                    "qty_pending":  float(ci.qty_sent) - float(ci.qty_sold) - float(ci.qty_returned),
                    "revenue":      float(ci.qty_sold) * float(ci.unit_price),
                }
                for ci in c.items
            ],
            "total_sent":     sum(float(ci.qty_sent)     for ci in c.items),
            "total_sold":     sum(float(ci.qty_sold)     for ci in c.items),
            "total_returned": sum(float(ci.qty_returned) for ci in c.items),
            "total_revenue":  sum(float(ci.qty_sold) * float(ci.unit_price) for ci in c.items),
        }
        for c in conses
    ]

@router.post("/api/consignments/{cons_id}/settle", dependencies=[Depends(require_action("b2b", "invoices", "settle"))])
async def settle_consignment(cons_id: int, data: ConsignmentSettle, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    """
    Settle consignment — for each qty sold, move from Deferred Revenue → Sales Revenue.
    Returned items restore stock.
    """
    _r = await db.execute(
        select(Consignment)
        .where(Consignment.id == cons_id)
        .options(
            selectinload(Consignment.items).selectinload(ConsignmentItem.product),
            selectinload(Consignment.client),
        )
    )
    cons = _r.scalar_one_or_none()
    if not cons:
        raise HTTPException(status_code=404, detail="Consignment not found")
    if cons.status == "closed":
        raise HTTPException(status_code=400, detail="Consignment already closed")

    total_revenue = 0
    for entry in data.items:
        ci_r = await db.execute(
            select(ConsignmentItem)
            .where(ConsignmentItem.id == entry["consignment_item_id"])
            .options(selectinload(ConsignmentItem.product))
        )
        ci = ci_r.scalar_one_or_none()
        if not ci: continue
        qty_sold     = float(entry.get("qty_sold", 0))
        qty_returned = float(entry.get("qty_returned", 0))
        pending      = float(ci.qty_sent) - float(ci.qty_sold) - float(ci.qty_returned)
        if qty_sold + qty_returned > pending + 0.001:
            raise HTTPException(status_code=400,
                detail=f"Total exceeds pending for {ci.product.name}. Pending: {pending:.2f}")
        ci.qty_sold     = Decimal(str(float(ci.qty_sold)     + qty_sold))
        ci.qty_returned = Decimal(str(float(ci.qty_returned) + qty_returned))
        if qty_returned > 0 and is_stock_tracked_product(ci.product):
            product = ci.product
            before  = float(product.stock); after = before + qty_returned
            product.stock = after
            db.add(StockMove(
                product_id=product.id, type="in",
                user_id=current_user.id,
                qty=qty_returned, qty_before=before, qty_after=after,
                ref_type="consignment_return", ref_id=cons.id,
                note=f"Returned from {cons.ref_number}",
            ))
        total_revenue += qty_sold * float(ci.unit_price)

    if total_revenue > 0:
        amount = round(total_revenue, 2)
        # Deferred Revenue → Sales Revenue (earned on settlement)
        # Cash ← AR (client paid for what they sold)
        await _post_journal(db, f"Consignment settlement - {cons.ref_number}", "consignment_settlement", [
            ("1000", amount, 0),
            ("1100", 0, amount),
            ("2200", amount, 0),
            ("4000", 0, amount),
        ], user_id=current_user.id)
        cons.client.outstanding = Decimal(str(max(0, float(cons.client.outstanding) - amount)))

    all_done = all(
        float(ci.qty_sold) + float(ci.qty_returned) >= float(ci.qty_sent)
        for ci in cons.items
    )
    cons.status = "closed" if all_done else "active"
    if all_done:
        cons.settled_at = datetime.now(timezone.utc)

    await db.commit()
    return {"ok": True, "total_revenue": round(total_revenue, 2), "status": cons.status}


class ConsignmentPayment(BaseModel):
    amount:      float
    month_label: Optional[str] = None
    notes:       Optional[str] = None

@router.post("/api/invoices/{invoice_id}/consignment-payment", dependencies=[Depends(require_action("b2b", "invoices", "approve"))])
async def consignment_payment(invoice_id: int, data: ConsignmentPayment, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    """
    Record a cash payment from a consignment client.
    Moves amount: Deferred Revenue → Sales Revenue, Cash ← AR.
    """
    _r = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(selectinload(B2BInvoice.client))
    )
    invoice = _r.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if invoice.invoice_type != "consignment":
        raise HTTPException(status_code=400, detail="This endpoint is for consignment invoices only")

    amount = round(data.amount, 2)
    balance = float(invoice.total) - float(invoice.amount_paid)
    if amount > balance + 0.01:
        raise HTTPException(status_code=400, detail=f"Amount exceeds remaining balance: {balance:.2f}")

    invoice.amount_paid = Decimal(str(float(invoice.amount_paid) + amount))
    if float(invoice.amount_paid) >= float(invoice.total):
        invoice.status = "paid"

    client = invoice.client
    client.outstanding = Decimal(str(max(0, float(client.outstanding) - amount)))

    note = f"Consignment payment - {invoice.invoice_number}"
    if data.month_label:
        note += f" - {data.month_label}"

    await _post_journal(db, note, "consignment_payment", [
        ("1000", amount, 0),
        ("1100", 0, amount),
        ("2200", amount, 0),
        ("4000", 0, amount),
    ], user_id=current_user.id, ref_id=invoice.id)

    await db.commit()
    return {"ok": True, "invoice_number": invoice.invoice_number, "amount": amount, "status": invoice.status}

async def create_client_refund_core(db: AsyncSession, current_user: User, data: ClientRefundCreate):
    """Create a B2B client refund: refund record + line items, stock returned,
    journal posted, client balance credited.

    Shared so the Accounting page issues exactly the same refund as the B2B
    page — one code path, one set of side effects, one refund record.
    """
    _r = await db.execute(select(B2BClient).where(B2BClient.id == data.client_id, B2BClient.is_active == True))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    if not data.items:
        raise HTTPException(status_code=400, detail="Refund must have at least one item")

    refund_number = await _next_refund_number(db)
    subtotal = 0.0
    for item in data.items:
        _r = await db.execute(select(Product).where(Product.id == item.product_id))
        product = _r.scalar_one_or_none()
        if not product:
            raise HTTPException(status_code=404, detail=f"Product not found: {item.product_id}")
        if item.qty <= 0:
            raise HTTPException(status_code=400, detail="Refund quantities must be greater than 0")
        if item.unit_price < 0:
            raise HTTPException(status_code=400, detail="Unit price cannot be negative")
        subtotal += round(item.qty * item.unit_price, 2)

    subtotal = round(subtotal, 2)
    discount_pct = _client_discount_pct(client)
    discount = round(subtotal * (discount_pct / 100), 2)
    total = round(subtotal - discount, 2)
    if total <= 0:
        raise HTTPException(status_code=400, detail="Refund total must be greater than 0")
    # Check against the live balance, the same one the clients list shows. The
    # stored client.outstanding drifts (it is not maintained by every path), so
    # using it here rejected valid refunds and allowed invalid ones.
    live_outstanding = await _client_outstanding_value(db, client.id)
    if total > live_outstanding + 0.01:
        raise HTTPException(
            status_code=400,
            detail=f"Refund exceeds client outstanding: {live_outstanding:.2f}",
        )

    refund = B2BRefund(
        refund_number=refund_number,
        client_id=client.id,
        user_id=current_user.id,
        subtotal=subtotal,
        discount=discount,
        total=total,
        notes=(data.notes or "").strip() or None,
    )
    db.add(refund); await db.flush()

    for item in data.items:
        _r = await db.execute(select(Product).where(Product.id == item.product_id))
        product = _r.scalar_one_or_none()
        line_total = round(item.qty * item.unit_price, 2)
        db.add(B2BRefundItem(
            refund_id=refund.id,
            product_id=product.id,
            qty=item.qty,
            unit_price=item.unit_price,
            total=line_total,
        ))
        if is_stock_tracked_product(product):
            before = float(product.stock)
            after  = before + item.qty
            product.stock = after
            db.add(StockMove(
                product_id=product.id, type="in", qty=float(item.qty),
                user_id=current_user.id,
                qty_before=before, qty_after=after,
                ref_type="b2b_refund", ref_id=refund.id,
                note=f"B2B refund {refund_number} - {client.name}",
            ))

    client.outstanding = Decimal(str(max(0, float(client.outstanding) - total)))

    note = (data.notes or "").strip()
    desc = f"B2B client refund - {refund_number} - {client.name}"
    if note:
        desc += f" - {note}"
    await _post_journal(db, desc, "b2b_refund", [
        ("2200", total, 0),
        ("1100", 0, total),
    ], user_id=current_user.id)

    await db.commit()
    return {
        "ok": True,
        "refund_id": refund.id,
        "refund_number": refund_number,
        "client": client.name,
        "subtotal": subtotal,
        "discount": discount,
        "discount_pct": discount_pct,
        "amount": total,
        "outstanding": await _client_outstanding_value(db, client.id),
    }


@router.post("/api/refunds", dependencies=[Depends(require_action("b2b", "invoices", "refund"))])
async def create_client_refund(
    data: ClientRefundCreate,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    return await create_client_refund_core(db, current_user, data)


# ── STATS ──────────────────────────────────────────────
@router.get("/api/stats")
async def get_stats(db: AsyncSession = Depends(get_async_session)):
    r1 = await db.execute(select(func.count(B2BClient.id)).where(B2BClient.is_active == True))
    r2 = await db.execute(
        select(func.sum(B2BInvoice.total - B2BInvoice.amount_paid))
        .where(B2BInvoice.status.in_(["unpaid", "partial"]))
    )
    r3 = await db.execute(
        select(func.count(B2BInvoice.id))
        .where(B2BInvoice.status.in_(["unpaid", "partial"]))
    )
    r4 = await db.execute(select(func.count(Consignment.id)).where(Consignment.status == "active"))
    return {
        "total_clients":     r1.scalar() or 0,
        "total_outstanding": float(r2.scalar() or 0),
        "unpaid_invoices":   r3.scalar() or 0,
        "active_consign":    r4.scalar() or 0,
    }


@router.get("/api/client-analysis", dependencies=[Depends(require_permission("tab_b2b_clients"))])
async def get_client_analysis(db: AsyncSession = Depends(get_async_session)):
    clients_res = await db.execute(
        select(B2BClient)
        .where(B2BClient.is_active == True)
        .options(selectinload(B2BClient.invoices))
        .order_by(B2BClient.name)
    )
    clients = clients_res.scalars().all()

    refunds_res = await db.execute(
        select(
            B2BRefund.client_id,
            func.coalesce(func.sum(B2BRefund.total), 0).label("refund_total"),
        )
        .group_by(B2BRefund.client_id)
    )
    refunds_by_client = {
        client_id: _b2b_num(refund_total)
        for client_id, refund_total in refunds_res.all()
    }
    
    top_products_by_client = await _get_b2b_client_top_products(db)

    today = date.today()
    six_months_ago = datetime.now(timezone.utc) - timedelta(days=180)
    
    client_rows = []
    terms_breakdown = {}
    for client in clients:
        invoices = list(client.invoices or [])
        invoice_count = len(invoices)
        gross_sales = sum(_b2b_num(invoice.total) for invoice in invoices)
        paid_amount = sum(_b2b_num(invoice.amount_paid) for invoice in invoices)
        invoice_outstanding = sum(
            max(_b2b_num(invoice.total) - _b2b_num(invoice.amount_paid), 0.0)
            for invoice in invoices
            if (invoice.status or "").lower() in {"unpaid", "partial"} or _b2b_num(invoice.total) > _b2b_num(invoice.amount_paid)
        )
        refund_total = refunds_by_client.get(client.id, 0.0)
        outstanding = max(invoice_outstanding - refund_total, 0.0)
        net_sales = gross_sales - refund_total
        average_invoice = gross_sales / invoice_count if invoice_count else 0.0
        payment_rate = (paid_amount / gross_sales * 100) if gross_sales > 0 else 0.0
        last_invoice = max((_date_key(invoice.created_at) for invoice in invoices), default=datetime.min)
        last_invoice_label = last_invoice.strftime("%Y-%m-%d") if last_invoice != datetime.min else "—"
        days_since_last_invoice = (today - last_invoice.date()).days if last_invoice != datetime.min else None
        credit_limit = _b2b_num(client.credit_limit)
        credit_used_pct = (outstanding / credit_limit * 100) if credit_limit > 0 else None
        terms = client.payment_terms or "cash"
        terms_bucket = terms_breakdown.setdefault(terms, {"clients": 0, "gross_sales": 0.0, "outstanding": 0.0})
        terms_bucket["clients"] += 1
        terms_bucket["gross_sales"] += gross_sales
        terms_bucket["outstanding"] += outstanding

        if outstanding > 0 and credit_limit > 0 and outstanding > credit_limit:
            risk_level = "over_limit"
        elif outstanding > 0 and days_since_last_invoice is not None and days_since_last_invoice >= 45:
            risk_level = "stale_outstanding"
        elif outstanding > 0:
            risk_level = "collect"
        elif invoice_count == 0:
            risk_level = "new"
        elif days_since_last_invoice is not None and days_since_last_invoice >= 60:
            risk_level = "quiet"
        else:
            risk_level = "healthy"

        # Compute new analysis fields
        return_rate = (refund_total / gross_sales * 100) if gross_sales > 0 else 0.0
        
        trends = {}
        for inv in invoices:
            inv_date = _date_key(inv.created_at)
            if inv_date >= six_months_ago.replace(tzinfo=None):
                month_str = inv_date.strftime("%Y-%m")
                trends[month_str] = trends.get(month_str, 0.0) + _b2b_num(inv.total)
        purchase_trends = [{"month": m, "volume": round(v, 2)} for m, v in sorted(trends.items())]

        client_rows.append({
            "id": client.id,
            "name": client.name,
            "contact_person": client.contact_person or "—",
            "phone": client.phone or "—",
            "payment_terms": terms,
            "invoice_count": invoice_count,
            "gross_sales": round(gross_sales, 2),
            "refunds": round(refund_total, 2),
            "net_sales": round(net_sales, 2),
            "paid_amount": round(paid_amount, 2),
            "outstanding": round(outstanding, 2),
            "average_invoice": round(average_invoice, 2),
            "payment_rate": round(payment_rate, 1),
            "credit_limit": round(credit_limit, 2),
            "credit_used_pct": round(credit_used_pct, 1) if credit_used_pct is not None else None,
            "last_invoice": last_invoice_label,
            "days_since_last_invoice": days_since_last_invoice,
            "risk_level": risk_level,
            # New fields
            "ltv": round(net_sales, 2),
            "total_outstanding": round(outstanding, 2),
            "average_order_value": round(average_invoice, 2),
            "return_rate": round(return_rate, 2),
            "purchase_trends": purchase_trends,
            "top_products": top_products_by_client.get(client.id, []),
        })

    total_gross = sum(row["gross_sales"] for row in client_rows)
    total_refunds = sum(row["refunds"] for row in client_rows)
    total_net = sum(row["net_sales"] for row in client_rows)
    total_paid = sum(row["paid_amount"] for row in client_rows)
    total_outstanding = sum(row["outstanding"] for row in client_rows)
    top_client = max(client_rows, key=lambda row: row["net_sales"], default=None)
    summary = {
        "active_clients": len(client_rows),
        "clients_with_sales": sum(1 for row in client_rows if row["invoice_count"] > 0),
        "gross_sales": round(total_gross, 2),
        "refunds": round(total_refunds, 2),
        "net_sales": round(total_net, 2),
        "paid_amount": round(total_paid, 2),
        "outstanding": round(total_outstanding, 2),
        "payment_rate": round((total_paid / total_gross * 100) if total_gross > 0 else 0.0, 1),
        "at_risk_clients": sum(1 for row in client_rows if row["risk_level"] in {"over_limit", "stale_outstanding", "collect"}),
        "top_client": top_client["name"] if top_client else "—",
        "top_client_net_sales": top_client["net_sales"] if top_client else 0.0,
    }
    return {
        "summary": summary,
        "clients": sorted(client_rows, key=lambda row: row["net_sales"], reverse=True),
        "top_clients": sorted(client_rows, key=lambda row: row["net_sales"], reverse=True)[:5],
        "collection_watch": sorted(
            [row for row in client_rows if row["outstanding"] > 0],
            key=lambda row: row["outstanding"],
            reverse=True,
        )[:5],
        "terms_breakdown": [
            {
                "payment_terms": terms,
                "clients": data["clients"],
                "gross_sales": round(data["gross_sales"], 2),
                "outstanding": round(data["outstanding"], 2),
            }
            for terms, data in sorted(terms_breakdown.items())
        ],
    }


@router.get("/api/products-list")
async def products_list(client_id: int = None, db: AsyncSession = Depends(get_async_session)):
    _r = await db.execute(select(Product).where(Product.is_active == True).order_by(Product.name))
    products = _r.scalars().all()
    custom = {}
    if client_id:
        cp_r = await db.execute(select(B2BClientPrice).where(B2BClientPrice.client_id == client_id))
        for cp in cp_r.scalars().all():
            custom[cp.product_id] = float(cp.price)
    return [
        {
            "id":            p.id,
            "name":          p.name,
            "sku":           p.sku,
            "price":         custom.get(p.id, float(p.price)),
            "default_price": float(p.price),
            "has_custom":    p.id in custom,
            "stock":         float(p.stock),
            "item_type":     normalize_item_type(p.item_type),
            "stock_tracked":  is_stock_tracked_product(p),
            "unit":          p.unit,
        }
        for p in products
    ]


# ── CLIENT PRICE LIST API ──────────────────────────────
@router.get("/api/clients/{client_id}/prices")
async def get_client_prices(client_id: int, db: AsyncSession = Depends(get_async_session)):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    cp_r = await db.execute(
        select(B2BClientPrice)
        .where(B2BClientPrice.client_id == client_id)
        .options(selectinload(B2BClientPrice.product))
    )
    prices = cp_r.scalars().all()
    return [
        {
            "id":            cp.id,
            "product_id":    cp.product_id,
            "product_name":  cp.product.name if cp.product else "—",
            "sku":           cp.product.sku  if cp.product else "—",
            "custom_price":  float(cp.price),
            "default_price": float(cp.product.price) if cp.product else 0,
        }
        for cp in prices
    ]


class ClientPriceUpsert(BaseModel):
    product_id: int
    price:      float

@router.put("/api/clients/{client_id}/prices", dependencies=[Depends(require_permission("action_b2b_clients_update"))])
async def upsert_client_price(client_id: int, data: ClientPriceUpsert,
                               db: AsyncSession = Depends(get_async_session),
                               current_user: User = Depends(get_current_user)):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    if data.price < 0:
        raise HTTPException(status_code=400, detail="Price must be >= 0")
    cp_r = await db.execute(select(B2BClientPrice).where(
        B2BClientPrice.client_id == client_id,
        B2BClientPrice.product_id == data.product_id,
    ))
    cp = cp_r.scalar_one_or_none()
    if cp:
        cp.price = data.price
    else:
        db.add(B2BClientPrice(client_id=client_id, product_id=data.product_id, price=data.price))
    await db.commit()
    return {"ok": True}


@router.delete("/api/clients/{client_id}/prices/{product_id}", dependencies=[Depends(require_permission("action_b2b_clients_update"))])
async def delete_client_price(client_id: int, product_id: int,
                               db: AsyncSession = Depends(get_async_session),
                               current_user: User = Depends(get_current_user)):
    cp_r = await db.execute(select(B2BClientPrice).where(
        B2BClientPrice.client_id == client_id,
        B2BClientPrice.product_id == product_id,
    ))
    cp = cp_r.scalar_one_or_none()
    if cp:
        await db.delete(cp)
        await db.commit()
    return {"ok": True}

@router.get("/api/refund-products/{client_id}")
async def refund_products(client_id: int, db: AsyncSession = Depends(get_async_session)):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id, B2BClient.is_active == True))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    latest_prices = {}
    items_r = await db.execute(
        select(B2BInvoiceItem, B2BInvoice)
        .join(B2BInvoice, B2BInvoice.id == B2BInvoiceItem.invoice_id)
        .where(B2BInvoice.client_id == client_id)
        .order_by(B2BInvoice.created_at.desc(), B2BInvoice.id.desc(), B2BInvoiceItem.id.desc())
    )
    for item, _inv in items_r.all():
        if item.product_id not in latest_prices:
            latest_prices[item.product_id] = float(item.unit_price)

    prod_r = await db.execute(select(Product).where(Product.is_active == True).order_by(Product.name))
    products = prod_r.scalars().all()
    return [
        {
            "id":    p.id,
            "name":  p.name,
            "sku":   p.sku,
            "price": latest_prices.get(p.id, float(p.price)),
            "stock": float(p.stock),
            "item_type": normalize_item_type(p.item_type),
            "stock_tracked": is_stock_tracked_product(p),
            "unit":  p.unit,
        }
        for p in products
    ]

@router.get("/api/refunds")
async def get_refunds(client_id: int = None, db: AsyncSession = Depends(get_async_session)):
    stmt = (
        select(B2BRefund)
        .options(
            selectinload(B2BRefund.client),
            selectinload(B2BRefund.items).selectinload(B2BRefundItem.product),
        )
        .order_by(B2BRefund.created_at.desc(), B2BRefund.id.desc())
    )
    if client_id:
        stmt = stmt.where(B2BRefund.client_id == client_id)
    _r = await db.execute(stmt)
    refunds = _r.scalars().all()
    return [
        {
            "id":           r.id,
            "refund_number": r.refund_number,
            "client":       r.client.name if r.client else "—",
            "client_id":    r.client_id,
            "subtotal":     float(r.subtotal),
            "discount":     float(r.discount),
            "discount_pct": round(float(r.discount) / float(r.subtotal) * 100, 1) if float(r.subtotal) > 0 else 0,
            "total":        float(r.total),
            "notes":        r.notes or "",
            "created_at":   r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else "—",
            "items": [
                {
                    "product":    item.product.name if item.product else "—",
                    "sku":        item.product.sku if item.product else "—",
                    "qty":        float(item.qty),
                    "unit_price": float(item.unit_price),
                    "total":      float(item.total),
                }
                for item in r.items
            ],
        }
        for r in refunds
    ]


@router.delete("/api/refunds/{refund_id}", dependencies=[Depends(require_admin)])
async def delete_refund(
    refund_id: int,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    _r = await db.execute(
        select(B2BRefund)
        .where(B2BRefund.id == refund_id)
        .options(
            selectinload(B2BRefund.client),
            selectinload(B2BRefund.items).selectinload(B2BRefundItem.product),
        )
    )
    refund = _r.scalar_one_or_none()
    if not refund:
        raise HTTPException(status_code=404, detail="Refund not found")

    await _reverse_refund_effects(refund, db, current_user)

    record(
        db,
        "B2B",
        "delete_refund",
        f"Deleted refund {refund.refund_number} for {refund.client.name if refund.client else 'Unknown client'}",
        current_user,
        "b2b_refund",
        refund.id,
    )
    await db.delete(refund)
    await db.commit()
    return {"ok": True, "refund_number": refund.refund_number}


@router.get("/invoice/{invoice_id}/print", response_class=HTMLResponse)
async def print_invoice(
    invoice_id: int,
    request: Request,
    db: AsyncSession = Depends(get_async_session),
):
    _r = await db.execute(
        select(B2BInvoice)
        .where(B2BInvoice.id == invoice_id)
        .options(
            selectinload(B2BInvoice.client),
            selectinload(B2BInvoice.items).selectinload(B2BInvoiceItem.product),
        )
    )
    inv = _r.scalar_one_or_none()
    if not inv:
        raise HTTPException(status_code=404, detail="Invoice not found")

    client = inv.client
    subtotal = float(inv.subtotal or 0)
    discount = float(inv.discount or 0)
    total = float(inv.total or 0)
    discount_pct = round(discount / subtotal * 100, 1) if subtotal > 0 else 0.0

    return templates.TemplateResponse(
        request,
        "b2b_invoice_print.html",
        {
            "invoice": inv,
            "client_name": client.name if client else "—",
            "client_code": f"C{str(client.id).zfill(4)}" if client else "—",
            "subtotal": subtotal,
            "discount": discount,
            "total": total,
            "discount_pct": discount_pct,
        },
    )

@router.get("/refund/{refund_id}/print", response_class=HTMLResponse)
async def print_refund(
    refund_id: int,
    request: Request,
    db: AsyncSession = Depends(get_async_session),
):
    _r = await db.execute(
        select(B2BRefund)
        .where(B2BRefund.id == refund_id)
        .options(
            selectinload(B2BRefund.client),
            selectinload(B2BRefund.items).selectinload(B2BRefundItem.product),
        )
    )
    refund = _r.scalar_one_or_none()
    if not refund:
        raise HTTPException(status_code=404, detail="Refund not found")

    client = refund.client
    subtotal = float(refund.subtotal or 0)
    discount = float(refund.discount or 0)
    total = float(refund.total or 0)
    discount_pct = round(discount / subtotal * 100, 1) if subtotal > 0 else 0.0

    return templates.TemplateResponse(
        request,
        "b2b_refund_print.html",
        {
            "refund": refund,
            "client_name": client.name if client else "—",
            "client_code": f"C{str(client.id).zfill(4)}" if client else "—",
            "subtotal": subtotal,
            "discount": discount,
            "total": total,
            "discount_pct": discount_pct,
        },
    )

# â”€â”€ CLIENT STATEMENT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async def _build_client_statement_payload(
    client_id: int,
    db: AsyncSession,
    *,
    as_of: Optional[date] = None,
):
    _r = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = _r.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    invoice_stmt = (
        select(B2BInvoice)
        .where(B2BInvoice.client_id == client_id)
        .order_by(B2BInvoice.created_at)
    )
    refund_stmt = (
        select(B2BRefund)
        .where(B2BRefund.client_id == client_id)
        .order_by(B2BRefund.created_at)
    )
    if as_of:
        cutoff = datetime.combine(as_of + timedelta(days=1), time.min, tzinfo=timezone.utc)
        invoice_stmt = invoice_stmt.where(B2BInvoice.created_at < cutoff)
        refund_stmt = refund_stmt.where(B2BRefund.created_at < cutoff)

    invoices = (await db.execute(invoice_stmt)).scalars().all()
    refunds = (await db.execute(refund_stmt)).scalars().all()
    payments = await _load_client_payment_activity(db, client_id=client_id, as_of=as_of)

    txns = []
    for inv in invoices:
        txns.append({
            "date": inv.created_at,
            "ref": inv.invoice_number,
            "type": "invoice",
            "desc": f"{(inv.invoice_type or 'b2b').replace('_', ' ').title()} Invoice",
            "debit": float(inv.total or 0),
            "credit": float(inv.amount_paid or 0),
            "status": inv.status,
        })
    for rfnd in refunds:
        txns.append({
            "date": rfnd.created_at,
            "ref": rfnd.refund_number,
            "type": "refund",
            "desc": "Credit / Refund",
            "debit": 0.0,
            "credit": float(rfnd.total or 0),
            "status": "refund",
        })

    txns.sort(key=lambda x: x["date"] or datetime.min.replace(tzinfo=timezone.utc))

    running = 0.0
    rows = []
    for t in txns:
        running += t["debit"] - t["credit"]
        rows.append({
            "date": t["date"].strftime("%d-%b-%Y") if t["date"] else "-",
            "ref": t["ref"],
            "type": t["type"],
            "desc": t["desc"],
            "debit": round(float(t["debit"] or 0), 2),
            "credit": round(float(t["credit"] or 0), 2),
            "balance": round(running, 2),
            "status": t["status"],
        })

    statement_date = as_of or date.today()
    return {
        "client": {
            "id": client.id,
            "code": f"C{str(client.id).zfill(4)}",
            "name": client.name,
            "contact_person": client.contact_person or "",
            "phone": client.phone or "",
            "email": client.email or "",
            "address": client.address or "",
            "payment_terms": client.payment_terms or "",
            "credit_limit": float(client.credit_limit or 0),
            "outstanding": round(running, 2),
        },
        "statement_date": statement_date.strftime("%d-%b-%Y"),
        "statement_period_label": f"As of {statement_date.strftime('%d-%b-%Y')}",
        "transactions": rows,
        "payment_activity": payments,
        "total_invoiced": round(sum(t["debit"] for t in rows), 2),
        "total_paid": round(sum(t["credit"] for t in rows), 2),
        "balance_due": round(running, 2),
        "as_of": as_of.isoformat() if as_of else None,
    }


# ── Client portal link (shareable, read-only, revocable) ────────────────────
def _portal_path(token: str) -> str:
    return f"/portal/c/{token}"


def _portal_url(request: Request, token: str) -> str:
    """Absolute URL to hand the client. Built from the request so it is correct
    behind a proxy/custom domain without extra configuration."""
    return str(request.base_url).rstrip("/") + _portal_path(token)


def _portal_state(client: B2BClient, request: Request) -> dict:
    enabled = bool(client.portal_enabled and client.portal_token)
    return {
        "enabled":        enabled,
        "url":            _portal_url(request, client.portal_token) if enabled else None,
        "created_at":     client.portal_created_at.strftime("%d-%b-%Y %H:%M") if client.portal_created_at else None,
        "last_viewed_at": client.portal_last_viewed_at.strftime("%d-%b-%Y %H:%M") if client.portal_last_viewed_at else None,
        "view_count":     int(client.portal_view_count or 0),
    }


async def _get_client_or_404(db: AsyncSession, client_id: int) -> B2BClient:
    result = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = result.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    return client


@router.get("/api/clients/{client_id}/portal")
async def get_client_portal_link(
    client_id: int,
    request: Request,
    db: AsyncSession = Depends(get_async_session),
):
    client = await _get_client_or_404(db, client_id)
    return _portal_state(client, request)


@router.post("/api/clients/{client_id}/portal", dependencies=[Depends(require_action("b2b", "clients", "update_client"))])
async def create_client_portal_link(
    client_id: int,
    request: Request,
    rotate: bool = False,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    """Issue (or re-issue) the client's portal link.

    Without ``rotate`` an existing token is kept, so re-opening the share dialog
    hands out the same URL the client already bookmarked. With ``rotate=true``
    a fresh token is minted and the old link stops working immediately.
    """
    import secrets

    client = await _get_client_or_404(db, client_id)
    rotated = False
    if rotate or not client.portal_token:
        client.portal_token = secrets.token_urlsafe(32)
        client.portal_created_at = datetime.now(timezone.utc)
        client.portal_view_count = 0
        client.portal_last_viewed_at = None
        rotated = True
    client.portal_enabled = True

    record(db, "B2B", "portal_link",
           f"{'Rotated' if rotated else 'Enabled'} client portal link — {client.name}",
           user=current_user, ref_type="b2b_client", ref_id=client.id)
    await db.commit()
    await db.refresh(client)
    return {**_portal_state(client, request), "rotated": rotated}


@router.delete("/api/clients/{client_id}/portal", dependencies=[Depends(require_action("b2b", "clients", "update_client"))])
async def revoke_client_portal_link(
    client_id: int,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    """Kill the link. Clears the token as well as the flag, so the old URL 404s
    even if the flag is ever flipped back on."""
    client = await _get_client_or_404(db, client_id)
    client.portal_token = None
    client.portal_enabled = False
    client.portal_created_at = None

    record(db, "B2B", "portal_link", f"Revoked client portal link — {client.name}",
           user=current_user, ref_type="b2b_client", ref_id=client.id)
    await db.commit()
    return {"ok": True, "enabled": False}


async def _build_client_products_payload(
    client_id: int,
    db: AsyncSession,
    *,
    as_of: Optional[date] = None,
):
    """
    Everything the client has physically received, netted against returns.

    Sources, chosen so nothing is counted twice:
      • invoice items — every invoice type, including consignment invoices
      • consignment items (qty_sent) — ONLY for consignments with no linked
        invoice, because a consignment invoice already writes the same lines
        onto both the invoice and the consignment
      • refund items and consignment qty_returned — subtracted; the settle
        flow moves stock back without writing a refund, so these never overlap

    Returns a per-product roll-up plus a dated delivery log.
    """
    cutoff = (
        datetime.combine(as_of + timedelta(days=1), time.min, tzinfo=timezone.utc)
        if as_of else None
    )

    invoice_stmt = (
        select(B2BInvoice)
        .where(B2BInvoice.client_id == client_id)
        .options(selectinload(B2BInvoice.items).selectinload(B2BInvoiceItem.product))
        .order_by(B2BInvoice.created_at)
    )
    consignment_stmt = (
        select(Consignment)
        .where(Consignment.client_id == client_id)
        .options(selectinload(Consignment.items).selectinload(ConsignmentItem.product))
        .order_by(Consignment.created_at)
    )
    refund_stmt = (
        select(B2BRefund)
        .where(B2BRefund.client_id == client_id)
        .options(selectinload(B2BRefund.items).selectinload(B2BRefundItem.product))
        .order_by(B2BRefund.created_at)
    )
    if cutoff is not None:
        invoice_stmt = invoice_stmt.where(B2BInvoice.created_at < cutoff)
        consignment_stmt = consignment_stmt.where(Consignment.created_at < cutoff)
        refund_stmt = refund_stmt.where(B2BRefund.created_at < cutoff)

    invoices = (await db.execute(invoice_stmt)).scalars().all()
    consignments = (await db.execute(consignment_stmt)).scalars().all()
    refunds = (await db.execute(refund_stmt)).scalars().all()

    products: dict[int, dict] = {}

    def bucket(product, product_id):
        entry = products.get(product_id)
        if entry is None:
            entry = {
                "product_id":    product_id,
                "name":          getattr(product, "name", None) or f"Product #{product_id}",
                "sku":           getattr(product, "sku", None) or "",
                "unit":          getattr(product, "unit", None) or "",
                "qty_received":  0.0,
                "qty_returned":  0.0,
                "value_received": 0.0,
                "value_returned": 0.0,
                "last_received": None,
            }
            products[product_id] = entry
        return entry

    def note_received(product, product_id, qty, value, when):
        entry = bucket(product, product_id)
        entry["qty_received"] += qty
        entry["value_received"] += value
        if when and (entry["last_received"] is None or when > entry["last_received"]):
            entry["last_received"] = when

    def note_returned(product, product_id, qty, value):
        entry = bucket(product, product_id)
        entry["qty_returned"] += qty
        entry["value_returned"] += value

    deliveries = []

    for inv in invoices:
        lines = []
        for it in inv.items:
            qty, value = float(it.qty or 0), float(it.total or 0)
            note_received(it.product, it.product_id, qty, value, inv.created_at)
            lines.append({
                "product":    getattr(it.product, "name", None) or f"Product #{it.product_id}",
                "qty":        round(qty, 3),
                "unit":       getattr(it.product, "unit", None) or "",
                "unit_price": round(float(it.unit_price or 0), 2),
                "total":      round(value, 2),
            })
        if not lines:
            continue
        deliveries.append({
            "date":     inv.created_at,
            "date_str": inv.created_at.strftime("%d-%b-%Y") if inv.created_at else "—",
            "ref":      inv.invoice_number or f"INV-{inv.id}",
            "kind":     "delivery",
            "label":    f"{(inv.invoice_type or 'b2b').replace('_', ' ').title()} Invoice",
            "items":    lines,
            "total":    round(sum(l["total"] for l in lines), 2),
        })

    for cons in consignments:
        lines = []
        for ci in cons.items:
            # Consignment invoices already contributed these lines above.
            if cons.invoice_id is None:
                qty, price = float(ci.qty_sent or 0), float(ci.unit_price or 0)
                if qty:
                    note_received(ci.product, ci.product_id, qty, qty * price, cons.created_at)
                    lines.append({
                        "product":    getattr(ci.product, "name", None) or f"Product #{ci.product_id}",
                        "qty":        round(qty, 3),
                        "unit":       getattr(ci.product, "unit", None) or "",
                        "unit_price": round(price, 2),
                        "total":      round(qty * price, 2),
                    })
            returned = float(ci.qty_returned or 0)
            if returned:
                note_returned(ci.product, ci.product_id, returned, returned * float(ci.unit_price or 0))
        if not lines:
            continue
        deliveries.append({
            "date":     cons.created_at,
            "date_str": cons.created_at.strftime("%d-%b-%Y") if cons.created_at else "—",
            "ref":      cons.ref_number or f"CONS-{cons.id}",
            "kind":     "delivery",
            "label":    "Consignment Delivery",
            "items":    lines,
            "total":    round(sum(l["total"] for l in lines), 2),
        })

    for rfnd in refunds:
        lines = []
        for it in rfnd.items:
            qty, value = float(it.qty or 0), float(it.total or 0)
            note_returned(it.product, it.product_id, qty, value)
            lines.append({
                "product":    getattr(it.product, "name", None) or f"Product #{it.product_id}",
                "qty":        round(qty, 3),
                "unit":       getattr(it.product, "unit", None) or "",
                "unit_price": round(float(it.unit_price or 0), 2),
                "total":      round(value, 2),
            })
        if not lines:
            continue
        deliveries.append({
            "date":     rfnd.created_at,
            "date_str": rfnd.created_at.strftime("%d-%b-%Y") if rfnd.created_at else "—",
            "ref":      rfnd.refund_number or f"REF-{rfnd.id}",
            "kind":     "return",
            "label":    "Return / Credit",
            "items":    lines,
            "total":    round(sum(l["total"] for l in lines), 2),
        })

    deliveries.sort(
        key=lambda d: d["date"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True
    )
    for d in deliveries:
        d.pop("date", None)

    rows = []
    for entry in products.values():
        qty_net = entry["qty_received"] - entry["qty_returned"]
        value_net = entry["value_received"] - entry["value_returned"]
        rows.append({
            "product_id":     entry["product_id"],
            "name":           entry["name"],
            "sku":            entry["sku"],
            "unit":           entry["unit"],
            "qty_received":   round(entry["qty_received"], 3),
            "qty_returned":   round(entry["qty_returned"], 3),
            "qty_net":        round(qty_net, 3),
            "value_received": round(entry["value_received"], 2),
            "value_returned": round(entry["value_returned"], 2),
            "value_net":      round(value_net, 2),
            "avg_unit_price": round(value_net / qty_net, 2) if qty_net else 0.0,
            "last_received":  entry["last_received"].strftime("%d-%b-%Y") if entry["last_received"] else "—",
        })
    rows.sort(key=lambda r: r["value_net"], reverse=True)

    return {
        "products": rows,
        "deliveries": deliveries,
        "totals": {
            "product_lines":  len(rows),
            "deliveries":     sum(1 for d in deliveries if d["kind"] == "delivery"),
            "returns":        sum(1 for d in deliveries if d["kind"] == "return"),
            "qty_net":        round(sum(r["qty_net"] for r in rows), 3),
            "value_received": round(sum(r["value_received"] for r in rows), 2),
            "value_returned": round(sum(r["value_returned"] for r in rows), 2),
            "value_net":      round(sum(r["value_net"] for r in rows), 2),
        },
    }


async def _build_client_consignment_stock_payload(client_id: int, db: AsyncSession) -> dict:
    """
    What the client is still physically holding on consignment - their own
    stock of our goods, neither sold on nor sent back.

    Where "sent" comes from
    -----------------------
    A consignment invoice writes the same lines onto BOTH the invoice and the
    Consignment mirror, and the two drift in practice (imports and edits leave
    the mirror at zero). ``_build_client_products_payload`` already resolves
    that by treating the invoice as authoritative and falling back to the
    mirror only when no invoice backs it; this does the same, so "Stock on
    hand" and "Products received" cannot contradict each other on one page.

    What takes it off the shelf
    ---------------------------
    Three records retire stock, and they do not overlap:

      1. ``qty_sold`` / ``qty_returned`` on the mirror - written by Settle,
         which never touches ``amount_paid``
      2. client refunds - goods physically handed back
      3. **payment** - on a consignment deal the client pays for what they
         sell, so money against the invoice IS the record of a sale

    (3) is what makes this honest for real accounts. A payment recorded with
    its sold items (ConsignmentSale) says exactly which goods went; a payment
    recorded before that feature existed, or written off, says only how much.
    Both are money against the invoice, so both count - but the itemised part
    is deducted per product and only the remainder is spread proportionally,
    by value, over what is left. Deducting the itemised lines AND the full
    paid amount would retire the same goods twice.

    The consequence is an invariant worth knowing: with no unrecorded sales,
    the net value left on the shelf equals the client's outstanding balance,
    because a consignment client is invoiced for everything sent and pays it
    down as they sell.
    """
    cons_result = await db.execute(
        select(Consignment)
        .where(Consignment.client_id == client_id)
        .options(selectinload(Consignment.items).selectinload(ConsignmentItem.product))
        .order_by(Consignment.created_at)
    )
    consignments = cons_result.scalars().all()
    if not consignments:
        return {"items": [], "totals": _empty_consignment_stock_totals()}

    # The invoices behind those consignments: the authoritative record of what
    # went out, and of how much of it has been paid for.
    invoice_ids = {c.invoice_id for c in consignments if c.invoice_id}
    invoices: dict[int, B2BInvoice] = {}
    if invoice_ids:
        inv_result = await db.execute(
            select(B2BInvoice)
            .where(B2BInvoice.id.in_(invoice_ids))
            .options(selectinload(B2BInvoice.items).selectinload(B2BInvoiceItem.product))
        )
        invoices = {inv.id: inv for inv in inv_result.scalars().all()}

    products: dict[int, dict] = {}

    def bucket(product, product_id):
        entry = products.get(product_id)
        if entry is None:
            entry = products[product_id] = {
                "product_id":    product_id,
                "name":          getattr(product, "name", None) or f"Product #{product_id}",
                "sku":           getattr(product, "sku", None) or "",
                "unit":          getattr(product, "unit", None) or "",
                "qty_sent":      0.0,
                "qty_settled":   0.0,   # Settle flow
                "qty_returned":  0.0,   # Settle flow
                "_price_qty":    0.0,
                "_price_value":  0.0,
                "last_received": None,
            }
        return entry

    def note_sent(entry, qty, price, when):
        entry["qty_sent"] += qty
        entry["_price_qty"] += qty
        entry["_price_value"] += qty * price
        if when and (entry["last_received"] is None or when > entry["last_received"]):
            entry["last_received"] = when

    gross_invoiced = 0.0     # subtotal, before the client discount
    net_invoiced = 0.0       # total, after it
    net_paid = 0.0
    for cons in consignments:
        invoice = invoices.get(cons.invoice_id) if cons.invoice_id else None
        if invoice is not None:
            gross_invoiced += float(invoice.subtotal or 0)
            net_invoiced += float(invoice.total or 0)
            net_paid += float(invoice.amount_paid or 0)
            for it in invoice.items:
                note_sent(bucket(it.product, it.product_id),
                          float(it.qty or 0), float(it.unit_price or 0), cons.created_at)
        for ci in cons.items:
            price = float(ci.unit_price or 0)
            entry = bucket(ci.product, ci.product_id)
            if invoice is None:
                note_sent(entry, float(ci.qty_sent or 0), price, cons.created_at)
            entry["qty_settled"] += float(ci.qty_sold or 0)
            entry["qty_returned"] += float(ci.qty_returned or 0)
            if not entry["_price_qty"] and price:      # drifted mirror: keep a price
                entry["_price_qty"], entry["_price_value"] = 1.0, price

    # Payments that named the goods they were for.
    sold_result = await db.execute(
        select(ConsignmentSaleItem)
        .join(ConsignmentSale, ConsignmentSaleItem.sale_id == ConsignmentSale.id)
        .where(ConsignmentSale.client_id == client_id)
    )
    reported_sold: dict[int, float] = {}
    for line in sold_result.scalars().all():
        reported_sold[line.product_id] = reported_sold.get(line.product_id, 0.0) + float(line.qty or 0)
    itemised_paid = float((await db.execute(
        select(func.coalesce(func.sum(ConsignmentSale.amount), 0))
        .where(ConsignmentSale.client_id == client_id)
    )).scalar() or 0)

    # Goods physically handed back. Only products actually placed on
    # consignment count, so a refund against an outright sale cannot eat into
    # consignment stock.
    refund_result = await db.execute(
        select(B2BRefundItem)
        .join(B2BRefund, B2BRefundItem.refund_id == B2BRefund.id)
        .where(B2BRefund.client_id == client_id)
    )
    refunded: dict[int, float] = {}
    for line in refund_result.scalars().all():
        if line.product_id in products:
            refunded[line.product_id] = refunded.get(line.product_id, 0.0) + float(line.qty or 0)

    # Everything the records name explicitly comes off first.
    for entry in products.values():
        pid = entry["product_id"]
        entry["_named_sold"] = round(reported_sold.get(pid, 0.0), 3)
        entry["_refunded"] = round(refunded.get(pid, 0.0), 3)
        entry["_unit_price"] = (round(entry["_price_value"] / entry["_price_qty"], 2)
                                if entry["_price_qty"] else 0.0)
        entry["_remaining"] = max(0.0, round(
            entry["qty_sent"] - entry["qty_settled"] - entry["qty_returned"]
            - entry["_named_sold"] - entry["_refunded"], 3))

    # Then the money that was paid without naming anything is spread over what
    # is left, by value. Line prices are gross while payments are net of the
    # client discount, so the shelf is valued through the same ratio the
    # invoices themselves used rather than a discount read off the client -
    # that way a rate that changed over time still reconciles.
    net_factor = (net_invoiced / gross_invoiced) if gross_invoiced > 0 else 1.0
    shelf_net_value = sum(e["_remaining"] * e["_unit_price"] for e in products.values()) * net_factor
    unnamed_paid = max(0.0, round(net_paid - itemised_paid, 2))
    sold_ratio = min(1.0, unnamed_paid / shelf_net_value) if shelf_net_value > 0.005 else 0.0

    rows = []
    for entry in products.values():
        remaining = entry["_remaining"]
        on_hand = round(remaining * (1 - sold_ratio), 3)
        unit_price = entry["_unit_price"]
        rows.append({
            "product_id":    entry["product_id"],
            "name":          entry["name"],
            "sku":           entry["sku"],
            "unit":          entry["unit"],
            "qty_sent":      round(entry["qty_sent"], 3),
            "qty_sold":      round(entry["qty_settled"] + entry["_named_sold"]
                                  + (remaining - on_hand), 3),
            "qty_returned":  round(entry["qty_returned"] + entry["_refunded"], 3),
            "qty_on_hand":   on_hand,
            "unit_price":    unit_price,
            "value_on_hand": round(on_hand * unit_price, 2),
            "last_received": entry["last_received"].strftime("%d-%b-%Y") if entry["last_received"] else "—",
        })
    rows.sort(key=lambda r: (-r["value_on_hand"], r["name"].lower()))

    in_stock = [r for r in rows if r["qty_on_hand"] > 0.0005]
    return {
        "items": rows,
        "totals": {
            "product_lines": len(in_stock),
            "qty_on_hand":   round(sum(r["qty_on_hand"] for r in rows), 3),
            "value_on_hand": round(sum(r["value_on_hand"] for r in rows), 2),
            "qty_sent":      round(sum(r["qty_sent"] for r in rows), 3),
            "qty_sold":      round(sum(r["qty_sold"] for r in rows), 3),
            "qty_returned":  round(sum(r["qty_returned"] for r in rows), 3),
            # What the shelf is worth at the price the client is billed.
            "net_value_on_hand": round(
                sum(r["value_on_hand"] for r in rows) * net_factor, 2),
        },
    }


def _empty_consignment_stock_totals() -> dict:
    return {
        "product_lines": 0, "qty_on_hand": 0.0, "value_on_hand": 0.0,
        "qty_sent": 0.0, "qty_sold": 0.0, "qty_returned": 0.0,
    }


@router.get("/api/clients/{client_id}/consignment-stock")
async def client_consignment_stock_data(
    client_id: int,
    db: AsyncSession = Depends(get_async_session),
):
    return await _build_client_consignment_stock_payload(client_id, db)


@router.get("/api/clients/{client_id}/products")
async def client_products_data(
    client_id: int,
    as_of: Optional[date] = None,
    db: AsyncSession = Depends(get_async_session),
):
    return await _build_client_products_payload(client_id, db, as_of=as_of)


@router.get("/api/clients/{client_id}/statement")
async def client_statement_data(
    client_id: int,
    as_of: Optional[date] = None,
    db: AsyncSession = Depends(get_async_session),
):
    return await _build_client_statement_payload(client_id, db, as_of=as_of)


@router.get("/client/{client_id}/statement", response_class=HTMLResponse)
async def client_statement_print(
    client_id: int,
    request: Request,
    as_of: Optional[date] = None,
    db: AsyncSession = Depends(get_async_session),
):
    payload = await _build_client_statement_payload(client_id, db, as_of=as_of)
    return templates.TemplateResponse(
        request,
        "b2b_client_statement_print.html",
        {
            "client": payload["client"],
            "transactions": payload["transactions"],
            "payment_activity": payload["payment_activity"],
            "statement_date": payload["statement_date"],
            "statement_period_label": payload["statement_period_label"],
            "total_invoiced": payload["total_invoiced"],
            "total_paid": payload["total_paid"],
            "balance_due": payload["balance_due"],
        },
    )

@router.get("/", response_class=HTMLResponse)
def b2b_ui(current_user: User = Depends(require_permission("page_b2b"))):
    return """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="/static/theme-init.js"></script>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>B2B — AZed Farm</title>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{
    --bg:#060810;--surface:#0a0d18;--card:#0f1424;--card2:#151c30;
    --border:rgba(255,255,255,0.06);--border2:rgba(255,255,255,0.11);
    --green:#00ff9d;--blue:#4d9fff;--purple:#a855f7;--orange:#fb923c;
    --danger:#ff4d6d;--warn:#ffb547;--teal:#2dd4bf;
    --text:#f0f4ff;--sub:#8899bb;--muted:#445066;
    --sans:'Outfit',sans-serif;--mono:'JetBrains Mono',monospace;--r:12px;
}
body.light{
    --bg:#f4f5ef;--surface:#f1f3eb;--card:#eceee6;--card2:#e4e6de;
    --border:rgba(0,0,0,0.08);--border2:rgba(0,0,0,0.14);
    --green:#0f8a43;
    --text:#1a1e14;--sub:#4a5040;--muted:#7b816f;
}
body.light nav{background:rgba(244,245,239,.92);}
body.light .nav-link:hover{background:rgba(0,0,0,.05);}
body.light tr:hover td{background:rgba(0,0,0,.03);}
.mode-btn{display:flex;align-items:center;justify-content:center;width:36px;height:36px;border-radius:10px;border:1px solid var(--border);background:var(--card);color:var(--sub);font-size:16px;cursor:pointer;transition:all .2s;font-family:var(--sans);}
.mode-btn:hover{border-color:var(--border2);transform:scale(1.06);}
.topbar-right{display:flex;align-items:center;gap:12px;}
.account-menu{position:relative;}
.user-pill{display:flex;align-items:center;gap:10px;background:var(--card);border:1px solid var(--border);border-radius:40px;padding:7px 16px 7px 10px;cursor:pointer;transition:all .2s;}
.user-pill:hover,.user-pill.open{border-color:var(--border2);}
.user-avatar{width:28px;height:28px;background:linear-gradient(135deg,#7ecb6f,#d4a256);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#0a0c08;}
.user-name{font-size:13px;font-weight:500;color:var(--sub);}
.menu-caret{font-size:11px;color:var(--muted);}
.account-dropdown{position:absolute;right:0;top:calc(100% + 10px);min-width:220px;background:var(--card);border:1px solid var(--border2);border-radius:14px;padding:8px;box-shadow:0 24px 50px rgba(0,0,0,.35);display:none;z-index:500;}
.account-dropdown.open{display:block;}
.account-head{padding:10px 12px 8px;border-bottom:1px solid var(--border);margin-bottom:6px;}
.account-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;}
.account-email{font-size:12px;color:var(--sub);margin-top:4px;word-break:break-word;}
.account-item{width:100%;display:flex;align-items:center;gap:10px;padding:10px 12px;border:none;background:transparent;border-radius:10px;color:var(--sub);font-family:var(--sans);font-size:13px;text-decoration:none;cursor:pointer;text-align:left;}
.account-item:hover{background:var(--card2);color:var(--text);}
.account-item.danger:hover{color:#c97a7a;}
.logout-btn{background:transparent;border:1px solid var(--border);color:var(--muted);font-family:var(--sans);font-size:12px;font-weight:500;padding:8px 16px;border-radius:8px;cursor:pointer;transition:all .2s;letter-spacing:.3px;}
.logout-btn:hover{border-color:#c97a7a;color:#c97a7a;}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;}
nav{position:sticky;top:0;z-index:100;display:flex;align-items:center;gap:8px;padding:0 24px;height:58px;background:rgba(10,13,24,.92);backdrop-filter:blur(20px);border-bottom:1px solid var(--border);flex-wrap:wrap;}
.logo{font-size:17px;font-weight:900;background:linear-gradient(135deg,var(--green),var(--blue));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;margin-right:10px;text-decoration:none;display:flex;align-items:center;gap:8px;}
.nav-link{padding:7px 12px;border-radius:8px;color:var(--sub);font-size:12px;font-weight:600;text-decoration:none;transition:all .2s;white-space:nowrap;}
.nav-link:hover{background:rgba(255,255,255,.05);color:var(--text);}
.nav-link.active{background:rgba(77,159,255,.1);color:var(--blue);}
.nav-spacer{flex:1;}
.content{max-width:1300px;margin:0 auto;padding:28px 24px;display:flex;flex-direction:column;gap:20px;}
.page-title{font-size:24px;font-weight:800;letter-spacing:-.5px;}
.page-sub{color:var(--muted);font-size:13px;margin-top:3px;}
.info-banner{background:rgba(77,159,255,.07);border:1px solid rgba(77,159,255,.2);border-radius:var(--r);padding:12px 16px;font-size:13px;color:var(--blue);display:flex;align-items:center;gap:10px;}
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;}
.stat-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;display:flex;flex-direction:column;gap:8px;position:relative;overflow:hidden;}
.stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.stat-card.blue::before  {background:linear-gradient(90deg,var(--blue),transparent);}
.stat-card.warn::before  {background:linear-gradient(90deg,var(--warn),transparent);}
.stat-card.danger::before{background:linear-gradient(90deg,var(--danger),transparent);}
.stat-card.teal::before  {background:linear-gradient(90deg,var(--teal),transparent);}
.stat-label{font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);}
.stat-value{font-family:var(--mono);font-size:26px;font-weight:700;}
.stat-value.blue  {color:var(--blue);}
.stat-value.warn  {color:var(--warn);}
.stat-value.danger{color:var(--danger);}
.stat-value.teal  {color:var(--teal);}
.analysis-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;}
.analysis-layout{display:grid;grid-template-columns:1.15fr .85fr;gap:14px;}
.analysis-panel{background:var(--card);border:1px solid var(--border);border-radius:var(--r);overflow:hidden;}
.analysis-panel-head{padding:14px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;}
.analysis-title{font-size:14px;font-weight:800;color:var(--text);}
.analysis-sub{font-size:12px;color:var(--muted);margin-top:2px;}
.analysis-list{display:flex;flex-direction:column;}
.analysis-row{display:grid;grid-template-columns:minmax(130px,1fr) 96px 82px;gap:10px;align-items:center;padding:12px 16px;border-top:1px solid var(--border);}
.analysis-row:first-child{border-top:none;}
.analysis-client{min-width:0;}
.analysis-client strong{display:block;color:var(--text);font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.analysis-client span{display:block;color:var(--muted);font-size:11px;margin-top:2px;}
.bar-track{height:7px;background:var(--card2);border-radius:999px;overflow:hidden;}
.bar-fill{height:100%;border-radius:999px;background:linear-gradient(90deg,var(--blue),var(--teal));}
.risk-pill{display:inline-flex;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;}
.risk-healthy{background:rgba(0,255,157,.1);color:var(--green);}
.risk-new{background:rgba(77,159,255,.1);color:var(--blue);}
.risk-quiet{background:rgba(255,181,71,.1);color:var(--warn);}
.risk-collect,.risk-stale_outstanding{background:rgba(255,181,71,.12);color:var(--warn);}
.risk-over_limit{background:rgba(255,77,109,.12);color:var(--danger);}
@media(max-width:900px){.analysis-layout{grid-template-columns:1fr;}.analysis-row{grid-template-columns:1fr 80px;}.analysis-row .bar-track{display:none;}}
.tabs{display:flex;gap:4px;background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:4px;flex-wrap:wrap;}
.tab{padding:8px 18px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;border:none;background:transparent;color:var(--muted);transition:all .2s;font-family:var(--sans);}
.tab.active{background:var(--card2);color:var(--text);}
.btn{display:flex;align-items:center;gap:7px;padding:10px 16px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;border:none;transition:all .2s;white-space:nowrap;}
.btn-blue {background:linear-gradient(135deg,var(--blue),var(--purple));color:white;}
.btn-blue:hover{filter:brightness(1.1);transform:translateY(-1px);}
.btn-green{background:linear-gradient(135deg,var(--green),#00d4ff);color:#021a10;}
.btn-green:hover{filter:brightness(1.1);transform:translateY(-1px);}
.btn-warn {background:linear-gradient(135deg,var(--warn),var(--orange));color:#1a0800;}
.btn-warn:hover{filter:brightness(1.1);transform:translateY(-1px);}
.btn-teal {background:linear-gradient(135deg,var(--teal),var(--blue));color:#001a18;}
.btn-teal:hover{filter:brightness(1.1);transform:translateY(-1px);}
.btn-danger{background:transparent;border:1px solid var(--danger);color:var(--danger);}
.btn-danger:hover{background:var(--danger);color:#fff;}
.toolbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
.search-box{display:flex;align-items:center;gap:9px;background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:0 14px;flex:1;min-width:200px;}
.search-box input{background:transparent;border:none;outline:none;color:var(--text);font-family:var(--sans);font-size:14px;padding:11px 0;width:100%;}
.search-box input::placeholder{color:var(--muted);}
.table-wrap{background:var(--card);border:1px solid var(--border);border-radius:var(--r);overflow:hidden;}
table{width:100%;border-collapse:collapse;}
thead{background:var(--card2);}
th{text-align:left;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:12px 16px;}
td{padding:12px 16px;border-top:1px solid var(--border);color:var(--sub);font-size:13px;}
tr:hover td{background:rgba(255,255,255,.02);}
td.name{color:var(--text);font-weight:600;}
.action-btn{background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;transition:all .15s;font-family:var(--sans);}
.action-btn:hover      {border-color:var(--blue);  color:var(--blue);}
.action-btn.danger:hover{border-color:var(--danger);color:var(--danger);}
.action-btn.green:hover {border-color:var(--green); color:var(--green);}
.action-btn.warn:hover  {border-color:var(--warn);  color:var(--warn);}
.action-btn.teal:hover  {border-color:var(--teal);  color:var(--teal);}
.badge{display:inline-flex;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;}
.badge-cash        {background:rgba(0,255,157,.1); color:var(--green);}
.badge-full_payment{background:rgba(77,159,255,.1);color:var(--blue);}
.badge-consignment {background:rgba(45,212,191,.1);color:var(--teal);}
.badge-paid        {background:rgba(0,255,157,.1); color:var(--green);}
.badge-unpaid      {background:rgba(255,181,71,.1);color:var(--warn);}
.badge-partial     {background:rgba(77,159,255,.1);color:var(--blue);}
.badge-active      {background:rgba(45,212,191,.1);color:var(--teal);}
.badge-closed      {background:rgba(0,255,157,.1); color:var(--green);}
.modal-bg{position:fixed;inset:0;z-index:500;background:rgba(0,0,0,.75);backdrop-filter:blur(4px);display:none;align-items:center;justify-content:center;}
.modal-bg.open{display:flex;}
.modal{background:var(--card);border:1px solid var(--border2);border-radius:16px;padding:28px;width:680px;max-width:95vw;max-height:90vh;overflow-y:auto;animation:modalIn .2s ease;}
@keyframes modalIn{from{opacity:0;transform:scale(.95)}to{opacity:1;transform:scale(1)}}
.modal-title{font-size:18px;font-weight:800;margin-bottom:4px;}
.modal-sub{font-size:13px;color:var(--muted);margin-bottom:20px;}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.fld{display:flex;flex-direction:column;gap:6px;margin-bottom:14px;}
.fld.span2{grid-column:span 2;}
.fld label{font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);}
.fld input,.fld select,.fld textarea{background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--sans);font-size:14px;outline:none;transition:border-color .2s;width:100%;}
.fld input:focus,.fld select:focus{border-color:rgba(77,159,255,.4);}
.modal-actions{display:flex;gap:10px;margin-top:8px;justify-content:flex-end;}
.btn-cancel{background:transparent;border:1px solid var(--border2);color:var(--sub);padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;}
.btn-cancel:hover{border-color:var(--danger);color:var(--danger);}
.type-selector{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:16px;}
.type-opt{background:var(--card2);border:2px solid var(--border2);border-radius:12px;padding:14px 10px;cursor:pointer;text-align:center;transition:all .2s;}
.type-opt:hover{border-color:var(--blue);}
.type-opt.selected.cash        {border-color:var(--green);background:rgba(0,255,157,.08);}
.type-opt.selected.full_payment{border-color:var(--blue); background:rgba(77,159,255,.08);}
.type-opt.selected.consignment {border-color:var(--teal); background:rgba(45,212,191,.08);}
.type-icon{font-size:24px;margin-bottom:6px;}
.type-label{font-size:13px;font-weight:700;color:var(--text);}
.type-desc{font-size:10px;color:var(--muted);margin-top:3px;}
.type-accounting{font-size:10px;color:var(--warn);margin-top:4px;font-style:italic;}
.item-row{display:grid;grid-template-columns:2fr 80px 100px 30px;gap:8px;align-items:center;margin-bottom:8px;}
.item-row select,.item-row input{background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;width:100%;}
.item-row select:focus,.item-row input:focus{border-color:rgba(77,159,255,.4);}
.rm-btn{background:none;border:none;color:var(--muted);font-size:18px;cursor:pointer;padding:0;transition:color .15s;}
.rm-btn:hover{color:var(--danger);}
.add-item-btn{border:1px dashed rgba(77,159,255,.3);color:var(--blue);font-family:var(--sans);font-size:13px;font-weight:600;padding:8px;border-radius:8px;cursor:pointer;width:100%;transition:all .2s;margin-bottom:14px;background:transparent;}
.add-item-btn:hover{background:rgba(77,159,255,.08);}
.invoice-summary{background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:14px;margin-bottom:14px;}
.inv-row{display:flex;justify-content:space-between;font-size:13px;padding:4px 0;}
.inv-row.total{font-size:18px;font-weight:800;border-top:1px solid var(--border2);margin-top:8px;padding-top:10px;}
.side-bg{position:fixed;inset:0;z-index:400;background:rgba(0,0,0,.5);display:none;}
.side-bg.open{display:block;}
.side-panel{position:fixed;right:0;top:0;bottom:0;width:500px;max-width:95vw;background:var(--card);border-left:1px solid var(--border2);display:flex;flex-direction:column;transform:translateX(100%);transition:transform .3s ease;z-index:401;}
.side-panel.open{transform:translateX(0);}
.side-header{padding:20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
.side-header h3{font-size:16px;font-weight:800;}
.close-btn{background:none;border:none;color:var(--muted);font-size:22px;cursor:pointer;padding:0;}
.close-btn:hover{color:var(--danger);}
.side-body{flex:1;overflow-y:auto;padding:16px 20px;}
.cons-item-card{background:var(--card2);border:1px solid var(--border);border-radius:10px;padding:14px;margin-bottom:10px;}
.cons-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px;}
.cons-input{background:var(--card);border:1px solid var(--border2);border-radius:7px;padding:7px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100%;}
.cons-input:focus{border-color:rgba(45,212,191,.4);}
.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(16px);background:var(--card2);border:1px solid var(--border2);border-radius:var(--r);padding:12px 20px;font-size:13px;font-weight:600;color:var(--text);box-shadow:0 20px 50px rgba(0,0,0,.5);opacity:0;pointer-events:none;transition:opacity .25s,transform .25s;z-index:999;}
.toast.show{opacity:1;transform:translateX(-50%) translateY(0);}
::-webkit-scrollbar{width:4px;}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px;}
</style>
    <script src="/static/auth-guard.js"></script>
</head>
<body>
""" + render_app_header(current_user, "page_b2b") + """

<div class="content">
    <div>
        <div class="page-title">B2B Sales</div>
        <div class="page-sub">Business clients — cash, full payment and consignment deals</div>
    </div>

    <div class="info-banner">
        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        <span><b>Accounting:</b> Cash invoices post directly to Revenue. Full Payment &amp; Consignment go to <b>Deferred Revenue</b> — revenue is only recognized when payment is collected or consignment is settled.</span>
    </div>

    <div class="stats-grid">
        <div class="stat-card blue"><div class="stat-label">B2B Clients</div><div class="stat-value blue" id="stat-clients">—</div></div>
        <div class="stat-card warn"><div class="stat-label">Outstanding</div><div class="stat-value warn" id="stat-outstanding">—</div></div>
        <div class="stat-card danger"><div class="stat-label">Unpaid Invoices</div><div class="stat-value danger" id="stat-unpaid">—</div></div>
        <div class="stat-card teal"><div class="stat-label">Active Consignments</div><div class="stat-value teal" id="stat-consign">—</div></div>
    </div>

    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">
        <div class="tabs">
            <button class="tab active" id="tab-clients"    onclick="switchTab('clients')">Clients</button>
            <button class="tab"        id="tab-invoices"  onclick="switchTab('invoices')">Invoices</button>
            <button class="tab"        id="tab-analysis"  onclick="switchTab('analysis')">Client Analysis</button>
            <button class="tab"        id="tab-refunds"   onclick="switchTab('refunds')">Client Refund</button>
            <button class="tab"        id="tab-pricelists" onclick="switchTab('pricelists')">&#127991; Price Lists</button>
        </div>
        <div style="display:flex;gap:10px;">
            <button class="btn btn-blue"  id="btn-add-client"  onclick="openClientModal()">+ Add Client</button>
            <button class="btn btn-green" id="btn-new-invoice" onclick="openInvoiceModal()" style="display:none">+ New Invoice</button>
        </div>
    </div>

    <!-- CLIENTS -->
    <div id="section-clients">
        <div class="toolbar">
            <div class="search-box">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                <input id="client-search" placeholder="Search clients..." oninput="onClientSearch()">
            </div>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Business</th><th>Contact</th><th>Phone</th><th>Default Terms</th><th>Discount %</th><th>Outstanding</th><th>Actions</th></tr></thead>
                <tbody id="clients-body"><tr><td colspan="7" style="text-align:center;color:var(--muted);padding:40px">Loading...</td></tr></tbody>
            </table>
        </div>
    </div>

    <!-- CLIENT ANALYSIS -->
    <div id="section-analysis" style="display:none">
        <div class="analysis-grid" style="margin-bottom:14px">
            <div class="stat-card blue"><div class="stat-label">Net B2B Sales</div><div class="stat-value blue" id="ana-net-sales">—</div></div>
            <div class="stat-card warn"><div class="stat-label">Outstanding</div><div class="stat-value warn" id="ana-outstanding">—</div></div>
            <div class="stat-card teal"><div class="stat-label">Payment Rate</div><div class="stat-value teal" id="ana-payment-rate">—</div></div>
            <div class="stat-card danger"><div class="stat-label">At-Risk Clients</div><div class="stat-value danger" id="ana-risk-count">—</div></div>
        </div>

        <div class="analysis-layout" style="margin-bottom:14px">
            <div class="analysis-panel">
                <div class="analysis-panel-head">
                    <div><div class="analysis-title">Top Clients</div><div class="analysis-sub">Ranked by net sales after refunds</div></div>
                </div>
                <div class="analysis-list" id="analysis-top-clients"></div>
            </div>
            <div class="analysis-panel">
                <div class="analysis-panel-head">
                    <div><div class="analysis-title">Collection Watch</div><div class="analysis-sub">Largest unpaid balances</div></div>
                </div>
                <div class="analysis-list" id="analysis-collection-watch"></div>
            </div>
        </div>

        <div class="table-wrap">
            <div style="padding:16px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
                <div>
                    <div class="modal-title" style="margin-bottom:2px">Client Performance</div>
                    <div class="modal-sub">Sales, collections, outstanding balance, credit usage, and recency by client.</div>
                </div>
                <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
                    <div class="search-box" style="min-width:220px;flex:0 1 260px">
                        <input id="analysis-search" placeholder="Search clients..." oninput="renderClientAnalysis()">
                    </div>
                    <select id="analysis-sort" onchange="renderClientAnalysis()" style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:9px 12px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none">
                        <option value="net_sales">Net sales</option>
                        <option value="outstanding">Outstanding</option>
                        <option value="invoice_count">Invoice count</option>
                        <option value="payment_rate">Payment rate</option>
                        <option value="last_invoice">Last invoice</option>
                    </select>
                </div>
            </div>
            <div style="overflow-x:auto">
                <table>
                    <thead><tr><th>Client</th><th>Terms</th><th>Invoices</th><th>Gross</th><th>Refunds</th><th>Net Sales</th><th>Paid</th><th>Outstanding</th><th>Payment Rate</th><th>Avg Invoice</th><th>Credit Used</th><th>Last Invoice</th><th>Status</th></tr></thead>
                    <tbody id="analysis-body"><tr><td colspan="13" style="text-align:center;color:var(--muted);padding:40px">Open this tab to load client analysis.</td></tr></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- INVOICES -->
    <div id="section-invoices" style="display:none">
        <div class="toolbar">
            <div class="search-box">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                <input id="invoice-search" placeholder="Search by client, invoice number..." oninput="filterInvoices()">
            </div>
            <select class="filter-sel" id="type-filter" onchange="filterInvoices()" style="background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:10px 14px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;">
                <option value="">All Types</option>
                <option value="cash">💵 Cash</option>
                <option value="full_payment">📋 Full Payment</option>
                <option value="consignment">🔄 Consignment</option>
            </select>
            <select class="filter-sel" id="status-filter" onchange="filterInvoices()" style="background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:10px 14px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;">
                <option value="">All Statuses</option>
                <option value="paid">Paid</option>
                <option value="unpaid">Unpaid</option>
                <option value="partial">Partial</option>
                <option value="consignment">Consignment</option>
            </select>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Invoice #</th><th>Client</th><th>Type</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th><th>Date</th><th>Actions</th></tr></thead>
                <tbody id="invoices-body"><tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">Loading...</td></tr></tbody>
            </table>
        </div>
    </div>

    <!-- REFUNDS -->
    <div id="section-refunds" style="display:none">
        <div class="table-wrap" style="padding:18px">
            <div class="modal-title" style="margin-bottom:4px">Client Refund</div>
            <div class="modal-sub" style="margin-bottom:16px">Select a client, add returned products, and the total will be calculated automatically.</div>

            <div class="form-row">
                <div class="fld">
                    <label>Client *</label>
                    <select id="refund-client" onchange="onRefundClientChange()"></select>
                </div>
                <div class="fld">
                    <label>Current Outstanding</label>
                    <input id="refund-outstanding" readonly value="0.00">
                </div>
            </div>

            <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);margin-bottom:8px">Returned Products</div>
            <div style="display:grid;grid-template-columns:2fr 80px 100px 30px;gap:8px;margin-bottom:6px;">
                <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Product</span>
                <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Qty</span>
                <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Unit Price</span>
                <span></span>
            </div>
            <div id="refund-items"></div>
            <button class="add-item-btn" onclick="addRefundItem()">+ Add Product</button>

            <div class="invoice-summary">
                <div class="inv-row"><span style="color:var(--muted)">Subtotal</span><span style="font-family:var(--mono)" id="refund-subtotal">0.00</span></div>
                <div class="inv-row"><span style="color:var(--muted)">Discount (<span id="refund-pct">0</span>%)</span><span style="font-family:var(--mono);color:var(--danger)" id="refund-discount">-0.00</span></div>
                <div class="inv-row total"><span>Refund Total</span><span style="font-family:var(--mono);color:var(--warn)" id="refund-total">0.00</span></div>
                <div class="inv-row"><span style="color:var(--muted)">Outstanding After Refund</span><span style="font-family:var(--mono);color:var(--green)" id="refund-after">0.00</span></div>
            </div>

            <div class="fld"><label>Notes</label><input id="refund-notes" placeholder="Optional return notes"></div>
            <div class="modal-actions" style="padding:0;margin-top:8px">
                <button class="btn-cancel" onclick="resetRefundForm()">Reset</button>
                <button class="btn btn-warn" onclick="saveRefund()">Record Refund</button>
            </div>
        </div>

        <div class="table-wrap" style="margin-top:18px">
            <div style="padding:16px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
                <div>
                    <div class="modal-title" style="margin-bottom:2px">Refund Records</div>
                    <div class="modal-sub">Recent client refunds with discount, notes, and print access.</div>
                </div>
                <button class="btn btn-outline" onclick="loadRefundRecords()">Refresh Records</button>
            </div>
            <table>
                <thead><tr><th>Refund #</th><th>Client</th><th>Subtotal</th><th>Discount</th><th>Total</th><th>Date</th><th>Actions</th></tr></thead>
                <tbody id="refund-records-body"><tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px">Loading refunds...</td></tr></tbody>
            </table>
        </div>
    </div>

    <!-- CONSIGNMENT SETTLE (inline cards, no separate tab) -->
    <div id="section-consignments" style="display:none"></div>

    <!-- PRICE LISTS -->
    <div id="section-pricelists" style="display:none">
        <div class="table-wrap">
            <div style="padding:16px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
                <div>
                    <div class="modal-title" style="margin-bottom:2px">Client Price Lists</div>
                    <div class="modal-sub">Set custom prices per client. These override the default product price on new invoices.</div>
                </div>
                <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
                    <select id="pl-client-select" onchange="loadPriceList()" style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 12px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;min-width:200px">
                        <option value="">— Select a client —</option>
                    </select>
                    <button class="btn btn-blue" onclick="openAddPriceModal()" id="btn-add-price" style="display:none">+ Add / Edit Price</button>
                </div>
            </div>
            <table>
                <thead><tr><th>Product</th><th>SKU</th><th>Default Price</th><th>Client Price</th><th>Difference</th><th>Actions</th></tr></thead>
                <tbody id="pl-body"><tr><td colspan="6" style="text-align:center;color:var(--muted);padding:40px">Select a client to view their price list.</td></tr></tbody>
            </table>
        </div>
    </div>
</div>

<!-- REVERSE PAYMENT MODAL -->
<div class="modal-bg" id="reverse-modal">
    <div class="modal" style="width:480px">
        <div class="modal-title">Reverse Payment</div>
        <div class="modal-sub" id="reverse-modal-sub"></div>
        <div style="background:rgba(255,181,71,.08);border:1px solid rgba(255,181,71,.2);border-radius:10px;padding:11px 14px;margin:14px 0;font-size:12px;color:var(--warn);line-height:1.55">
            A contra journal entry is posted — the original payment stays on the ledger and the
            reversal sits beside it, so the accounts still reconcile.
        </div>
        <div class="fld" style="margin-bottom:12px">
            <label>How much</label>
            <label style="display:flex;align-items:center;gap:8px;font-size:13px;color:var(--sub);font-weight:400;cursor:pointer;margin-top:6px">
                <input type="radio" name="reverse-mode" value="full" checked onchange="onReverseModeChange()"> Reverse the full payment
            </label>
            <label style="display:flex;align-items:center;gap:8px;font-size:13px;color:var(--sub);font-weight:400;cursor:pointer;margin-top:4px">
                <input type="radio" name="reverse-mode" value="partial" onchange="onReverseModeChange()"> Reverse part of it
            </label>
        </div>
        <div class="fld" id="reverse-amount-wrap" style="margin-bottom:12px">
            <label>Amount to reverse (ج.م.)</label>
            <input id="reverse-amount" type="number" min="0.01" step="any"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:9px 12px;color:var(--text);font-family:var(--mono);font-size:14px;outline:none;width:100%">
        </div>
        <div class="fld">
            <label>Reason (optional)</label>
            <input id="reverse-reason" placeholder="e.g. Cheque bounced, entered against the wrong invoice"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:9px 12px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;width:100%">
        </div>
        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeReverseModal()">Cancel</button>
            <button class="btn btn-warn" onclick="submitReversePayment()">Reverse Payment</button>
        </div>
    </div>
</div>

<!-- CLIENT PORTAL LINK MODAL -->
<div class="modal-bg" id="portal-modal">
    <div class="modal" style="width:600px">
        <div class="modal-title">Client Account Link</div>
        <div class="modal-sub" id="portal-modal-sub">Send this to the client so they can see their statement and received products, live.</div>

        <div id="portal-body" style="margin-top:16px"></div>

        <div class="modal-actions">
            <button class="btn-cancel" onclick="document.getElementById('portal-modal').classList.remove('open')">Close</button>
        </div>
    </div>
</div>

<!-- PRICE LIST MODAL -->
<div class="modal-bg" id="pl-modal">
    <div class="modal" style="width:460px">
        <div class="modal-title">Set Custom Price</div>
        <div class="modal-sub" id="pl-modal-sub">Override the default product price for this client.</div>
        <div class="fld" style="margin-top:14px">
            <label>Product *</label>
            <select id="pl-product" onchange="onPlProductChange()" style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:9px 12px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;width:100%"></select>
        </div>
        <div style="background:var(--card2);border:1px solid var(--border);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--muted);margin-bottom:12px" id="pl-default-hint"></div>
        <div class="fld">
            <label>Custom Price (ج.م.) *</label>
            <input id="pl-price" type="number" min="0" step="any" placeholder="0.00"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:9px 12px;color:var(--text);font-family:var(--mono);font-size:14px;outline:none;width:100%">
        </div>
        <div class="modal-actions">
            <button class="btn-cancel" onclick="document.getElementById('pl-modal').classList.remove('open')">Cancel</button>
            <button class="btn btn-blue" onclick="savePriceEntry()">Save Price</button>
        </div>
    </div>
</div>

<!-- CLIENT MODAL -->
<div class="modal-bg" id="client-modal">
    <div class="modal">
        <div class="modal-title" id="client-modal-title">Add B2B Client</div>
        <div class="modal-sub">Cafes, restaurants, retail stores, distributors</div>
        <div class="form-row">
            <div class="fld span2"><label>Business Name *</label><input id="c-name" placeholder="e.g. Green Cafe"></div>
            <div class="fld"><label>Contact Person</label><input id="c-contact" placeholder="Name"></div>
            <div class="fld"><label>Phone</label><input id="c-phone" placeholder="+20 100 000 0000"></div>
            <div class="fld"><label>Email</label><input id="c-email" placeholder="contact@business.com"></div>
            <div class="fld"><label>Address</label><input id="c-address" placeholder="City / Area"></div>
            <div class="fld"><label>Default Payment Terms</label>
                <select id="c-terms">
                    <option value="cash">Cash — Pay on delivery</option>
                    <option value="full_payment">Full Payment — Invoice then pay</option>
                    <option value="consignment">Consignment — Pay what you sell</option>
                </select>
            </div>
            <div class="fld"><label>Default Discount %</label>
                <input id="c-discount" type="number" placeholder="0" min="0" max="100" step="0.5" value="0">
            </div>
            <div class="fld span2"><label>Notes</label><input id="c-notes" placeholder="Internal notes"></div>
        </div>
        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeClientModal()">Cancel</button>
            <button class="btn btn-blue" onclick="saveClient()">Save Client</button>
        </div>
    </div>
</div>

<!-- INVOICE MODAL -->
<div class="modal-bg" id="invoice-modal">
    <div class="modal">
        <div class="modal-title" id="inv-modal-title">New B2B Invoice</div>
        <div class="modal-sub">Select client and products. Deal type and discount come from the client profile.</div>
        <div class="fld"><label>Client *</label>
            <select id="inv-client" onchange="onClientChange()"></select>
        </div>
        <div class="form-row" style="margin-bottom:16px">
            <div class="fld">
                <label>Deal Type</label>
                <input id="inv-deal-type" readonly>
            </div>
            <div class="fld">
                <label>Discount %</label>
                <input id="inv-discount-pct" type="number" readonly value="0">
            </div>
        </div>
        <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);margin-bottom:8px">Products</div>
        <div style="display:grid;grid-template-columns:2fr 80px 100px 30px;gap:8px;margin-bottom:6px;">
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Product</span>
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Qty</span>
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Unit Price</span>
            <span></span>
        </div>
        <div id="inv-items"></div>
        <button class="add-item-btn" onclick="addInvItem()">+ Add Product</button>
        <div class="invoice-summary">
            <div class="inv-row"><span style="color:var(--muted)">Subtotal</span><span style="font-family:var(--mono)" id="s-subtotal">0.00</span></div>
            <div class="inv-row"><span style="color:var(--muted)">Discount (<span id="s-pct">0</span>%)</span><span style="font-family:var(--mono);color:var(--danger)" id="s-discount">-0.00</span></div>
            <div class="inv-row total"><span>Total</span><span style="font-family:var(--mono);color:var(--green)" id="s-total">0.00</span></div>
        </div>
        <div class="fld"><label>Notes</label><input id="inv-notes" placeholder="Optional notes"></div>
        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeInvoiceModal()">Cancel</button>
            <button class="btn btn-green" id="inv-save-btn" onclick="saveInvoice()">Create Invoice</button>
        </div>
    </div>
</div>

<!-- PAYMENT MODAL REMOVED — payment collection happens in Accounting → B2B Clients -->

<!-- CONSIGNMENT SETTLE PANEL -->
<div class="side-bg" id="side-bg" onclick="closeSide()"></div>
<div class="side-panel" id="side-panel">
    <div class="side-header">
        <h3 id="side-title">Settle Consignment</h3>
        <button class="close-btn" onclick="closeSide()">×</button>
    </div>
    <div class="side-body" id="side-body"></div>
</div>

<!-- CONSIGNMENT PAYMENT MODAL REMOVED — consignment payments are recorded in Accounting → B2B Clients on the client account, not against a specific invoice -->

<div class="toast" id="toast"></div>

<script>
  function setModeButton(isLight){
    const btn = document.getElementById("mode-btn");
    if(btn) btn.innerText = isLight ? "☀️" : "🌙";
}
function toggleMode(){
    const isLight = document.body.classList.toggle("light");
    localStorage.setItem("colorMode", isLight ? "light" : "dark");
    setModeButton(isLight);
}
function initializeColorMode(){
    const isLight = localStorage.getItem("colorMode") === "light";
    document.body.classList.toggle("light", isLight);
    setModeButton(isLight);
}
async function initUser() {
    try {
        const r = await fetch("/auth/me");
        if (!r.ok) { _redirectToLogin(); return; }
        const u = await r.json();
        const nameEl = document.getElementById("user-name");
        const avatarEl = document.getElementById("user-avatar");
        const emailEl = document.getElementById("user-email");
        if (nameEl) nameEl.innerText = u.name;
        if (avatarEl) avatarEl.innerText = u.name.charAt(0).toUpperCase();
        if (emailEl) emailEl.innerText = u.email;
        return u;
    } catch(e) { _redirectToLogin(); }
}
function toggleAccountMenu(event){
    event.stopPropagation();
    const trigger = document.getElementById("account-trigger");
    const dropdown = document.getElementById("account-dropdown");
    const open = dropdown.classList.toggle("open");
    trigger.classList.toggle("open", open);
    trigger.setAttribute("aria-expanded", open ? "true" : "false");
}
document.addEventListener("click", e => {
    const menu = document.getElementById("account-dropdown");
    const trigger = document.getElementById("account-trigger");
    if(!menu || !trigger) return;
    if(menu.contains(e.target) || trigger.contains(e.target)) return;
    menu.classList.remove("open");
    trigger.classList.remove("open");
    trigger.setAttribute("aria-expanded", "false");
});
async function logout(){
    await fetch("/auth/logout", { method: "POST" });
    window.location.href = "/";
}
  let currentUser = null;
  function hasPermission(permission, u = currentUser){
      const role = u ? (u.role || "") : "";
      const perms = new Set(u ? (u.permissions || []) : []);
      return role === "admin" || perms.has(permission);
  }
  function configureB2BPermissions(u){
      currentUser = u;
      isAdmin = u.role === "admin";
      if(!hasPermission("tab_b2b_clients", u)){
          let el = document.getElementById("tab-clients");
          if(el) el.style.display = "none";
          let analysisEl = document.getElementById("tab-analysis");
          if(analysisEl) analysisEl.style.display = "none";
      }
    if(!hasPermission("tab_b2b_invoices", u)){
          let el = document.getElementById("tab-invoices");
          if(el) el.style.display = "none";
          let refundEl = document.getElementById("tab-refunds");
          if(refundEl) refundEl.style.display = "none";
      }
      if(!hasPermission("tab_b2b_clients", u) && hasPermission("tab_b2b_invoices", u)){
          setTimeout(() => switchTab("invoices"), 0);
      }
      renderRefundRecords(allRefunds || []);
  }
  initializeColorMode();
  initUser().then(u => { if(u) configureB2BPermissions(u); });
let allProducts   = [];
let refundProducts = [];
let allClients    = [];
let allInvoices   = [];
let allRefunds    = [];
let clientAnalysis = null;
let selectedType  = "cash";
let editingClientId  = null;
let editingInvoiceId = null;
let settlingConsId   = null;
let searchTimer      = null;
let isAdmin = false; // set by initUser() via configureB2BPermissions(u)

async function init(){
    // Run seeding and data loading independently
    fetch("/b2b/api/seed-accounts", {method:"POST"}).catch(e => console.warn("Seeding failed", e));
    
    try {
        const res = await fetch("/b2b/api/products-list");
        if(res.ok) {
            allProducts = await res.json();
            buildB2BProductDatalist();
        }
    } catch(e) { console.error("Products load failed", e); }

    loadStats().catch(e => console.error("Stats load failed", e));
    loadClients().catch(e => {
        console.error("Clients load failed", e);
        document.getElementById("clients-body").innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--danger);padding:40px">Error loading clients. Check permissions or database.</td></tr>`;
    });
}

async function loadStats(){
    let d = await (await fetch("/b2b/api/stats")).json();
    document.getElementById("stat-clients").innerText     = d.total_clients;
    document.getElementById("stat-outstanding").innerText = d.total_outstanding.toFixed(2);
    document.getElementById("stat-unpaid").innerText      = d.unpaid_invoices;
    document.getElementById("stat-consign").innerText     = d.active_consign;
}

/* ── TABS ── */
function switchTab(tab){
    const required = {
        clients: "tab_b2b_clients",
        analysis: "tab_b2b_clients",
        invoices: "tab_b2b_invoices",
        refunds: "tab_b2b_invoices",
        consignments: "tab_b2b_consignment",
    };
    if(required[tab] && !hasPermission(required[tab])) return;
    ["clients","analysis","invoices","refunds","consignments","pricelists"].forEach(t=>{
        let el = document.getElementById("section-"+t);
        if(el) el.style.display = t===tab?"":"none";
        let tb = document.getElementById("tab-"+t);
        if(tb) tb.classList.toggle("active", t===tab);
    });
    document.getElementById("btn-add-client").style.display  = tab==="clients"    ?"":"none";
    document.getElementById("btn-new-invoice").style.display = tab==="invoices"   ?"":"none";
    if(tab==="analysis")   loadClientAnalysis();
    if(tab==="invoices")   loadInvoices();
    if(tab==="refunds")    prepareRefundTab();
    if(tab==="pricelists") initPriceListTab();
}

/* ── CLIENT ANALYSIS ── */
function escHtml(value){
    return String(value ?? "").replace(/[&<>"']/g, ch => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[ch]));
}

function fmtMoney(value){
    return (Number(value) || 0).toFixed(2);
}

function riskLabel(risk){
    return {
        healthy: "Healthy",
        new: "New",
        quiet: "Quiet",
        collect: "Collect",
        stale_outstanding: "Stale",
        over_limit: "Over limit",
    }[risk] || "Review";
}

function termsLabel(type){
    return {
        cash: "Cash",
        full_payment: "Full Payment",
        consignment: "Consignment",
    }[type] || type || "—";
}

async function loadClientAnalysis(){
    const body = document.getElementById("analysis-body");
    if(body) body.innerHTML = `<tr><td colspan="13" style="text-align:center;color:var(--muted);padding:40px">Loading client analysis...</td></tr>`;
    try {
        const res = await fetch("/b2b/api/client-analysis");
        if(!res.ok) throw new Error(`API Error: ${res.status}`);
        clientAnalysis = await res.json();
        renderClientAnalysis();
    } catch(err) {
        console.error(err);
        if(body) body.innerHTML = `<tr><td colspan="13" style="text-align:center;color:var(--danger);padding:40px">Error loading client analysis.</td></tr>`;
    }
}

function renderAnalysisList(targetId, rows, valueKey, emptyText){
    const el = document.getElementById(targetId);
    if(!el) return;
    if(!rows || !rows.length){
        el.innerHTML = `<div style="padding:18px;color:var(--muted);font-size:13px">${emptyText}</div>`;
        return;
    }
    const maxValue = Math.max(...rows.map(row => Number(row[valueKey]) || 0), 1);
    el.innerHTML = rows.map(row => {
        const value = Number(row[valueKey]) || 0;
        const width = Math.max(4, Math.round((value / maxValue) * 100));
        return `<div class="analysis-row">
            <div class="analysis-client">
                <strong>${escHtml(row.name)}</strong>
                <span>${escHtml(termsLabel(row.payment_terms))} · ${row.invoice_count} invoices</span>
            </div>
            <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
            <div style="font-family:var(--mono);font-weight:700;color:${valueKey==="outstanding"?"var(--warn)":"var(--blue)"};text-align:right">${fmtMoney(value)}</div>
        </div>`;
    }).join("");
}

function renderClientAnalysis(){
    if(!clientAnalysis) return;
    const summary = clientAnalysis.summary || {};
    document.getElementById("ana-net-sales").innerText = fmtMoney(summary.net_sales);
    document.getElementById("ana-outstanding").innerText = fmtMoney(summary.outstanding);
    document.getElementById("ana-payment-rate").innerText = `${(Number(summary.payment_rate) || 0).toFixed(1)}%`;
    document.getElementById("ana-risk-count").innerText = summary.at_risk_clients || 0;

    renderAnalysisList("analysis-top-clients", clientAnalysis.top_clients || [], "net_sales", "No B2B sales yet.");
    renderAnalysisList("analysis-collection-watch", clientAnalysis.collection_watch || [], "outstanding", "No outstanding balances.");

    const q = (document.getElementById("analysis-search")?.value || "").trim().toLowerCase();
    const sort = document.getElementById("analysis-sort")?.value || "net_sales";
    let rows = [...(clientAnalysis.clients || [])].filter(row => {
        if(!q) return true;
        return [row.name, row.contact_person, row.phone, row.payment_terms].some(value => String(value || "").toLowerCase().includes(q));
    });
    rows.sort((a,b) => {
        if(sort === "last_invoice"){
            return (a.days_since_last_invoice ?? 999999) - (b.days_since_last_invoice ?? 999999);
        }
        return (Number(b[sort]) || 0) - (Number(a[sort]) || 0);
    });

    const body = document.getElementById("analysis-body");
    if(!rows.length){
        body.innerHTML = `<tr><td colspan="13" style="text-align:center;color:var(--muted);padding:40px">No clients match this analysis filter.</td></tr>`;
        return;
    }
    body.innerHTML = rows.map(row => {
        const creditUsed = row.credit_used_pct === null || row.credit_used_pct === undefined ? "—" : `${Number(row.credit_used_pct).toFixed(1)}%`;
        const last = row.days_since_last_invoice === null || row.days_since_last_invoice === undefined
            ? "—"
            : `${escHtml(row.last_invoice)}<br><span style="font-size:11px;color:var(--muted)">${row.days_since_last_invoice} days ago</span>`;
        return `<tr>
            <td class="name">${escHtml(row.name)}<br><span style="font-size:11px;color:var(--muted)">${escHtml(row.contact_person)}</span></td>
            <td><span class="badge badge-${escHtml(row.payment_terms)}">${escHtml(termsLabel(row.payment_terms))}</span></td>
            <td style="font-family:var(--mono)">${row.invoice_count}</td>
            <td style="font-family:var(--mono)">${fmtMoney(row.gross_sales)}</td>
            <td style="font-family:var(--mono);color:${row.refunds>0?"var(--danger)":"var(--muted)"}">${row.refunds>0?fmtMoney(row.refunds):"—"}</td>
            <td style="font-family:var(--mono);font-weight:700;color:var(--blue)">${fmtMoney(row.net_sales)}</td>
            <td style="font-family:var(--mono);color:var(--green)">${fmtMoney(row.paid_amount)}</td>
            <td style="font-family:var(--mono);color:${row.outstanding>0?"var(--warn)":"var(--muted)"}">${row.outstanding>0?fmtMoney(row.outstanding):"—"}</td>
            <td style="font-family:var(--mono);color:${row.payment_rate>=80?"var(--green)":row.payment_rate>=50?"var(--warn)":"var(--danger)"}">${Number(row.payment_rate).toFixed(1)}%</td>
            <td style="font-family:var(--mono)">${fmtMoney(row.average_invoice)}</td>
            <td style="font-family:var(--mono);color:${row.credit_used_pct>100?"var(--danger)":"var(--muted)"}">${creditUsed}</td>
            <td style="font-size:12px;color:var(--sub)">${last}</td>
            <td><span class="risk-pill risk-${escHtml(row.risk_level)}">${escHtml(riskLabel(row.risk_level))}</span></td>
        </tr>`;
    }).join("");
}

/* ── CLIENTS ── */
function onClientSearch(){ clearTimeout(searchTimer); searchTimer=setTimeout(loadClients,300); }

async function loadClients(){
    try {
        let q = document.getElementById("client-search").value.trim();
        const res = await fetch(`/b2b/api/clients${q?"?q="+encodeURIComponent(q):""}`);
        if (!res.ok) throw new Error(`API Error: ${res.status}`);
        allClients = await res.json();
        
        if(!allClients.length){
            document.getElementById("clients-body").innerHTML=`<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:40px">No clients yet.</td></tr>`;
            return;
        }
        const termsLabel={cash:"💵 Cash",full_payment:"📋 Full Payment",consignment:"🔄 Consignment"};
        document.getElementById("clients-body").innerHTML = allClients.map(c=>`
            <tr>
                <td class="name">${c.name}</td>
                <td style="font-size:12px">${c.contact_person}</td>
            <td style="font-family:var(--mono);font-size:12px">${c.phone}</td>
            <td><span class="badge badge-${c.payment_terms}">${termsLabel[c.payment_terms]||c.payment_terms}</span></td>
            <td style="font-family:var(--mono);color:var(--blue)">${c.discount_pct>0?c.discount_pct.toFixed(1)+"%":"—"}</td>
            <td style="font-family:var(--mono);color:${c.outstanding>0?"var(--warn)":"var(--muted)"}">
                ${c.outstanding>0?c.outstanding.toFixed(2):"—"}
            </td>
            <td style="display:flex;gap:6px;flex-wrap:wrap">
                ${hasPermission("tab_b2b_invoices")?`<button class="action-btn green" onclick="quickInvoice(${c.id})">+ Invoice</button>`:""}
                <button class="action-btn" onclick="window.open('/b2b/client/${c.id}/statement','_blank')" title="Account Statement">&#128196; Statement</button>
                ${hasPermission("action_b2b_clients_update")?`<button class="action-btn${c.portal_enabled?" green":""}" onclick="openPortalLink(${c.id})" title="Shareable live account link for this client">&#128279; ${c.portal_enabled?"Link&nbsp;on":"Share&nbsp;link"}</button>`:""}
                <button class="action-btn" onclick="openEditClient(${c.id})">Edit</button>
                ${hasPermission("action_b2b_delete")?`<button class="action-btn danger" onclick="deleteClient(${c.id},'${c.name.replace(/'/g,"\\'")}')">Remove</button>`:""}
            </td>
        </tr>`).join("");
    } catch (err) {
        console.error(err);
        document.getElementById("clients-body").innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--danger);padding:40px">Error loading clients.</td></tr>`;
    }
}

function openClientModal(){
    editingClientId = null;
    document.getElementById("client-modal-title").innerText = "Add B2B Client";
    ["c-name","c-contact","c-phone","c-email","c-address","c-notes"].forEach(id=>document.getElementById(id).value="");
    document.getElementById("c-terms").value    = "cash";
    document.getElementById("c-discount").value = "0";
    document.getElementById("client-modal").classList.add("open");
}

function openEditClient(id){
    let c = allClients.find(x=>x.id===id); if(!c) return;
    editingClientId = id;
    document.getElementById("client-modal-title").innerText = "Edit Client";
    document.getElementById("c-name").value    = c.name;
    document.getElementById("c-contact").value = c.contact_person==="—"?"":c.contact_person;
    document.getElementById("c-phone").value   = c.phone==="—"?"":c.phone;
    document.getElementById("c-email").value   = c.email==="—"?"":c.email;
    document.getElementById("c-address").value = c.address==="—"?"":c.address;
    document.getElementById("c-terms").value   = c.payment_terms;
    document.getElementById("c-discount").value= c.discount_pct;
    document.getElementById("c-notes").value   = c.notes==="—"?"":c.notes;
    document.getElementById("client-modal").classList.add("open");
}

function closeClientModal(){ document.getElementById("client-modal").classList.remove("open"); }

async function saveClient(){
    let name = document.getElementById("c-name").value.trim();
    if(!name){ showToast("Business name is required"); return; }
    let body = {
        name,
        contact_person: document.getElementById("c-contact").value.trim()||null,
        phone:          document.getElementById("c-phone").value.trim()||null,
        email:          document.getElementById("c-email").value.trim()||null,
        address:        document.getElementById("c-address").value.trim()||null,
        payment_terms:  document.getElementById("c-terms").value,
        discount_pct:   parseFloat(document.getElementById("c-discount").value)||0,
        notes:          document.getElementById("c-notes").value.trim()||null,
    };
    let url=editingClientId?`/b2b/api/clients/${editingClientId}`:"/b2b/api/clients";
    let method=editingClientId?"PUT":"POST";
    let res=await fetch(url,{method,headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});
    let data=await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    closeClientModal();
    showToast(editingClientId?"Client updated ✓":"Client added ✓");
    loadClients(); loadStats();
}

async function deleteClient(id,name){
    if(!confirm(`Remove "${name}"?`)) return;
    await fetch(`/b2b/api/clients/${id}`,{method:"DELETE"});
    showToast("Client removed ✓");
    loadClients(); loadStats();
}

/* ── CLIENT PORTAL LINK ── */
let portalClientId = null;

function portalEscape(v){
    return String(v == null ? "" : v).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}

async function openPortalLink(clientId){
    portalClientId = clientId;
    const client = allClients.find(c => c.id === clientId);
    document.getElementById("portal-modal-sub").innerText =
        `${client ? client.name : "Client"} — they see their own statement and received products, live. No login needed.`;
    document.getElementById("portal-body").innerHTML =
        `<div style="color:var(--muted);font-size:13px;padding:18px 0">Loading link…</div>`;
    document.getElementById("portal-modal").classList.add("open");
    await renderPortalState();
}

async function renderPortalState(){
    let state;
    try{
        const res = await fetch(`/b2b/api/clients/${portalClientId}/portal`, { credentials: "same-origin" });
        state = await res.json();
        if(!res.ok) throw new Error(state.detail || "Could not load the link");
    } catch(e){
        document.getElementById("portal-body").innerHTML =
            `<div style="color:var(--danger);font-size:13px;padding:18px 0">${portalEscape(e.message)}</div>`;
        return;
    }

    if(!state.enabled){
        document.getElementById("portal-body").innerHTML = `
            <div style="background:var(--card2);border:1px solid var(--border);border-radius:10px;padding:16px 18px;font-size:13px;color:var(--sub);line-height:1.6">
                No link has been issued for this client yet.<br>
                Creating one generates a private web address. <strong>Anyone who has that address can see this
                client's statement and deliveries</strong>, so send it only to them. You can revoke it at any time.
            </div>
            <button class="btn btn-blue" style="margin-top:14px" onclick="createPortalLink(false)">Create link</button>`;
        return;
    }

    const url = state.url;
    const seen = state.last_viewed_at
        ? `Opened ${state.view_count} time${state.view_count === 1 ? "" : "s"} · last on ${portalEscape(state.last_viewed_at)}`
        : "Not opened yet";
    document.getElementById("portal-body").innerHTML = `
        <div class="fld"><label>Client's private link</label>
            <input id="portal-url" readonly value="${portalEscape(url)}"
                onclick="this.select()"
                style="width:100%;background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:10px 12px;color:var(--text);font-family:var(--mono);font-size:12.5px;outline:none">
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:10px">
            <button class="btn btn-blue" onclick="copyPortalLink()">Copy link</button>
            <button class="btn" style="background:#25D366;color:#07260f" onclick="sharePortalWhatsApp()">Send on WhatsApp</button>
            <button class="btn-cancel" onclick="window.open(document.getElementById('portal-url').value,'_blank')">Preview</button>
        </div>
        <div style="font-size:12px;color:var(--muted);margin-top:14px;line-height:1.7">
            ${portalEscape(seen)}${state.created_at ? ` · issued ${portalEscape(state.created_at)}` : ""}<br>
            The page refreshes itself, so the client always sees current numbers.
        </div>
        <div style="border-top:1px solid var(--border);margin-top:16px;padding-top:14px;display:flex;gap:8px;flex-wrap:wrap">
            <button class="btn-cancel" onclick="createPortalLink(true)" title="Issue a new address and kill the old one">Replace link</button>
            <button class="btn btn-danger" onclick="revokePortalLink()" title="Stop the link working">Revoke access</button>
        </div>`;
}

async function createPortalLink(rotate){
    if(rotate && !confirm("Replace the link?\\n\\nThe address you already sent this client will stop working immediately.")) return;
    const res = await fetch(`/b2b/api/clients/${portalClientId}/portal?rotate=${rotate ? "true" : "false"}`,
        { method: "POST", credentials: "same-origin" });
    const data = await res.json();
    if(!res.ok){ showToast("Error: " + (data.detail || "Could not create the link")); return; }
    showToast(rotate ? "New link issued — old one is dead" : "Link created ✓");
    await renderPortalState();
    loadClients();
}

async function revokePortalLink(){
    if(!confirm("Revoke access?\\n\\nThe client's link stops working immediately.")) return;
    const res = await fetch(`/b2b/api/clients/${portalClientId}/portal`,
        { method: "DELETE", credentials: "same-origin" });
    const data = await res.json().catch(() => ({}));
    if(!res.ok){ showToast("Error: " + (data.detail || "Could not revoke the link")); return; }
    showToast("Link revoked ✓");
    await renderPortalState();
    loadClients();
}

async function copyPortalLink(){
    const input = document.getElementById("portal-url");
    try{
        await navigator.clipboard.writeText(input.value);
    } catch(e){
        // clipboard API needs HTTPS or permission — fall back to selecting it
        input.select();
        document.execCommand("copy");
    }
    showToast("Link copied ✓");
}

function sharePortalWhatsApp(){
    const client = allClients.find(c => c.id === portalClientId);
    const url = document.getElementById("portal-url").value;
    const msg = `Hello${client ? " " + client.name : ""}, here is your live account with Habiba Organic Farm — statement and products received: ${url}`;
    window.open("https://wa.me/?text=" + encodeURIComponent(msg), "_blank");
}

/* ── INVOICE MODAL ── */
async function openInvoiceModal(preClientId=null){
    editingInvoiceId = null;
    document.getElementById("inv-modal-title").innerText = "New B2B Invoice";
    document.getElementById("inv-save-btn").innerText    = "Create Invoice";
    let sel = document.getElementById("inv-client");
    sel.innerHTML = allClients.map(c=>
        `<option value="${c.id}" data-terms="${c.payment_terms}" data-discount="${c.discount_pct}" ${c.id===preClientId?"selected":""}>${c.name}</option>`
    ).join("");
    document.getElementById("inv-items").innerHTML = "";
    document.getElementById("inv-notes").value = "";
    // Await so allProducts is loaded with client prices before items are added
    await onClientChange();
    addInvItem();
    document.getElementById("invoice-modal").classList.add("open");
}

function quickInvoice(clientId){ switchTab("invoices"); setTimeout(()=>openInvoiceModal(clientId),50); }

function closeInvoiceModal(){
    editingInvoiceId = null;
    document.getElementById("inv-modal-title").innerText = "New B2B Invoice";
    document.getElementById("inv-save-btn").innerText    = "Create Invoice";
    document.getElementById("invoice-modal").classList.remove("open");
}

async function onClientChange(){
    let sel = document.getElementById("inv-client");
    let opt = sel.options[sel.selectedIndex];
    if(!opt || !opt.value) return;
    let terms    = opt.dataset.terms    || "cash";
    let discount = parseFloat(opt.dataset.discount) || 0;
    selectType(terms);
    document.getElementById("inv-deal-type").value = formatDealType(terms);
    document.getElementById("inv-discount-pct").value = discount;
    // Reload product list with client-specific prices
    let clientId = parseInt(opt.value);
    allProducts = await (await fetch(`/b2b/api/products-list?client_id=${clientId}`)).json();
    buildB2BProductDatalist();
    updateSummary();
}

function formatDealType(type){
    const labels = {
        cash: "Cash",
        full_payment: "Full Payment",
        consignment: "Consignment",
    };
    return labels[type] || type;
}

function selectType(type){
    selectedType = type;
}

function buildB2BProductDatalist(){
    let dl = document.getElementById("b2b-product-datalist");
    if(!dl){
        dl = document.createElement("datalist");
        dl.id = "b2b-product-datalist";
        document.body.appendChild(dl);
    }
    dl.innerHTML = allProducts.map(p=>
        `<option data-id="${p.id}" value="${p.sku} — ${p.name}" data-price="${p.price}" data-unit="${p.unit}" data-stock="${p.stock}">`
    ).join("");
}

function resolveB2BProduct(inputEl){
    let val = inputEl.value.trim().toLowerCase();
    let match = allProducts.find(p=>
        (p.sku+" — "+p.name).toLowerCase()===val ||
        p.sku.toLowerCase()===val ||
        p.name.toLowerCase()===val
    );
    if(!match) match = allProducts.find(p=>
        p.sku.toLowerCase().startsWith(val) ||
        p.name.toLowerCase().includes(val)
    );
    return match||null;
}

function productLabel(product){
    return `${product.sku} — ${product.name}`;
}

function isServiceProduct(product){
    return product && (product.stock_tracked === false || product.item_type === "service");
}

function stockShortLabel(product){
    return isServiceProduct(product) ? "service" : `stk ${product.stock.toFixed(0)} ${product.unit}`;
}

function stockHintLabel(product){
    return isServiceProduct(product) ? "service" : `stock: ${product.stock.toFixed(0)} ${product.unit}`;
}

function productMatches(products, query){
    let q = (query || "").trim().toLowerCase();
    if(!q) return products.slice(0, 8);
    let starts = [];
    let contains = [];
    products.forEach(p=>{
        let sku  = (p.sku || "").toLowerCase();
        let name = (p.name || "").toLowerCase();
        if(sku.startsWith(q) || name.startsWith(q)) starts.push(p);
        else if(sku.includes(q) || name.includes(q)) contains.push(p);
    });
    return starts.concat(contains).slice(0, 8);
}

// getProducts can be an array OR a function returning an array
function attachProductDropdown(inputEl, hiddenEl, hintEl, getProducts, accent, onPick){
    function resolveList(){ return typeof getProducts === "function" ? getProducts() : getProducts; }

    let box = document.createElement("div");
    box.style.cssText = "position:absolute;left:0;right:0;top:calc(100% + 4px);background:var(--card);border:1px solid var(--border2);border-radius:10px;box-shadow:0 18px 40px rgba(0,0,0,.4);max-height:280px;overflow-y:auto;z-index:9999;display:none;";
    inputEl.parentElement.appendChild(box);

    let activeIdx = -1;

    function hideBox(){ box.style.display = "none"; activeIdx = -1; }

    function draw(q){
        let items = productMatches(resolveList(), q);
        activeIdx = -1;
        if(!items.length){
            box.innerHTML = `<div style="padding:10px 14px;color:var(--muted);font-size:12px">No matching products</div>`;
            box.style.display = "block";
            return;
        }
        box.innerHTML = items.map((p, i)=>{
            let priceColor = (p.has_custom) ? "var(--blue)" : accent;
            let customTag  = p.has_custom ? `<span style="font-size:9px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;background:rgba(77,159,255,.15);color:var(--blue);padding:1px 5px;border-radius:4px;margin-left:4px">custom</span>` : "";
            return `<button type="button" data-idx="${i}" style="width:100%;text-align:left;background:transparent;border:none;padding:10px 14px;cursor:pointer;${i>0?"border-top:1px solid var(--border)":""};font-family:var(--sans);transition:background .1s;">
                <div style="display:flex;justify-content:space-between;gap:10px;align-items:center">
                    <div style="min-width:0;flex:1">
                        <div style="font-size:13px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${p.name}${customTag}</div>
                        <div style="font-family:var(--mono);font-size:11px;color:var(--muted);margin-top:1px">${p.sku || "—"}</div>
                    </div>
                    <div style="text-align:right;flex-shrink:0">
                        <div style="font-family:var(--mono);font-size:13px;font-weight:700;color:${priceColor}">${p.price.toFixed(2)}</div>
                        <div style="font-size:10px;color:var(--muted);margin-top:1px">${stockShortLabel(p)}</div>
                    </div>
                </div>
            </button>`;
        }).join("");
        box.querySelectorAll("button[data-idx]").forEach(btn=>{
            btn.addEventListener("mousedown", function(e){
                e.preventDefault();
                let p = items[parseInt(this.dataset.idx)];
                pick(p);
            });
            btn.addEventListener("mouseenter", function(){
                setActive(parseInt(this.dataset.idx));
            });
        });
        box.style.display = "block";
    }

    function setActive(idx){
        activeIdx = idx;
        box.querySelectorAll("button[data-idx]").forEach(btn=>{
            let active = parseInt(btn.dataset.idx) === idx;
            btn.style.background = active ? "var(--card2)" : "transparent";
        });
    }

    function pick(p){
        inputEl.value = productLabel(p);
        hiddenEl.value = p.id;
        hintEl.innerText = stockHintLabel(p);
        inputEl.style.borderColor = accent;
        onPick(p);
        hideBox();
    }

    inputEl.addEventListener("focus", function(){ draw(this.value); });
    inputEl.addEventListener("input", function(){ draw(this.value); });
    inputEl.addEventListener("keydown", function(e){
        let btns = Array.from(box.querySelectorAll("button[data-idx]"));
        if(!btns.length) return;
        if(e.key === "ArrowDown"){
            e.preventDefault();
            setActive(Math.min(activeIdx + 1, btns.length - 1));
        } else if(e.key === "ArrowUp"){
            e.preventDefault();
            setActive(Math.max(activeIdx - 1, 0));
        } else if(e.key === "Enter" && activeIdx >= 0){
            e.preventDefault();
            btns[activeIdx].dispatchEvent(new MouseEvent("mousedown", {bubbles:true}));
        } else if(e.key === "Escape"){
            hideBox();
        }
    });
    inputEl.addEventListener("blur", function(){ setTimeout(hideBox, 150); });
}

function addInvItem(selectedId=null, qty=1, price=null){
    let div = document.createElement("div");
    div.className = "item-row";
    div.innerHTML = `
        <div style="position:relative;flex:1;">
            <input type="text"
                class="b2b-prod-search"
                placeholder="Search by name or SKU…"
                autocomplete="off"
                style="width:100%;background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;transition:border-color .2s;">
            <input type="hidden" class="b2b-prod-id">
            <span class="b2b-stock-hint" style="position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:10px;color:var(--muted);pointer-events:none;"></span>
        </div>
        <input type="number" placeholder="1" min="0.001" step="any" value="${qty}" oninput="updateSummary()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:80px;">
        <input type="number" placeholder="0.00" min="0" step="any" value="${price!=null?price:""}" oninput="updateSummary()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100px;">
        <button class="rm-btn" onclick="this.closest('.item-row').remove();updateSummary()">×</button>
    `;
    let searchInp = div.querySelector(".b2b-prod-search");
    let hiddenId  = div.querySelector(".b2b-prod-id");
    let stockHint = div.querySelector(".b2b-stock-hint");
    let priceInp  = div.querySelectorAll("input[type=number]")[1];

    // Pre-fill when editing an existing invoice
    if(selectedId){
        let p = allProducts.find(x=>x.id===selectedId);
        if(p){
            searchInp.value = productLabel(p);
            hiddenId.value  = p.id;
            stockHint.innerText = stockHintLabel(p);
            searchInp.style.borderColor = "rgba(0,255,157,.4)";
        }
    }

    // Rich dropdown — always reads current allProducts via getter
    attachProductDropdown(
        searchInp, hiddenId, stockHint,
        () => allProducts,
        "rgba(0,255,157,.45)",
        function(p){
            priceInp.value = p.price.toFixed(2);
            if(p.has_custom){
                priceInp.style.borderColor = "rgba(77,159,255,.6)";
                priceInp.title = `Custom price (default: ${(p.default_price||p.price).toFixed(2)})`;
            } else {
                priceInp.style.borderColor = "";
                priceInp.title = "";
            }
            updateSummary();
        }
    );

    document.getElementById("inv-items").appendChild(div);
    if(price != null) updateSummary();
}

function updateSummary(){
    let rows = document.querySelectorAll("#inv-items .item-row");
    let subtotal = 0;
    rows.forEach(row=>{
        let qty   = parseFloat(row.querySelectorAll("input[type=number]")[0].value)||0;
        let price = parseFloat(row.querySelectorAll("input[type=number]")[1].value)||0;
        subtotal += qty * price;
    });
    let pct=parseFloat(document.getElementById("inv-discount-pct").value)||0;
    let discount=subtotal*pct/100;
    let total=subtotal-discount;
    document.getElementById("s-subtotal").innerText = subtotal.toFixed(2);
    document.getElementById("s-pct").innerText      = pct.toFixed(1);
    document.getElementById("s-discount").innerText = "-"+discount.toFixed(2);
    document.getElementById("s-total").innerText    = total.toFixed(2);
}

/* ── REFUNDS ── */
async function prepareRefundTab(){
    let refundClients = await (await fetch("/b2b/api/clients")).json();
    let sel = document.getElementById("refund-client");
    if(!sel) return;
    sel.innerHTML = refundClients.map(c=>
        `<option value="${c.id}" data-outstanding="${c.outstanding}">${c.name}</option>`
    ).join("");
    if(!refundClients.length){
        document.getElementById("refund-outstanding").value = "0.00";
        document.getElementById("refund-pct").innerText = "0";
        document.getElementById("refund-subtotal").innerText = "0.00";
        document.getElementById("refund-discount").innerText = "-0.00";
        document.getElementById("refund-total").innerText = "0.00";
        document.getElementById("refund-after").innerText = "0.00";
        document.getElementById("refund-items").innerHTML = "";
        return;
    }
    await loadRefundProducts(parseInt(sel.value));
    if(!document.getElementById("refund-items").children.length){
        addRefundItem();
    }
    await onRefundClientChange();
    loadRefundRecords();
}

async function onRefundClientChange(){
    let sel = document.getElementById("refund-client");
    let opt = sel.options[sel.selectedIndex];
    let outstanding = opt ? (parseFloat(opt.dataset.outstanding) || 0) : 0;
    let client = allClients.find(c => c.id === parseInt(opt?.value || "0"));
    let discountPct = client ? (parseFloat(client.discount_pct) || 0) : 0;
    document.getElementById("refund-outstanding").value = outstanding.toFixed(2);
    document.getElementById("refund-pct").innerText = discountPct.toFixed(1);
    if(opt && opt.value){
        await loadRefundProducts(parseInt(opt.value));
    }
    updateRefundSummary();
    loadRefundRecords();
}

async function loadRefundProducts(clientId){
    let res = await fetch(`/b2b/api/refund-products/${clientId}`);
    let data = await res.json();
    refundProducts = Array.isArray(data) ? data : [];
}

function addRefundItem(selectedId=null, qty=1, price=null){
    let div = document.createElement("div");
    div.className = "item-row";
    div.innerHTML = `
        <div style="position:relative;flex:1;">
            <input type="text"
                class="b2b-prod-search"
                placeholder="Search by name or SKU…"
                autocomplete="off"
                style="width:100%;background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;transition:border-color .2s;">
            <input type="hidden" class="b2b-prod-id">
            <span class="b2b-stock-hint" style="position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:10px;color:var(--muted);pointer-events:none;"></span>
        </div>
        <input type="number" placeholder="1" min="0.001" step="any" value="${qty}" oninput="updateRefundSummary()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:80px;">
        <input type="number" placeholder="0.00" min="0" step="any" value="${price!=null?price:""}" oninput="updateRefundSummary()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100px;">
        <button class="rm-btn" onclick="this.closest('.item-row').remove();updateRefundSummary()">×</button>
    `;
    let searchInp = div.querySelector(".b2b-prod-search");
    let hiddenId  = div.querySelector(".b2b-prod-id");
    let stockHint = div.querySelector(".b2b-stock-hint");

    if(selectedId){
        let p = allProducts.find(x=>x.id===selectedId);
        if(p){
            searchInp.value = productLabel(p);
            hiddenId.value  = p.id;
            stockHint.innerText = stockHintLabel(p);
            searchInp.style.borderColor = "rgba(255,181,71,.45)";
        }
    }

    let priceInp = div.querySelectorAll("input[type=number]")[1];
    attachProductDropdown(
        searchInp,
        hiddenId,
        stockHint,
        () => refundProducts,
        "rgba(255,181,71,.45)",
        function(p){
            priceInp.value = p.price.toFixed(2);
            updateRefundSummary();
        }
    );

    document.getElementById("refund-items").appendChild(div);
    if(price != null) updateRefundSummary();
}

function updateRefundSummary(){
    let rows = document.querySelectorAll("#refund-items .item-row");
    let subtotal = 0;
    rows.forEach(row=>{
        let qty   = parseFloat(row.querySelectorAll("input[type=number]")[0].value)||0;
        let price = parseFloat(row.querySelectorAll("input[type=number]")[1].value)||0;
        subtotal += qty * price;
    });
    let pct = parseFloat(document.getElementById("refund-pct").innerText)||0;
    let discount = subtotal * pct / 100;
    let total = Math.max(0, subtotal - discount);
    let outstanding = parseFloat(document.getElementById("refund-outstanding").value)||0;
    let after = Math.max(0, outstanding - total);
    document.getElementById("refund-subtotal").innerText = subtotal.toFixed(2);
    document.getElementById("refund-discount").innerText = "-" + discount.toFixed(2);
    document.getElementById("refund-total").innerText = total.toFixed(2);
    document.getElementById("refund-after").innerText = after.toFixed(2);
}

function resetRefundForm(){
    document.getElementById("refund-notes").value = "";
    document.getElementById("refund-items").innerHTML = "";
    addRefundItem();
    onRefundClientChange();
}

async function saveRefund(){
    let client_id = parseInt(document.getElementById("refund-client").value);
    if(!client_id){ showToast("Select a client"); return; }
    let rows = document.querySelectorAll("#refund-items .item-row");
    let items = [];
    for(let row of rows){
        let product_id = parseInt(row.querySelector(".b2b-prod-id").value)||0;
        let qty        = parseFloat(row.querySelectorAll("input[type=number]")[0].value)||0;
        let unit_price = parseFloat(row.querySelectorAll("input[type=number]")[1].value)||0;
        if(!product_id){ showToast("Select a product for all rows"); return; }
        if(qty<=0){ showToast("Refund quantity must be greater than 0"); return; }
        items.push({product_id, qty, unit_price});
    }
    if(!items.length){ showToast("Add at least one returned product"); return; }
    let res = await fetch("/b2b/api/refunds", {
        method: "POST",
        headers: {"Content-Type":"application/json"},
        body: JSON.stringify({
            client_id,
            notes: document.getElementById("refund-notes").value.trim() || null,
            items,
        }),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: " + data.detail); return; }
    showToast(`${data.refund_number} recorded for ${data.client} - ${data.amount.toFixed(2)} EGP`);
    resetRefundForm();
    await loadClients();
    await loadStats();
    await prepareRefundTab();
}

async function loadRefundRecords(){
    let sel = document.getElementById("refund-client");
    let clientId = parseInt(sel?.value || "0");
    let url = `/b2b/api/refunds${clientId ? "?client_id="+clientId : ""}`;
    allRefunds = await (await fetch(url)).json();
    renderRefundRecords(allRefunds);
}

function renderRefundRecords(refunds){
    let body = document.getElementById("refund-records-body");
    if(!body) return;
    if(!refunds.length){
        body.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px">No refund records yet.</td></tr>`;
        return;
    }
    body.innerHTML = refunds.map(r=>`
        <tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--warn)">${r.refund_number}</td>
            <td class="name">${r.client}</td>
            <td style="font-family:var(--mono)">${r.subtotal.toFixed(2)}</td>
            <td style="font-family:var(--mono);color:${r.discount>0?"var(--danger)":"var(--muted)"}">${r.discount>0?`${r.discount.toFixed(2)} (${r.discount_pct.toFixed(1)}%)`:"—"}</td>
            <td style="font-family:var(--mono);font-weight:700;color:var(--warn)">${r.total.toFixed(2)}</td>
            <td style="font-size:12px;color:var(--muted)">${r.created_at}</td>
            <td><div style="display:flex;gap:6px;flex-wrap:wrap">
                <button class="action-btn" onclick="window.open('/b2b/refund/${r.id}/print','_blank')">Print</button>
                ${isAdmin ? `<button class="action-btn danger js-delete-refund" data-refund-id="${r.id}" data-refund-number="${String(r.refund_number).replace(/"/g, "&quot;")}">Delete</button>` : ""}
            </div></td>
        </tr>
    `).join("");
}

document.addEventListener("click", function(event){
    const btn = event.target.closest(".js-delete-refund");
    if(!btn) return;
    deleteRefund(parseInt(btn.dataset.refundId || "0"), btn.dataset.refundNumber || "");
});

async function deleteRefund(id, refundNumber){
    if(!confirm(`Are you sure you want to delete refund ${refundNumber}?`)) return;
    let res = await fetch(`/b2b/api/refunds/${id}`, {method:"DELETE"});
    let data = await res.json().catch(()=>({detail:"Unable to delete refund"}));
    if(!res.ok || data.detail){
        showToast("Error: " + (data.detail || "Unable to delete refund"));
        return;
    }
    showToast(`${refundNumber} deleted`);
    await loadClients();
    await loadStats();
    await prepareRefundTab();
}

async function openEditInvoice(id){
    let data=await (await fetch("/b2b/api/invoices?limit=500")).json();
    let invoice=data.invoices.find(i=>i.id===id);
    if(!invoice){ showToast("Could not load invoice"); return; }
    editingInvoiceId = id;
    document.getElementById("inv-modal-title").innerText = `Edit Invoice — ${invoice.invoice_number}`;
    document.getElementById("inv-save-btn").innerText    = "Save Changes";
    let sel = document.getElementById("inv-client");
    sel.innerHTML = allClients.map(c=>
        `<option value="${c.id}" data-terms="${c.payment_terms}" data-discount="${c.discount_pct}" ${c.id===invoice.client_id?"selected":""}>${c.name}</option>`
    ).join("");
    await onClientChange();
    document.getElementById("inv-notes").value        = invoice.notes;
    document.getElementById("inv-items").innerHTML = "";
    invoice.items.forEach(item=>{ addInvItem(item.product_id, item.qty, item.unit_price); });
    updateSummary();
    document.getElementById("invoice-modal").classList.add("open");
}

async function saveInvoice(){
    let client_id = parseInt(document.getElementById("inv-client").value);
    if(!client_id){ showToast("Select a client"); return; }
    let rows = document.querySelectorAll("#inv-items .item-row");
    let items = [];
    for(let row of rows){
        let product_id = parseInt(row.querySelector(".b2b-prod-id").value)||0;
        let qty        = parseFloat(row.querySelectorAll("input[type=number]")[0].value)||0;
        let unit_price = parseFloat(row.querySelectorAll("input[type=number]")[1].value)||0;
        if(!product_id){ showToast("Select a product for all rows"); return; }
        if(qty<=0)      { showToast("Quantity must be greater than 0"); return; }
        items.push({product_id, qty, unit_price});
    }
    if(!items.length){ showToast("Add at least one product"); return; }
    let body={
        client_id,
        invoice_type:selectedType,
        payment_method:selectedType,
        discount_pct:parseFloat(document.getElementById("inv-discount-pct").value)||0,
        notes:document.getElementById("inv-notes").value.trim()||null,
        items,
    };
    let url=editingInvoiceId?`/b2b/api/invoices/${editingInvoiceId}`:"/b2b/api/invoices";
    let method=editingInvoiceId?"PUT":"POST";
    let res=await fetch(url,{method,headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});
    let data=await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    closeInvoiceModal();
    let action=editingInvoiceId?"updated":"created";
    showToast(`${data.invoice_number} ${action} ✓  Total: ${data.total.toFixed(2)} EGP`);
    loadInvoices(); loadClients(); loadStats();
}

async function deleteInvoice(id,number){
    if(!confirm(`Delete invoice ${number}? This will reverse all stock and accounting changes.`)) return;
    let res=await fetch(`/b2b/api/invoices/${id}`,{method:"DELETE"});
    let data=await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    showToast(`${number} deleted — reversed ✓`);
    loadInvoices(); loadClients(); loadStats();
}

/* ── INVOICES TABLE ── */
async function loadInvoices(){
    try {
        const res = await fetch("/b2b/api/invoices?limit=200");
        if (!res.ok) throw new Error(`API Error: ${res.status}`);
        let data = await res.json();
        allInvoices = data.invoices;
        renderInvoices(allInvoices);
    } catch (err) {
        document.getElementById("invoices-body").innerHTML = `<tr><td colspan="9" style="text-align:center;color:var(--danger);padding:40px">Error loading invoices.</td></tr>`;
    }
}

function filterInvoices(){
    let q      = document.getElementById("invoice-search").value.toLowerCase();
    let type   = document.getElementById("type-filter").value;
    let status = document.getElementById("status-filter").value;
    renderInvoices(allInvoices.filter(i=>{
        let matchQ = !q || i.client.toLowerCase().includes(q) || i.invoice_number.toLowerCase().includes(q);
        let matchT = !type   || i.invoice_type === type;
        let matchS = !status || i.status === status;
        return matchQ && matchT && matchS;
    }));
}

function renderInvoices(invoices){
    if(!invoices.length){
        document.getElementById("invoices-body").innerHTML=`<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">No invoices yet.</td></tr>`;
        return;
    }
    const typeLabel={cash:"💵 Cash",full_payment:"📋 Full Payment",consignment:"🔄 Consignment"};
    document.getElementById("invoices-body").innerHTML = invoices.map(i=>{
        // Payment collection is recorded in Accounting → B2B Clients, not from the B2B invoices list.
        let editable = !(i.status === "paid" && i.amount_paid > 0);
        let actionBtns=`<div style="display:flex;gap:5px;flex-wrap:wrap">
            <button class="action-btn" onclick="window.open('/b2b/invoice/${i.id}/print','_blank')">🖨 Print</button>
            ${(editable && (isAdmin || hasPermission("action_b2b_invoices_update")))?`<button class="action-btn teal" onclick="openEditInvoice(${i.id})">✏ Edit</button>`:""}
            ${(i.amount_paid > 0 && (isAdmin || hasPermission("action_b2b_collect")))?`<button class="action-btn warn" onclick="openReversePayment(${i.id},'${i.invoice_number.replace(/'/g,"\'")}',${i.amount_paid})" title="Undo a payment collected on this invoice">↩ Reverse Payment</button>`:""}
            ${(isAdmin || hasPermission("action_b2b_delete"))?`<button class="action-btn danger" onclick="deleteInvoice(${i.id},'${i.invoice_number}')">Delete</button>`:""}
        </div>`;
        return `<tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--blue)">${i.invoice_number}</td>
            <td class="name">${i.client}</td>
            <td><span class="badge badge-${i.invoice_type}">${typeLabel[i.invoice_type]||i.invoice_type}</span></td>
            <td style="font-family:var(--mono);font-weight:700">${i.total.toFixed(2)}</td>
            <td style="font-family:var(--mono);color:var(--green)">${i.amount_paid.toFixed(2)}</td>
            <td style="font-family:var(--mono);color:${i.balance_due>0?"var(--warn)":"var(--muted)"}">${i.balance_due>0?i.balance_due.toFixed(2):"—"}</td>
            <td><span class="badge badge-${i.status}">${i.status}</span></td>
            <td style="font-size:12px;color:var(--muted)">${i.created_at}</td>
            <td>${actionBtns}</td>
        </tr>`;
    }).join("");
}

/* ── PAYMENT ── Removed: payment collection happens in Accounting → B2B Clients */

/* ── CONSIGNMENT PAYMENT ── Removed: consignment payments are recorded on the client account in Accounting → B2B Clients */

/* ── REVERSE PAYMENT ── */
let reverseInvoiceId = null;
let reverseInvoicePaid = 0;

function closeReverseModal(){
    const el = document.getElementById("reverse-modal");
    if(el) el.classList.remove("open");
}

function openReversePayment(invoiceId, invoiceNumber, amountPaid){
    reverseInvoiceId = invoiceId;
    reverseInvoicePaid = Number(amountPaid || 0);
    document.getElementById("reverse-modal-sub").innerText =
        `${invoiceNumber} — ${reverseInvoicePaid.toFixed(2)} EGP has been collected on this invoice.`;
    const amt = document.getElementById("reverse-amount");
    amt.value = reverseInvoicePaid.toFixed(2);
    amt.max = reverseInvoicePaid;
    document.getElementById("reverse-reason").value = "";
    document.querySelector('input[name="reverse-mode"][value="full"]').checked = true;
    onReverseModeChange();
    document.getElementById("reverse-modal").classList.add("open");
}

function onReverseModeChange(){
    const full = document.querySelector('input[name="reverse-mode"]:checked').value === "full";
    const amt = document.getElementById("reverse-amount");
    amt.disabled = full;
    if(full) amt.value = reverseInvoicePaid.toFixed(2);
    document.getElementById("reverse-amount-wrap").style.opacity = full ? ".55" : "1";
}

async function submitReversePayment(){
    const full = document.querySelector('input[name="reverse-mode"]:checked').value === "full";
    const amount = full ? null : (parseFloat(document.getElementById("reverse-amount").value) || 0);
    const reason = document.getElementById("reverse-reason").value.trim();
    if(!full && (amount <= 0 || amount > reverseInvoicePaid + 0.01)){
        showToast(`Enter an amount between 0 and ${reverseInvoicePaid.toFixed(2)}`);
        return;
    }
    const label = full ? `all ${reverseInvoicePaid.toFixed(2)} EGP` : `${amount.toFixed(2)} EGP`;
    if(!confirm(`Reverse ${label} of payment?

A contra journal entry will be posted. The original payment stays on the ledger.`)) return;
    const res = await fetch(`/b2b/api/invoices/${reverseInvoiceId}/reverse-payment`, {
        method: "POST", headers: {"Content-Type":"application/json"},
        credentials: "same-origin",
        body: JSON.stringify({amount, reason: reason || null}),
    });
    const data = await res.json().catch(()=>({}));
    if(!res.ok || data.detail){ showToast("Error: " + (data.detail || "Could not reverse the payment")); return; }
    closeReverseModal();
    showToast(`Reversed ${Number(data.reversed).toFixed(2)} EGP — ${data.invoice_number} is now ${data.status}`);
    loadInvoices();
    loadClients();
}

/* ── CONSIGNMENTS ── */
async function loadConsignments(){
    let conses=await (await fetch("/b2b/api/consignments")).json();
    if(!conses.length){
        document.getElementById("cons-body").innerHTML=`<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">No consignments yet.</td></tr>`;
        return;
    }
    document.getElementById("cons-body").innerHTML = conses.map(c=>`
        <tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--teal)">${c.ref_number}</td>
            <td class="name">${c.client}</td>
            <td style="font-family:var(--mono)">${c.total_sent.toFixed(0)}</td>
            <td style="font-family:var(--mono);color:var(--green)">${c.total_sold.toFixed(0)}</td>
            <td style="font-family:var(--mono);color:var(--blue)">${c.total_returned.toFixed(0)}</td>
            <td style="font-family:var(--mono);color:var(--warn);font-weight:700">${c.total_revenue.toFixed(2)}</td>
            <td><span class="badge badge-${c.status}">${c.status}</span></td>
            <td style="font-size:12px;color:var(--muted)">${c.created_at}</td>
            <td>${c.status!=="closed"?`<button class="action-btn teal" onclick="openSettle(${c.id})">Settle</button>`:"✓"}</td>
        </tr>`).join("");
}

async function openSettleByInvoice(invoiceId){
    let conses = await (await fetch("/b2b/api/consignments")).json();
    let cons   = conses.find(c => c.items.length > 0 || c.ref_number);
    // Find the consignment linked to this invoice
    let allC   = await (await fetch("/b2b/api/consignments")).json();
    let found  = allC.find(c => {
        // The consignment is linked via invoice_id on backend, load all and match by client+date proximity
        return true; // we'll filter properly below
    });
    // Fetch full list and find by invoice association
    let res  = await fetch("/b2b/api/consignments");
    let data = await res.json();
    // Find the consignment for this invoice — match via the invoice items
    let invoice = allInvoices.find(i => i.id === invoiceId);
    if(!invoice){ showToast("Invoice not found"); return; }
    // Get consignments for this client and find active one matching invoice total
    let match = data.find(c =>
        c.client_id === invoice.client_id &&
        c.status !== "closed" &&
        Math.abs(c.items.reduce((s,ci)=>s+ci.qty_sent*ci.unit_price,0) - invoice.total) < 0.01
    );
    if(!match){
        // fallback: just get latest active consignment for this client
        match = data.find(c => c.client_id === invoice.client_id && c.status !== "closed");
    }
    if(!match){ showToast("No active consignment found for this invoice"); return; }
    openSettle(match.id);
}

async function openSettle(id){
    settlingConsId=id;
    let conses=await (await fetch("/b2b/api/consignments")).json();
    let cons=conses.find(c=>c.id===id);
    if(!cons) return;
    document.getElementById("side-title").innerText=`Settle — ${cons.ref_number} (${cons.client})`;
    document.getElementById("side-body").innerHTML=`
        <p style="color:var(--muted);font-size:13px;margin-bottom:6px">Enter qty sold and returned.</p>
        <div style="background:rgba(0,255,157,.06);border:1px solid rgba(0,255,157,.15);border-radius:8px;padding:10px 12px;margin-bottom:16px;font-size:12px;color:var(--green);">
            Revenue will be recognized only for qty sold — moving from Deferred Revenue to Sales Revenue.
        </div>
        ${cons.items.map(item=>`
            <div class="cons-item-card" data-item-id="${item.id}">
                <div style="font-weight:700;margin-bottom:4px">${item.product}</div>
                <div style="font-size:12px;color:var(--muted);margin-bottom:10px">
                    Sent: ${item.qty_sent.toFixed(0)} &nbsp;|&nbsp;
                    Sold so far: ${item.qty_sold.toFixed(0)} &nbsp;|&nbsp;
                    Pending: <b style="color:var(--warn)">${item.qty_pending.toFixed(0)}</b> &nbsp;|&nbsp;
                    Price: ${item.unit_price.toFixed(2)} EGP
                </div>
                <div class="cons-grid">
                    <div>
                        <div style="font-size:10px;color:var(--green);font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px">Qty Sold</div>
                        <input class="cons-input" type="number" placeholder="0" min="0" step="any" value="0" data-field="sold">
                    </div>
                    <div>
                        <div style="font-size:10px;color:var(--blue);font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px">Qty Returned</div>
                        <input class="cons-input" type="number" placeholder="0" min="0" step="any" value="0" data-field="returned">
                    </div>
                </div>
            </div>`).join("")}
        <button class="btn btn-teal" style="width:100%;margin-top:8px;justify-content:center" onclick="saveSettle()">Confirm Settlement & Recognize Revenue</button>
    `;
    document.getElementById("side-bg").classList.add("open");
    document.getElementById("side-panel").classList.add("open");
}

function closeSide(){
    document.getElementById("side-bg").classList.remove("open");
    document.getElementById("side-panel").classList.remove("open");
}

async function saveSettle(){
    let rows=document.querySelectorAll(".cons-item-card");
    let items=[];
    rows.forEach(row=>{
        items.push({
            consignment_item_id: parseInt(row.dataset.itemId),
            qty_sold:     parseFloat(row.querySelector('[data-field="sold"]').value)||0,
            qty_returned: parseFloat(row.querySelector('[data-field="returned"]').value)||0,
        });
    });
    let res=await fetch(`/b2b/api/consignments/${settlingConsId}/settle`,{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({items}),
    });
    let data=await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    closeSide();
    showToast(`Settlement done ✓  Revenue recognized: ${data.total_revenue.toFixed(2)} EGP`);
    loadConsignments(); loadClients(); loadStats();
}

["client-modal","invoice-modal","pl-modal"].forEach(id=>{
    let el = document.getElementById(id);
    if(el) el.addEventListener("click",function(e){ if(e.target===this) this.classList.remove("open"); });
});

let toastTimer=null;
function showToast(msg){
    let t=document.getElementById("toast");
    t.innerText=msg; t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer=setTimeout(()=>t.classList.remove("show"),4500);
}

/* ── PRICE LISTS ── */
let plClientPrices = [];   // current client's price entries from API
let plInitDone = false;

function initPriceListTab(){
    if(plInitDone) return;
    plInitDone = true;
    // Populate client dropdown
    let sel = document.getElementById("pl-client-select");
    sel.innerHTML = '<option value="">— Select a client —</option>';
    allClients.forEach(c=>{
        let opt = document.createElement("option");
        opt.value = c.id; opt.textContent = c.name;
        sel.appendChild(opt);
    });
}

async function loadPriceList(){
    let clientId = document.getElementById("pl-client-select").value;
    let addBtn   = document.getElementById("btn-add-price");
    let tbody    = document.getElementById("pl-body");
    if(!clientId){
        addBtn.style.display = "none";
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:40px">Select a client to view their price list.</td></tr>`;
        return;
    }
    addBtn.style.display = "";
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:28px">Loading...</td></tr>`;
    plClientPrices = await (await fetch(`/b2b/api/clients/${clientId}/prices`)).json();
    renderPriceList();
}

function renderPriceList(){
    let tbody = document.getElementById("pl-body");
    if(!plClientPrices.length){
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:40px">No custom prices set. All products use default pricing.</td></tr>`;
        return;
    }
    tbody.innerHTML = plClientPrices.map(cp=>{
        let diff = cp.custom_price - cp.default_price;
        let diffStr = diff === 0 ? "—" : (diff > 0 ? `<span style="color:var(--warn)">+${diff.toFixed(2)}</span>` : `<span style="color:var(--green)">${diff.toFixed(2)}</span>`);
        return `<tr>
            <td class="name">${cp.product_name}</td>
            <td style="font-family:var(--mono);font-size:12px;color:var(--muted)">${cp.sku}</td>
            <td style="font-family:var(--mono);color:var(--muted)">${cp.default_price.toFixed(2)}</td>
            <td style="font-family:var(--mono);font-weight:700;color:var(--blue)">${cp.custom_price.toFixed(2)}</td>
            <td style="font-family:var(--mono)">${diffStr}</td>
            <td style="display:flex;gap:6px">
                <button class="action-btn" onclick="editPriceEntry(${cp.product_id},${cp.custom_price})">Edit</button>
                <button class="action-btn danger" onclick="deletePriceEntry(${cp.product_id},'${cp.product_name.replace(/'/g,"\\'")}')">Remove</button>
            </td>
        </tr>`;
    }).join("");
}

function openAddPriceModal(){
    let clientId = document.getElementById("pl-client-select").value;
    if(!clientId){ showToast("Select a client first"); return; }
    let client = allClients.find(c=>c.id==clientId);
    document.getElementById("pl-modal-sub").innerText = `Client: ${client ? client.name : ""}`;
    // Build product dropdown — skip products already in the list
    let existing = new Set(plClientPrices.map(p=>p.product_id));
    let prodSel = document.getElementById("pl-product");
    prodSel.innerHTML = allProducts.map(p=>{
        let label = p.sku ? `${p.sku} — ${p.name}` : p.name;
        let note  = existing.has(p.id) ? " ★" : "";
        return `<option value="${p.id}" data-default="${p.default_price || p.price}">${label}${note}</option>`;
    }).join("");
    document.getElementById("pl-price").value = "";
    onPlProductChange();
    document.getElementById("pl-modal").classList.add("open");
}

function editPriceEntry(productId, currentPrice){
    openAddPriceModal();
    // Select the right product
    let sel = document.getElementById("pl-product");
    for(let opt of sel.options){ if(parseInt(opt.value)===productId){ sel.value=productId; break; } }
    onPlProductChange();
    document.getElementById("pl-price").value = currentPrice.toFixed(2);
}

function onPlProductChange(){
    let sel  = document.getElementById("pl-product");
    let opt  = sel.options[sel.selectedIndex];
    let hint = document.getElementById("pl-default-hint");
    if(!opt || !opt.value){ hint.innerText = ""; return; }
    let def = parseFloat(opt.dataset.default) || 0;
    // Check if client already has a custom price for this product
    let existing = plClientPrices.find(cp=>cp.product_id===parseInt(opt.value));
    hint.innerHTML = `Default price: <b style="font-family:var(--mono)">${def.toFixed(2)} ج.م.</b>` +
        (existing ? `&nbsp;&nbsp;|&nbsp;&nbsp;Current custom price: <b style="font-family:var(--mono);color:var(--blue)">${existing.custom_price.toFixed(2)} ج.م.</b>` : "");
    if(!document.getElementById("pl-price").value && existing)
        document.getElementById("pl-price").value = existing.custom_price.toFixed(2);
}

async function savePriceEntry(){
    let clientId  = document.getElementById("pl-client-select").value;
    let productId = parseInt(document.getElementById("pl-product").value);
    let price     = parseFloat(document.getElementById("pl-price").value);
    if(!clientId || !productId){ showToast("Select a product"); return; }
    if(isNaN(price) || price < 0){ showToast("Enter a valid price"); return; }
    let res  = await fetch(`/b2b/api/clients/${clientId}/prices`, {
        method:"PUT", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({product_id: productId, price}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    document.getElementById("pl-modal").classList.remove("open");
    showToast("Price saved ✓");
    plClientPrices = await (await fetch(`/b2b/api/clients/${clientId}/prices`)).json();
    renderPriceList();
}

async function deletePriceEntry(productId, productName){
    let clientId = document.getElementById("pl-client-select").value;
    if(!confirm(`Remove custom price for "${productName}"? The default product price will apply.`)) return;
    await fetch(`/b2b/api/clients/${clientId}/prices/${productId}`, {method:"DELETE"});
    showToast("Custom price removed ✓");
    plClientPrices = await (await fetch(`/b2b/api/clients/${clientId}/prices`)).json();
    renderPriceList();
}

init();
</script>
</body>
</html>"""
