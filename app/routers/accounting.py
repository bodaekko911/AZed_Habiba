from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import HTMLResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import and_, func, or_, select
from sqlalchemy.orm import selectinload
from typing import Optional, List
from pydantic import BaseModel, Field
from datetime import date, datetime, time, timedelta, timezone
from zoneinfo import ZoneInfo

from app.database import get_async_session
from app.core.config import settings
from app.core.permissions import get_current_user, require_admin, require_permission
from app.core.log import record
from app.core.navigation import render_app_header
from app.models.accounting import Account, Journal, JournalEntry
from app.models.b2b import (
    B2BClient, B2BInvoice, B2BInvoiceItem, B2BRefund,
    Consignment, ConsignmentItem, ConsignmentSale, ConsignmentSaleItem,
)
from app.models.product import Product
from app.models.expense import Expense, ExpenseCategory
from app.models.user import User
from app.models.invoice import Invoice
from app.models.refund import RetailRefund
from app.services.b2b_shared import (
    client_invoice_balance_subquery,
    client_outstanding_map,
    client_outstanding_value,
    client_refund_subquery,
)
from app.schemas.invoice import B2BPaymentRequest, ConsignmentSaleItemIn
from decimal import Decimal

router = APIRouter(
    prefix="/accounting",
    tags=["Accounting"],
    dependencies=[Depends(require_permission("page_accounting"))],
)


# ── Schemas ────────────────────────────────────────────
class AccountCreate(BaseModel):
    code:      str = Field(..., min_length=1, max_length=50)
    name:      str = Field(..., min_length=1, max_length=150)
    type:      str = Field(..., min_length=1, max_length=50)
    parent_id: Optional[int] = None

class JournalEntryIn(BaseModel):
    account_id: int
    debit:      float = Field(0, ge=0)
    credit:     float = Field(0, ge=0)
    note:       Optional[str] = Field(None, max_length=255)

class JournalCreate(BaseModel):
    ref_type:    Optional[str] = Field(None, max_length=100)
    description: Optional[str] = Field(None, max_length=500)
    entries:     List[JournalEntryIn]

class B2BRefundItemIn(BaseModel):
    product_id: int
    qty:        float = Field(..., gt=0)
    unit_price: float = Field(..., ge=0)


class B2BRefundIn(BaseModel):
    # Item-based, matching the B2B page. The old amount-only form could not
    # restock returned goods or record what was actually sent back.
    items:  List[B2BRefundItemIn]
    reason: Optional[str] = Field(None, max_length=255)


# ── ACCOUNTS API ───────────────────────────────────────
@router.get("/api/accounts")
async def get_accounts(db: AsyncSession = Depends(get_async_session)):
    result = await db.execute(select(Account).order_by(Account.code))
    accounts = result.scalars().all()
    return [
        {
            "id":      a.id,
            "code":    a.code,
            "name":    a.name,
            "type":    a.type,
            "balance": float(a.balance),
        }
        for a in accounts
    ]

@router.post("/api/accounts")
async def create_account(data: AccountCreate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(require_admin)):
    result = await db.execute(select(Account).where(Account.code == data.code))
    if result.scalar_one_or_none():
        raise HTTPException(status_code=400, detail="Account code already exists")
    a = Account(**data.model_dump())
    db.add(a); await db.commit(); await db.refresh(a)
    return {"id": a.id, "code": a.code, "name": a.name}

@router.delete("/api/accounts/{account_id}")
async def delete_account(account_id: int, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(require_admin)):
    result = await db.execute(select(Account).options(selectinload(Account.entries)).where(Account.id == account_id))
    a = result.scalar_one_or_none()
    if not a:
        raise HTTPException(status_code=404, detail="Account not found")
    if a.entries:
        raise HTTPException(status_code=400, detail="Cannot delete account with journal entries")
    await db.delete(a); await db.commit()
    return {"ok": True}

@router.post("/api/accounts/seed")
async def seed_accounts(db: AsyncSession = Depends(get_async_session), _: User = Depends(require_admin)):
    """Create a standard chart of accounts if none exist."""
    cnt_result = await db.execute(select(func.count()).select_from(Account))
    if cnt_result.scalar() > 0:
        return {"message": "Accounts already exist"}

    defaults = [
        # Assets
        ("1000", "Cash",                  "asset"),
        ("1100", "Accounts Receivable",   "asset"),
        ("1200", "Inventory",             "asset"),
        ("1300", "Prepaid Expenses",      "asset"),
        # Liabilities
        ("2000", "Accounts Payable",      "liability"),
        ("2100", "Salaries Payable",      "liability"),
        ("2200", "Tax Payable",           "liability"),
        # Equity
        ("3000", "Owner Equity",          "equity"),
        ("3100", "Retained Earnings",     "equity"),
        # Revenue
        ("4000", "Sales Revenue",         "revenue"),
        ("4100", "Other Income",          "revenue"),
        # Expenses
        ("5000", "Cost of Goods Sold",    "expense"),
        ("5100", "Salaries Expense",      "expense"),
        ("5200", "Rent Expense",          "expense"),
        ("5300", "Utilities Expense",     "expense"),
        ("5400", "Marketing Expense",     "expense"),
        ("5500", "Other Expenses",        "expense"),
    ]

    for code, name, atype in defaults:
        db.add(Account(code=code, name=name, type=atype, balance=0))
    await db.commit()
    return {"message": f"Created {len(defaults)} accounts"}


# ── JOURNALS API ───────────────────────────────────────
@router.get("/api/journals")
async def get_journals(
    skip: int = Query(0, ge=0),
    limit: int = Query(50, ge=1, le=200),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=200),
    from_date: Optional[date] = Query(None),
    to_date: Optional[date] = Query(None),
    db: AsyncSession = Depends(get_async_session),
):
    page, page_size, skip, limit = _normalize_journal_pagination(page, page_size, skip, limit)
    base_stmt = _apply_date_range(select(Journal), Journal.created_at, from_date, to_date)
    count_stmt = _apply_date_range(select(func.count()).select_from(Journal), Journal.created_at, from_date, to_date)
    cnt_result = await db.execute(count_stmt)
    total = int(cnt_result.scalar() or 0)
    total_pages = max(1, (total + page_size - 1) // page_size) if total else 1
    result = await db.execute(
        base_stmt
        .options(selectinload(Journal.entries).selectinload(JournalEntry.account))
        .order_by(Journal.created_at.desc())
        .offset(skip)
        .limit(limit)
    )
    journals = result.scalars().all()
    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "journals": [
            {
                "id":          j.id,
                "ref_type":    j.ref_type or "manual",
                "description": j.description or "—",
                "created_at":  j.created_at.strftime("%Y-%m-%d %H:%M") if j.created_at else "—",
                "entries_count": len(j.entries),
                "total_debit": sum(float(e.debit) for e in j.entries),
            }
            for j in journals
        ],
        "from_date": from_date.isoformat() if from_date else None,
        "to_date": to_date.isoformat() if to_date else None,
    }

@router.get("/api/journals/{journal_id}")
async def get_journal(journal_id: int, db: AsyncSession = Depends(get_async_session)):
    result = await db.execute(
        select(Journal)
        .options(selectinload(Journal.entries).selectinload(JournalEntry.account))
        .where(Journal.id == journal_id)
    )
    j = result.scalar_one_or_none()
    if not j:
        raise HTTPException(status_code=404, detail="Journal not found")
    return {
        "id":          j.id,
        "ref_type":    j.ref_type or "manual",
        "description": j.description or "—",
        "created_at":  j.created_at.strftime("%Y-%m-%d %H:%M") if j.created_at else "—",
        "entries": [
            {
                "account_code": e.account.code if e.account else "—",
                "account_name": e.account.name if e.account else "—",
                "debit":        float(e.debit),
                "credit":       float(e.credit),
                "note":         e.note or "",
            }
            for e in j.entries
        ],
    }

@router.post("/api/journals", dependencies=[Depends(require_permission("action_accounting_post_journal"))])
async def create_journal(data: JournalCreate, db: AsyncSession = Depends(get_async_session), current_user: User = Depends(get_current_user)):
    total_debit  = sum(e.debit  for e in data.entries)
    total_credit = sum(e.credit for e in data.entries)
    if round(total_debit, 2) != round(total_credit, 2):
        raise HTTPException(
            status_code=400,
            detail=f"Journal not balanced. Debits: {total_debit:.2f}, Credits: {total_credit:.2f}"
        )

    journal = Journal(
        ref_type=data.ref_type or "manual",
        description=data.description,
        user_id=current_user.id,
    )
    db.add(journal); await db.flush()

    for entry in data.entries:
        acc_result = await db.execute(select(Account).where(Account.id == entry.account_id))
        acc = acc_result.scalar_one_or_none()
        if not acc:
            raise HTTPException(status_code=404, detail=f"Account ID not found: {entry.account_id}")

        je = JournalEntry(
            journal_id=journal.id,
            account_id=entry.account_id,
            debit=entry.debit,
            credit=entry.credit,
            note=entry.note,
        )
        db.add(je)

        # Update account balance
        acc.balance += entry.debit - entry.credit

    record(db, "Accounting", "create_journal",
           f"Manual journal: {data.description or '—'} — debit total: {total_debit:.2f}",
           user=current_user, ref_type="journal", ref_id=journal.id)
    await db.commit(); await db.refresh(journal)
    return {"id": journal.id, "ok": True}



def _num(value) -> float:
    """Safely convert database numeric values to float."""
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _validate_date_range(from_date: Optional[date], to_date: Optional[date]) -> None:
    if from_date and to_date and from_date > to_date:
        raise HTTPException(status_code=400, detail="From date cannot be after To date")


def _normalize_journal_pagination(
    page: int | str | float | None,
    page_size: int | str | float | None,
    skip: int | str | float | None,
    limit: int | str | float | None,
) -> tuple[int, int, int, int]:
    page = int(page) if isinstance(page, (int, float, str)) else 1
    page_size = int(page_size) if isinstance(page_size, (int, float, str)) else 50
    if page < 1:
        page = 1
    if page_size < 1:
        page_size = 50
    if page_size > 200:
        page_size = 200

    # Keep backward compatibility for callers still sending skip/limit.
    if isinstance(skip, (int, float, str)) and int(skip) > 0 and not (isinstance(page, int) and page > 1):
        skip = int(skip)
        limit = int(limit) if isinstance(limit, (int, float, str)) else page_size
        limit = max(1, min(limit, 200))
        page_size = limit
        page = (skip // limit) + 1
        return page, page_size, skip, limit

    skip = (page - 1) * page_size
    limit = page_size
    return page, page_size, skip, limit


def _local_day_bounds_utc(
    from_date: Optional[date],
    to_date: Optional[date],
) -> tuple[Optional[datetime], Optional[datetime]]:
    """
    Convert local app dates to UTC datetime bounds.

    This keeps Accounting P&L aligned with the Dashboard date logic:
    selected local day/month -> equivalent UTC timestamp range.
    """
    tz = ZoneInfo(settings.APP_TIMEZONE)
    utc = ZoneInfo("UTC")

    utc_start = None
    utc_end = None

    if from_date:
        utc_start = datetime.combine(from_date, time.min, tzinfo=tz).astimezone(utc)

    if to_date:
        utc_end = datetime.combine(to_date + timedelta(days=1), time.min, tzinfo=tz).astimezone(utc)

    return utc_start, utc_end


def _apply_date_range(stmt, column, from_date: Optional[date], to_date: Optional[date]):
    _validate_date_range(from_date, to_date)

    utc_start, utc_end = _local_day_bounds_utc(from_date, to_date)

    if utc_start:
        stmt = stmt.where(column >= utc_start)

    if utc_end:
        stmt = stmt.where(column < utc_end)

    return stmt


def _apply_as_of_date(stmt, column, as_of: Optional[date]):
    if as_of:
        stmt = stmt.where(func.date(column) <= as_of)
    return stmt


def _apply_period_with_fallback_date(stmt, journal_column, from_date: Optional[date], to_date: Optional[date], fallback_date_column=None):
    _validate_date_range(from_date, to_date)
    if from_date:
        journal_from = datetime.combine(from_date, time.min, tzinfo=timezone.utc)
        if fallback_date_column is not None:
            stmt = stmt.where(
                or_(
                    and_(fallback_date_column.is_not(None), fallback_date_column >= from_date),
                    and_(fallback_date_column.is_(None), journal_column >= journal_from),
                )
            )
        else:
            stmt = stmt.where(journal_column >= journal_from)
    if to_date:
        journal_to = datetime.combine(to_date + timedelta(days=1), time.min, tzinfo=timezone.utc)
        if fallback_date_column is not None:
            stmt = stmt.where(
                or_(
                    and_(fallback_date_column.is_not(None), fallback_date_column <= to_date),
                    and_(fallback_date_column.is_(None), journal_column < journal_to),
                )
            )
        else:
            stmt = stmt.where(journal_column < journal_to)
    return stmt


# ── REPORTS API ────────────────────────────────────────
@router.get("/api/trial-balance")
async def trial_balance(
    as_of: Optional[date] = Query(None),
    db: AsyncSession = Depends(get_async_session),
):
    ledger_totals = (
        _apply_as_of_date(
            select(
                JournalEntry.account_id.label("account_id"),
                func.coalesce(func.sum(JournalEntry.debit), 0).label("ledger_debit"),
                func.coalesce(func.sum(JournalEntry.credit), 0).label("ledger_credit"),
            ).select_from(JournalEntry).join(Journal, Journal.id == JournalEntry.journal_id),
            Journal.created_at,
            as_of,
        )
        .group_by(JournalEntry.account_id)
        .subquery()
    )
    result = await db.execute(
        select(
            Account,
            func.coalesce(ledger_totals.c.ledger_debit, 0).label("ledger_debit"),
            func.coalesce(ledger_totals.c.ledger_credit, 0).label("ledger_credit"),
        )
        .outerjoin(ledger_totals, ledger_totals.c.account_id == Account.id)
        .order_by(Account.code)
    )
    account_rows = result.all()
    rows = []
    total_debit = 0.0
    total_credit = 0.0
    drift_accounts = []

    for a, ledger_debit_raw, ledger_credit_raw in account_rows:
        ledger_debit = float(ledger_debit_raw or 0)
        ledger_credit = float(ledger_credit_raw or 0)
        net_balance = round(ledger_debit - ledger_credit, 2)
        stored_balance = round(float(a.balance or 0), 2)
        debit_balance = net_balance if net_balance > 0 else 0.0
        credit_balance = abs(net_balance) if net_balance < 0 else 0.0
        drift = round(stored_balance - net_balance, 2)
        total_debit += debit_balance
        total_credit += credit_balance
        rows.append({
            "code": a.code,
            "name": a.name,
            "type": a.type,
            "debit": round(debit_balance, 2),
            "credit": round(credit_balance, 2),
            "ledger_debit": round(ledger_debit, 2),
            "ledger_credit": round(ledger_credit, 2),
            "net_balance": net_balance,
            "stored_balance": stored_balance,
            "balance_drift": drift,
        })
        if abs(drift) >= 0.01:
            drift_accounts.append({
                "code": a.code,
                "name": a.name,
                "ledger_balance": net_balance,
                "stored_balance": stored_balance,
                "drift": drift,
            })

    return {
        "rows": rows,
        "total_debit": round(total_debit, 2),
        "total_credit": round(total_credit, 2),
        "drift_count": len(drift_accounts),
        "drift_accounts": drift_accounts,
        "as_of": as_of.isoformat() if as_of else None,
    }

@router.get("/api/profit-loss")
async def profit_loss(
    from_date: Optional[date] = Query(None),
    to_date: Optional[date] = Query(None),
    db: AsyncSession = Depends(get_async_session),
):
    """
    Operational P&L using the same revenue definition as Dashboard and Sales Report.

    Revenue = paid POS invoices
            + paid B2B invoices
            - retail refunds
            - B2B refunds

    Operational Cost = expenses recorded in the Expenses module by expense_date.
    Operating Profit = Revenue - Operational Cost.
    """
    _validate_date_range(from_date, to_date)

    pos_sales_stmt = _apply_date_range(
        select(func.coalesce(func.sum(Invoice.total), 0)),
        Invoice.created_at,
        from_date,
        to_date,
    ).where(Invoice.status == "paid")
    pos_sales_result = await db.execute(pos_sales_stmt)
    pos_sales = _num(pos_sales_result.scalar())

    b2b_sales_stmt = _apply_date_range(
        select(func.coalesce(func.sum(B2BInvoice.total), 0)),
        B2BInvoice.created_at,
        from_date,
        to_date,
    ).where(B2BInvoice.status == "paid")
    b2b_sales_result = await db.execute(b2b_sales_stmt)
    b2b_sales = _num(b2b_sales_result.scalar())

    retail_refunds_stmt = _apply_date_range(
        select(
            func.coalesce(func.sum(RetailRefund.total), 0).label("total"),
            func.count(RetailRefund.id).label("count"),
        ),
        RetailRefund.created_at,
        from_date,
        to_date,
    )
    retail_refunds_row = (await db.execute(retail_refunds_stmt)).mappings().one()
    retail_refunds = _num(retail_refunds_row["total"])
    retail_refund_count = _num(retail_refunds_row["count"])

    b2b_refunds_stmt = _apply_date_range(
        select(
            func.coalesce(func.sum(B2BRefund.total), 0).label("total"),
            func.count(B2BRefund.id).label("count"),
        ),
        B2BRefund.created_at,
        from_date,
        to_date,
    )
    b2b_refunds_row = (await db.execute(b2b_refunds_stmt)).mappings().one()
    b2b_refunds = _num(b2b_refunds_row["total"])
    b2b_refund_count = _num(b2b_refunds_row["count"])

    total_refunds = retail_refunds + b2b_refunds
    total_revenue = round(pos_sales + b2b_sales - total_refunds, 2)

    expense_stmt = (
        select(
            ExpenseCategory.account_code.label("code"),
            ExpenseCategory.name.label("name"),
            func.coalesce(func.sum(Expense.amount), 0).label("amount"),
        )
        .select_from(Expense)
        .join(ExpenseCategory, ExpenseCategory.id == Expense.category_id, isouter=True)
    )

    if from_date:
        expense_stmt = expense_stmt.where(Expense.expense_date >= from_date)
    if to_date:
        expense_stmt = expense_stmt.where(Expense.expense_date <= to_date)

    expense_stmt = (
        expense_stmt
        .group_by(ExpenseCategory.account_code, ExpenseCategory.name)
        .order_by(ExpenseCategory.name)
    )
    expense_rows = (await db.execute(expense_stmt)).mappings().all()

    expenses = []
    for row in expense_rows:
        amount = round(_num(row["amount"]), 2)
        if abs(amount) < 0.01:
            continue
        expenses.append(
            {
                "code": row["code"] or "OPC",
                "name": row["name"] or "Operational Cost",
                "amount": amount,
            }
        )

    total_expense = round(sum(item["amount"] for item in expenses), 2)
    net_profit = round(total_revenue - total_expense, 2)

    revenues = []

    if abs(pos_sales) >= 0.01:
        revenues.append(
            {"code": "POS", "name": "Paid POS Sales", "amount": round(pos_sales, 2)}
        )

    if abs(b2b_sales) >= 0.01:
        revenues.append(
            {"code": "B2B", "name": "Paid B2B Sales", "amount": round(b2b_sales, 2)}
        )

    return {
        "revenues": revenues,
        "expenses": expenses,
        "total_revenue": total_revenue,
        "total_expense": total_expense,
        "net_profit": net_profit,
        "total_refunds": round(total_refunds, 2),
        "refund_count": int(retail_refund_count + b2b_refund_count),
        "refund_breakdown": {
            "retail": round(retail_refunds, 2),
            "b2b": round(b2b_refunds, 2),
        },
        "revenue_breakdown": {
            "pos_paid_sales": round(pos_sales, 2),
            "b2b_paid_sales": round(b2b_sales, 2),
            "retail_refunds": round(retail_refunds, 2),
            "b2b_refunds": round(b2b_refunds, 2),
        },
        "from_date": from_date.isoformat() if from_date else None,
        "to_date": to_date.isoformat() if to_date else None,
    }



# ── B2B INVOICES (for Accounting tab) ─────────────────
@router.get("/api/b2b-invoices")
async def get_b2b_invoices(
    invoice_type: Optional[str] = Query(None, max_length=50),
    status: Optional[str] = Query(None, max_length=50),
    search: Optional[str] = Query(None, max_length=150),
    from_date: Optional[date] = Query(None),
    to_date: Optional[date] = Query(None),
    db: AsyncSession = Depends(get_async_session),
):
    invoice_type = invoice_type.strip() if isinstance(invoice_type, str) and invoice_type.strip() else None
    status = status.strip() if isinstance(status, str) and status.strip() else None
    search = " ".join(search.split()) if isinstance(search, str) and search.strip() else None
    stmt = _apply_date_range(
        select(B2BInvoice)
        .options(
            selectinload(B2BInvoice.client),
            selectinload(B2BInvoice.items).selectinload(B2BInvoiceItem.product),
        )
        .order_by(B2BInvoice.created_at.desc()),
        B2BInvoice.created_at,
        from_date,
        to_date,
    )
    if invoice_type:
        stmt = stmt.where(B2BInvoice.invoice_type == invoice_type)
    if status:
        stmt = stmt.where(B2BInvoice.status == status)
    if search:
        search_term = f"%{search}%"
        stmt = (
            stmt.join(B2BClient, B2BClient.id == B2BInvoice.client_id, isouter=True)
            .where(
                or_(
                    B2BInvoice.invoice_number.ilike(search_term),
                    B2BClient.name.ilike(search_term),
                )
            )
        )
    result = await db.execute(stmt)
    invoices = result.scalars().all()
    # One grouped query rather than one per row, and the same definition the
    # clients list uses — this row used to read the stored client.outstanding,
    # which drifts and ignores refunds.
    outstanding_by_client = await client_outstanding_map(db)
    return [
        {
            "id":             i.id,
            "invoice_number": i.invoice_number,
            "client":         i.client.name if i.client else "—",
            "client_id":      i.client_id,
            "client_outstanding": outstanding_by_client.get(i.client_id, 0.0),
            "invoice_type":   i.invoice_type,
            "status":         i.status,
            "subtotal":       float(i.subtotal),
            "discount":       float(i.discount),
            "total":          float(i.total),
            "amount_paid":    float(i.amount_paid),
            "balance_due":    round(float(i.total) - float(i.amount_paid), 2),
            "created_at":     i.created_at.strftime("%Y-%m-%d") if i.created_at else "—",
            "items": [
                {
                    "product":    it.product.name if it.product else "—",
                    "qty":        float(it.qty),
                    "unit_price": float(it.unit_price),
                    "total":      float(it.total),
                }
                for it in i.items
            ],
        }
        for i in invoices
    ]


@router.get("/api/b2b-clients")
async def get_accounting_b2b_clients(
    q: Optional[str] = Query(None, max_length=150),
    db: AsyncSession = Depends(get_async_session),
):
    q = " ".join(q.split()) if isinstance(q, str) and q.strip() else None
    # Outstanding = unpaid invoice balances LESS refunds issued — the shared
    # definition, so this screen agrees with the B2B page and the statement.
    # The refund leg used to be missing here, so a refund never moved the
    # balance shown on this list.
    outstanding_sub = client_invoice_balance_subquery()
    refund_sub = client_refund_subquery()
    stmt = (
        select(
            B2BClient,
            (
                func.coalesce(outstanding_sub.c.outstanding, 0)
                - func.coalesce(refund_sub.c.refunded, 0)
            ).label("computed_outstanding"),
        )
        .outerjoin(outstanding_sub, outstanding_sub.c.client_id == B2BClient.id)
        .outerjoin(refund_sub, refund_sub.c.client_id == B2BClient.id)
        .where(B2BClient.is_active == True)
        .options(selectinload(B2BClient.invoices))
        .order_by(B2BClient.name)
    )
    if q:
        term = f"%{q}%"
        stmt = stmt.where(
            or_(
                B2BClient.name.ilike(term),
                B2BClient.phone.ilike(term),
                B2BClient.contact_person.ilike(term),
            )
        )
    result = await db.execute(stmt)
    rows = result.all()
    return [
        {
            "id": c.id,
            "name": c.name,
            "contact_person": c.contact_person or "—",
            "phone": c.phone or "—",
            "email": c.email or "—",
            "payment_terms": c.payment_terms,
            "credit_limit": float(c.credit_limit or 0),
            "discount_pct": float(c.discount_pct or 0),
            "outstanding": max(float(computed_outstanding or 0), 0.0),
            "invoice_count": len(c.invoices),
            "is_consignment": (c.payment_terms or "").strip().lower() == "consignment",
        }
        for c, computed_outstanding in rows
    ]

@router.post("/api/b2b-invoices/{invoice_id}/collect", dependencies=[Depends(require_permission("action_b2b_collect"))])
async def collect_b2b_payment(
    invoice_id: int,
    data: B2BPaymentRequest,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    inv_result = await db.execute(
        select(B2BInvoice)
        .options(selectinload(B2BInvoice.client))
        .where(B2BInvoice.id == invoice_id)
    )
    invoice = inv_result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if invoice.invoice_type not in ("cash", "full_payment"):
        raise HTTPException(status_code=400, detail="Use consignment-payment endpoint for consignment invoices")
    amount  = round(float(data.amount), 2)
    balance = round(float(invoice.total) - float(invoice.amount_paid), 2)
    if amount <= 0:
        raise HTTPException(status_code=400, detail="Amount must be greater than 0")
    if amount > balance + 0.01:
        raise HTTPException(status_code=400, detail=f"Amount exceeds balance: {balance:.2f}")
    invoice.amount_paid = Decimal(str(float(invoice.amount_paid) + amount))
    invoice.status = "paid" if float(invoice.amount_paid) >= float(invoice.total) else "partial"
    client = invoice.client
    if client:
        client.outstanding = Decimal(str(max(0, float(client.outstanding) - amount)))
    # Journal: Cash in, AR out, Deferred → Revenue
    journal = Journal(
        ref_type="b2b_collection",
        ref_id=invoice.id,
        description=f"B2B payment collected - {invoice.invoice_number}",
        user_id=current_user.id,
    )
    db.add(journal); await db.flush()
    for code, debit, credit in [
        ("1000", amount, 0),
        ("1100", 0, amount),
        ("2200", amount, 0),
        ("4000", 0, amount),
    ]:
        acc_result = await db.execute(select(Account).where(Account.code == code))
        acc = acc_result.scalar_one_or_none()
        if acc:
            db.add(JournalEntry(journal_id=journal.id, account_id=acc.id, debit=debit, credit=credit))
            acc.balance += Decimal(str(debit)) - Decimal(str(credit))
    record(db, "Accounting", "collect_b2b_payment",
           f"B2B payment collected — {invoice.invoice_number} — amount: {amount:.2f} — status: {invoice.status}",
           ref_type="b2b_invoice", ref_id=invoice_id)
    await db.commit()
    return {"ok": True, "status": invoice.status, "invoice_number": invoice.invoice_number}


async def _record_consignment_client_payment(
    db: AsyncSession,
    *,
    client: B2BClient,
    amount: float,
    month_label: str,
    current_user: User,
    sold_items: Optional[List["ConsignmentSaleItemIn"]] = None,
):
    # When sold items are supplied, the amount is the reconciled total of the
    # items the client reported sold, net of the client's agreed discount — the
    # consignment invoice it pays down is booked net of that same discount, so
    # collecting the gross line total would over-pay the invoice.
    discount_pct = max(0.0, min(100.0, float(client.discount_pct or 0)))
    sale_lines: list[dict] = []
    items_subtotal = 0.0
    items_discount = 0.0
    if sold_items:
        for line in sold_items:
            qty        = round(float(line.qty), 3)
            unit_price = round(float(line.unit_price), 2)
            if qty <= 0:
                continue
            sale_lines.append({
                "product_id": int(line.product_id),
                "qty": qty,
                "unit_price": unit_price,
                "total": round(qty * unit_price, 2),
            })
        if not sale_lines:
            raise HTTPException(status_code=400, detail="No valid sold items were provided")
        items_subtotal = round(sum(l["total"] for l in sale_lines), 2)
        items_discount = round(items_subtotal * (discount_pct / 100), 2)
        items_total    = round(items_subtotal - items_discount, 2)
        if abs(items_total - round(amount, 2)) > 0.01:
            detail = f"Amount ({amount:.2f}) does not match the sold items total ({items_total:.2f})"
            if items_discount > 0:
                detail += (
                    f" — {items_subtotal:.2f} less {discount_pct:g}% client discount"
                    f" ({items_discount:.2f})"
                )
            raise HTTPException(status_code=400, detail=detail)
        # Within tolerance the server's own net figure wins, so the stored
        # subtotal - discount always reconciles exactly to what was collected.
        amount = items_total

    open_result = await db.execute(
        select(B2BInvoice)
        .where(
            B2BInvoice.client_id == client.id,
            B2BInvoice.invoice_type == "consignment",
        )
        .order_by(B2BInvoice.created_at.asc(), B2BInvoice.id.asc())
    )
    invoices = open_result.scalars().all()
    open_invoices = [inv for inv in invoices if round(float(inv.total) - float(inv.amount_paid), 2) > 0.01]
    if not open_invoices:
        raise HTTPException(status_code=400, detail="This client has no open consignment invoices")

    outstanding = round(await client_outstanding_value(db, client.id), 2)
    open_balance_total = round(sum(float(inv.total) - float(inv.amount_paid) for inv in open_invoices), 2)
    if amount <= 0:
        raise HTTPException(status_code=400, detail="Amount must be greater than 0")
    if outstanding <= 0.01:
        raise HTTPException(status_code=400, detail="This client has no outstanding balance to reduce")
    if amount > outstanding + 0.01:
        raise HTTPException(status_code=400, detail=f"Amount exceeds client outstanding: {outstanding:.2f}")
    if amount > open_balance_total + 0.01:
        raise HTTPException(status_code=400, detail=f"Amount exceeds open consignment balance: {open_balance_total:.2f}")

    remaining = round(amount, 2)
    allocations = []
    for invoice in open_invoices:
        balance = round(float(invoice.total) - float(invoice.amount_paid), 2)
        if balance <= 0.01:
            continue
        applied = round(min(balance, remaining), 2)
        if applied <= 0:
            continue
        invoice.amount_paid = Decimal(str(round(float(invoice.amount_paid) + applied, 2)))
        invoice.status = "paid" if float(invoice.amount_paid) >= float(invoice.total) - 0.01 else "partial"
        allocations.append({
            "invoice_id": invoice.id,
            "invoice_number": invoice.invoice_number,
            "applied_amount": applied,
        })
        remaining = round(remaining - applied, 2)
        if remaining <= 0.009:
            remaining = 0.0
            break

    if remaining > 0.01:
        raise HTTPException(status_code=400, detail="Could not allocate the full payment across open consignment invoices")

    client.outstanding = Decimal(str(max(0, round(float(client.outstanding) - amount, 2))))

    note = f"Consignment client payment - {client.name}"
    if month_label:
        note += f" ({month_label})"
    if allocations:
        note += f" - {', '.join(a['invoice_number'] for a in allocations[:3])}"
        if len(allocations) > 3:
            note += f" +{len(allocations) - 3} more"

    journal = Journal(
        ref_type="consignment_client_payment",
        ref_id=client.id,
        description=note,
        user_id=current_user.id,
    )
    db.add(journal)
    await db.flush()
    for code, debit, credit in [
        ("1000", amount, 0),
        ("1100", 0, amount),
        ("2200", amount, 0),
        ("4000", 0, amount),
    ]:
        acc_result = await db.execute(select(Account).where(Account.code == code))
        acc = acc_result.scalar_one_or_none()
        if acc:
            db.add(JournalEntry(journal_id=journal.id, account_id=acc.id, debit=debit, credit=credit))
            acc.balance += Decimal(str(debit)) - Decimal(str(credit))

    # Persist the reconciled sale record (sold items + month) alongside the
    # payment. Bookkeeping only — consignment quantities/stock are untouched.
    sale_id = None
    if sale_lines:
        sale = ConsignmentSale(
            client_id=client.id,
            user_id=current_user.id,
            journal_id=journal.id,
            month_label=month_label or None,
            subtotal=Decimal(str(items_subtotal)),
            discount=Decimal(str(items_discount)),
            amount=Decimal(str(round(amount, 2))),
        )
        db.add(sale)
        await db.flush()
        for l in sale_lines:
            db.add(ConsignmentSaleItem(
                sale_id=sale.id,
                product_id=l["product_id"],
                qty=Decimal(str(l["qty"])),
                unit_price=Decimal(str(l["unit_price"])),
                total=Decimal(str(l["total"])),
            ))
        sale_id = sale.id

    record(
        db,
        "Accounting",
        "collect_consignment_client_payment",
        f"Consignment client payment - {client.name} - amount: {amount:.2f}"
        + (f" (sold {items_subtotal:.2f} less {discount_pct:g}% discount {items_discount:.2f})" if items_discount > 0 else "")
        + (f" - {month_label}" if month_label else ""),
        user=current_user,
        ref_type="b2b_client",
        ref_id=client.id,
    )
    return {
        "ok": True,
        "client_id": client.id,
        "client": client.name,
        "client_outstanding": round(float(client.outstanding), 2),
        "subtotal": items_subtotal,
        "discount": items_discount,
        "discount_pct": discount_pct,
        "amount": round(amount, 2),
        "allocations": allocations,
        "journal_id": journal.id,
        "sale_id": sale_id,
    }


@router.post("/api/b2b-clients/{client_id}/consignment-payment", dependencies=[Depends(require_permission("action_b2b_consignment_settle"))])
async def accounting_client_consignment_payment(
    client_id: int,
    data: B2BPaymentRequest,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    client_result = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = client_result.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")

    amount = round(float(data.amount), 2)
    month_label = (data.month_label or "").strip()
    payload = await _record_consignment_client_payment(
        db,
        client=client,
        amount=amount,
        month_label=month_label,
        current_user=current_user,
        sold_items=data.items,
    )
    await db.commit()
    return payload


@router.post("/api/b2b-invoices/{invoice_id}/consignment-payment", dependencies=[Depends(require_permission("action_b2b_consignment_settle"))])
async def accounting_consignment_payment(
    invoice_id: int,
    data: B2BPaymentRequest,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    """Record a monthly/partial cash payment against a consignment invoice."""
    inv_result = await db.execute(
        select(B2BInvoice)
        .options(selectinload(B2BInvoice.client))
        .where(B2BInvoice.id == invoice_id)
    )
    invoice = inv_result.scalar_one_or_none()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    if invoice.invoice_type != "consignment":
        raise HTTPException(status_code=400, detail="This endpoint is for consignment invoices only")
    client = invoice.client
    if not client:
        raise HTTPException(status_code=400, detail="Consignment invoice has no client")
    payload = await _record_consignment_client_payment(
        db,
        client=client,
        amount=round(float(data.amount), 2),
        month_label=(data.month_label or "").strip(),
        current_user=current_user,
        sold_items=data.items,
    )
    await db.commit()
    return payload


@router.get("/api/b2b-clients/{client_id}/consignment-items")
async def get_client_consignment_items(client_id: int, db: AsyncSession = Depends(get_async_session)):
    """
    Sellable items for a consignment client — the products this client was sent
    on consignment, with their consignment unit prices. Used to populate the
    sold-items picker on the Record Consignment Payment screen. Aggregated
    across the client's non-closed consignments and grouped by product + price.

    Unit prices are gross (as sent); the client's discount is returned
    alongside them under ``discount_pct`` and is applied to the reported sold
    total, matching how the consignment invoice itself was booked.
    """
    client_result = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = client_result.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    discount_pct = max(0.0, min(100.0, float(client.discount_pct or 0)))

    result = await db.execute(
        select(ConsignmentItem)
        .join(Consignment, ConsignmentItem.consignment_id == Consignment.id)
        .where(Consignment.client_id == client_id, Consignment.status != "closed")
        .options(selectinload(ConsignmentItem.product))
    )
    items = result.scalars().all()

    grouped: dict[tuple, dict] = {}
    for ci in items:
        price = round(float(ci.unit_price), 2)
        key = (ci.product_id, price)
        remaining = float(ci.qty_sent) - float(ci.qty_sold) - float(ci.qty_returned)
        row = grouped.get(key)
        if not row:
            row = grouped[key] = {
                "product_id": ci.product_id,
                "name": ci.product.name if ci.product else f"#{ci.product_id}",
                "sku": ci.product.sku if ci.product else "",
                "unit": (ci.product.unit if ci.product else "") or "piece",
                "unit_price": price,
                "qty_sent": 0.0,
                "qty_remaining": 0.0,
            }
        row["qty_sent"] += float(ci.qty_sent)
        row["qty_remaining"] += max(0.0, remaining)

    out = sorted(grouped.values(), key=lambda r: r["name"].lower())

    # Quantities already reported sold on earlier payments. That flow is
    # bookkeeping only — it never touches qty_sold — so without this the
    # picker would keep offering goods the client has already paid for, and
    # would disagree with the stock figure the client sees on their portal.
    sold_result = await db.execute(
        select(ConsignmentSaleItem)
        .join(ConsignmentSale, ConsignmentSaleItem.sale_id == ConsignmentSale.id)
        .where(ConsignmentSale.client_id == client_id)
    )
    reported_sold: dict[int, float] = {}
    for line in sold_result.scalars().all():
        reported_sold[line.product_id] = reported_sold.get(line.product_id, 0.0) + float(line.qty or 0)
    # Reported sales name a product, not a price band, so drain that product's
    # rows in order until the reported quantity is used up.
    for r in out:
        outstanding = reported_sold.get(r["product_id"], 0.0)
        if outstanding <= 0:
            continue
        taken = min(r["qty_remaining"], outstanding)
        r["qty_remaining"] -= taken
        reported_sold[r["product_id"]] = outstanding - taken

    for r in out:
        r["qty_sent"] = round(r["qty_sent"], 3)
        r["qty_remaining"] = round(max(0.0, r["qty_remaining"]), 3)
    return {"discount_pct": discount_pct, "items": out}


@router.get("/api/b2b-clients/{client_id}/consignment-sales")
async def get_client_consignment_sales(client_id: int, db: AsyncSession = Depends(get_async_session)):
    """Recorded consignment payments for a client, each with the items the
    client reported sold and the month — newest first."""
    result = await db.execute(
        select(ConsignmentSale)
        .where(ConsignmentSale.client_id == client_id)
        .options(selectinload(ConsignmentSale.items).selectinload(ConsignmentSaleItem.product))
        .order_by(ConsignmentSale.created_at.desc(), ConsignmentSale.id.desc())
    )
    sales = result.scalars().all()
    return [
        {
            "id": s.id,
            "month_label": s.month_label,
            "subtotal": round(float(s.subtotal or 0), 2),
            "discount": round(float(s.discount or 0), 2),
            "amount": round(float(s.amount or 0), 2),
            "created_at": s.created_at.isoformat() if s.created_at else None,
            "items": [
                {
                    "product_id": it.product_id,
                    "name": it.product.name if it.product else f"#{it.product_id}",
                    "unit": (it.product.unit if it.product else "") or "piece",
                    "qty": round(float(it.qty), 3),
                    "unit_price": round(float(it.unit_price), 2),
                    "total": round(float(it.total), 2),
                }
                for it in s.items
            ],
        }
        for s in sales
    ]


@router.get("/api/products-list")
async def products_for_refund(client_id: int = None, db: AsyncSession = Depends(get_async_session)):
    """Active products for the client-refund picker, with the client's own
    price where one is set.

    Proxied here (gated by page_accounting at the router level) so an
    accountant who can raise a client refund doesn't also need page_b2b or
    page_products just to populate the dropdown — same pattern as the farm
    picker on the expenses page.
    """
    from app.models.b2b import B2BClientPrice
    from app.models.product import Product

    result = await db.execute(
        select(Product).where(Product.is_active == True).order_by(Product.name)  # noqa: E712
    )
    products = result.scalars().all()

    custom: dict[int, float] = {}
    if client_id:
        price_rows = await db.execute(
            select(B2BClientPrice).where(B2BClientPrice.client_id == client_id)
        )
        for row in price_rows.scalars().all():
            custom[row.product_id] = float(row.price)

    return [
        {
            "id": p.id,
            "name": p.name,
            "sku": p.sku,
            "unit": p.unit or "",
            "price": custom.get(p.id, float(p.price or 0)),
        }
        for p in products
    ]


@router.post("/api/b2b-clients/{client_id}/refund", dependencies=[Depends(require_permission("action_b2b_refund"))])
async def refund_b2b_client_account(
    client_id: int,
    data: B2BRefundIn,
    db: AsyncSession = Depends(get_async_session),
    current_user: User = Depends(get_current_user),
):
    """Refund a B2B client from the accounting screen.

    Delegates to the same routine the B2B page uses, so a refund raised here is
    identical to one raised there: a numbered B2BRefund with line items, stock
    returned, journal posted, and the client's balance credited. It previously
    wrote only a journal and decremented a stored balance field, which meant
    the refund never showed up in the B2B refunds list, never restocked the
    returned goods, and left the client's outstanding untouched on screen.
    """
    from app.routers.b2b import ClientRefundCreate, RefundItemIn, create_client_refund_core

    client_result = await db.execute(select(B2BClient).where(B2BClient.id == client_id))
    client = client_result.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Client not found")
    if not data.items:
        raise HTTPException(status_code=400, detail="Refund must have at least one item")

    payload = ClientRefundCreate(
        client_id=client_id,
        notes=(data.reason or None),
        items=[
            RefundItemIn(product_id=i.product_id, qty=i.qty, unit_price=i.unit_price)
            for i in data.items
        ],
    )
    result = await create_client_refund_core(db, current_user, payload)
    return {**result, "client_outstanding": result["outstanding"]}


# ── UI ─────────────────────────────────────────────────
@router.get("/", response_class=HTMLResponse)
def accounting_ui(current_user: User = Depends(require_permission("page_accounting"))):
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="/static/theme-init.js"></script>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Accounting — AZed Farm</title>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root {
    --bg:      #060810;
    --surface: #0a0d18;
    --card:    #0f1424;
    --card2:   #151c30;
    --border:  rgba(255,255,255,0.06);
    --border2: rgba(255,255,255,0.11);
    --green:   #00ff9d;
    --blue:    #4d9fff;
    --purple:  #a855f7;
    --danger:  #ff4d6d;
    --warn:    #ffb547;
    --text:    #f0f4ff;
    --sub:     #8899bb;
    --muted:   #445066;
    --sans:    'Outfit', sans-serif;
    --mono:    'JetBrains Mono', monospace;
    --r:       12px;
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
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--sans); background: var(--bg); color: var(--text); min-height: 100vh; font-size: 14px; }
nav {
    position: sticky; top: 0; z-index: 100;
    display: flex; align-items: center; gap: 8px;
    padding: 0 24px; height: 58px;
    background: rgba(10,13,24,.92); backdrop-filter: blur(20px);
    border-bottom: 1px solid var(--border); flex-wrap: wrap;
}
.logo { font-size: 18px; font-weight: 900; background: linear-gradient(135deg,#f59e0b,#fbbf24); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; margin-right:10px; text-decoration:none; display:flex; align-items:center; gap:8px; cursor:pointer; }
.nav-link { padding:7px 12px; border-radius:8px; color:var(--sub); font-size:12px; font-weight:600; text-decoration:none; transition:all .2s; white-space:nowrap; }
.nav-link:hover { background:rgba(255,255,255,.05); background:linear-gradient(135deg,var(--green),var(--blue));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text; }
.nav-link.active { background:rgba(0,255,157,.1); color:var(--green); }
.nav-spacer { flex:1; }
.content { max-width:1300px; margin:0 auto; padding:28px 24px; display:flex; flex-direction:column; gap:20px; }
.page-title { font-size:24px; font-weight:800; letter-spacing:-.5px; }
.page-sub   { color:var(--muted); font-size:13px; margin-top:3px; }
.tabs { display:flex; gap:4px; background:var(--card); border:1px solid var(--border); border-radius:var(--r); padding:4px; width:fit-content; flex-wrap:wrap; }
.tab { padding:8px 18px; border-radius:9px; font-size:13px; font-weight:700; cursor:pointer; border:none; background:transparent; color:var(--muted); transition:all .2s; font-family:var(--sans); }
.tab.active { background:var(--card2); color:var(--text); }
.subtabs { display:flex; gap:4px; background:var(--card2); border:1px solid var(--border); border-radius:12px; padding:4px; width:fit-content; flex-wrap:wrap; margin-bottom:14px; }
.subtab { padding:7px 15px; border-radius:9px; font-size:12px; font-weight:700; cursor:pointer; border:none; background:transparent; color:var(--muted); transition:all .2s; font-family:var(--sans); }
.subtab.active { background:var(--card); color:var(--text); box-shadow:0 6px 16px rgba(0,0,0,.12); }
.section-card { background:var(--card); border:1px solid var(--border); border-radius:var(--r); padding:18px; margin-bottom:14px; }
.section-card-title { font-size:16px; font-weight:800; margin-bottom:4px; }
.section-card-sub { color:var(--muted); font-size:12px; }
.toolbar { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
.btn { display:flex; align-items:center; gap:7px; padding:10px 16px; border-radius:var(--r); font-family:var(--sans); font-size:13px; font-weight:700; cursor:pointer; border:none; transition:all .2s; white-space:nowrap; }
.btn-green  { background:linear-gradient(135deg,var(--green),#00d4ff); color:#021a10; }
.btn-green:hover  { filter:brightness(1.1); transform:translateY(-1px); }
.btn-blue   { background:linear-gradient(135deg,var(--blue),var(--purple)); color:white; }
.btn-blue:hover   { filter:brightness(1.1); transform:translateY(-1px); }
.btn-outline { background:transparent; border:1px solid var(--border2); color:var(--sub); }
.btn-outline:hover { border-color:var(--green); color:var(--green); }
.table-wrap { background:var(--card); border:1px solid var(--border); border-radius:var(--r); overflow:hidden; }
table { width:100%; border-collapse:collapse; }
thead { background:var(--card2); }
th { text-align:left; font-size:10px; font-weight:700; letter-spacing:1px; text-transform:uppercase; color:var(--muted); padding:12px 16px; }
td { padding:12px 16px; border-top:1px solid var(--border); color:var(--sub); font-size:13px; }
tr:hover td { background:rgba(255,255,255,.02); }
td.name { color:var(--text); font-weight:600; }
td.mono { font-family:var(--mono); }
td.dr { font-family:var(--mono); color:var(--green); }
td.cr { font-family:var(--mono); color:var(--blue); }
.type-badge { display:inline-flex; padding:2px 8px; border-radius:20px; font-size:11px; font-weight:700; }
.type-asset     { background:rgba(0,255,157,.1);  color:var(--green);  }
.type-liability { background:rgba(255,77,109,.1); color:var(--danger); }
.type-equity    { background:rgba(168,85,247,.1); color:var(--purple); }
.type-revenue   { background:rgba(77,159,255,.1); color:var(--blue);   }
.type-expense   { background:rgba(255,181,71,.1); color:var(--warn);   }
.type-refund    { background:rgba(255,77,109,.15); color:var(--danger); border:1px solid rgba(255,77,109,.3); }
.action-btn { background:transparent; border:1px solid var(--border2); color:var(--sub); font-size:12px; font-weight:600; padding:5px 10px; border-radius:7px; cursor:pointer; transition:all .15s; font-family:var(--sans); }
.action-btn:hover { border-color:var(--blue); color:var(--blue); }
.action-btn.danger:hover { border-color:var(--danger); color:var(--danger); }
.action-btn.green:hover  { border-color:var(--green); color:var(--green); }

/* PL REPORT */
.pl-section { background:var(--card); border:1px solid var(--border); border-radius:var(--r); overflow:hidden; margin-bottom:14px; }
.pl-header  { background:var(--card2); padding:12px 16px; font-size:11px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; color:var(--muted); }
.pl-row     { display:flex; justify-content:space-between; padding:11px 16px; border-top:1px solid var(--border); font-size:13px; }
.pl-row:hover { background:rgba(255,255,255,.02); }
.pl-total   { display:flex; justify-content:space-between; padding:14px 16px; border-top:1px solid var(--border2); font-size:15px; font-weight:800; }
.pl-net     { display:flex; justify-content:space-between; padding:16px; background:var(--card2); border:1px solid var(--border2); border-radius:var(--r); font-size:18px; font-weight:800; }

/* MODAL */
.modal-bg { position:fixed; inset:0; z-index:500; background:rgba(0,0,0,.7); backdrop-filter:blur(4px); display:none; align-items:center; justify-content:center; }
.modal-bg.open { display:flex; }
.modal { background:var(--card); border:1px solid var(--border2); border-radius:16px; padding:28px; width:600px; max-width:95vw; max-height:90vh; overflow-y:auto; animation:modalIn .2s ease; }
@keyframes modalIn { from{opacity:0;transform:scale(.95)} to{opacity:1;transform:scale(1)} }
.modal-title { font-size:18px; font-weight:800; margin-bottom:20px; }
.fld { display:flex; flex-direction:column; gap:6px; margin-bottom:14px; }
.fld label { font-size:11px; font-weight:700; letter-spacing:1px; text-transform:uppercase; color:var(--muted); }
.fld input, .fld select { background:var(--card2); border:1px solid var(--border2); border-radius:10px; padding:10px 12px; color:var(--text); font-family:var(--sans); font-size:14px; outline:none; transition:border-color .2s; width:100%; }
.fld input:focus, .fld select:focus { border-color:rgba(0,255,157,.4); }
.modal-actions { display:flex; gap:10px; margin-top:6px; justify-content:flex-end; }
.btn-cancel { background:transparent; border:1px solid var(--border2); color:var(--sub); padding:10px 18px; border-radius:var(--r); font-family:var(--sans); font-size:13px; font-weight:700; cursor:pointer; }
.btn-cancel:hover { border-color:var(--danger); color:var(--danger); }

/* JOURNAL ENTRY ROWS */
.entry-row { display:grid; grid-template-columns:2fr 1fr 1fr 30px; gap:8px; align-items:center; margin-bottom:8px; }
.entry-row select, .entry-row input { background:var(--card2); border:1px solid var(--border2); border-radius:8px; padding:8px 10px; color:var(--text); font-family:var(--sans); font-size:13px; outline:none; width:100%; }
.entry-row select:focus, .entry-row input:focus { border-color:rgba(0,255,157,.4); }
.rm-btn { background:none; border:none; color:var(--muted); font-size:18px; cursor:pointer; padding:0; transition:color .15s; }
.rm-btn:hover { color:var(--danger); }
.add-entry-btn { background:rgba(77,159,255,.1); border:1px dashed rgba(77,159,255,.3); color:var(--blue); font-family:var(--sans); font-size:13px; font-weight:600; padding:9px; border-radius:8px; cursor:pointer; width:100%; transition:all .2s; margin-bottom:14px; }
.add-entry-btn:hover { background:rgba(77,159,255,.2); }
.balance-display { display:flex; justify-content:space-between; background:var(--card2); border:1px solid var(--border2); border-radius:10px; padding:12px 14px; margin-bottom:14px; }
.balance-ok   { color:var(--green); font-family:var(--mono); font-weight:700; }
.balance-fail { color:var(--danger); font-family:var(--mono); font-weight:700; }

/* SIDE PANEL */
.side-bg { position:fixed; inset:0; z-index:400; background:rgba(0,0,0,.5); display:none; }
.side-bg.open { display:block; }
.side-panel { position:fixed; right:0; top:0; bottom:0; width:460px; max-width:95vw; background:var(--card); border-left:1px solid var(--border2); display:flex; flex-direction:column; transform:translateX(100%); transition:transform .3s ease; z-index:401; }
.side-panel.open { transform:translateX(0); }
.side-header { padding:20px; border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; }
.side-header h3 { font-size:16px; font-weight:800; }
.close-btn { background:none; border:none; color:var(--muted); font-size:22px; cursor:pointer; padding:0; transition:color .15s; }
.close-btn:hover { color:var(--danger); }
.side-body { flex:1; overflow-y:auto; padding:16px 20px; }

.toast { position:fixed; bottom:22px; left:50%; transform:translateX(-50%) translateY(16px); background:var(--card2); border:1px solid var(--border2); border-radius:var(--r); padding:12px 20px; font-size:13px; font-weight:600; color:var(--text); box-shadow:0 20px 50px rgba(0,0,0,.5); opacity:0; pointer-events:none; transition:opacity .25s,transform .25s; z-index:999; }
.toast.show { opacity:1; transform:translateX(-50%) translateY(0); }
::-webkit-scrollbar { width:4px; }
::-webkit-scrollbar-thumb { background:var(--border2); border-radius:4px; }
</style>
    <script src="/static/auth-guard.js"></script>
</head>
<body>
""" + render_app_header(current_user, "page_accounting") + """

<div class="content">
    <div>
        <div class="page-title">Accounting</div>
        <div class="page-sub">Chart of accounts, journal entries and financial reports</div>
    </div>

    <!-- TABS -->
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">
        <div class="tabs">
            <button class="tab active" id="tab-accounts" onclick="switchTab('accounts')">Chart of Accounts</button>
            <button class="tab"        id="tab-journals" onclick="switchTab('journals')">Journal Entries</button>
            <button class="tab"        id="tab-pl"       onclick="switchTab('pl')">Profit & Loss</button>
            <button class="tab"        id="tab-tb"       onclick="switchTab('tb')">Trial Balance</button>
            <button class="tab"        id="tab-b2b"      onclick="switchTab('b2b')">B2B Clients</button>
        </div>
        <div style="display:flex;gap:10px;">
            <button class="btn btn-outline" id="btn-seed"   onclick="seedAccounts()" style="display:none">⚡ Setup Default Accounts</button>
            <button class="btn btn-green"   id="btn-add-acc" onclick="openAddAccModal()">+ Add Account</button>
            <button class="btn btn-blue"    id="btn-add-je"  onclick="openJEModal()" style="display:none">+ New Journal Entry</button>
        </div>
    </div>

    <!-- ACCOUNTS -->
    <div id="section-accounts">
        <div class="table-wrap">
            <table>
                <thead><tr><th>Code</th><th>Name</th><th>Type</th><th>Balance</th><th>Actions</th></tr></thead>
                <tbody id="accounts-body"><tr><td colspan="5" style="text-align:center;color:var(--muted);padding:40px">Loading…</td></tr></tbody>
            </table>
        </div>
    </div>

    <!-- JOURNALS -->
    <div id="section-journals" style="display:none">
        <div style="display:flex;align-items:end;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
            <div class="fld" style="margin:0;min-width:170px;">
                <label>From Date</label>
                <input type="date" id="journals-from-date">
            </div>
            <div class="fld" style="margin:0;min-width:170px;">
                <label>To Date</label>
                <input type="date" id="journals-to-date">
            </div>
            <button class="btn btn-outline" onclick="resetJournalFilters()">Clear</button>
            <div id="journals-active-range" style="font-size:12px;color:var(--muted);padding-bottom:10px"></div>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>ID</th><th>Type</th><th>Description</th><th>Entries</th><th>Total Debit</th><th>Date</th><th>Actions</th></tr></thead>
                <tbody id="journals-body"><tr><td colspan="7" style="text-align:center;color:var(--muted);padding:40px">Loading…</td></tr></tbody>
            </table>
        </div>
        <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-top:12px;">
            <div id="journals-pagination-summary" style="font-size:12px;color:var(--muted)">Loading entries...</div>
            <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <button class="btn btn-outline" id="journals-prev-btn" onclick="changeJournalPage(-1)">Previous</button>
                <div id="journals-page-indicator" style="font-size:12px;color:var(--sub);min-width:96px;text-align:center">Page 1 of 1</div>
                <button class="btn btn-outline" id="journals-next-btn" onclick="changeJournalPage(1)">Next</button>
            </div>
        </div>
    </div>
    <!-- P&L -->
    <div id="section-pl" style="display:none">
        <div style="display:flex;align-items:end;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
            <div class="fld" style="margin:0;min-width:170px;">
                <label>From Date</label>
                <input type="date" id="pl-from-date" onchange="loadPL()">
            </div>
            <div class="fld" style="margin:0;min-width:170px;">
                <label>To Date</label>
                <input type="date" id="pl-to-date" onchange="loadPL()">
            </div>
            <button class="btn btn-outline" onclick="resetPLFilters()">Clear</button>
        </div>
        <div id="pl-content"><div style="color:var(--muted);padding:40px;text-align:center">Loading…</div></div>
    </div>

    <!-- TRIAL BALANCE -->
    <div id="section-tb" style="display:none">
        <div style="display:flex;align-items:end;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
            <div class="fld" style="margin:0;min-width:170px;">
                <label>As Of</label>
                <input type="date" id="tb-as-of-date" onchange="loadTB()">
            </div>
            <button class="btn btn-outline" onclick="resetTBFilters()">Clear</button>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Code</th><th>Account</th><th>Type</th><th>Debit</th><th>Credit</th></tr></thead>
                <tbody id="tb-body"><tr><td colspan="5" style="text-align:center;color:var(--muted);padding:40px">Loading…</td></tr></tbody>
                <tfoot id="tb-foot"></tfoot>
            </table>
        </div>
    </div>

    <!-- B2B INVOICES -->
    <div id="section-b2b" style="display:none">
        <div class="section-card">
            <div class="section-card-title">B2B Clients</div>
            <div class="section-card-sub">Client-first accounting view using the same B2B client records as the main B2B page.</div>
        </div>
        <div class="subtabs">
            <button class="subtab active" id="b2b-subtab-clients" onclick="switchB2BSubtab('clients')">Clients</button>
            <button class="subtab" id="b2b-subtab-invoices" onclick="switchB2BSubtab('invoices')">Invoices</button>
        </div>

        <div id="b2b-clients-panel">
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
                <div class="fld" style="margin:0;min-width:250px;flex:1 1 260px;">
                    <label>Search Clients</label>
                    <input id="b2b-client-search" type="search" placeholder="Search client name, contact, or phone..." oninput="queueB2BClientSearch()" autocomplete="off">
                </div>
                <div class="fld" style="margin:0;min-width:170px;">
                    <label>Statement As Of</label>
                    <input id="b2b-client-statement-date" type="date">
                </div>
                <button class="btn btn-outline" onclick="resetB2BClientFilters()">Clear</button>
            </div>
            <div class="table-wrap">
                <table>
                    <thead><tr><th>Client</th><th>Contact</th><th>Phone</th><th>Terms</th><th>Invoices</th><th>Outstanding</th><th>Credit Limit</th><th>Discount</th><th>Actions</th></tr></thead>
                    <tbody id="b2b-clients-body"><tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">Loading…</td></tr></tbody>
                </table>
            </div>
        </div>

        <div id="b2b-invoices-panel" style="display:none">
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
                <div class="fld" style="margin:0;min-width:250px;flex:1 1 260px;">
                    <label>Search</label>
                    <input id="b2b-search" type="search" placeholder="Search client or invoice number..." oninput="queueB2BSearch()" autocomplete="off">
                </div>
                <select id="b2b-type-filter" onchange="loadB2BInvoices()" style="background:var(--card2);border:1px solid var(--border2);border-radius:var(--r);padding:9px 13px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;">
                    <option value="">All Types</option>
                    <option value="cash">💵 Cash</option>
                    <option value="full_payment">📋 Full Payment</option>
                    <option value="consignment">🔄 Consignment</option>
                </select>
                <select id="b2b-status-filter" onchange="loadB2BInvoices()" style="background:var(--card2);border:1px solid var(--border2);border-radius:var(--r);padding:9px 13px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;">
                    <option value="">All Statuses</option>
                    <option value="paid">Paid</option>
                    <option value="unpaid">Unpaid</option>
                    <option value="partial">Partial</option>
                </select>
                <div class="fld" style="margin:0;min-width:170px;">
                    <label>From Date</label>
                    <input type="date" id="b2b-from-date" onchange="loadB2BInvoices()">
                </div>
                <div class="fld" style="margin:0;min-width:170px;">
                    <label>To Date</label>
                    <input type="date" id="b2b-to-date" onchange="loadB2BInvoices()">
                </div>
                <button class="btn btn-outline" onclick="resetB2BFilters()">Clear</button>
            </div>
            <div class="table-wrap">
                <table>
                    <thead><tr><th>Invoice #</th><th>Client</th><th>Type</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th><th>Date</th><th>Actions</th></tr></thead>
                    <tbody id="b2b-invoices-body"><tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">Loading…</td></tr></tbody>
                </table>
            </div>

            <!-- CONSIGNMENT PAYMENTS SUB-SECTION -->
            <div id="cons-payment-section" style="display:none;margin-top:20px">
                <div style="font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--muted);margin-bottom:12px;display:flex;align-items:center;gap:8px;">
                    Consignment Payment History
                    <span style="flex:1;height:1px;background:linear-gradient(90deg,var(--border2),transparent)"></span>
                </div>
                <div id="cons-payment-list"></div>
            </div>
        </div>
    </div>
</div>

<!-- INVOICE DETAIL MODAL -->
<div class="modal-bg" id="inv-detail-modal">
    <div class="modal" style="width:520px">
        <div style="text-align:center;margin-bottom:16px">
            <img src="/static/Logo.png" style="height:120px;object-fit:contain;margin-bottom:6px">
            <div style="font-size:16px;font-weight:900;color:#2a7a2a">Habiba Organic Farm</div>
            <div style="font-size:12px;color:var(--muted);margin-top:2px" id="inv-detail-num"></div>
        </div>
        <div style="border-top:1px dashed var(--border2);margin:12px 0"></div>
        <div id="inv-detail-meta" style="margin-bottom:12px"></div>
        <div style="border-top:1px dashed var(--border2);margin:12px 0"></div>
        <table style="width:100%;border-collapse:collapse;margin-bottom:12px">
            <thead><tr>
                <th style="text-align:left;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Product</th>
                <th style="text-align:center;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">QTY</th>
                <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Price</th>
                <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Total</th>
            </tr></thead>
            <tbody id="inv-detail-items"></tbody>
        </table>
        <div style="border-top:1px dashed var(--border2);margin:12px 0"></div>
        <div id="inv-detail-totals"></div>
        <div style="display:flex;gap:8px;margin-top:16px;justify-content:flex-end">
            <button onclick="printInvDetail()" style="background:linear-gradient(135deg,#2a7a2a,#217346);color:white;border:none;padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;">🖨 Print</button>
            <button onclick="document.getElementById('inv-detail-modal').classList.remove('open')" style="background:var(--card2);border:1px solid var(--border2);color:var(--sub);padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;">Close</button>
        </div>
    </div>
</div>

<!-- COLLECT PAYMENT MODAL (cash / full_payment) -->
<div class="modal-bg" id="collect-modal">
    <div class="modal" style="width:420px">
        <div class="modal-title">Collect Payment</div>
        <div class="modal-sub" id="collect-modal-sub"></div>
        <div style="background:rgba(0,255,157,.06);border:1px solid rgba(0,255,157,.15);border-radius:10px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:var(--green)">
            Recording payment moves: <b>Deferred Revenue → Sales Revenue</b>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Confirm Invoice Number *</label>
            <input id="collect-inv-num" placeholder="e.g. B2B-00012"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--mono);font-size:14px;outline:none;width:100%">
            <span style="font-size:11px;color:var(--muted)">Type the invoice number to confirm you're collecting the right payment</span>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Amount *</label>
            <input id="collect-amount" type="number" placeholder="0.00" min="0.01" step="any"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--sans);font-size:14px;outline:none;width:100%">
        </div>
        <div style="display:flex;gap:10px;justify-content:flex-end">
            <button onclick="document.getElementById('collect-modal').classList.remove('open')" style="background:transparent;border:1px solid var(--border2);color:var(--sub);padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;">Cancel</button>
            <button onclick="saveCollect()" style="background:linear-gradient(135deg,#00ff9d,#00d4ff);border:none;color:#021a10;padding:12px 28px;border-radius:var(--r);font-family:var(--sans);font-size:14px;font-weight:800;cursor:pointer;letter-spacing:.3px;box-shadow:0 4px 20px rgba(0,255,157,.3);">✓ Confirm Payment</button>
        </div>
    </div>
</div>

<!-- CONSIGNMENT PAYMENT MODAL -->
<div class="modal-bg" id="cons-modal">
    <div class="modal" style="width:640px;max-width:94vw">
        <div class="modal-title">💰 Record Consignment Payment</div>
        <div class="modal-sub" id="cons-modal-sub" style="color:var(--muted);font-size:13px;margin-bottom:14px"></div>
        <div style="background:rgba(45,212,191,.06);border:1px solid rgba(45,212,191,.15);border-radius:10px;padding:10px 14px;margin-bottom:16px;font-size:12px;color:var(--teal)">
            Record the items this client reported sold for the month. The amount is the total of those items less the client's agreed discount, and is matched to them automatically.
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:16px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">For which month's sales? *</label>
            <select id="cons-month" style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--sans);font-size:14px;outline:none;width:100%">
                <option value="">General payment (no specific month)</option>
            </select>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:8px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Items sold — enter quantity for each</label>
        </div>
        <div style="border:1px solid var(--border2);border-radius:10px;overflow-y:auto;max-height:280px;margin-bottom:14px">
            <table style="width:100%;border-collapse:collapse">
                <thead>
                    <tr style="background:var(--card2);position:sticky;top:0;z-index:1">
                        <th style="text-align:left;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 10px">Product</th>
                        <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 10px">Unit Price</th>
                        <th style="text-align:center;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 10px;width:110px">Qty Sold</th>
                        <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 10px">Line Total</th>
                    </tr>
                </thead>
                <tbody id="cons-items-body"></tbody>
            </table>
        </div>
        <div style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:12px 16px;margin-bottom:16px">
            <div style="display:flex;align-items:center;justify-content:space-between;font-size:13px;padding:2px 0">
                <span style="color:var(--muted)">Items sold subtotal</span>
                <span class="mono" id="cons-subtotal-display" style="color:var(--text)">0.00</span>
            </div>
            <div id="cons-discount-row" style="display:none;align-items:center;justify-content:space-between;font-size:13px;padding:2px 0">
                <span style="color:var(--muted)">Client discount (<span id="cons-discount-pct">0</span>%)</span>
                <span class="mono" id="cons-discount-display" style="color:var(--danger)">-0.00</span>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;border-top:1px solid var(--border2);margin-top:8px;padding-top:10px">
                <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Amount Paid (after discount)</div>
                <div id="cons-amount-display" class="mono" style="font-size:22px;font-weight:900;color:var(--teal)">0.00 EGP</div>
            </div>
            <input id="cons-amount" type="hidden" value="0">
        </div>
        <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:8px">
            <button onclick="document.getElementById('cons-modal').classList.remove('open')"
                style="background:transparent;border:1px solid var(--border2);color:var(--sub);padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;">Cancel</button>
            <button onclick="saveConsPayment()"
                style="background:linear-gradient(135deg,#2dd4bf,#4d9fff);border:none;color:#001a18;padding:12px 28px;border-radius:var(--r);font-family:var(--sans);font-size:14px;font-weight:800;cursor:pointer;letter-spacing:.3px;box-shadow:0 4px 20px rgba(45,212,191,.35);">💰 Record Payment</button>
        </div>
    </div>
</div>

<!-- B2B REFUND MODAL -->
<div class="modal-bg" id="refund-modal">
    <div class="modal" style="width:640px">
        <div class="modal-title">Client Refund</div>
        <div class="modal-sub" id="refund-modal-sub"></div>
        <div style="background:rgba(255,181,71,.08);border:1px solid rgba(255,181,71,.18);border-radius:10px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:var(--warn)">
            Records a numbered refund, returns the goods to stock and credits the client's balance —
            the same refund the B2B page creates.
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Confirm Client Name *</label>
            <input id="refund-inv-num" placeholder="Type client name exactly"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--mono);font-size:14px;outline:none;width:100%">
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Returned Items *</label>
            <div id="refund-items"></div>
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-top:6px">
                <button onclick="addRefundItem()" style="background:transparent;border:1px dashed var(--border2);color:var(--sub);padding:8px 14px;border-radius:10px;font-family:var(--sans);font-size:12px;font-weight:700;cursor:pointer">+ Add item</button>
                <div style="font-family:var(--mono);font-size:14px;font-weight:800" id="refund-total">0.00 EGP</div>
            </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:16px">
            <label style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted)">Reason</label>
            <input id="refund-reason" placeholder="Returned items / supplier return note"
                style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:10px 12px;color:var(--text);font-family:var(--sans);font-size:14px;outline:none;width:100%">
        </div>
        <div style="display:flex;gap:10px;justify-content:flex-end">
            <button onclick="document.getElementById('refund-modal').classList.remove('open')" style="background:transparent;border:1px solid var(--border2);color:var(--sub);padding:10px 18px;border-radius:var(--r);font-family:var(--sans);font-size:13px;font-weight:700;cursor:pointer;">Cancel</button>
            <button onclick="saveRefund()" style="background:linear-gradient(135deg,#ffb547,#ff7a45);border:none;color:#241200;padding:12px 28px;border-radius:var(--r);font-family:var(--sans);font-size:14px;font-weight:800;cursor:pointer;letter-spacing:.3px;">Record Return</button>
        </div>
    </div>
</div>

<!-- ADD ACCOUNT MODAL -->
<div class="modal-bg" id="acc-modal">
    <div class="modal" style="width:420px">
        <div class="modal-title">Add Account</div>
        <div class="fld"><label>Account Code *</label><input id="ac-code" placeholder="e.g. 1010"></div>
        <div class="fld"><label>Account Name *</label><input id="ac-name" placeholder="e.g. Petty Cash"></div>
        <div class="fld"><label>Type *</label>
            <select id="ac-type">
                <option value="asset">Asset</option>
                <option value="liability">Liability</option>
                <option value="equity">Equity</option>
                <option value="revenue">Revenue</option>
                <option value="expense">Expense</option>
            </select>
        </div>
        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeAccModal()">Cancel</button>
            <button class="btn btn-green" onclick="saveAccount()">Add Account</button>
        </div>
    </div>
</div>

<!-- JOURNAL ENTRY MODAL -->
<div class="modal-bg" id="je-modal">
    <div class="modal">
        <div class="modal-title">New Journal Entry</div>
        <div class="fld"><label>Description</label><input id="je-desc" placeholder="e.g. Monthly rent payment"></div>

        <div style="display:grid;grid-template-columns:2fr 1fr 1fr 30px;gap:8px;margin-bottom:6px;">
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Account</span>
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Debit</span>
            <span style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:1px">Credit</span>
            <span></span>
        </div>

        <div id="je-entries"></div>
        <button class="add-entry-btn" onclick="addEntryRow()">+ Add Line</button>

        <div class="balance-display">
            <span style="color:var(--muted);font-size:13px;font-weight:600">Balance Check</span>
            <span id="balance-check" class="balance-ok">Debit 0.00 = Credit 0.00 ✓</span>
        </div>

        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeJEModal()">Cancel</button>
            <button class="btn btn-blue" onclick="saveJournal()">Post Journal Entry</button>
        </div>
    </div>
</div>

<!-- JOURNAL DETAIL SIDE PANEL -->
<div class="side-bg" id="side-bg" onclick="closeSide()"></div>
<div class="side-panel" id="side-panel">
    <div class="side-header">
        <h3 id="side-title">Journal Entry</h3>
        <button class="close-btn" onclick="closeSide()">×</button>
    </div>
    <div class="side-body" id="side-body"></div>
</div>

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
        currentUserRole = u.role || "";
        currentUserPermissions = new Set(
            (typeof u.permissions === "string" ? u.permissions.split(",") : (u.permissions || []))
                .map(v => String(v).trim())
                .filter(Boolean)
        );
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
  function hasPermission(permission, u){
      const role = u ? (u.role || "") : currentUserRole;
      const perms = u
          ? new Set(typeof u.permissions === "string" ? u.permissions.split(",").map(v => v.trim()).filter(Boolean) : (u.permissions || []))
          : currentUserPermissions;
      return role === "admin" || perms.has(permission);
  }
  function configureAccountingPermissions(u){
      const tabMap = [
          {id:"tab-journals", permission:"tab_accounting_journal"},
          {id:"tab-pl", permission:"tab_accounting_pl"},
          {id:"tab-b2b", permission:"tab_accounting_b2b"},
      ];
      tabMap.forEach(conf => {
          let el = document.getElementById(conf.id);
          if(el && !hasPermission(conf.permission, u)) el.style.display = "none";
      });
      if(!hasPermission("action_accounting_post_journal", u)){
          let btn = document.getElementById("btn-add-je");
          if(btn) btn.style.display = "none";
      }
  }
  initializeColorMode();
  initUser().then(u => { if(u) configureAccountingPermissions(u); });
  let accounts    = [];
let currentTab  = "accounts";
let currentUserRole = "";
let currentUserPermissions = new Set();

function formatLocalDateInputValue(d){
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function todayIso(){
    return formatLocalDateInputValue(new Date());
}

function monthStartIso(){
    const d = new Date();
    d.setDate(1);
    return formatLocalDateInputValue(d);
}

function setDefaultAccountingFilters(){
    const today = todayIso();
    const monthStart = monthStartIso();
    const defaults = [
        ["journals-from-date", monthStart],
        ["journals-to-date", today],
        ["pl-from-date", monthStart],
        ["pl-to-date", today],
        ["tb-as-of-date", today],
        ["b2b-from-date", monthStart],
        ["b2b-to-date", today],
    ];
    defaults.forEach(([id, value]) => {
        const el = document.getElementById(id);
        if(el && !el.value) el.value = value;
    });
}

function appendDateRangeParams(params, fromId, toId){
    const fromValue = document.getElementById(fromId)?.value || "";
    const toValue = document.getElementById(toId)?.value || "";
    if(fromValue) params.set("from_date", fromValue);
    if(toValue) params.set("to_date", toValue);
}

function appendAsOfParam(params, id){
    const value = document.getElementById(id)?.value || "";
    if(value) params.set("as_of", value);
}

async function init(){
    setDefaultAccountingFilters();
    setupJournalFilters();
    const {fromDate, toDate} = getJournalFilterValues();
    updateJournalsActiveRange(fromDate, toDate);
    updateJournalPaginationUi();
    await loadAccounts();
}

/* ── TABS ── */
function switchTab(tab){
    const required = {
        journals: "tab_accounting_journal",
        pl: "tab_accounting_pl",
        b2b: "tab_accounting_b2b",
    };
    if(required[tab] && !hasPermission(required[tab])) return;
    currentTab = tab;
    ["accounts","journals","pl","tb","b2b"].forEach(t=>{
        document.getElementById("section-"+t).style.display = t===tab?"":"none";
        document.getElementById("tab-"+t).classList.toggle("active", t===tab);
    });
    document.getElementById("btn-add-acc").style.display  = tab==="accounts"?"":"none";
    document.getElementById("btn-seed").style.display     = tab==="accounts"?"":"none";
    document.getElementById("btn-add-je").style.display   = tab==="journals" && hasPermission("action_accounting_post_journal")?"":"none";

    if(tab==="journals") loadJournals();
    if(tab==="pl")       loadPL();
    if(tab==="tb")       loadTB();
    if(tab==="b2b")      switchB2BSubtab("clients");
}

/* ── ACCOUNTS ── */
async function loadAccounts(){
    accounts = await (await fetch("/accounting/api/accounts")).json();
    if(!accounts.length){
        document.getElementById("accounts-body").innerHTML =
            `<tr><td colspan="5" style="text-align:center;color:var(--muted);padding:40px">
                No accounts yet. Click "⚡ Setup Default Accounts" to get started.
            </td></tr>`;
        document.getElementById("btn-seed").style.display="";
        return;
    }
    document.getElementById("btn-seed").style.display="none";
    document.getElementById("accounts-body").innerHTML = accounts.map(a=>`
        <tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--blue)">${a.code}</td>
            <td class="name">${a.name}</td>
            <td><span class="type-badge type-${a.type}">${a.type}</span></td>
            <td class="${a.balance>=0?'dr':'cr'}">${Math.abs(a.balance).toFixed(2)}</td>
            <td><button class="action-btn danger" onclick="deleteAccount(${a.id},'${a.name.replace(/'/g,"\\'")}')">Delete</button></td>
        </tr>`).join("");
}

async function seedAccounts(){
    let res  = await fetch("/accounting/api/accounts/seed",{method:"POST"});
    let data = await res.json();
    showToast(data.message);
    loadAccounts();
}

function openAddAccModal(){ document.getElementById("acc-modal").classList.add("open"); }
function closeAccModal()  { document.getElementById("acc-modal").classList.remove("open"); }

async function saveAccount(){
    let code = document.getElementById("ac-code").value.trim();
    let name = document.getElementById("ac-name").value.trim();
    if(!code||!name){ showToast("Code and name are required"); return; }
    let res  = await fetch("/accounting/api/accounts",{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({code, name, type:document.getElementById("ac-type").value}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    closeAccModal();
    showToast("Account added ✓");
    loadAccounts();
}

async function deleteAccount(id,name){
    if(!confirm(`Delete account "${name}"?`)) return;
    let res  = await fetch(`/accounting/api/accounts/${id}`,{method:"DELETE"});
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    showToast("Account deleted ✓");
    loadAccounts();
}

/* ── JOURNALS ── */
let journalsRequestSeq = 0;
let journalsAbortController = null;
let journalsPage = 1;
let journalsPageSize = 50;
let journalsTotal = 0;
let journalsTotalPages = 1;

function setJournalsTableState(message, color="var(--muted)"){
    document.getElementById("journals-body").innerHTML =
        `<tr><td colspan="7" style="text-align:center;color:${color};padding:40px">${message}</td></tr>`;
}

function setJournalsUnauthorizedState(message="Your session expired or you do not have access to Journal Entries. Please sign in again."){
    setJournalsTableState(message, "var(--danger)");
}

function updateJournalPaginationUi(){
    const summaryEl = document.getElementById("journals-pagination-summary");
    const indicatorEl = document.getElementById("journals-page-indicator");
    const prevBtn = document.getElementById("journals-prev-btn");
    const nextBtn = document.getElementById("journals-next-btn");
    if(summaryEl){
        if(!journalsTotal){
            summaryEl.textContent = "0 total entries";
        }else{
            const startRow = ((journalsPage - 1) * journalsPageSize) + 1;
            const endRow = Math.min(journalsPage * journalsPageSize, journalsTotal);
            summaryEl.textContent = `Showing ${startRow}-${endRow} of ${journalsTotal} entries`;
        }
    }
    if(indicatorEl){
        indicatorEl.textContent = `Page ${journalsPage} of ${journalsTotalPages}`;
    }
    if(prevBtn) prevBtn.disabled = journalsPage <= 1;
    if(nextBtn) nextBtn.disabled = journalsPage >= journalsTotalPages || journalsTotal === 0;
}

function getJournalFilterValues(){
    return {
        fromDate: document.getElementById("journals-from-date")?.value || "",
        toDate: document.getElementById("journals-to-date")?.value || "",
    };
}

function updateJournalsActiveRange(fromDate, toDate){
    const el = document.getElementById("journals-active-range");
    if(!el) return;
    if(fromDate && toDate){
        el.textContent = `Applied range: ${fromDate} to ${toDate}`;
        return;
    }
    if(fromDate){
        el.textContent = `Applied range: from ${fromDate}`;
        return;
    }
    if(toDate){
        el.textContent = `Applied range: up to ${toDate}`;
        return;
    }
    el.textContent = "Applied range: all dates";
}

function debugJournalFilters(stage, details){
    console.debug("[Accounting][Journals]", stage, details);
}

async function tryJournalSessionRefresh(){
    const rawFetch = window._origFetch || window.fetch;
    const refreshRes = await rawFetch("/auth/refresh", {
        method: "POST",
        credentials: "same-origin",
    });
    debugJournalFilters("refresh-attempt", {ok: refreshRes.ok, status: refreshRes.status});
    return refreshRes.ok;
}

async function fetchJournalsWithAuth(url, signal){
    let res = await fetch(url, {
        cache: "no-store",
        credentials: "same-origin",
        signal,
    });
    if(res.status !== 401) return res;

    debugJournalFilters("unauthorized-response", {url, status: res.status});
    const refreshed = await tryJournalSessionRefresh();
    if(!refreshed) return res;

    const retryRes = await fetch(url, {
        cache: "no-store",
        credentials: "same-origin",
        signal,
    });
    debugJournalFilters("retry-response", {url, status: retryRes.status});
    return retryRes;
}

function handleJournalFilterChange(){
    const {fromDate, toDate} = getJournalFilterValues();
    debugJournalFilters("ui-change", {fromDate, toDate});
    journalsPage = 1;
    loadJournals();
}

function setupJournalFilters(){
    ["journals-from-date", "journals-to-date"].forEach(id => {
        const el = document.getElementById(id);
        if(!el || el.dataset.bound === "1") return;
        el.addEventListener("change", handleJournalFilterChange);
        el.addEventListener("input", handleJournalFilterChange);
        el.dataset.bound = "1";
    });
}

async function loadJournals(){
    const requestSeq = ++journalsRequestSeq;
    if(journalsAbortController) journalsAbortController.abort();
    journalsAbortController = new AbortController();
    const params = new URLSearchParams();
    appendDateRangeParams(params, "journals-from-date", "journals-to-date");
    params.set("page", String(journalsPage));
    params.set("page_size", String(journalsPageSize));
    const {fromDate, toDate} = getJournalFilterValues();
    updateJournalsActiveRange(fromDate, toDate);
    debugJournalFilters("request", {
        requestSeq,
        fromDate,
        toDate,
        page: journalsPage,
        pageSize: journalsPageSize,
        url: `/accounting/api/journals?${params.toString()}`,
    });
    setJournalsTableState("Loading...");
    const summaryEl = document.getElementById("journals-pagination-summary");
    const indicatorEl = document.getElementById("journals-page-indicator");
    if(summaryEl) summaryEl.textContent = "Loading entries...";
    if(indicatorEl) indicatorEl.textContent = `Page ${journalsPage} of ${journalsTotalPages}`;
    try{
        let res = await fetchJournalsWithAuth(
            `/accounting/api/journals?${params.toString()}`,
            journalsAbortController.signal,
        );
        let data = await res.json();
        if(requestSeq !== journalsRequestSeq) return;
        debugJournalFilters("response", {
            requestSeq,
            fromDate: data.from_date || null,
            toDate: data.to_date || null,
            total: data.total,
            page: data.page,
            pageSize: data.page_size,
            totalPages: data.total_pages,
            rowCount: data.journals?.length || 0,
        });
        if(!res.ok){
            if(res.status === 401){
                showToast("Session expired. Please sign in again.");
                setJournalsUnauthorizedState();
                return;
            }
            if(res.status === 403){
                showToast("You do not have permission to view journal entries.");
                setJournalsUnauthorizedState("You do not have permission to view Journal Entries.");
                return;
            }
            showToast("Error: " + (data.detail || "Unable to load journal entries"));
            setJournalsTableState(data.detail || "Unable to load journal entries", "var(--danger)");
            return;
        }
        updateJournalsActiveRange(data.from_date || "", data.to_date || "");
        journalsTotal = Number(data.total || 0);
        journalsPage = Math.max(1, Number(data.page || 1));
        journalsPageSize = Math.max(1, Number(data.page_size || journalsPageSize));
        journalsTotalPages = Math.max(1, Number(data.total_pages || 1));
        updateJournalPaginationUi();
        if(!data.journals.length){
            debugJournalFilters("render-empty", {requestSeq});
            setJournalsTableState("No journal entries found for the selected date range.");
            return;
        }
        const renderedRows = data.journals.map(j=>{
            const isRefund = j.ref_type === "retail_refund" || j.ref_type === "retail_refund_void";
            const badgeClass = isRefund ? "type-refund"
                : j.ref_type === "manual" ? "type-equity"
                : j.ref_type.includes("b2b_refund") ? "type-refund"
                : "type-revenue";
            const rowStyle = isRefund ? 'style="background:rgba(255,77,109,.03);"' : '';
            const amtColor = isRefund ? "var(--danger)" : "var(--green)";
            const amtPrefix = isRefund ? "−" : "";
            return `<tr ${rowStyle}>
                <td style="font-family:var(--mono);color:var(--muted);font-size:12px">#${j.id}</td>
                <td><span class="type-badge ${badgeClass}">${j.ref_type}</span></td>
                <td class="name">${j.description}</td>
                <td style="color:var(--sub)">${j.entries_count} lines</td>
                <td class="dr" style="color:${amtColor}">${amtPrefix}${j.total_debit.toFixed(2)}</td>
                <td style="font-size:12px;color:var(--muted)">${j.created_at}</td>
                <td><button class="action-btn green" onclick="viewJournal(${j.id})">View</button></td>
            </tr>`;
        }).join("");
        document.getElementById("journals-body").innerHTML = renderedRows;
        debugJournalFilters("render-complete", {
            requestSeq,
            renderedCount: data.journals.length,
            firstJournalId: data.journals[0]?.id || null,
        });
    }catch(_err){
        if(_err?.name === "AbortError"){
            debugJournalFilters("request-aborted", {requestSeq});
            return;
        }
        if(requestSeq !== journalsRequestSeq) return;
        debugJournalFilters("request-failed", {
            requestSeq,
            message: _err?.message || "Unknown error",
        });
        showToast("Error: Unable to load journal entries");
        setJournalsTableState("Unable to load journal entries", "var(--danger)");
    }
}

function resetJournalFilters(){
    document.getElementById("journals-from-date").value = monthStartIso();
    document.getElementById("journals-to-date").value = todayIso();
    journalsPage = 1;
    loadJournals();
}

function changeJournalPage(direction){
    const nextPage = journalsPage + direction;
    if(nextPage < 1 || nextPage > journalsTotalPages) return;
    journalsPage = nextPage;
    loadJournals();
}

function openJEModal(){
    document.getElementById("je-desc").value="";
    document.getElementById("je-entries").innerHTML="";
    addEntryRow(); addEntryRow();
    updateBalanceCheck();
    document.getElementById("je-modal").classList.add("open");
}
function closeJEModal(){ document.getElementById("je-modal").classList.remove("open"); }

function addEntryRow(){
    let div = document.createElement("div");
    div.className = "entry-row";
    div.innerHTML = `
        <select onchange="updateBalanceCheck()">
            <option value="">Select account…</option>
            ${accounts.map(a=>`<option value="${a.id}">${a.code} — ${a.name}</option>`).join("")}
        </select>
        <input type="number" placeholder="0.00" min="0" step="any" oninput="updateBalanceCheck()">
        <input type="number" placeholder="0.00" min="0" step="any" oninput="updateBalanceCheck()">
        <button class="rm-btn" onclick="this.parentElement.remove();updateBalanceCheck()">×</button>
    `;
    document.getElementById("je-entries").appendChild(div);
}

function updateBalanceCheck(){
    let rows   = document.querySelectorAll("#je-entries .entry-row");
    let totalD = 0, totalC = 0;
    rows.forEach(row=>{
        let inputs = row.querySelectorAll("input");
        totalD += parseFloat(inputs[0].value)||0;
        totalC += parseFloat(inputs[1].value)||0;
    });
    let el = document.getElementById("balance-check");
    let ok = Math.abs(totalD-totalC) < 0.01;
    el.className = ok ? "balance-ok" : "balance-fail";
    el.innerText  = ok
        ? `Debit ${totalD.toFixed(2)} = Credit ${totalC.toFixed(2)} ✓`
        : `Debit ${totalD.toFixed(2)} ≠ Credit ${totalC.toFixed(2)} ✗`;
}

async function saveJournal(){
    let desc = document.getElementById("je-desc").value.trim();
    let rows = document.querySelectorAll("#je-entries .entry-row");
    let entries = [];
    for(let row of rows){
        let acc_id = parseInt(row.querySelector("select").value);
        let debit  = parseFloat(row.querySelectorAll("input")[0].value)||0;
        let credit = parseFloat(row.querySelectorAll("input")[1].value)||0;
        if(!acc_id) continue;
        entries.push({account_id:acc_id, debit, credit});
    }
    if(!entries.length){ showToast("Add at least one entry"); return; }

    let res  = await fetch("/accounting/api/journals",{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({description:desc, entries}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    closeJEModal();
    showToast("Journal entry posted ✓");
    loadJournals(); loadAccounts();
}

async function viewJournal(id){
    document.getElementById("side-body").innerHTML=`<div style="color:var(--muted);font-size:13px">Loading…</div>`;
    document.getElementById("side-bg").classList.add("open");
    document.getElementById("side-panel").classList.add("open");
    let j = await (await fetch(`/accounting/api/journals/${id}`)).json();
    document.getElementById("side-title").innerText = `Journal #${j.id}`;
    document.getElementById("side-body").innerHTML = `
        <div style="display:flex;flex-direction:column;gap:14px">
            <div style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:14px">
                <div style="display:flex;justify-content:space-between;margin-bottom:6px">
                    <span style="color:var(--muted);font-size:12px">Type</span>
                    <span style="font-weight:700;text-transform:capitalize">${j.ref_type}</span>
                </div>
                <div style="display:flex;justify-content:space-between;margin-bottom:6px">
                    <span style="color:var(--muted);font-size:12px">Description</span>
                    <span style="font-size:13px">${j.description}</span>
                </div>
                <div style="display:flex;justify-content:space-between">
                    <span style="color:var(--muted);font-size:12px">Date</span>
                    <span style="font-size:12px">${j.created_at}</span>
                </div>
            </div>
            <table style="width:100%;border-collapse:collapse">
                <thead><tr>
                    <th style="text-align:left;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Account</th>
                    <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Debit</th>
                    <th style="text-align:right;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--muted);padding:8px 0">Credit</th>
                </tr></thead>
                <tbody>
                ${j.entries.map(e=>`
                    <tr>
                        <td style="padding:9px 0;border-top:1px solid var(--border);font-size:13px">
                            <div style="color:var(--text);font-weight:600">${e.account_name}</div>
                            <div style="font-family:var(--mono);font-size:10px;color:var(--muted)">${e.account_code}</div>
                        </td>
                        <td style="padding:9px 0;border-top:1px solid var(--border);text-align:right;font-family:var(--mono);color:var(--green)">${e.debit>0?e.debit.toFixed(2):""}</td>
                        <td style="padding:9px 0;border-top:1px solid var(--border);text-align:right;font-family:var(--mono);color:var(--blue)">${e.credit>0?e.credit.toFixed(2):""}</td>
                    </tr>`).join("")}
                </tbody>
            </table>
        </div>`;
}

function closeSide(){
    document.getElementById("side-bg").classList.remove("open");
    document.getElementById("side-panel").classList.remove("open");
}

/* ── P&L ── */
async function loadPL(){
    const params = new URLSearchParams();
    appendDateRangeParams(params, "pl-from-date", "pl-to-date");
    let res = await fetch(`/accounting/api/profit-loss?${params.toString()}`);
    let d = await res.json();
    if(!res.ok){
        showToast("Error: " + (d.detail || "Unable to load Profit & Loss"));
        document.getElementById("pl-content").innerHTML = `<div style="color:var(--danger);padding:40px;text-align:center">${d.detail || "Unable to load Profit & Loss"}</div>`;
        return;
    }
    let profitColor = d.net_profit>=0?"var(--green)":"var(--danger)";
    const refundLine = d.total_refunds > 0
        ? `<div class="pl-row" style="background:rgba(255,77,109,.04);border-left:3px solid rgba(255,77,109,.4);">
               <span style="color:var(--danger)">↩ Refunds Deducted (${d.refund_count} refunds — already deducted from net revenue)</span>
               <span style="font-family:var(--mono);color:var(--danger)">−${d.total_refunds.toFixed(2)}</span>
           </div>`
        : "";
    document.getElementById("pl-content").innerHTML = `
        <div class="pl-section">
            <div class="pl-header">Revenue</div>
            ${d.revenues.map(r=>`
                <div class="pl-row">
                    <span style="color:var(--sub)">${r.code} — ${r.name}</span>
                    <span style="font-family:var(--mono);color:var(--green)">${Math.abs(r.amount).toFixed(2)}</span>
                </div>`).join("") || `<div class="pl-row" style="color:var(--muted)">No revenue recorded yet</div>`}
            ${refundLine}
            <div class="pl-total">
                <span>Net Revenue (after refunds)</span>
                <span style="font-family:var(--mono);color:var(--green)">${Math.abs(d.total_revenue).toFixed(2)}</span>
            </div>
        </div>

        <div class="pl-section">
            <div class="pl-header">Expenses</div>
            ${d.expenses.map(e=>`
                <div class="pl-row">
                    <span style="color:var(--sub)">${e.code} — ${e.name}</span>
                    <span style="font-family:var(--mono);color:var(--warn)">${Math.abs(e.amount).toFixed(2)}</span>
                </div>`).join("") || `<div class="pl-row" style="color:var(--muted)">No expenses recorded yet</div>`}
            <div class="pl-total">
                <span>Total Expenses</span>
                <span style="font-family:var(--mono);color:var(--warn)">${Math.abs(d.total_expense).toFixed(2)}</span>
            </div>
        </div>

        <div class="pl-net">
            <span>${d.net_profit>=0?"Net Profit":"Net Loss"}</span>
            <span style="font-family:var(--mono);color:${profitColor}">${Math.abs(d.net_profit).toFixed(2)}</span>
        </div>`;
}

function resetPLFilters(){
    document.getElementById("pl-from-date").value = monthStartIso();
    document.getElementById("pl-to-date").value = todayIso();
    loadPL();
}

/* ── TRIAL BALANCE ── */
async function loadTB(){
    const params = new URLSearchParams();
    appendAsOfParam(params, "tb-as-of-date");
    let res = await fetch(`/accounting/api/trial-balance?${params.toString()}`);
    let d = await res.json();
    if(!res.ok){
        showToast("Error: " + (d.detail || "Unable to load Trial Balance"));
        document.getElementById("tb-body").innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--danger);padding:40px">${d.detail || "Unable to load Trial Balance"}</td></tr>`;
        document.getElementById("tb-foot").innerHTML = "";
        return;
    }
    document.getElementById("tb-body").innerHTML = d.rows.map(r=>`
        <tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--blue)">${r.code}</td>
            <td class="name">${r.name}</td>
            <td><span class="type-badge type-${r.type}">${r.type}</span></td>
            <td class="dr">${r.debit>0?r.debit.toFixed(2):""}</td>
            <td class="cr">${r.credit>0?r.credit.toFixed(2):""}</td>
        </tr>`).join("");
    let balanced = Math.abs(d.total_debit-d.total_credit)<0.01;
    document.getElementById("tb-foot").innerHTML = `
        <tr style="background:var(--card2)">
            <td colspan="3" style="padding:12px 16px;font-weight:800;color:var(--sub)">
                Total ${balanced?"✓ Balanced":"✗ Not Balanced"}
            </td>
            <td style="padding:12px 16px;font-family:var(--mono);font-size:14px;font-weight:800;color:var(--green)">${d.total_debit.toFixed(2)}</td>
            <td style="padding:12px 16px;font-family:var(--mono);font-size:14px;font-weight:800;color:var(--blue)">${d.total_credit.toFixed(2)}</td>
        </tr>
        ${d.drift_count ? `<tr style="background:rgba(255,181,71,.08)">
            <td colspan="5" style="padding:10px 16px;color:var(--warn);font-size:12px">
                ${d.drift_count} account${d.drift_count===1?"":"s"} have stored balance drift versus journal-derived balance.
            </td>
        </tr>` : ""}`;
}

function resetTBFilters(){
    document.getElementById("tb-as-of-date").value = todayIso();
    loadTB();
}

["acc-modal","je-modal"].forEach(id=>{
    document.getElementById(id).addEventListener("click",function(e){ if(e.target===this) this.classList.remove("open"); });
});

let toastTimer=null;
function showToast(msg){
    let t=document.getElementById("toast");
    t.innerText=msg; t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer=setTimeout(()=>t.classList.remove("show"),3500);
}

/* ── B2B INVOICES ── */
let allB2BInvoices  = [];
let allB2BClients   = [];
let collectInvoiceId = null;
let consClientId     = null;
let consClientName   = null;
let consItems        = [];
let consDiscountPct  = 0;
let currentInvDetail = null;
let refundInvoiceId  = null;
let refundInvoiceNum = null;
let b2bSearchTimer   = null;
let b2bClientSearchTimer = null;
let currentB2BSubtab = "clients";

function switchB2BSubtab(subtab){
    currentB2BSubtab = subtab === "invoices" ? "invoices" : "clients";
    document.getElementById("b2b-subtab-clients").classList.toggle("active", currentB2BSubtab === "clients");
    document.getElementById("b2b-subtab-invoices").classList.toggle("active", currentB2BSubtab === "invoices");
    document.getElementById("b2b-clients-panel").style.display = currentB2BSubtab === "clients" ? "" : "none";
    document.getElementById("b2b-invoices-panel").style.display = currentB2BSubtab === "invoices" ? "" : "none";
    if(currentB2BSubtab === "clients") loadB2BClients();
    else loadB2BInvoices();
}

function getB2BClientSearchValue(){
    return (document.getElementById("b2b-client-search")?.value || "").trim().replace(/\\s+/g, " ");
}

function getB2BClientStatementDate(){
    const input = document.getElementById("b2b-client-statement-date");
    if(!input) return "";
    if(!input.value) input.value = todayIso();
    return input.value;
}

function openB2BClientStatement(clientId){
    const params = new URLSearchParams();
    const asOf = getB2BClientStatementDate();
    if(asOf) params.set("as_of", asOf);
    const qs = params.toString();
    window.open(`/b2b/client/${clientId}/statement${qs ? `?${qs}` : ""}`, "_blank", "noopener");
}

function queueB2BClientSearch(){
    clearTimeout(b2bClientSearchTimer);
    document.getElementById("b2b-clients-body").innerHTML =
        `<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:40px">Searching clients…</td></tr>`;
    b2bClientSearchTimer = setTimeout(() => loadB2BClients(), 180);
}

async function loadB2BClients(){
    const search = getB2BClientSearchValue();
    const params = new URLSearchParams();
    if(search) params.set("q", search);
    let res = await fetch(`/accounting/api/b2b-clients?${params.toString()}`);
    let data = await res.json();
    if(!res.ok){
        showToast("Error: " + (data.detail || "Unable to load B2B clients"));
        document.getElementById("b2b-clients-body").innerHTML =
            `<tr><td colspan="8" style="text-align:center;color:var(--danger);padding:40px">${data.detail || "Unable to load B2B clients"}</td></tr>`;
        allB2BClients = [];
        return;
    }
    allB2BClients = data;
    renderB2BClients(allB2BClients);
}

function resetB2BClientFilters(){
    clearTimeout(b2bClientSearchTimer);
    document.getElementById("b2b-client-search").value = "";
    document.getElementById("b2b-client-statement-date").value = todayIso();
    loadB2BClients();
}

function renderB2BClients(clients){
    if(!clients.length){
        const search = getB2BClientSearchValue();
        document.getElementById("b2b-clients-body").innerHTML =
            `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">${search ? "No B2B clients matched your search" : "No B2B clients found"}</td></tr>`;
        return;
    }
    document.getElementById("b2b-clients-body").innerHTML = clients.map(client => `
        <tr>
            <td style="color:var(--text);font-weight:700">
                ${client.name}
                ${client.is_consignment ? `<div style="font-size:10px;color:var(--teal);font-weight:700;letter-spacing:.5px;margin-top:3px">Consignment Client</div>` : ``}
            </td>
            <td style="font-size:12px;color:var(--sub)">${client.contact_person}</td>
            <td class="mono" style="font-size:12px;color:var(--muted)">${client.phone}</td>
            <td><span style="font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;background:${client.is_consignment?"rgba(45,212,191,.1)":"rgba(77,159,255,.1)"};color:${client.is_consignment?"var(--teal)":"var(--blue)"}">${(client.payment_terms || "—").replace(/_/g," ")}</span></td>
            <td class="mono" style="font-size:12px">${client.invoice_count}</td>
            <td class="mono" style="font-size:12px;color:${client.outstanding>0?"var(--warn)":"var(--muted)"};font-weight:${client.outstanding>0?"700":"400"}">${client.outstanding>0?client.outstanding.toFixed(2):"—"}</td>
            <td class="mono" style="font-size:12px">${client.credit_limit.toFixed(2)}</td>
            <td class="mono" style="font-size:12px">${client.discount_pct.toFixed(2)}%</td>
            <td>
                <div style="display:flex;gap:5px;flex-wrap:wrap">
                    <button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                        onmouseenter="this.style.borderColor='var(--blue)';this.style.color='var(--blue)'"
                        onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                        onclick="openB2BClientStatement(${client.id})">Statement</button>
                    ${client.is_consignment && client.outstanding > 0.01
                        ? `<button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                            onmouseenter="this.style.borderColor='var(--teal)';this.style.color='var(--teal)'"
                            onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                            onclick="openConsModal(${client.id},'${client.name.replace(/'/g,"\\'")}',${client.outstanding},${client.discount_pct||0})">💰 Record Payment</button>`
                        : ""}
                    ${client.outstanding > 0.01
                        ? `<button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                            onmouseenter="this.style.borderColor='var(--danger)';this.style.color='var(--danger)'"
                            onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                            onclick="openRefundModal(${client.id},'${client.name.replace(/'/g,"\\'")}',${client.outstanding})">Client Refund</button>`
                        : ""}
                </div>
            </td>
        </tr>`).join("");
}

function getB2BSearchValue(){
    return (document.getElementById("b2b-search")?.value || "").trim().replace(/\\s+/g, " ");
}

function queueB2BSearch(){
    clearTimeout(b2bSearchTimer);
    document.getElementById("b2b-invoices-body").innerHTML =
        `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">Searching invoices…</td></tr>`;
    b2bSearchTimer = setTimeout(() => loadB2BInvoices(), 180);
}

async function loadB2BInvoices(){
    let type   = document.getElementById("b2b-type-filter").value;
    let status = document.getElementById("b2b-status-filter").value;
    let search = getB2BSearchValue();
    const params = new URLSearchParams();
    if(type) params.set("invoice_type", type);
    if(status) params.set("status", status);
    if(search) params.set("search", search);
    appendDateRangeParams(params, "b2b-from-date", "b2b-to-date");
    let res = await fetch(`/accounting/api/b2b-invoices?${params.toString()}`);
    let data = await res.json();
    if(!res.ok){
        showToast("Error: " + (data.detail || "Unable to load B2B invoices"));
        document.getElementById("b2b-invoices-body").innerHTML =
            `<tr><td colspan="9" style="text-align:center;color:var(--danger);padding:40px">${data.detail || "Unable to load B2B invoices"}</td></tr>`;
        allB2BInvoices = [];
        return;
    }
    allB2BInvoices = data;
    renderB2BInvoices(allB2BInvoices);

    // Show consignment payment history section if filtering consignment
    let showCons = type==="consignment";
    document.getElementById("cons-payment-section").style.display = showCons?"":"none";
    if(showCons){ loadConsPaymentHistory(); }
}

function resetB2BFilters(){
    clearTimeout(b2bSearchTimer);
    document.getElementById("b2b-search").value = "";
    document.getElementById("b2b-type-filter").value = "";
    document.getElementById("b2b-status-filter").value = "";
    document.getElementById("b2b-from-date").value = monthStartIso();
    document.getElementById("b2b-to-date").value = todayIso();
    loadB2BInvoices();
}

function renderB2BInvoices(invoices){
    if(!invoices.length){
        const search = getB2BSearchValue();
        document.getElementById("b2b-invoices-body").innerHTML =
            `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:40px">${search ? "No B2B invoices matched your search" : "No invoices found"}</td></tr>`;
        return;
    }
    const typeLabel = {cash:"💵 Cash", full_payment:"📋 Full Payment", consignment:"🔄 Consignment"};
    const typeBadge = {cash:"badge-cash", full_payment:"badge-full_payment", consignment:"badge-consignment"};
    const statusColor = {paid:"var(--green)", unpaid:"var(--warn)", partial:"var(--blue)"};

    document.getElementById("b2b-invoices-body").innerHTML = invoices.map(i=>{
        let isCons   = i.invoice_type === "consignment";
        let isPaid   = i.status === "paid";
        let hasBalance = i.balance_due > 0.01;

        let actions = `<div style="display:flex;gap:5px;flex-wrap:wrap">
            <button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                onmouseenter="this.style.borderColor='var(--blue)';this.style.color='var(--blue)'"
                onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                onclick="openInvDetail(${i.id})">View</button>
            ${!isCons && hasBalance
                ? `<button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                    onmouseenter="this.style.borderColor='var(--warn)';this.style.color='var(--warn)'"
                    onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                    onclick="openCollectModal(${i.id},'${i.invoice_number}',${i.balance_due})">💵 Collect</button>`
                : ""}
            ${i.client_outstanding > 0.01
                ? `<button style="background:transparent;border:1px solid var(--border2);color:var(--sub);font-size:12px;font-weight:600;padding:5px 10px;border-radius:7px;cursor:pointer;font-family:var(--sans);"
                    onmouseenter="this.style.borderColor='var(--danger)';this.style.color='var(--danger)'"
                    onmouseleave="this.style.borderColor='var(--border2)';this.style.color='var(--sub)'"
                    onclick="openRefundModal(${i.client_id},'${i.client.replace(/'/g,"\\'")}',${i.client_outstanding})">Client Refund</button>`
                : ""}
        </div>`;

        return `<tr>
            <td style="font-family:var(--mono);font-size:12px;color:var(--blue)">${i.invoice_number}</td>
            <td style="color:var(--text);font-weight:600">${i.client}</td>
            <td><span style="font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;background:${isCons?"rgba(45,212,191,.1)":i.invoice_type==="cash"?"rgba(0,255,157,.1)":"rgba(77,159,255,.1)"};color:${isCons?"var(--teal)":i.invoice_type==="cash"?"var(--green)":"var(--blue)"}">${typeLabel[i.invoice_type]||i.invoice_type}</span></td>
            <td style="font-family:var(--mono);font-weight:700">${i.total.toFixed(2)}</td>
            <td style="font-family:var(--mono);color:var(--green)">${i.amount_paid.toFixed(2)}</td>
            <td style="font-family:var(--mono);color:${i.balance_due>0?"var(--warn)":"var(--muted)"};font-weight:${i.balance_due>0?"700":"400"}">${i.balance_due>0?i.balance_due.toFixed(2):"—"}</td>
            <td><span style="font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;background:rgba(0,0,0,.2);color:${statusColor[i.status]||"var(--muted)"}">${i.status}</span></td>
            <td style="font-size:12px;color:var(--muted)">${i.created_at}</td>
            <td>${actions}</td>
        </tr>`;
    }).join("");
}

/* ── INVOICE DETAIL ── */
function openInvDetail(id){
    let inv = allB2BInvoices.find(i=>i.id===id);
    if(!inv) return;
    currentInvDetail = inv;

    document.getElementById("inv-detail-num").innerText = inv.invoice_number + " — " + inv.created_at;
    document.getElementById("inv-detail-meta").innerHTML = `
        <div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Client</span><span style="font-weight:700">${inv.client}</span></div>
        <div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Type</span><span>${inv.invoice_type.split("_").map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(" ")}</span></div>
        <div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Status</span><span style="font-weight:700;color:${inv.status==="paid"?"var(--green)":"var(--warn)"}">${inv.status.toUpperCase()}</span></div>`;

    document.getElementById("inv-detail-items").innerHTML = inv.items.map(item=>`
        <tr>
            <td style="font-size:13px;padding:8px 0;border-bottom:1px solid var(--border);color:var(--text)">${item.product}</td>
            <td style="text-align:center;font-family:var(--mono);padding:8px 0;border-bottom:1px solid var(--border)">${item.qty.toFixed(0)}</td>
            <td style="text-align:right;font-family:var(--mono);padding:8px 0;border-bottom:1px solid var(--border)">${item.unit_price.toFixed(2)}</td>
            <td style="text-align:right;font-family:var(--mono);font-weight:700;padding:8px 0;border-bottom:1px solid var(--border);color:var(--green)">${item.total.toFixed(2)}</td>
        </tr>`).join("");

    document.getElementById("inv-detail-totals").innerHTML = `
        <div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Subtotal</span><span style="font-family:var(--mono)">${inv.subtotal.toFixed(2)}</span></div>
        ${inv.discount>0?`<div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Discount</span><span style="font-family:var(--mono);color:var(--danger)">-${inv.discount.toFixed(2)}</span></div>`:""}
        <div style="display:flex;justify-content:space-between;font-size:16px;font-weight:800;padding:8px 0;border-top:1px solid var(--border2);margin-top:6px"><span>Total</span><span style="font-family:var(--mono);color:var(--green)">${inv.total.toFixed(2)} EGP</span></div>
        ${inv.balance_due>0?`<div style="display:flex;justify-content:space-between;font-size:13px;padding:4px 0"><span style="color:var(--muted)">Balance Due</span><span style="font-family:var(--mono);font-weight:700;color:var(--warn)">${inv.balance_due.toFixed(2)} EGP</span></div>`:""}`;

    document.getElementById("inv-detail-modal").classList.add("open");
}

function printInvDetail(){
    let inv = currentInvDetail;
    if(!inv) return;
    let rows = inv.items.map(item=>`
        <tr>
            <td style="padding:8px 12px;border-bottom:1px solid #eee">${item.product}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center">${item.qty.toFixed(0)}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">${item.unit_price.toFixed(2)}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:700">${item.total.toFixed(2)}</td>
        </tr>`).join("");
    let win = window.open("","_blank","width=650,height=900");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${inv.invoice_number}</title>
    <style>body{font-family:Arial,sans-serif;padding:30px;color:#111;max-width:600px;margin:0 auto}
    .header{text-align:center;margin-bottom:20px;padding-bottom:16px;border-bottom:2px solid #2a7a2a}
    .logo{max-height:120px;margin-bottom:6px}
    .company{font-size:18px;font-weight:900;color:#2a7a2a;margin-bottom:4px}
    .meta{display:flex;justify-content:space-between;font-size:13px;margin-bottom:16px}
    .meta-label{color:#555}
    table{width:100%;border-collapse:collapse;margin-bottom:16px}
    thead{background:#f0f0f0}
    th{padding:8px 12px;text-align:left;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
    .totals{text-align:right}
    .total-final{font-size:18px;font-weight:900;color:#2a7a2a;border-top:2px solid #2a7a2a;padding-top:8px;margin-top:8px}
    .footer{text-align:center;margin-top:30px;font-size:11px;color:#888;border-top:1px solid #eee;padding-top:10px;font-style:italic}
    @media print{button{display:none}}
    </style></head><body>
    <div class="header">
        <img src="/static/Logo.png" class="logo"><br>
        <div class="company">Habiba Organic Farm</div>
        <div style="font-size:12px;color:#555">Commercial registry: 126278 | Tax ID: 560042604</div>
    </div>
    <div class="meta">
        <div><div class="meta-label">Invoice #</div><b>${inv.invoice_number}</b></div>
        <div><div class="meta-label">Client</div><b>${inv.client}</b></div>
        <div><div class="meta-label">Date</div>${inv.created_at}</div>
        <div><div class="meta-label">Type</div>${inv.invoice_type.split("_").map(w=>w.charAt(0).toUpperCase()+w.slice(1)).join(" ")}</div>
    </div>
    <table><thead><tr><th>Product</th><th>QTY</th><th style="text-align:right">Price</th><th style="text-align:right">Total</th></tr></thead>
    <tbody>${rows}</tbody></table>
    <div class="totals">
        <div style="font-size:13px">Subtotal: ${inv.subtotal.toFixed(2)}</div>
        ${inv.discount>0?`<div style="font-size:13px;color:#c0392b">Discount: -${inv.discount.toFixed(2)}</div>`:""}
        <div class="total-final">Total: ${inv.total.toFixed(2)} EGP</div>
        ${inv.balance_due>0?`<div style="color:#c0392b;font-size:13px;margin-top:6px">Balance Due: ${inv.balance_due.toFixed(2)} EGP</div>`:""}
    </div>
    <div style="margin-top:40px;display:flex;justify-content:space-between;font-size:12px;color:#555;border-top:1px solid #ddd;padding-top:16px">
        <div><div style="border-bottom:1px solid #aaa;width:160px;margin-bottom:4px;padding-bottom:20px"></div>Received by</div>
        <div><div style="border-bottom:1px solid #aaa;width:160px;margin-bottom:4px;padding-bottom:20px"></div>Receipt Date</div>
    </div>
    <div class="footer">Desert going green | habibaorganicfarm | habibacommunity.com</div>
    <br><button onclick="window.print()">🖨 Print</button>
    </body></html>`);
    win.document.close();
}

/* ── COLLECT PAYMENT (cash / full_payment) ── */
let collectInvoiceNum = null;

function openCollectModal(id, num, balance){
    collectInvoiceId  = id;
    collectInvoiceNum = num;
    document.getElementById("collect-modal-sub").innerText = `${num} — Balance: ${balance.toFixed(2)} EGP`;
    document.getElementById("collect-amount").value  = balance.toFixed(2);
    document.getElementById("collect-inv-num").value = "";
    document.getElementById("collect-modal").classList.add("open");
}

async function saveCollect(){
    let typed  = document.getElementById("collect-inv-num").value.trim();
    let amount = parseFloat(document.getElementById("collect-amount").value)||0;
    if(!typed){ showToast("Please enter the invoice number to confirm"); return; }
    if(typed !== collectInvoiceNum){
        showToast(`Invoice number doesn't match — expected: ${collectInvoiceNum}`);
        document.getElementById("collect-inv-num").style.border = "1px solid var(--danger)";
        return;
    }
    document.getElementById("collect-inv-num").style.border = "";
    if(amount<=0){ showToast("Enter a valid amount"); return; }
    let res  = await fetch(`/accounting/api/b2b-invoices/${collectInvoiceId}/collect`,{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({amount}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    document.getElementById("collect-modal").classList.remove("open");
    showToast(`✓ Payment collected — Revenue recognized! Status: ${data.status}`);
    loadB2BInvoices();
}

/* ── CONSIGNMENT PAYMENT ── */
async function openConsModal(clientId, clientName, outstanding, discountPct){
    consClientId = clientId;
    consClientName = clientName;
    consItems = [];
    consDiscountPct = Number(discountPct) || 0;
    document.getElementById("cons-modal-sub").innerText = `${clientName} — Client outstanding balance: ${outstanding.toFixed(2)} EGP`;
    document.getElementById("cons-amount").value = "0";
    document.getElementById("cons-amount-display").innerText = "0.00 EGP";
    document.getElementById("cons-subtotal-display").innerText = "0.00";
    document.getElementById("cons-discount-display").innerText = "-0.00";
    // Fill month selector
    let sel = document.getElementById("cons-month");
    sel.innerHTML = '<option value="">General payment (no specific month)</option>';
    let d = new Date();
    for(let i=0;i<12;i++){
        let label = d.toLocaleDateString("en-GB",{month:"long",year:"numeric"});
        sel.innerHTML += `<option value="${label}">${label}</option>`;
        d.setMonth(d.getMonth()-1);
    }
    // Load the items this client was sent on consignment
    let body = document.getElementById("cons-items-body");
    body.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:24px">Loading items…</td></tr>`;
    document.getElementById("cons-modal").classList.add("open");
    try {
        let res = await fetch(`/accounting/api/b2b-clients/${clientId}/consignment-items`);
        let data = await res.json();
        if(!res.ok){ throw new Error((data && data.detail) || "Unable to load items"); }
        consItems = Array.isArray(data) ? data : (data.items || []);
        // The server is the authority on the client's discount; the caller's
        // value is only a hint from the already-rendered client row.
        if(data && data.discount_pct !== undefined) consDiscountPct = Number(data.discount_pct) || 0;
    } catch(e){
        consItems = [];
        body.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--danger);padding:24px">${e.message}</td></tr>`;
        return;
    }
    renderConsItemRows();
}

function renderConsItemRows(){
    let body = document.getElementById("cons-items-body");
    if(!consItems.length){
        body.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:24px">No consignment items on record for this client.</td></tr>`;
        recalcConsTotal();
        return;
    }
    body.innerHTML = consItems.map((it, idx) => `
        <tr style="border-top:1px solid var(--border2)">
            <td style="padding:8px 10px;color:var(--text);font-weight:600">
                ${it.name}
                <div style="font-size:10px;color:var(--muted)">Remaining sent: ${it.qty_remaining} ${it.unit}</div>
            </td>
            <td class="mono" style="padding:8px 10px;text-align:right;color:var(--sub)">${it.unit_price.toFixed(2)}</td>
            <td style="padding:8px 10px;text-align:center">
                <input type="number" min="0" step="any" value="" placeholder="0"
                    data-idx="${idx}" oninput="recalcConsTotal()"
                    style="width:90px;background:var(--card2);border:1px solid var(--border2);border-radius:8px;padding:6px 8px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;text-align:center">
            </td>
            <td class="mono cons-line-total" data-idx="${idx}" style="padding:8px 10px;text-align:right;color:var(--text);font-weight:700">0.00</td>
        </tr>`).join("");
    recalcConsTotal();
}

function recalcConsTotal(){
    let subtotal = 0;
    document.querySelectorAll("#cons-items-body input[data-idx]").forEach(inp => {
        let idx = parseInt(inp.dataset.idx);
        let qty = parseFloat(inp.value) || 0;
        if(qty < 0){ qty = 0; }
        let line = qty * (consItems[idx] ? consItems[idx].unit_price : 0);
        let cell = document.querySelector(`.cons-line-total[data-idx="${idx}"]`);
        if(cell){ cell.innerText = line.toFixed(2); }
        subtotal += line;
    });
    // Item prices are gross; the client's discount comes off the total, the
    // same way it does on the consignment invoice this payment pays down.
    let pct      = Math.min(100, Math.max(0, consDiscountPct || 0));
    subtotal     = Math.round(subtotal * 100) / 100;
    let discount = Math.round(subtotal * pct) / 100;
    let total    = Math.round((subtotal - discount) * 100) / 100;
    document.getElementById("cons-subtotal-display").innerText = subtotal.toFixed(2);
    document.getElementById("cons-discount-pct").innerText     = String(Math.round(pct * 100) / 100);
    document.getElementById("cons-discount-display").innerText = `-${discount.toFixed(2)}`;
    document.getElementById("cons-discount-row").style.display = pct > 0 ? "flex" : "none";
    document.getElementById("cons-amount").value = total;
    document.getElementById("cons-amount-display").innerText = `${total.toFixed(2)} EGP`;
}

let refundProducts = [];
let refundProductsClientId = null;

async function ensureRefundProducts(clientId){
    // Re-fetch per client: the picker must offer that client's own agreed
    // price, not the generic list price.
    if(refundProducts.length && refundProductsClientId === clientId) return;
    try{
        const res = await fetch(`/accounting/api/products-list?client_id=${clientId}`, { credentials: "same-origin" });
        refundProducts = res.ok ? await res.json() : [];
        refundProductsClientId = clientId;
    }catch(e){ refundProducts = []; }
}

function refundProductOptions(selected){
    return refundProducts.map(p =>
        `<option value="${p.id}" data-price="${p.price}" ${String(p.id)===String(selected)?"selected":""}>${p.name}</option>`
    ).join("");
}

function addRefundItem(){
    const wrap = document.getElementById("refund-items");
    const row = document.createElement("div");
    row.className = "refund-item-row";
    row.style.cssText = "display:grid;grid-template-columns:1fr 90px 110px 30px;gap:8px;margin-bottom:8px;align-items:center";
    row.innerHTML = `
        <select data-field="product" onchange="onRefundProductChange(this)"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:9px 10px;color:var(--text);font-family:var(--sans);font-size:13px;outline:none;width:100%">
            ${refundProductOptions(null)}
        </select>
        <input data-field="qty" type="number" min="0" step="any" placeholder="Qty" oninput="refreshRefundTotal()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:9px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100%">
        <input data-field="price" type="number" min="0" step="any" placeholder="Unit price" oninput="refreshRefundTotal()"
            style="background:var(--card2);border:1px solid var(--border2);border-radius:10px;padding:9px 10px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100%">
        <button onclick="this.parentElement.remove();refreshRefundTotal()" title="Remove"
            style="background:transparent;border:none;color:var(--danger);font-size:16px;cursor:pointer">×</button>`;
    wrap.appendChild(row);
    onRefundProductChange(row.querySelector('[data-field="product"]'));
}

function onRefundProductChange(sel){
    const opt = sel.options[sel.selectedIndex];
    const priceInput = sel.parentElement.querySelector('[data-field="price"]');
    if(opt && priceInput && !priceInput.value) priceInput.value = opt.dataset.price || "";
    refreshRefundTotal();
}

function collectRefundItems(){
    return [].slice.call(document.querySelectorAll("#refund-items .refund-item-row")).map(row => ({
        product_id: parseInt(row.querySelector('[data-field="product"]').value, 10),
        qty:        parseFloat(row.querySelector('[data-field="qty"]').value) || 0,
        unit_price: parseFloat(row.querySelector('[data-field="price"]').value) || 0,
    }));
}

function refreshRefundTotal(){
    const total = collectRefundItems().reduce((s, i) => s + i.qty * i.unit_price, 0);
    const el = document.getElementById("refund-total");
    if(el) el.innerText = total.toFixed(2) + " EGP";
}

async function openRefundModal(clientId, clientName, outstanding){
    refundInvoiceId  = clientId;
    refundInvoiceNum = clientName;
    document.getElementById("refund-modal-sub").innerText = `${clientName} — Outstanding balance: ${Number(outstanding||0).toFixed(2)} EGP`;
    document.getElementById("refund-inv-num").value = "";
    document.getElementById("refund-reason").value = "";
    document.getElementById("refund-inv-num").style.border = "1px solid var(--border2)";
    document.getElementById("refund-items").innerHTML = "";
    document.getElementById("refund-modal").classList.add("open");
    await ensureRefundProducts(clientId);
    addRefundItem();
}

async function saveRefund(){
    let typed  = document.getElementById("refund-inv-num").value.trim();
    let reason = document.getElementById("refund-reason").value.trim();
    if(!typed){ showToast("Please enter the client name to confirm"); return; }
    if(typed !== refundInvoiceNum){
        showToast(`Client name doesn't match — expected: ${refundInvoiceNum}`);
        document.getElementById("refund-inv-num").style.border = "1px solid var(--danger)";
        return;
    }
    document.getElementById("refund-inv-num").style.border = "1px solid var(--border2)";
    const items = collectRefundItems().filter(i => i.product_id && i.qty > 0);
    if(!items.length){ showToast("Add at least one returned item"); return; }
    let res  = await fetch(`/accounting/api/b2b-clients/${refundInvoiceId}/refund`,{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({items, reason:reason||null}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    document.getElementById("refund-modal").classList.remove("open");
    showToast(`✓ ${data.refund_number} recorded — New outstanding: ${Number(data.client_outstanding||0).toFixed(2)} EGP`);
    loadB2BInvoices();
    if(typeof loadB2BClients === "function") loadB2BClients();
}

async function saveConsPayment(){
    let month  = document.getElementById("cons-month").value;
    // Gather sold items with a positive qty
    let items = [];
    document.querySelectorAll("#cons-items-body input[data-idx]").forEach(inp => {
        let idx = parseInt(inp.dataset.idx);
        let qty = parseFloat(inp.value) || 0;
        if(qty > 0 && consItems[idx]){
            items.push({product_id: consItems[idx].product_id, qty, unit_price: consItems[idx].unit_price});
        }
    });
    if(!items.length){ showToast("Enter the quantity sold for at least one item"); return; }
    let amount = parseFloat(document.getElementById("cons-amount").value)||0;
    if(amount<=0){ showToast("Amount must be greater than 0"); return; }
    let res    = await fetch(`/accounting/api/b2b-clients/${consClientId}/consignment-payment`,{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({amount, month_label:month||null, items}),
    });
    let data = await res.json();
    if(data.detail){ showToast("Error: "+data.detail); return; }
    document.getElementById("cons-modal").classList.remove("open");
    showToast(`✓ ${amount.toFixed(2)} EGP recorded for ${data.client || consClientName}${month?" ("+month+")":""}`);
    loadB2BInvoices();
    if(document.getElementById("b2b-type-filter").value === "consignment"){ loadConsPaymentHistory(); }
}

/* ── CONSIGNMENT PAYMENT HISTORY ── */
async function loadConsPaymentHistory(){
    let list = document.getElementById("cons-payment-list");
    if(!list) return;
    // Make sure we have the client list (may not be loaded yet on this tab)
    if(!allB2BClients || !allB2BClients.length){
        try {
            let cr = await fetch(`/accounting/api/b2b-clients`);
            let cd = await cr.json();
            if(cr.ok && Array.isArray(cd)) allB2BClients = cd;
        } catch(e){}
    }
    // Gather the consignment clients currently shown, pull each one's sales
    let clients = (allB2BClients || []).filter(c => c.is_consignment);
    if(!clients.length){
        list.innerHTML = `<div style="color:var(--muted);font-size:13px;padding:12px">No consignment clients.</div>`;
        return;
    }
    let results = await Promise.all(clients.map(async c => {
        try {
            let r = await fetch(`/accounting/api/b2b-clients/${c.id}/consignment-sales`);
            let sales = await r.json();
            return {client:c, sales: Array.isArray(sales)?sales:[]};
        } catch(e){ return {client:c, sales:[]}; }
    }));
    let rows = [];
    results.forEach(({client, sales}) => {
        sales.forEach(s => {
            let itemsHtml = s.items.map(it =>
                `<span style="display:inline-block;background:var(--card2);border:1px solid var(--border2);border-radius:6px;padding:2px 8px;margin:2px 4px 2px 0;font-size:11px">${it.name} × ${it.qty} ${it.unit} @ ${it.unit_price.toFixed(2)} = <b>${it.total.toFixed(2)}</b></span>`
            ).join("");
            let when = s.created_at ? new Date(s.created_at).toLocaleDateString("en-GB") : "";
            rows.push(`
                <div style="border:1px solid var(--border2);border-radius:10px;padding:12px 14px;margin-bottom:10px">
                    <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px">
                        <div style="font-weight:700;color:var(--text)">${client.name}
                            ${s.month_label?`<span style="font-size:11px;color:var(--teal);font-weight:700;margin-left:8px">${s.month_label}</span>`:``}
                        </div>
                        <div style="text-align:right">
                            <div class="mono" style="font-weight:900;color:var(--teal)">${s.amount.toFixed(2)} EGP</div>
                            ${(s.discount||0) > 0
                                ? `<div class="mono" style="font-size:10px;color:var(--muted)">${(s.subtotal||0).toFixed(2)} sold - ${s.discount.toFixed(2)} discount</div>`
                                : ``}
                        </div>
                    </div>
                    <div style="font-size:11px;color:var(--muted);margin-bottom:6px">${when}</div>
                    <div>${itemsHtml || '<span style="color:var(--muted);font-size:11px">No item detail</span>'}</div>
                </div>`);
        });
    });
    list.innerHTML = rows.length ? rows.join("")
        : `<div style="color:var(--muted);font-size:13px;padding:12px">No consignment payments recorded yet.</div>`;
}

["inv-detail-modal","collect-modal","cons-modal","refund-modal"].forEach(id=>{
    let el = document.getElementById(id);
    if(el) el.addEventListener("click",function(e){ if(e.target===this) this.classList.remove("open"); });
});

init();
</script>
</body>
</html>
"""