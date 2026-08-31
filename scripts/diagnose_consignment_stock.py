"""
Consignment stock diagnostic — why does a client's "stock on hand" read wrong?

Usage:
    python -m scripts.diagnose_consignment_stock "Yo Studio"

Prints, for the matching client(s), every source that feeds the on-hand
figure, so the divergence is visible rather than inferred:

  • each consignment: its invoice lines (authoritative) beside the Consignment
    mirror's qty_sent / qty_sold / qty_returned, so drift between them shows
  • each recorded payment's sold items                          (the Accounting flow)
  • the resulting on-hand per product, and whether the two flows overlap
  • deliveries booked on NON-consignment invoices, which are excluded from
    consignment stock by design and are a common source of surprise

Read-only, and deliberately built on explicit column selects and raw SQL so it
still runs against a database whose schema is behind the models.
"""
from __future__ import annotations

import asyncio
import sys

from sqlalchemy import text

from app.db.session import AsyncSessionLocal


def f(v) -> float:
    return float(v or 0)


async def _has_table(db, table: str) -> bool:
    r = await db.execute(text(
        "select 1 from information_schema.tables where table_name = :t"
    ), {"t": table})
    return r.first() is not None


async def _has_column(db, table: str, column: str) -> bool:
    r = await db.execute(text(
        "select 1 from information_schema.columns "
        "where table_name = :t and column_name = :c"
    ), {"t": table, "c": column})
    return r.first() is not None


async def report(db, client) -> None:
    cid, cname, terms, disc, outstanding = client
    print("=" * 82)
    print(f"CLIENT #{cid}  {cname}")
    print(f"  terms={terms}  discount={f(disc)}%  outstanding={f(outstanding):.2f}")
    print("=" * 82)

    cons = (await db.execute(text("""
        select c.id, c.ref_number, c.status, c.invoice_id, c.created_at,
               ci.product_id, coalesce(p.name, '#' || ci.product_id) as pname,
               ci.qty_sent, ci.qty_sold, ci.qty_returned, ci.unit_price
        from consignments c
        join consignment_items ci on ci.consignment_id = c.id
        left join products p on p.id = ci.product_id
        where c.client_id = :cid
        order by c.created_at, c.id, pname
    """), {"cid": cid})).all()

    # The invoice behind a consignment is the authoritative record of what went
    # out; the Consignment mirror can drift from it. Both are printed so the
    # divergence is visible instead of inferred.
    inv_lines = (await db.execute(text("""
        select c.id, ii.product_id, coalesce(p.name, '#' || ii.product_id) as pname,
               ii.qty, ii.unit_price
        from consignments c
        join b2b_invoice_items ii on ii.invoice_id = c.invoice_id
        left join products p on p.id = ii.product_id
        where c.client_id = :cid and c.invoice_id is not null
        order by c.id, pname
    """), {"cid": cid})).all()
    inv_by_cons = {}
    for (cons_id, pid, pname, qty, price) in inv_lines:
        inv_by_cons.setdefault(cons_id, []).append((pid, pname, f(qty), f(price)))

    print(f"\n-- CONSIGNMENTS ({len({r[0] for r in cons})} refs, {len(cons)} lines) " + "-" * 34)
    sent, settled, returned, names = {}, {}, {}, {}
    if not cons:
        print("  (none - this client has no consignment records at all, so consignment")
        print("   stock is empty by definition. Check the invoice types below.)")
    seen = None
    drift = []
    for (_id, ref, status, inv_id, created, pid, pname, q_sent, q_sold, q_ret, price) in cons:
        if seen != _id:
            seen = _id
            print(f"  {ref or _id}  status={status}  invoice_id={inv_id}  "
                  f"{created.strftime('%d-%b-%Y') if created else '-'}")
            for (ipid, ipname, iqty, iprice) in inv_by_cons.get(_id, []):
                names[ipid] = ipname
                print(f"      [invoice] {ipname[:26]:<26} qty={iqty:>10.3f} "
                      f"@ {iprice:.2f}   <- authoritative")
                sent[ipid] = sent.get(ipid, 0.0) + iqty
        pending = f(q_sent) - f(q_sold) - f(q_ret)
        names[pid] = pname
        print(f"      [mirror]  {pname[:26]:<26} sent={f(q_sent):>10.3f} "
              f"sold={f(q_sold):>9.3f} returned={f(q_ret):>9.3f} "
              f"pending={pending:>10.3f} @ {f(price):.2f}")
        if _id not in inv_by_cons:
            # Nothing invoiced behind it, so the mirror IS the record of what went out.
            sent[pid] = sent.get(pid, 0.0) + f(q_sent)
        else:
            inv_qty = sum(q for (ip, _n, q, _p) in inv_by_cons[_id] if ip == pid)
            if abs(inv_qty - f(q_sent)) > 0.001:
                drift.append(f"{pname}: invoice {inv_qty:.3f} vs mirror {f(q_sent):.3f}")
        settled[pid]  = settled.get(pid, 0.0) + f(q_sold)
        returned[pid] = returned.get(pid, 0.0) + f(q_ret)

    if drift:
        print("\n  !! Invoice / mirror mismatch - stock follows the invoice:")
        for d in drift:
            print("       " + d)

    has_sales = await _has_table(db, "consignment_sales")
    has_split = has_sales and await _has_column(db, "consignment_sales", "subtotal")
    money_cols = ("s.subtotal, s.discount, s.amount" if has_split
                  else "0 as subtotal, 0 as discount, s.amount")
    sales = [] if not has_sales else (await db.execute(text(f"""
        select s.id, s.month_label, s.created_at, {money_cols},
               si.product_id, coalesce(p.name, '#' || si.product_id) as pname,
               si.qty, si.unit_price
        from consignment_sales s
        join consignment_sale_items si on si.sale_id = s.id
        left join products p on p.id = si.product_id
        where s.client_id = :cid
        order by s.created_at, s.id, pname
    """), {"cid": cid})).all()

    print(f"\n-- RECORDED PAYMENTS / REPORTED SALES ({len({r[0] for r in sales})}) " + "-" * 30)
    if not has_sales:
        print("  (this database has no consignment_sales tables - start the app once")
        print("   so the schema guard creates them, or run alembic upgrade head)")
    elif not has_split:
        print("  (this database predates the subtotal/discount split - amounts only)")
    reported = {}
    seen = None
    for (sid, month, created, sub, dis, amt, pid, pname, qty, price) in sales:
        if seen != sid:
            seen = sid
            print(f"  sale#{sid}  {month or 'no month'}  "
                  f"{created.strftime('%d-%b-%Y') if created else '-'}  "
                  f"subtotal={f(sub):.2f} discount={f(dis):.2f} amount={f(amt):.2f}")
        names.setdefault(pid, pname)
        print(f"      {pname[:28]:<28} qty={f(qty):>10.3f} @ {f(price):.2f}")
        reported[pid] = reported.get(pid, 0.0) + f(qty)

    print("\n-- ON HAND PER PRODUCT " + "-" * 58)
    print(f"  {'product':<28}{'sent':>10}{'settled':>10}{'reported':>10}"
          f"{'returned':>10}{'on hand':>10}")
    both_flows = []
    for pid in sorted(set(sent) | set(reported), key=lambda p: str(names.get(p, ""))):
        s_, se = sent.get(pid, 0.0), settled.get(pid, 0.0)
        rp, rt = reported.get(pid, 0.0), returned.get(pid, 0.0)
        raw = s_ - se - rp - rt
        print(f"  {str(names.get(pid, pid))[:28]:<28}{s_:>10.3f}{se:>10.3f}{rp:>10.3f}"
              f"{rt:>10.3f}{max(0.0, raw):>10.3f}" + ("   (clamped from %.3f)" % raw if raw < 0 else ""))
        if se > 0 and rp > 0:
            both_flows.append(str(names.get(pid, pid)))

    if both_flows:
        print("\n  !! Both flows were used for: " + ", ".join(both_flows))
        print("     Settle wrote qty_sold AND payments recorded sold items. If those")
        print("     cover the SAME goods the quantity is subtracted twice here, and")
        print("     the revenue was booked twice in the accounts.")

    invs = (await db.execute(text("""
        select invoice_number, invoice_type, status, total, amount_paid
        from b2b_invoices where client_id = :cid order by created_at, id
    """), {"cid": cid})).all()
    non_cons = [i for i in invs if (i[1] or "") != "consignment"]
    print(f"\n-- INVOICES ({len(invs)} total, {len(non_cons)} NOT consignment) " + "-" * 24)
    for (num, itype, status, total, paid) in invs:
        flag = "" if (itype or "") == "consignment" else "   <- not consignment stock"
        print(f"  {str(num):<16} {str(itype):<14} {str(status):<12} "
              f"total={f(total):>10.2f} paid={f(paid):>10.2f}{flag}")
    if non_cons:
        print("\n  Note: goods on the invoices marked above were sold outright, not")
        print("  placed on consignment, so they are deliberately absent from stock")
        print("  on hand. If they should be consignment, the invoice type is wrong.")


async def main() -> None:
    needle = " ".join(sys.argv[1:]).strip()
    if not needle:
        print(__doc__)
        raise SystemExit(2)
    async with AsyncSessionLocal() as db:
        clients = (await db.execute(text(
            "select id, name, payment_terms, discount_pct, outstanding "
            "from b2b_clients where name ilike :n order by id"
        ), {"n": f"%{needle}%"})).all()
        if not clients:
            print(f"No B2B client matching {needle!r}.")
            raise SystemExit(1)
        for c in clients:
            await report(db, c)


if __name__ == "__main__":
    asyncio.run(main())
