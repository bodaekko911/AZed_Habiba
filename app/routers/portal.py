"""
Client Portal — public, read-only, token-addressed
==================================================
Prefix : /portal
Auth   : NONE. This is the one router in the system with no permission gate.

A B2B client opens ``/portal/c/<token>`` and sees their own live account:
statement with running balance, products received (netted against returns),
and the delivery log behind those numbers. No login, no account to manage.

Security model
--------------
The token IS the credential. It is 32 bytes from ``secrets.token_urlsafe``
(~256 bits), issued per client from the B2B clients screen, and it can be
revoked or rotated there at any time — revoking clears the token so the old
URL 404s. Every lookup requires BOTH a matching token and ``portal_enabled``,
so a leaked-then-revoked link is dead even if the row is later re-enabled with
a new token.

Because anyone holding the link sees that client's financials:
  • pages are served ``noindex, nofollow`` and ``Cache-Control: no-store``
  • an unknown, disabled, or inactive-client token returns a plain 404 that
    reveals nothing about whether the token ever existed
  • the portal reads one client's rows only — the token never widens scope
"""

from __future__ import annotations

from datetime import datetime, timezone
from html import escape
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.encoders import jsonable_encoder
from fastapi.responses import HTMLResponse, JSONResponse
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_async_session
from app.models.b2b import B2BClient

router = APIRouter(prefix="/portal", tags=["Client Portal"])

# Applied to every portal response: keep these pages out of search indexes and
# out of shared caches, since the URL alone unlocks the data.
PORTAL_HEADERS = {
    "X-Robots-Tag": "noindex, nofollow, noarchive",
    "Cache-Control": "no-store, max-age=0",
    "Referrer-Policy": "no-referrer",
}


async def _resolve_client(db: AsyncSession, token: str) -> B2BClient:
    """Token → client, or 404. Requires an enabled portal on an active client."""
    if not token or len(token) < 16:
        raise HTTPException(status_code=404, detail="Not found")
    result = await db.execute(
        select(B2BClient).where(
            B2BClient.portal_token == token,
            B2BClient.portal_enabled == True,   # noqa: E712 — SQL boolean, not Python
            B2BClient.is_active == True,        # noqa: E712
        )
    )
    client = result.scalar_one_or_none()
    if not client:
        raise HTTPException(status_code=404, detail="Not found")
    return client


async def _portal_payload(db: AsyncSession, client: B2BClient) -> dict:
    # Imported here rather than at module scope: app.routers.b2b pulls in the
    # permission stack, and this router must stay importable without it.
    from app.routers.b2b import (
        _build_client_consignment_stock_payload,
        _build_client_products_payload,
        _build_client_statement_payload,
    )

    statement = await _build_client_statement_payload(client.id, db)
    products = await _build_client_products_payload(client.id, db)
    stock = await _build_client_consignment_stock_payload(client.id, db)
    return {
        "client": {
            "name":           statement["client"]["name"],
            "code":           statement["client"]["code"],
            "contact_person": statement["client"]["contact_person"],
            "phone":          statement["client"]["phone"],
            "email":          statement["client"]["email"],
            "address":        statement["client"]["address"],
            "payment_terms":  statement["client"]["payment_terms"],
        },
        "as_of":            statement["statement_date"],
        "generated_at":     datetime.now(timezone.utc).strftime("%d-%b-%Y %H:%M UTC"),
        "transactions":     statement["transactions"],
        # Projected down to the four fields the page shows, deliberately:
        #   • the raw "date" datetime is dropped — JSONResponse serialises with
        #     json.dumps, which cannot encode a datetime (it 500s the endpoint)
        #   • "user_name" is dropped — that is OUR staff member who recorded the
        #     payment, and it has no business being in a client-facing payload
        "payment_activity": [
            {
                "date":   p.get("date_str") or "—",
                "ref":    p.get("ref") or "—",
                "desc":   p.get("desc") or "Payment received",
                "amount": float(p.get("amount") or 0),
            }
            for p in statement["payment_activity"]
        ],
        "total_invoiced":   statement["total_invoiced"],
        "total_paid":       statement["total_paid"],
        "balance_due":      statement["balance_due"],
        "products":         products["products"],
        "deliveries":       products["deliveries"],
        "product_totals":   products["totals"],
        # Goods the client still holds on consignment — their own stock of our
        # products. Empty for clients who never took anything on consignment.
        "stock":            [r for r in stock["items"] if r["qty_on_hand"] > 0],
        "stock_totals":     stock["totals"],
    }


@router.get("/c/{token}/data")
async def portal_data(token: str, db: AsyncSession = Depends(get_async_session)):
    """JSON behind the page — polled so the view stays live without a reload."""
    client = await _resolve_client(db, token)
    payload = await _portal_payload(db, client)
    # jsonable_encoder, not raw JSONResponse: building the Response ourselves
    # skips FastAPI's encoding step, so any Decimal or datetime that ever
    # reaches this payload would raise inside json.dumps and 500 the page.
    return JSONResponse(jsonable_encoder(payload), headers=PORTAL_HEADERS)


@router.get("/c/{token}", response_class=HTMLResponse)
async def portal_page(token: str, request: Request, db: AsyncSession = Depends(get_async_session)):
    client = await _resolve_client(db, token)

    # Count real opens only — the polling endpoint above deliberately does not
    # touch these, so "12 views" means the client opened it 12 times.
    client.portal_view_count = int(client.portal_view_count or 0) + 1
    client.portal_last_viewed_at = datetime.now(timezone.utc)
    await db.commit()

    return HTMLResponse(_portal_html(client.name), headers=PORTAL_HEADERS)


def _portal_html(client_name: str) -> str:
    return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Account — __CLIENT__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#f6f7f2;--card:#ffffff;--line:#e6e8e0;--line2:#d3d7ca;
  --text:#15180f;--sub:#5c6350;--muted:#8b917d;
  --green:#0f8a43;--red:#c0392b;--amber:#b7791f;--blue:#2563eb;
  --sans:'Outfit',sans-serif;--mono:'JetBrains Mono',monospace;--r:14px;
}
@media (prefers-color-scheme: dark){
  :root{
    --bg:#0b0e07;--card:#141810;--line:#242a1c;--line2:#333a28;
    --text:#f1f4ea;--sub:#a8b096;--muted:#6f7761;
    --green:#5ee08a;--red:#ff7a6b;--amber:#f0b54a;--blue:#7aa8ff;
  }
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:var(--sans);background:var(--bg);color:var(--text);font-size:15px;line-height:1.5;}
.wrap{max-width:1080px;margin:0 auto;padding:26px 18px 60px;display:flex;flex-direction:column;gap:18px;}
header.top{display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap;
  background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:20px 22px;}
.brand{display:flex;align-items:center;gap:14px;}
.brand img{height:52px;width:auto;object-fit:contain;}
.brand-name{font-size:17px;font-weight:800;letter-spacing:-.3px;}
.brand-meta{font-size:11px;color:var(--muted);margin-top:2px;}
.client-name{font-size:22px;font-weight:800;letter-spacing:-.5px;text-align:right;}
.client-meta{font-size:12px;color:var(--muted);margin-top:3px;text-align:right;}
.live{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:700;color:var(--green);
  text-transform:uppercase;letter-spacing:1px;}
.live .dot{width:7px;height:7px;border-radius:50%;background:var(--green);animation:pulse 2s infinite;}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.25}}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:16px 18px;}
.card-label{font-size:10px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:var(--muted);margin-bottom:7px;}
.card-value{font-family:var(--mono);font-size:23px;font-weight:700;letter-spacing:-.5px;}
.card-note{font-size:11px;color:var(--muted);margin-top:5px;}
.v-red{color:var(--red);} .v-green{color:var(--green);} .v-blue{color:var(--blue);} .v-amber{color:var(--amber);}
.tabs{display:flex;gap:5px;background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:5px;flex-wrap:wrap;}
.tab{flex:1;min-width:130px;padding:10px 14px;border:none;border-radius:10px;background:transparent;cursor:pointer;
  font-family:var(--sans);font-size:13px;font-weight:700;color:var(--muted);transition:all .18s;}
.tab.active{background:var(--bg);color:var(--text);box-shadow:0 1px 3px rgba(0,0,0,.07);}
.panel{display:none;flex-direction:column;gap:14px;}
.panel.active{display:flex;}
.block{background:var(--card);border:1px solid var(--line);border-radius:var(--r);overflow:hidden;}
.block-title{padding:13px 18px;border-bottom:1px solid var(--line);font-size:11px;font-weight:700;
  letter-spacing:1.2px;text-transform:uppercase;color:var(--muted);display:flex;justify-content:space-between;
  align-items:center;gap:10px;flex-wrap:wrap;}
.scroll{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:13.5px;}
thead th{text-align:left;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;
  color:var(--muted);padding:11px 16px;background:rgba(0,0,0,.02);white-space:nowrap;}
@media (prefers-color-scheme: dark){thead th{background:rgba(255,255,255,.03);}}
td{padding:11px 16px;border-top:1px solid var(--line);color:var(--sub);vertical-align:top;}
td.name{color:var(--text);font-weight:600;}
td.num{font-family:var(--mono);text-align:right;white-space:nowrap;}
th.num{text-align:right;}
tr:last-child td{border-bottom:none;}
.tag{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;white-space:nowrap;}
.tag-paid{background:rgba(15,138,67,.12);color:var(--green);}
.tag-unpaid{background:rgba(192,57,43,.12);color:var(--red);}
.tag-partial{background:rgba(183,121,31,.14);color:var(--amber);}
.tag-refund{background:rgba(37,99,235,.12);color:var(--blue);}
.tag-return{background:rgba(37,99,235,.12);color:var(--blue);}
.tag-delivery{background:rgba(15,138,67,.12);color:var(--green);}
.empty{padding:34px 18px;text-align:center;color:var(--muted);font-size:13.5px;}
.delivery{border-top:1px solid var(--line);padding:14px 18px;}
.delivery:first-child{border-top:none;}
.delivery-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:9px;}
.delivery-ref{font-family:var(--mono);font-size:13px;font-weight:700;color:var(--text);}
.delivery-date{font-size:12px;color:var(--muted);}
.delivery-items{display:flex;flex-direction:column;gap:5px;}
.delivery-item{display:flex;justify-content:space-between;gap:12px;font-size:13px;color:var(--sub);}
.delivery-item .q{font-family:var(--mono);color:var(--muted);white-space:nowrap;}
footer.foot{text-align:center;font-size:11.5px;color:var(--muted);line-height:1.7;padding-top:6px;}
.err{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:28px;text-align:center;color:var(--red);font-weight:600;}
.actions{display:flex;gap:8px;flex-wrap:wrap;}
.btn{padding:8px 14px;border-radius:9px;border:1px solid var(--line2);background:var(--card);color:var(--sub);
  font-family:var(--sans);font-size:12px;font-weight:700;cursor:pointer;}
.btn:hover{color:var(--text);}
@media print{
  .tabs,.actions,.live{display:none!important;}
  body{background:#fff;color:#111;}
  .panel{display:flex!important;}
  .block,.card,header.top{border-color:#ddd;break-inside:avoid;}
}
@media(max-width:640px){
  .client-name,.client-meta{text-align:left;}
  header.top{flex-direction:column;align-items:flex-start;}
  td,thead th{padding:9px 11px;}
}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div class="brand">
      <img src="/static/Logo.png" alt="" onerror="this.style.display='none'">
      <div>
        <div class="brand-name">Habiba Organic Farm</div>
        <div class="brand-meta">Commercial registry: 126278 &nbsp;|&nbsp; Tax ID: 560042604</div>
      </div>
    </div>
    <div>
      <div class="client-name" id="client-name">__CLIENT__</div>
      <div class="client-meta" id="client-meta">Loading account…</div>
      <div class="client-meta" style="margin-top:6px"><span class="live"><span class="dot"></span> Live</span></div>
    </div>
  </header>

  <div id="error" style="display:none"><div class="err" id="error-text"></div></div>

  <div class="cards" id="cards"></div>

  <div class="tabs">
    <button class="tab active" data-panel="statement" type="button">Statement</button>
    <button class="tab" data-panel="products" type="button">Products received</button>
    <button class="tab" data-panel="stock" type="button" id="tab-stock" style="display:none">Stock on hand</button>
    <button class="tab" data-panel="deliveries" type="button">Delivery log</button>
  </div>

  <div class="panel active" id="panel-statement"></div>
  <div class="panel" id="panel-products"></div>
  <div class="panel" id="panel-stock"></div>
  <div class="panel" id="panel-deliveries"></div>

  <div class="actions">
    <button class="btn" type="button" onclick="window.print()">Print / Save as PDF</button>
    <button class="btn" type="button" onclick="load()">Refresh now</button>
  </div>

  <footer class="foot">
    <div id="generated"></div>
    <div>This page updates automatically. Questions about your account? Contact us and quote your client code.</div>
  </footer>
</div>

<script>
const DATA_URL = window.location.pathname.replace(/\\/+$/, "") + "/data";

const ESC = {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"};
function esc(v){ return v == null ? "" : String(v).replace(/[&<>"']/g, c => ESC[c]); }
function money(v){ return Number(v || 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2}); }
function qty(v){ return Number(v || 0).toLocaleString(undefined, {maximumFractionDigits:3}); }
function empty(cols, msg){ return `<tr><td colspan="${cols}" class="empty">${esc(msg)}</td></tr>`; }

function statusTag(status){
  const s = String(status || "").toLowerCase();
  const cls = s === "paid" ? "tag-paid" : s === "partial" ? "tag-partial"
            : s === "refund" ? "tag-refund" : s === "unpaid" ? "tag-unpaid" : "tag-partial";
  return `<span class="tag ${cls}">${esc(s || "—")}</span>`;
}

function renderCards(d){
  const due = Number(d.balance_due || 0);
  const st = d.stock_totals || {};
  const hasStock = (d.stock || []).length > 0;
  document.getElementById("cards").innerHTML = `
    <div class="card">
      <div class="card-label">Balance due</div>
      <div class="card-value ${due > 0 ? "v-red" : "v-green"}">${money(due)}</div>
      <div class="card-note">EGP &middot; as of ${esc(d.as_of)}</div>
    </div>
    <div class="card">
      <div class="card-label">Total invoiced</div>
      <div class="card-value">${money(d.total_invoiced)}</div>
      <div class="card-note">${d.transactions.length} transaction${d.transactions.length === 1 ? "" : "s"}</div>
    </div>
    <div class="card">
      <div class="card-label">Total paid</div>
      <div class="card-value v-green">${money(d.total_paid)}</div>
      <div class="card-note">${d.payment_activity.length} payment${d.payment_activity.length === 1 ? "" : "s"}</div>
    </div>
    <div class="card">
      <div class="card-label">Products received</div>
      <div class="card-value v-blue">${money(d.product_totals.value_net)}</div>
      <div class="card-note">${d.product_totals.product_lines} product${d.product_totals.product_lines === 1 ? "" : "s"}
        &middot; ${d.product_totals.deliveries} deliver${d.product_totals.deliveries === 1 ? "y" : "ies"}</div>
    </div>` + (hasStock ? `
    <div class="card">
      <div class="card-label">Stock on hand</div>
      <div class="card-value v-amber">${money(st.value_on_hand)}</div>
      <div class="card-note">${qty(st.qty_on_hand)} across ${st.product_lines} product${st.product_lines === 1 ? "" : "s"}</div>
    </div>` : "");
}

function renderStatement(d){
  const rows = d.transactions.length ? d.transactions.map(t => `
    <tr>
      <td class="num" style="text-align:left">${esc(t.date)}</td>
      <td class="name">${esc(t.ref)}</td>
      <td>${esc(t.desc)}</td>
      <td class="num">${t.debit ? money(t.debit) : "—"}</td>
      <td class="num">${t.credit ? money(t.credit) : "—"}</td>
      <td class="num" style="color:var(--text);font-weight:700">${money(t.balance)}</td>
      <td>${statusTag(t.status)}</td>
    </tr>`).join("") : empty(7, "No transactions on your account yet.");

  const pays = d.payment_activity.length ? d.payment_activity.map(p => `
    <tr>
      <td class="num" style="text-align:left">${esc(p.date)}</td>
      <td class="name">${esc(p.ref)}</td>
      <td>${esc(p.desc)}</td>
      <td class="num v-green">${money(p.amount)}</td>
    </tr>`).join("") : empty(4, "No payments recorded yet.");

  document.getElementById("panel-statement").innerHTML = `
    <div class="block">
      <div class="block-title"><span>Account statement</span><span>Amounts in EGP</span></div>
      <div class="scroll"><table>
        <thead><tr><th>Date</th><th>Reference</th><th>Description</th>
          <th class="num">Charged</th><th class="num">Paid</th><th class="num">Balance</th><th>Status</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>
    </div>
    <div class="block">
      <div class="block-title"><span>Payments received from you</span></div>
      <div class="scroll"><table>
        <thead><tr><th>Date</th><th>Reference</th><th>Description</th><th class="num">Amount</th></tr></thead>
        <tbody>${pays}</tbody>
      </table></div>
    </div>`;
}

function renderProducts(d){
  const t = d.product_totals;
  const rows = d.products.length ? d.products.map(p => `
    <tr>
      <td class="name">${esc(p.name)}${p.sku ? ` <span style="color:var(--muted);font-weight:400;font-size:11px">${esc(p.sku)}</span>` : ""}</td>
      <td class="num">${qty(p.qty_received)}</td>
      <td class="num">${p.qty_returned ? qty(p.qty_returned) : "—"}</td>
      <td class="num" style="color:var(--text);font-weight:700">${qty(p.qty_net)}</td>
      <td>${esc(p.unit || "—")}</td>
      <td class="num">${money(p.avg_unit_price)}</td>
      <td class="num" style="color:var(--text);font-weight:700">${money(p.value_net)}</td>
      <td class="num" style="text-align:left">${esc(p.last_received)}</td>
    </tr>`).join("") : empty(8, "No products received yet.");

  document.getElementById("panel-products").innerHTML = `
    <div class="block">
      <div class="block-title">
        <span>Products you have received</span>
        <span>Returns already deducted &middot; amounts in EGP</span>
      </div>
      <div class="scroll"><table>
        <thead><tr><th>Product</th><th class="num">Received</th><th class="num">Returned</th>
          <th class="num">Net qty</th><th>Unit</th><th class="num">Avg price</th>
          <th class="num">Net value</th><th>Last delivery</th></tr></thead>
        <tbody>${rows}</tbody>
        ${d.products.length ? `<tfoot><tr>
          <td class="name">Total</td>
          <td class="num"></td><td class="num"></td>
          <td class="num" style="color:var(--text);font-weight:700">${qty(t.qty_net)}</td>
          <td></td><td class="num"></td>
          <td class="num" style="color:var(--text);font-weight:700">${money(t.value_net)}</td>
          <td></td>
        </tr></tfoot>` : ""}
      </table></div>
    </div>`;
}

function renderStock(d){
  // Consignment goods only, so the tab stays hidden for clients who buy outright.
  const items = d.stock || [];
  const st = d.stock_totals || {};
  const tab = document.getElementById("tab-stock");
  tab.style.display = items.length ? "" : "none";
  if(!items.length){
    document.getElementById("panel-stock").innerHTML = "";
    return;
  }
  const rows = items.map(p => `
    <tr>
      <td class="name">${esc(p.name)}${p.sku ? ` <span style="color:var(--muted);font-weight:400;font-size:11px">${esc(p.sku)}</span>` : ""}</td>
      <td class="num" style="color:var(--text);font-weight:700">${qty(p.qty_on_hand)}</td>
      <td>${esc(p.unit || "—")}</td>
      <td class="num">${money(p.unit_price)}</td>
      <td class="num" style="color:var(--text);font-weight:700">${money(p.value_on_hand)}</td>
      <td class="num">${qty(p.qty_sent)}</td>
      <td class="num">${p.qty_sold ? qty(p.qty_sold) : "—"}</td>
      <td class="num">${p.qty_returned ? qty(p.qty_returned) : "—"}</td>
      <td class="num" style="text-align:left">${esc(p.last_received)}</td>
    </tr>`).join("");

  document.getElementById("panel-stock").innerHTML = `
    <div class="block">
      <div class="block-title">
        <span>Goods you still hold on consignment</span>
        <span>Sold and returned already deducted &middot; amounts in EGP</span>
      </div>
      <div class="scroll"><table>
        <thead><tr><th>Product</th><th class="num">On hand</th><th>Unit</th>
          <th class="num">Unit price</th><th class="num">Value</th>
          <th class="num">Sent</th><th class="num">Sold</th><th class="num">Returned</th>
          <th>Last delivery</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr>
          <td class="name">Total</td>
          <td class="num" style="color:var(--text);font-weight:700">${qty(st.qty_on_hand)}</td>
          <td></td><td class="num"></td>
          <td class="num" style="color:var(--text);font-weight:700">${money(st.value_on_hand)}</td>
          <td class="num">${qty(st.qty_sent)}</td>
          <td class="num">${qty(st.qty_sold)}</td>
          <td class="num">${qty(st.qty_returned)}</td>
          <td></td>
        </tr></tfoot>
      </table></div>
    </div>
    <div class="block">
      <div class="block-title"><span>You are billed for what you sell, not for what you hold</span></div>
      <div class="empty" style="text-align:left;color:var(--sub)">
        These goods are with you on consignment. Anything still on hand can be sold or
        returned — tell us the quantities you sold and we will record the payment against them.
      </div>
    </div>`;
}

function renderDeliveries(d){
  const body = d.deliveries.length ? d.deliveries.map(x => `
    <div class="delivery">
      <div class="delivery-head">
        <div>
          <span class="delivery-ref">${esc(x.ref)}</span>
          <span class="tag ${x.kind === "return" ? "tag-return" : "tag-delivery"}" style="margin-left:8px">${esc(x.label)}</span>
        </div>
        <div style="text-align:right">
          <div class="delivery-date">${esc(x.date_str)}</div>
          <div style="font-family:var(--mono);font-weight:700">${money(x.total)}</div>
        </div>
      </div>
      <div class="delivery-items">
        ${x.items.map(i => `<div class="delivery-item">
          <span>${esc(i.product)}</span>
          <span class="q">${qty(i.qty)} ${esc(i.unit)} &times; ${money(i.unit_price)} = ${money(i.total)}</span>
        </div>`).join("")}
      </div>
    </div>`).join("") : `<div class="empty">No deliveries recorded yet.</div>`;

  document.getElementById("panel-deliveries").innerHTML = `
    <div class="block">
      <div class="block-title"><span>Every delivery and return, newest first</span></div>
      ${body}
    </div>`;
}

async function load(){
  try{
    const res = await fetch(DATA_URL, { cache: "no-store" });
    if(!res.ok) throw new Error(res.status === 404
      ? "This link is no longer active. Please ask us for a new one."
      : "Could not load your account right now.");
    const d = await res.json();

    document.getElementById("error").style.display = "none";
    document.getElementById("client-name").textContent = d.client.name;
    document.title = "Account — " + d.client.name;
    const bits = [d.client.code, d.client.payment_terms ? d.client.payment_terms.replace(/_/g, " ") : "", d.client.phone]
      .filter(v => v && v !== "—");
    document.getElementById("client-meta").textContent = bits.join("  ·  ");
    document.getElementById("generated").textContent = "Last updated " + d.generated_at;

    renderCards(d);
    renderStatement(d);
    renderProducts(d);
    renderStock(d);
    renderDeliveries(d);
  } catch(e){
    document.getElementById("error").style.display = "block";
    document.getElementById("error-text").textContent = e.message || "Could not load your account.";
  }
}

document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b === btn));
    document.querySelectorAll(".panel").forEach(p =>
      p.classList.toggle("active", p.id === "panel-" + btn.dataset.panel));
  });
});

load();
// Live: refresh on a timer, and immediately when the tab regains focus so a
// client who leaves it open never reads a stale balance.
setInterval(load, 60000);
document.addEventListener("visibilitychange", () => { if(!document.hidden) load(); });
</script>
</body>
</html>""".replace("__CLIENT__", escape(client_name or "Account"))
