from contextlib import asynccontextmanager
from pathlib import Path

from urllib.parse import quote

from fastapi import Depends, FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from starlette.middleware.trustedhost import TrustedHostMiddleware
from starlette_csrf import CSRFMiddleware

from slowapi import _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded

from app.api.routes import ROUTERS
from app.core.config import settings
from app.core.log import configure_logging, logger
from app.core.middleware import RequestLoggingMiddleware, SecurityHeadersMiddleware
from app.core.migrations import verify_migration_status
from app.core.monitoring import configure_monitoring
from app.core.rate_limit import limiter
from app.database import get_async_session


async def ensure_payroll_columns() -> None:
    """Self-healing guard: ensure recently-added payroll columns exist even if
    the alembic migration hasn't applied yet. Idempotent and safe to run on
    every startup. Without these columns, every query against the payroll table
    fails (e.g. the Payroll tab hangs on 'Loading...')."""
    from sqlalchemy import text
    from app.db.session import AsyncSessionLocal

    statements = [
        "ALTER TABLE payroll ADD COLUMN IF NOT EXISTS paid_amount NUMERIC(12,2)",
        "ALTER TABLE payroll ADD COLUMN IF NOT EXISTS days_off_credited NUMERIC(8,2) NOT NULL DEFAULT 0",
    ]
    try:
        async with AsyncSessionLocal() as db:
            ok = 0
            for stmt in statements:
                # One statement per transaction: on Postgres, a single failed
                # statement poisons the whole transaction (InFailedSQLTransaction)
                # and would silently kill every statement after it.
                try:
                    await db.execute(text(stmt))
                    await db.commit()
                    ok += 1
                except Exception:
                    await db.rollback()
                    logger.exception("ensure_payroll_columns: statement failed: %s", stmt)
            if ok == len(statements):
                logger.info("ensure_payroll_columns: payroll columns ready")
            else:
                logger.error(
                    "ensure_payroll_columns: only %d/%d statements succeeded — "
                    "payroll pages may fail; see tracebacks above", ok, len(statements)
                )
    except Exception:
        logger.exception("ensure_payroll_columns: failed (could not open DB session)")


async def ensure_price_precision() -> None:
    """Self-healing guard: widen per-UNIT price/cost columns to NUMERIC(12,3).

    Lets unit prices carry 3 decimals (e.g. 0.065 EGP per gram). Only per-unit
    columns are widened — line totals and cash amounts stay at 2 decimals.
    Idempotent: ALTER ... TYPE NUMERIC(12,3) is a no-op if already (12,3).
    Each statement runs in its own transaction so one failure can't poison
    the rest (mirrors the other guards).
    """
    from sqlalchemy import text
    from app.db.session import AsyncSessionLocal

    targets = [
        ("products", "price"),
        ("products", "cost"),
        ("invoice_items", "unit_price"),
        ("retail_refund_items", "unit_price"),
        ("product_receipts", "unit_cost"),
        ("purchase_items", "unit_cost"),
    ]
    try:
        async with AsyncSessionLocal() as db:
            ok = 0
            for table, column in targets:
                try:
                    # to_regclass is NULL when the table doesn't exist → skip.
                    reg = (await db.execute(
                        text("SELECT to_regclass(:t)"), {"t": table}
                    )).scalar()
                    if reg is None:
                        continue
                    await db.execute(text(
                        f'ALTER TABLE {table} ALTER COLUMN {column} TYPE NUMERIC(12,3)'
                    ))
                    await db.commit()
                    ok += 1
                except Exception:
                    await db.rollback()
                    logger.exception(
                        "ensure_price_precision: failed for %s.%s", table, column)
            logger.info("ensure_price_precision: %d/%d unit-price columns at (12,3)",
                        ok, len(targets))
    except Exception:
        logger.exception("ensure_price_precision: failed (could not open DB session)")


async def ensure_delivery_transport_columns() -> None:
    """Self-healing guard: transport columns on farm_deliveries used by the
    carbon module to auto-log Scope 1 delivery transport emissions. Idempotent
    and safe to run on every startup (mirrors migration 20260611_0036)."""
    from sqlalchemy import text
    from app.db.session import AsyncSessionLocal

    statements = [
        "ALTER TABLE farm_deliveries ADD COLUMN IF NOT EXISTS distance_km NUMERIC(8,1)",
        "ALTER TABLE farm_deliveries ADD COLUMN IF NOT EXISTS vehicle_type VARCHAR(20)",
        # Average weight of one piece in kg — lets piece/bunch products be
        # counted in mass-based carbon metrics (mirrors migration 20260611_0037).
        "ALTER TABLE products ADD COLUMN IF NOT EXISTS unit_weight_kg NUMERIC(8,3)",
    ]
    try:
        async with AsyncSessionLocal() as db:
            ok = 0
            for stmt in statements:
                try:
                    await db.execute(text(stmt))
                    await db.commit()
                    ok += 1
                except Exception:
                    await db.rollback()
                    logger.exception("ensure_delivery_transport_columns: statement failed: %s", stmt)
            if ok == len(statements):
                logger.info("ensure_delivery_transport_columns: delivery transport columns ready")
            else:
                logger.error(
                    "ensure_delivery_transport_columns: only %d/%d statements succeeded — "
                    "delivery transport logging may fail; see tracebacks above", ok, len(statements)
                )
    except Exception:
        logger.exception("ensure_delivery_transport_columns: failed (could not open DB session)")


async def ensure_product_categories_table() -> None:
    """Self-healing guard for standalone product categories."""
    from app.db.session import AsyncSessionLocal

    statement = """
        CREATE TABLE IF NOT EXISTS product_categories (
            id SERIAL PRIMARY KEY,
            name VARCHAR(100) NOT NULL UNIQUE,
            created_at TIMESTAMPTZ DEFAULT now()
        )
    """
    indexes = [
        "CREATE INDEX IF NOT EXISTS ix_product_categories_id ON product_categories (id)",
        "CREATE UNIQUE INDEX IF NOT EXISTS ix_product_categories_name ON product_categories (name)",
    ]
    try:
        async with AsyncSessionLocal() as db:
            try:
                await db.execute(text(statement))
                for index_statement in indexes:
                    await db.execute(text(index_statement))
                await db.commit()
                logger.info("ensure_product_categories_table: product_categories ready")
            except Exception:
                await db.rollback()
                logger.exception("ensure_product_categories_table: failed")
    except Exception:
        logger.exception("ensure_product_categories_table: failed (could not open DB session)")


async def ensure_consignment_sales_tables() -> None:
    """Self-healing guard for consignment sale records (sold items captured
    with each consignment payment). Mirrors migration 20260724_0043 so the
    tables exist even if alembic is blocked (e.g. by a multi-head conflict).
    Also mirrors 20260831_0046 (subtotal / discount on each payment).
    Idempotent — CREATE / ALTER ... IF NOT EXISTS is safe on every startup."""
    from app.db.session import AsyncSessionLocal

    statements = [
        """
        CREATE TABLE IF NOT EXISTS consignment_sales (
            id SERIAL PRIMARY KEY,
            client_id INTEGER NOT NULL REFERENCES b2b_clients(id),
            user_id INTEGER REFERENCES users(id),
            journal_id INTEGER REFERENCES journals(id),
            month_label VARCHAR(100),
            subtotal NUMERIC(14, 2) DEFAULT 0,
            discount NUMERIC(14, 2) DEFAULT 0,
            amount NUMERIC(14, 2) DEFAULT 0,
            notes TEXT,
            created_at TIMESTAMPTZ DEFAULT now()
        )
        """,
        # Mirrors migration 20260831_0046 for databases created before the
        # client discount was applied to consignment payments.
        "ALTER TABLE consignment_sales ADD COLUMN IF NOT EXISTS subtotal NUMERIC(14, 2) DEFAULT 0",
        "ALTER TABLE consignment_sales ADD COLUMN IF NOT EXISTS discount NUMERIC(14, 2) DEFAULT 0",
        "UPDATE consignment_sales SET subtotal = COALESCE(amount, 0) WHERE subtotal IS NULL OR subtotal = 0",
        """
        CREATE TABLE IF NOT EXISTS consignment_sale_items (
            id SERIAL PRIMARY KEY,
            sale_id INTEGER NOT NULL REFERENCES consignment_sales(id),
            product_id INTEGER NOT NULL REFERENCES products(id),
            qty NUMERIC(12, 3) NOT NULL,
            unit_price NUMERIC(14, 2) NOT NULL,
            total NUMERIC(14, 2) NOT NULL
        )
        """,
        "CREATE INDEX IF NOT EXISTS ix_consignment_sales_client_id ON consignment_sales (client_id)",
        "CREATE INDEX IF NOT EXISTS ix_consignment_sale_items_sale_id ON consignment_sale_items (sale_id)",
    ]
    try:
        async with AsyncSessionLocal() as db:
            for stmt in statements:
                try:
                    await db.execute(text(stmt))
                    await db.commit()
                except Exception:
                    await db.rollback()
                    logger.exception("ensure_consignment_sales_tables: statement failed: %s", stmt.split("(")[0].strip())
            logger.info("ensure_consignment_sales_tables: consignment sale tables ready")
    except Exception:
        logger.exception("ensure_consignment_sales_tables: failed (could not open DB session)")


async def ensure_b2b_portal_columns() -> None:
    """Self-healing guard for the shareable client-portal link on b2b_clients.
    Mirrors migration 20260812_0044 so a client link keeps working even if
    alembic is blocked. Idempotent — ADD COLUMN IF NOT EXISTS on every boot."""
    from sqlalchemy import text
    from app.db.session import AsyncSessionLocal

    statements = [
        "ALTER TABLE b2b_clients ADD COLUMN IF NOT EXISTS portal_token VARCHAR(64)",
        "ALTER TABLE b2b_clients ADD COLUMN IF NOT EXISTS portal_enabled BOOLEAN DEFAULT FALSE",
        "ALTER TABLE b2b_clients ADD COLUMN IF NOT EXISTS portal_created_at TIMESTAMPTZ",
        "ALTER TABLE b2b_clients ADD COLUMN IF NOT EXISTS portal_last_viewed_at TIMESTAMPTZ",
        "ALTER TABLE b2b_clients ADD COLUMN IF NOT EXISTS portal_view_count INTEGER DEFAULT 0",
        # The token is the only credential on the portal URL — two clients must
        # never be able to share one.
        "CREATE UNIQUE INDEX IF NOT EXISTS ix_b2b_clients_portal_token "
        "ON b2b_clients (portal_token)",
    ]
    try:
        async with AsyncSessionLocal() as db:
            for stmt in statements:
                try:
                    await db.execute(text(stmt))
                    await db.commit()
                except Exception:
                    await db.rollback()
                    logger.exception("ensure_b2b_portal_columns: statement failed: %s", stmt)
            logger.info("ensure_b2b_portal_columns: client portal columns ready")
    except Exception:
        logger.exception("ensure_b2b_portal_columns: failed (could not open DB session)")


async def ensure_carbon_methodology() -> None:
    """Self-healing guard for the carbon module's methodology upgrade.

    1. Adds GHG Protocol provenance columns to carbon_emission_factors
       (scope, methodology_source, source_year, region) if missing.
    2. Seeds a set of documented default emission factors (inserted only
       when the source_key does not already exist — never overwrites
       user-edited values).
    3. Backfills methodology fields on known seeded keys where they are
       still NULL, so pre-existing installs gain citations too.

    Idempotent and safe to run on every startup. Default factor values are
    indicative figures from widely used public datasets (DEFRA GHG
    Conversion Factors; IFI Harmonised Grid Emission Factors for the Egypt
    grid) and are fully editable in the Factors admin page.
    """
    from sqlalchemy import text
    from app.db.session import AsyncSessionLocal

    # DDL to (re)create the carbon tables if alembic migration
    # 20260510_0018_carbon_footprint never applied. entrypoint.sh tolerates a
    # failed `alembic upgrade head`, so a stuck migration chain leaves these
    # tables missing entirely — in which case every ALTER/INSERT below would
    # fail with UndefinedTable. Mirrors app/models/carbon.py exactly.
    # The 0018 migration is now idempotent, so a later successful alembic run
    # will not collide with tables created here.
    table_statements = [
        """CREATE TABLE IF NOT EXISTS carbon_emission_factors (
               id SERIAL PRIMARY KEY,
               source_type VARCHAR(30) NOT NULL,
               source_key VARCHAR(60) NOT NULL UNIQUE,
               label VARCHAR(150) NOT NULL,
               factor_kg_co2e_per_unit NUMERIC(10,4) NOT NULL,
               unit VARCHAR(30) NOT NULL,
               description TEXT,
               is_active BOOLEAN NOT NULL DEFAULT TRUE,
               created_at TIMESTAMPTZ DEFAULT now()
           )""",
        """CREATE TABLE IF NOT EXISTS carbon_logs (
               id SERIAL PRIMARY KEY,
               factor_id INTEGER NOT NULL REFERENCES carbon_emission_factors(id),
               farm_id INTEGER REFERENCES farms(id),
               user_id INTEGER REFERENCES users(id),
               log_date DATE NOT NULL,
               quantity NUMERIC(14,3) NOT NULL,
               kg_co2e NUMERIC(14,4) NOT NULL,
               ref_type VARCHAR(40),
               ref_id INTEGER,
               notes TEXT,
               created_at TIMESTAMPTZ DEFAULT now()
           )""",
        "CREATE INDEX IF NOT EXISTS ix_carbon_logs_log_date ON carbon_logs (log_date)",
        "CREATE INDEX IF NOT EXISTS ix_carbon_logs_ref ON carbon_logs (ref_type, ref_id)",
        "CREATE INDEX IF NOT EXISTS ix_carbon_logs_farm_id ON carbon_logs (farm_id)",
        """CREATE TABLE IF NOT EXISTS carbon_targets (
               id SERIAL PRIMARY KEY,
               label VARCHAR(150) NOT NULL,
               period_start DATE NOT NULL,
               period_end DATE NOT NULL,
               target_kg_co2e NUMERIC(14,2) NOT NULL,
               notes TEXT,
               created_at TIMESTAMPTZ DEFAULT now()
           )""",
    ]

    column_statements = [
        "ALTER TABLE carbon_emission_factors ADD COLUMN IF NOT EXISTS scope INTEGER",
        "ALTER TABLE carbon_emission_factors ADD COLUMN IF NOT EXISTS methodology_source VARCHAR(200)",
        "ALTER TABLE carbon_emission_factors ADD COLUMN IF NOT EXISTS source_year INTEGER",
        "ALTER TABLE carbon_emission_factors ADD COLUMN IF NOT EXISTS region VARCHAR(80)",
    ]

    # (source_type, source_key, label, factor, unit, scope, methodology_source, source_year, region)
    DEFAULT_FACTORS = [
        ("energy", "diesel_liter", "Diesel fuel (per litre)", 2.68, "litre", 1,
         "DEFRA GHG Conversion Factors 2024 — diesel (average biofuel blend)", 2024, "Global default"),
        ("energy", "petrol_liter", "Petrol fuel (per litre)", 2.31, "litre", 1,
         "DEFRA GHG Conversion Factors 2024 — petrol (average biofuel blend)", 2024, "Global default"),
        ("energy", "lpg_liter", "LPG (per litre)", 1.56, "litre", 1,
         "DEFRA GHG Conversion Factors 2024 — LPG", 2024, "Global default"),
        ("energy", "electricity_kwh", "Grid electricity (per kWh)", 0.46, "kWh", 2,
         "IFI Harmonised Grid Emission Factors — Egypt national grid (indicative; verify against latest dataset)", 2023, "Egypt"),
        ("transport", "van_km", "Van / pickup transport (per km)", 0.23, "km", 1,
         "DEFRA GHG Conversion Factors 2024 — average van, diesel", 2024, "Global default"),
        ("transport", "truck_km", "Rigid truck transport (per km)", 0.81, "km", 1,
         "DEFRA GHG Conversion Factors 2024 — rigid HGV, average laden", 2024, "Global default"),
        ("waste", "organic_waste_kg", "Organic waste to landfill (per kg)", 0.58, "kg", 3,
         "DEFRA GHG Conversion Factors 2024 — organic/food waste to landfill (indicative)", 2024, "Global default"),
        ("waste", "compost_kg", "Organic waste composted (per kg)", 0.01, "kg", 3,
         "DEFRA GHG Conversion Factors 2024 — open-loop composting", 2024, "Global default"),
        ("production", "processing_kwh", "Processing energy (per kWh)", 0.46, "kWh", 2,
         "Grid electricity factor — Egypt (IFI Harmonised Grid Emission Factors)", 2023, "Egypt"),
        # ── Livestock (IPCC 2006 Tier 1 enteric CH₄, Africa defaults; ────
        # kg CH₄/head/yr × GWP₁₀₀ 28 (AR5) ÷ 365 → kg CO₂e per head-day.
        # cattle 31 kg CH₄/yr → 2.38 · sheep/goats 5 kg CH₄/yr → 0.38.
        # Poultry have no enteric factor; their small manure-management
        # CH₄ (0.02 kg/yr) is included so poultry groups aren't silently 0.
        ("livestock", "enteric_cattle_head_day", "Cattle — enteric methane (per head-day)", 2.38, "head-day", 1,
         "IPCC 2006 GL Vol.4 Tab.10.11 — other cattle, Africa; CH4 GWP100=28 (AR5)", 2006, "Africa default"),
        ("livestock", "enteric_sheep_head_day", "Sheep — enteric methane (per head-day)", 0.38, "head-day", 1,
         "IPCC 2006 GL Vol.4 Tab.10.10 — sheep, developing regions; CH4 GWP100=28 (AR5)", 2006, "Africa default"),
        ("livestock", "enteric_goat_head_day", "Goats — enteric methane (per head-day)", 0.38, "head-day", 1,
         "IPCC 2006 GL Vol.4 Tab.10.10 — goats, developing regions; CH4 GWP100=28 (AR5)", 2006, "Africa default"),
        ("livestock", "poultry_manure_head_day", "Poultry — manure methane (per head-day)", 0.0015, "head-day", 1,
         "IPCC 2006 GL Vol.4 Tab.10.14 — poultry manure CH4, warm climate; GWP100=28 (AR5)", 2006, "Africa default"),
    ]

    # ── Phase 0: do the carbon tables even exist? ─────────────────────────
    try:
        async with AsyncSessionLocal() as db:
            res = await db.execute(text(
                "SELECT to_regclass('carbon_emission_factors'), "
                "       to_regclass('carbon_logs'), "
                "       to_regclass('carbon_targets')"
            ))
            factors_t, logs_t, targets_t = res.one()
            missing = [name for name, reg in (
                ("carbon_emission_factors", factors_t),
                ("carbon_logs", logs_t),
                ("carbon_targets", targets_t),
            ) if reg is None]
            if missing:
                logger.warning(
                    "ensure_carbon_methodology: carbon table(s) MISSING: %s — "
                    "alembic migration 20260510_0018 likely never (fully) applied "
                    "(entrypoint tolerates failed upgrades). Creating now.",
                    ", ".join(missing),
                )
                # All statements are CREATE ... IF NOT EXISTS, so running the
                # full set is safe even when only one table is missing.
                for stmt in table_statements:
                    try:
                        await db.execute(text(stmt))
                        await db.commit()
                    except Exception:
                        await db.rollback()
                        logger.exception(
                            "ensure_carbon_methodology: table create failed: %s",
                            stmt.split("(")[0].strip(),
                        )
    except Exception:
        logger.exception("ensure_carbon_methodology: table existence check failed")

    # ── Phase 1: methodology columns (one statement per transaction so a
    #    single failure cannot poison the rest — Postgres aborts the whole
    #    transaction on the first error otherwise) ──────────────────────────
    try:
        async with AsyncSessionLocal() as db:
            for stmt in column_statements:
                try:
                    await db.execute(text(stmt))
                    await db.commit()
                except Exception:
                    await db.rollback()
                    logger.exception("ensure_carbon_methodology: column add failed: %s", stmt)
    except Exception:
        logger.exception("ensure_carbon_methodology: column add failed (could not open DB session)")

    # ── Phase 2: seed + backfill, one factor per transaction ──────────────
    try:
        async with AsyncSessionLocal() as db:
            failed_keys: list[str] = []
            for (stype, skey, label, factor, unit, scope, msrc, syear, region) in DEFAULT_FACTORS:
                try:
                    await db.execute(text("""
                        INSERT INTO carbon_emission_factors
                            (source_type, source_key, label, factor_kg_co2e_per_unit, unit,
                             scope, methodology_source, source_year, region, is_active)
                        SELECT :stype, :skey, :label, :factor, :unit,
                               :scope, :msrc, :syear, :region, TRUE
                        WHERE NOT EXISTS (
                            -- CAST is required: asyncpg deduces conflicting types
                            -- (varchar vs text) when :skey is reused both as an
                            -- inserted value and in a comparison, and fails with
                            -- AmbiguousParameterError. This was the root cause of
                            -- the silent 9/9 seed failure at startup.
                            SELECT 1 FROM carbon_emission_factors
                             WHERE source_key = CAST(:skey AS VARCHAR(60))
                        )
                    """), {"stype": stype, "skey": skey, "label": label, "factor": factor,
                           "unit": unit, "scope": scope, "msrc": msrc, "syear": syear, "region": region})

                    # Backfill methodology on pre-existing rows with the same key
                    # (never touches factor values or labels the user may have edited).
                    await db.execute(text("""
                        UPDATE carbon_emission_factors
                           SET scope = COALESCE(scope, :scope),
                               methodology_source = COALESCE(methodology_source, :msrc),
                               source_year = COALESCE(source_year, :syear),
                               region = COALESCE(region, :region)
                         WHERE source_key = :skey
                    """), {"skey": skey, "scope": scope, "msrc": msrc, "syear": syear, "region": region})
                    await db.commit()
                except Exception:
                    # One bad row must not abort the rest — roll back just this
                    # statement's transaction and continue with the next factor.
                    await db.rollback()
                    failed_keys.append(skey)
                    logger.exception("ensure_carbon_methodology: seed failed for key=%s", skey)
            if failed_keys:
                logger.error(
                    "ensure_carbon_methodology: %d/%d factor seeds FAILED (%s) — "
                    "see tracebacks above",
                    len(failed_keys), len(DEFAULT_FACTORS), ", ".join(failed_keys),
                )
            else:
                logger.info("ensure_carbon_methodology: carbon methodology columns and default factors ready")
    except Exception:
        logger.exception("ensure_carbon_methodology: factor seed failed")


async def sync_livestock_emissions_on_boot() -> None:
    """Generate any missing monthly livestock (enteric methane) carbon logs.
    One log per animal group per complete month; idempotent; never blocks
    startup — a failure is logged and the next boot (or the carbon backfill
    endpoint) catches up."""
    try:
        from app.db.session import AsyncSessionLocal
        from app.services.livestock_carbon_service import sync_livestock_carbon_logs
        async with AsyncSessionLocal() as db:
            created = await sync_livestock_carbon_logs(db)
            if created:
                logger.info("sync_livestock_emissions_on_boot: created %d monthly livestock logs", created)
    except Exception:
        logger.exception("sync_livestock_emissions_on_boot: failed")


async def seed_chart_of_accounts() -> None:
    """Ensure core accounting accounts exist. Safe to run on every startup."""
    from decimal import Decimal
    from sqlalchemy import select
    from app.db.session import AsyncSessionLocal
    from app.models.accounting import Account

    CORE_ACCOUNTS = [
        ("1000", "Cash",              "asset"),
        ("1100", "Accounts Receivable", "asset"),
        ("2000", "Accounts Payable",  "liability"),
        ("2200", "Deferred Revenue",  "liability"),
        ("4000", "Revenue",           "revenue"),
        ("5000", "Cost of Goods Sold","expense"),
        ("6000", "Expenses",          "expense"),
    ]
    try:
        async with AsyncSessionLocal() as db:
            for code, name, atype in CORE_ACCOUNTS:
                r = await db.execute(select(Account).where(Account.code == code))
                if r.scalar_one_or_none() is None:
                    db.add(Account(code=code, name=name, type=atype, balance=Decimal("0")))
            await db.commit()
            logger.info("seed_chart_of_accounts: core accounts ready")
    except Exception:
        logger.exception("seed_chart_of_accounts: failed — journals will not post until accounts exist")

STATIC_DIR = Path(__file__).resolve().parent / "static"

# ── HTML shown when an unhandled 500 occurs during an HTML page navigation ─
_ERROR_HTML_PAGE = """<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
<title>Something went wrong</title>
<style>
  :root{--card:rgba(15,20,36,0.88);--border:rgba(255,255,255,0.08);--text:#fff;--sub:#8899bb}
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Segoe UI',sans-serif;min-height:100vh;display:flex;align-items:center;
       justify-content:center;color:var(--text);padding:24px;
       background:linear-gradient(rgba(6,8,16,.68),rgba(6,8,16,.68)),
                 url('/static/home1.jpg.jpeg') center/cover no-repeat}
  .box{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:40px;
       width:360px;backdrop-filter:blur(8px);box-shadow:0 24px 60px rgba(0,0,0,.35);text-align:center}
  h2{color:#ff4d6d;font-size:22px;margin-bottom:12px}
  p{color:var(--sub);font-size:14px;margin-bottom:28px}
  a{display:inline-block;padding:12px 28px;background:linear-gradient(135deg,#00ff9d,#00d4ff);
    border-radius:10px;color:#021a10;font-weight:800;font-size:14px;text-decoration:none}
  a:hover{filter:brightness(1.1)}
</style>
</head><body>
<div class="box">
  <h2>Something went wrong</h2>
  <p>An unexpected error occurred. Please try again or go back.</p>
  <a href="javascript:history.back()">Go back</a>
</div>
</body></html>"""


async def _try_silent_refresh(refresh_token_value: str) -> tuple[str, str] | None:
    """
    Open a fresh DB session and attempt to rotate the given raw refresh token.
    Returns ``(access_token, refresh_token)`` on success, or None on any
    failure, so callers can safely fall through to a login redirect without
    raising.

    The refresh token is rotated server-side on success, so callers MUST write
    both returned cookies back to the browser. Defined at module level (not
    inside create_app) so tests can monkeypatch
    ``app.app_factory._try_silent_refresh`` without a real DB connection.
    """
    from app.db.session import AsyncSessionLocal
    from app.core import security
    try:
        async with AsyncSessionLocal() as db:
            return await security.try_refresh_access_token(db, refresh_token_value)
    except Exception:
        return None


@asynccontextmanager
async def lifespan(_: FastAPI):
    configure_logging()
    configure_monitoring()
    await verify_migration_status()
    await ensure_payroll_columns()
    await ensure_price_precision()
    await ensure_delivery_transport_columns()
    await ensure_product_categories_table()
    await ensure_consignment_sales_tables()
    await ensure_b2b_portal_columns()
    await ensure_carbon_methodology()
    await sync_livestock_emissions_on_boot()
    await seed_chart_of_accounts()
    from app.core.cache import init_redis_pool, close_redis_pool
    await init_redis_pool()
    yield
    await close_redis_pool()


def create_app() -> FastAPI:
    app = FastAPI(
        title=settings.APP_NAME,
        debug=settings.DEBUG,
        lifespan=lifespan,
    )

    app.state.limiter = limiter
    app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

    STATIC_DIR.mkdir(parents=True, exist_ok=True)
    (STATIC_DIR / "uploads").mkdir(parents=True, exist_ok=True)
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.CORS_ALLOW_ORIGINS,
        allow_credentials=settings.CORS_ALLOW_CREDENTIALS,
        allow_methods=settings.CORS_ALLOW_METHODS,
        allow_headers=settings.CORS_ALLOW_HEADERS,
    )
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=settings.ALLOWED_HOSTS)
    app.add_middleware(SecurityHeadersMiddleware)
    # CSRF protection: only triggers on requests that carry the access_token cookie.
    # Auth endpoints (/auth/*) are exempt — they use credentials as proof, not a session.
    # Pure JSON API calls (/*/api/*) are also exempt — protected by CORS same-origin policy.
    import re
    app.add_middleware(
        CSRFMiddleware,
        secret=settings.SECRET_KEY,
        sensitive_cookies={"access_token"},
        exempt_urls=[
            re.compile(r"^/auth/.*"),
            re.compile(r".*/api/.*"),
            re.compile(r"^/hr/clear-data$"),
            re.compile(r"^/import/.*"),
            re.compile(r"^/invoice.*"),
            re.compile(r"^/health.*"),
        ],
    )
    app.add_middleware(RequestLoggingMiddleware)

    @app.exception_handler(Exception)
    async def unhandled_exception_handler(request: Request, exc: Exception):
        logger.exception(
            "Unhandled application error",
            exc_info=exc,
            extra={
                "method": request.method,
                "path": request.url.path,
                "query": str(request.url.query) or None,
            },
        )
        # Return a styled HTML page for browser navigations so users never see
        # raw JSON from an unhandled 500.  API callers (JSON Accept) still get
        # the machine-readable JSON body.
        if (
            request.method == "GET"
            and "text/html" in request.headers.get("accept", "")
        ):
            return HTMLResponse(content=_ERROR_HTML_PAGE, status_code=500)
        return JSONResponse(
            status_code=500,
            content={"detail": "An internal server error occurred."},
        )

    # ── Session-expiry middleware ────────────────────────────────────────────
    # Intercepts 401 responses on HTML GET navigations (e.g. the browser
    # requests /dashboard after the access token has expired):
    #
    #  • If a refresh_token cookie is present → attempt a silent refresh and
    #    307-redirect back to the same URL with the new access_token cookie so
    #    the browser retries with a fresh token.
    #  • Otherwise → 307-redirect to /?next=<path>&reason=expired so the login
    #    page can show a friendly "session expired" message and bounce the user
    #    back after sign-in.
    #
    # Only HTML GETs are rewritten.  JSON/API callers and POST/PUT/DELETE
    # requests keep receiving the plain 401 JSON so auth-guard.js keeps
    # working and API clients are unaffected.
    # /auth/* and /health* are explicitly excluded.
    @app.middleware("http")
    async def _session_expiry(request: Request, call_next):
        response = await call_next(request)

        if (
            response.status_code == 401
            and request.method == "GET"
            and "text/html" in request.headers.get("accept", "")
            and not request.url.path.startswith("/auth/")
            and not request.url.path.startswith("/health")
        ):
            # Reconstruct the original path + query so ?next= roundtrips.
            path = request.url.path
            if request.url.query:
                path += "?" + request.url.query

            refresh_token_value = request.cookies.get("refresh_token")
            if refresh_token_value:
                refreshed = await _try_silent_refresh(refresh_token_value)
                if refreshed:
                    new_access_token, new_refresh_token = refreshed
                    # Redirect to the same URL; the new cookie rides along so
                    # the retry succeeds without another round-trip.
                    redirect = RedirectResponse(url=path, status_code=307)
                    redirect.set_cookie(
                        key="access_token",
                        value=new_access_token,
                        httponly=True,
                        samesite="lax",
                        secure=settings.COOKIE_SECURE,
                        path="/",
                        max_age=settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60,
                    )
                    redirect.set_cookie(
                        key="logged_in",
                        value="true",
                        httponly=False,
                        samesite="lax",
                        secure=settings.COOKIE_SECURE,
                        path="/",
                        max_age=settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60,
                    )
                    redirect.set_cookie(
                        key="refresh_token",
                        value=new_refresh_token,
                        httponly=True,
                        samesite="lax",
                        secure=settings.COOKIE_SECURE,
                        path="/",
                        max_age=settings.REFRESH_TOKEN_EXPIRE_DAYS * 86400,
                    )
                    return redirect

            # No valid refresh token — send to login with context.
            login_url = "/?next=" + quote(path, safe="") + "&reason=expired"
            return RedirectResponse(url=login_url, status_code=307)

        return response

    for router in ROUTERS:
        app.include_router(router)

    @app.get("/health/live")
    async def liveness():
        return {"status": "ok"}

    @app.get("/health/ready")
    async def readiness(db: AsyncSession = Depends(get_async_session)):
        try:
            await db.execute(text("SELECT 1"))
            return {"status": "ok", "db": "ok"}
        except Exception:
            return JSONResponse(
                status_code=503,
                content={"status": "error", "db": "unreachable"},
            )

    # Backward-compat alias
    @app.get("/health")
    async def health():
        return {"status": "ok", "app": settings.APP_NAME, "environment": settings.APP_ENV}

    logger.info("Application configured")
    return app