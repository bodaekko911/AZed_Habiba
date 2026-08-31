from sqlalchemy import Column, Integer, String, Numeric, DateTime, ForeignKey, Text, Boolean, Date, UniqueConstraint
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from app.database import Base


class B2BClient(Base):
    __tablename__ = "b2b_clients"

    id              = Column(Integer, primary_key=True, index=True)
    name            = Column(String(200), nullable=False)
    contact_person  = Column(String(150))
    phone           = Column(String(50))
    email           = Column(String(150))
    address         = Column(String(300))
    payment_terms   = Column(String(50), default="immediate")  # immediate | net15 | net30 | net60 | consignment
    discount_pct    = Column(Numeric(6, 2), default=0)   # client-specific discount percentage
    credit_limit    = Column(Numeric(14,2), default=0)
    outstanding     = Column(Numeric(14,2), default=0)
    notes           = Column(Text)
    is_active       = Column(Boolean, default=True)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())

    # ── Client portal (shareable read-only link) ──
    # A client opens /portal/c/<portal_token> and sees their own live statement
    # and received-products list with no login. The token is the only secret,
    # so it is generated with secrets.token_urlsafe and can be revoked or
    # rotated per client at any time. portal_enabled is the kill switch:
    # revoking clears the token AND flips this to False.
    portal_token          = Column(String(64), unique=True, index=True, nullable=True)
    portal_enabled        = Column(Boolean, default=False)
    portal_created_at     = Column(DateTime(timezone=True), nullable=True)
    portal_last_viewed_at = Column(DateTime(timezone=True), nullable=True)
    portal_view_count     = Column(Integer, default=0)

    invoices        = relationship("B2BInvoice", back_populates="client")
    consignments    = relationship("Consignment", back_populates="client")


class B2BInvoice(Base):
    __tablename__ = "b2b_invoices"

    id              = Column(Integer, primary_key=True, index=True)
    invoice_number  = Column(String(30), unique=True, index=True)
    client_id       = Column(Integer, ForeignKey("b2b_clients.id"), nullable=False)
    user_id         = Column(Integer, ForeignKey("users.id"), nullable=True)
    invoice_type    = Column(String(20), nullable=False)  # full_payment | credit | consignment
    status          = Column(String(20), default="unpaid")  # unpaid | paid | partial | consignment
    payment_method  = Column(String(30))  # cash | transfer | —
    subtotal        = Column(Numeric(14,2), default=0)
    discount        = Column(Numeric(14,2), default=0)
    total           = Column(Numeric(14,2), default=0)
    amount_paid     = Column(Numeric(14,2), default=0)
    due_date        = Column(Date, nullable=True)
    notes           = Column(Text)
    import_batch_id = Column(String(64), nullable=True, index=True)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())

    client          = relationship("B2BClient", back_populates="invoices")
    user            = relationship("User")
    items           = relationship("B2BInvoiceItem", back_populates="invoice", cascade="all, delete-orphan")


class B2BInvoiceItem(Base):
    __tablename__ = "b2b_invoice_items"

    id          = Column(Integer, primary_key=True, index=True)
    invoice_id  = Column(Integer, ForeignKey("b2b_invoices.id"), nullable=False)
    product_id  = Column(Integer, ForeignKey("products.id"), nullable=False)
    qty         = Column(Numeric(12,3), nullable=False)
    unit_price  = Column(Numeric(14,2), nullable=False)
    total       = Column(Numeric(14,2), nullable=False)

    invoice     = relationship("B2BInvoice", back_populates="items")
    product     = relationship("Product")


class Consignment(Base):
    __tablename__ = "consignments"

    id              = Column(Integer, primary_key=True, index=True)
    ref_number      = Column(String(30), unique=True, index=True)
    client_id       = Column(Integer, ForeignKey("b2b_clients.id"), nullable=False)
    invoice_id      = Column(Integer, ForeignKey("b2b_invoices.id"), nullable=True)
    user_id         = Column(Integer, ForeignKey("users.id"), nullable=True)
    status          = Column(String(20), default="active")  # active | settled | closed
    notes           = Column(Text)
    import_batch_id = Column(String(64), nullable=True, index=True)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())
    settled_at      = Column(DateTime(timezone=True), nullable=True)

    client          = relationship("B2BClient", back_populates="consignments")
    invoice         = relationship("B2BInvoice")
    user            = relationship("User")
    items           = relationship("ConsignmentItem", back_populates="consignment", cascade="all, delete-orphan")


class ConsignmentItem(Base):
    __tablename__ = "consignment_items"

    id                  = Column(Integer, primary_key=True, index=True)
    consignment_id      = Column(Integer, ForeignKey("consignments.id"), nullable=False)
    product_id          = Column(Integer, ForeignKey("products.id"), nullable=False)
    qty_sent            = Column(Numeric(12,3), default=0)
    qty_sold            = Column(Numeric(12,3), default=0)
    qty_returned        = Column(Numeric(12,3), default=0)
    unit_price          = Column(Numeric(14,2), default=0)

    consignment         = relationship("Consignment", back_populates="items")
    product             = relationship("Product")


class ConsignmentSale(Base):
    """
    A recorded consignment-client payment together with the items the client
    reported sold for a given month. ``subtotal`` is the gross sum of the line
    items (qty × unit_price), ``discount`` is the client's agreed discount on
    that subtotal, and ``amount`` — what the client actually pays — is the two
    netted off. This is a bookkeeping record only — it does
    NOT modify consignment quantities or stock (that stays with the separate
    Settle flow). It exists so each payment can be reconciled against the
    specific items sold and the month they were sold in.
    """
    __tablename__ = "consignment_sales"

    id          = Column(Integer, primary_key=True, index=True)
    client_id   = Column(Integer, ForeignKey("b2b_clients.id"), nullable=False, index=True)
    user_id     = Column(Integer, ForeignKey("users.id"), nullable=True)
    journal_id  = Column(Integer, ForeignKey("journals.id"), nullable=True)
    month_label = Column(String(100))                 # e.g. "July 2026"
    subtotal    = Column(Numeric(14, 2), default=0)    # = sum of item totals (gross)
    discount    = Column(Numeric(14, 2), default=0)    # client discount on the subtotal
    amount      = Column(Numeric(14, 2), default=0)    # = subtotal - discount (collected)
    notes       = Column(Text)
    created_at  = Column(DateTime(timezone=True), server_default=func.now())

    client      = relationship("B2BClient")
    user        = relationship("User")
    items       = relationship("ConsignmentSaleItem", back_populates="sale", cascade="all, delete-orphan")


class ConsignmentSaleItem(Base):
    __tablename__ = "consignment_sale_items"

    id          = Column(Integer, primary_key=True, index=True)
    sale_id     = Column(Integer, ForeignKey("consignment_sales.id"), nullable=False, index=True)
    product_id  = Column(Integer, ForeignKey("products.id"), nullable=False)
    qty         = Column(Numeric(12, 3), nullable=False)
    unit_price  = Column(Numeric(14, 2), nullable=False)
    total       = Column(Numeric(14, 2), nullable=False)

    sale        = relationship("ConsignmentSale", back_populates="items")
    product     = relationship("Product")


class B2BRefund(Base):
    __tablename__ = "b2b_refunds"

    id              = Column(Integer, primary_key=True, index=True)
    refund_number   = Column(String(30), unique=True, index=True)
    client_id       = Column(Integer, ForeignKey("b2b_clients.id"), nullable=False)
    user_id         = Column(Integer, ForeignKey("users.id"), nullable=True)
    subtotal        = Column(Numeric(14,2), default=0)
    discount        = Column(Numeric(14,2), default=0)
    total           = Column(Numeric(14,2), default=0)
    notes           = Column(Text)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())

    client          = relationship("B2BClient")
    user            = relationship("User")
    items           = relationship("B2BRefundItem", back_populates="refund", cascade="all, delete-orphan")


class B2BRefundItem(Base):
    __tablename__ = "b2b_refund_items"

    id          = Column(Integer, primary_key=True, index=True)
    refund_id   = Column(Integer, ForeignKey("b2b_refunds.id"), nullable=False)
    product_id  = Column(Integer, ForeignKey("products.id"), nullable=False)
    qty         = Column(Numeric(12,3), nullable=False)
    unit_price  = Column(Numeric(14,2), nullable=False)
    total       = Column(Numeric(14,2), nullable=False)

    refund      = relationship("B2BRefund", back_populates="items")
    product     = relationship("Product")


class B2BClientPrice(Base):
    __tablename__ = "b2b_client_prices"
    __table_args__ = (UniqueConstraint("client_id", "product_id", name="uq_client_product_price"),)

    id          = Column(Integer, primary_key=True, index=True)
    client_id   = Column(Integer, ForeignKey("b2b_clients.id"), nullable=False)
    product_id  = Column(Integer, ForeignKey("products.id"), nullable=False)
    price       = Column(Numeric(14, 2), nullable=False)

    client      = relationship("B2BClient")
    product     = relationship("Product")