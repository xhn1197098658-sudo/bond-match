-- Database schema for Bond Buyer Match application

-- Issuers table: Stores information about bond issuers
CREATE TABLE issuers (
    issuer_id INTEGER PRIMARY KEY AUTOINCREMENT,
    issuer_name TEXT NOT NULL,
    issuer_code TEXT,
    credit_rating TEXT,
    industry TEXT,
    additional_info TEXT
);

-- Bonds table: Stores individual bond information
CREATE TABLE bonds (
    bond_id INTEGER PRIMARY KEY AUTOINCREMENT,
    bond_code TEXT NOT NULL UNIQUE,
    issuer_id INTEGER NOT NULL,
    bond_name TEXT,
    issue_date DATE,
    maturity_date DATE,
    coupon_rate REAL,
    FOREIGN KEY (issuer_id) REFERENCES issuers(issuer_id)
);

-- Buyside Companies table: Fund/asset management companies
CREATE TABLE buyside_companies (
    company_id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name TEXT NOT NULL,
    company_type TEXT,
    aum REAL,
    location TEXT,
    additional_info TEXT
);

-- Funds table: For fund companies with multiple funds
CREATE TABLE funds (
    fund_id INTEGER PRIMARY KEY AUTOINCREMENT,
    fund_name TEXT NOT NULL,
    company_id INTEGER NOT NULL,
    fund_type TEXT,
    fund_size REAL,
    additional_info TEXT,
    FOREIGN KEY (company_id) REFERENCES buyside_companies(company_id)
);

-- Import batches: track source file for each import (for batch delete)
CREATE TABLE import_batches (
    batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    data_type TEXT NOT NULL,
    row_count INTEGER DEFAULT 0,
    imported_at TEXT DEFAULT (datetime('now','localtime'))
);

-- Contacts table: Employee contact information
CREATE TABLE contacts (
    contact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL,
    fund_id INTEGER,
    name TEXT NOT NULL,
    position TEXT,
    email TEXT,
    phone TEXT,
    qt TEXT,
    qq TEXT,
    wechat TEXT,
    mobil TEXT,
    is_primary BOOLEAN DEFAULT 0,
    additional_info TEXT,
    batch_id INTEGER REFERENCES import_batches(batch_id),
    FOREIGN KEY (company_id) REFERENCES buyside_companies(company_id),
    FOREIGN KEY (fund_id) REFERENCES funds(fund_id)
);

-- Holdings table: Which bonds are held by which companies/funds
CREATE TABLE holdings (
    holding_id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL,
    fund_id INTEGER,
    bond_id INTEGER NOT NULL,
    amount REAL,
    purchase_date DATE,
    batch_id INTEGER REFERENCES import_batches(batch_id),
    FOREIGN KEY (company_id) REFERENCES buyside_companies(company_id),
    FOREIGN KEY (fund_id) REFERENCES funds(fund_id),
    FOREIGN KEY (bond_id) REFERENCES bonds(bond_id)
);

-- CanBuy Lists: Explicit can-buy relationships
CREATE TABLE can_buy_lists (
    list_id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL,
    issuer_id INTEGER NOT NULL,
    date_added DATE DEFAULT CURRENT_DATE,
    additional_criteria TEXT,
    batch_id INTEGER REFERENCES import_batches(batch_id),
    FOREIGN KEY (company_id) REFERENCES buyside_companies(company_id),
    FOREIGN KEY (issuer_id) REFERENCES issuers(issuer_id)
);

CREATE INDEX idx_bonds_issuer ON bonds(issuer_id);
CREATE INDEX idx_holdings_company ON holdings(company_id);
CREATE INDEX idx_holdings_fund ON holdings(fund_id);
CREATE INDEX idx_holdings_bond ON holdings(bond_id);
CREATE INDEX idx_can_buy_company ON can_buy_lists(company_id);
CREATE INDEX idx_can_buy_issuer ON can_buy_lists(issuer_id);
CREATE INDEX idx_contacts_company ON contacts(company_id);
CREATE INDEX idx_funds_company ON funds(company_id);
