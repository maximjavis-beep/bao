"""SQLite 表结构 DDL"""

DDL_DECLARATIONS = """
CREATE TABLE IF NOT EXISTS declarations (
    id              TEXT PRIMARY KEY,
    declaration_id  TEXT NOT NULL,
    trade_mode      TEXT NOT NULL,
    domestic_shipper TEXT NOT NULL,
    shipper_code    TEXT DEFAULT '',
    overseas_consignee TEXT NOT NULL,
    destination_country TEXT NOT NULL,
    transport_mode  TEXT DEFAULT '',
    vessel_name     TEXT DEFAULT '',
    voyage_no       TEXT DEFAULT '',
    bill_of_lading TEXT DEFAULT '',
    port_of_loading TEXT DEFAULT '',
    port_of_discharge TEXT DEFAULT '',
    incoterm        TEXT NOT NULL,
    currency        TEXT DEFAULT 'USD',
    freight         REAL DEFAULT 0.0,
    insurance       REAL DEFAULT 0.0,
    package_type    TEXT DEFAULT '',
    package_count   INTEGER DEFAULT 0,
    contract_no     TEXT DEFAULT '',
    invoice_no      TEXT DEFAULT '',
    declaration_date TEXT NOT NULL,
    total_amount    REAL DEFAULT 0.0,
    total_net_weight REAL DEFAULT 0.0,
    total_gross_weight REAL DEFAULT 0.0,
    fob_amount      REAL DEFAULT 0.0,
    estimated_tax_rebate REAL DEFAULT 0.0,
    remarks         TEXT DEFAULT '',
    created_at      TEXT DEFAULT (datetime('now'))
);
"""

DDL_ITEMS = """
CREATE TABLE IF NOT EXISTS declaration_items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    declaration_fk TEXT NOT NULL REFERENCES declarations(id) ON DELETE CASCADE,
    seq         INTEGER NOT NULL,
    hs_code     TEXT NOT NULL,
    name_zh     TEXT NOT NULL,
    name_en     TEXT DEFAULT '',
    specs       TEXT DEFAULT '',
    quantity    REAL NOT NULL,
    unit        TEXT NOT NULL,
    unit_price  REAL NOT NULL,
    total_price REAL NOT NULL,
    currency    TEXT DEFAULT 'USD',
    net_weight  REAL NOT NULL,
    gross_weight REAL NOT NULL,
    origin      TEXT DEFAULT '中国',
    destination_country TEXT DEFAULT '',
    tax_rebate_rate REAL DEFAULT 0.0,
    declaration_elements TEXT DEFAULT ''
);
"""

DDL_TEMPLATES = """
CREATE TABLE IF NOT EXISTS templates (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL UNIQUE,
    direction   TEXT NOT NULL DEFAULT 'export',
    trade_mode  TEXT DEFAULT '',
    mapping     TEXT NOT NULL DEFAULT '{}',
    created_at  TEXT DEFAULT (datetime('now'))
);
"""

ALL_DDL = [DDL_DECLARATIONS, DDL_ITEMS, DDL_TEMPLATES]
