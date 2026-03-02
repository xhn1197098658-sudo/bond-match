import sqlite3
import os
import re
import pandas as pd
from pathlib import Path
import traceback

# Excel 表头：中文列名 -> 程序使用的英文列名（支持中文 Excel）
COLUMN_ALIASES = {
    "holdings": {
        "公司名称": "company_name", "机构名称": "company_name", "公司": "company_name",
        "基金名称": "fund_name", "基金": "fund_name", "产品名称": "fund_name",
        "发行人": "issuer_name", "发行人名称": "issuer_name",
        "债券代码": "bond_code", "代码": "bond_code",
        "上证代码": "bond_code_sse", "深证代码": "bond_code_sz", "其他通用代码": "bond_code_other",
        "债券名称": "bond_name", "债券": "bond_name", "持仓资产名称": "bond_name",
        "金额": "amount", "持仓金额": "amount", "数量": "amount",
        "资产规模": "amount", "amount": "amount",
        "资产规模（万元）": "amount", "资产规模（万元）/amount": "amount",
        "资产规模 (万元)": "amount", "资产规模(万元)": "amount",
        "机构类型": "company_type", "类型": "company_type",
    },
    "can_buy_lists": {
        "公司名称": "company_name", "机构名称": "company_name",
        "发行人": "issuer_name", "发行人名称": "issuer_name",
        "机构类型": "company_type",
        "备注": "additional_criteria",
    },
    "contacts": {
        "公司类型": "company_type", "机构类型": "company_type", "类型": "company_type",
        "公司名称": "company_name", "机构名称": "company_name",
        "发行人": "issuer_name", "发行人名称": "issuer_name",
        "基金名称": "fund_name", "基金": "fund_name",
        "姓名": "name", "名字": "name", "联系人": "name",
        "职位": "position", "岗位": "position",
        "邮箱": "email", "电子邮件": "email", "email": "email",
        "电话": "phone", "手机": "phone", "mobile": "phone", "mobil": "phone",
        "是否主联系人": "is_primary", "主联系人": "is_primary",
        "备注": "additional_info", "qt": "qt", "qq": "qq", "wechat": "wechat",
    },
}


def _normalize_df_columns(df, sheet_type):
    """把 DataFrame 的列名统一为英文（支持中文表头）。"""
    if df is None or df.empty:
        return df
    aliases = COLUMN_ALIASES.get(sheet_type, {})
    rename = {}
    for col in list(df.columns):
        c = str(col).strip()
        if c in aliases:
            rename[col] = aliases[c]
        elif sheet_type == "holdings" and "amount" not in rename.values():
            if "资产规模" in c or c.lower() == "amount" or (c and "amount" in c.lower()):
                rename[col] = "amount"
    if rename:
        df = df.rename(columns=rename)
    return df


def _safe_str(v, default=""):
    """将单元格值转为可写入数据库的字符串（NaN/None/空 转为 default）。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return default
    return s


def _has_required_columns(df, required):
    """检查 DataFrame 是否包含所需列名。"""
    if df is None:
        return False
    cols = set(str(c).strip() for c in df.columns)
    return all(r in cols for r in required)


def _has_holdings_columns(df):
    """持仓需：公司名称 + 债券代码（可从 bond_code/上证代码→bond_code_sse/深证代码→bond_code_sz/其他通用代码→bond_code_other 任选）"""
    if df is None or df.empty:
        return False
    cols = set(str(c).strip() for c in df.columns)
    if "company_name" not in cols:
        return False
    return any(c in cols for c in ("bond_code", "bond_code_sse", "bond_code_sz", "bond_code_other"))


def _get_bond_code_from_row(row):
    """从行中取债券代码（优先 bond_code，否则从上证/深证/其他通用代码取首个非空）"""
    for key in ("bond_code", "bond_code_sse", "bond_code_sz", "bond_code_other"):
        v = _safe_str(row.get(key))
        if v:
            return v
    return ""


def _extract_issuer_from_bond_name(bond_name):
    """从债券名称中尝试提取发行人（如含「股份有限公司」「有限公司」等）"""
    if not bond_name:
        return ""
    s = str(bond_name).strip()
    # 匹配：XXX股份有限公司、XXX有限公司、XXX有限责任公司
    for pat in [
        r"(.+?股份有限公司)",
        r"(.+?有限公司)(?:\s|$|[2-9]\d{3})",
        r"(.+?有限责任公司)(?:\s|$|[2-9]\d{3})",
    ]:
        m = re.search(pat, s)
        if m:
            return m.group(1).strip()
    return ""


def _get_issuer_from_row(row):
    """从行中取发行人，若无则尝试从持仓资产名称/债券名称提取"""
    v = _safe_str(row.get("issuer_name"))
    if v:
        return v
    return _extract_issuer_from_bond_name(_safe_str(row.get("bond_name"))) or "未知"


def _safe_float(v, default=0.0):
    """将单元格值转为浮点数（金额等）"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default


class DatabaseManager:
    """Database manager for Bond Buyer Match application"""

    def __init__(self, db_path='bond_buyer_match.db'):
        """Initialize database connection"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.db_path = os.path.join(script_dir, db_path)
        self.conn = None
        self.initialize_db()

    def get_connection(self):
        """Get database connection"""
        if self.conn is None:
            self.conn = sqlite3.connect(self.db_path)
            self.conn.execute('PRAGMA foreign_keys = ON')
            self.conn.row_factory = sqlite3.Row
        return self.conn

    def initialize_db(self):
        """Initialize database if it doesn't exist"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if not os.path.exists(self.db_path):
            script_dir = Path(__file__).parent
            schema_path = script_dir / 'schema.sql'
            with open(schema_path, 'r', encoding='utf-8') as f:
                schema_script = f.read()
            cursor.executescript(schema_script)
            conn.commit()
            print(f"Database initialized at {self.db_path}")
        self._migrate_import_batches(conn, cursor)

    def _migrate_import_batches(self, conn, cursor):
        """Add import_batches table and batch_id columns for tracking import source"""
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='import_batches'")
        if cursor.fetchone():
            return
        cursor.execute("""
            CREATE TABLE import_batches (
                batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                data_type TEXT NOT NULL,
                row_count INTEGER DEFAULT 0,
                imported_at TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        for tbl, col in [('holdings', 'batch_id'), ('can_buy_lists', 'batch_id'), ('contacts', 'batch_id')]:
            try:
                cursor.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} INTEGER REFERENCES import_batches(batch_id)")
            except sqlite3.OperationalError:
                pass
        conn.commit()

    def close_connection(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()
            self.conn = None

    def get_issuer_by_bond_code(self, bond_code):
        """Get issuer information for a bond code"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT i.* FROM issuers i
            JOIN bonds b ON i.issuer_id = b.issuer_id
            WHERE b.bond_code = ?
        """, (bond_code,))
        return cursor.fetchone()

    def add_bond_and_issuer(self, bond_code, bond_name, issuer_name, **kwargs):
        """Add bond and issuer if they don't exist"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (issuer_name,))
        issuer = cursor.fetchone()
        if issuer:
            issuer_id = issuer['issuer_id']
        else:
            issuer_data = {
                'issuer_name': issuer_name,
                'issuer_code': kwargs.get('issuer_code'),
                'credit_rating': kwargs.get('credit_rating'),
                'industry': kwargs.get('industry'),
                'additional_info': kwargs.get('additional_info')
            }
            columns = ', '.join(issuer_data.keys())
            placeholders = ', '.join(['?'] * len(issuer_data))
            cursor.execute(f"INSERT INTO issuers ({columns}) VALUES ({placeholders})", list(issuer_data.values()))
            issuer_id = cursor.lastrowid
        cursor.execute("SELECT bond_id FROM bonds WHERE bond_code = ?", (bond_code,))
        bond = cursor.fetchone()
        if not bond:
            bond_data = {
                'bond_code': bond_code,
                'bond_name': bond_name,
                'issuer_id': issuer_id,
                'issue_date': kwargs.get('issue_date'),
                'maturity_date': kwargs.get('maturity_date'),
                'coupon_rate': kwargs.get('coupon_rate')
            }
            columns = ', '.join(bond_data.keys())
            placeholders = ', '.join(['?'] * len(bond_data))
            cursor.execute(f"INSERT INTO bonds ({columns}) VALUES ({placeholders})", list(bond_data.values()))
        conn.commit()
        return issuer_id

    def _normalize_issuer_name_for_match(self, s):
        """去掉前导数字/空格，便于匹配 Excel 中「0西安高新控股有限公司」与「西安高新控股有限公司」"""
        if not s:
            return ""
        s = str(s).strip()
        return re.sub(r"^[\d\s]+", "", s).strip() or s

    def _issuer_keyword_from_bond_like_name(self, name):
        """若名称像债券简称（如 24西安高新SCP002），提取主体关键词「西安高新」用于匹配发行人"""
        if not name:
            return None
        # 匹配：数字 + 中文/英文 + SCP/CP/PPN/MTN 等
        m = re.search(r"[\d]*\s*([\u4e00-\u9fff\u3040-\u30ffa-zA-Z]+?)\s*(?:SCP|CP|PPN|MTN|PCPN)?\s*[\d]*$", str(name))
        if m:
            kw = m.group(1).strip()
            if len(kw) >= 2:
                return kw
        return None

    def get_issuer_ids_by_name_fuzzy(self, issuer_id):
        """根据发行人 id 查名称，再模糊匹配（Excel 前导0、iFinD 债券简称如 24西安高新SCP002 与 西安高新控股有限公司）"""
        if issuer_id is None:
            return []
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT issuer_name FROM issuers WHERE issuer_id = ?", (issuer_id,))
        row = cursor.fetchone()
        if not row or not row['issuer_name']:
            return [issuer_id]
        name = str(row['issuer_name']).strip()
        if not name:
            return [issuer_id]
        normalized = self._normalize_issuer_name_for_match(name)
        key = normalized.replace("股份有限公司", "").replace("有限公司", "").replace("有限责任公司", "").strip() or normalized
        keyword_from_bond = self._issuer_keyword_from_bond_like_name(name)
        # 条件：完全一致、包含关键词、或库中发行人去掉前导数字后与 normalized 一致（如 0西安高新控股有限公司）
        params = [name, normalized, f"%{key}%", f"%{normalized}%", normalized]
        sql = """
            SELECT issuer_id FROM issuers
            WHERE issuer_name = ?
               OR TRIM(LTRIM(issuer_name, '0123456789 ')) = ?
               OR issuer_name LIKE ?
               OR issuer_name LIKE ?
               OR ? LIKE '%' || issuer_name || '%'
        """
        if keyword_from_bond:
            params.extend([f"%{keyword_from_bond}%"])
            sql += " OR issuer_name LIKE ?"
        sql += " ORDER BY issuer_id"
        cursor.execute(sql, params)
        ids = [r['issuer_id'] for r in cursor.fetchall()]
        if not ids:
            ids = [issuer_id]
        return list(dict.fromkeys(ids))

    def get_can_buy_companies(self, issuer_id, bond_code=None):
        """Get companies that can buy from this issuer（含发行人名称模糊匹配；若有 bond_code 则也按债券代码查持仓，避免发行人缺失或名称不一致时查不到）"""
        conn = self.get_connection()
        cursor = conn.cursor()
        issuer_ids = self.get_issuer_ids_by_name_fuzzy(issuer_id) if issuer_id else []
        placeholders = ",".join("?" * len(issuer_ids)) if issuer_ids else ""
        explicit_can_buy = []
        holdings_can_buy = []
        if placeholders:
            cursor.execute(f"""
                SELECT DISTINCT bc.* FROM buyside_companies bc
                JOIN can_buy_lists cbl ON bc.company_id = cbl.company_id
                WHERE cbl.issuer_id IN ({placeholders})
            """, issuer_ids)
            explicit_can_buy = cursor.fetchall()
            cursor.execute(f"""
                SELECT DISTINCT bc.* FROM buyside_companies bc
                JOIN holdings h ON bc.company_id = h.company_id
                JOIN bonds b ON h.bond_id = b.bond_id
                WHERE b.issuer_id IN ({placeholders})
            """, issuer_ids)
            holdings_can_buy = cursor.fetchall()
        # 按债券代码查持仓（支持无发行人 / 发行人名称不一致的导入数据）
        if bond_code:
            cursor.execute("""
                SELECT DISTINCT bc.* FROM buyside_companies bc
                JOIN holdings h ON bc.company_id = h.company_id
                JOIN bonds b ON h.bond_id = b.bond_id
                WHERE b.bond_code = ?
            """, (bond_code,))
            by_bond_code = cursor.fetchall()
        combined_results = {}
        for company in explicit_can_buy:
            combined_results[company['company_id']] = {
                'company_id': company['company_id'],
                'company_name': company['company_name'],
                'company_type': company['company_type'],
                'match_type': 'explicit'
            }
        for company in holdings_can_buy:
            if company['company_id'] in combined_results:
                combined_results[company['company_id']]['match_type'] = 'both'
            else:
                combined_results[company['company_id']] = {
                    'company_id': company['company_id'],
                    'company_name': company['company_name'],
                    'company_type': company['company_type'],
                    'match_type': 'holdings'
                }
        if bond_code:
            for company in by_bond_code:
                if company['company_id'] in combined_results:
                    if combined_results[company['company_id']]['match_type'] == 'explicit':
                        combined_results[company['company_id']]['match_type'] = 'both'
                else:
                    combined_results[company['company_id']] = {
                        'company_id': company['company_id'],
                        'company_name': company['company_name'],
                        'company_type': company['company_type'],
                        'match_type': 'holdings'
                    }
        return list(combined_results.values())

    def get_company_contacts(self, company_id):
        """Get contacts for a company"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM contacts WHERE company_id = ? ORDER BY is_primary DESC, name", (company_id,))
        return cursor.fetchall()

    def get_fund_holdings(self, company_id, issuer_id):
        """Get fund holdings for a specific issuer"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT f.fund_id, f.fund_name, b.bond_id, b.bond_code, b.bond_name, h.amount
            FROM funds f
            JOIN holdings h ON f.fund_id = h.fund_id
            JOIN bonds b ON h.bond_id = b.bond_id
            WHERE f.company_id = ? AND b.issuer_id = ?
        """, (company_id, issuer_id))
        return cursor.fetchall()

    def get_fund_holdings_by_bond_code(self, company_id, bond_code):
        """按债券代码查询该机构下基金的持仓（与发行人名称无关，避免导入与搜索发行人名不一致时查不到）"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT f.fund_id, f.fund_name, b.bond_id, b.bond_code, b.bond_name, h.amount
            FROM funds f
            JOIN holdings h ON f.fund_id = h.fund_id
            JOIN bonds b ON h.bond_id = b.bond_id
            WHERE f.company_id = ? AND b.bond_code = ?
        """, (company_id, bond_code))
        return cursor.fetchall()

    def clear_data(self, data_type):
        """清空指定类型的数据：holdings / can_buy_lists / contacts / all。返回删除条数。"""
        conn = self.get_connection()
        cursor = conn.cursor()
        total = 0
        if data_type == "holdings":
            cursor.execute("DELETE FROM holdings")
            total = cursor.rowcount
        elif data_type == "can_buy_lists":
            cursor.execute("DELETE FROM can_buy_lists")
            total = cursor.rowcount
        elif data_type == "contacts":
            cursor.execute("DELETE FROM contacts")
            total = cursor.rowcount
        elif data_type == "all":
            for tbl in ("holdings", "can_buy_lists", "contacts", "funds", "bonds", "issuers", "buyside_companies", "import_batches"):
                try:
                    cursor.execute(f"DELETE FROM {tbl}")
                    total += cursor.rowcount
                except sqlite3.OperationalError:
                    pass
        else:
            return 0
        conn.commit()
        return total

    def get_import_batches(self):
        """返回所有导入批次，用于「删除指定导入」功能。每条: {batch_id, filename, data_type, row_count, imported_at}"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='import_batches'")
        if not cursor.fetchone():
            return []
        cursor.execute("""
            SELECT batch_id, filename, data_type, row_count, imported_at
            FROM import_batches ORDER BY batch_id DESC
        """)
        return [dict(r) for r in cursor.fetchall()]

    def delete_batch(self, batch_id):
        """按 batch_id 删除该批次导入的 holdings/can_buy_lists/contacts 记录，返回删除条数"""
        conn = self.get_connection()
        cursor = conn.cursor()
        total = 0
        for tbl in ("holdings", "can_buy_lists", "contacts"):
            cursor.execute(f"DELETE FROM {tbl} WHERE batch_id = ?", (batch_id,))
            total += cursor.rowcount
        cursor.execute("DELETE FROM import_batches WHERE batch_id = ?", (batch_id,))
        conn.commit()
        return total

    def _read_table_file(self, path, sheet_type="holdings"):
        """读取 Excel 或 CSV，返回 DataFrame。CSV 适合 50 万+ 行大批量。"""
        p = str(path)
        if p.lower().endswith(".csv"):
            try:
                df = pd.read_csv(p, encoding='utf-8-sig', dtype=str, on_bad_lines='skip')
            except UnicodeDecodeError:
                df = pd.read_csv(p, encoding='gbk', dtype=str, on_bad_lines='skip')
            df = df.dropna(how='all')
        else:
            df = pd.read_excel(p)
        return _normalize_df_columns(df, sheet_type)

    def _create_batch(self, filename, data_type, row_count):
        """Create import_batches record, return batch_id"""
        conn = self.get_connection()
        cursor = conn.cursor()
        name = os.path.basename(filename) if filename else ""
        cursor.execute(
            "INSERT INTO import_batches (filename, data_type, row_count) VALUES (?, ?, ?)",
            (name, data_type, row_count)
        )
        conn.commit()
        return cursor.lastrowid

    def import_from_excel(self, holdings_file, can_buy_file, contacts_file):
        """导入 Excel/CSV 到数据库。支持中文表头。CSV 适合大批量（50万+）。返回 (成功?, 持仓条数, 可买条数, 联系人条数)"""
        counts = {"holdings": 0, "can_buy_lists": 0, "contacts": 0}
        try:
            self.initialize_db()
            if holdings_file and os.path.exists(holdings_file):
                holdings_df = self._read_table_file(holdings_file, "holdings")
                if not holdings_df.empty and _has_holdings_columns(holdings_df):
                    n = len(holdings_df)
                    batch_id = self._create_batch(holdings_file, "holdings", n)
                    if n >= 10000:
                        self._import_holdings_bulk(holdings_df, batch_id)
                    else:
                        self._import_holdings(holdings_df, batch_id)
                    counts["holdings"] = n
                print(f"Imported {counts['holdings']} holdings records")
            if can_buy_file and os.path.exists(can_buy_file):
                can_buy_df = self._read_table_file(can_buy_file, "can_buy_lists")
                if not can_buy_df.empty and _has_required_columns(can_buy_df, ["company_name", "issuer_name"]):
                    n = len(can_buy_df)
                    batch_id = self._create_batch(can_buy_file, "can_buy_lists", n)
                    self._import_can_buy_lists(can_buy_df, batch_id)
                    counts["can_buy_lists"] = n
                print(f"Imported {counts['can_buy_lists']} can-buy list records")
            if contacts_file and os.path.exists(contacts_file):
                contacts_df = self._read_table_file(contacts_file, "contacts")
                if 'company_name' in contacts_df.columns:
                    contacts_df['company_name'] = contacts_df['company_name'].astype(str).str.strip()
                if not contacts_df.empty and _has_required_columns(contacts_df, ["company_name"]):
                    n = len(contacts_df)
                    batch_id = self._create_batch(contacts_file, "contacts", n)
                    self._import_contacts(contacts_df, batch_id)
                    counts["contacts"] = n
                print(f"Imported {counts['contacts']} contact records")
            return True, counts
        except Exception as e:
            print(f"Error importing data: {e}")
            traceback.print_exc()
            return False, counts

    def _import_holdings(self, df, batch_id=None):
        """Import holdings data from DataFrame"""
        conn = self.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            company_name = _safe_str(row.get("company_name"))
            if not company_name:
                continue
            bond_code = _get_bond_code_from_row(row)
            if not bond_code:
                continue
            issuer_name = _get_issuer_from_row(row)
            bond_name = _safe_str(row.get("bond_name", ""))
            cursor.execute("SELECT company_id FROM buyside_companies WHERE company_name = ?", (company_name,))
            company = cursor.fetchone()
            if company:
                company_id = company['company_id']
            else:
                cursor.execute("INSERT INTO buyside_companies (company_name, company_type) VALUES (?, ?)",
                               (company_name, row.get('company_type', 'Unknown')))
                company_id = cursor.lastrowid
            fund_id = None
            fund_name = _safe_str(row.get('fund_name'))
            if fund_name:
                cursor.execute("SELECT fund_id FROM funds WHERE fund_name = ? AND company_id = ?",
                               (fund_name, company_id))
                fund = cursor.fetchone()
                if fund:
                    fund_id = fund['fund_id']
                else:
                    cursor.execute("INSERT INTO funds (fund_name, company_id) VALUES (?, ?)", (fund_name, company_id))
                    fund_id = cursor.lastrowid
            cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (issuer_name,))
            issuer = cursor.fetchone()
            if issuer:
                issuer_id = issuer['issuer_id']
            else:
                cursor.execute("INSERT INTO issuers (issuer_name) VALUES (?)", (issuer_name,))
                issuer_id = cursor.lastrowid
            cursor.execute("SELECT bond_id FROM bonds WHERE bond_code = ?", (bond_code,))
            bond = cursor.fetchone()
            if bond:
                bond_id = bond['bond_id']
            else:
                cursor.execute("INSERT INTO bonds (bond_code, bond_name, issuer_id) VALUES (?, ?, ?)",
                               (bond_code, bond_name, issuer_id))
                bond_id = cursor.lastrowid
            amt = _safe_float(row.get('amount'), 0.0)
            cursor.execute(
                "INSERT INTO holdings (company_id, fund_id, bond_id, amount, batch_id) VALUES (?, ?, ?, ?, ?)",
                (company_id, fund_id, bond_id, amt, batch_id)
            )
        conn.commit()

    def _import_holdings_bulk(self, df, batch_id=None):
        """大批量导入持仓（50万+），使用批量插入和预建索引，显著提速"""
        conn = self.get_connection()
        cursor = conn.cursor()
        BATCH = 10000
        company_map = {}
        fund_map = {}
        issuer_map = {}
        bond_map = {}
        valid_rows = []
        for _, row in df.iterrows():
            company_name = _safe_str(row.get("company_name"))
            bond_code = _get_bond_code_from_row(row)
            if not company_name or not bond_code:
                continue
            issuer_name = _get_issuer_from_row(row)
            bond_name = _safe_str(row.get("bond_name", ""))
            fund_name = _safe_str(row.get("fund_name"))
            amt = _safe_float(row.get("amount"), 0.0)
            valid_rows.append((company_name, fund_name, issuer_name, bond_code, bond_name, amt))
        if not valid_rows:
            conn.commit()
            return
        for company_name, fund_name, issuer_name, bond_code, bond_name, amt in valid_rows:
            company_map.setdefault(company_name, None)
            issuer_map.setdefault(issuer_name, None)
        for name in company_map:
            cursor.execute("SELECT company_id FROM buyside_companies WHERE company_name = ?", (name,))
            r = cursor.fetchone()
            if r:
                company_map[name] = r['company_id']
            else:
                cursor.execute("INSERT INTO buyside_companies (company_name, company_type) VALUES (?, ?)",
                               (name, 'Unknown'))
                company_map[name] = cursor.lastrowid
        for name in issuer_map:
            cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (name,))
            r = cursor.fetchone()
            if r:
                issuer_map[name] = r['issuer_id']
            else:
                cursor.execute("INSERT INTO issuers (issuer_name) VALUES (?)", (name,))
                issuer_map[name] = cursor.lastrowid
        for company_name, fund_name, issuer_name, bond_code, bond_name, amt in valid_rows:
            issuer_id = issuer_map[issuer_name]
            if bond_code not in bond_map:
                cursor.execute("SELECT bond_id FROM bonds WHERE bond_code = ?", (bond_code,))
                r = cursor.fetchone()
                if r:
                    bond_map[bond_code] = r['bond_id']
                else:
                    cursor.execute("INSERT INTO bonds (bond_code, bond_name, issuer_id) VALUES (?, ?, ?)",
                                   (bond_code, bond_name, issuer_id))
                    bond_map[bond_code] = cursor.lastrowid
            key = (company_name, fund_name)
            if key not in fund_map:
                company_id = company_map[company_name]
                if fund_name:
                    cursor.execute("SELECT fund_id FROM funds WHERE fund_name = ? AND company_id = ?",
                                   (fund_name, company_id))
                    r = cursor.fetchone()
                    if r:
                        fund_map[key] = r['fund_id']
                    else:
                        cursor.execute("INSERT INTO funds (fund_name, company_id) VALUES (?, ?)",
                                       (fund_name, company_id))
                        fund_map[key] = cursor.lastrowid
                else:
                    fund_map[key] = None
        holdings_rows = []
        for company_name, fund_name, issuer_name, bond_code, bond_name, amt in valid_rows:
            company_id = company_map[company_name]
            fund_id = fund_map[(company_name, fund_name)]
            bond_id = bond_map[bond_code]
            holdings_rows.append((company_id, fund_id, bond_id, amt, batch_id))
        for i in range(0, len(holdings_rows), BATCH):
            chunk = holdings_rows[i:i + BATCH]
            cursor.executemany(
                "INSERT INTO holdings (company_id, fund_id, bond_id, amount, batch_id) VALUES (?, ?, ?, ?, ?)",
                chunk
            )
        conn.commit()

    def _import_can_buy_lists(self, df, batch_id=None):
        """Import can-buy lists from DataFrame"""
        conn = self.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            cursor.execute("SELECT company_id FROM buyside_companies WHERE company_name = ?", (row['company_name'],))
            company = cursor.fetchone()
            if company:
                company_id = company['company_id']
            else:
                cursor.execute("INSERT INTO buyside_companies (company_name, company_type) VALUES (?, ?)",
                               (row['company_name'], row.get('company_type', 'Unknown')))
                company_id = cursor.lastrowid
            cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (row['issuer_name'],))
            issuer = cursor.fetchone()
            if issuer:
                issuer_id = issuer['issuer_id']
            else:
                cursor.execute("INSERT INTO issuers (issuer_name) VALUES (?)", (row['issuer_name'],))
                issuer_id = cursor.lastrowid
            cursor.execute("SELECT list_id FROM can_buy_lists WHERE company_id = ? AND issuer_id = ?",
                           (company_id, issuer_id))
            if not cursor.fetchone():
                cursor.execute(
                    "INSERT INTO can_buy_lists (company_id, issuer_id, additional_criteria, batch_id) VALUES (?, ?, ?, ?)",
                    (company_id, issuer_id, row.get('additional_criteria', ''), batch_id)
                )
        conn.commit()

    def _import_contacts(self, df, batch_id=None):
        """Import contacts from DataFrame. Supports:
        - Full: 公司名称 + 姓名/联系人 + 职位/邮箱/电话等
        - Minimal: 仅 公司类型、公司名称、发行人（无姓名时用「—」，并同步写入可买名单）
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            company_name = _safe_str(row.get('company_name'))
            if not company_name:
                continue
            company_type = _safe_str(row.get('company_type'), 'Unknown') or 'Unknown'
            cursor.execute("SELECT company_id FROM buyside_companies WHERE company_name = ?", (company_name,))
            company = cursor.fetchone()
            if not company:
                cursor.execute("INSERT INTO buyside_companies (company_name, company_type) VALUES (?, ?)",
                               (company_name, company_type))
                company_id = cursor.lastrowid
            else:
                company_id = company['company_id']
            issuer_name = _safe_str(row.get('issuer_name'))
            if issuer_name:
                cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (issuer_name,))
                issuer = cursor.fetchone()
                if issuer:
                    issuer_id = issuer['issuer_id']
                else:
                    cursor.execute("INSERT INTO issuers (issuer_name) VALUES (?)", (issuer_name,))
                    issuer_id = cursor.lastrowid
                cursor.execute("SELECT list_id FROM can_buy_lists WHERE company_id = ? AND issuer_id = ?",
                               (company_id, issuer_id))
                if not cursor.fetchone():
                    cursor.execute(
                        "INSERT INTO can_buy_lists (company_id, issuer_id, additional_criteria, batch_id) VALUES (?, ?, ?, ?)",
                        (company_id, issuer_id, _safe_str(row.get('additional_criteria')), batch_id)
                    )
            fund_id = None
            fund_name = _safe_str(row.get('fund_name'))
            if fund_name:
                cursor.execute("SELECT fund_id FROM funds WHERE fund_name = ? AND company_id = ?", (fund_name, company_id))
                fund = cursor.fetchone()
                if fund:
                    fund_id = fund['fund_id']
                else:
                    cursor.execute("INSERT INTO funds (fund_name, company_id) VALUES (?, ?)", (fund_name, company_id))
                    fund_id = cursor.lastrowid
            is_primary = 0
            if 'is_primary' in row and pd.notna(row['is_primary']):
                v = row['is_primary']
                if isinstance(v, bool):
                    is_primary = 1 if v else 0
                elif isinstance(v, str):
                    is_primary = 1 if v == '是' else 0
                elif isinstance(v, (int, float)) and not pd.isna(v):
                    is_primary = 1 if v > 0 else 0
            name_val = _safe_str(row.get('name')) or "—"
            cursor.execute("""
                INSERT INTO contacts (company_id, fund_id, name, position, email, phone, qt, qq, wechat, mobil, is_primary, additional_info, batch_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (company_id, fund_id, name_val,
                  _safe_str(row.get('position')), _safe_str(row.get('email')), _safe_str(row.get('phone')),
                  _safe_str(row.get('qt')), _safe_str(row.get('qq')), _safe_str(row.get('wechat')), _safe_str(row.get('mobil')),
                  is_primary, _safe_str(row.get('additional_info')), batch_id))
        conn.commit()
