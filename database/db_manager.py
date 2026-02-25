import sqlite3
import os
import pandas as pd
from pathlib import Path
import traceback


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
        if not os.path.exists(self.db_path):
            conn = self.get_connection()
            cursor = conn.cursor()
            script_dir = Path(__file__).parent
            schema_path = script_dir / 'schema.sql'
            with open(schema_path, 'r', encoding='utf-8') as f:
                schema_script = f.read()
            cursor.executescript(schema_script)
            conn.commit()
            print(f"Database initialized at {self.db_path}")

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

    def get_can_buy_companies(self, issuer_id):
        """Get companies that can buy from this issuer"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT DISTINCT bc.* FROM buyside_companies bc
            JOIN can_buy_lists cbl ON bc.company_id = cbl.company_id
            WHERE cbl.issuer_id = ?
        """, (issuer_id,))
        explicit_can_buy = cursor.fetchall()
        cursor.execute("""
            SELECT DISTINCT bc.* FROM buyside_companies bc
            JOIN holdings h ON bc.company_id = h.company_id
            JOIN bonds b ON h.bond_id = b.bond_id
            WHERE b.issuer_id = ?
        """, (issuer_id,))
        holdings_can_buy = cursor.fetchall()
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

    def import_from_excel(self, holdings_file, can_buy_file, contacts_file):
        """Import data from Excel files into the database"""
        try:
            self.initialize_db()
            if holdings_file and os.path.exists(holdings_file):
                holdings_df = pd.read_excel(holdings_file)
                self._import_holdings(holdings_df)
                print(f"Imported {len(holdings_df)} holdings records")
            if can_buy_file and os.path.exists(can_buy_file):
                can_buy_df = pd.read_excel(can_buy_file)
                self._import_can_buy_lists(can_buy_df)
                print(f"Imported {len(can_buy_df)} can-buy list records")
            if contacts_file and os.path.exists(contacts_file):
                contacts_df = pd.read_excel(contacts_file)
                if 'company_name' in contacts_df.columns:
                    contacts_df['company_name'] = contacts_df['company_name'].astype(str).str.strip()
                self._import_contacts(contacts_df)
                print(f"Imported {len(contacts_df)} contact records")
            return True
        except Exception as e:
            print(f"Error importing data: {e}")
            traceback.print_exc()
            return False

    def _import_holdings(self, df):
        """Import holdings data from DataFrame"""
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
            fund_id = None
            if 'fund_name' in row and pd.notna(row['fund_name']):
                cursor.execute("SELECT fund_id FROM funds WHERE fund_name = ? AND company_id = ?",
                               (row['fund_name'], company_id))
                fund = cursor.fetchone()
                if fund:
                    fund_id = fund['fund_id']
                else:
                    cursor.execute("INSERT INTO funds (fund_name, company_id) VALUES (?, ?)", (row['fund_name'], company_id))
                    fund_id = cursor.lastrowid
            cursor.execute("SELECT issuer_id FROM issuers WHERE issuer_name = ?", (row['issuer_name'],))
            issuer = cursor.fetchone()
            if issuer:
                issuer_id = issuer['issuer_id']
            else:
                cursor.execute("INSERT INTO issuers (issuer_name) VALUES (?)", (row['issuer_name'],))
                issuer_id = cursor.lastrowid
            cursor.execute("SELECT bond_id FROM bonds WHERE bond_code = ?", (row['bond_code'],))
            bond = cursor.fetchone()
            if bond:
                bond_id = bond['bond_id']
            else:
                cursor.execute("INSERT INTO bonds (bond_code, bond_name, issuer_id) VALUES (?, ?, ?)",
                               (row['bond_code'], row.get('bond_name', ''), issuer_id))
                bond_id = cursor.lastrowid
            cursor.execute("INSERT INTO holdings (company_id, fund_id, bond_id, amount) VALUES (?, ?, ?, ?)",
                           (company_id, fund_id, bond_id, row.get('amount', 0)))
        conn.commit()

    def _import_can_buy_lists(self, df):
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
                cursor.execute("INSERT INTO can_buy_lists (company_id, issuer_id, additional_criteria) VALUES (?, ?, ?)",
                               (company_id, issuer_id, row.get('additional_criteria', '')))
        conn.commit()

    def _import_contacts(self, df):
        """Import contacts from DataFrame"""
        conn = self.get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            cursor.execute("SELECT company_id FROM buyside_companies WHERE company_name = ?", (row['company_name'],))
            company = cursor.fetchone()
            if not company:
                cursor.execute("INSERT INTO buyside_companies (company_name) VALUES (?)", (row['company_name'],))
                company_id = cursor.lastrowid
            else:
                company_id = company['company_id']
            fund_id = None
            if 'fund_name' in row and pd.notna(row['fund_name']):
                cursor.execute("SELECT fund_id FROM funds WHERE fund_name = ? AND company_id = ?",
                               (row['fund_name'], company_id))
                fund = cursor.fetchone()
                if fund:
                    fund_id = fund['fund_id']
                else:
                    cursor.execute("INSERT INTO funds (fund_name, company_id) VALUES (?, ?)", (row['fund_name'], company_id))
                    fund_id = cursor.lastrowid
            is_primary = 0
            if 'is_primary' in row and pd.notna(row['is_primary']):
                if isinstance(row['is_primary'], bool):
                    is_primary = 1 if row['is_primary'] else 0
                elif isinstance(row['is_primary'], str):
                    is_primary = 1 if row['is_primary'] == '是' else 0
                elif isinstance(row['is_primary'], (int, float)):
                    is_primary = 1 if row['is_primary'] > 0 else 0
            cursor.execute("""
                INSERT INTO contacts (company_id, fund_id, name, position, email, phone, qt, qq, wechat, mobil, is_primary, additional_info)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (company_id, fund_id, row['name'], row.get('position', ''), row.get('email', ''),
                 row.get('phone', ''), row.get('qt', ''), row.get('qq', ''), row.get('wechat', ''),
                 row.get('mobil', ''), is_primary, row.get('additional_info', '')))
        conn.commit()
