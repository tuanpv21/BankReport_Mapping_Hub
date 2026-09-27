# -*- coding: utf-8 -*-
"""
Database abstraction layer for BankReport Mapping Hub.
Seamlessly supports:
 1. Local SQLite (mapping_hub.db) - Default for local development
 2. PostgreSQL (Render Free Database / Supabase) - Used when DATABASE_URL is set in environment
Automatically initializes schema and migrates seed data from SQLite to PostgreSQL on first launch.
"""

import os
import sys
import sqlite3

try:
    import psycopg2
    import psycopg2.extras
    PSYCOPG2_AVAILABLE = True
except ImportError:
    PSYCOPG2_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SQLITE_PATH = os.path.join(BASE_DIR, "mapping_hub.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

# Render sometimes gives postgres:// instead of postgresql://
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)


class DictLikeRow:
    """Wrapper that allows both dict access (row['name']) and index access (row[0]), plus dict(row)."""
    def __init__(self, data_dict, columns, values):
        self._dict = data_dict
        self._columns = columns
        self._values = values

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        return self._dict.get(key)

    def __iter__(self):
        return iter(self._dict)

    def keys(self):
        return self._dict.keys()

    def values(self):
        return self._dict.values()

    def items(self):
        return self._dict.items()

    def get(self, key, default=None):
        return self._dict.get(key, default)

    def __repr__(self):
        return repr(self._dict)


class PostgresCursorWrapper:
    def __init__(self, raw_cursor):
        self._cursor = raw_cursor
        self.lastrowid = None

    def execute(self, sql, params=None):
        # Translate '?' to '%s'
        clean_sql = sql.replace("?", "%s")
        
        # Check if INSERT into tables with auto-increment ID to capture lastrowid
        upper_sql = clean_sql.strip().upper()
        needs_returning = (
            upper_sql.startswith("INSERT INTO") and
            "RETURNING" not in upper_sql and
            any(t in upper_sql for t in ["MAPPINGS", "COMMENTS", "USERS", "CODE_LOOKUPS"])
        )
        if needs_returning:
            clean_sql = clean_sql.rstrip(";") + " RETURNING id;"

        if params:
            # Convert list to tuple if necessary
            params_tuple = tuple(params)
            self._cursor.execute(clean_sql, params_tuple)
        else:
            self._cursor.execute(clean_sql)

        if needs_returning:
            try:
                row = self._cursor.fetchone()
                if row:
                    self.lastrowid = row[0]
            except Exception:
                pass
        return self

    def executemany(self, sql, seq_of_params):
        clean_sql = sql.replace("?", "%s")
        self._cursor.executemany(clean_sql, seq_of_params)
        return self

    def fetchone(self):
        row = self._cursor.fetchone()
        if row is None:
            return None
        if hasattr(self._cursor, "description") and self._cursor.description:
            cols = [col.name.lower() for col in self._cursor.description]
            val_list = list(row)
            d = dict(zip(cols, val_list))
            return DictLikeRow(d, cols, val_list)
        return row

    def fetchall(self):
        rows = self._cursor.fetchall()
        if not rows:
            return []
        if hasattr(self._cursor, "description") and self._cursor.description:
            cols = [col.name.lower() for col in self._cursor.description]
            res = []
            for r in rows:
                val_list = list(r)
                d = dict(zip(cols, val_list))
                res.append(DictLikeRow(d, cols, val_list))
            return res
        return rows

    def close(self):
        self._cursor.close()


class PostgresConnectionWrapper:
    def __init__(self, raw_conn):
        self._conn = raw_conn

    def cursor(self):
        return PostgresCursorWrapper(self._conn.cursor())

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()


def get_db():
    """Returns database connection: PostgreSQL if DATABASE_URL is set, else SQLite."""
    if DATABASE_URL and PSYCOPG2_AVAILABLE:
        try:
            conn = psycopg2.connect(DATABASE_URL)
            return PostgresConnectionWrapper(conn)
        except Exception as e:
            print(f"[Database Error] Could not connect to PostgreSQL ({e}). Falling back to SQLite.")

    # SQLite fallback
    conn = sqlite3.connect(SQLITE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def is_postgres():
    return bool(DATABASE_URL and PSYCOPG2_AVAILABLE)


def init_schema():
    """Initializes tables for either PostgreSQL or SQLite."""
    if is_postgres():
        _init_postgres_schema()
    else:
        _init_sqlite_schema()


def _init_sqlite_schema():
    conn = sqlite3.connect(SQLITE_PATH)
    cur = conn.cursor()
    # Reports columns check
    r_cols = [r[1] for r in cur.execute("PRAGMA table_info(reports)").fetchall()] if cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='reports'").fetchone() else []
    for col, col_t in [("template_file", "VARCHAR(200)"), ("cycle", "VARCHAR(50)"), ("unit", "VARCHAR(100)"), ("has_matrix", "INTEGER DEFAULT 0")]:
        if col not in r_cols and r_cols:
            try:
                cur.execute(f"ALTER TABLE reports ADD COLUMN {col} {col_t}")
            except Exception:
                pass

    m_cols = [r[1] for r in cur.execute("PRAGMA table_info(mappings)").fetchall()] if cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mappings'").fetchone() else []
    for col, col_t in [("calc_method", "VARCHAR(30) DEFAULT 'CORE_TABLE'"), ("row_id", "VARCHAR(50)"), ("col_id", "VARCHAR(50)"), ("gl_account", "VARCHAR(100)"), ("balance_type", "VARCHAR(50)"), ("formula_expr", "TEXT")]:
        if col not in m_cols and m_cols:
            try:
                cur.execute(f"ALTER TABLE mappings ADD COLUMN {col} {col_t}")
            except Exception:
                pass

    # Users table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username VARCHAR(50) UNIQUE NOT NULL,
            password_hash VARCHAR(200) NOT NULL,
            full_name VARCHAR(100) NOT NULL,
            department VARCHAR(100),
            role VARCHAR(50) DEFAULT 'Chuyên viên',
            avatar_color VARCHAR(20) DEFAULT '#059669',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # Comments table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username VARCHAR(50) NOT NULL,
            full_name VARCHAR(100) NOT NULL,
            department VARCHAR(100),
            avatar_color VARCHAR(20) DEFAULT '#1e3a8a',
            target_type VARCHAR(20) NOT NULL,
            target_id VARCHAR(100) NOT NULL,
            report_code VARCHAR(50),
            field_code VARCHAR(100),
            content TEXT NOT NULL,
            tag VARCHAR(50) DEFAULT 'THAO_LUAN',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    conn.commit()
    conn.close()


def _init_postgres_schema():
    raw_conn = psycopg2.connect(DATABASE_URL)
    cur = raw_conn.cursor()

    # 1. systems
    cur.execute('''
        CREATE TABLE IF NOT EXISTS systems (
            code VARCHAR(20) PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            regulation VARCHAR(255),
            description TEXT
        );
    ''')

    # 2. reports
    cur.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            report_code VARCHAR(50) NOT NULL,
            system_code VARCHAR(20) NOT NULL,
            report_name VARCHAR(255) NOT NULL,
            frequency VARCHAR(50),
            target_table VARCHAR(100),
            procedure_name VARCHAR(100),
            regulation_ref VARCHAR(255),
            description TEXT,
            procedure_code TEXT,
            template_file VARCHAR(200),
            cycle VARCHAR(50),
            unit VARCHAR(100),
            has_matrix INTEGER DEFAULT 0,
            PRIMARY KEY (system_code, report_code)
        );
    ''')

    # 3. mappings
    cur.execute('''
        CREATE TABLE IF NOT EXISTS mappings (
            id SERIAL PRIMARY KEY,
            system_code VARCHAR(20) NOT NULL,
            report_code VARCHAR(50) NOT NULL,
            field_code VARCHAR(100) NOT NULL,
            field_name_vi TEXT,
            data_type VARCHAR(100),
            target_table VARCHAR(100),
            target_column VARCHAR(100),
            source_system VARCHAR(50),
            source_table VARCHAR(100),
            source_column VARCHAR(100),
            transformation_rule TEXT,
            lookup_ref VARCHAR(255),
            implemented_by VARCHAR(255),
            regulatory_ref VARCHAR(255),
            status VARCHAR(50) DEFAULT 'Active',
            notes TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            calc_method VARCHAR(30) DEFAULT 'CORE_TABLE',
            row_id VARCHAR(50),
            col_id VARCHAR(50),
            gl_account VARCHAR(100),
            balance_type VARCHAR(50),
            formula_expr TEXT
        );
    ''')

    # 4. code_lookups
    cur.execute('''
        CREATE TABLE IF NOT EXISTS code_lookups (
            id SERIAL PRIMARY KEY,
            system_code VARCHAR(20),
            category_name VARCHAR(100),
            bank_code VARCHAR(50),
            bank_desc TEXT,
            target_code VARCHAR(50),
            target_desc TEXT,
            source_table VARCHAR(100),
            source_column VARCHAR(100),
            notes TEXT
        );
    ''')

    # 5. users
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username VARCHAR(50) UNIQUE NOT NULL,
            password_hash VARCHAR(200) NOT NULL,
            full_name VARCHAR(100) NOT NULL,
            department VARCHAR(100),
            role VARCHAR(50) DEFAULT 'Chuyên viên',
            avatar_color VARCHAR(20) DEFAULT '#059669',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    # 6. comments
    cur.execute('''
        CREATE TABLE IF NOT EXISTS comments (
            id SERIAL PRIMARY KEY,
            user_id INTEGER,
            username VARCHAR(50) NOT NULL,
            full_name VARCHAR(100) NOT NULL,
            department VARCHAR(100),
            avatar_color VARCHAR(20) DEFAULT '#1e3a8a',
            target_type VARCHAR(20) NOT NULL,
            target_id VARCHAR(100) NOT NULL,
            report_code VARCHAR(50),
            field_code VARCHAR(100),
            content TEXT NOT NULL,
            tag VARCHAR(50) DEFAULT 'THAO_LUAN',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')

    raw_conn.commit()

    # Check if PostgreSQL is empty and needs data migration from SQLite
    cur.execute("SELECT COUNT(*) FROM mappings;")
    cnt = cur.fetchone()[0]
    if cnt == 0 and os.path.exists(SQLITE_PATH):
        print("[Startup] PostgreSQL is empty. Migrating all reports, mappings, templates, and users from SQLite snapshot...")
        _migrate_sqlite_to_postgres(raw_conn)

    raw_conn.close()


def _migrate_sqlite_to_postgres(pg_conn):
    """Copies all seed data from mapping_hub.db to PostgreSQL."""
    sq_conn = sqlite3.connect(SQLITE_PATH)
    sq_conn.row_factory = sqlite3.Row
    sq_cur = sq_conn.cursor()
    pg_cur = pg_conn.cursor()

    tables = ["systems", "reports", "mappings", "code_lookups", "users", "comments"]
    for tbl in tables:
        try:
            rows = sq_cur.execute(f"SELECT * FROM {tbl}").fetchall()
            if not rows:
                continue
            cols = [col[0] for col in sq_cur.description]
            col_str = ", ".join(cols)
            placeholders = ", ".join(["%s"] * len(cols))
            insert_sql = f"INSERT INTO {tbl} ({col_str}) VALUES ({placeholders}) ON CONFLICT DO NOTHING"
            
            data_tuples = [tuple(r) for r in rows]
            pg_cur.executemany(insert_sql, data_tuples)
            pg_conn.commit()
            print(f"[Migration] Copied {len(data_tuples)} records into table '{tbl}' in PostgreSQL.")
        except Exception as e:
            print(f"[Migration Warning] Error migrating table {tbl}: {e}")
            pg_conn.rollback()

    sq_conn.close()
