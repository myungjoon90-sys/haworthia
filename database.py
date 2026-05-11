import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'haworthia.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS live_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        live_date TEXT,
        check_start TEXT,
        check_end TEXT,
        created_at TEXT
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER,
        buyer_name TEXT,
        item TEXT,
        amount INTEGER,
        pay_type TEXT,
        status TEXT DEFAULT 'pending',
        confirmed_at TEXT,
        bank_date TEXT,
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS sms_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sender TEXT,
        body TEXT,
        parsed_name TEXT,
        parsed_amount INTEGER,
        received_at TEXT,
        matched INTEGER DEFAULT 0
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS nick_mappings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nickname TEXT UNIQUE,
        realname TEXT
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS check_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER,
        check_date TEXT,
        check_time TEXT,
        imweb_status TEXT DEFAULT 'pending',
        sms_status TEXT DEFAULT 'pending',
        imweb_confirmed INTEGER DEFAULT 0,
        sms_confirmed INTEGER DEFAULT 0,
        error_message TEXT,
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')

    conn.commit()
    conn.close()
    print("✅ DB 초기화 완료")

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn
