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
        item_no TEXT,
        amount INTEGER,
        pay_type TEXT,
        status TEXT DEFAULT 'pending',
        confirmed_at TEXT,
        bank_date TEXT,
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')
    # 기존 DB 마이그레이션 (item_no 컬럼 없으면 추가)
    try:
        c.execute("ALTER TABLE orders ADD COLUMN item_no TEXT")
    except sqlite3.OperationalError:
        pass

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
        realname TEXT,
        negative INTEGER DEFAULT 0
    )''')
    # 기존 DB에 negative 컬럼이 없으면 ALTER로 추가 (마이그레이션)
    try:
        c.execute("ALTER TABLE nick_mappings ADD COLUMN negative INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass  # 이미 있음

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

    # 🤔 의심후보: 아임웹/SMS에서 들어왔지만 거래명세표 구매자와
    #    100% 일치하지 않아 자동확인을 보류한 항목.
    c.execute('''CREATE TABLE IF NOT EXISTS match_candidates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER,
        source TEXT,
        source_ref TEXT,
        paid_name TEXT,
        paid_name2 TEXT,
        amount INTEGER,
        paid_at TEXT,
        candidate_order_id INTEGER,
        candidate_buyer_name TEXT,
        candidate_amount INTEGER,
        confidence TEXT,
        reason TEXT,
        status TEXT DEFAULT 'open',
        created_at TEXT,
        decided_at TEXT,
        UNIQUE(session_id, source, source_ref, candidate_order_id),
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')

    # 거래명세표 발송 진행상태 (식물보관 / 배송완료)
    c.execute('''CREATE TABLE IF NOT EXISTS buyer_status (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER,
        buyer_name TEXT,
        status TEXT,
        updated_at TEXT,
        UNIQUE(session_id, buyer_name),
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')

    conn.commit()
    conn.close()
    print("✅ DB 초기화 완료")

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn
