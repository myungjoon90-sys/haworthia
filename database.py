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

    # 🤔 의심후보: 아임웹/SMS에서 들어왔지만 거래명세표 구매자와
    #    100% 일치하지 않아 자동확인을 보류한 항목.
    #    사용자가 결제 및 입금완료 페이지에서 직접 승인/거절합니다.
    c.execute('''CREATE TABLE IF NOT EXISTS match_candidates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER,
        source TEXT,                  -- 'imweb' | 'sms'
        source_ref TEXT,              -- 아임웹 order_no 또는 sms_payments.id
        paid_name TEXT,               -- 아임웹/SMS에서 받은 이름 (실명/닉네임)
        paid_name2 TEXT,              -- 아임웹 괄호 안/밖 두번째 이름 (있다면)
        amount INTEGER,
        paid_at TEXT,
        candidate_order_id INTEGER,   -- 추정되는 orders.id (없으면 NULL)
        candidate_buyer_name TEXT,    -- 추정되는 거래명세표 구매자명
        candidate_amount INTEGER,     -- 거래명세표 합계금액
        confidence TEXT,              -- 'high' | 'medium' | 'low'
        reason TEXT,                  -- 추정 근거 (요약)
        status TEXT DEFAULT 'open',   -- 'open' | 'approved' | 'rejected'
        created_at TEXT,
        decided_at TEXT,
        UNIQUE(session_id, source, source_ref, candidate_order_id),
        FOREIGN KEY (session_id) REFERENCES live_sessions(id)
    )''')

    conn.commit()
    conn.close()
    print("✅ DB 초기화 완료")

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn
