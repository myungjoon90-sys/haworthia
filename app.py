from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from apscheduler.schedulers.background import BackgroundScheduler
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO
from datetime import datetime, timedelta
import logging
import os
import socket

from database import init_db, get_conn
from imweb_api import get_paid_orders, extract_order_info
from sms_parser import parse_sms

import sys

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='static')
CORS(app)

# ══════════════════════════════════════════════════════════════════
#  메인 페이지 서빙
# ══════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    response = send_file('index.html')
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


# ══════════════════════════════════════════════════════════════════
#  SMS Webhook  ← MacroDroid가 여기로 POST 전송
# ══════════════════════════════════════════════════════════════════
@app.route('/sms', methods=['POST'])
def receive_sms():
    try:
        data = request.get_json(force=True) or {}
        body      = data.get('body', '')
        sender    = data.get('sender', '')
        recv_time = data.get('time', datetime.now().isoformat())

        parsed = parse_sms(body)

        try:
            conn = get_conn()
            conn.execute(
                '''INSERT INTO sms_payments (sender, body, parsed_name, parsed_amount, received_at)
                   VALUES (?, ?, ?, ?, ?)''',
                (sender, body,
                 parsed['name']   if parsed else None,
                 parsed['amount'] if parsed else None,
                 recv_time)
            )
            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"SMS DB 저장 오류: {e}")

        if parsed:
            try:
                match_sms_to_order(parsed, recv_time)
                logger.info(f"SMS 입금: {parsed['bank']} {parsed['name']} {parsed['amount']:,}원")
            except Exception as e:
                logger.error(f"SMS 매칭 오류: {e}")
        else:
            logger.info(f"SMS 수신 (파싱불가): {body[:40]}")

        return jsonify({'ok': True})

    except Exception as e:
        logger.error(f"SMS 수신 처리 오류: {e}")
        return jsonify({'ok': False, 'error': str(e)}), 200


# ══════════════════════════════════════════════════════════════════
#  거래명세표 업로드
# ══════════════════════════════════════════════════════════════════
@app.route('/api/session/upload', methods=['POST'])
def upload_session():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400

    file      = request.files['file']
    live_date = request.form.get('live_date', datetime.now().strftime('%Y-%m-%d'))

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        orders = parse_invoice_excel(ws)
    except Exception as e:
        return jsonify({'error': f'엑셀 파싱 오류: {e}'}), 400

    live_dt     = datetime.strptime(live_date, '%Y-%m-%d')
    check_start = (live_dt + timedelta(days=1)).strftime('%Y-%m-%d')
    check_end   = (live_dt + timedelta(days=7)).strftime('%Y-%m-%d')

    conn = get_conn()
    c = conn.execute(
        '''INSERT INTO live_sessions (filename, live_date, check_start, check_end, created_at)
           VALUES (?, ?, ?, ?, ?)''',
        (file.filename, live_date, check_start, check_end, datetime.now().isoformat())
    )
    session_id = c.lastrowid

    for o in orders:
        conn.execute(
            '''INSERT INTO orders (session_id, buyer_name, item, amount, pay_type, status)
               VALUES (?, ?, ?, ?, ?, 'pending')''',
            (session_id, o['name'], o['item'], o['amount'], o.get('pay_type', ''))
        )

    conn.commit()
    conn.close()

    logger.info(f"📋 세션 등록: {file.filename} ({len(orders)}명, {check_start}~{check_end})")

    # 업로드 즉시 한 번 확인 실행
    run_auto_check()

    return jsonify({
        'ok': True,
        'session_id': session_id,
        'orders_count': len(orders),
        'check_start': check_start,
        'check_end': check_end
    })


def parse_invoice_excel(ws):
    """거래명세표 엑셀에서 구매자/금액 파싱"""
    rows = list(ws.values)
    orders = []

    header_idx = name_col = amount_col = item_col = -1

    NAME_HINTS   = ['이름', '구매자', '닉네임', '성함']
    AMOUNT_HINTS = ['금액', '합계', '가격', '총액', '결제']
    ITEM_HINTS   = ['상품', '품목', '내역', '식물']

    for i, row in enumerate(rows[:15]):
        if not row:
            continue
        cells = [str(c or '').strip() for c in row]
        ni = next((j for j, c in enumerate(cells) if any(h in c for h in NAME_HINTS)), -1)
        ai = next((j for j, c in enumerate(cells) if any(h in c for h in AMOUNT_HINTS)), -1)
        if ni >= 0 and ai >= 0:
            header_idx = i
            name_col   = ni
            amount_col = ai
            item_col   = next((j for j, c in enumerate(cells) if any(h in c for h in ITEM_HINTS)), -1)
            break

    if header_idx < 0:
        return orders

    for row in rows[header_idx + 1:]:
        if not row or all((c is None or str(c).strip() == '') for c in row):
            continue
        row = list(row)

        name = str(row[name_col] or '').strip() if name_col < len(row) else ''
        try:
            raw  = str(row[amount_col] or '').replace(',', '').replace('원', '')
            amt  = int(float(raw))
        except Exception:
            continue
        item = str(row[item_col] or '').strip() if item_col >= 0 and item_col < len(row) else ''

        if name and amt > 0:
            orders.append({'name': name, 'item': item, 'amount': amt})

    return orders


# ══════════════════════════════════════════════════════════════════
#  세션 목록
# ══════════════════════════════════════════════════════════════════
@app.route('/api/sessions', methods=['GET'])
def get_sessions():
    conn = get_conn()
    sessions = conn.execute('SELECT * FROM live_sessions ORDER BY created_at DESC').fetchall()
    result = []
    for s in sessions:
        s = dict(s)
        orders = conn.execute('SELECT * FROM orders WHERE session_id=?', (s['id'],)).fetchall()
        s['orders']    = [dict(o) for o in orders]
        s['confirmed'] = sum(1 for o in s['orders'] if o['status'] == 'confirmed')
        s['pending']   = sum(1 for o in s['orders'] if o['status'] == 'pending')
        s['total']     = len(s['orders'])
        result.append(s)
    conn.close()
    return jsonify(result)


# ══════════════════════════════════════════════════════════════════
#  배송 목록
# ══════════════════════════════════════════════════════════════════
@app.route('/api/orders/pending', methods=['GET'])
def get_pending_orders():
    """미확인 대기 주문 - 구매자별 합산"""
    session_id = request.args.get('session_id')
    conn = get_conn()

    sql = '''SELECT o.buyer_name, 
                    GROUP_CONCAT(o.item, ' / ') as items,
                    SUM(o.amount) as total_amount,
                    COUNT(*) as item_count,
                    ls.live_date, ls.filename, o.session_id
             FROM orders o
             JOIN live_sessions ls ON o.session_id = ls.id
             WHERE o.status = 'pending' '''
    params = []
    if session_id:
        sql += ' AND o.session_id = ?'
        params.append(session_id)
    sql += ' GROUP BY o.session_id, o.buyer_name ORDER BY o.buyer_name'

    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/orders/confirmed', methods=['GET'])
def get_confirmed_orders():
    """입금확인 완료 주문 - 구매자별 합산"""
    session_id = request.args.get('session_id')
    conn = get_conn()

    sql = '''SELECT o.buyer_name,
                    GROUP_CONCAT(o.item, ' / ') as items,
                    SUM(o.amount) as total_amount,
                    COUNT(*) as item_count,
                    MAX(o.pay_type) as pay_type,
                    MAX(o.confirmed_at) as confirmed_at,
                    ls.live_date, ls.filename, o.session_id
             FROM orders o
             JOIN live_sessions ls ON o.session_id = ls.id
             WHERE o.status = 'confirmed' '''
    params = []
    if session_id:
        sql += ' AND o.session_id = ?'
        params.append(session_id)
    sql += ' GROUP BY o.session_id, o.buyer_name ORDER BY o.confirmed_at DESC'

    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])



    session_id = request.args.get('session_id')
    conn = get_conn()

    sql = '''SELECT o.id, o.buyer_name, o.item, o.amount, o.pay_type,
                    o.confirmed_at, o.bank_date, ls.live_date, ls.filename
             FROM orders o
             JOIN live_sessions ls ON o.session_id = ls.id
             WHERE o.status = 'confirmed' '''
    params = []
    if session_id:
        sql += ' AND o.session_id = ?'
        params.append(session_id)
    sql += ' ORDER BY o.confirmed_at DESC'

    items = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(i) for i in items])


@app.route('/api/orders/confirm-by-buyer', methods=['POST'])
def confirm_by_buyer():
    """구매자 이름으로 해당 세션의 모든 주문 수동 확인"""
    data = request.get_json(force=True) or {}
    session_id = data.get('session_id')
    buyer_name = data.get('buyer_name')
    if not session_id or not buyer_name:
        return jsonify({'error': '필수 파라미터 없음'}), 400
    conn = get_conn()
    conn.execute('''UPDATE orders SET status='confirmed', confirmed_at=?, pay_type='manual'
                   WHERE session_id=? AND buyer_name=? AND status='pending',
                 (datetime.now().isoformat(), session_id, buyer_name))
    conn.commit()
    conn.close()
    logger.info(f"수동확인: {buyer_name} (session {session_id})")
    return jsonify({'ok': True})



def manual_confirm_order(order_id):
    """수동으로 입금확인 처리"""
    conn = get_conn()
    conn.execute('''UPDATE orders SET status='confirmed', confirmed_at=?, pay_type='manual'
                   WHERE id=?''', (datetime.now().isoformat(), order_id))
    conn.commit()
    conn.close()
    logger.info(f"수동 입금확인: order_id={order_id}")
    return jsonify({'ok': True})


@app.route('/api/orders/<int:order_id>/unconfirm', methods=['POST'])
def manual_unconfirm_order(order_id):
    """입금확인 취소 (대기로 되돌리기)"""
    conn = get_conn()
    conn.execute('''UPDATE orders SET status='pending', confirmed_at=NULL, pay_type=NULL
                   WHERE id=?''', (order_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})



    session_id = request.args.get('session_id')
    conn = get_conn()
    sql = '''SELECT o.id, o.buyer_name, o.item, o.amount, o.pay_type,
                    o.confirmed_at, o.bank_date, ls.live_date, ls.filename
             FROM orders o
             JOIN live_sessions ls ON o.session_id = ls.id
             WHERE o.status = 'confirmed' '''
    params = []
    if session_id:
        sql += ' AND o.session_id = ?'
        params.append(session_id)
    sql += ' ORDER BY o.confirmed_at DESC'
    items = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(i) for i in items])


@app.route('/api/delivery/excel', methods=['GET'])
def download_delivery_excel():
    session_id = request.args.get('session_id')
    conn = get_conn()

    sql = '''SELECT o.buyer_name, o.item, o.amount, o.pay_type, o.confirmed_at, ls.live_date
             FROM orders o
             JOIN live_sessions ls ON o.session_id = ls.id
             WHERE o.status = 'confirmed' '''
    params = []
    if session_id:
        sql += ' AND o.session_id = ?'
        params.append(session_id)
    sql += ' ORDER BY o.confirmed_at DESC'

    items = conn.execute(sql, params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '배송목록'

    # 헤더
    headers = ['구매자명', '상품', '금액', '결제방법', '입금확인일시', '라이브날짜']
    header_fill = PatternFill(fill_type='solid', fgColor='1F6B2E')
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill   = header_fill
        cell.font   = Font(color='FFFFFF', bold=True)

    # 데이터
    pay_labels = {'card': '카드결제', 'transfer': '계좌이체', '': ''}
    for ri, item in enumerate(items, 2):
        row = list(item)
        row[3] = pay_labels.get(row[3], row[3])
        for ci, val in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=val)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    fname = f'배송목록_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(output, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ══════════════════════════════════════════════════════════════════
#  닉네임 매핑
# ══════════════════════════════════════════════════════════════════
@app.route('/api/mappings', methods=['GET'])
def get_mappings():
    conn = get_conn()
    rows = conn.execute('SELECT * FROM nick_mappings ORDER BY nickname').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/mappings', methods=['POST'])
def add_mapping():
    data = request.get_json(force=True) or {}
    conn = get_conn()
    conn.execute('INSERT OR REPLACE INTO nick_mappings (nickname, realname) VALUES (?, ?)',
                 (data['nickname'], data['realname']))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/api/mappings/<nickname>', methods=['DELETE'])
def delete_mapping(nickname):
    conn = get_conn()
    conn.execute('DELETE FROM nick_mappings WHERE nickname=?', (nickname,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ══════════════════════════════════════════════════════════════════
#  수동 확인 실행 버튼
# ══════════════════════════════════════════════════════════════════
@app.route('/api/check/run', methods=['POST'])
def manual_check():
    run_auto_check()
    return jsonify({'ok': True, 'message': '입금확인 완료'})



def get_check_logs():
    """세션별 7일 자동확인 실행 로그"""
    session_id = request.args.get('session_id')
    conn = get_conn()

    # 세션 정보 조회
    if session_id:
        session = conn.execute('SELECT * FROM live_sessions WHERE id=?', (session_id,)).fetchone()
    else:
        session = conn.execute('SELECT * FROM live_sessions ORDER BY created_at DESC LIMIT 1').fetchone()

    if not session:
        conn.close()
        return jsonify({'session': None, 'days': []})

    session = dict(session)

    # 7일 날짜 생성
    from datetime import date
    start = datetime.strptime(session['check_start'], '%Y-%m-%d').date()
    days = []
    for i in range(7):
        d = start + timedelta(days=i)
        log = conn.execute(
            'SELECT * FROM check_logs WHERE session_id=? AND check_date=? ORDER BY id DESC LIMIT 1',
            (session['id'], str(d))
        ).fetchone()
        today = date.today()
        if d < today:
            day_status = 'done' if log else 'missed'
        elif d == today:
            day_status = 'today'
        else:
            day_status = 'future'

        days.append({
            'date': str(d),
            'day_status': day_status,
            'log': dict(log) if log else None
        })

    conn.close()
    return jsonify({'session': session, 'days': days})



    run_auto_check()
    return jsonify({'ok': True, 'message': '입금확인 완료'})


# ══════════════════════════════════════════════════════════════════
#  SMS 수신 목록 (모니터링용)
# ══════════════════════════════════════════════════════════════════
@app.route('/api/sms/recent', methods=['GET'])
def recent_sms():
    conn = get_conn()
    rows = conn.execute('''SELECT * FROM sms_payments
                           ORDER BY received_at DESC LIMIT 30''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


# ══════════════════════════════════════════════════════════════════
#  상태 요약
# ══════════════════════════════════════════════════════════════════
@app.route('/api/myip', methods=['GET'])
def get_my_ip():
    """이 PC의 로컬 IP 자동 감지"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = "127.0.0.1"
    return jsonify({'ip': ip, 'sms_url': f'http://{ip}:5000/sms'})


@app.route('/api/status', methods=['GET'])
def get_status():
    conn = get_conn()
    total     = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    confirmed = conn.execute("SELECT COUNT(*) FROM orders WHERE status='confirmed'").fetchone()[0]
    pending   = conn.execute("SELECT COUNT(*) FROM orders WHERE status='pending'").fetchone()[0]
    sms_today = conn.execute(
        "SELECT COUNT(*) FROM sms_payments WHERE DATE(received_at)=DATE('now')"
    ).fetchone()[0]
    conn.close()
    return jsonify({'total': total, 'confirmed': confirmed, 'pending': pending, 'sms_today': sms_today})


# ══════════════════════════════════════════════════════════════════
#  핵심 매칭 로직
# ══════════════════════════════════════════════════════════════════
def resolve_names(conn, name):
    """닉네임 → 실명 매핑 포함한 검색 이름 목록"""
    names = [name]
    row = conn.execute('SELECT realname FROM nick_mappings WHERE nickname=?', (name,)).fetchone()
    if row:
        names.append(row['realname'])
    return names


def amount_matches(paid, ordered):
    """금액 일치 여부: 정확 일치 or 배송비 4000원 포함 or 100원 이하 오차"""
    diff = abs(paid - ordered)
    return diff == 0 or diff == 4000 or diff <= 100


def match_sms_to_order(parsed, recv_time):
    """SMS 입금 → 대기중 주문 매칭 (이름+금액 or 금액만)"""
    name   = parsed.get('name')
    amount = parsed['amount']
    today  = datetime.now().strftime('%Y-%m-%d')
    now    = datetime.now().isoformat()

    conn = get_conn()
    try:
        if name:
            # 이름 + 금액 매칭 (KB국민은행 / 이름있는 농협)
            names = resolve_names(conn, name)
            for sname in names:
                order = conn.execute('''
                    SELECT o.* FROM orders o
                    JOIN live_sessions ls ON o.session_id = ls.id
                    WHERE o.status = 'pending'
                      AND (o.buyer_name = ? OR o.buyer_name LIKE ?)
                      AND ? BETWEEN ls.check_start AND ls.check_end
                    LIMIT 1
                ''', (sname, f'%{sname}%', today)).fetchone()

                if order and amount_matches(amount, order['amount']):
                    conn.execute('''UPDATE orders
                                   SET status='confirmed', confirmed_at=?,
                                       pay_type='transfer', bank_date=?
                                   WHERE id=?''',
                                 (now, recv_time, order['id']))
                    conn.execute('''UPDATE sms_payments SET matched=1
                                   WHERE parsed_name=? AND parsed_amount=? AND matched=0''',
                                 (name, amount))
                    conn.commit()
                    logger.info(f"SMS 매칭: {order['buyer_name']} {order['amount']:,}원")
                    return
        else:
            # 이름 없음 (농협) → 금액만으로 매칭
            orders = conn.execute('''
                SELECT o.* FROM orders o
                JOIN live_sessions ls ON o.session_id = ls.id
                WHERE o.status = 'pending'
                  AND ? BETWEEN ls.check_start AND ls.check_end
            ''', (today,)).fetchall()

            matched = [dict(o) for o in orders if amount_matches(amount, dict(o)['amount'])]
            if len(matched) == 1:
                order = matched[0]
                conn.execute('''UPDATE orders
                               SET status='confirmed', confirmed_at=?,
                                   pay_type='transfer', bank_date=?
                               WHERE id=?''',
                             (now, recv_time, order['id']))
                conn.execute('''UPDATE sms_payments SET matched=1
                               WHERE parsed_amount=? AND parsed_name IS NULL AND matched=0''',
                             (amount,))
                conn.commit()
                logger.info(f"SMS 금액매칭(농협): {order['buyer_name']} {order['amount']:,}원")
            elif len(matched) > 1:
                logger.info(f"SMS 농협 금액중복 {len(matched)}명 - 수동확인필요: {amount:,}원")
    except Exception as e:
        logger.error(f"SMS 매칭 오류: {e}")
    finally:
        conn.close()


def run_auto_check():
    """매일 오전 11시 + 수동 실행: 아임웹 카드결제 + 미매칭 SMS 재대조"""
    logger.info("자동 입금확인 시작...")
    today    = datetime.now().strftime('%Y-%m-%d')
    today_ym = today.replace('-', '')
    now_time = datetime.now().strftime('%H:%M:%S')

    conn = get_conn()
    try:
        sessions = conn.execute('''
            SELECT * FROM live_sessions
            WHERE ? BETWEEN check_start AND check_end
        ''', (today,)).fetchall()

        imweb_confirmed = 0
        sms_confirmed = 0

        for session in sessions:
            session = dict(session)
            sid = session['id']
            logger.info(f"세션 확인: {session['filename']}")

            imweb_status = 'success'
            sms_status = 'success'
            imweb_confirmed = 0
            sms_confirmed = 0
            error_msg = None

            # ── 1. 아임웹 카드결제 확인 ──────────────────────────────
            try:
                imweb_orders = get_paid_orders(
                    session['check_start'].replace('-', ''), today_ym
                )
                for iorder in imweb_orders:
                    info = extract_order_info(iorder)
                    if not info['name'] or not info['amount']:
                        continue
                    names = resolve_names(conn, info['name'])
                    for name in names:
                        order = conn.execute('''
                            SELECT * FROM orders
                            WHERE session_id=? AND status='pending'
                              AND (buyer_name=? OR buyer_name LIKE ?)
                            LIMIT 1
                        ''', (sid, name, f'%{name}%')).fetchone()
                        if order and amount_matches(info['amount'], order['amount']):
                            conn.execute('''UPDATE orders
                                           SET status='confirmed', confirmed_at=?, pay_type='card'
                                           WHERE id=?''', (info['paid_at'], order['id']))
                            imweb_confirmed += 1
                            logger.info(f"카드결제 확인: {order['buyer_name']} {order['amount']:,}원")
                            break
            except Exception as e:
                imweb_status = 'error'
                error_msg = f"아임웹:{str(e)}"
                logger.error(f"아임웹 조회 오류: {e}")

            # ── 2. 미매칭 SMS 재대조 ─────────────────────────────────
            try:
                pending_orders = conn.execute(
                    "SELECT * FROM orders WHERE session_id=? AND status='pending'", (sid,)
                ).fetchall()
                unmatched_sms = conn.execute(
                    "SELECT * FROM sms_payments WHERE matched=0 AND parsed_name IS NOT NULL"
                ).fetchall()
                for order in pending_orders:
                    order = dict(order)
                    names = resolve_names(conn, order['buyer_name'])
                    for sms in unmatched_sms:
                        sms = dict(sms)
                        if sms['parsed_name'] in names or \
                           any(sms['parsed_name'] in n or n in sms['parsed_name'] for n in names):
                            if amount_matches(sms['parsed_amount'], order['amount']):
                                conn.execute('''UPDATE orders
                                               SET status='confirmed', confirmed_at=?,
                                                   pay_type='transfer', bank_date=?
                                               WHERE id=?''',
                                             (datetime.now().isoformat(), sms['received_at'], order['id']))
                                conn.execute('UPDATE sms_payments SET matched=1 WHERE id=?', (sms['id'],))
                                sms_confirmed += 1
                                logger.info(f"SMS 재매칭: {order['buyer_name']} {order['amount']:,}원")
                                break
            except Exception as e:
                sms_status = 'error'
                error_msg = (error_msg or '') + f" SMS:{str(e)}"
                logger.error(f"SMS 재대조 오류: {e}")

            # ── 로그 기록 ─────────────────────────────────────────────
            conn.execute('''INSERT INTO check_logs
                (session_id, check_date, check_time, imweb_status, sms_status,
                 imweb_confirmed, sms_confirmed, error_message)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (sid, today, now_time, imweb_status, sms_status,
                 imweb_confirmed, sms_confirmed, error_msg))

        conn.commit()
        logger.info(f"자동 입금확인 완료 (카드:{imweb_confirmed} SMS:{sms_confirmed})")

    except Exception as e:
        logger.error(f"자동확인 오류: {e}")
        conn.rollback()
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════════
#  서버 시작
# ══════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    init_db()

    scheduler = BackgroundScheduler(timezone='Asia/Seoul')
    scheduler.add_job(run_auto_check, 'cron', hour=11, minute=0,
                      id='daily_check', replace_existing=True)
    scheduler.start()
    logger.info("⏰ 스케줄러 시작 (매일 오전 11:00 자동 확인)")

    PORT = int(os.environ.get('PORT', 5000))
    logger.info(f"🌿 지양하월시아 서버 시작! 포트: {PORT}")
    app.run(host='0.0.0.0', port=PORT, debug=False, use_reloader=False)
