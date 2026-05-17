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
import re

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
#  ★ 요청 추적 미들웨어 (디버깅용)
#  모든 들어오는 요청을 무조건 기록 → Automate/외부에서 보낸 요청이
#  서버에 닿기만 하면 Railway 로그에 한 줄이 무조건 찍힘.
# ══════════════════════════════════════════════════════════════════
@app.before_request
def _log_every_request():
    try:
        # 정적 파일이나 favicon 등 시끄러운 요청은 제외
        if request.path in ('/favicon.ico',):
            return
        client_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        ctype = request.content_type or ''
        clen = request.content_length or 0
        # 짧은 본문이면 내용도 같이 (SMS POST 디버깅 핵심)
        body_preview = ''
        if request.method == 'POST' and clen < 2000:
            try:
                raw = request.get_data(cache=True, as_text=True) or ''
                body_preview = ' body=' + raw.replace('\n', '\\n')[:300]
            except Exception:
                body_preview = ' body=<read-failed>'
        logger.info(
            f"➡️  {request.method} {request.path} from={client_ip} "
            f"ct={ctype} len={clen}{body_preview}"
        )
    except Exception as e:
        logger.error(f"요청 로깅 오류: {e}")


# ══════════════════════════════════════════════════════════════════
#  헬스체크 (외부 모니터링/연결 테스트용)
# ══════════════════════════════════════════════════════════════════
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'ok': True, 'service': 'jiyang-haworthia', 'time': datetime.now().isoformat()})


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
#  SMS Webhook  ← Automate/MacroDroid가 여기로 POST 전송
#  여러 경로 모두 같은 핸들러로: /sms, /api/sms, /webhook/sms
# ══════════════════════════════════════════════════════════════════
@app.route('/sms',         methods=['POST'])
@app.route('/api/sms',     methods=['POST'])
@app.route('/webhook/sms', methods=['POST'])
def receive_sms():
    try:
        # JSON 우선, 실패 시 form/text도 시도 (Automate 설정 차이 대응)
        data = request.get_json(silent=True)
        if not data:
            # JSON 헤더 없이 form-encoded로 보내는 경우
            if request.form:
                data = request.form.to_dict()
            else:
                # raw text를 body로 취급
                raw = request.get_data(as_text=True) or ''
                if raw.strip():
                    data = {'body': raw}
                else:
                    data = {}

        body      = (data.get('body')   or data.get('message') or data.get('text')   or '').strip()
        sender    = (data.get('sender') or data.get('from')    or data.get('source') or '').strip()
        recv_time = data.get('time') or datetime.now().isoformat()

        if not body:
            logger.warning(f"⚠️  SMS 본문 없음 (data keys: {list(data.keys()) if data else 'none'})")
            return jsonify({'ok': False, 'error': 'body 필드가 비어있음'}), 400

        logger.info(f"📨 SMS 수신: sender='{sender}' body='{body[:80]}'")

        parsed = parse_sms(body)
        sms_id = None
        try:
            conn = get_conn()
            cur = conn.execute(
                '''INSERT INTO sms_payments (sender, body, parsed_name, parsed_amount, received_at)
                   VALUES (?, ?, ?, ?, ?)''',
                (sender, body,
                 parsed['name']   if parsed else None,
                 parsed['amount'] if parsed else None,
                 recv_time)
            )
            sms_id = cur.lastrowid
            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"SMS DB 저장 오류: {e}")

        if parsed:
            try:
                match_sms_to_order(parsed, recv_time, sms_id=sms_id)
                logger.info(f"✅ SMS 파싱 성공: {parsed['bank']} {parsed.get('name','?')} {parsed['amount']:,}원")
            except Exception as e:
                logger.error(f"SMS 매칭 오류: {e}")
        else:
            logger.info(f"ℹ️  SMS 파싱 불가: {body[:60]}")

        return jsonify({'ok': True, 'received': True, 'parsed': bool(parsed), 'sms_id': sms_id})

    except Exception as e:
        logger.exception(f"SMS 수신 처리 예외: {e}")
        return jsonify({'ok': False, 'error': str(e)}), 200


# ══════════════════════════════════════════════════════════════════
#  거래명세표 업로드
# ══════════════════════════════════════════════════════════════════
def extract_live_date_from_filename(filename):
    """파일명에서 라이브방송 날짜(YYYY-MM-DD) 추출"""
    if not filename: return None
    m = re.search(r'(20\d{2})-(\d{2})-(\d{2})', filename)
    if m: return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'(20\d{2})(\d{2})(\d{2})', filename)
    if m:
        try:
            datetime.strptime(f"{m.group(1)}-{m.group(2)}-{m.group(3)}", '%Y-%m-%d')
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        except ValueError:
            return None
    return None


@app.route('/api/session/upload', methods=['POST'])
def upload_session():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400

    file = request.files['file']

    # live_date 결정: 1) 폼 입력값  2) 파일명 추출  3) 오늘 (최후)
    live_date_form = (request.form.get('live_date') or '').strip()
    if live_date_form:
        live_date = live_date_form
        live_date_source = 'form'
    else:
        extracted = extract_live_date_from_filename(file.filename)
        if extracted:
            live_date = extracted
            live_date_source = 'filename'
        else:
            live_date = datetime.now().strftime('%Y-%m-%d')
            live_date_source = 'today_fallback'

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        orders = parse_invoice_excel(ws)
    except Exception as e:
        return jsonify({'error': f'엑셀 파싱 오류: {e}'}), 400

    try:
        live_dt = datetime.strptime(live_date, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': f'라이브 날짜 형식 오류: {live_date}'}), 400

    check_start = live_dt.strftime('%Y-%m-%d')
    check_end   = (live_dt + timedelta(days=7)).strftime('%Y-%m-%d')
    logger.info(f"📅 live_date 결정: {live_date} (source={live_date_source})")

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

        # 총합계/소계/합계 같은 요약행 스킵 (구매자가 아님)
        clean = name.replace(' ', '').replace('\u00a0', '')
        if clean in {'총합계', '총계', '소계', '합계', '총합', '계', '총주문', '주문합계'}:
            continue
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


@app.route('/api/orders/confirm-by-buyer', methods=['POST'])
def confirm_by_buyer():
    """구매자 이름으로 해당 세션의 모든 주문 수동 확인"""
    data = request.get_json(force=True) or {}
    session_id = data.get('session_id')
    buyer_name = data.get('buyer_name')
    if not session_id or not buyer_name:
        return jsonify({'error': '필수 파라미터 없음'}), 400
    conn = get_conn()
    conn.execute(
        "UPDATE orders SET status='confirmed', confirmed_at=?, pay_type='manual' WHERE session_id=? AND buyer_name=? AND status='pending'",
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


# ══════════════════════════════════════════════════════════════════
#  7일 자동확인 진행현황 조회  ⭐ (이전 버전에서 데코레이터 누락이 있었음)
# ══════════════════════════════════════════════════════════════════
@app.route('/api/check/logs', methods=['GET'])
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
    sms_today = conn.execute("SELECT COUNT(*) FROM sms_payments WHERE DATE(received_at)=DATE('now')").fetchone()[0]
    candidates_open = conn.execute("SELECT COUNT(*) FROM match_candidates WHERE status='open'").fetchone()[0]
    conn.close()
    return jsonify({'total': total, 'confirmed': confirmed, 'pending': pending,
                    'sms_today': sms_today, 'candidates_open': candidates_open})


# ══════════════════════════════════════════════════════════════════
#  의심후보 조회 / 승인 / 거절
# ══════════════════════════════════════════════════════════════════
@app.route('/api/candidates', methods=['GET'])
def list_candidates():
    session_id = request.args.get('session_id')
    status     = request.args.get('status', 'open')
    conn = get_conn()
    sql = """SELECT mc.*, ls.live_date, ls.filename
             FROM match_candidates mc
             LEFT JOIN live_sessions ls ON mc.session_id=ls.id
             WHERE mc.status=?"""
    params = [status]
    if session_id:
        sql += ' AND mc.session_id=?'
        params.append(session_id)
    sql += """ ORDER BY CASE mc.confidence WHEN 'high' THEN 0 WHEN 'medium' THEN 1
                          WHEN 'low' THEN 2 ELSE 3 END, mc.created_at DESC"""
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/candidates/<int:cand_id>/approve', methods=['POST'])
def approve_candidate(cand_id):
    """승인 — 같은 buyer_name + session의 모든 pending 행을 일괄 confirmed."""
    data = request.get_json(silent=True) or {}
    override_buyer = (data.get('buyer_name') or '').strip()
    conn = get_conn()
    try:
        cand = conn.execute('SELECT * FROM match_candidates WHERE id=?', (cand_id,)).fetchone()
        if not cand: return jsonify({'error': '후보 없음'}), 404
        cand = dict(cand)
        target_buyer = override_buyer or cand.get('candidate_buyer_name')
        target_session = cand.get('session_id')
        if not target_buyer:
            return jsonify({'error': '구매자 이름 필요'}), 400
        pending = conn.execute(
            "SELECT id, amount FROM orders WHERE session_id=? AND buyer_name=? AND status='pending'",
            (target_session, target_buyer)).fetchall()
        if not pending:
            return jsonify({'error': f"세션{target_session} '{target_buyer}' pending 없음"}), 404
        pay_type = 'card' if cand['source'] == 'imweb' else 'transfer'
        now = datetime.now().isoformat()
        for row in pending:
            conn.execute("UPDATE orders SET status='confirmed', confirmed_at=?, pay_type=?, bank_date=? WHERE id=?",
                         (now, pay_type, cand.get('paid_at'), row['id']))
        conn.execute("UPDATE match_candidates SET status='approved', decided_at=?, candidate_buyer_name=? WHERE id=?",
                     (now, target_buyer, cand_id))
        if cand['source'] == 'sms' and cand.get('source_ref'):
            try: conn.execute('UPDATE sms_payments SET matched=1 WHERE id=?', (int(cand['source_ref']),))
            except: pass
        # ⭐ 닉네임 매핑 자동 추가 (동일인) — 받은 이름과 구매자명이 다를 때만 의미있음
        paid_name = cand.get('paid_name')
        if paid_name and target_buyer and paid_name != target_buyer:
            try:
                conn.execute(
                    "INSERT OR REPLACE INTO nick_mappings (nickname, realname, negative) VALUES (?, ?, 0)",
                    (paid_name, target_buyer)
                )
                logger.info(f"  🔗 매핑 자동 추가 (동일인): '{paid_name}' ↔ '{target_buyer}'")
            except Exception as e:
                logger.warning(f"매핑 자동추가 실패: {e}")
        conn.commit()
        total = sum(r['amount'] for r in pending)
        logger.info(f"  👍 후보 승인: cand={cand_id} → '{target_buyer}' ({len(pending)}건 합{total:,}원)")
        return jsonify({'ok': True, 'buyer_name': target_buyer, 'rows_confirmed': len(pending), 'total_amount': total})
    finally:
        conn.close()


@app.route('/api/candidates/<int:cand_id>/reject', methods=['POST'])
def reject_candidate(cand_id):
    """거절 — 후보 닫기 + 닉네임 매핑에 '동일인 아님(negative=1)' 자동 추가"""
    conn = get_conn()
    try:
        cand = conn.execute('SELECT * FROM match_candidates WHERE id=?', (cand_id,)).fetchone()
        if not cand:
            return jsonify({'error': '후보 없음'}), 404
        cand = dict(cand)
        conn.execute("UPDATE match_candidates SET status='rejected', decided_at=? WHERE id=?",
                     (datetime.now().isoformat(), cand_id))
        # ⭐ 닉네임 매핑 자동 추가 (동일인 아님) — 같은 pair 다시 후보로 만들지 않게
        paid_name  = cand.get('paid_name')
        buyer_name = cand.get('candidate_buyer_name')
        if paid_name and buyer_name and paid_name != buyer_name:
            try:
                conn.execute(
                    "INSERT OR REPLACE INTO nick_mappings (nickname, realname, negative) VALUES (?, ?, 1)",
                    (paid_name, buyer_name)
                )
                logger.info(f"  🚫 매핑 자동 추가 (동일인 아님): '{paid_name}' ↔ '{buyer_name}'")
            except Exception as e:
                logger.warning(f"매핑 자동추가 실패: {e}")
        conn.commit()
        return jsonify({'ok': True})
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════════
#  ★ SMS 진단 엔드포인트 (Automate 디버깅)
# ══════════════════════════════════════════════════════════════════
@app.route('/api/sms/echo', methods=['GET', 'POST'])
def sms_echo():
    info = {
        'ok': True, 'received_at': datetime.now().isoformat(),
        'method': request.method, 'path': request.path,
        'remote_addr': request.headers.get('X-Forwarded-For', request.remote_addr),
        'content_type': request.content_type,
        'content_length': request.content_length,
        'headers': {k: v for k, v in request.headers.items()},
        'args': request.args.to_dict(flat=False),
        'form': request.form.to_dict(flat=False),
        'json': request.get_json(silent=True),
        'body_preview': (request.get_data(as_text=True) or '')[:500],
    }
    logger.info(f"🩺 /api/sms/echo {request.method}")
    return jsonify(info)


@app.route('/api/sms/simulate', methods=['GET', 'POST'])
def sms_simulate():
    body = (request.args.get('body') or request.form.get('body')
            or (request.get_json(silent=True) or {}).get('body') or '').strip()
    if not body:
        return jsonify({'ok': False, 'error': 'body 필요'}), 400
    parsed = parse_sms(body)
    recv_iso = datetime.now().isoformat()
    sms_id = None
    try:
        conn = get_conn()
        cur = conn.execute("INSERT INTO sms_payments (sender, body, parsed_name, parsed_amount, received_at) VALUES (?,?,?,?,?)",
                            ('SIMULATE', body,
                             parsed['name']   if parsed else None,
                             parsed['amount'] if parsed else None,
                             recv_iso))
        sms_id = cur.lastrowid
        conn.commit()
        conn.close()
    except Exception as e:
        logger.error(f"simulate 저장: {e}")
    if parsed:
        match_sms_to_order(parsed, recv_iso, sms_id=sms_id)
    return jsonify({'ok': True, 'parsed': parsed, 'sms_id': sms_id})



# ══════════════════════════════════════════════════════════════════
#  은행 입금내역 엑셀 업로드 (수동 일괄 매칭)
# ══════════════════════════════════════════════════════════════════
def parse_bank_excel(file_storage):
    """KB/농협 등 은행 입금내역 엑셀(.xls/.xlsx)에서 (이름, 입금액, 일시) 추출"""
    fn = (file_storage.filename or '').lower()
    rows = []
    if fn.endswith('.xls') and not fn.endswith('.xlsx'):
        try:
            import xlrd
        except ImportError:
            raise RuntimeError("xls 파일을 읽으려면 xlrd 라이브러리가 필요합니다 (requirements.txt에 xlrd 추가)")
        data = file_storage.read()
        wb = xlrd.open_workbook(file_contents=data)
        sheet = wb.sheet_by_index(0)
        for r in range(sheet.nrows):
            rows.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
    else:
        wb = openpyxl.load_workbook(file_storage, data_only=True)
        ws = wb.active
        for r in ws.values:
            rows.append(list(r))

    header_idx = name_col = amount_col = date_col = -1
    for i, row in enumerate(rows[:20]):
        if not row: continue
        cells = [str(c or '').strip() for c in row]
        ni = next((j for j, c in enumerate(cells) if '보낸' in c or '받는' in c), -1)
        ai = next((j for j, c in enumerate(cells) if c == '입금액(원)' or c == '입금액' or '입금' in c), -1)
        di = next((j for j, c in enumerate(cells) if '거래일' in c or '날짜' in c), -1)
        if ni >= 0 and ai >= 0:
            header_idx = i
            name_col = ni
            amount_col = ai
            date_col = di
            break
    if header_idx < 0:
        return []

    deposits = []
    for row in rows[header_idx + 1:]:
        if not row: continue
        name_v = row[name_col] if name_col < len(row) else ''
        name = str(name_v or '').strip()
        if not name: continue
        try:
            raw = str(row[amount_col] if amount_col < len(row) else '').replace(',', '').replace('원', '').strip()
            if not raw or raw == '0': continue
            amt = int(float(raw))
        except (ValueError, TypeError):
            continue
        if amt <= 0: continue
        date_str = ''
        if date_col >= 0 and date_col < len(row):
            date_str = str(row[date_col] or '').strip()
        deposits.append({'name': name, 'amount': amt, 'date': date_str})
    return deposits


@app.route('/api/bank/upload', methods=['POST'])
def upload_bank():
    """은행 입금내역 엑셀 업로드 → 각 행을 SMS와 동일하게 매칭 (자동확인 or 의심후보)"""
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400
    file = request.files['file']
    try:
        deposits = parse_bank_excel(file)
    except Exception as e:
        return jsonify({'error': f'엑셀 파싱 오류: {e}'}), 400

    if not deposits:
        return jsonify({'error': '입금 내역을 찾을 수 없습니다 (헤더에 "보낸분/받는분" + "입금액"이 있어야 함)'}), 400

    matched_high = matched_cand = no_match = 0
    for d in deposits:
        try:
            recv_iso = datetime.now().isoformat()
            conn = get_conn()
            cur = conn.execute(
                "INSERT INTO sms_payments (sender, body, parsed_name, parsed_amount, received_at) VALUES (?,?,?,?,?)",
                ('BANK_EXCEL', f"엑셀:{d['name']} {d['amount']:,}원 {d.get('date','')}",
                 d['name'], d['amount'], recv_iso)
            )
            sms_id = cur.lastrowid
            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"은행 입금 저장 오류: {e}")
            no_match += 1
            continue
        try:
            parsed = {'name': d['name'], 'amount': d['amount'], 'bank': '엑셀'}
            match_sms_to_order(parsed, recv_iso, sms_id=sms_id)
            conn = get_conn()
            row = conn.execute("SELECT matched FROM sms_payments WHERE id=?", (sms_id,)).fetchone()
            matched = (row['matched'] == 1) if row else False
            if matched:
                matched_high += 1
            else:
                cand = conn.execute(
                    "SELECT 1 FROM match_candidates WHERE source='sms' AND source_ref=? AND status='open'",
                    (str(sms_id),)
                ).fetchone()
                if cand: matched_cand += 1
                else: no_match += 1
            conn.close()
        except Exception as e:
            logger.error(f"은행 입금 매칭 오류: {e}")
            no_match += 1

    logger.info(f"📥 은행입금 업로드: 총 {len(deposits)}건 → 자동확인 {matched_high}, 의심후보 {matched_cand}, 매칭실패 {no_match}")
    return jsonify({
        'ok': True,
        'total': len(deposits),
        'matched_high': matched_high,
        'matched_candidate': matched_cand,
        'no_match': no_match
    })


# ══════════════════════════════════════════════════════════════════
#  닉네임 매핑 CSV Export / Import
# ══════════════════════════════════════════════════════════════════
@app.route('/api/mappings/export', methods=['GET'])
def export_mappings():
    """닉네임 매핑 전체를 텍스트 파일로 다운로드.
    양식:  nickname=realname     (동일인)
           nickname≠realname     (동일인 아님)
    """
    from flask import Response
    conn = get_conn()
    rows = conn.execute(
        'SELECT nickname, realname, COALESCE(negative,0) as negative FROM nick_mappings ORDER BY negative, nickname'
    ).fetchall()
    conn.close()
    lines = []
    for r in rows:
        sep = '≠' if r['negative'] == 1 else '='
        lines.append(f"{r['nickname']}{sep}{r['realname']}")
    text = '\n'.join(lines) + '\n'
    fname = f'nick_mappings_{datetime.now().strftime("%Y%m%d_%H%M")}.txt'
    return Response(text,
                    mimetype='text/plain; charset=utf-8',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


@app.route('/api/mappings/import', methods=['POST'])
def import_mappings():
    """텍스트 양식 업로드로 닉네임 매핑 일괄 추가.
    한 줄당 한 매핑:
        nickname=realname     → 동일인 (negative=0)
        nickname≠realname  → 동일인 아님 (negative=1)
        nickname!=realname    → 동일인 아님 (대체 표기)
    빈 줄, # 으로 시작하는 줄은 무시.
    같은 nickname이 이미 있으면 덮어쓰기.
    """
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400
    file = request.files['file']
    try:
        raw = file.read()
        text = None
        for enc in ('utf-8-sig', 'utf-8', 'cp949', 'euc-kr'):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            return jsonify({'error': '파일 인코딩 해석 실패'}), 400
    except Exception as e:
        return jsonify({'error': f'파일 읽기 오류: {e}'}), 400

    added = updated = skipped = 0
    conn = get_conn()
    try:
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                skipped += 1
                continue
            sep = None
            neg = 0
            if '≠' in line:
                sep, neg = '≠', 1
            elif '!=' in line:
                sep, neg = '!=', 1
            elif '=' in line:
                sep, neg = '=', 0
            if not sep:
                skipped += 1
                continue
            parts = line.split(sep, 1)
            if len(parts) != 2:
                skipped += 1
                continue
            nick = parts[0].strip()
            real = parts[1].strip()
            if not nick or not real:
                skipped += 1
                continue
            try:
                existing = conn.execute('SELECT id FROM nick_mappings WHERE nickname=?', (nick,)).fetchone()
                conn.execute(
                    "INSERT OR REPLACE INTO nick_mappings (nickname, realname, negative) VALUES (?,?,?)",
                    (nick, real, neg)
                )
                if existing: updated += 1
                else: added += 1
            except Exception as e:
                skipped += 1
                logger.warning(f"매핑 import 행 스킵: {e}")
        conn.commit()
    finally:
        conn.close()
    logger.info(f"📤 매핑 import (텍스트): 추가 {added}, 수정 {updated}, 스킵 {skipped}")
    return jsonify({'ok': True, 'added': added, 'updated': updated, 'skipped': skipped})


# ══════════════════════════════════════════════════════════════════
#  핵심 매칭 로직 (구매자별 SUM 합산 매칭 + 의심후보 시스템)
# ══════════════════════════════════════════════════════════════════
def _norm(s):
    """이름 정규화: 공백/특수문자 제거, 소문자화"""
    if not s: return ''
    return re.sub(r'[\s\(\)\[\]\.\,\-_]+', '', str(s)).lower()


def resolve_names(conn, name):
    """닉네임 → 실명 매핑 (양방향, negative=1 mapping은 제외)"""
    names = [name]
    if not name: return names
    row = conn.execute(
        "SELECT realname FROM nick_mappings WHERE nickname=? AND COALESCE(negative,0)=0",
        (name,)
    ).fetchone()
    if row and row['realname']: names.append(row['realname'])
    row = conn.execute(
        "SELECT nickname FROM nick_mappings WHERE realname=? AND COALESCE(negative,0)=0",
        (name,)
    ).fetchone()
    if row and row['nickname']: names.append(row['nickname'])
    return list(dict.fromkeys(filter(None, names)))


def amount_matches(paid, ordered):
    diff = abs(paid - ordered)
    return diff == 0 or diff == 4000 or diff <= 100


def find_buyer_match(conn, session_id, search_names, amount):
    """
    퍼지 매칭 - 구매자별 SUM(amount)를 기준으로.
    Returns: ({'buyer_name','total_amount','order_ids'}, confidence, reason)
    """
    if not search_names: return (None, None, '검색이름 없음')
    grouped_rows = conn.execute('''
        SELECT buyer_name, SUM(amount) AS total_amount, COUNT(*) AS row_count,
               GROUP_CONCAT(id) AS order_ids
        FROM orders WHERE session_id=? AND status='pending' GROUP BY buyer_name
    ''', (session_id,)).fetchall()
    if not grouped_rows: return (None, None, '대기 주문 없음')
    grouped = [{
        'buyer_name': r['buyer_name'],
        'total_amount': int(r['total_amount'] or 0),
        'order_ids': [int(x) for x in (r['order_ids'] or '').split(',') if x],
    } for r in grouped_rows]

    # 1) 이름 정확일치 + 합산금액 일치 → high
    for name in search_names:
        for g in grouped:
            if g['buyer_name'] == name and amount_matches(amount, g['total_amount']):
                return (g, 'high', f"이름정확({name})+합산일치")

    # 2) 정규화/부분 + 합산금액 일치 → medium
    norm_searches = {n: _norm(n) for n in search_names if n}
    for name, nname in norm_searches.items():
        if not nname: continue
        for g in grouped:
            buyer = g['buyer_name'] or ''; nbuyer = _norm(buyer)
            if not nbuyer: continue
            if nname == nbuyer and amount_matches(amount, g['total_amount']):
                return (g, 'medium', f"정규화일치({name}↔{buyer})+합산일치")
            if (nname in nbuyer or nbuyer in nname) and amount_matches(amount, g['total_amount']):
                return (g, 'medium', f"부분포함({name}↔{buyer})+합산일치")

    # 3) 합산금액만 일치 (1명) → medium
    amount_only = [g for g in grouped if amount_matches(amount, g['total_amount'])]
    if len(amount_only) == 1:
        g = amount_only[0]
        return (g, 'medium', f"합산금액만 일치({g['buyer_name']})")
    elif len(amount_only) > 1:
        return (amount_only[0], 'low', f"합산금액 같은 {len(amount_only)}명")

    # 4) 이름 유사 + 합산금액 다름 → low
    for name, nname in norm_searches.items():
        if not nname: continue
        for g in grouped:
            buyer = g['buyer_name'] or ''; nbuyer = _norm(buyer)
            if nname == nbuyer or (nname and nbuyer and (nname in nbuyer or nbuyer in nname)):
                return (g, 'low', f"이름유사({name}↔{buyer}) 합산다름({amount:,}↔{g['total_amount']:,})")

    return (None, None, f"후보없음 ({search_names}, {amount:,}원)")


def save_candidate(conn, session_id, source, source_ref,
                   paid_name, paid_name2, amount, paid_at,
                   buyer_dict, confidence, reason):
    if buyer_dict:
        buyer_name = buyer_dict.get('buyer_name')
        cand_amt   = buyer_dict.get('total_amount')
        first_id   = (buyer_dict.get('order_ids') or [None])[0]
    else:
        buyer_name = None; cand_amt = None; first_id = None
    # ⭐ 사용자가 이미 [❌ 거절]한 (paid_name ↔ buyer_name) pair는 다시 후보로 만들지 않음
    if paid_name and buyer_name:
        neg = conn.execute(
            "SELECT 1 FROM nick_mappings WHERE nickname=? AND realname=? AND COALESCE(negative,0)=1",
            (paid_name, buyer_name)
        ).fetchone()
        if neg:
            logger.info(f"  🚫 후보 스킵 (이전 거절): '{paid_name}'↔'{buyer_name}'")
            return
    try:
        conn.execute('''INSERT INTO match_candidates
              (session_id, source, source_ref, paid_name, paid_name2,
               amount, paid_at, candidate_order_id, candidate_buyer_name,
               candidate_amount, confidence, reason, status, created_at)
              VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'open',?)''',
              (session_id, source, str(source_ref) if source_ref else '',
               paid_name, paid_name2, amount, paid_at,
               first_id, buyer_name, cand_amt, confidence, reason,
               datetime.now().isoformat()))
        logger.info(f"  🤔 의심후보: src={source} '{paid_name}' {amount:,}원 → '{buyer_name}' 합계{cand_amt} ({confidence}: {reason})")
    except Exception as e:
        if 'UNIQUE' not in str(e):
            logger.warning(f"후보 저장 실패: {e}")


def match_sms_to_order(parsed, recv_time, sms_id=None):
    """SMS 입금 → SUM 합산 기준 매칭. high면 자동확인, 그 외는 의심후보 등록."""
    name   = parsed.get('name')
    amount = parsed['amount']
    today  = datetime.now().strftime('%Y-%m-%d')
    now    = datetime.now().isoformat()

    conn = get_conn()
    try:
        # 가장 최근 세션 (활성/비활성 무관) — pending 1건이라도 있으면 매칭 시도
        sessions = conn.execute('''
            SELECT * FROM live_sessions
            WHERE EXISTS (SELECT 1 FROM orders o WHERE o.session_id=live_sessions.id AND o.status='pending')
            ORDER BY live_date DESC
        ''').fetchall()
        if not sessions:
            logger.info(f"SMS: pending 보유 세션 없음 → 보류 ({name or '-'} {amount:,}원)")
            return

        search_names = resolve_names(conn, name) if name else []

        for sess in sessions:
            sid = sess['id']
            buyer, confidence, reason = find_buyer_match(
                conn, sid, search_names or [name or ''], amount
            )
            if confidence == 'high':
                for oid in buyer['order_ids']:
                    conn.execute('''UPDATE orders SET status='confirmed', confirmed_at=?,
                                           pay_type='transfer', bank_date=? WHERE id=?''',
                                 (now, recv_time, oid))
                if sms_id is not None:
                    conn.execute('UPDATE sms_payments SET matched=1 WHERE id=?', (sms_id,))
                else:
                    conn.execute('''UPDATE sms_payments SET matched=1
                                   WHERE parsed_amount=? AND COALESCE(parsed_name,'')=COALESCE(?,'') AND matched=0''',
                                 (amount, name))
                conn.commit()
                logger.info(f"  ✅ SMS 자동매칭: {buyer['buyer_name']} 합계 {buyer['total_amount']:,}원 ({len(buyer['order_ids'])}건)")
                return
            elif confidence in ('medium', 'low'):
                save_candidate(conn, sid, 'sms', sms_id, name, None, amount, recv_time,
                               buyer, confidence, reason)
                conn.commit()
                return
        logger.info(f"  ⚠️ SMS 매칭/후보 모두 실패: '{name or '-'}' {amount:,}원")
    except Exception as e:
        logger.error(f"SMS 매칭 오류: {e}")
    finally:
        conn.close()



def run_auto_check():
    """매일 30분마다 + 수동: 아임웹 카드결제 + 미매칭 SMS 재대조 (SUM 합산 + 의심후보)"""
    logger.info("자동 입금확인 시작...")
    today    = datetime.now().strftime('%Y-%m-%d')
    today_ym = today.replace('-', '')
    now_time = datetime.now().strftime('%H:%M:%S')

    conn = get_conn()
    try:
        # ⭐ pending 주문이 1건이라도 있는 모든 세션 처리 (check_end 만료 무관)
        sessions = conn.execute('''
            SELECT * FROM live_sessions
            WHERE EXISTS (SELECT 1 FROM orders o WHERE o.session_id=live_sessions.id AND o.status='pending')
            ORDER BY live_date DESC
        ''').fetchall()
        logger.info(f"처리할 세션: {len(sessions)}건 (pending 보유)")

        for session in sessions:
            session = dict(session)
            sid = session['id']
            logger.info(f"세션 확인: {session['filename']}")

            imweb_status = 'success'; sms_status = 'success'
            imweb_confirmed = 0; sms_confirmed = 0; error_msg = None

            # 1) 아임웹 카드결제 — 조회기간 라이브 ~ +30일 (오늘로 캡)
            try:
                live_dt = datetime.strptime(session['live_date'], '%Y-%m-%d')
                start_dt = live_dt
                end_dt = live_dt + timedelta(days=30)
                if end_dt > datetime.now(): end_dt = datetime.now()
                start_ym = start_dt.strftime('%Y%m%d')
                end_ym   = end_dt.strftime('%Y%m%d')
                imweb_orders = get_paid_orders(start_ym, end_ym)
                logger.info(f"아임웹 조회: {len(imweb_orders)}건 ({start_ym}~{end_ym})")

                for iorder in imweb_orders:
                    info = extract_order_info(iorder)
                    if not info['amount']:
                        continue
                    search_names = []
                    if info.get('name'):  search_names += resolve_names(conn, info['name'])
                    if info.get('name2'): search_names += resolve_names(conn, info['name2'])
                    search_names = list(dict.fromkeys(filter(None, search_names)))

                    order_no = str(iorder.get('order_no') or iorder.get('orderNo') or '')
                    buyer, conf, reason = find_buyer_match(conn, sid, search_names, info['amount'])

                    if conf == 'high':
                        for oid in buyer['order_ids']:
                            conn.execute("UPDATE orders SET status='confirmed', confirmed_at=?, pay_type='card' WHERE id=?",
                                         (info['paid_at'], oid))
                        imweb_confirmed += 1
                        logger.info(f"  ✅ 카드결제: {buyer['buyer_name']} 합계 {buyer['total_amount']:,}원")
                    elif conf in ('medium', 'low'):
                        save_candidate(conn, sid, 'imweb', order_no,
                                       info.get('name'), info.get('name2'),
                                       info['amount'], info['paid_at'],
                                       buyer, conf, reason)
            except Exception as e:
                imweb_status = 'error'
                error_msg = f"아임웹:{str(e)}"
                logger.error(f"아임웹 조회 오류: {e}")

            # 2) 미매칭 SMS 재대조 (SUM 합산)
            try:
                unmatched_sms = conn.execute("SELECT * FROM sms_payments WHERE matched=0").fetchall()
                for sms in unmatched_sms:
                    sms = dict(sms)
                    amt = sms.get('parsed_amount') or 0
                    sname = sms.get('parsed_name')
                    if not amt: continue
                    search_names = resolve_names(conn, sname) if sname else [sname or '']
                    buyer, conf, reason = find_buyer_match(conn, sid, search_names, amt)
                    if conf == 'high':
                        for oid in buyer['order_ids']:
                            conn.execute("UPDATE orders SET status='confirmed', confirmed_at=?, pay_type='transfer', bank_date=? WHERE id=?",
                                         (datetime.now().isoformat(), sms['received_at'], oid))
                        conn.execute("UPDATE sms_payments SET matched=1 WHERE id=?", (sms['id'],))
                        sms_confirmed += 1
                        logger.info(f"  ✅ SMS 재매칭: {buyer['buyer_name']} 합계 {buyer['total_amount']:,}원")
                    elif conf in ('medium', 'low'):
                        save_candidate(conn, sid, 'sms', sms['id'], sname, None,
                                       amt, sms['received_at'], buyer, conf, reason)
            except Exception as e:
                sms_status = 'error'
                error_msg = (error_msg or '') + f" SMS:{str(e)}"
                logger.error(f"SMS 재대조 오류: {e}")

            conn.execute("INSERT INTO check_logs (session_id, check_date, check_time, imweb_status, sms_status, imweb_confirmed, sms_confirmed, error_message) VALUES (?,?,?,?,?,?,?,?)",
                         (sid, today, now_time, imweb_status, sms_status, imweb_confirmed, sms_confirmed, error_msg))

        conn.commit()
        logger.info("자동 입금확인 종료")
    except Exception as e:
        logger.error(f"자동확인 오류: {e}")
        conn.rollback()
    finally:
        conn.close()


if __name__ == '__main__':
    init_db()

    scheduler = BackgroundScheduler(timezone='Asia/Seoul')
    scheduler.add_job(run_auto_check, 'interval', minutes=30,
                      id='interval_check', replace_existing=True)
    scheduler.start()
    logger.info("⏰ 스케줄러 시작 (30분마다 자동 확인)")

    PORT = int(os.environ.get('PORT', 5000))
    logger.info(f"🌿 지양하월시아 서버 시작! 포트: {PORT}")
    app.run(host='0.0.0.0', port=PORT, debug=False, use_reloader=False)
