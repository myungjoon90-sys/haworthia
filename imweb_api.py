import requests
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

IMWEB_API_KEY = "d62fe2aadfe7344ec34482b6c1654eb7d076320092"
IMWEB_SECRET  = "2e98914f0564d9eb1cdaa6"

_token = None
_token_expires = None


def invalidate_token():
    """캐시된 토큰을 강제로 무효화 (아임웹이 토큰 거부 시 호출됨)"""
    global _token, _token_expires
    _token = None
    _token_expires = None
    logger.info("🔄 아임웹 토큰 캐시 무효화 → 다음 요청 시 재발급")


def get_access_token(force_refresh=False):
    """
    아임웹 액세스 토큰 발급/조회
    force_refresh=True 면 캐시 무시하고 무조건 새로 받음
    """
    global _token, _token_expires
    now = datetime.now()
    if not force_refresh and _token and _token_expires and now < _token_expires:
        return _token
    try:
        resp = requests.get(
            "https://api.imweb.me/v2/auth",
            params={"key": IMWEB_API_KEY, "secret": IMWEB_SECRET},
            timeout=10
        )
        data = resp.json()
        if data.get("code") == 200:
            _token = data["access_token"]
            # 안전마진: 1시간 → 50분으로 단축 (아임웹이 조기 만료시키는 경우 대비)
            _token_expires = now + timedelta(minutes=50)
            logger.info("✅ 아임웹 토큰 발급 성공 (50분간 캐시)")
            return _token
        else:
            logger.error(f"아임웹 토큰 발급 실패: {data}")
            return None
    except Exception as e:
        logger.error(f"아임웹 연결 오류: {e}")
        return None


def _is_token_error(data):
    """응답이 토큰 관련 에러인지 판단"""
    if not data:
        return False
    code = data.get("code")
    msg = (data.get("msg") or "").lower()
    # code = -2 ('Error Token'), 401 (Unauthorized), 또는 msg에 'token' 포함
    return code == -2 or code == 401 or "token" in msg


def get_paid_orders(start_date, end_date):
    """
    start_date, end_date: 'YYYYMMDD' 형식
    결제완료 주문 목록 반환
    ★ 토큰 만료/오류 시 자동으로 새 토큰 받아 재시도 (사이클당 1회)
    """
    # YYYYMMDD → YYYY-MM-DD 형식 변환
    def fmt(d):
        if d and len(d) == 8:
            return f"{d[:4]}-{d[4:6]}-{d[6:]}"
        return d

    start_fmt = fmt(start_date) + " 00:00:00"
    end_fmt   = fmt(end_date)   + " 23:59:59"
    logger.info(f"아임웹 주문 조회 기간: {start_fmt} ~ {end_fmt}")

    all_orders = []
    page = 1
    retried = False  # 한 호출당 토큰 재시도 1회 제한

    while True:
        token = get_access_token()
        if not token:
            logger.warning("아임웹 토큰 없음 → 조회 중단")
            break

        try:
            resp = requests.get(
                "https://api.imweb.me/v2/shop/orders",
                headers={"access-token": token},
                params={
                    "order_date_from": start_fmt,
                    "order_date_to":   end_fmt,
                    "limit": 100,
                    "page": page
                },
                timeout=10
            )
            data = resp.json()

            logger.info(f"아임웹 응답: code={data.get('code')} msg={data.get('msg','')[:50]}")

            # ★★ 토큰 에러 자동 복구 ★★
            if _is_token_error(data) and not retried:
                logger.warning(f"⚠️  토큰 에러 감지(code={data.get('code')}) → 토큰 재발급 후 재시도")
                invalidate_token()
                retried = True
                continue  # 같은 page로 다시 시도

            if data.get("code") != 200:
                logger.warning(f"아임웹 주문 조회 실패: {str(data)[:300]}")
                break

            items = data.get("data", {}).get("list", [])
            logger.info(f"아임웹 주문 수: {len(items)}건 (page {page})")

            # 첫 페이지 첫 주문의 키 구조 로깅 (디버그용)
            if items and page == 1:
                first = items[0]
                payment = first.get('payment') or {}
                orderer = first.get('orderer') or {}
                logger.info(f"아임웹 주문 필드: {list(first.keys())[:20]}")
                logger.info(f"아임웹 payment 필드: {list(payment.keys())[:15]}")
                logger.info(f"아임웹 샘플: name={orderer.get('name','')} price={payment.get('price','')} pay_price={payment.get('pay_price','')}")

            if not items:
                break

            all_orders.extend(items)

            if len(items) < 100:
                break
            page += 1

        except Exception as e:
            logger.error(f"아임웹 주문 조회 오류: {e}")
            break

    return all_orders


def extract_order_info(iorder):
    """아임웹 V2 주문 데이터에서 핵심 정보 추출"""
    import re as _re
    orderer = iorder.get("orderer") or {}
    payment = iorder.get("payment") or {}

    # 주문자 이름: orderer.name 우선
    raw_name = (
        orderer.get("name") or
        iorder.get("member_id") or
        iorder.get("member_name") or
        iorder.get("orderer_name") or
        iorder.get("name") or ""
    ).strip()

    # "문성옥(미아옹)" → 닉네임/실명 분리
    name2 = None
    m = _re.search(r'\(([^)]+)\)', raw_name)
    if m:
        name  = m.group(1).strip()
        name2 = raw_name[:raw_name.index('(')].strip()
    else:
        name = raw_name

    # 결제 금액: ⭐ payment.pay_price = 실제 결제금액 (상품+배송비-할인)
    #   payment.price 는 상품금액만이라 배송비가 빠진다 → fallback으로만 사용
    #   pay_price 없으면 price + delivery_price 로 보정
    pay_price_v = (payment.get("pay_price") or payment.get("payment_price")
                   or payment.get("total_price"))
    if not pay_price_v:
        base = payment.get("price") or iorder.get("price") or 0
        ship = (payment.get("delivery_price") or payment.get("shipping_price")
                or payment.get("deliv_price") or iorder.get("delivery_price") or 0)
        try:
            base = int(base); ship = int(ship)
        except Exception:
            base = 0; ship = 0
        pay_price_v = base + ship if base else 0

    try:
        amount = int(pay_price_v or 0)
    except (TypeError, ValueError):
        amount = 0

    # 진단: price와 pay_price가 다르면 (=배송비 있음) 어느 값이 쓰였는지 로그
    try:
        _p = payment.get("price"); _pp = payment.get("pay_price")
        _dp = payment.get("delivery_price")
        if _p and _pp and int(_p) != int(_pp):
            logger.info(f"  💳 결제금액: price={_p} pay_price={_pp} delivery={_dp} → 사용={amount}")
    except Exception:
        pass

    # 결제 시각
    paid_at = (
        payment.get("complete_time") or
        payment.get("pay_time") or
        iorder.get("complete_time") or
        iorder.get("order_time") or
        datetime.now().isoformat()
    )

    # 상품명
    prod_list = iorder.get("prod_list") or iorder.get("prod_items") or []
    item = ", ".join(p.get("prod_name", "") for p in prod_list) if prod_list else ""

    return {"name": name, "name2": name2, "amount": amount, "paid_at": paid_at, "item": item}



# 알려진 "조용히 넘어가도 되는" 코드들 — 이미 그 상태거나 의미 없는 호출
_STANDBY_SKIP_KEYWORDS = (
    '이미', '동일', '준비중이 아', '발송', '배송중', '배송완', '취소', '환불',
    'already', 'same state', 'invalid status', 'not preparing'
)

def get_order_products(order_no):
    """주문의 품목 상품명들을 ', '로 합쳐 반환. 실패 시 ''.
    GET /v2/shop/orders/{order_no}/prod-orders → data.list[].items[].prod_name"""
    if not order_no:
        return ''
    token = get_access_token()
    if not token:
        return ''
    url = f"https://api.imweb.me/v2/shop/orders/{order_no}/prod-orders"
    def _send(tok):
        return requests.get(url, headers={"access-token": tok},
                            params={"order_version": "v2"}, timeout=10)
    try:
        resp = _send(token)
        data = resp.json() if resp.text else {}
        if _is_token_error(data):
            invalidate_token()
            token = get_access_token(force_refresh=True)
            if token:
                resp = _send(token)
                data = resp.json() if resp.text else {}
        if data.get('code') != 200:
            return ''
        d = data.get('data')
        plist = []
        if isinstance(d, dict):
            plist = d.get('list') or d.get('prod_orders') or []
        elif isinstance(d, list):
            plist = d
        names = []
        for po in (plist or []):
            for it in (po.get('items') or []):
                nm = (it.get('prod_name') or '').strip()
                if nm:
                    names.append(nm)
        return ', '.join(dict.fromkeys(names))
    except Exception as e:
        logger.warning(f"prod-orders 조회 실패 {order_no}: {e}")
        return ''


def set_order_to_standby(order_no, prod_order_nos=None):
    """아임웹 주문을 '상품 준비중' → '배송대기'로 전환.
    PATCH /v2/shop/orders/{order_no}/place
    body: {prod_order_nos: [...]} — 아임웹 V2는 prod_order_no 리스트가 필수.
    [v22] 자세한 진단 로그 + 알려진 상태(이미 배송대기/배송중/취소 등)는 info로 강등.
    """
    token = get_access_token()
    if not token:
        logger.warning(f"[imweb] 토큰 없음 → 배송대기 전환 스킵: {order_no}")
        return False
    url = f"https://api.imweb.me/v2/shop/orders/{order_no}/place"
    body = {}
    if prod_order_nos:
        body['prod_order_nos'] = prod_order_nos if isinstance(prod_order_nos, list) else [prod_order_nos]

    def _send(tok):
        return requests.patch(
            url,
            headers={"access-token": tok, "Content-Type": "application/json"},
            json=body, timeout=10
        )

    try:
        resp = _send(token)
        try:
            data = resp.json() if resp.text else {}
        except Exception:
            data = {'_raw_text': resp.text[:300]}

        # 토큰 만료 시 1회 재시도
        if _is_token_error(data):
            logger.info(f"  [imweb standby] 토큰 만료 감지 → 재발급 후 재시도: {order_no}")
            invalidate_token()
            token = get_access_token(force_refresh=True)
            if token:
                resp = _send(token)
                try: data = resp.json() if resp.text else {}
                except Exception: data = {'_raw_text': resp.text[:300]}

        code = data.get('code')
        msg  = (data.get('msg') or data.get('message') or '').strip()

        if code == 200:
            logger.info(f"  🚚 아임웹 배송대기 전환 성공: {order_no}")
            return True

        # 알려진 "괜찮은" 실패 — 이미 그 상태이거나 의미가 없는 경우
        msg_l = msg.lower()
        if any(k in msg or k.lower() in msg_l for k in _STANDBY_SKIP_KEYWORDS):
            logger.info(f"  ℹ️ 아임웹 배송대기: '{order_no}' → 스킵 (사유: {msg or 'code='+str(code)})")
            return True  # 사용자 입장에선 "이미 됐다"이므로 성공 처리

        # 진단 강화: HTTP 상태/코드/메시지/요청바디 전부 로깅
        logger.warning(
            f"  ⚠️ 아임웹 배송대기 전환 실패: order={order_no} "
            f"http={resp.status_code} api_code={code} msg='{msg or str(data)[:160]}' "
            f"body_sent={body}"
        )
        return False
    except requests.exceptions.Timeout:
        logger.error(f"  ❌ 아임웹 배송대기 타임아웃 (10초): {order_no}")
        return False
    except requests.exceptions.ConnectionError as e:
        logger.error(f"  ❌ 아임웹 배송대기 연결 오류: {order_no} → {e}")
        return False
    except Exception as e:
        logger.error(f"  ❌ 아임웹 배송대기 전환 예외: {order_no} → {type(e).__name__}: {e}")
        return False
