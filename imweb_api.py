import requests
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

IMWEB_API_KEY = "d62fe2aadfe7344ec34482b6c1654eb7d076320092"
IMWEB_SECRET  = "2e98914f0564d9eb1cdaa6"

_token = None
_token_expires = None

def get_access_token():
    global _token, _token_expires
    now = datetime.now()
    if _token and _token_expires and now < _token_expires:
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
            _token_expires = now + timedelta(hours=1)
            logger.info("✅ 아임웹 토큰 발급 성공")
            return _token
        else:
            logger.error(f"아임웹 토큰 발급 실패: {data}")
            return None
    except Exception as e:
        logger.error(f"아임웹 연결 오류: {e}")
        return None

def get_paid_orders(start_date, end_date):
    """
    start_date, end_date: 'YYYYMMDD' 형식
    결제완료 주문 목록 반환
    """
    token = get_access_token()
    if not token:
        return []

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

    while True:
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

            # 디버그: 응답 코드와 첫 주문 구조 로깅
            logger.info(f"아임웹 응답: code={data.get('code')} msg={data.get('msg','')[:50]}")

            if data.get("code") != 200:
                logger.warning(f"아임웹 주문 조회 실패: {str(data)[:300]}")
                break

            items = data.get("data", {}).get("list", [])
            logger.info(f"아임웹 주문 수: {len(items)}건 (page {page})")

            # 첫 주문의 키 구조 로깅
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

    # 결제 금액: payment 객체 안에서 추출
    amount = int(
        payment.get("price") or
        payment.get("pay_price") or
        payment.get("total_price") or
        payment.get("payment_price") or
        iorder.get("pay_price") or
        iorder.get("total_price") or 0
    )

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
