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

    all_orders = []
    page = 1

    while True:
        try:
            resp = requests.get(
                "https://api.imweb.me/v2/shop/orders",
                headers={"access-token": token},
                params={
                    "order_date_from": start_date,
                    "order_date_to": end_date,
                    "limit": 100,
                    "page": page
                },
                timeout=10
            )
            data = resp.json()

            if data.get("code") != 200:
                logger.warning(f"아임웹 주문 조회 실패: {data}")
                break

            items = data.get("data", {}).get("list", [])
            if not items:
                break

            all_orders.extend(items)
            logger.info(f"아임웹 주문 조회: {len(items)}건 (page {page})")

            if len(items) < 100:
                break
            page += 1

        except Exception as e:
            logger.error(f"아임웹 주문 조회 오류: {e}")
            break

    return all_orders

def extract_order_info(iorder):
    """아임웹 주문 데이터에서 핵심 정보 추출"""
    import re as _re
    orderer = iorder.get("orderer") or {}

    # 여러 필드 시도 (아임웹은 버전마다 필드명이 다름)
    raw_name = (
        iorder.get("member_id") or       # 아임웹 회원 아이디 (닉네임)
        orderer.get("name") or
        iorder.get("member_name") or
        iorder.get("orderer_name") or
        iorder.get("name") or ""
    ).strip()

    # "문성옥(미아옹)" 형태 → 닉네임과 실명 분리
    name2 = None
    m = _re.search(r'\(([^)]+)\)', raw_name)
    if m:
        name   = m.group(1).strip()              # 닉네임: 미아옹
        name2  = raw_name[:raw_name.index('(')].strip()  # 실명: 문성옥
    else:
        name = raw_name

    amount = int(
        iorder.get("pay_price") or
        iorder.get("total_price") or
        iorder.get("price") or 0
    )
    paid_at = (
        iorder.get("pay_date") or
        iorder.get("order_date") or
        datetime.now().isoformat()
    )
    prod_list = iorder.get("prod_list") or []
    item = ", ".join(p.get("prod_name", "") for p in prod_list) if prod_list else ""

    return {"name": name, "name2": name2, "amount": amount, "paid_at": paid_at, "item": item}
