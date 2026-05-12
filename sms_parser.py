import re

# ══════════════════════════════════════════════════════════════════
#  실제 SMS 형식
#
#  NH농협:
#  알림
#  [Web발신]
#  농협 입금250,000원
#  05/03 13:26 301-****-2024-21 양보라 잔액1,865,041원
#  → 이름: 양보라, 금액: 250,000
#
#  KB국민:
#  [Web발신]
#  [KB]01/20 10:13
#  971372**874
#  고체샤
#  입금
#  100,000
#  잔액1,033,072
#  → 이름: 고체샤, 금액: 100,000
# ══════════════════════════════════════════════════════════════════

BANK_WORDS = {'은행', '농협', '국민', '입금', '잔액', '출금', '이체', '거래', 'web발신', 'kb', '알림'}


def _parse_amount(raw):
    try:
        return int(str(raw).replace(',', '').replace('원', '').strip())
    except Exception:
        return 0


def _is_valid_name(name):
    name = name.replace('님', '').strip().lower()
    if not name or len(name) < 2:
        return False
    if name in BANK_WORDS:
        return False
    if re.fullmatch(r'[\d\-\*\.\s]+', name):
        return False
    return True


def parse_nh_sms(body):
    """
    NH농협 형식 파싱
    예:
    농협 입금250,000원
    05/03 13:26 301-****-2024-21 양보라 잔액1,865,041원
    → 이름: 양보라, 금액: 250,000
    """
    amt_m = re.search(r'입금([0-9,]+)원', body)
    if not amt_m:
        return None
    amount = _parse_amount(amt_m.group(1))
    if amount <= 0:
        return None

    # 이름: "잔액" 바로 앞 한글 단어
    name = None
    name_m = re.search(r'([가-힣]{2,6})\s*잔액', body)
    if name_m:
        candidate = name_m.group(1).strip()
        if _is_valid_name(candidate):
            name = candidate

    return {'name': name, 'amount': amount, 'bank': '농협'}


def parse_kb_sms(body):
    """
    KB국민은행 형식 파싱 (여러 줄)
    줄 순서: [Web발신] → [KB]날짜 → 계좌번호 → 이름 → 입금 → 금액 → 잔액
    """
    lines = [l.strip() for l in body.replace('\r', '').split('\n') if l.strip()]

    name = None
    amount = None

    for i, line in enumerate(lines):
        if line == '입금' or line == '입금액':
            # 이름: "입금" 앞 줄 중 유효한 이름 (한글 or 영문 닉네임)
            for j in range(i - 1, max(i - 4, -1), -1):
                candidate = lines[j]
                if re.fullmatch(r'[가-힣a-zA-Z\w\-_\.]{2,20}', candidate) and _is_valid_name(candidate):
                    name = candidate.replace('님', '').strip()
                    break
            # 금액: "입금" 다음 줄
            if i + 1 < len(lines):
                amount = _parse_amount(lines[i + 1])
            break

    if not amount:
        m = re.search(r'입금\s*([0-9,]+)원?', body)
        if m:
            amount = _parse_amount(m.group(1))

    if not name:
        m = re.search(r'([가-힣]{2,6})\s*잔액', body)
        if m and _is_valid_name(m.group(1)):
            name = m.group(1)

    if amount and amount > 0:
        return {'name': name, 'amount': amount, 'bank': '국민'}
    return None


def parse_sms(body):
    """
    농협 or 국민은행 입금 문자 파싱
    Returns: {'name': str or None, 'amount': int, 'bank': str} | None
    """
    if not body or '입금' not in body:
        return None

    body_clean = body.replace('\r', '').strip()

    is_nh = '농협' in body_clean
    is_kb = 'KB' in body_clean or '국민은행' in body_clean or '국민' in body_clean

    if is_nh:
        return parse_nh_sms(body_clean)
    elif is_kb:
        return parse_kb_sms(body_clean)

    return None
