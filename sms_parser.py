import re

# ── 농협 NH농협은행 패턴 ────────────────────────────────────────────────────────
NH_PATTERNS = [
    r'(?:농협|NH)은?행?.*?([가-힣]{2,5})(?:님)?[\s]*([0-9,]+)원.*?입금',
    r'(?:농협|NH)은?행?.*?입금.*?([가-힣]{2,5})(?:님)?[\s]*([0-9,]+)원',
    r'([가-힣]{2,5})(?:님)?.*?([0-9,]+)원.*?입금.*?(?:농협|NH)',
    r'(?:농협|NH)은?행?.*?입금\s*([0-9,]+)원\s*([가-힣]{2,5})',
]

# ── KB국민은행 패턴 ────────────────────────────────────────────────────────────
# 핵심: "국민은행" 텍스트 이후에 나오는 한글 이름을 잡아야 함
# 예1: "[KB국민은행] 05/01 14:23 홍길동 입금 30,000원 잔액 1,234,567원"
# 예2: "[Web발신][KB국민은행]입금 30,000원 홍길동 잔액..."
KB_PATTERNS = [
    # 이름 → 입금 → 금액
    r'(?:KB)?국민은행[^\가-힣]{0,20}([가-힣]{2,5})(?:님)?\s*입금\s*([0-9,]+)원',
    # 입금 → 금액 → 이름
    r'(?:KB)?국민은행.*?입금\s*([0-9,]+)원\s*([가-힣]{2,5})(?:님)?',
    # 이름이 앞에 오는 경우
    r'([가-힣]{2,5})(?:님)?\s*입금\s*([0-9,]+)원.*?(?:KB)?국민',
]


def _clean(text):
    return text.replace('\n', ' ').replace('\r', ' ').strip()


def _parse_amount(raw):
    try:
        return int(raw.replace(',', ''))
    except Exception:
        return 0


# 은행명 관련 단어 (이름으로 잡히면 안 되는 것들)
BANK_WORDS = {'은행', '농협', '국민', '입금', '잔액', '출금', '이체', '거래'}


def _is_valid_name(name):
    """이름으로 유효한지 확인 (은행명 단어 제외)"""
    name = name.replace('님', '').strip()
    if name in BANK_WORDS:
        return False
    if len(name) < 2:
        return False
    return True


def parse_sms(body):
    """
    농협 or 국민은행 입금 문자 파싱
    Returns: {'name': str, 'amount': int, 'bank': str} | None
    """
    if not body:
        return None

    body = _clean(body)

    if '입금' not in body:
        return None

    is_nh = '농협' in body or 'NH' in body
    is_kb = 'KB국민' in body or '국민은행' in body

    if not is_nh and not is_kb:
        return None

    patterns = NH_PATTERNS if is_nh else KB_PATTERNS
    bank = '농협' if is_nh else '국민'

    for pattern in patterns:
        m = re.search(pattern, body)
        if not m:
            continue
        g = m.groups()
        if len(g) < 2:
            continue

        g0_is_amount = bool(re.fullmatch(r'[0-9,]+', g[0]))
        if g0_is_amount:
            amount = _parse_amount(g[0])
            name = g[1].replace('님', '').strip()
        else:
            name = g[0].replace('님', '').strip()
            amount = _parse_amount(g[1])

        if _is_valid_name(name) and amount > 0:
            return {'name': name, 'amount': amount, 'bank': bank}

    return None
