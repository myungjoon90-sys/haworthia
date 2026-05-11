import re

# ── 농협 NH농협은행 패턴 ────────────────────────────────────────────────────────
NH_PATTERNS = [
    r'(?:농협|NH)은?행?.*?([가-힣]{2,5})(?:님)?[\s]*([0-9,]+)원.*?입금',
    r'(?:농협|NH)은?행?.*?입금.*?([가-힣]{2,5})(?:님)?[\s]*([0-9,]+)원',
    r'([가-힣]{2,5})(?:님)?.*?([0-9,]+)원.*?입금.*?(?:농협|NH)',
    r'(?:농협|NH)은?행?.*?입금\s*([0-9,]+)원\s*([가-힣]{2,5})',
]

# ── KB국민은행 패턴 ────────────────────────────────────────────────────────────
# 실제 문자 형식:
# [Web발신]\n[KB]01/20 10:13\n971372**874\n박희자\n입금\n141,000\n잔액1,033,072
KB_PATTERNS = [
    # 실제 KB 형식: 계좌번호 다음 줄에 이름, 그 다음 입금, 그 다음 금액
    r'(?:KB|국민).*?(?:\d[\d*]+)\s+([가-힣]{2,5})\s+입금\s+([0-9,]+)',
    # 이름 → 입금 → 금액 (한 줄)
    r'(?:KB)?국민은?행?[^\가-힣]{0,20}([가-힣]{2,5})(?:님)?\s*입금\s*([0-9,]+)원?',
    # 입금 → 금액 → 이름
    r'(?:KB)?국민은?행?.*?입금\s*([0-9,]+)원?\s*([가-힣]{2,5})(?:님)?',
    # 이름이 앞에 오는 경우
    r'([가-힣]{2,5})(?:님)?\s*입금\s*([0-9,]+)원?.*?(?:KB)?국민',
]

# 은행명 관련 단어 (이름으로 잡히면 안 되는 것들)
BANK_WORDS = {'은행', '농협', '국민', '입금', '잔액', '출금', '이체', '거래'}


def _clean(text):
    return text.replace('\r', ' ').strip()


def _parse_amount(raw):
    try:
        return int(str(raw).replace(',', '').replace('원', '').strip())
    except Exception:
        return 0


def _is_valid_name(name):
    name = name.replace('님', '').strip()
    if name in BANK_WORDS:
        return False
    if len(name) < 2:
        return False
    return True


def _parse_kb_multiline(body):
    """
    실제 KB국민은행 줄바꿈 형식 파싱
    예: [Web발신]\n[KB]01/20 10:13\n971372**874\n박희자\n입금\n141,000\n잔액1,033,072
    """
    lines = [l.strip() for l in body.split('\n') if l.strip()]
    name = None
    amount = None

    for i, line in enumerate(lines):
        # 이름 찾기: 한글 2~5자이고 은행 단어 아닌 것
        if re.fullmatch(r'[가-힣]{2,5}', line) and line not in BANK_WORDS:
            name = line
        # 금액 찾기: 숫자+콤마로만 구성
        if re.fullmatch(r'[\d,]+', line):
            amt = _parse_amount(line)
            if amt > 0 and amt < 100000000:  # 1억 이하
                amount = amt

    if name and amount:
        return {'name': name, 'amount': amount, 'bank': '국민'}
    return None


def _parse_nh_multiline(body):
    """농협 줄바꿈 형식 파싱"""
    lines = [l.strip() for l in body.split('\n') if l.strip()]
    name = None
    amount = None

    for line in lines:
        if re.fullmatch(r'[가-힣]{2,5}', line) and line not in BANK_WORDS:
            name = line
        if re.fullmatch(r'[\d,]+', line):
            amt = _parse_amount(line)
            if amt > 0 and amt < 100000000:
                amount = amt
        # "30,000원" 형식
        m = re.fullmatch(r'([\d,]+)원', line)
        if m:
            amt = _parse_amount(m.group(1))
            if amt > 0:
                amount = amt

    if name and amount:
        return {'name': name, 'amount': amount, 'bank': '농협'}
    return None


def parse_sms(body):
    """
    농협 or 국민은행 입금 문자 파싱
    Returns: {'name': str, 'amount': int, 'bank': str} | None
    """
    if not body:
        return None

    body_clean = _clean(body)

    if '입금' not in body_clean:
        return None

    is_nh = '농협' in body_clean or 'NH' in body_clean
    is_kb = 'KB' in body_clean or '국민' in body_clean

    if not is_nh and not is_kb:
        return None

    # 줄바꿈 있는 경우 먼저 멀티라인 파싱 시도
    if '\n' in body:
        if is_kb:
            result = _parse_kb_multiline(body)
            if result:
                return result
        if is_nh:
            result = _parse_nh_multiline(body)
            if result:
                return result

    # 한 줄 패턴 매칭
    body_oneline = body_clean.replace('\n', ' ')
    patterns = NH_PATTERNS if is_nh else KB_PATTERNS
    bank = '농협' if is_nh else '국민'

    for pattern in patterns:
        m = re.search(pattern, body_oneline)
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
