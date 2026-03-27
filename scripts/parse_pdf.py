#!/usr/bin/env python3
"""
매일아침 — PDF-only 파서

레시피 엑셀 없이, 어린이집 식단표 PDF 한 장에서
메뉴명 + 알레르기 번호를 추출하여 build.py가 사용하는 JSON을 생성한다.

지원 형식:
  - 주 5일 (월~금) 테이블
  - 5-row-per-week 블록: 일자 / 오전간식 / 점심 / 영양량 / 오후간식
  - 알레르기 표기: 메뉴명(1,2,5,6) — 식약처 전국 표준

사용법:
  from parse_pdf import parse_pdf_menu
  data = parse_pdf_menu("식단표.pdf")
  # → { year, month, menus: { "2026-03-02": [...], ... } }
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import pdfplumber

# ── 19종 알레르기 ──
ALLERGEN_NAMES = {
    1: "난류(계란)", 2: "우유", 3: "메밀", 4: "땅콩", 5: "대두",
    6: "밀", 7: "고등어", 8: "게", 9: "새우", 10: "돼지고기",
    11: "복숭아", 12: "토마토", 13: "아황산류", 14: "호두", 15: "닭고기",
    16: "쇠고기", 17: "오징어", 18: "조개류", 19: "잣"
}

# 알레르기 번호 패턴: (1,2,5,6) or （1,2,5,6）
ALLERGY_PATTERN = re.compile(r'[\(（]([\d,.\s]+)[\)）]')

# 메뉴명 + 알레르기 번호를 분리하는 패턴
# "돈육표고장조림(5,6,10)" → ("돈육표고장조림", [5,6,10])
ITEM_SPLIT = re.compile(r'([^(（]+)[\(（]([\d,.\s]+)[\)）]')

# 날짜 패턴: "2(월)", "10(화)", "31(화)", "12목)" (괄호 누락 허용) 등
DATE_PATTERN = re.compile(r'(\d{1,2})\s*[\(（]?\s*[월화수목금토일]\s*[\)）]')


def _parse_allergy_nums(s: str) -> list[int]:
    """'1,2,5,6' → [1, 2, 5, 6]"""
    nums = []
    for n in re.findall(r'\d+', s):
        v = int(n)
        if 1 <= v <= 19:
            nums.append(v)
    return nums


def _split_menu_items(cell_text: str) -> list[dict]:
    """
    셀 텍스트에서 개별 메뉴 아이템을 추출한다.

    처리해야 할 패턴:
    1. 줄바꿈으로 구분: "잡곡밥\\n냉이된장국(5,6)\\n돈육표고장조림(5,6,10)"
    2. 쉼표 구분: "우유(2),미니단팥빵(2,5,6)"
    3. 붙어있는 아이템: "저당시리얼(2,5,6)우유(2)"
    4. 알레르기 없는 아이템: "바나나", "딸기"
    5. 잘린 텍스트: "오이사과무침(5," → 알레르기 번호 불완전
    """
    if not cell_text or not cell_text.strip():
        return []

    items = []
    # 먼저 줄바꿈으로 분리
    lines = cell_text.strip().split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # "점심\n(공통)" 같은 라벨은 스킵
        if line in ('(공통)', '점심', '오전간식', '오후간식', '방과후과정'):
            continue

        # 줄 안에서 붙어있는 아이템 분리
        # 패턴: "메뉴A(1,2)메뉴B(3,4)" or "메뉴A(1,2),메뉴B(3,4)"
        sub_items = _split_concatenated(line)
        items.extend(sub_items)

    return items


def _split_concatenated(text: str) -> list[dict]:
    """
    붙어있거나 쉼표로 연결된 아이템을 분리한다.

    "저당시리얼(2,5,6)우유(2)" → [저당시리얼, 우유]
    "우유(2),미니단팥빵(2,5,6)" → [우유, 미니단팥빵]
    "크림스프(2,5,6,15,16)모닝빵(1,2,5,6)" → [크림스프, 모닝빵]
    "미니붕어빵(1,5,6) 포도주스" → [미니붕어빵, 포도주스]
    """
    results = []

    # 전략: 알레르기 괄호 닫힘 뒤에 한글이 오면 새 아이템 시작
    # "메뉴A(1,2,3)메뉴B(4,5)" → split at ")메"
    # Also handle comma separator: "메뉴A(1,2),메뉴B(4,5)"

    # 모든 아이템 찾기: 이름 + (숫자들) 패턴
    # 정규식으로 "이름(숫자)" 쌍을 모두 찾는다
    parts = []
    pos = 0
    pattern = re.compile(
        r'([가-힣a-zA-Z0-9\-&*/ ]+?)'  # 메뉴명 (최소 매칭)
        r'[\(（]([\d,.\s]+?)[\)）]'       # (알레르기 번호)
    )

    for m in pattern.finditer(text):
        # 매치 앞에 남은 텍스트 처리 (알레르기 없는 아이템)
        before = text[pos:m.start()].strip().strip(',').strip()
        if before and not re.match(r'^[\d,.\s()（）]+$', before):
            # 알레르기 없는 독립 아이템
            for name in _split_plain(before):
                results.append({"name": name, "allergy": []})

        name = m.group(1).strip().strip(',').strip()
        nums = _parse_allergy_nums(m.group(2))

        if name:
            results.append({"name": name, "allergy": nums})

        pos = m.end()

    # 매치 뒤에 남은 텍스트 (알레르기 없는 아이템)
    remainder = text[pos:].strip().strip(',').strip()
    if remainder and not re.match(r'^[\d,.\s()（）]+$', remainder):
        for name in _split_plain(remainder):
            results.append({"name": name, "allergy": []})

    # 매치가 하나도 없으면 전체를 알레르기 없는 아이템으로
    if not results and text.strip():
        for name in _split_plain(text.strip()):
            results.append({"name": name, "allergy": []})

    return results


def _split_plain(text: str) -> list[str]:
    """알레르기 번호가 없는 텍스트에서 메뉴명 추출 (쉼표/공백 구분)"""
    names = []
    for part in re.split(r'[,，]', text):
        part = part.strip()
        if part and len(part) >= 1 and re.search(r'[가-힣a-zA-Z]', part):
            # 영양량 같은 숫자만 있는 건 스킵
            if not re.match(r'^[\d.]+\s*(kcal|Kcal|g)$', part, re.IGNORECASE):
                names.append(part)
    return names


def _detect_year_month(header_text: str, table) -> tuple[int, int]:
    """테이블 헤더에서 연도/월 감지"""
    # "2026년 3월 식단표" 패턴
    m = re.search(r'(\d{4})\s*년\s*(\d{1,2})\s*월', header_text)
    if m:
        return int(m.group(1)), int(m.group(2))

    # 날짜 행에서 추론
    for row in table[1:6]:
        for cell in row:
            if cell and DATE_PATTERN.search(str(cell)):
                # 날짜는 있지만 연/월은 헤더에서 못 찾은 경우
                # 현재 연도 가정
                from datetime import date as dt
                today = dt.today()
                return today.year, today.month

    raise ValueError("연도/월을 감지할 수 없습니다")


def parse_pdf_menu(pdf_path: str) -> dict:
    """
    식단표 PDF에서 메뉴 + 알레르기 데이터를 추출한다.

    Returns:
        {
            "year": 2026,
            "month": 3,
            "menus": {
                "2026-03-02": [
                    {"name": "잡곡밥", "allergy": [], "sec": "점심"},
                    {"name": "냉이된장국", "allergy": [5,6], "sec": "점심"},
                    ...
                ],
                ...
            }
        }
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()

    if not tables:
        raise ValueError(f"PDF에서 테이블을 찾을 수 없습니다: {pdf_path}")

    table = tables[0]
    if len(table) < 6:
        raise ValueError(f"테이블 행 수가 너무 적습니다: {len(table)} rows")

    # 연/월 감지
    header = str(table[0][0] or '')
    year, month = _detect_year_month(header, table)
    print(f"  [PDF] 감지: {year}년 {month}월")

    # 테이블 파싱
    # 구조: Row 0 = 헤더, 그 다음 5행씩 반복 (일자/오전간식/점심/영양량/오후간식)
    menus = {}  # date_str → list of items

    ri = 1  # Row 0은 헤더
    while ri + 4 < len(table):
        row_label = str(table[ri][0] or '').strip()

        # "일자" 행 찾기
        if '일자' not in row_label:
            ri += 1
            continue

        date_row = table[ri]
        snack_am_row = table[ri + 1]
        lunch_row = table[ri + 2]
        # nutrition_row = table[ri + 3]  # 영양량 — 스킵
        snack_pm_row = table[ri + 4]

        # 날짜 파싱: 컬럼 1,3,5,7,9 (홀수 인덱스는 월~금)
        day_cols = [1, 3, 5, 7, 9]

        for ci in day_cols:
            if ci >= len(date_row):
                continue

            date_cell = str(date_row[ci] or '').strip()
            dm = DATE_PATTERN.search(date_cell)
            if not dm:
                continue

            day_num = int(dm.group(1))
            try:
                ds = date(year, month, day_num).isoformat()
            except ValueError:
                continue

            day_items = []

            # 오전간식
            am_cell = str(snack_am_row[ci] or '') if ci < len(snack_am_row) else ''
            for item in _split_menu_items(am_cell):
                item['sec'] = '오전간식'
                day_items.append(item)

            # 점심
            lunch_cell = str(lunch_row[ci] or '') if ci < len(lunch_row) else ''
            for item in _split_menu_items(lunch_cell):
                item['sec'] = '점심'
                day_items.append(item)

            # 오후간식
            pm_cell = str(snack_pm_row[ci] or '') if ci < len(snack_pm_row) else ''
            for item in _split_menu_items(pm_cell):
                item['sec'] = '오후간식'
                day_items.append(item)

            if day_items:
                menus[ds] = day_items

        ri += 5  # 다음 주 블록으로

    print(f"  [PDF] {len(menus)}일 추출, 총 {sum(len(v) for v in menus.values())}개 아이템")
    return {
        "year": year,
        "month": month,
        "menus": menus
    }


def build_json_from_pdf(pdf_data: dict) -> dict:
    """
    parse_pdf_menu() 결과를 build.py/index.html이 사용하는 JSON 포맷으로 변환.
    레시피 없는 PDF-only 데이터이므로 recipe='', sauce 분석 없음.
    """
    from datetime import date as dt, timedelta

    year = pdf_data['year']
    month = pdf_data['month']
    menus = pdf_data['menus']

    # 메뉴 데이터 변환
    alt_pattern = re.compile(r'^만\s*\d')
    menu_data = {}

    for ds, items in sorted(menus.items()):
        js_items = []
        for item in items:
            name = item['name']

            # 대체 식단 감지
            if alt_pattern.match(name) and js_items:
                js_items[-1]['alt'] = name
                continue

            js_item = {
                "name": name,
                "recipe": "",  # PDF-only: 레시피 없음
                "allergy": item.get('allergy', []),
                "sec": item.get('sec', '점심'),
            }
            js_items.append(js_item)

        menu_data[ds] = js_items

    # 주간 배열
    first_day = dt(year, month, 1)
    days_since_monday = first_day.weekday()
    week_start = first_day - timedelta(days=days_since_monday)
    weeks = []
    d = week_start
    for _ in range(6):
        wk = []
        for _ in range(7):
            if d.weekday() != 6:
                wk.append(d.isoformat() if d.month == month else None)
            d += timedelta(days=1)
        if any(wk):
            weeks.append(wk)

    return {
        "year": year,
        "month": month,
        "menus": menu_data,
        "allergens": {str(k): v for k, v in ALLERGEN_NAMES.items()},
        "weeks": weeks,
    }


# ── CLI ──
if __name__ == '__main__':
    import json
    import sys

    if len(sys.argv) < 2:
        print("Usage: python3 parse_pdf.py <식단표.pdf> [output.json]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    result = parse_pdf_menu(pdf_path)
    data = build_json_from_pdf(result)

    if len(sys.argv) >= 3:
        out = sys.argv[2]
    else:
        out = Path(pdf_path).stem + '.json'

    Path(out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"\n[OK] → {out}")
    print(f"  {data['year']}년 {data['month']}월, {len(data['menus'])}일, "
          f"{sum(len(v) for v in data['menus'].values())}개 아이템")
