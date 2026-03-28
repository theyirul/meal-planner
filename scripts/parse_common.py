#!/usr/bin/env python3
"""
매일아침 — 공유 파싱 유틸리티

모든 포맷 파서가 공유하는 상수, 알레르기 추출, JSON 빌더.
"""
from __future__ import annotations

import re
from datetime import date, timedelta

# ── 19종 알레르기 ──
ALLERGEN_NAMES = {
    1: "난류(계란)", 2: "우유", 3: "메밀", 4: "땅콩", 5: "대두",
    6: "밀", 7: "고등어", 8: "게", 9: "새우", 10: "돼지고기",
    11: "복숭아", 12: "토마토", 13: "아황산류", 14: "호두", 15: "닭고기",
    16: "쇠고기", 17: "오징어", 18: "조개류", 19: "잣"
}

# 원형숫자 → 일반숫자 매핑
CIRCLED_NUMS = {
    '①': 1, '②': 2, '③': 3, '④': 4, '⑤': 5,
    '⑥': 6, '⑦': 7, '⑧': 8, '⑨': 9, '⑩': 10,
    '⑪': 11, '⑫': 12, '⑬': 13, '⑭': 14, '⑮': 15,
    '⑯': 16, '⑰': 17, '⑱': 18, '⑲': 19,
}
CIRCLED_PATTERN = re.compile('[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]+')

# 장식 문자 제거
DECO_CHARS = re.compile(r'[♬★◉♥♣✔\U000f0074]')


def extract_allergy_circled(text: str) -> tuple[str, list[int]]:
    """원형숫자 알레르기 추출. "쇠고기장조림⑤⑥⑯" → ("쇠고기장조림", [5, 6, 16])"""
    nums = []
    # 원문자 번호 수집 후 한 번에 제거 (인덱스 밀림 방지)
    for match in CIRCLED_PATTERN.finditer(text):
        for ch in match.group():
            if ch in CIRCLED_NUMS:
                nums.append(CIRCLED_NUMS[ch])
    clean = CIRCLED_PATTERN.sub('', text)
    # 괄호형도 처리 (혼용 대비)
    paren_pat = re.compile(r'[\(（]([\d,.\s]+)[\)）]')
    for m in paren_pat.finditer(clean):
        for n in re.findall(r'\d+', m.group(1)):
            v = int(n)
            if 1 <= v <= 19 and v not in nums:
                nums.append(v)
        clean = clean[:m.start()] + clean[m.end():]
    return clean.strip(), sorted(set(nums))


def extract_allergy_inline(text: str) -> tuple[str, list[int]]:
    """인라인 콤마형 알레르기 추출. "조갯살호박국5,6,18" → ("조갯살호박국", [5, 6, 18])"""
    if not text or not text.strip():
        return ('', [])
    text = text.strip()
    nums = []

    # 원형숫자도 체크 (혼용 대비)
    for ch, n in CIRCLED_NUMS.items():
        if ch in text:
            nums.append(n)
            text = text.replace(ch, '')

    # 괄호 안 대체메뉴 제거: "백김치9(백깍두기9)"
    alt_pattern = re.compile(r'\(([^)]*\d+[^)]*)\)')
    for m in reversed(list(alt_pattern.finditer(text))):
        inner = m.group(1)
        if re.match(r'만\s*\d', inner):
            continue
        text = text[:m.start()] + text[m.end():]

    # 끝에 붙은 숫자+콤마: "조갯살호박국5,6,18"
    m = re.search(r'[\s,]*(\d{1,2}(?:\s*,\s*\d{1,2})*)\s*$', text)
    if m:
        for n_str in re.findall(r'\d{1,2}', m.group(1)):
            v = int(n_str)
            if 1 <= v <= 19 and v not in nums:
                nums.append(v)
        text = text[:m.start()].strip()

    return text, sorted(set(nums))


def clean_menu_name(name: str) -> str:
    """메뉴명 정리: 장식 문자, 연령 표기, 앞뒤 괄호 제거"""
    name = DECO_CHARS.sub('', name).strip()
    name = re.sub(r'\(만\s*\d+-?\d*세\)', '', name).strip()
    # 앞뒤가 괄호로 감싸져 있으면 제거: "(견과류(부럼))" → "견과류(부럼)"
    if name.startswith('(') and name.endswith(')'):
        inner = name[1:-1]
        if inner.count('(') == inner.count(')'):
            name = inner.strip()
    # '+' 접두어 제거 (용인시 음료 동반 표기: "+둥굴레차" → "둥굴레차")
    name = re.sub(r'^\+\s*', '', name)
    # 열린 괄호가 닫히지 않은 경우 제거: "떡만두국(비" → "떡만두국"
    # (pdfplumber 셀 분할로 잘린 태그)
    if name.count('(') > name.count(')'):
        name = re.sub(r'\([^)]*$', '', name).strip()
    return name


def build_weeks(year: int, month: int) -> list[list]:
    """월간 주 배열 생성 (일요일 제외)"""
    first_day = date(year, month, 1)
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
    return weeks


def build_output_json(year: int, month: int, menus: dict, recipes: dict | None = None) -> dict:
    """표준 JSON 출력 포맷 생성"""
    recipes = recipes or {}
    matched = 0
    total = 0
    menu_data = {}

    for ds, items in sorted(menus.items()):
        js_items = []
        for item in items:
            name = item['name']
            recipe = match_recipe(name, recipes)
            if recipe:
                matched += 1
            total += 1
            js_items.append({
                "name": name,
                "recipe": recipe,
                "allergy": item.get('allergy', []),
                "sec": item.get('sec', '점심'),
            })
        menu_data[ds] = js_items

    if total > 0 and recipes:
        print(f"  [레시피 매칭] {matched}/{total} ({matched * 100 // total}%)")

    return {
        "year": year,
        "month": month,
        "menus": menu_data,
        "allergens": {str(k): v for k, v in ALLERGEN_NAMES.items()},
        "weeks": build_weeks(year, month),
    }


def match_recipe(menu_name: str, recipes: dict) -> str:
    """메뉴명으로 레시피 매칭 (정확 → 공백무시 → 부분)"""
    if menu_name in recipes:
        return recipes[menu_name]
    clean = menu_name.replace(' ', '')
    for rname, rtext in recipes.items():
        if rname.replace(' ', '') == clean:
            return rtext
    for rname, rtext in recipes.items():
        if rname in menu_name or menu_name in rname:
            return rtext
    return ''
