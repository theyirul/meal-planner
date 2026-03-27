#!/usr/bin/env python3
"""
어린이집 점심 메뉴 인터랙티브 HTML 생성기

레시피 엑셀(.xlsx)과 알레르기 식단표 PDF로부터
월간/주간 탭, 알레르기 필터, 양념 제외 분석이 포함된
인터랙티브 HTML 페이지를 자동 생성한다.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pdfplumber


# ── 19종 알레르기 ──
ALLERGEN_NAMES = {
    1: "난류(계란)", 2: "우유", 3: "메밀", 4: "땅콩", 5: "대두",
    6: "밀", 7: "고등어", 8: "게", 9: "새우", 10: "돼지고기",
    11: "복숭아", 12: "토마토", 13: "아황산류", 14: "호두", 15: "닭고기",
    16: "쇠고기", 17: "오징어", 18: "조개류", 19: "잣"
}

# ── 양념 제외 분석용 키워드 ──
SOY_INTEGRAL_KW = ['두부', '유부', '된장', '청국장', '콩가루', '콩나물']
WHEAT_INTEGRAL_KW = ['밀가루', '부침가루', '빵가루', '당면', '어묵', '카레',
                     '짜장', '국수', '수제비', '만두', '크래미', '게살']
SAUCE_KW = ['간장', '국간장', '진간장', '양조간장', '양념장']

DAY_NAMES_KR = ['월', '화', '수', '목', '금', '토', '일']
DAY_NAMES_CAL = ['일', '월', '화', '수', '목', '금', '토']


# ═══════════════════════════════════════
# 1. 엑셀에서 메뉴 데이터 추출
# ═══════════════════════════════════════

def find_recipe_sheet(xlsx_path: str) -> str:
    """'일반형 레시피' 패턴의 시트를 찾는다."""
    xls = pd.ExcelFile(xlsx_path)
    skip_kw = ['인쇄', '출력용', '출력']
    for name in xls.sheet_names:
        if '일반형' in name and '레시피' in name and not any(k in name for k in skip_kw):
            return name
    # 못 찾으면 '레시피'가 포함된 첫 시트 (출력용 제외)
    for name in xls.sheet_names:
        if '레시피' in name and not any(k in name for k in skip_kw):
            return name
    raise ValueError(f"레시피 시트를 찾을 수 없습니다. 시트 목록: {xls.sheet_names}")


def extract_menus(xlsx_path: str) -> dict:
    """엑셀에서 일별 점심 메뉴와 레시피를 추출한다."""
    sheet_name = find_recipe_sheet(xlsx_path)
    print(f"[INFO] 시트 사용: '{sheet_name}'")
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)

    menus = {}  # {date_str: [{"name": ..., "recipe": ..., "section": ...}]}
    current_date = None
    current_section = None
    SECTIONS = ('오전간식', '점심', '오후간식')

    for _, row in df.iterrows():
        val0 = row.iloc[0] if len(row) > 0 else None

        # 날짜 감지 (col 0) — 날짜 변경 시 섹션 리셋
        if pd.notna(val0):
            ds = _parse_date(val0)
            if ds:
                current_date = ds
                current_section = None

        # 구분 감지 (col 0 — 오전간식/점심/오후간식)
        if pd.notna(val0) and isinstance(val0, str):
            v = val0.strip()
            if v in SECTIONS:
                current_section = v

        if current_section not in SECTIONS or not current_date:
            continue

        # 메뉴명 (컬럼 1) — 헤더 행 라벨 제외
        menu_name = row.iloc[1] if len(row) > 1 else None
        if pd.isna(menu_name) or not str(menu_name).strip():
            continue
        menu_name = str(menu_name).strip()
        if menu_name in ('메뉴명', '식재료명', '구분'):
            continue

        # 조리방법 (컬럼 10)
        recipe = ''
        if len(row) > 10 and pd.notna(row.iloc[10]):
            recipe = str(row.iloc[10]).strip()

        if current_date not in menus:
            menus[current_date] = []
        menus[current_date].append({"name": menu_name, "recipe": recipe, "section": current_section})

    # 후처리: 점심 섹션이 없는 날의 보정
    # 엑셀에서 점심 라벨이 누락된 경우 (예: 토요일 데이터 입력 오류)
    # 오전간식과 오후간식 사이의 메인 요리를 점심으로 재분류
    SNACK_KW = ['우유', '두유', '요구르트', '주스', '토마토', '바나나', '사과',
                '배', '포도', '오렌지', '귤', '딸기', '치즈', '과일', '비스킷',
                '과자', '쿠키', '마들렌', '달걀', '삶은']
    for ds, items in menus.items():
        sections = {it['section'] for it in items}
        if '점심' not in sections and '오전간식' in sections:
            # 오전간식 첫 항목(간식류)만 유지, 나머지를 점심으로 재분류
            snack_done = False
            for it in items:
                if it['section'] != '오전간식':
                    continue
                if not snack_done and any(kw in it['name'] for kw in SNACK_KW):
                    snack_done = True  # 첫 간식 항목은 유지
                elif snack_done or not any(kw in it['name'] for kw in SNACK_KW):
                    it['section'] = '점심'
                    snack_done = True

    print(f"[INFO] {len(menus)}일치 메뉴 추출 완료")
    return menus


def _parse_date(val) -> str | None:
    """다양한 날짜 형식을 'YYYY-MM-DD' 문자열로 변환."""
    if isinstance(val, (pd.Timestamp, date)):
        return str(val)[:10]
    s = str(val).strip()
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return None


# ═══════════════════════════════════════
# 2. PDF에서 알레르기 정보 추출
# ═══════════════════════════════════════

def extract_allergies(pdf_path: str, menu_dates: list[str]) -> dict:
    """PDF 테이블에서 날짜별 메뉴 알레르기 정보를 추출한다.

    PDF 구조: 첫 페이지에 큰 테이블.
    - 행 0: 헤더(날짜) → '3(화)', '4(수)' 등
    - 행 2: 점심 메뉴 셀 → 줄바꿈으로 구분된 메뉴, 각 메뉴 뒤에 (5,6,10) 형태 알레르기
    """
    with pdfplumber.open(pdf_path) as pdf:
        table = pdf.pages[0].extract_tables()[0]

    # menu_dates에서 연/월 감지
    if menu_dates:
        ref = date.fromisoformat(menu_dates[0])
        year, month = ref.year, ref.month
    else:
        year, month = 2026, 3

    # 테이블에서 헤더 행(일자)과 모든 섹션(오전간식/점심/오후간식)을 추출
    # 구조: [헤더행, 오전간식행, 점심행, 오후간식행, 열량행] × 주 수
    all_allergy = {}
    allergy_pattern = re.compile(r'[\(（]([\d,.\s]+)[\)）]')

    current_date_cols = {}  # {col_idx: day_number}
    # 추출 대상 섹션 키워드
    SECTION_KW = ('오전', '점심', '오후')
    BOUNDARY_KW = ('일자', '열량')

    for ri, row in enumerate(table):
        cell0 = str(row[0]).strip() if row[0] else ''

        # 헤더 행 감지: '일자' 또는 날짜 패턴
        if '일자' in cell0 or re.match(r'\d+\s*[\(（]', cell0):
            current_date_cols = {}
            for ci, cell in enumerate(row):
                if not cell:
                    continue
                m = re.match(r'(\d+)\s*[\(（]', str(cell))
                if m:
                    current_date_cols[ci] = int(m.group(1))
            continue

        # 오전간식 / 점심 / 오후간식 행 감지
        is_section = any(kw in cell0 for kw in SECTION_KW)
        if not is_section:
            continue
        if not current_date_cols:
            continue

        # 섹션 행 + 후속 연속 행들도 수집 (PDF에서 메뉴가 여러 행에 걸쳐 있는 경우)
        section_rows = [row]
        for ri2 in range(ri + 1, min(ri + 8, len(table))):
            next_row = table[ri2]
            next_c0 = str(next_row[0]).strip() if next_row[0] else ''
            # 다른 섹션이나 경계 키워드를 만나면 중단
            if any(kw in next_c0 for kw in SECTION_KW) or any(kw in next_c0 for kw in BOUNDARY_KW):
                break
            # 데이터가 있는 후속 행이면 수집
            has_data = any(next_row[ci] for ci in current_date_cols if ci < len(next_row))
            if has_data:
                section_rows.append(next_row)

        for ci, day_num in current_date_cols.items():
            try:
                ds = date(year, month, day_num).isoformat()
            except ValueError:
                continue

            # 모든 섹션 행에서 해당 컬럼의 텍스트를 합산
            all_lines = []
            for lr in section_rows:
                cell = lr[ci] if ci < len(lr) else None
                if cell:
                    all_lines.extend(str(cell).split('\n'))

            if not all_lines:
                continue

            if ds not in all_allergy:
                all_allergy[ds] = {}
            merged = _merge_menu_lines(all_lines)

            # '/' 로 구분된 복합 메뉴 분리
            # 예: '치즈(2)/ 배' → ['치즈(2)', '배']
            expanded = []
            for menu_text in merged:
                if '/' in menu_text:
                    parts = [p.strip() for p in menu_text.split('/') if p.strip()]
                    expanded.extend(parts)
                else:
                    expanded.append(menu_text)

            for menu_text in expanded:
                m = allergy_pattern.search(menu_text)
                if m:
                    name = allergy_pattern.sub('', menu_text).strip()
                    nums = [int(n) for n in re.findall(r'\d+', m.group(1)) if 1 <= int(n) <= 19]
                    if name and nums:
                        all_allergy[ds][name] = nums
                else:
                    name = menu_text.strip()
                    if name:
                        all_allergy[ds][name] = []

    total = sum(len(v) for v in all_allergy.values())
    print(f"[INFO] PDF에서 {len(all_allergy)}일, {total}개 메뉴의 알레르기 정보 추출")

    if not all_allergy:
        print("[WARN] 테이블 추출 실패, 텍스트 기반으로 대체합니다.")
        return _fallback_text_extraction(pdf_path, menu_dates)

    return all_allergy


def _merge_menu_lines(lines: list[str]) -> list[str]:
    """줄바꿈으로 분리된 메뉴명과 알레르기 번호를 병합.
    예: ['닭살카레라이스', '(2,5,6,10,15,16,18)'] → ['닭살카레라이스(2,5,6,10,15,16,18)']
    주의: '(백)배추김치'처럼 (한글)로 시작하는 메뉴명은 새 항목으로 분리.
    """
    merged = []
    buf = ''
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 알레르기 번호 연속행: (숫자... 로 시작
        if re.match(r'[\(（]\d', line):
            buf += line
        else:
            if buf:
                merged.append(buf)
            buf = line
    if buf:
        merged.append(buf)
    return merged


def _fallback_text_extraction(pdf_path: str, menu_dates: list[str]) -> dict:
    """테이블 추출 실패 시 텍스트 기반으로 추출."""
    text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ''

    pattern = re.compile(r'([가-힣]+(?:[가-힣a-zA-Z0-9\-&]+)*)\s*[\(（](\d[\d,.\s]*?)[\)）]')
    raw = {}
    for name_raw, nums_raw in pattern.findall(text):
        name = name_raw.strip()
        nums = [int(n) for n in re.findall(r'\d+', nums_raw) if 1 <= int(n) <= 19]
        if nums and name:
            raw[name] = nums

    return {"_fallback": True, "_raw": raw}


def match_allergies(menus: dict, allergy_data: dict) -> dict:
    """메뉴명과 PDF 알레르기 데이터를 매칭한다."""
    # 테이블 기반 추출인 경우: 날짜별로 이미 매칭됨
    if '_fallback' not in allergy_data:
        result = {}
        for ds, items in menus.items():
            result[ds] = {}
            pdf_menus = allergy_data.get(ds, {})
            for item in items:
                name = item['name'].split('\n')[0].strip()
                matched = _find_allergy_in_day(name, pdf_menus)
                if matched is not None:
                    result[ds][name] = matched
        return result

    # 폴백: 전역 매칭
    raw = allergy_data['_raw']
    result = {}
    for ds, items in menus.items():
        result[ds] = {}
        for item in items:
            name = item['name'].split('\n')[0].strip()
            matched = _find_allergy_match(name, raw)
            if matched:
                result[ds][name] = matched
    return result


def _find_allergy_in_day(menu_name: str, pdf_menus: dict) -> list[int] | None:
    """같은 날짜의 PDF 메뉴에서 알레르기 번호를 찾는다."""
    clean = menu_name.replace('(백)', '').replace('*', '').strip()

    # 정확히 일치
    for k, v in pdf_menus.items():
        k_clean = k.replace('(백)', '').replace('*', '').strip()
        if k_clean == clean:
            return v

    # 부분 매칭
    for k, v in sorted(pdf_menus.items(), key=lambda x: len(x[0]), reverse=True):
        k_clean = k.replace('(백)', '').replace('*', '').strip()
        if k_clean in clean or clean in k_clean:
            return v
        if len(k_clean) >= 3 and len(clean) >= 3 and k_clean[:3] == clean[:3]:
            return v

    return None


def _find_allergy_match(menu_name: str, raw: dict) -> list[int] | None:
    """전역 메뉴명으로 알레르기 번호를 찾는다 (폴백용)."""
    clean = menu_name.replace('(백)', '').replace('*', '').strip()
    if clean in raw:
        return raw[clean]
    for k in sorted(raw.keys(), key=len, reverse=True):
        k_clean = k.replace('(백)', '').replace('*', '').strip()
        if k_clean in clean or clean in k_clean:
            return raw[k]
    return None


# ═══════════════════════════════════════
# 3. 양념 제외 분석
# ═══════════════════════════════════════

def get_main_recipe(recipe: str) -> tuple[str, str]:
    """레시피를 본 조리 과정과 팁 라인으로 분리.
    팁 라인(* 또는 ※로 시작)은 선택적 제안이므로 분석에서 제외."""
    main_lines, tip_lines = [], []
    for line in recipe.split('\n'):
        stripped = line.strip()
        if stripped.startswith('*') or stripped.startswith('※'):
            tip_lines.append(stripped)
        else:
            main_lines.append(stripped)
    return '\n'.join(main_lines), '\n'.join(tip_lines)


def analyze_sauce(menus: dict, allergy_match: dict) -> dict:
    """양념 제외 가능 여부를 알레르기별로 분석한다."""
    result = {}

    for ds, items in menus.items():
        day_result = []
        match_data = allergy_match.get(ds, {})

        for item in items:
            name = item['name'].split('\n')[0].strip()
            allergens = match_data.get(name, [])
            recipe = item.get('recipe', '')

            if not recipe or (5 not in allergens and 6 not in allergens):
                continue

            # 간장류 사용 여부 확인
            if not any(kw in recipe for kw in SAUCE_KW):
                continue

            main_recipe, _ = get_main_recipe(recipe)

            removable = []
            integral = []
            sauce_step = ''

            # 간장 포함 단계 찾기
            for line in recipe.split('\n'):
                if any(kw in line for kw in SAUCE_KW):
                    sauce_step = line.strip()
                    break

            # 대두(5) 분석
            if 5 in allergens:
                if any(kw in main_recipe for kw in SOY_INTEGRAL_KW):
                    integral.append(5)
                else:
                    removable.append(5)

            # 밀(6) 분석
            if 6 in allergens:
                if any(kw in main_recipe for kw in WHEAT_INTEGRAL_KW):
                    integral.append(6)
                else:
                    removable.append(6)

            if removable:
                day_result.append({
                    "name": name,
                    "removable": removable,
                    "integral": integral,
                    "sauceStep": sauce_step
                })

        if day_result:
            result[ds] = day_result

    sauce_count = sum(len(v) for v in result.values())
    print(f"[INFO] {sauce_count}개 메뉴에서 양념 제외 가능 분석 완료")
    return result


# ═══════════════════════════════════════
# 4. HTML 생성
# ═══════════════════════════════════════

def detect_month_year(menus: dict) -> tuple[int, int]:
    """메뉴 데이터에서 연/월을 감지한다."""
    first_date = sorted(menus.keys())[0]
    d = date.fromisoformat(first_date)
    return d.year, d.month


def build_calendar_html(menus: dict, year: int, month: int) -> str:
    """월간 달력 HTML을 생성한다 (월~토 6열, 일요일 제외)."""
    first_day = date(year, month, 1)
    # 월요일 시작: weekday()가 0=월이므로 그대로 사용
    # 단, 첫 날이 일요일(6)이면 빈 셀 0개
    if first_day.weekday() == 6:
        empty_cells = 0
    else:
        empty_cells = first_day.weekday()

    # 해당 월의 마지막 날
    if month == 12:
        last_day = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(year, month + 1, 1) - timedelta(days=1)
    total_days = last_day.day

    html = '<div class="cal">\n'
    html += '<div class="ch">월</div><div class="ch">화</div><div class="ch">수</div>'
    html += '<div class="ch">목</div><div class="ch">금</div>'
    html += '<div class="ch sa">토</div>\n'

    # 빈 셀
    for _ in range(empty_cells):
        html += '<div class="cd em"></div>\n'

    for day in range(1, total_days + 1):
        d = date(year, month, day)
        dow = d.weekday()  # 0=월 ~ 6=일
        if dow == 6:  # 일요일은 건너뛴다
            continue
        ds = d.isoformat()

        is_sat = dow == 5
        cls_extra = ' sa' if is_sat else ''

        items = menus.get(ds, [])
        if not items:
            html += f'<div class="cd em{cls_extra}"><div class="dn">{day}</div></div>\n'
        else:
            # 전체 메뉴 — 섹션별로 모두 표시, 구분선만
            sec_order = ['오전간식', '점심', '오후간식']
            sec_map = {}
            for it in items:
                s = it.get('section', '점심')
                if s not in sec_map:
                    sec_map[s] = []
                n = it['name'].split('\n')[0]
                sec_map[s].append(n)
            parts = []
            for sec in sec_order:
                names_s = sec_map.get(sec, [])
                if names_s:
                    parts.append('<br>'.join(_esc(n) for n in names_s))
            preview_text = '<span class="mp-div"></span>'.join(parts)
            html += (f'<div class="cd{cls_extra}" data-d="{ds}" onclick="mc(\'{ds}\')">'
                     f'<div class="dn">{day}</div>'
                     f'<div class="dt"><span class="ad"></span><span class="sd"></span></div>'
                     f'<div class="mp">{preview_text}</div></div>\n')

    # 마지막 주 빈 셀 채우기 (6열 기준)
    last_shown = last_day
    if last_shown.weekday() == 6:  # 마지막 날이 일요일이면 토요일이 마지막
        last_shown = last_shown - timedelta(days=1)
    remaining = 5 - last_shown.weekday()  # 토요일(5)이면 0, 금요일(4)이면 1 ...
    for _ in range(remaining):
        html += '<div class="cd em"></div>\n'

    html += '</div>'
    return html


def build_js_data(menus: dict, allergy_match: dict, sauce_data: dict) -> str:
    """JavaScript 데이터 객체를 생성한다."""
    js_data = {}

    for ds, items in sorted(menus.items()):
        match_data = allergy_match.get(ds, {})
        sauce_items = {s['name']: s for s in sauce_data.get(ds, [])}
        js_items = []

        for item in items:
            name = item['name'].split('\n')[0].strip()
            allergy_nums = match_data.get(name, [])
            sauce = sauce_items.get(name, {})

            js_item = {
                "name": item['name'],
                "recipe": item.get('recipe', ''),
                "allergy": allergy_nums,
                "sec": item.get('section', '점심'),
            }
            if sauce.get('removable'):
                js_item['rm'] = sauce['removable']
            if sauce.get('integral'):
                js_item['ig'] = sauce['integral']
            if sauce.get('sauceStep'):
                js_item['ss'] = sauce['sauceStep']

            js_items.append(js_item)

        js_data[ds] = js_items

    return f"const D={json.dumps(js_data, ensure_ascii=False)};"


def build_weeks_js(year: int, month: int) -> str:
    """주간 배열 JS 코드를 생성한다 (월~토 6일, 일요일 제외)."""
    first_day = date(year, month, 1)
    # 월요일 시작 주간 → 첫 날이 속한 주의 월요일을 찾는다
    days_since_monday = first_day.weekday()  # 0=월 이면 0, 6=일 이면 6
    week_start = first_day - timedelta(days=days_since_monday)

    code = "const wks=[];\n(function(){\n"
    code += f"  let d=new Date('{week_start.isoformat()}T00:00:00');\n"
    code += "  for(let w=0;w<6;w++){let wk=[];for(let i=0;i<7;i++){"
    code += "const s=d.toISOString().slice(0,10);"
    code += f"if(d.getDay()!==0)wk.push(d.getMonth()==={month-1}?s:null);"
    code += "d.setDate(d.getDate()+1)}if(wk.some(x=>x))wks.push(wk)}\n"
    code += "})();"
    return code


def generate_html(menus: dict, allergy_match: dict, sauce_data: dict, year: int, month: int) -> str:
    """최종 인터랙티브 HTML을 생성한다."""
    calendar_html = build_calendar_html(menus, year, month)
    js_data = build_js_data(menus, allergy_match, sauce_data)
    weeks_js = build_weeks_js(year, month)
    an_js = f"const AN={json.dumps(ALLERGEN_NAMES, ensure_ascii=False)};"

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{month}월 어린이집 식단</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Pretendard+Variable:wght@300;400;500;600;700&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
:root{{
  --accent:#1F2937;--accent-light:#F3F4F6;--accent-border:#D1D5DB;
  --alert:#DC2626;--alert-bg:#FEF2F2;--alert-light:#FECACA;
  --ok:#16A34A;--ok-bg:#F0FDF4;--ok-border:#BBF7D0;
  --bg:#FFFFFF;--card:#FFFFFF;
  --text:#111827;--text2:#6B7280;--text3:#9CA3AF;
  --border:#E5E7EB;--border-light:#F3F4F6;
  --radius:10px;
}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Pretendard Variable','Noto Sans KR',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);padding:28px 20px;max-width:860px;margin:0 auto;font-size:15px;line-height:1.6;-webkit-font-smoothing:antialiased}}
h1{{text-align:center;font-size:1.5rem;font-weight:700;color:var(--text);margin:16px 0 24px;letter-spacing:-.02em}}

/* ── Tabs ── */
.tabs{{display:flex;gap:4px;margin-bottom:28px;background:var(--border-light);border-radius:var(--radius);padding:3px}}
.tb{{flex:1;padding:11px 0;text-align:center;font-size:.9rem;font-weight:600;color:var(--text3);background:none;border:none;cursor:pointer;font-family:inherit;transition:all .2s;border-radius:8px}}
.tb.on{{color:var(--text);background:var(--card);box-shadow:0 1px 3px rgba(0,0,0,.08)}}
.tc{{display:none}}.tc.on{{display:block}}

/* (요일 필터 제거됨 — 일요일은 달력에서 완전 제외) */

/* ── Search ── */
.srch{{margin-bottom:20px;display:flex;align-items:center;gap:8px}}
.srch input{{flex:1;border:1.5px solid var(--border);border-radius:var(--radius);padding:10px 14px;font-size:.9rem;font-family:inherit;outline:none;transition:border-color .2s;background:var(--card)}}
.srch input:focus{{border-color:var(--accent)}}
.srch input::placeholder{{color:var(--text3)}}
.srch .sr-cl{{background:none;border:none;font-size:1.1rem;color:var(--text3);cursor:pointer;padding:4px 8px;display:none}}
.srch .sr-cl.sh{{display:block}}
.cd.sr-hl{{border-color:var(--text)!important;background:var(--accent-light)!important;opacity:1!important}}
.cd.sr-dim{{opacity:.2!important}}
.sr-cnt{{text-align:center;font-size:.82rem;color:var(--text2);margin:-12px 0 18px}}

/* ── Filter (collapsible) ── */
.fs{{margin-bottom:8px}}
.fs-toggle{{display:flex;align-items:center;gap:6px;padding:10px 0;cursor:pointer;-webkit-tap-highlight-color:transparent}}
.fs-toggle .fs-icon{{width:20px;height:20px;border-radius:6px;background:var(--border-light);display:flex;align-items:center;justify-content:center;font-size:.7rem;color:var(--text2);transition:all .2s}}
.fs-toggle.on .fs-icon{{background:var(--alert);color:#fff}}
.fs-toggle .fs-label{{font-weight:500;font-size:.88rem;color:var(--text2)}}
.fs-toggle .fs-cnt{{font-size:.78rem;color:var(--alert);font-weight:600}}
.fs-toggle .fs-arr{{margin-left:auto;color:var(--text3);font-size:.7rem;transition:transform .2s}}
.fs-toggle.open .fs-arr{{transform:rotate(180deg)}}
.fg-wrap{{max-height:0;overflow:hidden;transition:max-height .3s ease}}
.fg-wrap.open{{max-height:200px}}
.fg{{display:flex;flex-wrap:wrap;gap:6px;padding:8px 0 12px}}
.fc{{font-size:.78rem;padding:5px 10px;border-radius:20px;background:var(--border-light);color:var(--text2);cursor:pointer;transition:all .15s;border:1.5px solid transparent;user-select:none;font-weight:500}}
.fc:hover{{background:var(--alert-bg);border-color:var(--alert-light)}}
.fc.on{{background:var(--alert);color:#fff;border-color:var(--alert)}}
.fc .n{{font-weight:700}}

/* Active filter bar */
.fb{{display:none;background:var(--alert-bg);border:1px solid var(--alert-light);border-radius:var(--radius);padding:10px 14px;margin-bottom:20px;font-size:.84rem;color:var(--alert);align-items:center;gap:8px;flex-wrap:wrap}}
.fb.sh{{display:flex}}
.fb .ts{{display:flex;flex-wrap:wrap;gap:4px;flex:1}}
.ft{{font-size:.76rem;padding:3px 8px;border-radius:6px;background:var(--alert-light);color:var(--alert);font-weight:600}}
.fb .cl{{margin-left:auto;background:var(--alert);color:#fff;border:none;padding:5px 14px;border-radius:8px;cursor:pointer;font-size:.8rem;font-family:inherit;white-space:nowrap;font-weight:500}}

/* ── MONTHLY CALENDAR ── */
.cal{{display:grid;grid-template-columns:repeat(6,1fr);gap:5px;margin-bottom:28px}}
.ch{{text-align:center;font-weight:600;font-size:.78rem;padding:10px 2px 6px;color:var(--text3)}}
.ch.sa{{color:#3B82F6}}
.cd{{background:var(--card);border-radius:var(--radius);padding:8px 7px;min-height:0;cursor:pointer;transition:all .15s;border:1.5px solid var(--border-light);position:relative}}
.cd:hover{{border-color:var(--accent-border);box-shadow:0 2px 8px rgba(0,0,0,.06)}}
.cd.ac{{border-color:var(--accent);background:var(--accent-light)}}
.cd.em{{background:transparent;cursor:default;border-color:transparent;opacity:.3}}.cd.em:hover{{box-shadow:none;border-color:transparent}}
/* .dx 클래스 제거됨 */
.cd .dn{{font-weight:700;font-size:.88rem;margin-bottom:4px;color:var(--text)}}
.cd.sa .dn{{color:#3B82F6}}
.cd .dt{{position:absolute;top:7px;right:7px;display:flex;gap:3px}}
.cd .ad{{width:7px;height:7px;border-radius:50%;background:var(--alert);display:none}}
.cd .sd{{width:7px;height:7px;border-radius:50%;background:var(--ok);display:none}}
.cd .mp{{font-size:.7rem;color:var(--text2);line-height:1.45}}
.cd .mp-div{{display:block;height:1px;background:var(--border-light);margin:3px 0}}
.cd .mp .ma{{color:var(--alert);font-weight:600}}
.cd .mp .ms{{color:var(--ok);font-weight:600}}

/* ── Allergy & sauce tags (inline) ── */
.at{{display:none;flex-wrap:wrap;gap:3px;align-items:center}}
.at.sh{{display:inline-flex}}
.ag{{font-size:.65rem;padding:1px 5px;border-radius:3px;background:var(--alert-bg);color:var(--alert);font-weight:600;line-height:1.4}}
.ag.hl{{background:var(--alert);color:#fff}}
.sb{{display:none;font-size:.68rem;padding:2px 7px;border-radius:4px;font-weight:600;white-space:nowrap;line-height:1.4}}
.sb.sh{{display:inline}}
.sb.f{{background:var(--ok-bg);color:var(--ok)}}
.sb.p{{background:#FFFBEB;color:#D97706}}

/* ── WEEKLY ── */
.wn{{display:flex;align-items:center;justify-content:center;gap:14px;margin-bottom:20px}}
.wn button{{background:var(--card);border:1.5px solid var(--border);border-radius:var(--radius);padding:9px 20px;cursor:pointer;font-family:inherit;font-size:.88rem;color:var(--text2);transition:all .15s;-webkit-tap-highlight-color:transparent;font-weight:500}}
.wn button:hover,.wn button:active{{border-color:var(--accent);color:var(--accent)}}
.wn .wl{{font-weight:600;font-size:1.05rem;color:var(--text)}}
.wr{{display:grid;gap:6px;margin-bottom:20px}}
.wc{{background:var(--card);border-radius:var(--radius);padding:10px 6px;cursor:pointer;border:1.5px solid var(--border-light);transition:all .15s;text-align:center;-webkit-tap-highlight-color:transparent;position:relative;opacity:.5}}
.wc:hover{{border-color:var(--accent-border);box-shadow:0 2px 8px rgba(0,0,0,.06);opacity:.8}}
.wc.ac{{border-color:var(--text);background:var(--accent-light);opacity:1}}
.wc .wd{{font-size:.72rem;color:var(--text3);margin-bottom:1px}}
.wc .wdt{{font-weight:700;font-size:1.1rem;color:var(--text)}}
.wc .wdots{{position:absolute;top:5px;right:5px;display:flex;gap:2px}}
.wc .wdots .wd-a,.wc .wdots .wd-s{{width:6px;height:6px;border-radius:50%;display:none}}
.wc .wdots .wd-a{{background:var(--alert)}}
.wc .wdots .wd-s{{background:var(--ok)}}
.wc.em{{opacity:.2;cursor:default}}.wc.em:hover{{box-shadow:none;border-color:var(--border-light);opacity:.2}}

/* Weekly detail panel */
.wp{{background:var(--card);border-radius:14px;padding:28px;box-shadow:0 1px 8px rgba(0,0,0,.04);border:1px solid var(--border-light);display:none;position:relative}}
.wp.sh{{display:block}}
.wp-top{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:6px}}
.wp .wpt{{font-size:1.15rem;font-weight:700;color:var(--text)}}
.cp-btn{{background:var(--border-light);border:1px solid var(--border);border-radius:8px;padding:6px 14px;font-size:.78rem;font-family:inherit;color:var(--text2);cursor:pointer;transition:all .15s;font-weight:500;white-space:nowrap;flex-shrink:0}}
.cp-btn:hover{{background:var(--border);color:var(--text)}}
.cp-btn.ok{{background:var(--ok-bg);border-color:var(--ok-border);color:var(--ok)}}
.wl2{{list-style:none}}
.sec-hd{{font-size:.76rem;font-weight:600;color:var(--text2);padding:16px 0 2px;letter-spacing:.02em;border-bottom:none}}
.sec-hd:first-child{{padding-top:4px}}
.wi{{border-bottom:1px solid var(--border-light);display:flex;align-items:flex-start;gap:10px;flex-wrap:wrap}}.wi:last-child{{border-bottom:none}}
.wh{{padding:12px 0;cursor:pointer;display:flex;align-items:center;flex-wrap:wrap;gap:6px;font-size:.95rem;font-weight:500;color:var(--text);transition:color .15s;-webkit-tap-highlight-color:transparent;flex:1;min-width:0}}
.wh-name{{flex-shrink:1;min-width:0}}
.wh:hover,.wh:active{{color:var(--text2)}}
.wh .ar{{transition:transform .3s;font-size:.72rem;color:var(--text3);margin-left:auto}}
.wh.op .ar{{transform:rotate(180deg);color:var(--text)}}

/* Recipe expand */
.rc{{max-height:0;overflow:hidden;transition:max-height .4s ease;font-size:.88rem;line-height:1.8;color:var(--text2);background:var(--bg);border-radius:var(--radius);width:100%}}
.rc.op{{max-height:800px;padding:16px;margin-bottom:8px}}
.rc p{{margin:0}}
.sn{{background:var(--ok-bg);border-radius:8px;padding:12px 16px;margin-top:12px;font-size:.84rem;color:#15803D;line-height:1.6;border-left:3px solid var(--ok)}}
.sn b{{color:#166534}}
.sn .rm-list{{font-weight:600}}
.nr{{color:var(--text3);font-style:italic;font-size:.86rem}}

/* ── Strikethrough (weekly check-off) ── */
.wi .ck{{width:20px;height:20px;border-radius:6px;border:1.5px solid var(--border);cursor:pointer;flex-shrink:0;display:flex;align-items:center;justify-content:center;transition:all .15s;-webkit-tap-highlight-color:transparent;margin-top:14px}}
.wi .ck:hover{{border-color:var(--text3)}}
.wi .ck.done{{background:var(--border);border-color:var(--border);color:#fff;font-size:.7rem}}
.wi.chk .wh-name{{text-decoration:line-through;color:var(--border)}}
.wi.chk .sb,.wi.chk .ag{{opacity:.2}}
.wi.chk .wh .ar{{color:var(--border)}}

/* ── Mobile ── */
@media(max-width:600px){{
  body{{padding:16px 14px;font-size:14px}}
  h1{{font-size:1.3rem;margin:12px 0 18px}}
  .tabs{{margin-bottom:20px}}
  .tb{{padding:9px 0;font-size:.86rem}}
  .cal{{gap:3px;margin-bottom:20px}}
  .ch{{font-size:.72rem;padding:8px 1px 4px}}
  .cd{{min-height:0;padding:5px 4px;border-radius:8px}}
  .cd .dn{{font-size:.82rem;margin-bottom:2px}}
  .cd .mp{{font-size:.6rem;line-height:1.35}}
  .cd .mp-div{{margin:2px 0}}
  .cd .ad,.cd .sd{{width:6px;height:6px}}
  .fc{{font-size:.74rem;padding:5px 9px}}
  .wn{{gap:10px;margin-bottom:16px}}
  .wn button{{padding:8px 14px;font-size:.84rem}}
  .wn .wl{{font-size:.96rem}}
  .wr{{gap:5px;margin-bottom:16px}}
  .wc{{padding:8px 4px;border-radius:8px}}
  .wc .wd{{font-size:.68rem}}
  .wc .wdt{{font-size:1rem}}
  .wp{{padding:22px 18px;border-radius:12px}}
  .wp .wpt{{font-size:1.05rem}}
  .wh{{padding:12px 0 6px;font-size:.9rem;gap:6px}}
  .rc.op{{padding:12px;font-size:.84rem;line-height:1.7}}
  .sn{{padding:10px 12px;font-size:.8rem}}
  .sb{{font-size:.68rem;padding:2px 7px}}
  .ag{{font-size:.68rem;padding:2px 6px}}
  .at{{gap:3px}}
  .sec-hd{{padding:12px 0 2px}}
}}
@media(max-width:380px){{
  .cd{{padding:4px 3px}}
  .cd .dn{{font-size:.76rem}}
  .cd .mp{{font-size:.54rem;line-height:1.3}}
  .wc{{padding:8px 3px}}
  .wc .wdt{{font-size:.92rem}}
}}
</style>
</head>
<body>
<h1>{month}월 어린이집 식단</h1>

<div class="tabs">
  <button class="tb on" onclick="sw('m')">월간</button>
  <button class="tb" onclick="sw('w')">주간</button>
</div>

<!-- Filter (collapsible) -->
<div class="fs">
  <div class="fs-toggle" id="fsTg" onclick="tgFs()">
    <span class="fs-icon" id="fsIc">&#9888;</span>
    <span class="fs-label">알레르기 필터</span>
    <span class="fs-cnt" id="fsCnt"></span>
    <span class="fs-arr">&#9660;</span>
  </div>
  <div class="fg-wrap" id="fgW">
    <div class="fg" id="aG"></div>
  </div>
</div>
<div class="fb" id="fB">
  <span>선택:</span><span class="ts" id="fT"></span>
  <button class="cl" onclick="clr()">해제</button>
</div>

<!-- MONTHLY -->
<div class="tc on" id="t-m">
<div class="srch">
  <input type="text" id="sQ" placeholder="메뉴 검색 (예: 카레, 불고기)" oninput="doSr()">
  <button class="sr-cl" id="sC" onclick="clrSr()">&#10005;</button>
</div>
<div class="sr-cnt" id="sR" style="display:none"></div>
{calendar_html}
</div>

<!-- WEEKLY -->
<div class="tc" id="t-w">
<div class="wn"><button onclick="pw()">&#8249; 이전</button><span class="wl" id="wL"></span><button onclick="nw()">다음 &#8250;</button></div>
<div class="wr" id="wR"></div>
<div class="wp" id="wP"></div>
</div>

<script>
{js_data}
{an_js}
const dk=['일','월','화','수','목','금','토'];
{weeks_js}
let wI=0;
const F=new Set();
let fsOpen=false;

function bld(){{
  const g=document.getElementById('aG');
  Object.entries(AN).forEach(([n,nm])=>{{
    const c=document.createElement('span');c.className='fc';c.dataset.n=n;
    c.innerHTML=`<span class="n">${{n}}</span> ${{nm}}`;
    c.onclick=()=>tg(+n);g.appendChild(c);
  }});
}}

/* Filter toggle */
function tgFs(){{
  fsOpen=!fsOpen;
  document.getElementById('fgW').classList.toggle('open',fsOpen);
  document.getElementById('fsTg').classList.toggle('open',fsOpen);
}}

function tg(n){{if(F.has(n))F.delete(n);else F.add(n);
  document.querySelectorAll('.fc').forEach(c=>c.classList.toggle('on',F.has(+c.dataset.n)));ub();af();
  document.getElementById('fsTg').classList.toggle('on',F.size>0);
  document.getElementById('fsIc').textContent=F.size>0?F.size:'\\u26A0';
  document.getElementById('fsCnt').textContent=F.size>0?F.size+'종 선택':'';
}}
function clr(){{F.clear();document.querySelectorAll('.fc').forEach(c=>c.classList.remove('on'));ub();af();
  document.getElementById('fsTg').classList.remove('on');
  document.getElementById('fsIc').textContent='\\u26A0';
  document.getElementById('fsCnt').textContent='';
}}
function ub(){{const b=document.getElementById('fB'),t=document.getElementById('fT');
  if(!F.size){{b.classList.remove('sh');return}}b.classList.add('sh');
  t.innerHTML=[...F].sort((a,b)=>a-b).map(n=>`<span class="ft">${{n}}. ${{AN[n]}}</span>`).join('');}}

function ihf(it){{return F.size>0&&it.allergy&&it.allergy.some(n=>F.has(n));}}
function dhf(ds){{return(D[ds]||[]).some(it=>ihf(it));}}
function dhs(ds){{return(D[ds]||[]).some(it=>ihf(it)&&it.rm&&it.rm.length>0);}}

function af(){{
  document.querySelectorAll('.cd[data-d]').forEach(el=>{{
    const ds=el.dataset.d,ad=el.querySelector('.ad'),sd=el.querySelector('.sd');
    const h=F.size>0&&dhf(ds),s=F.size>0&&dhs(ds);
    ad.style.display=h?'block':'none';sd.style.display=s?'block':'none';
    el.style.borderColor=h?'var(--alert-light)':'';
    el.style.background=h?'var(--alert-bg)':'';
    // 월간 메뉴명 색상 적용
    const mp=el.querySelector('.mp');
    if(mp){{
      const its=D[ds]||[];
      const secOrder=['오전간식','점심','오후간식'];
      const secMap={{}};
      its.forEach(it=>{{const sc=it.sec||'점심';if(!secMap[sc])secMap[sc]=[];secMap[sc].push(it);}});
      const parts=[];
      secOrder.forEach(sc=>{{
        const items=secMap[sc]||[];
        if(!items.length)return;
        const names=items.map(it=>{{
          const n=it.name.split('\\n')[0];
          const en=esc(n);
          if(F.size===0)return en;
          const hit=it.allergy&&it.allergy.some(a=>F.has(a));
          if(!hit)return en;
          const hasSauce=it.rm&&it.rm.length>0;
          return hasSauce?`<span class="ms">${{en}}</span>`:`<span class="ma">${{en}}</span>`;
        }});
        parts.push(names.join('<br>'));
      }});
      mp.innerHTML=parts.join('<span class="mp-div"></span>');
    }}
  }});
  // 주간 상세의 알레르기·양념 태그 토글
  document.querySelectorAll('.wh .at').forEach(el=>el.classList.toggle('sh',F.size>0));
  document.querySelectorAll('.wh .sb').forEach(el=>el.classList.toggle('sh',F.size>0));
  rw();
  const wp=document.getElementById('wP');
  if(wp.classList.contains('sh')&&wp.dataset.d)rwD(wp.dataset.d);
}}

function sw(t){{
  document.querySelectorAll('.tb').forEach((b,i)=>b.classList.toggle('on',t==='m'?i===0:i===1));
  document.querySelectorAll('.tc').forEach((c,i)=>c.classList.toggle('on',t==='m'?i===0:i===1));
  if(t==='w')rw();
}}

// Sauce badge HTML — sh class when filter active
function sbH(it){{
  if(!it.rm||!it.rm.length) return '';
  const v=F.size>0?' sh':'';
  const rmL=it.rm.map(n=>AN[n]).join('\\u00b7');
  if(it.ig&&it.ig.length){{
    return `<span class="sb p${{v}}" title="양념 제외 시 일부 대응 가능">양념 제외 가능</span>`;
  }}
  return `<span class="sb f${{v}}" title="${{rmL}}: 양념 제외로 대응 가능">양념 제외 가능</span>`;
}}

// Allergy tags HTML — sh class when filter active
function atH(it){{
  if(!it.allergy||!it.allergy.length) return '';
  const v=F.size>0?' sh':'';
  let h=`<span class="at${{v}}">`;
  it.allergy.forEach(n=>{{const hl=F.has(n)?' hl':'';h+=`<span class="ag${{hl}}">${{n}}</span>`}});
  return h+'</span>';
}}

// MONTHLY -> click switches to weekly tab
function mc(ds){{
  const wi=wks.findIndex(wk=>wk.includes(ds));
  if(wi<0)return;
  wI=wi;
  sw('w');
  setTimeout(()=>wc2(ds),50);
}}

// WEEKLY
function pw(){{if(wI>0){{wI--;rw();}}}}
function nw(){{if(wI<wks.length-1){{wI++;rw();}}}}
function rw(){{
  const wk=wks[wI];if(!wk)return;
  const fds=wk.filter(d=>d);
  if(!fds.length){{document.getElementById('wR').innerHTML='';return;}}
  const fd=new Date(fds[0]+'T00:00:00'),ld=new Date(fds[fds.length-1]+'T00:00:00');
  document.getElementById('wL').textContent=`${{fd.getMonth()+1}}/${{fd.getDate()}} ~ ${{ld.getMonth()+1}}/${{ld.getDate()}}`;
  const row=document.getElementById('wR');row.innerHTML='';
  row.style.gridTemplateColumns=`repeat(${{fds.length}},1fr)`;
  wk.forEach((ds)=>{{
    if(!ds)return;
    const d=new Date(ds+'T00:00:00');
    const its=D[ds]||[],c=document.createElement('div');c.className='wc';c.dataset.d=ds;
    if(!its.length){{
      c.classList.add('em');
      c.innerHTML=`<div class="wd">${{dk[d.getDay()]}}</div><div class="wdt">${{d.getDate()}}</div>`;
      row.appendChild(c);return;
    }}
    const hw=F.size>0&&dhf(ds),hs2=F.size>0&&dhs(ds);
    const adS=hw?'display:block':'',sdS=hs2?'display:block':'';
    c.innerHTML=`<div class="wdots"><span class="wd-a" style="${{adS}}"></span><span class="wd-s" style="${{sdS}}"></span></div><div class="wd">${{dk[d.getDay()]}}</div><div class="wdt">${{d.getDate()}}</div>`;
    const wp=document.getElementById('wP');
    if(wp.dataset.d===ds)c.classList.add('ac');
    c.onclick=()=>wc2(ds);row.appendChild(c);
  }});
  const wp=document.getElementById('wP');
  if(wp.dataset.d&&wk.includes(wp.dataset.d))rwD(wp.dataset.d);
  else wp.classList.remove('sh');
}}
function wc2(ds){{
  document.querySelectorAll('.wc').forEach(e=>e.classList.remove('ac'));
  const el=document.querySelector(`.wc[data-d="${{ds}}"]`);if(el)el.classList.add('ac');
  rwDS(ds);
}}
function rwD(ds){{
  const p=document.getElementById('wP');p.dataset.d=ds;
  const its=D[ds]||[];if(!its.length){{p.classList.remove('sh');return}}
  const d=new Date(ds+'T00:00:00'),dkn=dk[d.getDay()];
  let h=`<div class="wp-top"><div class="wpt">${{d.getMonth()+1}}월 ${{d.getDate()}}일 (${{dkn}})</div><button class="cp-btn" onclick="cpDay('${{ds}}')">복사하기</button></div><ul class="wl2">`;
  let prevSec='';
  its.forEach((it,idx)=>{{
    if(it.sec&&it.sec!==prevSec){{h+=`<li class="sec-hd">${{it.sec}}</li>`;prevSec=it.sec;}}
    const id=`r_${{ds}}_${{idx}}`;
    let rc;
    if(it.recipe){{
      rc=it.recipe.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\n/g,'<br>');
      if(it.rm&&it.rm.length&&it.ss){{
        const se=it.ss.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        const rmL=it.rm.map(n=>AN[n]).join('\\u00b7');
        let note;
        if(it.ig&&it.ig.length){{
          const igL=it.ig.map(n=>AN[n]).join('\\u00b7');
          note=`<b>양념 제외 시 <span class="rm-list">${{rmL}}</span> 대응 가능</b> (${{igL}}은 재료 자체에 포함)`;
        }} else {{
          note=`<b>양념 제외 시 <span class="rm-list">${{rmL}}</span> 모두 대응 가능</b> — 간장 대신 소금으로 대체 요청하세요.`;
        }}
        rc+=`<div class="sn">${{note}}<br><em>${{se}}</em></div>`;
      }}
    }} else {{ rc='<span class="nr">별도 조리방법 없음</span>'; }}
    h+=`<li class="wi" id="li_${{id}}"><div class="ck" onclick="tgCk(this.closest('.wi'),event)"></div><div class="wh" onclick="tr('${{id}}')"><span class="wh-name">${{esc(it.name)}}</span>${{sbH(it)}}${{atH(it)}}<span class="ar">&#9660;</span></div><div class="rc" id="${{id}}"><p>${{rc}}</p></div></li>`;
  }});
  p.innerHTML=h+'</ul>';p.classList.add('sh');
}}
function rwDS(ds){{rwD(ds);const p=document.getElementById('wP');if(p.classList.contains('sh'))p.scrollIntoView({{behavior:'smooth',block:'nearest'}});}}
function tr(id){{const e=document.getElementById(id),h=e.previousElementSibling;e.classList.toggle('op');h.classList.toggle('op');}}
function esc(s){{return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}}

// ── Search ──
function doSr(){{
  const q=document.getElementById('sQ').value.trim().toLowerCase();
  const cl=document.getElementById('sC');
  const sr=document.getElementById('sR');
  cl.classList.toggle('sh',q.length>0);
  if(!q){{clrSr();return;}}
  let cnt=0;
  document.querySelectorAll('.cd[data-d]').forEach(el=>{{
    const ds=el.dataset.d,its=D[ds]||[];
    const match=its.some(it=>it.name.toLowerCase().includes(q));
    el.classList.toggle('sr-hl',match);
    el.classList.toggle('sr-dim',!match);
    if(match)cnt++;
  }});
  sr.style.display='block';
  sr.textContent=cnt>0?`"${{q}}" 검색 결과: ${{cnt}}일`:`"${{q}}" 검색 결과 없음`;
}}
function clrSr(){{
  document.getElementById('sQ').value='';
  document.getElementById('sC').classList.remove('sh');
  document.getElementById('sR').style.display='none';
  document.querySelectorAll('.cd[data-d]').forEach(el=>{{
    el.classList.remove('sr-hl','sr-dim');
  }});
}}

// ── Weekly check-off (strikethrough) ──
function tgCk(li,ev){{
  ev.stopPropagation();
  li.classList.toggle('chk');
  const ck=li.querySelector('.ck');
  ck.classList.toggle('done');
  ck.textContent=ck.classList.contains('done')?'\\u2713':'';
}}

// ── Copy day menu as text ──
function cpDay(ds){{
  const d=new Date(ds+'T00:00:00'),dkn=dk[d.getDay()];
  const title=`${{d.getMonth()+1}}월 ${{d.getDate()}}일 (${{dkn}})`;
  const its=D[ds]||[];
  const secMap={{'오전간식':'오전','점심':'점심','오후간식':'오후'}};
  const secOrder=['오전간식','점심','오후간식'];
  const byS={{}};
  its.forEach((it,idx)=>{{
    const s=it.sec||'점심';
    if(!byS[s])byS[s]=[];
    const li=document.getElementById(`li_r_${{ds}}_${{idx}}`);
    const checked=li&&li.classList.contains('chk');
    const name=it.name.split('\\n')[0];
    byS[s].push(checked?`${{name}}(X)`:`${{name}}(O)`);
  }});
  const lines=[title];
  secOrder.forEach(s=>{{
    if(byS[s]&&byS[s].length){{
      lines.push(`${{secMap[s]}}: ${{byS[s].join(', ')}}`);
    }}
  }});
  const text=lines.join('\\n');
  navigator.clipboard.writeText(text).then(()=>{{
    const btn=document.querySelector('.cp-btn');
    btn.textContent='복사됨!';btn.classList.add('ok');
    setTimeout(()=>{{btn.textContent='복사하기';btn.classList.remove('ok');}},1500);
  }});
}}

document.addEventListener('DOMContentLoaded',()=>{{bld();rw();}});
</script>
</body>
</html>"""


def _esc(s: str) -> str:
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ═══════════════════════════════════════
# 5. JSON 데이터 내보내기 (뷰어 분리용)
# ═══════════════════════════════════════

def build_json_data(menus: dict, allergy_match: dict, sauce_data: dict,
                    year: int, month: int) -> dict:
    """뷰어에서 사용할 JSON 데이터를 생성한다."""
    # 메뉴 데이터 (기존 build_js_data 로직)
    import re
    alt_pattern = re.compile(r'^만\s*\d')  # "만1-2세", "만 3-5세" 등

    menu_data = {}
    for ds, items in sorted(menus.items()):
        match_data = allergy_match.get(ds, {})
        sauce_items = {s['name']: s for s in sauce_data.get(ds, [])}
        js_items = []
        for item in items:
            name = item['name'].split('\n')[0].strip()
            allergy_nums = match_data.get(name, [])
            sauce = sauce_items.get(name, {})

            # 대체 식단 감지: "만1-2세 ..." 패턴이면 바로 앞 메뉴의 alt로 묶기
            if alt_pattern.match(name) and js_items:
                alt_text = item['name'].replace('\n', ' ').strip()
                js_items[-1]['alt'] = alt_text
                continue

            js_item = {
                "name": item['name'],
                "recipe": item.get('recipe', ''),
                "allergy": allergy_nums,
                "sec": item.get('section', '점심'),
            }
            if sauce.get('removable'):
                js_item['rm'] = sauce['removable']
            if sauce.get('integral'):
                js_item['ig'] = sauce['integral']
            if sauce.get('sauceStep'):
                js_item['ss'] = sauce['sauceStep']
            js_items.append(js_item)
        menu_data[ds] = js_items

    # 주간 배열 (월~토, 일요일 제외)
    first_day = date(year, month, 1)
    days_since_monday = first_day.weekday()
    week_start = first_day - timedelta(days=days_since_monday)
    weeks = []
    d = week_start
    for _ in range(6):
        wk = []
        for _ in range(7):
            if d.weekday() != 6:  # 일요일 제외
                wk.append(d.isoformat() if d.month == month else None)
            d += timedelta(days=1)
        if any(wk):
            weeks.append(wk)

    return {
        "year": year,
        "month": month,
        "menus": menu_data,
        "allergens": ALLERGEN_NAMES,
        "weeks": weeks,
    }


def export_json(menus: dict, allergy_match: dict, sauce_data: dict,
                year: int, month: int, output_path: str):
    """JSON 데이터 파일을 내보낸다."""
    data = build_json_data(menus, allergy_match, sauce_data, year, month)
    Path(output_path).write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8'
    )
    print(f"[INFO] JSON 데이터 내보내기 완료: {output_path}")
    return data


# ═══════════════════════════════════════
# Main
# ═══════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description='어린이집 점심 메뉴 HTML 생성기')
    parser.add_argument('--recipe', required=True, help='레시피 엑셀 파일 경로')
    parser.add_argument('--allergy', required=True, help='알레르기 식단표 PDF 경로')
    parser.add_argument('--output', help='출력 HTML 경로 (미지정 시 자동)')
    args = parser.parse_args()

    print("=" * 50)
    print("어린이집 점심 메뉴 HTML 생성기")
    print("=" * 50)

    # Step 1: 메뉴 추출
    print("\n[1/4] 레시피 엑셀에서 메뉴 추출 중...")
    menus = extract_menus(args.recipe)

    # 연/월 감지
    year, month = detect_month_year(menus)
    print(f"[INFO] 감지된 기간: {year}년 {month}월")

    # Step 2: 알레르기 추출 및 매칭
    print("\n[2/4] PDF에서 알레르기 정보 추출 중...")
    allergy_data = extract_allergies(args.allergy, list(menus.keys()))
    allergy_match = match_allergies(menus, allergy_data)

    # Step 3: 양념 분석
    print("\n[3/4] 양념 제외 가능 분석 중...")
    sauce_data = analyze_sauce(menus, allergy_match)

    # Step 4: HTML 생성
    print("\n[4/4] HTML 생성 중...")
    html = generate_html(menus, allergy_match, sauce_data, year, month)

    output_path = args.output or f'{month}월_점심메뉴.html'
    Path(output_path).write_text(html, encoding='utf-8')
    print(f"\n[완료] {output_path} 생성됨 ({len(html):,} bytes)")


if __name__ == '__main__':
    main()
