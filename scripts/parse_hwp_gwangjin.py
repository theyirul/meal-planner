#!/usr/bin/env python3
"""광진구 HWP 식단표 파서.

창원(parse_hwp.py)과 구조가 다르다:
- 헤더가 '날짜'가 아니라 '일자', 날짜 표기가 '31(월)' 괄호형(뒤에 '광복절' 등 꼬리표 가능)
- 알레르기가 원문자(①)가 아니라 아라비아 숫자 괄호 '(5,6)'
- 메뉴가 셀 안 공백 구분이 아니라 **행마다 한 칸씩 세로로** 쌓임 (메뉴명에 공백이 들어감)
- 한 달치가 표 하나에 들어있고, 고아 월요일(8/31)이 첫 행에 끼워짐

병합셀 rowspan carry 때문에 같은 메뉴가 여러 행에 반복된다. 끼니 라벨(col 0)과
메뉴 셀의 rowspan 경계가 어긋나므로, 메뉴가 걸친 행들의 끼니 라벨을
다수결로 정한다(동수면 뒤쪽 끼니).

사용: python3 scripts/parse_hwp_gwangjin.py <hwp> --year 2026 --month 8
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_hwp import _Grid, _find_hwp5html  # noqa: E402

SEC_ORDER = ['오전간식', '점심', '오후간식', '저녁']
DATE_RE = re.compile(r'^(\d{1,2})\s*\(([월화수목금토일])\)')
KCAL_RE = re.compile(r'^[\d.]+\s*/\s*[\d.]+$')          # '411/15' 열량/단백질 값
ALLERGY_RE = re.compile(r'\(\s*\d+(?:\s*,\s*\d+)*\s*\)')  # '(5,6)' 숫자만
ALT_RE = re.compile(r'\s+\((.+)\)$')                      # ' (두유(5))' 앞 공백 = 대체메뉴


def _to_html(hwp_path: str) -> str:
    tmp = tempfile.mkdtemp(prefix='hwp5_gj_')
    try:
        subprocess.run([_find_hwp5html(), '--output', tmp, hwp_path],
                       check=True, capture_output=True)
        return open(os.path.join(tmp, 'index.xhtml'), encoding='utf-8').read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _split_top(text: str) -> list[str]:
    """괄호 밖의 '/'로만 분리. '떠먹는요구르트(2)/ (데친)당근스틱' → 2개."""
    parts, buf, depth = [], [], 0
    for ch in text:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth = max(0, depth - 1)
        if ch == '/' and depth == 0:
            parts.append(''.join(buf)); buf = []
        else:
            buf.append(ch)
    parts.append(''.join(buf))
    return [p.strip() for p in parts if p.strip()]


def _strip_allergy(text: str):
    """'(백)배추김치(9)' → ('(백)배추김치', [9]). 숫자만 든 괄호만 알레르기로 본다."""
    nums = []
    for g in ALLERGY_RE.findall(text):
        nums += [int(x) for x in re.findall(r'\d+', g)]
    name = ALLERGY_RE.sub('', text)
    name = re.sub(r'\s+', ' ', name).strip()
    return name, sorted(set(nums))


def _parse_cell(raw: str, sec: str) -> list[dict]:
    """셀 텍스트 → 메뉴 항목들. 대체메뉴는 별도 항목으로 분리(7월 데이터 관례)."""
    raw = re.sub(r'\s+', ' ', raw).strip()
    if not raw or KCAL_RE.match(raw):
        return []
    items = []
    for part in _split_top(raw):
        alt = None
        m = ALT_RE.search(part)
        if m and re.search(r'[가-힣]', ALLERGY_RE.sub('', m.group(1))):
            alt = m.group(1).strip()
            part = part[:m.start()].strip()
        name, allergy = _strip_allergy(part)
        if not name:
            continue
        items.append({'name': name, 'allergy': allergy, 'sec': sec})
        if alt:
            aname, aallergy = _strip_allergy(alt)
            if aname:
                items.append({'name': f'{aname}\n*{name} 대체메뉴',
                              'allergy': aallergy, 'sec': sec})
    return items


def parse(hwp_path: str, year: int, month: int):
    g = _Grid()
    g.feed(_to_html(hwp_path))
    grid = max(g.tables, key=len) if g.tables else None
    if not grid:
        raise RuntimeError('HWP에서 표를 찾지 못함')

    header_rows = [i for i, row in enumerate(grid)
                   if any((v or '').strip() == '일자' for v in row.values())]
    if not header_rows:
        raise RuntimeError("'일자' 헤더 행을 찾지 못함")

    menus, notes = {}, []
    for bi, hr in enumerate(header_rows):
        end = header_rows[bi + 1] if bi + 1 < len(header_rows) else len(grid)
        date_cols = {}
        for c, v in grid[hr].items():
            m = DATE_RE.match((v or '').strip())
            if m:
                date_cols[c] = int(m.group(1))

        # 끼니 라벨은 col 0. '열량...' 행부터는 식단이 아니다.
        sec_by_row = {}
        for r in range(hr + 1, end):
            lbl = re.sub(r'\s+', '', (grid[r].get(0) or ''))
            if lbl.startswith('열량'):
                break
            if lbl in SEC_ORDER:
                sec_by_row[r] = lbl
        if not sec_by_row:
            continue
        body = sorted(sec_by_row)

        for col, day in date_cols.items():
            # 같은 텍스트가 이어지는 행들 = 하나의 병합셀 → 끼니 다수결
            runs, prev, start = [], None, None
            for r in body + [None]:
                cur = re.sub(r'\s+', ' ', (grid[r].get(col) or '').strip()) if r is not None else None
                if cur != prev:
                    if prev:
                        runs.append((prev, start, last))
                    prev, start = cur, r
                last = r
            items = []
            for text, r0, r1 in runs:
                rows = [r for r in body if r0 <= r <= r1]
                counts = {}
                for r in rows:
                    counts[sec_by_row[r]] = counts.get(sec_by_row[r], 0) + 1
                sec = max(counts, key=lambda s: (counts[s], SEC_ORDER.index(s)))
                if len(counts) > 1:
                    notes.append(f'{month}/{day} "{text[:24]}" 끼니 경계 어긋남 '
                                 f'→ {counts} 중 {sec} 선택')
                items += _parse_cell(text, sec)
            if items:
                key = f'{year:04d}-{month:02d}-{day:02d}'
                items.sort(key=lambda i: SEC_ORDER.index(i['sec']))
                menus[key] = items

    return {k: menus[k] for k in sorted(menus)}, notes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('hwp')
    ap.add_argument('--year', type=int, required=True)
    ap.add_argument('--month', type=int, required=True)
    ap.add_argument('--out')
    args = ap.parse_args()

    menus, notes = parse(args.hwp, args.year, args.month)
    total = sum(len(v) for v in menus.values())
    print(f'날짜 {len(menus)}일 / 항목 {total}개')
    if notes:
        print(f'[끼니 경계 보정 {len(notes)}건]')
        for n in notes[:20]:
            print('  -', n)

    if args.out:
        doc = {
            '_source': os.path.basename(args.hwp) + ' (광진구 식단표 HWP)',
            '_note': ('HWP 그리드 파싱. 알레르기는 원본 괄호숫자 표기 그대로(추정 없음). '
                      '대체메뉴는 별도 항목으로 분리. 룰 검증 필요.'),
            '_legend': ('1난류 2우유 3메밀 4땅콩 5대두 6밀 7고등어 8게 9새우 10돼지고기 '
                        '11복숭아 12토마토 13아황산류 14호두 15닭고기 16소고기 17오징어 '
                        '18조개류 19잣'),
            'year': args.year, 'month': args.month, 'menus': menus,
        }
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump(doc, f, ensure_ascii=False, indent=1)
        print('→', args.out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
