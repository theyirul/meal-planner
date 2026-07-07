#!/usr/bin/env python3
"""HWP 식단표 파서 (창원시 등 한글 파일 입력).

HWP 5.x → HTML(pyhwp hwp5html) → 병합셀(colspan/rowspan) 그리드 전개 → 메뉴 JSON.
- 셀 안 메뉴는 공백으로 구분, 각 메뉴 뒤 동그라미(①..⑲)가 알레르기.
- ♠(자연간식)·(고춧가루제외)·(1/2)·(냉) 등 괄호/기호 표기는 이름에 그대로 보존.
- 병합셀이 많아 반드시 그리드 좌표로 날짜열↔끼니열을 맞춰야 함(순서 인덱스 매핑 금지).

의존성: pyhwp (pip install pyhwp) → hwp5html 커맨드.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from glob import glob
from html.parser import HTMLParser

CIRCLED = {chr(0x2460 + i): i + 1 for i in range(20)}  # ①..⑳ → 1..20
CIRCLED_RE = re.compile('[' + ''.join(CIRCLED.keys()) + ']')
SEC_MAP = {'오전간식': '오전간식', '점심': '점심', '점 심': '점심', '오후간식': '오후간식'}


def _find_hwp5html() -> str:
    """hwp5html 실행 파일 경로 탐색."""
    p = shutil.which('hwp5html')
    if p:
        return p
    for cand in glob(os.path.expanduser('~/Library/Python/*/bin/hwp5html')) + \
                glob(os.path.expanduser('~/.local/bin/hwp5html')):
        if os.path.exists(cand):
            return cand
    raise RuntimeError("hwp5html 없음 → `pip3 install pyhwp` 후 재시도")


class _Grid(HTMLParser):
    """colspan/rowspan을 반영해 표를 절대 그리드 좌표(행별 {col:text})로 전개."""
    def __init__(self):
        super().__init__()
        self.tables = []; self.rows = None; self.carry = {}
        self.r = -1; self.col = 0; self.cell = None; self.attrs = None; self.on = False

    def handle_starttag(self, t, a):
        if t == 'table':
            self.rows = []; self.tables.append(self.rows); self.carry = {}; self.r = -1; self.on = True
        elif t == 'tr' and self.on:
            self.r += 1; self.col = 0; row = {}; self.rows.append(row)
            for c, (left, txt) in list(self.carry.items()):
                row[c] = txt
                self.carry[c] = (left - 1, txt) if left - 1 > 0 else None
                if self.carry[c] is None:
                    del self.carry[c]
        elif t in ('td', 'th') and self.on:
            self.cell = []; self.attrs = dict(a)

    def handle_endtag(self, t):
        if t == 'table' and self.on:
            self.on = False
        elif t in ('td', 'th') and self.cell is not None:
            txt = ' '.join(''.join(self.cell).split())
            cs = int(self.attrs.get('colspan', 1)); rs = int(self.attrs.get('rowspan', 1))
            row = self.rows[self.r]
            while self.col in row:
                self.col += 1
            start = self.col
            row[start] = txt
            for cc in range(start, start + cs):
                if cc != start:
                    row[cc] = ''  # 병합 점유 표시
                if rs > 1:
                    self.carry[cc] = (rs - 1, txt if cc == start else '')
            self.col = start + cs
            self.cell = None; self.attrs = None

    def handle_data(self, d):
        if self.cell is not None:
            self.cell.append(d)


def _parse_cell(raw: str, sec: str) -> list[dict]:
    items = []
    for tok in raw.split():
        allergy = [CIRCLED[c] for c in tok if c in CIRCLED]
        name = CIRCLED_RE.sub('', tok).strip()
        if not name:
            continue
        # 완전 괄호 토큰(재료설명 등)은 앞 메뉴명에 병합. "(냉)콩국수"처럼 괄호 뒤 글자 있으면 독립 항목.
        if name.startswith('(') and name.endswith(')') and items:
            items[-1]['name'] += name
            for x in allergy:
                if x not in items[-1]['allergy']:
                    items[-1]['allergy'].append(x)
            continue
        items.append({'name': name, 'allergy': allergy, 'sec': sec})
    return items


def _select_grid(tables: list, age: str):
    """연령(예: '1-2세')에 해당하는 식단 그리드 표 선택.
    표 순서: [제목(…연령…)][그리드][제목][그리드]… 제목에 연령 문자열이 있는 그리드를 고름."""
    age_key = age.replace(' ', '')
    grids = []          # (index, table)
    titles = {}         # index -> title text
    for idx, tb in enumerate(tables):
        flat = ' '.join(v for row in tb for v in row.values())
        is_grid = any((row.get(min(row)) or '').strip() in ('날짜',) or '날짜' in row.values()
                      for row in tb if row) and bool(re.search(r'\d{1,2}\s*\[', flat))
        if is_grid:
            grids.append((idx, tb))
        elif '식단표' in flat and re.search(r'\d\s*-\s*\d\s*세', flat):
            titles[idx] = flat
    # 그리드 직전의 제목에서 연령 매칭
    for gi, tb in grids:
        prev = max([i for i in titles if i < gi], default=None)
        if prev is not None and age_key in titles[prev].replace(' ', ''):
            return tb
    # 못 찾으면 첫 그리드
    return grids[0][1] if grids else None


def parse_hwp_menu(hwp_path: str, year: int, month: int, age: str = '1-2세') -> dict:
    """HWP 식단표 → {'year','month','menus'}."""
    hwp5html = _find_hwp5html()
    tmp = tempfile.mkdtemp(prefix='hwp5_')
    try:
        subprocess.run([hwp5html, '--output', tmp, hwp_path],
                       check=True, capture_output=True)
        html = open(os.path.join(tmp, 'index.xhtml'), encoding='utf-8').read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    g = _Grid(); g.feed(html)
    grid = _select_grid(g.tables, age)
    if grid is None:
        raise RuntimeError("HWP에서 식단 그리드를 찾지 못함")

    menus = {}
    i = 0
    while i < len(grid):
        row = grid[i]
        if '날짜' not in [(row.get(k) or '') for k in row]:
            i += 1; continue
        date_cols = {}
        for c in sorted(row):
            m = re.match(r'(\d{1,2})\s*\[', (row[c] or '').strip())
            if m:
                date_cols[c] = int(m.group(1))
        block = {}
        j = i + 1
        while j < len(grid):
            rj = grid[j]
            if '날짜' in [(rj.get(k) or '') for k in rj]:
                break
            for k in sorted(rj):
                lbl = (rj[k] or '').strip()
                if lbl in SEC_MAP:
                    block[SEC_MAP[lbl]] = rj
                    break
            j += 1
        for c, day in date_cols.items():
            items = []
            for sec in ('오전간식', '점심', '오후간식'):
                rj = block.get(sec)
                if rj and c in rj:
                    items.extend(_parse_cell((rj[c] or '').strip(), sec))
            if items:
                menus[f"{year}-{month:02d}-{day:02d}"] = items
        i = j

    return {'year': year, 'month': month, 'menus': menus}


if __name__ == '__main__':
    import json, sys
    d = parse_hwp_menu(sys.argv[1], int(sys.argv[2]), int(sys.argv[3]),
                       sys.argv[4] if len(sys.argv) > 4 else '1-2세')
    print(f"{len(d['menus'])}일자")
    print(json.dumps(d, ensure_ascii=False)[:500])
