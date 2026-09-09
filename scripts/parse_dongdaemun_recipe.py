#!/usr/bin/env python3
"""동대문구 표준조리법 PDF → 수동 JSON.

식단표 PDF 없이 표준조리법만 받았을 때 쓴다. 조리법 PDF의 표가
날짜·끼니·음식명(알레르기 원문자 포함)·조리방법을 모두 담고 있어 식단 복원이 가능하다.

**조리방법은 음식명과 같은 행에서 그대로 들고 온다.** 예전엔 이 파서로 메뉴만 뽑고
조리법은 같은 PDF를 `parse_recipe_pdf`로 다시 읽어 이름으로 재결합했는데,
두 파서의 이름 정규화가 달라 매달 다른 항목이 조용히 새어나갔다. 같은 행에 이미
있는 값을 쓰면 붙이는 단계 자체가 없어진다.

사용: python3 scripts/parse_dongdaemun_recipe.py <pdf> <out.json> --year 2026 --month 8
"""
import argparse
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

import pdfplumber

sys.path.insert(0, str(Path(__file__).resolve().parent))
from parse_common import is_recipe_bleed

CIRCLED = {chr(0x2460 + i): i + 1 for i in range(20)}  # ①..⑳
CIRCLED_RE = re.compile('[' + ''.join(CIRCLED) + ']')
DATE_RE = re.compile(r'(\d{1,2})\s*\[([월화수목금토일])\]')
SECTIONS = ('오전간식', '점심', '오후간식', '저녁')
LEGEND = ('①난류②우유③메밀④땅콩⑤대두⑥밀⑦고등어⑧게⑨새우⑩돼지고기⑪복숭아'
          '⑫토마토⑬아황산류⑭호두⑮닭고기⑯소고기⑰오징어⑱조개류⑲잣⑳기타')


def split_name(cell: str):
    """'햄채소볶음밥\n①②⑤⑥⑩⑮\n⑯' → ('햄채소볶음밥', [1,2,5,6,10,15,16])"""
    allergy = sorted({CIRCLED[c] for c in CIRCLED_RE.findall(cell)})
    name = CIRCLED_RE.sub('', cell)
    name = re.sub(r'\s+', '', name)
    return name, allergy


def split_recipe(cell: str, prev_cell: str) -> str:
    """조리방법 칸 정리. 원본이 '-'면 조리법 없음, 위 행에서 밀려온 것도 버린다."""
    rec = (cell or '').replace('_x000D_', '').strip()
    if rec in ('', '-', 'None'):
        return ''
    if is_recipe_bleed(rec, prev_cell):
        return ''
    return rec


def parse(pdf_path: str, year: int, month: int):
    menus = OrderedDict()
    cur_date = None
    cur_sec = None
    skipped = []
    prev_rec = ''

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    cells = [(c or '').strip() for c in row]
                    if len(cells) < 3:
                        continue
                    d_cell, s_cell, n_cell = cells[0], cells[1], cells[2]
                    r_cell = cells[5] if len(cells) > 5 else ''
                    recipe = split_recipe(r_cell, prev_rec)
                    prev_rec = r_cell

                    if d_cell.startswith('날짜') or '[동대문구센터]' in d_cell:
                        continue

                    m = DATE_RE.search(d_cell)
                    if m:
                        cur_date = f'{year:04d}-{month:02d}-{int(m.group(1)):02d}'
                        menus.setdefault(cur_date, [])
                    elif d_cell:
                        # '[생일식단]' 등 날짜 없는 블록. 앞 날짜에 딸려 들어가지 않게 끊는다.
                        cur_date = None
                        skipped.append(('날짜 없는 블록 시작', re.sub(r'\s+', '', d_cell)))

                    if s_cell:
                        sec = re.sub(r'\s+', '', s_cell)
                        if sec in SECTIONS:
                            cur_sec = sec
                        else:
                            skipped.append(('끼니 미상', s_cell))

                    if not n_cell:
                        continue
                    if cur_date is None or cur_sec is None:
                        skipped.append(('날짜/끼니 미정 상태의 음식명', n_cell))
                        continue

                    name, allergy = split_name(n_cell)
                    if not name:
                        skipped.append(('이름 없는 음식명 셀', n_cell))
                        continue
                    menus[cur_date].append(
                        {'name': name, 'allergy': allergy,
                         'sec': cur_sec, 'recipe': recipe})

    return menus, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('out')
    ap.add_argument('--year', type=int, required=True)
    ap.add_argument('--month', type=int, required=True)
    args = ap.parse_args()

    menus, skipped = parse(args.pdf, args.year, args.month)
    if not menus:
        print('[ERROR] 추출된 날짜가 없다. 표 구조를 다시 확인할 것.', file=sys.stderr)
        return 1

    # 같은 날 같은 끼니에서 이름 중복 제거 (표가 페이지를 넘길 때 헤더 반복 대비).
    # 중복 중 조리법이 있는 쪽을 살린다 — 페이지가 갈리면 뒤쪽에만 실려 있을 수 있다.
    for date, items in menus.items():
        seen = {}
        uniq = []
        for it in items:
            key = (it['sec'], it['name'])
            if key in seen:
                kept = seen[key]
                if not kept.get('recipe') and it.get('recipe'):
                    kept['recipe'] = it['recipe']
                continue
            seen[key] = it
            uniq.append(it)
        menus[date] = uniq

    doc = {
        '_source': args.pdf.split('/')[-1] + ' (동대문구 표준조리법)',
        '_note': ('식단표 없이 표준조리법 PDF만 받아 파싱. 조리법 표의 날짜·끼니·음식명·'
                  '알레르기 원문자·조리방법을 같은 행에서 그대로 사용(추정 없음, 이름 매칭 없음). '
                  '조리방법이 빈 항목은 원본 칸이 `-`인 것 — 생과일·우유·완제품 간식.'),
        '_legend': LEGEND,
        'year': args.year,
        'month': args.month,
        'menus': menus,
    }
    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)

    total = sum(len(v) for v in menus.values())
    with_rec = sum(1 for v in menus.values() for it in v if it.get('recipe'))
    print(f'날짜 {len(menus)}일 / 항목 {total}개 → {args.out}')
    print(f'조리방법 {with_rec}/{total} (나머지 {total - with_rec}건은 원본 칸이 `-`)')
    if skipped:
        print(f'[확인 필요] 건너뛴 셀 {len(skipped)}건:')
        for reason, val in skipped[:15]:
            print(f'  - {reason}: {val[:50]!r}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
