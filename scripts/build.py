#!/usr/bin/env python3
"""
매일아침 — 데이터 빌드 스크립트

input/uploads/ 폴더의 파일을 읽어서 data/ 폴더에 JSON을 생성하고,
index.html에 임베딩합니다.

파일 네이밍: {YYYYMM}_{지역}_{연령}_{유형}_{타입}.{확장자}
예: 202603_광진구_1-2세_일반형_식단표.pdf

사용법:
  python3 scripts/build.py
"""

import json
import re
import sys
from pathlib import Path
from collections import defaultdict

ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT / 'input'
DATA_DIR = ROOT / 'data'
SCRIPTS_DIR = ROOT / 'scripts'

sys.path.insert(0, str(SCRIPTS_DIR))
from generate_menu import extract_menus, detect_month_year, extract_allergies, match_allergies, analyze_sauce, build_json_data
from parse_pdf import parse_pdf_menu, build_json_from_pdf
from parse_yeongdeungpo import parse_yd_pdf, parse_yd_excel, build_yd_json
from parse_seongnam import parse_sn_pdf, parse_sn_excel, build_sn_json

# 지역명 → ID 매핑
REGION_MAP = {
    "광진구": "gwangjin",
    "영등포구": "yeongdeungpo",
    "수원시": "suwon",
    "김포시": "gimpo",
    "용인시": "yongin",
    "경기도성남시": "seongnam",
    "성남시": "seongnam",
}

REGION_NAMES = {v: k for k, v in REGION_MAP.items()}

# 표시명 (UI에 보여줄 이름)
REGION_DISPLAY = {
    "광진구": "서울 광진구",
    "영등포구": "서울 영등포구",
    "수원시": "경기 수원시",
    "김포시": "경기 김포시",
    "용인시": "경기 용인시",
    "경기도성남시": "경기 성남시",
    "성남시": "경기 성남시",
}


def parse_upload_filename(filename: str) -> dict | None:
    """파일명에서 메타데이터 추출"""
    import unicodedata
    filename = unicodedata.normalize('NFC', filename)
    name = Path(filename).stem
    ext = Path(filename).suffix.lstrip('.')
    pattern = r'^(\d{6})_([가-힣]+)_(.+?)_(.+?)_(식단표|레시피)$'
    m = re.match(pattern, name)
    if not m:
        return None
    return {
        "yyyymm": m.group(1), "region": m.group(2),
        "age": m.group(3), "type": m.group(4),
        "subtype": m.group(5), "ext": ext, "filename": filename,
    }


def scan_uploads() -> dict:
    """uploads/ 폴더 스캔 → 지역+월별 그룹화"""
    upload_dir = INPUT_DIR / 'uploads'
    if not upload_dir.exists():
        return {}
    groups = defaultdict(list)
    for fpath in upload_dir.iterdir():
        if not fpath.is_file():
            continue
        meta = parse_upload_filename(fpath.name)
        if meta:
            key = f"{meta['yyyymm']}_{meta['region']}"
            groups[key].append(meta)
    return dict(groups)


def detect_pdf_format(pdf_path: Path) -> str:
    """
    PDF 구조를 분석해서 포맷 타입을 반환.
    - "yeongdeungpo": 1페이지 대형 테이블 (19열), 원형숫자 알레르기
    - "standard": 기존 광진구/수원 형식 (5행/주 블록, 괄호형 알레르기)
    """
    import pdfplumber
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            tables = pdf.pages[0].extract_tables()
            if not tables:
                return "standard"
            # 가장 큰 테이블의 열 수로 판단
            biggest = max(tables, key=lambda t: len(t))
            ncols = len(biggest[0]) if biggest else 0
            # 영등포구: 19열 대형 테이블, 첫 행에 '월','화' 등 요일
            if ncols >= 15:
                header = biggest[0]
                days = [str(c or '') for c in header]
                if '월' in days and '화' in days:
                    return "yeongdeungpo"
            # 성남시: 11열 테이블, 첫 셀에 "N주차"
            if ncols == 11:
                first_cell = str(biggest[0][0] or '').strip()
                if re.match(r'\d+주차', first_cell):
                    return "seongnam"
            return "standard"
    except Exception:
        return "standard"


def _process_yeongdeungpo(pdf_file: Path, xlsx_file: Path | None = None) -> dict | None:
    """영등포구 포맷 처리"""
    print(f"  [영등포구 포맷] PDF 파싱...")
    pdf_data = parse_yd_pdf(str(pdf_file))
    recipes = {}
    if xlsx_file and xlsx_file.exists():
        print(f"  [영등포구 포맷] Excel 레시피 파싱...")
        try:
            recipes = parse_yd_excel(str(xlsx_file))
        except Exception as e:
            print(f"  [WARN] Excel 파싱 실패: {e}")
    return build_yd_json(pdf_data, recipes)


def _process_seongnam(pdf_file: Path, xlsx_file: Path | None = None) -> dict | None:
    """성남시 포맷 처리"""
    print(f"  [성남시 포맷] PDF 파싱...")
    pdf_data = parse_sn_pdf(str(pdf_file))
    recipes = {}
    if xlsx_file and xlsx_file.exists():
        print(f"  [성남시 포맷] Excel 레시피 파싱...")
        try:
            recipes = parse_sn_excel(str(xlsx_file))
        except Exception as e:
            print(f"  [WARN] Excel 파싱 실패: {e}")
    return build_sn_json(pdf_data, recipes)


def _process_excel_pdf(recipe_file: Path, allergy_file: Path) -> dict | None:
    """엑셀 + PDF 조합으로 처리"""
    print(f"  메뉴 추출 중...")
    menus = extract_menus(str(recipe_file))
    year, month = detect_month_year(menus)
    print(f"  감지: {year}년 {month}월")
    print(f"  알레르기 매칭 중...")
    allergy_data = extract_allergies(str(allergy_file), list(menus.keys()))
    allergy_match = match_allergies(menus, allergy_data)
    print(f"  양념 분석 중...")
    sauce_data = analyze_sauce(menus, allergy_match)
    print(f"  JSON 생성 중...")
    return build_json_data(menus, allergy_match, sauce_data, year, month)


def _process_pdf_only(pdf_file: Path) -> dict | None:
    """PDF 하나에서 메뉴+알레르기 추출"""
    try:
        pdf_data = parse_pdf_menu(str(pdf_file))
        return build_json_from_pdf(pdf_data)
    except Exception as e:
        print(f"  [ERROR] PDF 파싱 실패: {e}")
        return None


def process_upload_group(group_key: str, files: list[dict]) -> dict | None:
    """uploads 그룹 처리 (같은 지역+월의 파일들)"""
    meta = files[0]
    region_name = meta['region']
    region_id = REGION_MAP.get(region_name)
    if not region_id:
        print(f"  [SKIP] 지역 매핑 없음: {region_name} → REGION_MAP에 추가 필요")
        return None

    upload_dir = INPUT_DIR / 'uploads'
    pdf_file = None
    xlsx_file = None
    for f in files:
        if f['subtype'] == '식단표' and f['ext'].lower() == 'pdf':
            pdf_file = upload_dir / f['filename']
        elif f['subtype'] == '레시피' and f['ext'].lower() == 'xlsx':
            xlsx_file = upload_dir / f['filename']

    if not pdf_file or not pdf_file.exists():
        print(f"  [SKIP] 식단표 PDF 없음")
        return None

    # 포맷 자동 감지
    fmt = detect_pdf_format(pdf_file)
    print(f"  [포맷 감지] {fmt}")

    if fmt == "seongnam":
        data = _process_seongnam(pdf_file, xlsx_file)
    elif fmt == "yeongdeungpo":
        data = _process_yeongdeungpo(pdf_file, xlsx_file)
    elif xlsx_file and xlsx_file.exists():
        print(f"  [Excel+PDF] {xlsx_file.name} + {pdf_file.name}")
        try:
            data = _process_excel_pdf(xlsx_file, pdf_file)
        except Exception as e:
            print(f"  [WARN] Excel 파싱 실패 ({e}) → PDF-only로 전환")
            data = _process_pdf_only(pdf_file)
    else:
        print(f"  [PDF-only] {pdf_file.name}")
        data = _process_pdf_only(pdf_file)

    if data is None:
        return None

    year, month = data['year'], data['month']
    month_key = f"{year}-{month:02d}"
    out_dir = DATA_DIR / region_id
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{month_key}.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"  → {out_path.relative_to(ROOT)}")

    display_name = REGION_DISPLAY.get(region_name, region_name)
    return {"id": region_id, "name": display_name, "region": display_name, "month_key": month_key}


def build_centers_json(results: list[dict]):
    """centers.json 생성 (지역 기반)"""
    merged = {}
    for r in results:
        rid = r['id']
        if rid in merged:
            months = set(merged[rid]['months'])
            months.add(r['month_key'])
            merged[rid]['months'] = sorted(months)
        else:
            merged[rid] = {
                "id": rid, "name": r['name'],
                "region": r['region'], "months": [r['month_key']]
            }
    centers = sorted(merged.values(), key=lambda c: c['name'])
    centers_path = DATA_DIR / 'centers.json'
    centers_path.write_text(json.dumps(centers, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"\n[OK] centers.json 업데이트 ({len(centers)}개 지역)")


def build_embedded_html():
    """index.html에 데이터 임베딩"""
    index_src = ROOT / 'index_server.html'
    if not index_src.exists():
        print("[WARN] index_server.html 없음 → 스킵")
        return
    centers = json.loads((DATA_DIR / 'centers.json').read_text(encoding='utf-8'))
    embedded_data = {}
    for c in centers:
        for m in c['months']:
            fpath = DATA_DIR / c['id'] / f"{m}.json"
            if fpath.exists():
                embedded_data[f"{c['id']}/{m}"] = json.loads(fpath.read_text(encoding='utf-8'))
    html = index_src.read_text(encoding='utf-8')
    inject = f'<script>const EMBEDDED_CENTERS={json.dumps(centers, ensure_ascii=False)};'
    inject += f'const EMBEDDED_DATA={json.dumps(embedded_data, ensure_ascii=False)};</script>'
    html = html.replace('</body>', inject + '\n</body>')
    out = ROOT / 'index.html'
    out.write_text(html, encoding='utf-8')
    print(f"[OK] index.html 임베딩 완료 ({len(html):,} bytes)")


def main():
    print("=" * 50)
    print("매일아침 — 데이터 빌드")
    print("=" * 50)
    DATA_DIR.mkdir(exist_ok=True)

    results = []

    # uploads 폴더 처리
    groups = scan_uploads()
    if groups:
        for key in sorted(groups.keys()):
            print(f"\n[{key}] 처리 중...")
            r = process_upload_group(key, groups[key])
            if r:
                results.append(r)
    else:
        print("\n[uploads] 비어있음")

    # 기존 input/{center_id}/ 하위호환
    for cd in sorted(INPUT_DIR.iterdir()):
        if not cd.is_dir() or cd.name.startswith('.') or cd.name in ('uploads', 'raw'):
            continue
        allergy = cd / 'allergy.pdf'
        recipe = cd / 'recipe.xlsx'
        config_path = cd / 'config.json'
        if not allergy.exists():
            continue
        config = json.loads(config_path.read_text(encoding='utf-8')) if config_path.exists() else {"id": cd.name, "name": cd.name, "region": "미지정"}
        # 이미 uploads에서 같은 지역(ID)이 처리됐으면 스킵
        region_id_from_config = REGION_MAP.get(config.get('region', ''), config['id'])
        already = any(r['id'] == region_id_from_config for r in results)
        if already:
            print(f"\n[{cd.name}] 스킵 (uploads에서 이미 처리됨)")
            continue
        print(f"\n[{cd.name}] 처리 중...")
        if recipe.exists():
            data = _process_excel_pdf(recipe, allergy)
        else:
            print(f"  [PDF-only 모드]")
            data = _process_pdf_only(allergy)
        if data:
            year, month = data['year'], data['month']
            month_key = f"{year}-{month:02d}"
            out_dir = DATA_DIR / config['id']
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{month_key}.json"
            out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
            print(f"  → {out_path.relative_to(ROOT)}")
            results.append({"id": config['id'], "name": config['name'], "region": config['region'], "month_key": month_key})

    if results:
        build_centers_json(results)
        build_embedded_html()
        print(f"\n{'=' * 50}")
        print(f"완료! {len(set(r['id'] for r in results))}개 지역 데이터 생성됨")
    else:
        print("\n처리된 데이터가 없습니다.")


if __name__ == '__main__':
    main()
