#!/usr/bin/env python3
"""
매일아침 — 데이터 빌드 스크립트

uploads/ 폴더의 파일을 읽어서 index.html에 임베딩합니다.

파일 네이밍: {YYYYMM}_{지역}_{연령}_{유형}_{타입}.{확장자}
예: 202603_광진구_1-2세_일반형_식단표.pdf

사용법:
  python3 scripts/build.py
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path
from collections import defaultdict

ROOT = Path(__file__).resolve().parent.parent
UPLOAD_DIR = ROOT / 'uploads'
SCRIPTS_DIR = ROOT / 'scripts'

sys.path.insert(0, str(SCRIPTS_DIR))
from generate_menu import extract_menus, detect_month_year, extract_allergies, match_allergies, analyze_sauce, build_json_data
from parse_pdf import parse_pdf_menu, build_json_from_pdf
from parse_table import detect_table_format, parse_table_pdf, parse_recipe_xlsx, parse_recipe_pdf, _parse_vertical_pdf, _parse_school_pdf
from parse_common import build_output_json
from parse_hwp import parse_hwp_menu

# 지역명 → ID 매핑
REGION_MAP = {
    "광진구": "gwangjin",
    "영등포구": "yeongdeungpo",
    "수원시": "suwon",
    "김포시": "gimpo",
    "용인시": "yongin",
    "경기도용인시": "yongin",
    "경기도의왕시": "uiwang",
    "경기도성남시": "seongnam",
    "성남시": "seongnam",
    "동대문구": "dongdaemun",
    "강동구": "gangdong",
    "충북옥천군": "okcheon",
    "대전중구": "daejeon-junggu",
    "서울시강서구": "gangseo",
    "의왕시": "uiwang",
    "창원시": "changwon",
    "대구서구": "daegu-seogu",
    "미소시루": "misosiru",
}

REGION_NAMES = {v: k for k, v in REGION_MAP.items()}

# 표시명 (UI에 보여줄 이름)
REGION_DISPLAY = {
    "광진구": "서울 광진구",
    "영등포구": "서울 영등포구",
    "수원시": "경기 수원시",
    "김포시": "경기 김포시",
    "용인시": "경기 용인시",
    "경기도용인시": "경기 용인시",
    "경기도의왕시": "경기 의왕시",
    "경기도성남시": "경기 성남시",
    "성남시": "경기 성남시",
    "동대문구": "서울 동대문구",
    "강동구": "서울 강동구 강솔초등학교",
    "충북옥천군": "충북 옥천군",
    "대전중구": "대전 중구",
    "서울시강서구": "서울 강서구",
    "의왕시": "경기 의왕시",
    "창원시": "경남 창원시",
    "대구서구": "대구 서구 (구립화성파크드림어린이집)",
    "미소시루": "미소시루 유치원",
}


def parse_upload_filename(filename: str) -> dict | None:
    """파일명에서 메타데이터 추출"""
    filename = unicodedata.normalize('NFC', filename)
    name = Path(filename).stem
    ext = Path(filename).suffix.lstrip('.')
    # 수동/패치 JSON: {YYYYMM}_{지역}_{연령}_{유형}_{수동|패치}.json
    special_pat = r'^(\d{6})_([가-힣]+)_(.+?)_(.+?)_(수동|패치)$'
    mm = re.match(special_pat, name)
    if mm and ext == 'json':
        return {
            "yyyymm": mm.group(1), "region": mm.group(2),
            "age": mm.group(3), "type": mm.group(4),
            "subtype": mm.group(5), "ext": ext, "filename": filename,
        }
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
    """uploads/ 폴더 + archives/ 서브폴더 스캔 → 지역+월별 그룹화"""
    if not UPLOAD_DIR.exists():
        return {}
    groups = defaultdict(list)
    scan_dirs = [UPLOAD_DIR, UPLOAD_DIR / 'archives']
    for scan_dir in scan_dirs:
        if not scan_dir.exists():
            continue
        for fpath in scan_dir.iterdir():
            if not fpath.is_file():
                continue
            meta = parse_upload_filename(fpath.name)
            if meta:
                meta['filename'] = str(fpath)  # 절대 경로로 업데이트
                key = f"{meta['yyyymm']}_{meta['region']}"
                groups[key].append(meta)
    return dict(groups)


def _process_table_format(fmt: str, pdf_file: Path, recipe_file: Path | None = None) -> dict | None:
    """통합 테이블 파서로 처리 (영등포구, 성남시, 동대문구 등)"""
    # 세로형 포맷: 메뉴+레시피가 한 PDF에 모두 포함
    if fmt == "dongdaemun_vertical":
        pdf_data, recipes = _parse_vertical_pdf(str(pdf_file))
        return build_output_json(pdf_data['year'], pdf_data['month'], pdf_data['menus'], recipes)

    # 학교 급식: 레시피 없는 단일 PDF
    if fmt == "school":
        pdf_data = _parse_school_pdf(str(pdf_file))
        return build_output_json(pdf_data['year'], pdf_data['month'], pdf_data['menus'])

    pdf_data = parse_table_pdf(str(pdf_file), fmt)
    recipes = {}
    if recipe_file and recipe_file.exists():
        ext = recipe_file.suffix.lower()
        try:
            if ext == '.xlsx':
                recipes = parse_recipe_xlsx(str(recipe_file))
            elif ext == '.pdf':
                recipes = parse_recipe_pdf(str(recipe_file))
        except Exception as e:
            print(f"  [WARN] 레시피 파싱 실패: {e}")
    return build_output_json(pdf_data['year'], pdf_data['month'], pdf_data['menus'], recipes)


def _process_standard(pdf_file: Path, xlsx_file: Path | None = None) -> dict | None:
    """광진구/수원 표준 포맷 처리"""
    if xlsx_file and xlsx_file.exists():
        print(f"  [Excel+PDF] {xlsx_file.name} + {pdf_file.name}")
        try:
            menus = extract_menus(str(xlsx_file))
            year, month = detect_month_year(menus)
            print(f"  감지: {year}년 {month}월")
            allergy_data = extract_allergies(str(pdf_file), list(menus.keys()))
            allergy_match = match_allergies(menus, allergy_data)
            sauce_data = analyze_sauce(menus, allergy_match)
            return build_json_data(menus, allergy_match, sauce_data, year, month)
        except Exception as e:
            print(f"  [WARN] Excel 파싱 실패 ({e}) → PDF-only로 전환")

    print(f"  [PDF-only] {pdf_file.name}")
    try:
        pdf_data = parse_pdf_menu(str(pdf_file))
        return build_json_from_pdf(pdf_data)
    except Exception as e:
        print(f"  [ERROR] PDF 파싱 실패: {e}")
        return None


def _apply_patch(data: dict, patch_path: Path) -> dict:
    """패치 JSON을 파싱 결과에 적용. replace/add/remove 지원."""
    with open(patch_path, encoding='utf-8') as f:
        patch = json.load(f)

    menus = data.get('menus', {})
    applied = 0
    for date_str, ops in patch.items():
        if date_str.startswith('_'):  # _설명, _사용법 등 메타 필드 스킵
            continue
        if date_str not in menus:
            print(f"    [패치] {date_str} — 해당 날짜 없음, 스킵")
            continue
        items = menus[date_str]

        # replace: {"현재이름": "새이름"} — 이름만 교체
        for old_name, new_name in ops.get('replace', {}).items():
            for it in items:
                if it['name'] == old_name:
                    it['name'] = new_name
                    applied += 1
                    break

        # remove: ["이름1", "이름2"] — 해당 아이템 제거
        for rm_name in ops.get('remove', []):
            before = len(items)
            items = [it for it in items if it['name'] != rm_name]
            if len(items) < before:
                applied += 1
            menus[date_str] = items

        # add: [{"name": "...", "allergy": [...], "sec": "..."}] — 아이템 추가
        for new_item in ops.get('add', []):
            items.append(new_item)
            applied += 1
            menus[date_str] = items

    print(f"    [패치] {applied}건 적용 완료")
    data['menus'] = menus
    return data


def process_upload_group(group_key: str, files: list[dict]) -> tuple[dict | None, dict | None]:
    """uploads 그룹 처리. (메타정보, 파싱데이터) 반환"""
    meta = files[0]
    region_name = meta['region']
    region_id = REGION_MAP.get(region_name)
    if not region_id:
        print(f"  [SKIP] 지역 매핑 없음: {region_name} → REGION_MAP에 추가 필요")
        return None, None

    # 수동 JSON 파일 체크
    for f in files:
        if f['subtype'] == '수동' and f['ext'] == 'json':
            json_path = UPLOAD_DIR / f['filename']
            if json_path.exists():
                print(f"  [수동 JSON] {f['filename']}")
                with open(json_path, encoding='utf-8') as jf:
                    data = json.load(jf)
                # 수동 JSON엔 레시피가 없으므로, 같은 그룹에 레시피 파일이 있으면 함께 파싱
                recipes = {}
                for rf in files:
                    if rf['subtype'] == '레시피':
                        rpath = UPLOAD_DIR / rf['filename']
                        if rpath.exists():
                            ext = rpath.suffix.lower()
                            try:
                                if ext == '.xlsx':
                                    recipes = parse_recipe_xlsx(str(rpath))
                                elif ext == '.pdf':
                                    recipes = parse_recipe_pdf(str(rpath))
                            except Exception as e:
                                print(f"  [WARN] 레시피 파싱 실패: {e}")
                # build_output_json 형식으로 변환
                data = build_output_json(data['year'], data['month'], data['menus'], recipes)
                year, month = data['year'], data['month']
                month_key = f"{year}-{month:02d}"
                display_name = REGION_DISPLAY.get(region_name, region_name)
                result = {"id": region_id, "name": display_name, "region": display_name, "month_key": month_key}
                return result, data

    # HWP 식단표 (창원시 등 한글 파일) → parse_hwp로 그리드 파싱
    for f in files:
        if f['subtype'] == '식단표' and f['ext'].lower() == 'hwp':
            hwp_path = UPLOAD_DIR / f['filename']
            if hwp_path.exists():
                print(f"  [HWP 식단표] {f['filename']}")
                yyyymm = f['yyyymm']
                pdata = parse_hwp_menu(str(hwp_path), int(yyyymm[:4]), int(yyyymm[4:6]), f['age'])
                recipes = {}
                for rf in files:
                    if rf['subtype'] == '레시피':
                        rpath = UPLOAD_DIR / rf['filename']
                        if rpath.exists():
                            ext = rpath.suffix.lower()
                            try:
                                if ext == '.xlsx':
                                    recipes = parse_recipe_xlsx(str(rpath))
                                elif ext == '.pdf':
                                    recipes = parse_recipe_pdf(str(rpath))
                            except Exception as e:
                                print(f"  [WARN] 레시피 파싱 실패: {e}")
                data = build_output_json(pdata['year'], pdata['month'], pdata['menus'], recipes)
                month_key = f"{data['year']}-{data['month']:02d}"
                display_name = REGION_DISPLAY.get(region_name, region_name)
                result = {"id": region_id, "name": display_name, "region": display_name, "month_key": month_key}
                return result, data

    # 파일 분류 (식단표 PDF + 레시피 xlsx/pdf)
    pdf_file = None
    recipe_file = None
    for f in files:
        fpath = UPLOAD_DIR / f['filename']
        if f['subtype'] == '식단표' and f['ext'].lower() == 'pdf':
            pdf_file = fpath
        elif f['subtype'] == '레시피':
            recipe_file = fpath

    if not pdf_file or not pdf_file.exists():
        print(f"  [SKIP] 식단표 PDF 없음")
        return None, None

    # 포맷 자동 감지: 테이블 포맷이면 통합 파서, 아니면 표준 파서
    fmt = detect_table_format(str(pdf_file))

    if fmt:
        print(f"  [포맷 감지] {fmt}")
        data = _process_table_format(fmt, pdf_file, recipe_file)
    else:
        print(f"  [포맷 감지] standard")
        xlsx_file = recipe_file if recipe_file and recipe_file.suffix.lower() == '.xlsx' else None
        data = _process_standard(pdf_file, xlsx_file)

    if data is None:
        return None, None

    # 패치 JSON 적용 (있으면)
    for f in files:
        if f['subtype'] == '패치' and f['ext'] == 'json':
            patch_path = UPLOAD_DIR / f['filename']
            if patch_path.exists():
                print(f"  [패치 적용] {f['filename']}")
                data = _apply_patch(data, patch_path)

    year, month = data['year'], data['month']
    month_key = f"{year}-{month:02d}"
    display_name = REGION_DISPLAY.get(region_name, region_name)
    result = {"id": region_id, "name": display_name, "region": display_name, "month_key": month_key}
    return result, data


def load_existing_embedded() -> tuple[list[dict], dict]:
    """현재 index.html에서 EMBEDDED_CENTERS·EMBEDDED_DATA 추출.
    uploads/는 매달 갈아끼우는 임시 폴더 — 과거 데이터의 SoT는 index.html.
    빌드는 항상 기존 index.html에 신규 uploads 결과를 머지한다.
    """
    out_path = ROOT / 'index.html'
    if not out_path.exists():
        return [], {}
    html = out_path.read_text(encoding='utf-8')
    m = re.search(r'const EMBEDDED_CENTERS=(\[.*?\]);const EMBEDDED_DATA=(\{.*?\});', html, re.DOTALL)
    if not m:
        return [], {}
    try:
        return json.loads(m.group(1)), json.loads(m.group(2))
    except json.JSONDecodeError as e:
        print(f"[WARN] 기존 index.html EMBEDDED 파싱 실패: {e} → 빈 상태에서 시작")
        return [], {}


def build_embedded_html(centers: list[dict], all_data: dict):
    """index.html에 데이터 직접 임베딩 (중간 파일 없이)"""
    index_src = ROOT / 'index_server.html'
    if not index_src.exists():
        print("[WARN] index_server.html 없음 → 스킵")
        return

    html = index_src.read_text(encoding='utf-8')
    inject = f'<script>const EMBEDDED_CENTERS={json.dumps(centers, ensure_ascii=False)};'
    inject += f'const EMBEDDED_DATA={json.dumps(all_data, ensure_ascii=False)};</script>'
    html = html.replace('</body>', inject + '\n</body>')
    out = ROOT / 'index.html'
    out.write_text(html, encoding='utf-8')
    print(f"[OK] index.html 임베딩 완료 ({len(html):,} bytes)")


def main():
    print("=" * 50)
    print("매일아침 — 데이터 빌드")
    print("=" * 50)

    clean = '--clean' in sys.argv
    if clean:
        print("[--clean] 기존 index.html 무시, uploads/만으로 빌드")
        old_centers, old_data = [], {}
    else:
        old_centers, old_data = load_existing_embedded()
        if old_data:
            print(f"[merge 모드] 기존 index.html에서 {len(old_centers)}지역 / {len(old_data)} month키 로드")

    results = []
    all_data = dict(old_data)

    groups = scan_uploads()
    if groups:
        for key in sorted(groups.keys()):
            print(f"\n[{key}] 처리 중...")
            result, data = process_upload_group(key, groups[key])
            if result and data:
                results.append(result)
                key_str = f"{result['id']}/{result['month_key']}"
                if key_str in all_data:
                    print(f"  [덮어씀] {key_str} (기존 데이터 위에 새 데이터)")
                all_data[key_str] = data
    else:
        print("\n[uploads] 비어있음")

    if not results and not old_data:
        print("\n처리된 데이터가 없습니다.")
        return

    # centers 머지: 기존 + 신규 결과
    merged = {c['id']: dict(c) for c in old_centers}
    for r in results:
        rid = r['id']
        if rid in merged:
            months = set(merged[rid].get('months', []))
            months.add(r['month_key'])
            merged[rid]['months'] = sorted(months)
            # 표시명 새 빌드 결과로 갱신 (REGION_DISPLAY 변경 반영)
            merged[rid]['name'] = r['name']
            merged[rid]['region'] = r['region']
        else:
            merged[rid] = {
                "id": rid, "name": r['name'],
                "region": r['region'], "months": [r['month_key']]
            }
    centers = sorted(merged.values(), key=lambda c: c['name'])
    build_embedded_html(centers, all_data)
    print(f"\n{'=' * 50}")
    print(f"완료! {len(centers)}개 지역 / {len(all_data)} month키 데이터 임베딩")
    if results:
        new_keys = [f"{r['id']}/{r['month_key']}" for r in results]
        print(f"이번 빌드 신규/갱신: {', '.join(new_keys)}")


if __name__ == '__main__':
    main()
