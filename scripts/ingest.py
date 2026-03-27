#!/usr/bin/env python3
"""
매일아침 — 식단 파일 자동 처리 (ingest)

uploads/ 폴더에 아무 이름으로 PDF/xlsx를 넣고 실행하면:
1. 컨벤션에 안 맞는 파일 탐지
2. PDF 내용에서 지역, 연월 자동 감지
3. 파일명 자동 변경
4. build.py 자동 실행

사용법:
  python scripts/ingest.py
"""
from __future__ import annotations

import os
import re
import sys
import shutil
import unicodedata
from pathlib import Path

import pdfplumber

ROOT = Path(__file__).resolve().parent.parent
UPLOAD_DIR = ROOT / 'uploads'
SCRIPTS_DIR = ROOT / 'scripts'

sys.path.insert(0, str(SCRIPTS_DIR))
from build import parse_upload_filename, REGION_MAP, REGION_DISPLAY

# PDF 텍스트에서 지역을 감지하는 키워드 매핑
# key: PDF에서 찾을 키워드, value: 파일명에 쓸 지역명
REGION_KEYWORDS = {
    "강서구": "서울시강서구",
    "광진구": "광진구",
    "영등포구": "영등포구",
    "동대문구": "동대문구",
    "수원": "수원시",
    "성남": "경기도성남시",
    "용인": "경기도용인시",
    "옥천": "충북옥천군",
    "대전": "대전중구",
}

# 연령 감지 패턴
AGE_PATTERNS = [
    (r'만\s*1[~\-]2세|1-2세|1~2세|13[~\-]35개월', "1-2세"),
    (r'만\s*3[~\-]5세|3-5세|3~5세', "3-5세"),
]


def detect_region_from_pdf(pdf_path: str) -> str | None:
    """PDF 텍스트에서 지역명 감지"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''
            text = text[:500]  # 상단 500자만
    except Exception:
        return None

    for keyword, region_name in REGION_KEYWORDS.items():
        if keyword in text:
            return region_name
    return None


def detect_month_from_pdf(pdf_path: str) -> str | None:
    """PDF 텍스트에서 연월 감지 → 'YYYYMM' 형식"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''
            text = text[:500]
    except Exception:
        return None

    # "2026년 3월" or "2026년3월" or "2026. 3." 등
    m = re.search(r'(\d{4})\s*년\s*(\d{1,2})\s*월', text)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}"

    m = re.search(r'(\d{4})\.\s*(\d{1,2})\.', text)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}"

    return None


def detect_age_from_pdf(pdf_path: str) -> str:
    """PDF 텍스트에서 연령 감지 → '1-2세' 등. 기본값 '1-2세'"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''
            text = text[:500]
    except Exception:
        return "1-2세"

    for pattern, age in AGE_PATTERNS:
        if re.search(pattern, text):
            return age
    return "1-2세"


def detect_file_type(filename: str) -> str:
    """파일 확장자 + 이름에서 식단표/레시피 구분"""
    name_lower = filename.lower()
    if '레시피' in name_lower or 'recipe' in name_lower:
        return "레시피"
    ext = Path(filename).suffix.lower()
    if ext == '.xlsx':
        return "레시피"
    return "식단표"


def find_unprocessed_files() -> list[Path]:
    """컨벤션에 안 맞는 파일 찾기"""
    if not UPLOAD_DIR.exists():
        return []

    unprocessed = []
    for fpath in UPLOAD_DIR.iterdir():
        if not fpath.is_file():
            continue
        ext = fpath.suffix.lower()
        if ext not in ('.pdf', '.xlsx', '.json'):
            continue
        # 이미 컨벤션에 맞는 파일은 스킵
        meta = parse_upload_filename(fpath.name)
        if meta:
            continue
        unprocessed.append(fpath)
    return unprocessed


def generate_filename(region: str, yyyymm: str, age: str, file_type: str, ext: str) -> str:
    """컨벤션에 맞는 파일명 생성"""
    return f"{yyyymm}_{region}_{age}_일반형_{file_type}{ext}"


def process_file(fpath: Path) -> str | None:
    """파일 하나를 분석하고 이름 변경. 성공하면 새 파일명 반환."""
    ext = fpath.suffix.lower()
    name = fpath.name

    print(f"\n📄 {name}")

    if ext == '.pdf':
        region = detect_region_from_pdf(str(fpath))
        yyyymm = detect_month_from_pdf(str(fpath))
        age = detect_age_from_pdf(str(fpath))
        file_type = detect_file_type(name)

        if not region:
            print(f"  ❌ 지역을 감지할 수 없습니다. 파일명에 지역을 포함해주세요.")
            return None
        if not yyyymm:
            print(f"  ❌ 연월을 감지할 수 없습니다.")
            return None

        print(f"  → 지역: {region}, 연월: {yyyymm}, 연령: {age}, 유형: {file_type}")

    elif ext == '.xlsx':
        # xlsx는 PDF와 같은 지역/월의 레시피라고 가정
        # 파일명에서 힌트 찾기
        region = None
        yyyymm = None
        age = "1-2세"
        file_type = "레시피"

        # 파일명에서 월 추출 시도
        m = re.search(r'(\d{1,2})월', name)
        if m:
            month = int(m.group(1))
            # 연도는 같은 폴더의 PDF에서 추론
            yyyymm_candidates = set()
            for other in UPLOAD_DIR.iterdir():
                if other.suffix.lower() == '.pdf':
                    meta = parse_upload_filename(other.name)
                    if meta and int(meta['yyyymm'][4:]) == month:
                        yyyymm_candidates.add(meta['yyyymm'])
                        if not region:
                            region = meta['region']
            if yyyymm_candidates:
                yyyymm = sorted(yyyymm_candidates)[0]

        # 파일명에서 연령 추출
        for pattern, age_val in AGE_PATTERNS:
            if re.search(pattern, name):
                age = age_val
                break

        if not region or not yyyymm:
            # 같은 폴더에서 가장 최근 PDF 기반으로 추론
            pdfs = sorted(
                [f for f in UPLOAD_DIR.iterdir() if f.suffix.lower() == '.pdf'],
                key=lambda f: f.stat().st_mtime,
                reverse=True,
            )
            for pdf in pdfs:
                meta = parse_upload_filename(pdf.name)
                if meta:
                    if not region:
                        region = meta['region']
                    if not yyyymm:
                        yyyymm = meta['yyyymm']
                    break

        if not region or not yyyymm:
            print(f"  ❌ 지역/연월을 추론할 수 없습니다.")
            return None

        print(f"  → 지역: {region}, 연월: {yyyymm}, 연령: {age}, 유형: {file_type}")

    else:
        return None

    new_name = generate_filename(region, yyyymm, age, file_type, ext)
    new_path = UPLOAD_DIR / new_name

    if new_path.exists() and new_path != fpath:
        print(f"  ⚠️  {new_name} 이미 존재합니다. 덮어쓸까요? (y/n) ", end="")
        answer = input().strip().lower()
        if answer != 'y':
            print(f"  ⏭️  건너뜀")
            return None

    if fpath != new_path:
        shutil.move(str(fpath), str(new_path))
        print(f"  ✅ → {new_name}")
    else:
        print(f"  ✅ 이름 변경 불필요")

    return new_name


def main():
    print("=" * 50)
    print("매일아침 — 식단 파일 자동 처리")
    print("=" * 50)

    unprocessed = find_unprocessed_files()

    if not unprocessed:
        print("\n✅ 처리할 새 파일이 없습니다.")
        print("   (컨벤션에 맞지 않는 파일만 처리합니다)")
    else:
        print(f"\n📦 처리할 파일: {len(unprocessed)}개")

        success = 0
        for fpath in sorted(unprocessed):
            result = process_file(fpath)
            if result:
                success += 1

        print(f"\n{'=' * 50}")
        print(f"파일 처리 완료: {success}/{len(unprocessed)}개 성공")

    # build.py 실행
    print(f"\n{'=' * 50}")
    print("빌드 실행 중...")
    print("=" * 50)
    os.chdir(ROOT)
    exit_code = os.system(f"{sys.executable} scripts/build.py")

    if exit_code == 0:
        print("\n🎉 완료!")
    else:
        print(f"\n❌ 빌드 실패 (exit code: {exit_code})")
        sys.exit(1)


if __name__ == "__main__":
    main()
