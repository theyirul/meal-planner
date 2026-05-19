#!/usr/bin/env python3
"""
매일아침 — 식단표 이미지 주차별 크롭 유틸 (검토용)

작은 식단표 이미지(전체 800-1200px급)는 한 번에 OCR하면 텍스트가 작아 정확도가 떨어진다.
주차별 horizontal strip으로 나누면 동일 픽셀에서 디테일이 더 살아남.

사용법:
  python3 scripts/crop_for_review.py <image_path> [--out <dir>] [--rows y1 y2 y3 ...]

기본 전략: 헤더 자동 추정 후 4-5개 균등 분할. 정확한 y 라인을 알면 --rows로 명시.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image


def crop_weeks(
    image_path: str | Path,
    week_y_offsets: list[tuple[int, int]],
    output_dir: str | Path,
    prefix: str | None = None,
) -> list[Path]:
    """이미지를 주차별 horizontal strip으로 크롭.

    week_y_offsets: [(y_start, y_end), ...] 각 주의 y 범위.
    output_dir: 결과 저장 디렉터리 (없으면 생성).
    prefix: 출력 파일명 접두어 (기본: 입력 파일 stem).
    """
    image_path = Path(image_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    img = Image.open(image_path)
    w, _ = img.size
    base = prefix or image_path.stem

    paths: list[Path] = []
    for i, (y1, y2) in enumerate(week_y_offsets, start=1):
        out = output_dir / f"{base}_w{i}.jpg"
        img.crop((0, y1, w, y2)).save(out, quality=95)
        paths.append(out)
    return paths


def auto_split(image_path: str | Path, n_weeks: int, header_ratio: float = 0.2) -> list[tuple[int, int]]:
    """헤더 비율 추정 후 n_weeks 균등 분할.
    예: 1178px 이미지, header_ratio=0.2 → 헤더 236px, 나머지 942px / n_weeks.
    실제 식단표마다 다르므로 결과를 본 뒤 명시 좌표로 조정 권장.
    """
    img = Image.open(image_path)
    _, h = img.size
    header = int(h * header_ratio)
    body = h - header
    step = body // n_weeks
    return [(header + i * step, header + (i + 1) * step) for i in range(n_weeks)]


def main():
    p = argparse.ArgumentParser()
    p.add_argument("image", help="식단표 이미지 경로")
    p.add_argument("--out", default=None, help="출력 디렉터리 (기본: 입력파일과 같은 폴더의 _crops/)")
    p.add_argument("--weeks", type=int, default=4, help="주 개수 (auto split 시)")
    p.add_argument("--header-ratio", type=float, default=0.2, help="헤더가 차지하는 비율 (auto split)")
    p.add_argument("--rows", nargs="+", type=int, help="명시적 y 라인. 예: --rows 235 445 645 850 1100")
    args = p.parse_args()

    image_path = Path(args.image)
    if not image_path.exists():
        print(f"[ERROR] not found: {image_path}", file=sys.stderr)
        sys.exit(1)

    out_dir = Path(args.out) if args.out else image_path.parent / "_crops"

    if args.rows:
        offsets = list(zip(args.rows[:-1], args.rows[1:]))
    else:
        offsets = auto_split(image_path, args.weeks, args.header_ratio)

    paths = crop_weeks(image_path, offsets, out_dir)
    img = Image.open(image_path)
    print(f"입력: {image_path.name} ({img.size[0]}x{img.size[1]})")
    print(f"출력: {out_dir}")
    for i, (yr, path) in enumerate(zip(offsets, paths), start=1):
        print(f"  w{i}: y={yr[0]}~{yr[1]} → {path.name}")


if __name__ == "__main__":
    main()
