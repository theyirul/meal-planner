# 매일아침 (MealMorning)

## 프로젝트 개요
알레르기를 보유한 아이를 키우는 부모가 어린이집/학교 식단을 확인하고,
먹을 수 있는 것/없는 것을 체크해서 공유하는 서비스.

핵심 플로우: 매일 아침 → 오늘 식단 확인 → 알레르기 항목 자동 체크 → 복사 → 메신저로 전송

## 작업 원칙
1. **구현 전 반드시 방향 합의** — 바로 만들지 말고, 어떻게 할지 먼저 이야기하고 합의.
2. 실행 지시가 아니라 고민으로 보이면, 의견으로 답변해줘.
3. 중간중간 "Julie Zhuo critique" 역할로 UX 관점 개선 의견을 남겨줘.

## 빌드 구조
- 소스 템플릿: `index_server.html`
- 빌드 결과: `index.html` (build.py가 데이터 주입해서 생성)
- HTML/JS 수정 시 반드시 `index_server.html`을 수정해야 함. `index.html`만 고치면 다음 빌드 때 덮어씌워짐.

## 데이터 SoT와 빌드 모드
- `uploads/`는 매달 갈아끼우는 **임시 폴더** (gitignore됨). 이번 달 새로 받은 PDF/JSON만 들어감.
- 누적 데이터의 **SoT는 `index.html`의 EMBEDDED_DATA** (git에 커밋됨).
- `build.py`는 **merge 모드 기본**: 기존 `index.html`의 EMBEDDED를 로드 → 신규 uploads 결과만 갈아끼움 → 다시 임베딩.
- 같은 `region_id/YYYY-MM` 키가 충돌하면 새 데이터가 덮어씀.
- 처음부터 다시 빌드하려면 `python3 scripts/build.py --clean` (기존 index.html 무시).

## 원본 보정 (패치 JSON)
- 파싱 결과를 손으로 고칠 땐 `uploads/{YYYYMM}_{지역}_{연령}_{유형}_패치.json`. 연산: `replace`(이름) / `remove` / `add` / **`allergy`(번호만 병합, 이름·레시피 보존)**.
- 패치는 세 입력 경로(수동 JSON·HWP·PDF) 전부에서 적용된다(`_apply_patches`). 파일 안에 `_근거`를 남겨 왜 고쳤는지 추적 가능하게 할 것.
- **알레르기 판단 기준: 조리지시서 식재료에 근거가 있으면 식단표 원본이 미표기여도 보강한다** (룰 지시 2026-08-03). 안전 > 원본 충실. 근거 없는 추정으로는 넣지 않는다.

## 입력 형식 우선순위
1. **PDF (강력 권장)** — 텍스트 레이어가 있어 `pdfplumber`로 거의 100% 정확 파싱. `parse_pdf.py`·`parse_table.py`가 처리.
2. **HWP** — `parse_hwp.py`(pyhwp→HTML→병합셀 그리드 파싱). `.hwp` 식단표를 규칙 파일명으로 uploads에 넣으면 build가 자동 처리. **병합셀(colspan/rowspan) 때문에 순서 인덱스 매핑 금지 — 반드시 그리드 좌표로.** (창원시 등)
3. xlsx — 레시피 매칭용으로 함께 받으면 좋음. **조리지시서 형식(영양소분석표+주차별 지시서)이면 build가 자동으로 `crosscheck.py` 실행** → 배정 대조 + 알레르기 누락 flag. **조리지시서 있으면 항상 함께 교차검증**(룰 지침).
4. 이미지 (JPG/PNG) — **OCR 오류 불가피. 사람 검증 단계 필수.** 800px급 저해상도는 어휘 다양한 어린이집 메뉴에서 60-70% 정확도 (2026-05-19 PoC). 가능하면 베타테스터·셀프 업로드 가이드에서 "PDF로 부탁드려요"를 디폴트로.

## OCR 워크플로 (이미지 입력 시)
- `scripts/crop_for_review.py`로 주차별 horizontal strip 크롭 → 클로드 multimodal vision 또는 외부 OCR로 셀별 추출.
- 결과를 `uploads/{YYYYMM}_{지역}_{연령}_{유형}_수동.json` 형식으로 저장하면 build.py가 파서 우회하고 그대로 임베딩.
- 알레르기 번호는 잘못 입력 시 아이 건강 직결 → **사람 검증 전까지 라이브에 안 올림**.

## 배포
- GitHub Pages (theyirul/meal-planner, main 브랜치)
- Claude Code CLI에서 git push 가능 (확인됨 2026-05-10)

## 세션 구조
- **공동 창업자 세션** — 제품 기획, 개발, 베타테스터 대응, 운영 전반. 프로젝트의 메인 세션.
- 특정 업무가 반복적으로 커지면 그때 세션을 분리한다.
