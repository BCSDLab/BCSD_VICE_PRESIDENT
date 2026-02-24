# BCSD 부회장 자동화 도구

## 실행 환경

- Python 3.10+
- HWP 증빙 서류 생성 기능(`pyhwpx`)은 **Windows 전용**입니다.

## 설치

```bash
pip install -r requirements.txt
```

---

## 1. 장부 자동화 (`main.py`)

재학생 회비 관리 문서를 파싱하여 기간별 장부 Excel을 생성하고, 지출 내역에 대한 HWP 증빙 서류를 자동으로 만들어줍니다.

### 환경 설정

`.env.example`을 복사하여 `.env`를 생성한 후 값을 채웁니다.

```bash
cp .env.example .env
```

| 변수 | 설명 |
|---|---|
| `DEBUG` | `True`면 HWP 창을 화면에 표시 |
| `GOOGLE_SECRET_JSON` | Google 서비스 계정 JSON 파일 경로 |

> Google 서비스 계정에는 Google Docs API 및 Google Drive API 읽기 권한이 필요합니다.

### 사용법

```bash
python main.py <파일경로> <시작기간> <종료기간>
```

**예시 — 2025년 11월 ~ 2026년 2월 장부 생성:**

```bash
python main.py "재학생 회비 관리 문서_20260225.xlsx" 2025-11 2026-02
```

### 출력 결과

`output/` 디렉토리에 두 파일이 생성됩니다.

| 파일 | 설명 |
|---|---|
| `BCSD_YYYYMM_YYYYMM_장부.xlsx` | 기간별 장부 (순서 / 날짜 / 종류 / 내용 / 금액 / 잔액) |
| `BCSD_YYYYMM_YYYYMM_증빙자료.hwp(x)` | 지출 항목별 증빙 이미지가 삽입된 HWP 서류 |

---

## 2. 회비 미납자 확인 (`fee_check.py`)

재학생 회비 납부 문서에서 미납자를 집계하고, 개인별 알림 메시지 파일을 생성합니다.

### 사용법

```bash
python fee_check.py [-e <트랙명>]
```

실행 시 `재학생 회비 납부 문서_*.xlsx` 파일 중 가장 최신 날짜 파일을 자동으로 선택합니다.

**예시 — 전체 대상:**

```bash
python fee_check.py
```

**예시 — 특정 트랙 제외:**

```bash
python fee_check.py -e "Back-End" -e "Android"
python fee_check.py -e "Back-End,Android,Design"
```

### 출력 결과

`output/YYYY-MM/` 디렉토리에 미납자별 `.txt` 파일이 생성됩니다.

---

## 프로젝트 구조

```
├── main.py                        # 장부 자동화 진입점
├── fee_check.py                   # 회비 미납자 확인 진입점
├── ledger/
│   ├── membership_fee_parser.py   # 재학생 회비 관리 문서 → 장부 변환
│   └── hwp/
│       ├── hwp_generator.py       # HWP 증빙 서류 생성
│       ├── image_downloader.py    # Google Drive 이미지 다운로드
│       └── image_packer.py        # HWP 테이블 셀 이미지 배치
├── fee_checker/
│   └── checker.py                 # 회비 미납자 확인 로직
└── templates/
    └── fee_notice.md              # 미납 알림 메시지 템플릿 (한글·엑셀 파일은 gitignore)
```
