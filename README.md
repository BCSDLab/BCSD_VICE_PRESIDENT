# BCSD 부회장 자동화 도구

## 실행 환경

- Python 3.10+
- HWP 증빙 서류 생성
  - **Windows**: `pyhwpx`를 사용합니다.
  - **Linux / macOS**: HWPX(ZIP+XML) 직접 조작 방식으로 동작합니다. 템플릿은 `.hwpx` 형식이어야 합니다.

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
| `DEBUG` | `True`면 HWP 창을 화면에 표시 (Windows 전용) |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Google 서비스 계정 키 JSON 경로 (Google Docs/Drive 읽기) |

> Google 서비스 계정에는 Google Docs API 및 Google Drive API 읽기 권한이 필요합니다.

### 사용법

```bash
python main.py <시작기간> <종료기간>
```

**예시 — 2025년 11월 ~ 2026년 2월 장부 생성:**

```bash
python main.py 2025-11 2026-02
```

### 출력 결과

`output/` 디렉토리에 두 파일이 생성됩니다.

| 파일 | 설명 |
|---|---|
| `BCSD_YYYYMM_YYYYMM_장부.xlsx` | 기간별 장부 (순서 / 날짜 / 종류 / 내용 / 금액 / 잔액) |
| `BCSD_YYYYMM_YYYYMM_증빙자료.hwp(x)` | 지출 항목별 증빙 이미지가 삽입된 HWP 서류 |

---

## 2. 거래내역 → 관리 문서 자동 기입 (`fill_ledger.py`)

신한은행 거래내역 파일을 읽어 재학생 회비 관리 문서에 자동으로 기입합니다.

| 구분 | 열 | 처리 방식 |
|---|---|---|
| 날짜 | D열 | **자동** |
| 이름 (거래내역 내용) | F열 | **자동** |
| 입/출 금액 | H열 | **자동** (입금 양수, 출금 음수) |
| 잔액 | I열 | **자동** |
| 내용 (회비 / 서버비 등) | E열 | **수동** |
| 비고 (납부 월 등) | G열 | **수동** |

소계·합계 수식도 행 번호에 맞게 자동 갱신됩니다.

### 환경 설정

| 변수 | 설명 |
|---|---|
| `GOOGLE_OAUTH_CLIENT_JSON` | OAuth 데스크톱 앱 클라이언트 시크릿 JSON 경로 |
| `TRANSACTION_DRIVE_URL` | 신한 거래내역 파일들이 있는 Google Drive 폴더 링크 |
| `MANAGEMENT_SHEET_URL` | 재학생 회비 관리 문서 Google Sheets URL |
| `RECEIPT_DIR` | 영수증 Google Drive 폴더 URL (선택) — `{년}/{월}/` 하위 폴더 구조 |

> 최초 실행 시 브라우저 인증이 열립니다. 이후 `.google_token.json`에 토큰이 캐시되어 재인증 없이 사용 가능합니다.

### 준비

`TRANSACTION_DRIVE_URL`에 신한 거래내역 파일들이 있는 Google Drive 폴더 링크를 설정합니다.
폴더 내 `신한_거래내역_YYMM.xlsx` 중 가장 최신 파일을 자동으로 선택합니다.

### 사용법

```bash
python fill_ledger.py                              # TRANSACTION_DRIVE_URL에서 최신 파일 자동 선택
python fill_ledger.py 신한_거래내역/신한_거래내역_2602.xlsx
python fill_ledger.py --force                      # 이미 기입된 월도 덮어쓰기
```

### 출력 결과

`MANAGEMENT_SHEET_URL`에 지정된 Google Sheets 파일에 직접 업로드됩니다.

---

## 3. 회비 미납자 확인 및 Slack DM 발송 (`fee_check.py`)

재학생 회비 납부 문서에서 미납자를 집계하고, 개인별 알림 메시지 파일을 생성합니다.
`--send-dm` 옵션을 사용하면 미납자에게 Slack DM을 자동으로 발송합니다.

### 회비 구조

| 기간 | 단위 | 금액 |
|---|---|---|
| ~2026년 2월 | 월별 | 10,000원/월 |
| 2026년 3월~ | 학기별 (1학기: 3~8월, 2학기: 9~2월) | 60,000원/학기 |

### 환경 설정

`.env`에 아래 항목을 추가합니다. (SSH 터널을 통해 DB에 **readonly**로 접근합니다.)

| 변수 | 설명 |
|---|---|
| `SLACK_BOT_TOKEN` | Slack Bot Token (`xoxb-...`) |
| `SLACK_SENDER_ID` | 발신자(본인) Slack user_id |
| `SENDER_NAME` | 발신자 이름 (메시지 서명에 사용) |
| `SENDER_PHONE` | 발신자 전화번호 (메시지 문의처에 사용) |
| `FEE_SHEET_URL` | 납부 문서 URL (메시지 내 하이퍼링크에 사용) |
| `SSH_HOST` | SSH 서버 호스트 |
| `SSH_PORT` | SSH 포트 (기본값: `22`) |
| `SSH_USER` | SSH 사용자명 |
| `SSH_KEY_PATH` | SSH 개인키 경로 (없으면 `SSH_PASSWORD` 사용) |
| `SSH_PASSWORD` | SSH 비밀번호 (키 파일 미사용 시) |
| `DB_HOST` | MySQL 호스트 (기본값: `127.0.0.1`) |
| `DB_PORT` | MySQL 포트 (기본값: `3306`) |
| `DB_NAME` | 데이터베이스명 |
| `DB_USER` | DB 사용자명 (readonly 계정 권장) |
| `DB_PASSWORD` | DB 비밀번호 |
| `DB_TABLE` | 회원 테이블명 |
| `DB_COL_NAME` | 이름 컬럼명 (기본값: `name`) |
| `DB_COL_SLACK_ID` | Slack ID 컬럼명 (기본값: `slack_id`) |
| `DB_COL_TRACK_ID` | 트랙 FK 컬럼명 (기본값: `track_id`) |
| `DB_COL_IS_DELETED` | 삭제 여부 컬럼명 (기본값: `is_deleted`) |
| `DB_TRACK_TABLE` | 트랙 테이블명 |
| `DB_TRACK_COL_ID` | 트랙 PK 컬럼명 (기본값: `id`) |
| `DB_TRACK_COL_NAME` | 트랙명 컬럼명 (기본값: `name`) |
| `DB_TRACK_COL_IS_DELETED` | 트랙 삭제 여부 컬럼명 (기본값: `DB_COL_IS_DELETED` 값) |

> Slack Bot에는 `users:read`, `im:write`, `chat:write` 권한이 필요합니다.

### 사용법

```bash
python fee_check.py [-e <트랙명>] [-p <이름_트랙>] [--send-dm]
```

실행 시 `.env`의 `FEE_SHEET_URL`에서 납부 문서를 자동으로 다운로드합니다.

**예시 — 파일 생성만:**

```bash
python fee_check.py
```

**예시 — 특정 트랙 제외:**

```bash
python fee_check.py -e "Back-End" -e "Android"
python fee_check.py -e "Back-End,Android,Design"
```

**예시 — 특정 인원 제외:**

```bash
python fee_check.py -p "홍길동_Android"
python fee_check.py -p "홍길동_Android" -p "김철수_Back-End"
python fee_check.py -p "홍길동_Android,김철수_Back-End"
```

**예시 — 파일 생성 + Slack DM 발송:**

```bash
python fee_check.py --send-dm
```

### 출력 결과

`output/YYYY-MM/` 디렉토리에 미납자별 `.txt` 파일이 생성됩니다.
`--send-dm` 사용 시 각 미납자에게 `templates/fee_notice.md` 내용으로 DM이 발송됩니다.

미납 내역은 아래 3가지 케이스에 따라 자동으로 다르게 생성됩니다.

| 케이스 | 조건 | 메시지 형태 |
|---|---|---|
| 월별 회비만 미납 | ~2026년 2월 미납분만 있는 경우 | `회비가 N원 미납되었습니다.` |
| 학기 회비만 미납 | 2026년 3월~ 미납분만 있는 경우 | `26-1학기 회비가 미납되었습니다.` |
| 혼재 | 월별 + 학기 미납이 모두 있는 경우 | `회비가 총 N원 미납되었습니다.`<br>항목별 내역 포함 |

> 트랙명 매칭은 영문자만 추출 후 소문자 변환하여 비교합니다. (예: `FrontEnd` = `frontend`)

---

## 프로젝트 구조

```
├── main.py                        # 장부 자동화 진입점
├── fill_ledger.py                 # 거래내역 → 관리 문서 자동 기입 진입점
├── fee_check.py                   # 회비 미납자 확인 진입점
├── ledger/
│   ├── membership_fee_parser.py   # 재학생 회비 관리 문서 → 장부 변환
│   └── hwp/
│       ├── hwp_generator.py       # HWP 증빙 서류 생성 (Windows, pyhwpx)
│       ├── hwp_generator_xml.py   # HWPX 증빙 서류 생성 (Linux/macOS, XML 직접 조작)
│       ├── image_downloader.py    # Google Drive 이미지 다운로드
│       └── image_packer.py        # 이미지 그리드 레이아웃 계산
├── ledger_filler/
│   └── filler.py                  # 거래내역 → 관리 문서 자동 기입 로직
├── fee_checker/
│   └── checker.py                 # 회비 미납자 확인 로직
└── templates/
    └── fee_notice.md              # 미납 알림 메시지 템플릿 (한글·엑셀 파일은 gitignore)
```
