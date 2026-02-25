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

## 2. 회비 미납자 확인 및 Slack DM 발송 (`fee_check.py`)

재학생 회비 납부 문서에서 미납자를 집계하고, 개인별 알림 메시지 파일을 생성합니다.
`--send-dm` 옵션을 사용하면 미납자에게 Slack DM을 자동으로 발송합니다.

### 환경 설정

`.env`에 아래 항목을 추가합니다. (SSH 터널을 통해 DB에 **readonly**로 접근합니다.)

| 변수 | 설명 |
|---|---|
| `SLACK_BOT_TOKEN` | Slack Bot Token (`xoxb-...`) |
| `SLACK_SENDER_ID` | 발신자(본인) Slack user_id |
| `SENDER_NAME` | 발신자 이름 (메시지 서명에 사용) |
| `SENDER_PHONE` | 발신자 전화번호 (메시지 문의처에 사용) |
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
| `DB_TRACK_TABLE` | 트랙 테이블명 |
| `DB_TRACK_COL_ID` | 트랙 PK 컬럼명 (기본값: `id`) |
| `DB_TRACK_COL_NAME` | 트랙명 컬럼명 (기본값: `name`) |

> Slack Bot에는 `users:read`, `im:write`, `chat:write` 권한이 필요합니다.

### 사용법

```bash
python fee_check.py [-e <트랙명>] [--send-dm]
```

실행 시 `재학생 회비 납부 문서_*.xlsx` 파일 중 가장 최신 날짜 파일을 자동으로 선택합니다.

**예시 — 파일 생성만:**

```bash
python fee_check.py
```

**예시 — 특정 트랙 제외:**

```bash
python fee_check.py -e "Back-End" -e "Android"
python fee_check.py -e "Back-End,Android,Design"
```

**예시 — 파일 생성 + Slack DM 발송:**

```bash
python fee_check.py --send-dm
```

### 출력 결과

`output/YYYY-MM/` 디렉토리에 미납자별 `.txt` 파일이 생성됩니다.
`--send-dm` 사용 시 각 미납자에게 `templates/fee_notice.md` 내용으로 DM이 발송됩니다.

> 트랙명 매칭은 영문자만 추출 후 소문자 변환하여 비교합니다. (예: `FrontEnd` = `frontend`)

---

## 프로젝트 구조

```
├── main.py                        # 장부 자동화 진입점
├── fee_check.py                   # 회비 미납자 확인 진입점
├── ledger/
│   ├── membership_fee_parser.py   # 재학생 회비 관리 문서 → 장부 변환
│   └── hwp/
│       ├── hwp_generator.py       # HWP 증빙 서류 생성 (Windows, pyhwpx)
│       ├── hwp_generator_xml.py   # HWPX 증빙 서류 생성 (Linux/macOS, XML 직접 조작)
│       ├── image_downloader.py    # Google Drive 이미지 다운로드
│       └── image_packer.py        # 이미지 그리드 레이아웃 계산
├── fee_checker/
│   └── checker.py                 # 회비 미납자 확인 로직
└── templates/
    └── fee_notice.md              # 미납 알림 메시지 템플릿 (한글·엑셀 파일은 gitignore)
```
