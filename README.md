# BCSD 부회장 회계 자동화

BCSD 부회장의 반복적인 회계 업무를 자동화하는 Python 도구 모음입니다.

| 도구 | 기능 |
|---|---|
| `fill_ledger.py` | 신한 거래내역 → 회비 관리 문서 자동 기입 |
| `main.py` | 회비 관리 문서 → 기간별 장부 + 지출 증빙(HWP/HWPX) 생성 |
| `fee_check.py` | 미납자 집계 + 안내문 파일 생성 (선택: Slack DM 발송) |

---

## 요구사항

- Python 3.10+
- Google API 접근 권한 (Docs / Drive / Sheets)
- 증빙 서류 생성: Windows → `pyhwpx` 기반 `.hwp`/`.hwpx`, Linux/macOS → HWPX XML 직접 조작

## 설치

```bash
pip install -r requirements.txt
cp .env.example .env   # .env 구성 후 실행
```

---

## 환경 변수

### 공통

| 변수 | 설명 |
|---|---|
| `MANAGEMENT_SHEET_URL` | 재학생 회비 관리 문서 Google Sheets URL |
| `GOOGLE_OAUTH_CLIENT_JSON` | OAuth 클라이언트 JSON 경로 (`fill_ledger.py` 전용) |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | 서비스 계정 JSON 경로 (`main.py` 이미지 수집 전용) |
| `DEBUG` | Windows HWP 자동화 시 창 표시 여부 (`True` / `False`) |

### `fill_ledger.py` 추가

| 변수 | 설명 |
|---|---|
| `TRANSACTION_DRIVE_URL` | 거래내역 파일이 있는 Drive 폴더 URL |
| `RECEIPT_DIR` | 영수증 폴더 URL (선택) |

### `fee_check.py` 추가 (`--send-dm` 포함)

| 변수 | 설명 |
|---|---|
| `FEE_SHEET_URL` | 회비 납부 문서 URL |
| `SENDER_NAME` | 발신자 이름 |
| `SENDER_PHONE` | 발신자 연락처 |
| `SLACK_BOT_TOKEN` | Slack Bot Token |
| `SLACK_SENDER_ID` | 발신자 Slack user_id |
| `SSH_HOST`, `SSH_PORT`, `SSH_USER` | DB 터널용 SSH 정보 |
| `SSH_KEY_PATH` 또는 `SSH_PASSWORD` | SSH 인증 정보 |
| `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD` | MySQL 연결 정보 |
| `DB_TABLE`, `DB_TRACK_TABLE` | 회원 / 트랙 테이블명 |
| `DB_COL_*`, `DB_TRACK_COL_*` | 컬럼명 오버라이드 (선택) |

> Slack 봇 권한: `users:read`, `im:write`, `chat:write`

---

## 사용법

### 1. 거래내역 자동 기입 — `fill_ledger.py`

신한 거래내역 파일을 읽어 회비 관리 문서의 월 섹션을 자동 기입합니다.

```bash
python fill_ledger.py                                      # Drive에서 최신 파일 자동 선택
python fill_ledger.py 신한_거래내역/신한_거래내역_2602.xlsx  # 로컬 파일 직접 지정
python fill_ledger.py --force                              # 이미 기입된 월 덮어쓰기
```

최초 실행 시 Google OAuth 인증 창이 열리며, 토큰은 `.google_token.json`에 캐시됩니다.

**자동/수동 기입 범위**

| 열 | 항목 | 방식 |
|---|---|---|
| D | 날짜 | 자동 |
| E | 내용 | 수동 (영수증 매칭 시 일부 자동 링크) |
| F | 이름(거래내역 내용) | 자동 |
| G | 비고 | 수동 |
| H | 입/출 금액 | 자동 |
| I | 잔액 | 자동 |

소계/합계 수식은 행 수 변화에 맞춰 자동 갱신됩니다.

---

### 2. 장부 + 증빙 생성 — `main.py`

회비 관리 문서를 기간별로 파싱해 장부를 만들고, 지출 건에 대한 증빙 문서를 생성합니다.

```bash
python main.py <시작기간> <종료기간>
# 예시
python main.py 2025-11 2026-02
```

`output/` 디렉토리에 아래 파일이 생성됩니다.

```
output/
├── BCSD_202511_202602_장부.xlsx
└── BCSD_202511_202602_증빙자료.hwpx
```

---

### 3. 미납자 집계 + 안내문 생성 — `fee_check.py`

회비 납부 문서를 기준으로 미납자를 집계하고 개인별 안내문 파일을 생성합니다.

```bash
python fee_check.py                              # 전체 대상
python fee_check.py -e "Back-End,Android"        # 트랙 제외
python fee_check.py -p "홍길동_Android"           # 개인 제외
python fee_check.py --send-dm                    # Slack DM 발송 포함
```

**회비 규칙**

| 기간 | 부과 방식 | 금액 |
|---|---|---|
| ~2026-02 | 월별 | 10,000원/월 |
| 2026-03~ | 학기별 (1학기 3~8월 / 2학기 9~2월) | 60,000원/학기 |

생성된 안내문 파일은 `output/YYYY-MM/` 디렉토리에 저장됩니다.

---

## 프로젝트 구조

```
.
├── main.py                         # 장부 + 증빙 생성 진입점
├── fill_ledger.py                  # 거래내역 기입 진입점
├── fee_check.py                    # 미납 집계 진입점
├── common/
│   ├── google_drive.py             # Google Drive/Sheets URL 파싱 + 다운로드 공통 유틸
│   └── fee_notice.py               # 회비 안내문 템플릿 렌더링 공통 유틸
├── ledger_filler/
│   └── filler.py                   # 거래내역 → 회비 관리 문서 기입 로직
├── fee_checker/
│   └── checker.py                  # 미납 집계 + 안내문 생성 + Slack DM 발송 로직
├── ledger/
│   ├── membership_fee_parser.py    # 회비 관리 문서 파싱 → 장부 생성
│   └── hwp/
│       ├── hwp_generator.py        # 증빙 생성 (Windows, pyhwpx)
│       ├── hwp_generator_xml.py    # 증빙 생성 (Linux/macOS, XML)
│       ├── image_downloader.py     # 증빙 이미지 수집
│       └── image_packer.py         # 이미지 패킹
├── templates/
│   ├── ledger_format.xlsx          # 장부 출력 템플릿
│   ├── evid_format.hwpx            # 증빙 문서 템플릿
│   └── fee_notice.md               # 미납 안내문 템플릿
├── output/                         # 실행 결과물
└── receipt_images/                 # 증빙 이미지 캐시
```

---

## 주의사항

- Google API JSON 키 파일은 저장소에 커밋하지 마세요.
- 운영 전 테스트 문서에서 먼저 검증하는 것을 권장합니다.
- 서비스 계정이 접근해야 하는 Google Docs/Drive 파일은 서비스 계정 이메일에 공유되어 있어야 합니다.