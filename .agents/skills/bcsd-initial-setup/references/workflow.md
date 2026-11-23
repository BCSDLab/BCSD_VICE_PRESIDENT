# 초기 설정 절차

명령은 프로젝트 루트에서 실행한다. 아래 예시는 실제 초기 설정을 요청받았을 때 사용하며, 문서 작성·검수만 요청받았다면 설치·인증·온라인 접근을 실행하지 않는다. 기존 설정과 자료는 보존한다.

## 1. 환경과 목적 확인

- OS·Python 버전, `.venv`의 사용 가능 여부, 요청 기능을 확인한다. 기존 가상환경이 정상이라면 새로 만들거나 지우지 않는다. Python 최소 버전은 README와 실제 소스·패키지 지원 범위를 함께 확인한다.
- 거래 기입, 미납 계산, 제출 장부 작성, PDF 확인증, 실제 렌더링, 안내문/DM 중 이번에 준비할 기능을 구분한다. 선택 기능의 자료가 없다고 다른 기능까지 준비 불가라고 하지 않는다.
- `.env`, OAuth 클라이언트 JSON, 토큰의 존재·형식·필요 키만 확인한다. 파일 내용, `printenv` 전체, 토큰, 계정별 연락처를 로그로 출력하지 않는다. 현재 설정은 `.env.example`과 비교하되 예시 값을 실제 자격정보로 사용하지 않는다.

## 2. 가상환경과 패키지

새 가상환경이 필요한 경우에만 OS에 맞는 명령을 고른다.

```sh
# Linux/macOS: 기존 .venv가 없을 때만
python3 -m venv .venv
.venv/bin/python -m pip --version
```

```powershell
# Windows PowerShell: 기존 .venv가 없을 때만
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip --version
```

프로젝트 가상환경의 Python으로 아래 코드를 실행하면 `requirements.txt`를 단일 의존성 원본으로 사용한다. 현재 저장소는 Windows 전용 `pyhwpx`에 환경 마커가 없으므로 Linux/macOS에서는 이 항목만 임시 설치 목록에서 제외한다. 원본 파일을 고치거나 임의의 별도 패키지 목록을 유지하지 않는다. 다른 패키지가 실패하면 원인을 확인하며 임의로 더 제외하지 않는다.

```python
import re
import subprocess
import sys
import tempfile
from pathlib import Path

source = Path('requirements.txt').read_text(encoding='utf-8')
lines = source.splitlines()
if sys.platform != 'win32':
    lines = [line for line in lines
             if not re.match(r'^pyhwpx(?:\s|[<>=!~;\[]|$)', line.strip(), re.I)]
with tempfile.TemporaryDirectory(prefix='bcsd-setup-deps-') as directory:
    requirements = Path(directory) / 'requirements.txt'
    requirements.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', str(requirements)],
                   check=True)
```

- Windows의 HWP 자동화는 `pyhwpx`뿐 아니라 실제 한컴 한글 설치와 실행 가능한 환경도 필요하다. 설치 여부만으로 한글 자동화가 성공한다고 말하지 않는다.
- Linux/macOS의 증빙 생성은 `.hwpx` XML 경로다. `.hwp` 바이너리 템플릿만 있으면 사용할 수 없다.
- PDF 확인증 처리에는 `pdftotext`, `pdftoppm`이 필요하다. OS와 이미 있는 패키지 관리자를 확인한 뒤 poppler 설치 방법을 고른다. `pdfinfo`도 페이지 수 확인에 유용하다. 프로젝트 명령이 아닌 시스템 패키지 설치에는 사용자의 권한·환경 제약을 따른다.
- 실제 출력 검수에는 한글 또는 LibreOffice와 HWPX를 읽을 수 있는 변환 환경이 필요할 수 있다. LibreOffice 설치만으로 HWPX 가져오기가 보장되지 않으며 H2Orestart 등 현재 지원 경로를 확인한다. 렌더러는 실제 출력 검수용이지 단순 Sheets 조회의 필수 조건은 아니다.

## 3. .env와 입력 자료

새 `.env`가 필요한 경우에만 예시에서 생성한다. 기존 파일은 열어 덮어쓰지 않는다. 다음 코드는 새 파일을 독점 생성하므로 기존 파일이 있으면 실패한다. 파일 권한은 POSIX에서 소유자 읽기/쓰기만 허용하고, Windows에서는 적절한 사용자 ACL도 확인한다.

```python
import os
from pathlib import Path

example = Path('.env.example').read_bytes()
fd = os.open('.env', os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
with os.fdopen(fd, 'wb') as target:
    target.write(example)
```

| 준비할 기능 | 필요한 설정·자료 |
|---|---|
| Google 공통 인증 | `GOOGLE_OAUTH_CLIENT_JSON`: 실제 데스크톱 OAuth 클라이언트 JSON 파일 경로 |
| 거래 관리 문서 | `MANAGEMENT_SHEET_URL`: 실제 Sheets URL, 해당 `YYYY년` 시트 |
| 거래 파일 수집 | `TRANSACTION_DRIVE_URL`: 거래 원본 Drive 폴더, 요청 연월 파일 |
| 영수증 매칭 | `RECEIPT_DIR`: 영수증 폴더. 선택 기능이며 실제 폴더 구조 확인 |
| 납부 조회·미납 계산 | `FEE_SHEET_URL`: 별도의 납부 Sheets, 해당 `YYYY` 시트 |
| Docs API quota project | `GOOGLE_DOCS_QUOTA_PROJECT`: 필요한 경우에만 로그인 계정이 사용할 수 있는 실제 프로젝트 |
| 제출 문서 | `templates/ledger_format.xlsx`, Linux/macOS의 `templates/evid_format.hwpx`, Windows의 적절한 증빙 템플릿 |
| 안내문·DM 준비 | `templates/fee_notice.md`, 실제 발신 정보. DM에는 `SLACK_BOT_TOKEN`, `SLACK_SENDER_ID`, 발신자 이름·연락처, `member.csv`, `track.csv`도 필요 |

- 템플릿·명부·`.env`·키·토큰은 `.gitignore` 대상일 수 있어 새 체크아웃에 없을 수 있다. 정당하게 제공된 원본을 찾고, 필요한 입력만 요청한다. 가짜 영수증·회원·토큰·새 양식을 만들어 통과시키지 않는다.
- CSV 매핑은 `member.csv`의 `name,slack_id,track_id,is_deleted`와 `track.csv`의 `id,name,is_deleted`를 현재 코드와 대조한다. DB나 SSH 터널은 설정하지 않는다. 명부가 없더라도 DM을 준비하지 않는 작업은 진행할 수 있다.
- 상대 인증 경로와 `.google_token.json`은 실행 작업 폴더의 영향을 받는다. 최초 인증은 프로젝트 루트에서 수행한다. 장부 생성의 격리 작업 폴더는 장부 작성 스킬의 절차로 구성한다.
- `.env.example`의 프로젝트 ID, 문서 URL, Slack 토큰·ID, 연락처는 예시다. 실제 리소스와 값으로 확인하지 않은 상태에서 준비 완료라고 하지 않는다.
- Git 무시 규칙만으로 이미 추적된 비밀파일이 안전해지는 것은 아니다. 비밀 파일이 추적된 흔적을 발견하면 노출 값을 출력하거나 그대로 커밋하지 말고 위험과 별도 정리 필요성을 알린다. 키/토큰 삭제·폐기·재발급을 임의로 실행하지 않는다.

## 4. OAuth와 API 접근

1. 사용자가 접근 권한을 가진 Google Cloud 프로젝트에서 Drive API·Sheets API·Docs API가 필요한 기능에 맞게 활성화되어 있는지 확인한다. 데스크톱 OAuth 클라이언트와 동의 화면, 필요한 테스트 사용자/조직 정책을 확인한다. 프로젝트·클라이언트가 없으면 정상적인 사용자 제공/생성 절차가 필요하며 서비스 계정이나 임의 클라이언트로 우회하지 않는다.
2. 현재 `ledger_filler/filler.py::_get_credentials()`는 `https://www.googleapis.com/auth/drive` 전체 범위를 요청한다. 권한 범위를 인증 전에 설명하고, 계정 소유자가 실제 로그인 계정과 동의를 확인하게 한다. 초기 설정이 원격 장부 쓰기 동의는 아니다.
3. 프로젝트 루트에서 가상환경 Python으로 기존 인증 흐름을 호출한다. 기존 유효 토큰을 재사용하며 필요한 경우 갱신/로컬 서버 OAuth를 수행한다. 새 캐시 기록은 설정 작업에 필요한 로컬 변경이다.

```python
from pathlib import Path
from dotenv import load_dotenv
from ledger_filler.filler import _get_credentials

load_dotenv(Path.cwd() / '.env')
credentials = _get_credentials()
print('OAuth 인증 완료' if credentials.valid else 'OAuth 인증 미완료')
```

4. `.google_token.json`은 출력·커밋하지 않는다. POSIX 권한이 과도하면 소유자 권한으로 보호하는 조정을 검토한다. 기존 캐시를 무조건 삭제해 재인증하지 말고 갱신 실패·스코프 변경·권한 오류의 원인을 구분한다. 불가피한 재인증은 사용자에게 설명하고 기존 파일 교체 범위와 동의를 확인한다.
5. Cloud 프로젝트 권한과 대상 Docs/Drive/Sheets 파일 권한은 별개다. 로그인 계정이 실제 대상 리소스를 읽을 수 있어야 한다. quota project를 사용하면 Docs API 활성화와 해당 프로젝트의 사용 권한도 확인한다. 필요 권한이 없으면 최소 권한의 정상적인 요청 절차를 안내하고 임의 IAM 확대는 하지 않는다.

### 읽기 전용 온라인 점검

현재 코드의 `_get_drive_service()`, `_get_sheets_service()`와 URL 파서를 재사용하고 **요청된 리소스의 메타데이터만** 읽는다.

- Sheets: 두 URL을 각각 파싱하여 `spreadsheets().get(..., fields='sheets(properties(title))')`로 대상 연도 시트 존재를 확인한다. 문서 읽기 성공을 쓰기 성공으로 주장하지 않는다.
- Drive: 지정한 폴더에 `files().get(..., fields='mimeType,capabilities(canEdit)')` 등으로 폴더 유형과 필요한 접근을 확인한다. 거래 파일이 필요한 경우 지정 폴더의 파일 메타데이터를 확인한다. 사용자 Drive 전체를 뒤지지 않는다.
- Docs: 증빙 업무에서 실제로 사용할 문서가 있으면 동일 OAuth 자격과 선택적 quota project로 `documents().get(..., fields='title')` 등 필요한 메타데이터만 읽는다. 임의의 개인 문서를 골라 점검하지 않는다. 문서가 없으면 Docs 실제 접근은 미검증으로 남긴다.
- 단순 API 클라이언트 생성은 API 호출 성공이 아니다. `image_downloader`는 import 때 OAuth/API 클라이언트를 초기화하므로 오프라인 의존성 점검을 위해 불필요하게 import하지 않는다.
- 401/인증 갱신 오류, 403/API 비활성·quota project·파일 권한 오류, 404/리소스 ID·접근 문제를 구분한다. 오류를 숨기거나 같은 명령을 반복해 성공으로 취급하지 않는다. 권한 확대·토큰 폐기 대신 부족한 권한/입력을 구체적으로 보고한다.

## 5. 안전한 로컬 확인

다음 `--help` 경로는 현재 코드에서 실제 문서 작업 전에 argparse로 종료한다. 실행 시에도 현재 진입점에서 이 순서를 먼저 확인한다.

```sh
.venv/bin/python main.py --help
.venv/bin/python fill_ledger.py --help
.venv/bin/python fee_check.py --help
```

Windows에서는 실행 파일을 `.\.venv\Scripts\python.exe`로 바꾼다. `fee_check.py`를 옵션 없이 실행하면 안내문 생성·대화형 키워드 처리가 시작될 수 있으므로 초기 설정 smoke로 사용하지 않는다. `--send-dm`, `--force`, 실제 기간 인자 또는 실제 거래 파일 인자를 초기 확인에 붙이지 않는다.

선택 기능 명령의 존재만 확인하는 예시다. 가상환경 Python으로 실행한다.

```python
import shutil

for command in ('pdftotext', 'pdftoppm', 'pdfinfo', 'libreoffice'):
    print(command + ': ' + ('사용 가능' if shutil.which(command) else '미설치/현재 PATH에 없음'))
```

이는 명령의 존재 확인이지 실제 PDF 변환·한글 렌더링 성공 확인이 아니다. 초기 설정을 끝내려고 운영 거래나 가짜 영수증을 생성하지 않는다.

## 6. 완료 판정과 다음 단계

완료 보고는 실제 확인한 범위만 나눈다: Python/가상환경, Python 의존성, CLI 도움말, OAuth, Sheets·Drive·Docs 각각의 접근, 공식 템플릿, PDF 도구, 렌더링 환경, 선택적 안내/DM 입력. 필요한 항목이 미확인·미설치이면 어떤 기능이 준비됐고 무엇이 남았는지 명시한다. 설치나 인증을 실행하지 않은 문서 작성 작업에서는 실행했다고 하지 않는다.

- [장부 작성](../../bcsd-ledger-writing/SKILL.md): 공식 양식의 산출물 생성·확인증 추가.
- [관리 문서 업데이트](../../bcsd-fee-management/SKILL.md): 거래·납부 반영과 사유/미납 판단.
- [장부 검수](../../bcsd-ledger-audit/SKILL.md): 운영 원본을 바꾸지 않는 검증.

초기 설정은 이 작업들을 자동 실행하거나 커밋·푸시하는 단계가 아니다.
