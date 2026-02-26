#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
신한 거래내역 → 재학생 회비 관리 문서 자동 기입

자동 기입: 날짜, 이름(거래내역 내용), 입/출, 잔액, 소계/합계 수식
수동 기입: 내용(E열), 비고(G열)

사용법:
    python fill_ledger.py                        # 가장 최신 거래내역 파일 사용
    python fill_ledger.py 신한_거래내역/신한_거래내역_2602.xlsx
    python fill_ledger.py --force                # 이미 기입된 월도 덮어쓰기
"""

import io
import os
import re
import sys
import argparse
import tempfile
from collections import Counter
from copy import copy
from datetime import datetime
from urllib.parse import urlparse, parse_qs

import openpyxl
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter

try:
    from googleapiclient.errors import HttpError as _HttpError
except ImportError:
    _HttpError = OSError  # type: ignore[assignment,misc]

# ============================================================================
# Constants
# ============================================================================

# 관리 문서 컬럼 (1-based)
COL_MONTH = 3   # C: 월
COL_DATE = 4    # D: 날짜
COL_DESC = 5    # E: 내용 (수동 기입)
COL_NAME = 6    # F: 이름
COL_NOTE = 7    # G: 비고 (수동 기입)
COL_AMOUNT = 8  # H: 입/출
COL_BALANCE = 9 # I: 잔액

# 수식 작성에 사용할 열 문자 (COL_* 상수에서 파생)
_L_DESC    = get_column_letter(COL_DESC)    # E
_L_NOTE    = get_column_letter(COL_NOTE)    # G
_L_AMOUNT  = get_column_letter(COL_AMOUNT)  # H


# ============================================================================
# Google Drive integration
# ============================================================================

GOOGLE_TOKEN_FILE = '.google_token.json'
_GOOGLE_SCOPES = ['https://www.googleapis.com/auth/drive']


def _extract_sheet_id(url):
    """Google Sheets URL에서 스프레드시트 ID 추출."""
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not match:
        raise ValueError(f"Google Sheets URL에서 ID를 파싱할 수 없습니다: {url}")
    return match.group(1)


def _get_drive_service():
    """OAuth 인증을 통한 Drive 서비스 객체 반환. 토큰은 GOOGLE_TOKEN_FILE에 캐시."""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as err:
        raise ImportError(
            "Google Drive 연동에 필요한 패키지가 없습니다: pip install google-auth google-auth-oauthlib google-api-python-client"
        ) from err

    creds = None
    if os.path.exists(GOOGLE_TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, _GOOGLE_SCOPES)
        except (ValueError, OSError) as e:
            print(f"[WARNING] 토큰 파일이 손상되어 재인증합니다: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"[WARNING] 토큰 갱신 실패, 재인증합니다: {e}")
                creds = None
        if not creds or not creds.valid:
            secret_json = os.getenv('GOOGLE_OAUTH_CLIENT_JSON') or os.getenv('GOOGLE_SECRET_JSON')
            if not secret_json:
                raise ValueError("[ERROR] GOOGLE_OAUTH_CLIENT_JSON 환경변수가 설정되지 않았습니다.")
            flow = InstalledAppFlow.from_client_secrets_file(secret_json, _GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
            if not creds or not creds.valid:
                raise RuntimeError("OAuth 인증에 실패했습니다.")
        fd = os.open(GOOGLE_TOKEN_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, 'w') as f:
            f.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def download_sheet_as_xlsx(url):
    """
    Google Sheets를 임시 xlsx 파일로 내보내기.

    Returns:
        (sheet_id, tmp_path) — 호출자가 tmp_path를 사용 후 삭제 책임
    """
    from googleapiclient.http import MediaIoBaseDownload

    sheet_id = _extract_sheet_id(url)
    drive = _get_drive_service()

    request = drive.files().export_media(
        fileId=sheet_id,
        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    try:
        tmp.write(buf.getvalue())
        tmp.close()
    except Exception:
        tmp.close()
        os.unlink(tmp.name)
        raise

    return sheet_id, tmp.name


def _extract_drive_folder_id(url):
    """Google Drive 폴더 URL에서 폴더 ID 추출."""
    if '/folders/' in url:
        return url.split('/folders/')[1].split('/')[0].split('?')[0]
    return parse_qs(urlparse(url).query).get('id', [None])[0]


def _find_latest_transaction_in_folder(drive, folder_id):
    """폴더 내 신한_거래내역_YYMM.xlsx 파일 중 가장 최신 파일의 (file_id, name) 반환."""
    pattern = re.compile(r'신한_거래내역_(\d{4})\.xlsx$')
    files = []
    page_token = None
    while True:
        kwargs = dict(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
            fields='nextPageToken, files(id, name)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        if page_token:
            kwargs['pageToken'] = page_token
        result = drive.files().list(**kwargs).execute()
        for f in result.get('files', []):
            m = pattern.search(f['name'])
            if m:
                s = m.group(1)
                if 1 <= int(s[2:]) <= 12:
                    files.append((int(s), f['name'], f['id']))
        page_token = result.get('nextPageToken')
        if not page_token:
            break
    if not files:
        raise FileNotFoundError("폴더에서 신한_거래내역_YYMM.xlsx 파일을 찾을 수 없습니다.")
    files.sort(key=lambda x: x[0])
    _, name, file_id = files[-1]
    return file_id, name


def download_transaction_from_drive(url):
    """
    Google Drive 폴더에서 최신 거래내역 xlsx 파일 다운로드.

    폴더 내 신한_거래내역_YYMM.xlsx 중 가장 최신 파일을 다운로드.

    Returns:
        (original_filename, tmp_path) — 호출자가 tmp_path를 사용 후 삭제 책임
    """
    from googleapiclient.http import MediaIoBaseDownload

    folder_id = _extract_drive_folder_id(url)
    if not folder_id:
        raise ValueError(f"Google Drive 폴더 URL에서 ID를 파싱할 수 없습니다: {url}")

    drive = _get_drive_service()
    file_id, original_name = _find_latest_transaction_in_folder(drive, folder_id)

    request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    try:
        tmp.write(buf.getvalue())
        tmp.close()
    except Exception:
        tmp.close()
        os.unlink(tmp.name)
        raise

    return original_name, tmp.name


def upload_xlsx_to_sheet(sheet_id, local_path):
    """로컬 xlsx 파일을 Google Sheets 파일로 업로드하여 내용을 교체."""
    from googleapiclient.http import MediaFileUpload

    drive = _get_drive_service()
    media = MediaFileUpload(
        local_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True,
    )
    drive.files().update(
        fileId=sheet_id,
        body={'mimeType': 'application/vnd.google-apps.spreadsheet'},
        media_body=media,
        supportsAllDrives=True,
    ).execute()


# ============================================================================
# Parsing
# ============================================================================

def parse_transaction_file(filepath):
    """
    거래내역 엑셀 파일 파싱.

    Returns:
        list of (date_str, amount, name, balance)
        - amount: 입금이면 양수, 출금이면 음수
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    try:
        ws = wb.active
        transactions = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 6:
                continue
            no, date_str, deposit, withdrawal, name, balance = row[:6]
            if no is None:
                continue
            def _to_number(value):
                if isinstance(value, (int, float)):
                    return float(value)
                if isinstance(value, str):
                    raw = value.replace(',', '').strip()
                    try:
                        return float(raw) if raw else 0.0
                    except ValueError:
                        return 0.0
                return 0.0

            deposit = _to_number(deposit)
            withdrawal = _to_number(withdrawal)
            if deposit > 0:
                amount = deposit
            elif withdrawal > 0:
                amount = -withdrawal
            else:
                continue
            if isinstance(date_str, datetime):
                date_value = date_str.strftime('%Y.%m.%d')
            else:
                date_value = str(date_str).split()[0]
            if isinstance(balance, (int, float)):
                safe_balance = balance
            elif isinstance(balance, str):
                raw = balance.replace(',', '').strip()
                try:
                    safe_balance = float(raw) if raw else None
                except ValueError:
                    safe_balance = None
            else:
                safe_balance = None
            transactions.append((date_value, amount, str(name) if name else "", safe_balance))
    finally:
        wb.close()

    return transactions


def get_year_month_from_filename(filepath):
    """
    파일명에서 연도/월 추출.

    Examples:
        신한_거래내역_2601.xlsx → (2026, 1)
        신한_거래내역_2512.xlsx → (2025, 12)
    """
    basename = os.path.basename(filepath)
    match = re.search(r'_(\d{2})(\d{2})\.xlsx$', basename)
    if not match:
        raise ValueError(f"파일명에서 연도/월을 파싱할 수 없습니다: {basename}")
    yy, mm = int(match.group(1)), int(match.group(2))
    if not 1 <= mm <= 12:
        raise ValueError(f"파일명의 월 값이 유효하지 않습니다 (mm={mm}): {basename}")
    return 2000 + yy, mm


# ============================================================================
# Sheet manipulation
# ============================================================================

def find_month_section(ws, month):
    """
    월 헤더 행과 소계 행 번호 반환.

    Returns:
        (header_row, sogyeyu_row) or (None, None) if not found
    """
    month_label = f"{month}월"
    header_row = None
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row=row_idx, column=COL_MONTH).value == month_label:
            header_row = row_idx
            break
    if header_row is None:
        return None, None

    sogyeyu_row = None
    for row_idx in range(header_row + 1, ws.max_row + 1):
        if ws.cell(row=row_idx, column=COL_MONTH).value == '소계':
            sogyeyu_row = row_idx
            break

    return header_row, sogyeyu_row


def find_total_row(ws):
    """합계 행 번호 반환."""
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row=row_idx, column=COL_MONTH).value == '합계':
            return row_idx
    return None


def is_month_filled(ws, header_row):
    """월 데이터가 이미 기입되어 있는지 확인 (날짜 셀 기준)."""
    return ws.cell(row=header_row, column=COL_DATE).value is not None


def _unmerge_c_column(ws, header_row):
    """해당 월의 C열 병합 해제. header_row에서 시작하는 C열 병합만 제거."""
    to_remove = [
        str(m) for m in ws.merged_cells.ranges
        if m.min_col == COL_MONTH and m.max_col == COL_MONTH and m.min_row == header_row
    ]
    for range_str in to_remove:
        ws.unmerge_cells(range_str)


def _remerge_c_column(ws, header_row, data_end_row):
    """데이터 범위의 C열 재병합."""
    if data_end_row > header_row:
        ws.merge_cells(
            start_row=header_row, start_column=COL_MONTH,
            end_row=data_end_row, end_column=COL_MONTH,
        )


def _collect_and_unmerge_downstream_c(ws, from_row):
    """
    from_row 이후에 시작하는 C열 병합 범위를 수집하고 해제.

    openpyxl의 insert_rows/delete_rows가 병합 셀 범위를 자동으로
    갱신하지 않는 경우를 대비하여, 행 삽입/삭제 전에 호출한다.

    Returns:
        list of (min_row, max_row) tuples
    """
    affected = [
        (m.min_row, m.max_row)
        for m in list(ws.merged_cells.ranges)
        if m.min_col == COL_MONTH and m.max_col == COL_MONTH and m.min_row >= from_row
    ]
    for min_row, max_row in affected:
        ws.unmerge_cells(
            start_row=min_row, start_column=COL_MONTH,
            end_row=max_row, end_column=COL_MONTH,
        )
    return affected


def _remerge_with_offset(ws, ranges, delta):
    """저장된 C열 병합 범위를 delta만큼 행 번호를 이동하여 재병합."""
    for min_row, max_row in ranges:
        new_min = min_row + delta
        new_max = max_row + delta
        if new_max >= new_min:
            ws.merge_cells(
                start_row=new_min, start_column=COL_MONTH,
                end_row=new_max, end_column=COL_MONTH,
            )


def _capture_row_styles(ws, row_idx):
    """지정 행의 C~I열 셀 서식 캡처."""
    styles = {}
    for col in range(COL_MONTH, COL_BALANCE + 1):
        cell = ws.cell(row=row_idx, column=col)
        styles[col] = {
            'font': copy(cell.font),
            'fill': copy(cell.fill),
            'border': copy(cell.border),
            'alignment': copy(cell.alignment),
            'number_format': cell.number_format,
            'protection': copy(cell.protection),
        }
    return styles


def _apply_row_styles(ws, row_idx, styles):
    """캡처된 서식을 지정 행에 적용."""
    for col, style in styles.items():
        cell = ws.cell(row=row_idx, column=col)
        cell.font = copy(style['font'])
        cell.fill = copy(style['fill'])
        cell.border = copy(style['border'])
        cell.alignment = copy(style['alignment'])
        cell.number_format = style['number_format']
        cell.protection = copy(style['protection'])


def fill_month(ws, month, transactions, force=False, receipt_map=None):
    """
    해당 월의 거래내역을 관리 문서 시트에 기입.

    Returns:
        True  — 기입 완료
        None  — 이미 기입된 월을 건너뜀 (성공적 no-op)
        False — 오류
    """
    month_label = f"{month}월"
    header_row, sogyeyu_row = find_month_section(ws, month)

    if header_row is None or sogyeyu_row is None:
        print(f"[ERROR] {month_label} 섹션을 찾을 수 없습니다.")
        return False

    if is_month_filled(ws, header_row):
        if not force:
            print(f"[WARNING] {month_label} 데이터가 이미 존재합니다. 건너뜁니다. (덮어쓰려면 --force 사용)")
            return None
        print(f"[INFO] {month_label} 기존 데이터를 덮어씁니다.")

    tx_count = len(transactions)
    if tx_count == 0:
        print(f"[WARNING] {month_label} 거래 내역이 없습니다.")
        return True

    # 현재 플레이스홀더 행 수 (서식 캡처 전에 먼저 계산)
    placeholder_rows = sogyeyu_row - header_row  # 예: 2월: 53-48=5

    # 서식 캡처: 새 행에 적용할 "내부 행" 스타일
    # header_row는 top=medium(굵은 상단)이므로, 내부 행(header_row+1)에서 가져옴
    style_template_row = header_row + 1 if placeholder_rows > 1 else header_row
    row_style = _capture_row_styles(ws, style_template_row)

    # C열 병합 해제 (MergedCell 쓰기 오류 방지)
    _unmerge_c_column(ws, header_row)

    if tx_count > placeholder_rows:
        rows_to_insert = tx_count - placeholder_rows
        # 이후 월의 C열 병합 범위를 먼저 수집·해제 (insert_rows가 자동으로
        # 갱신하지 않는 경우 대비)
        downstream = _collect_and_unmerge_downstream_c(ws, sogyeyu_row)
        ws.insert_rows(sogyeyu_row, rows_to_insert)
        sogyeyu_row += rows_to_insert
        _remerge_with_offset(ws, downstream, rows_to_insert)
    elif tx_count < placeholder_rows:
        rows_to_delete = placeholder_rows - tx_count
        # 삭제될 행에 수동 기입 항목이 있으면 경고
        for r in range(header_row + tx_count, header_row + placeholder_rows):
            desc = ws.cell(row=r, column=COL_DESC).value
            note = ws.cell(row=r, column=COL_NOTE).value
            if desc or note:
                cols = []
                if desc:
                    cols.append("E열")
                if note:
                    cols.append("G열")
                print(f"[WARNING] 행 {r}의 수동 기입 항목({', '.join(cols)})이 삭제됩니다.")
        downstream = _collect_and_unmerge_downstream_c(ws, sogyeyu_row)
        ws.delete_rows(header_row + tx_count, rows_to_delete)
        sogyeyu_row -= rows_to_delete
        _remerge_with_offset(ws, downstream, -rows_to_delete)

    # 데이터 기입 (C열은 첫 행만 값, 나머지는 None → 이후 재병합)
    for i, (date_str, amount, name, balance) in enumerate(transactions):
        row_idx = header_row + i
        # 새로 삽입된 행(기존 플레이스홀더 범위 초과)에만 서식 복사
        if i >= placeholder_rows:
            _apply_row_styles(ws, row_idx, row_style)
        ws.cell(row=row_idx, column=COL_MONTH).value = month_label if i == 0 else None
        ws.cell(row=row_idx, column=COL_DATE).value = date_str
        ws.cell(row=row_idx, column=COL_NAME).value = name
        ws.cell(row=row_idx, column=COL_AMOUNT).value = amount
        ws.cell(row=row_idx, column=COL_BALANCE).value = balance

        # E열: 출금이고 영수증 매칭 결과가 있으면 하이퍼링크 기입
        if amount < 0 and receipt_map:
            key = (date_str, int(abs(amount)))
            if key in receipt_map:
                title, url = receipt_map[key]
                cell = ws.cell(row=row_idx, column=COL_DESC)
                cell.value = title
                cell.hyperlink = url
                # E6 하이퍼링크 셀 스타일과 동일하게 적용
                cell.font = Font(underline='single', color='0000FF')

    # 내부 행 상단 테두리 복원: header_row만 top=medium, 나머지는 top=None
    # (이전 실행에서 잘못 적용된 medium border가 남아 있는 경우도 교정)
    _no_top = Side(border_style=None)
    for i in range(1, tx_count):  # header_row 제외
        row_idx = header_row + i
        for col in range(COL_DATE, COL_BALANCE + 1):  # D~I (C는 병합으로 처리)
            cell = ws.cell(row=row_idx, column=col)
            b = cell.border
            cell.border = Border(
                top=_no_top,
                bottom=copy(b.bottom),
                left=copy(b.left),
                right=copy(b.right),
            )

    # C열 재병합
    data_end = sogyeyu_row - 1
    _remerge_c_column(ws, header_row, data_end)

    # 소계 수식 갱신
    ws.cell(row=sogyeyu_row, column=COL_MONTH).value = '소계'
    ws.cell(row=sogyeyu_row, column=COL_DATE).value = '입금'
    ws.cell(row=sogyeyu_row, column=COL_DESC).value = f'=SUMIF({_L_AMOUNT}{header_row}:{_L_AMOUNT}{data_end},">0")'
    ws.cell(row=sogyeyu_row, column=COL_NAME).value = '출금'
    ws.cell(row=sogyeyu_row, column=COL_NOTE).value = f'=SUMIF({_L_AMOUNT}{header_row}:{_L_AMOUNT}{data_end},"<0")*-1'
    ws.cell(row=sogyeyu_row, column=COL_AMOUNT).value = '합계'
    ws.cell(row=sogyeyu_row, column=COL_BALANCE).value = f'={_L_DESC}{sogyeyu_row}-{_L_NOTE}{sogyeyu_row}'

    print(f"[INFO] {month_label} 거래 {tx_count}건 기입 완료 (행 {header_row}~{data_end}, 소계 행 {sogyeyu_row})")
    return True


def update_total_formula(ws):
    """합계 행의 SUM 수식을 현재 소계 행 위치에 맞게 갱신."""
    total_row = find_total_row(ws)
    if total_row is None:
        return

    sogyeyu_rows = [
        row_idx
        for row_idx in range(2, total_row)
        if ws.cell(row=row_idx, column=COL_MONTH).value == '소계'
    ]
    if not sogyeyu_rows:
        return

    sum_e = ','.join(f'{_L_DESC}{r}' for r in sogyeyu_rows)
    sum_g = ','.join(f'{_L_NOTE}{r}' for r in sogyeyu_rows)

    ws.cell(row=total_row, column=COL_MONTH).value = '합계'
    ws.cell(row=total_row, column=COL_DATE).value = '입금'
    ws.cell(row=total_row, column=COL_DESC).value = f'=SUM({sum_e})'
    ws.cell(row=total_row, column=COL_NAME).value = '출금'
    ws.cell(row=total_row, column=COL_NOTE).value = f'=SUM({sum_g})'
    ws.cell(row=total_row, column=COL_AMOUNT).value = '합계'
    ws.cell(row=total_row, column=COL_BALANCE).value = f'={_L_DESC}{total_row}-{_L_NOTE}{total_row}'

    print(f"[INFO] 합계 행({total_row}) 수식 갱신 완료")


# ============================================================================
# Receipt matching (Google Drive)
# ============================================================================

def _find_drive_subfolder(drive, parent_id, name):
    """Drive 폴더 내 이름이 일치하는 하위 폴더 ID 반환. 없으면 None."""
    result = drive.files().list(
        q=(
            f"'{parent_id}' in parents"
            f" and name='{name}'"
            " and mimeType='application/vnd.google-apps.folder'"
            " and trashed=false"
        ),
        fields='files(id)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = result.get('files', [])
    return files[0]['id'] if files else None


def _list_receipt_candidates(drive, root_folder_id, date_str):
    """
    date_str(yyyy.mm.dd)로 시작하는 영수증 파일 목록 반환.

    폴더 구조를 두 가지 방식으로 탐색:
    1. root/{yyyy}/{mm} — 이름에 슬래시가 포함된 단일 폴더 (예: "2026/02")
    2. root/{yyyy}/{mm}/ — 중첩 폴더

    Returns list of {id, name, mimeType, webViewLink}.
    """
    parts = date_str.split('.')
    if len(parts) < 3:
        return []
    year, month = parts[0], parts[1]

    # 방식 1: "yyyy/mm" 이름의 단일 폴더
    month_id = _find_drive_subfolder(drive, root_folder_id, f"{year}/{month}")

    # 방식 2: 중첩 폴더 (year 폴더 → month 폴더)
    if not month_id:
        year_id = _find_drive_subfolder(drive, root_folder_id, year)
        if year_id:
            month_id = _find_drive_subfolder(drive, year_id, month)
            if not month_id:
                month_id = _find_drive_subfolder(drive, year_id, str(int(month)))

    if not month_id:
        return []

    # 월 폴더 내 전체 파일을 받아 클라이언트 사이드에서 날짜 접두사 필터
    # (Drive API의 name contains 는 점(.)을 구분자로 처리하여 오매칭 발생)
    all_files = []
    page_token = None
    while True:
        kwargs = dict(
            q=f"'{month_id}' in parents and trashed=false",
            fields='nextPageToken, files(id, name, mimeType, webViewLink)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        if page_token:
            kwargs['pageToken'] = page_token
        result = drive.files().list(**kwargs).execute()
        all_files.extend(result.get('files', []))
        page_token = result.get('nextPageToken')
        if not page_token:
            break

    return sorted(
        (f for f in all_files if f['name'].startswith(date_str)),
        key=lambda f: f['name'],
    )


def _export_drive_file_as_text(drive, file_id):
    """
    Google Drive 파일을 텍스트로 내보내기.

    text/plain → 실패 시 text/html(태그 제거) 순으로 시도.
    모두 실패하면 None 반환.
    """
    from googleapiclient.http import MediaIoBaseDownload

    for mime in ('text/plain', 'text/html'):
        buf = io.BytesIO()
        try:
            downloader = MediaIoBaseDownload(
                buf,
                drive.files().export_media(fileId=file_id, mimeType=mime),
            )
            done = False
            while not done:
                _, done = downloader.next_chunk()
        except Exception:
            continue

        text = buf.getvalue().decode('utf-8', errors='ignore')
        if mime == 'text/html':
            text = re.sub(r'<[^>]+>', '', text)
        return text

    return None


def _extract_amounts_from_drive_file(drive, file_id):
    """
    Google Drive 파일(Google Docs 등)을 텍스트로 내보내 포함된 정수 집합 반환.

    Returns set of int.
    """
    text = _export_drive_file_as_text(drive, file_id)
    if text is None:
        print(f"[WARNING] 파일 텍스트 내보내기 실패 (file_id={file_id}), 건너뜁니다.")
        return set()

    amounts = set()
    for m in re.finditer(r'([\d,]+)원', text):
        raw = m.group(1).replace(',', '')
        if raw:
            try:
                amounts.add(int(raw))
            except ValueError:
                pass
    return amounts


def _normalize_receipt_title(title):
    """영수증 파일 제목 정규화."""
    # 비기너 환급이 포함된 경우 → 비기너 환급
    if '비기너 환급' in title:
        return '비기너 환급'
    # "의 사본" 제거
    title = re.sub(r'의 사본\s*$', '', title).strip()
    # 이름 패턴 제거: 한글이름(트랙)님 또는 한글이름님
    title = re.sub(r'[가-힣]{2,5}(?:\s*\([A-Za-z]+\))?님\s*', '', title).strip()
    return title


def build_receipt_map(folder_url, transactions):
    """
    출금 거래 목록에 대해 영수증 파일 매칭 맵 생성.

    1차: 날짜(파일명 접두사 yyyy.mm.dd) 매칭
    2차: 파일 텍스트에 거래 금액(절댓값) 포함 여부 확인

    Returns:
        dict mapping (date_str, abs_amount_int) → (title, url)
        - title: 파일명에서 날짜 접두사와 확장자를 제거한 문자열
        - url:   Google Drive webViewLink
    """
    folder_id = _extract_drive_folder_id(folder_url)
    if not folder_id:
        raise ValueError(f"영수증 Drive 폴더 URL에서 ID를 파싱할 수 없습니다: {folder_url}")

    drive = _get_drive_service()
    receipt_map = {}

    # 동일 날짜·금액 출금이 2건 이상인 키는 어느 영수증인지 특정 불가 → 제외
    tx_counts = Counter(
        (date_str, int(abs(amount)))
        for date_str, amount, *_ in transactions
        if amount < 0
    )
    ambiguous = {key for key, cnt in tx_counts.items() if cnt > 1}

    withdrawal_dates = {date_str for date_str, *_ in tx_counts}

    for date_str in sorted(withdrawal_dates):
        candidates = _list_receipt_candidates(drive, folder_id, date_str)
        for f in candidates:
            amounts = _extract_amounts_from_drive_file(drive, f['id'])
            title = _normalize_receipt_title(f['name'][len(date_str):].strip())
            for amt in amounts:
                key = (date_str, amt)
                if key not in ambiguous and key not in receipt_map:
                    receipt_map[key] = (title, f['webViewLink'])
            # 이체 수수료 500원이 별도 기재된 경우: main + 500 키도 등록
            if 500 in amounts:
                for amt in amounts - {500}:
                    fee_key = (date_str, amt + 500)
                    if fee_key not in ambiguous and fee_key not in receipt_map:
                        receipt_map[fee_key] = (title, f['webViewLink'])

    return receipt_map


# ============================================================================
# Entry point
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="신한 거래내역 → 재학생 회비 관리 문서 자동 기입",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
자동 기입 항목: 날짜, 이름(거래내역 내용), 입/출, 잔액, 소계/합계 수식
수동 기입 항목: 내용(E열), 비고(G열)

예시:
  python fill_ledger.py
  python fill_ledger.py 신한_거래내역/신한_거래내역_2602.xlsx
  python fill_ledger.py --force
""",
    )
    parser.add_argument(
        "transaction_file",
        nargs="?",
        help="거래내역 파일 경로 (미지정 시 TRANSACTION_DRIVE_URL에서 다운로드)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="이미 데이터가 기입된 월도 강제로 덮어쓰기",
    )
    args = parser.parse_args()

    print("=" * 60)
    print("신한 거래내역 → 재학생 회비 관리 문서 자동 기입")
    print("=" * 60)

    tx_tmp_path = None
    tmp_path = None
    upload_ok = False
    preserve_tmp_path = False
    wb = None

    try:
        # 거래내역 파일 결정
        tx_original_name = None
        if args.transaction_file:
            tx_file = args.transaction_file
        else:
            tx_drive_url = os.getenv('TRANSACTION_DRIVE_URL')
            if not tx_drive_url:
                print("[ERROR] TRANSACTION_DRIVE_URL 환경변수가 설정되지 않았습니다.")
                sys.exit(1)
            print("\n[INFO] 거래내역 Drive에서 다운로드 중...")
            try:
                tx_original_name, tx_tmp_path = download_transaction_from_drive(tx_drive_url)
                tx_file = tx_tmp_path
                print(f"[INFO] 다운로드 완료: {tx_original_name}")
            except (FileNotFoundError, ValueError, OSError, _HttpError) as e:
                print(f"[ERROR] 거래내역 Drive 다운로드 실패: {e}")
                sys.exit(1)

        print(f"\n[INFO] 거래내역 파일: {tx_original_name or tx_file}")

        # 연도/월 파싱: Drive에서 받은 경우 원본 파일명 기준
        try:
            year, month = get_year_month_from_filename(tx_original_name or tx_file)
        except ValueError as e:
            print(f"[ERROR] {e}")
            sys.exit(1)
        print(f"[INFO] 대상: {year}년 {month}월")

        # 관리 문서 결정 (MANAGEMENT_SHEET_URL)
        management_sheet_url = os.getenv('MANAGEMENT_SHEET_URL')
        if not management_sheet_url:
            print("[ERROR] MANAGEMENT_SHEET_URL 환경변수가 설정되지 않았습니다.")
            sys.exit(1)

        print(f"[INFO] 관리 문서 (원격): {management_sheet_url}")
        try:
            sheet_id, tmp_path = download_sheet_as_xlsx(management_sheet_url)
            mgmt_file = tmp_path
            print(f"[INFO] 다운로드 완료 → 임시 파일: {tmp_path}")
        except (ValueError, OSError, _HttpError) as e:
            print(f"[ERROR] Google Sheets 다운로드 실패: {e}")
            sys.exit(1)

        # 거래내역 파싱
        try:
            transactions = parse_transaction_file(tx_file)
        except Exception as e:
            print(f"[ERROR] 거래내역 파일이 손상되었거나 읽을 수 없습니다: {e}")
            sys.exit(1)
        print(f"[INFO] 파싱된 거래 건수: {len(transactions)}건")

        # 관리 문서 열기
        try:
            wb = openpyxl.load_workbook(mgmt_file)
        except Exception as e:
            print(f"[ERROR] 관리 문서 파일이 손상되었거나 읽을 수 없습니다: {e}")
            sys.exit(1)
        sheet_name = f"{year}년"
        if sheet_name not in wb.sheetnames:
            print(f"[ERROR] 시트 '{sheet_name}'를 찾을 수 없습니다. (존재하는 시트: {wb.sheetnames})")
            sys.exit(1)
        ws = wb[sheet_name]

        # 영수증 매칭
        receipt_map = {}
        receipt_drive_url = os.getenv('RECEIPT_DIR')
        if receipt_drive_url:
            print("\n[INFO] 영수증 Drive 폴더에서 매칭 중...")
            try:
                receipt_map = build_receipt_map(receipt_drive_url, transactions)
                print(f"[INFO] 영수증 매칭 완료: {len(receipt_map)}건")
            except Exception as e:
                print(f"[WARNING] 영수증 매칭 실패 (건너뜀): {e}")

        # 데이터 기입
        print()
        success = fill_month(ws, month, transactions, force=args.force, receipt_map=receipt_map)
        if success is False:
            sys.exit(1)
        if success is None:
            sys.exit(0)

        # 합계 수식 갱신
        update_total_formula(ws)

        # 저장 후 업로드
        try:
            wb.save(tmp_path)
        except OSError as e:
            print(f"[ERROR] 임시 파일 저장 실패: {e}")
            sys.exit(1)

        print()
        print("=" * 60)
        try:
            print("[INFO] Google Sheets로 업로드 중...")
            upload_xlsx_to_sheet(sheet_id, tmp_path)
            print(f"[INFO] 업로드 완료: {management_sheet_url}")
            upload_ok = True
        except (OSError, ValueError, _HttpError) as e:
            print(f"[ERROR] 업로드 실패: {e}")
            preserve_tmp_path = True

        if not upload_ok:
            sys.exit(1)

        print("[INFO] 아래 항목은 수동으로 기입해주세요:")
        if not receipt_drive_url:
            print("       - E열 (내용): 회비 / 서버비 / 회식비 등")
        else:
            print("       - E열 (내용): 영수증 미매칭 출금 항목")
        print("       - G열 (비고): 납부 월 등")
        print("=" * 60)

    finally:
        if wb is not None:
            wb.close()
        if tx_tmp_path and os.path.exists(tx_tmp_path):
            os.remove(tx_tmp_path)
        if tmp_path and os.path.exists(tmp_path):
            if preserve_tmp_path:
                print(f"[INFO] 로컬 임시 파일은 보존됩니다: {tmp_path}")
            else:
                os.remove(tmp_path)
