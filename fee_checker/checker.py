#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BCSD 회비 납부 검증 및 미납 메시지 생성 프로그램
"""

import os
import sys
import re
import glob
import argparse
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl


# ============================================================================
# Constants
# ============================================================================

EXCEL_FILE_PATTERN = "재학생 회비 납부 문서_*.xlsx"
EXCEL_FILE_PREFIX = "재학생 회비 납부 문서_"
TEMPLATE_FILE = "templates/fee_notice.md"
OUTPUT_BASE_DIR = "output"

# 제외 키워드 (비고 열에서 검사)
EXCLUDE_KEYWORDS = ["졸업", "활동 중지", "휴학", "군 휴학", "트랙장", "교육장"]

# 엑셀 시트 설정
SHEETS_TO_PROCESS = ["2025", "2026"]
DATA_START_ROW = 5  # Row 5부터 데이터 시작

# 컬럼 매핑 (0-indexed)
COL_TRACK = 2  # C열: 트랙
COL_NAME = 3  # D열: 이름
COL_NOTES = 4  # E열: 비고
COL_MONTHS_START = 5  # F열: 1월 (F=5, G=6, ..., Q=16)

# 월별 컬럼 (F=1월, G=2월, ..., Q=12월)
MONTH_COLUMNS = {
    1: 5,  # F
    2: 6,  # G
    3: 7,  # H
    4: 8,  # I
    5: 9,  # J
    6: 10,  # K
    7: 11,  # L
    8: 12,  # M
    9: 13,  # N
    10: 14,  # O
    11: 15,  # P
    12: 16,  # Q
}


# ============================================================================
# Helper Functions
# ============================================================================


def find_latest_excel_file():
    """
    최신 날짜 접미사를 가진 엑셀 파일 찾기

    패턴: "재학생 회비 납부 문서_YYYYMMDD.xlsx"

    Returns:
        str: 최신 파일명, 또는 None (파일 없음)
    """
    matching_files = glob.glob(EXCEL_FILE_PATTERN)

    if not matching_files:
        return None

    date_pattern = re.compile(r"_(\d{8})\.xlsx$")
    files_with_dates = []

    for filepath in matching_files:
        filename = os.path.basename(filepath)
        match = date_pattern.search(filename)
        if match:
            date_str = match.group(1)
            try:
                date_obj = datetime.strptime(date_str, "%Y%m%d")
                files_with_dates.append((date_obj, filename))
            except ValueError:
                continue

    if not files_with_dates:
        return None

    files_with_dates.sort(key=lambda x: x[0], reverse=True)
    return files_with_dates[0][1]


def parse_sheet(ws, sheet_name):
    """
    시트에서 데이터 행 파싱

    Returns:
        list of dict: 각 행의 데이터
            {
                'track': str,
                'name': str,
                'notes': str,
                'months': {1: 'O'/'-'/None, 2: ..., 12: ...}
            }
    """
    rows = []
    row_idx = DATA_START_ROW

    while True:
        name_cell = ws.cell(row=row_idx, column=COL_NAME + 1)
        name = name_cell.value

        if name is None or str(name).strip() == "":
            break

        track_cell = ws.cell(row=row_idx, column=COL_TRACK + 1)
        notes_cell = ws.cell(row=row_idx, column=COL_NOTES + 1)

        track = track_cell.value if track_cell.value else ""
        notes = notes_cell.value if notes_cell.value else ""

        months = {}
        for month_num, col_idx in MONTH_COLUMNS.items():
            cell = ws.cell(row=row_idx, column=col_idx + 1)
            val = cell.value

            if val == "O":
                months[month_num] = "O"
            elif val == "-" or val == "−":
                months[month_num] = "-"
            else:
                months[month_num] = None

        rows.append(
            {
                "track": str(track).strip(),
                "name": str(name).strip(),
                "notes": str(notes).strip(),
                "months": months,
            }
        )

        row_idx += 1

    return rows


def should_exclude_row(row_data):
    """
    행 제외 여부 판단

    제외 조건:
    1. 비고에 제외 키워드 포함
    2. 모든 체크 대상 월이 "-" (면제)
    """
    notes = row_data["notes"]
    for keyword in EXCLUDE_KEYWORDS:
        if keyword in notes:
            return True

    months = row_data["months"]
    all_exempt = all(val == "-" for val in months.values())
    if all_exempt:
        return True

    return False


def calculate_unpaid_months(row_data, sheet_name, current_month):
    """
    미납 월 수 계산

    Args:
        sheet_name: "2025" 또는 "2026"
        current_month: 현재 월 (1~12)

    Returns:
        int: 미납 월 수
    """
    months = row_data["months"]

    if sheet_name == "2025":
        check_until = 12
    elif sheet_name == "2026":
        check_until = max(0, current_month - 1)
    else:
        check_until = 12

    unpaid_count = 0
    for month_num in range(1, check_until + 1):
        val = months.get(month_num)
        if val is None:
            unpaid_count += 1

    return unpaid_count


def _normalize_track(track):
    """트랙명 정규화: 영문자만 추출 후 소문자 변환 (예: 'FrontEnd' → 'frontend')"""
    return re.sub(r'[^a-zA-Z]', '', track).lower()


@contextmanager
def _ssh_tunnel(ssh_host, ssh_port, ssh_user, remote_host, remote_port, ssh_key_path=None, ssh_password=None):
    """paramiko 기반 SSH 포트 포워딩 터널"""
    import select
    import threading
    import socketserver
    import paramiko

    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    connect_kwargs: dict = {"port": ssh_port, "username": ssh_user}
    if ssh_key_path:
        connect_kwargs["key_filename"] = os.path.expanduser(ssh_key_path)
    else:
        connect_kwargs["password"] = ssh_password

    client.connect(ssh_host, **connect_kwargs)
    transport = client.get_transport()
    assert transport is not None

    class _ForwardHandler(socketserver.BaseRequestHandler):
        def handle(self):
            chan = transport.open_channel(
                "direct-tcpip",
                (remote_host, remote_port),
                self.request.getpeername(),
            )
            if chan is None:
                return
            while True:
                readable = select.select([self.request, chan], [], [], 5)[0]
                if self.request in readable:
                    data = self.request.recv(1024)
                    if not data:
                        break
                    chan.send(data)
                if chan in readable:
                    data = chan.recv(1024)
                    if not data:
                        break
                    self.request.send(data)
            chan.close()

    server = socketserver.ThreadingTCPServer(("127.0.0.1", 0), _ForwardHandler)
    local_port = server.server_address[1]
    threading.Thread(target=server.serve_forever, daemon=True).start()

    try:
        yield local_port
    finally:
        server.shutdown()
        client.close()


@contextmanager
def _db_connection():
    """SSH 터널 경유 MySQL 연결 컨텍스트 매니저 (readonly)"""
    try:
        import pymysql
    except ImportError:
        raise ImportError("필요한 패키지: pip install pymysql paramiko")

    ssh_host = os.getenv("SSH_HOST", "")
    ssh_port = int(os.getenv("SSH_PORT", "22"))
    ssh_user = os.getenv("SSH_USER", "")
    ssh_key_path = os.getenv("SSH_KEY_PATH")
    ssh_password = os.getenv("SSH_PASSWORD", "")

    db_host = os.getenv("DB_HOST", "127.0.0.1")
    db_port = int(os.getenv("DB_PORT", "3306"))
    db_name = os.getenv("DB_NAME", "")
    db_user = os.getenv("DB_USER", "")
    db_password = os.getenv("DB_PASSWORD", "")

    assert ssh_host, "SSH_HOST 환경 변수가 설정되지 않았습니다."
    assert ssh_user, "SSH_USER 환경 변수가 설정되지 않았습니다."

    with _ssh_tunnel(ssh_host, ssh_port, ssh_user, db_host, db_port, ssh_key_path, ssh_password) as local_port:
        conn = pymysql.connect(
            host="127.0.0.1",
            port=local_port,
            user=db_user,
            password=db_password,
            database=db_name,
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor,
        )
        try:
            yield conn
        finally:
            conn.close()


def fetch_slack_id_map():
    """DB에서 (이름, 트랙명) → Slack ID 매핑 조회 (readonly SELECT)"""
    table = os.getenv("DB_TABLE", "")
    col_name = os.getenv("DB_COL_NAME", "")
    col_slack_id = os.getenv("DB_COL_SLACK_ID", "")
    track_table = os.getenv("DB_TRACK_TABLE", "")
    track_col_id = os.getenv("DB_TRACK_COL_ID", "id")
    track_col_name = os.getenv("DB_TRACK_COL_NAME", "name")

    assert table, "DB_TABLE 환경 변수가 설정되지 않았습니다."
    assert col_name, "DB_COL_NAME 환경 변수가 설정되지 않았습니다."
    assert col_slack_id, "DB_COL_SLACK_ID 환경 변수가 설정되지 않았습니다."
    assert track_table, "DB_TRACK_TABLE 환경 변수가 설정되지 않았습니다."

    with _db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"SELECT m.`{col_name}`, t.`{track_col_name}` AS track_name, m.`{col_slack_id}`"
                f" FROM `{table}` m"
                f" JOIN `{track_table}` t ON m.`track_id` = t.`{track_col_id}`"
                f" WHERE m.`{col_slack_id}` IS NOT NULL AND m.`{col_slack_id}` != ''"
                f" AND m.`is_deleted` = 0 AND t.`is_deleted` = 0"
            )
            return {
                (row[col_name], _normalize_track(row["track_name"])): row[col_slack_id]
                for row in cur.fetchall()
            }


def aggregate_unpaid_fees(wb, current_month, excluded_tracks=None):
    """
    2025/2026 시트 데이터 통합 및 미납 금액 계산

    Returns:
        dict: {(name, track): {'name': str, 'track': str, 'unpaid_amount': int}}
    """
    if excluded_tracks is None:
        excluded_tracks = set()
    else:
        excluded_tracks = set(excluded_tracks)

    aggregated = {}
    excluded_keys = set()
    exempt_names_from_2025 = set()

    for sheet_name in SHEETS_TO_PROCESS:
        if sheet_name not in wb.sheetnames:
            print(f"[WARNING] 시트 '{sheet_name}' 없음, 건너뜀")
            continue

        ws = wb[sheet_name]
        rows = parse_sheet(ws, sheet_name)

        print(f"\n[INFO] 시트 '{sheet_name}': {len(rows)}개 행 파싱됨")

        included_count = 0
        excluded_count = 0

        if sheet_name == "2025":
            for row_data in rows:
                notes = row_data["notes"]
                for keyword in EXCLUDE_KEYWORDS:
                    if keyword in notes:
                        exempt_names_from_2025.add(row_data["name"])
                        break

        for row_data in rows:
            name = row_data["name"]
            track = row_data["track"]
            key = (name, track)

            if name in exempt_names_from_2025:
                excluded_keys.add(key)
                excluded_count += 1
                continue

            if track in excluded_tracks:
                excluded_keys.add(key)
                excluded_count += 1
                continue

            if should_exclude_row(row_data):
                excluded_keys.add(key)
                excluded_count += 1
                continue

            included_count += 1

            unpaid_months = calculate_unpaid_months(row_data, sheet_name, current_month)
            unpaid_amount = unpaid_months * 10000

            if key not in aggregated:
                aggregated[key] = {"name": name, "track": track, "unpaid_amount": 0}

            aggregated[key]["unpaid_amount"] += unpaid_amount

        print(f"[INFO] 시트 '{sheet_name}': 포함 {included_count}명, 제외 {excluded_count}명")

    for excluded_key in excluded_keys:
        aggregated.pop(excluded_key, None)

    filtered = {k: v for k, v in aggregated.items() if v["unpaid_amount"] > 0}
    return filtered


# ============================================================================
# Main Functions
# ============================================================================


def get_output_directory():
    """현재 연-월 기준으로 output 디렉토리 생성 및 반환 (예: output/2026-02/)"""
    now = datetime.now()
    output_dir = os.path.join(OUTPUT_BASE_DIR, f"{now.year:04d}-{now.month:02d}")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def get_previous_month_end_date():
    """전월 말일을 한국어 형식으로 반환 (예: "2026년 1월 31일")"""
    today = datetime.now()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

    year = last_day_of_previous_month.year
    month = last_day_of_previous_month.month
    day = last_day_of_previous_month.day

    return f"{year}년 {month}월 {day}일"


def format_amount(amount):
    """금액을 천 단위 콤마 형식으로 변환 (예: 120000 → "120,000")"""
    return f"{amount:,}"


def generate_unique_filename(name, track, used_filenames):
    """중복 이름 처리를 위한 고유 파일명 생성"""
    track_filename = f"{name}_{track}.txt"

    if track_filename not in used_filenames:
        used_filenames.add(track_filename)
        return track_filename

    counter = 1
    while True:
        numbered_filename = f"{name}_{track}_{counter}.txt"
        if numbered_filename not in used_filenames:
            used_filenames.add(numbered_filename)
            return numbered_filename
        counter += 1


def generate_message_files(unpaid_data, output_dir, template_path):
    """
    미납 회원별 메시지 파일 생성

    Returns:
        tuple: (생성된 파일 수, 총 미납 금액)
    """
    for existing_file in Path(output_dir).glob("*.txt"):
        existing_file.unlink()

    with open(template_path, "r", encoding="utf-8") as f:
        template_content = f.read()

    previous_month_date = get_previous_month_end_date()

    used_filenames = set()
    files_generated = 0
    total_unpaid_amount = 0

    for (name, track), data in unpaid_data.items():
        unpaid_amount = data["unpaid_amount"]
        formatted_amount = format_amount(unpaid_amount)

        message = template_content.replace("{이름}", name)
        message = message.replace("{금액}", formatted_amount)
        message = message.replace("{year}년 {month}월 {day}일", previous_month_date)

        filename = generate_unique_filename(name, track, used_filenames)
        filepath = os.path.join(output_dir, filename)

        with open(filepath, "w", encoding="utf-8") as f:
            f.write(message)

        files_generated += 1
        total_unpaid_amount += unpaid_amount

    return files_generated, total_unpaid_amount


def send_slack_dms(unpaid_data, template_path):
    """
    미납 회원에게 Slack DM 발송

    Returns:
        tuple: (발송 성공 수, 실패/미매칭 수)
    """
    try:
        from slack_sdk import WebClient
        from slack_sdk.errors import SlackApiError
    except ImportError:
        raise ImportError("slack_sdk 패키지가 필요합니다: pip install slack-sdk")

    token = os.getenv("SLACK_BOT_TOKEN")
    if not token:
        raise ValueError("SLACK_BOT_TOKEN 환경 변수가 설정되지 않았습니다.")

    sender_id = os.getenv("SLACK_SENDER_ID", "")
    if not sender_id:
        raise ValueError("SLACK_SENDER_ID 환경 변수가 설정되지 않았습니다.")

    client = WebClient(token=token)

    with open(template_path, "r", encoding="utf-8") as f:
        template_content = f.read()

    previous_month_date = get_previous_month_end_date()

    print("[INFO] DB에서 Slack ID 조회 중...")
    name_to_user_id = fetch_slack_id_map()
    print(f"[INFO] 조회된 멤버 수: {len(name_to_user_id)}명")

    sent = 0
    failed = 0

    for (name, track), data in unpaid_data.items():
        user_id = name_to_user_id.get((name, _normalize_track(track)))
        if not user_id:
            print(f"[WARNING] Slack 유저를 찾을 수 없음: {name} ({track})")
            failed += 1
            continue

        formatted_amount = format_amount(data["unpaid_amount"])
        message = template_content.replace("{이름}", name)
        message = message.replace("{멘션}", f"<@{sender_id}>")
        message = message.replace("{금액}", formatted_amount)
        message = message.replace("{year}년 {month}월 {day}일", previous_month_date)

        try:
            dm_resp = client.conversations_open(users=[user_id])
            channel_id = dm_resp["channel"]["id"]
            client.chat_postMessage(channel=channel_id, text=message)
            print(f"[INFO] DM 발송 완료: {name} ({track})")
            sent += 1
        except SlackApiError as e:
            print(f"[ERROR] DM 발송 실패: {name} ({track}) - {e.response['error']}")
            failed += 1

    return sent, failed


def parse_excluded_tracks(exclude_args):
    """CLI 인자에서 제외할 트랙 파싱"""
    excluded_tracks = set()
    if exclude_args:
        for arg in exclude_args:
            tracks = [t.strip() for t in arg.split(",")]
            excluded_tracks.update(tracks)
    return excluded_tracks


def main():
    parser = argparse.ArgumentParser(description="BCSD 회비 납부 검증 및 미납 메시지 생성 프로그램")
    parser.add_argument(
        "-e",
        "--exclude-track",
        action="append",
        dest="exclude_tracks",
        help="제외할 트랙명 (반복 사용 가능, 쉼표로 구분 가능)",
    )
    parser.add_argument(
        "--send-dm",
        action="store_true",
        help="미납 회원에게 Slack DM 발송 (SLACK_BOT_TOKEN 환경 변수 필요)",
    )
    args = parser.parse_args()

    excluded_tracks = parse_excluded_tracks(args.exclude_tracks)

    print("=" * 70)
    print("BCSD 회비 납부 검증 및 미납 메시지 생성 프로그램")
    print("=" * 70)

    output_dir = get_output_directory()
    print(f"\n[INFO] 출력 디렉토리: {output_dir}")

    if excluded_tracks:
        print(f"[INFO] 제외할 트랙: {', '.join(sorted(excluded_tracks))}")

    excel_file = find_latest_excel_file()
    if not excel_file:
        print(f"[ERROR] 엑셀 파일을 찾을 수 없습니다 (패턴: {EXCEL_FILE_PATTERN})")
        sys.exit(1)

    if not os.path.exists(TEMPLATE_FILE):
        print(f"[ERROR] 템플릿 파일을 찾을 수 없습니다 ({TEMPLATE_FILE})")
        sys.exit(1)

    print(f"[INFO] 선택된 엑셀 파일: {excel_file}")
    print(f"[INFO] 선택된 템플릿 파일: {TEMPLATE_FILE}")

    print("\n" + "=" * 70)
    print("엑셀 파싱 및 미납 계산")
    print("=" * 70)

    wb = openpyxl.load_workbook(excel_file, data_only=False)
    current_month = datetime.now().month

    print(f"[INFO] 현재 월: {current_month}")
    print(f"[INFO] 처리 대상 시트: {SHEETS_TO_PROCESS}")

    unpaid_data = aggregate_unpaid_fees(wb, current_month, excluded_tracks)

    print("\n" + "=" * 70)
    print("미납 데이터 요약")
    print("=" * 70)
    print(f"[INFO] 총 미납 대상자: {len(unpaid_data)}명")

    if unpaid_data:
        print("\n미납 데이터 샘플 (최대 5명):")
        for idx, ((name, track), data) in enumerate(unpaid_data.items()):
            if idx >= 5:
                break
            print(f"  - {name} ({track}): {data['unpaid_amount']:,}원")

    print("\n" + "=" * 70)
    print("메시지 생성 및 파일 출력")
    print("=" * 70)

    files_generated, total_unpaid_amount = generate_message_files(
        unpaid_data, output_dir, TEMPLATE_FILE
    )

    print(f"\n[INFO] 생성된 파일 수: {files_generated}개")
    print(f"[INFO] 총 미납 금액: {format_amount(total_unpaid_amount)}원")

    if args.send_dm and unpaid_data:
        print("\n" + "=" * 70)
        print("Slack DM 발송")
        print("=" * 70)
        dm_sent, dm_failed = send_slack_dms(unpaid_data, TEMPLATE_FILE)
        print(f"[INFO] DM 발송 완료: {dm_sent}명, 실패/미매칭: {dm_failed}명")

    print("\n" + "=" * 70)
    print("처리 완료 요약")
    print("=" * 70)
    print(f"[INFO] 처리된 시트: {', '.join(SHEETS_TO_PROCESS)}")
    print(f"[INFO] 미납 대상자: {len(unpaid_data)}명")
    print(f"[INFO] 생성된 메시지 파일: {files_generated}개")
    print(f"[INFO] 총 미납 금액: {format_amount(total_unpaid_amount)}원")
    print(f"[INFO] 출력 디렉토리: {output_dir}")
    print("=" * 70)
