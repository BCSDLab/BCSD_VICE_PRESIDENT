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
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl


# ============================================================================
# Constants
# ============================================================================

EXCEL_FILE_PATTERN = "재학생 회비 납부 문서_*.xlsx"
EXCEL_FILE_PREFIX = "재학생 회비 납부 문서_"
TEMPLATE_FILE = "fee_checker/fee_notice.md"
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

    print("\n" + "=" * 70)
    print("처리 완료 요약")
    print("=" * 70)
    print(f"[INFO] 처리된 시트: {', '.join(SHEETS_TO_PROCESS)}")
    print(f"[INFO] 미납 대상자: {len(unpaid_data)}명")
    print(f"[INFO] 생성된 메시지 파일: {files_generated}개")
    print(f"[INFO] 총 미납 금액: {format_amount(total_unpaid_amount)}원")
    print(f"[INFO] 출력 디렉토리: {output_dir}")
    print("=" * 70)
