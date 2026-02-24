import os
import argparse
from dotenv import load_dotenv
from ledger.excel_parser import parse
from ledger import membership_fee_parser as mfp
import hwp.image_downloader as imgd
import hwp.hwp_generator as hwpg

load_dotenv()

OUTPUT_DIR        = os.getenv('OUTPUT_DIR',        'output')
LEDGER_PATH       = os.getenv('LEDGER_PATH',       'ledger.xlsx')
HWP_TEMPLATE      = os.getenv('HWP_TEMPLATE_PATH', 'evid_format.hwpx')
HWP_OUTPUT        = os.getenv('HWP_OUTPUT_PATH',   'output.hwpx')
IMAGE_DIR         = os.getenv('IMAGE_DIR',          'receipt_images')
MEMBERSHIP_SOURCE = os.getenv('MEMBERSHIP_SOURCE')

HEADER_ROW     = 0
DATA_START_ROW = 1


def main():
    parser = argparse.ArgumentParser(description='BCSD 부회장 회계 자동화')
    parser.add_argument('start', help='시작 기간 (예: 2025-11)')
    parser.add_argument('end',   help='종료 기간 (예: 2026-02)')
    args = parser.parse_args()

    if not MEMBERSHIP_SOURCE:
        raise EnvironmentError("MEMBERSHIP_SOURCE가 .env에 설정되지 않았습니다.")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    ledger_output = os.path.join(
        OUTPUT_DIR,
        f'ledger_{args.start.replace("-", "")}_{args.end.replace("-", "")}.xlsx'
    )
    hwp_output = os.path.join(OUTPUT_DIR, os.path.basename(HWP_OUTPUT))

    mfp.run(MEMBERSHIP_SOURCE, args.start, args.end,
            output_path=ledger_output, ledger_path=LEDGER_PATH)

    data = parse(LEDGER_PATH, HEADER_ROW, DATA_START_ROW)
    if data.empty:
        print("empty data")
        return

    data_with_paths = imgd.run(data, IMAGE_DIR)
    hwpg.run(data_with_paths, HWP_TEMPLATE, hwp_output)
    print('done')


if __name__ == '__main__':
    main()
