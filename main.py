import os
import argparse
from dotenv import load_dotenv
from ledger import membership_fee_parser as mfp
import hwp.image_downloader as imgd
import hwp.hwp_generator as hwpg

load_dotenv()

OUTPUT_DIR        = 'output'
IMAGE_DIR         = 'receipt_images'
HWP_TEMPLATE      = 'templates/evid_format.hwpx'
MEMBERSHIP_SOURCE = os.getenv('MEMBERSHIP_SOURCE')


def main():
    parser = argparse.ArgumentParser(description='BCSD 부회장 회계 자동화')
    parser.add_argument('start', help='시작 기간 (예: 2025-11)')
    parser.add_argument('end',   help='종료 기간 (예: 2026-02)')
    args = parser.parse_args()

    if not MEMBERSHIP_SOURCE:
        raise EnvironmentError("MEMBERSHIP_SOURCE가 .env에 설정되지 않았습니다.")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    period = f'{args.start.replace("-", "")}_{args.end.replace("-", "")}'
    ledger_output = os.path.join(OUTPUT_DIR, f'BCSD_{period}_장부.xlsx')
    hwp_output    = os.path.join(OUTPUT_DIR, f'BCSD_{period}_증빙자료.hwpx')

    df = mfp.run(MEMBERSHIP_SOURCE, args.start, args.end, output_path=ledger_output)
    if df is None:
        return

    expenses = df[df['입/출'] < 0].copy().reset_index(drop=True)
    expenses['종류'] = expenses['내용']
    expenses['링크'] = None

    data_with_paths = imgd.run(expenses, IMAGE_DIR)
    hwpg.run(data_with_paths, HWP_TEMPLATE, hwp_output)
    print('done')


if __name__ == '__main__':
    main()
