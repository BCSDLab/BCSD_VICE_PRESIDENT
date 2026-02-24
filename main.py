import os
import argparse
from dotenv import load_dotenv
from ledger import membership_fee_parser as mfp
import hwp.image_downloader as imgd
import hwp.hwp_generator as hwpg

load_dotenv()

OUTPUT_DIR        = 'output'
IMAGE_DIR         = 'receipt_images'
HWP_TEMPLATE      = next(
    (f'templates/{f}' for f in os.listdir('templates') if f.startswith('evid_format')),
    'templates/evid_format.hwpx'
)
def main():
    parser = argparse.ArgumentParser(description='BCSD 부회장 회계 자동화')
    parser.add_argument('source', help='재학생 회비 관리 문서 파일 경로')
    parser.add_argument('start',  help='시작 기간 (예: 2025-11)')
    parser.add_argument('end',    help='종료 기간 (예: 2026-02)')
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    period = f'{args.start.replace("-", "")}_{args.end.replace("-", "")}'
    ledger_output = os.path.join(OUTPUT_DIR, f'BCSD_{period}_장부.xlsx')
    hwp_ext       = os.path.splitext(HWP_TEMPLATE)[1]
    hwp_output    = os.path.join(OUTPUT_DIR, f'BCSD_{period}_증빙자료{hwp_ext}')

    print(f"[1/3] 장부 파싱 중... ({args.source})")
    df = mfp.run(args.source, args.start, args.end, output_path=ledger_output)
    if df is None:
        return
    print(f"      → {len(df)}건 ({ledger_output})")

    expenses = df[df['입/출'] < 0].copy().reset_index(drop=True)
    expenses['종류'] = expenses['내용']
    expenses['링크'] = None
    print(f"      → 지출 {len(expenses)}건 증빙 서류 생성 대상")

    print(f"[2/3] 이미지 다운로드 중...")
    data_with_paths = imgd.run(expenses, IMAGE_DIR)

    print(f"[3/3] HWP 생성 중... ({hwp_output})")
    hwpg.run(data_with_paths, HWP_TEMPLATE, hwp_output)
    print("완료")


if __name__ == '__main__':
    main()
