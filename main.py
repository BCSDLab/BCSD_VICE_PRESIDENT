import os
import sys
import argparse
from dotenv import load_dotenv
from ledger import membership_fee_parser as mfp

load_dotenv()

OUTPUT_DIR = 'output'
IMAGE_DIR  = 'receipt_images'
def main():
    parser = argparse.ArgumentParser(description='BCSD 부회장 회계 자동화')
    parser.add_argument('start', help='시작 기간 (예: 2025-11)')
    parser.add_argument('end',   help='종료 기간 (예: 2026-02)')
    parser.add_argument(
        '--transfer-proof', metavar='PDF',
        help='이체 내역 입금확인증 PDF 경로 (지정 시 해당 증빙 이미지를 표에 추가)',
    )
    args = parser.parse_args()

    sheet_url = os.getenv('MANAGEMENT_SHEET_URL')
    if not sheet_url:
        print("[ERROR] MANAGEMENT_SHEET_URL 환경변수가 설정되지 않았습니다.", file=sys.stderr)
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    hwp_template = next(
        (f'templates/{f}' for f in os.listdir('templates') if f.startswith('evid_format')),
        None
    )
    if not hwp_template:
        raise FileNotFoundError("templates/evid_format.hwp(x) 파일을 찾을 수 없습니다.")

    period = f'{args.start.replace("-", "")}_{args.end.replace("-", "")}'
    ledger_output = os.path.join(OUTPUT_DIR, f'BCSD_{period}_장부.xlsx')
    hwp_ext       = os.path.splitext(hwp_template)[1]
    hwp_output    = os.path.join(OUTPUT_DIR, f'BCSD_{period}_증빙자료{hwp_ext}')

    from ledger_filler.filler import download_sheet_as_xlsx
    print(f"[INFO] 관리 문서 다운로드 중... ({sheet_url})")
    _, tmp_path = download_sheet_as_xlsx(sheet_url)
    print(f"[INFO] 다운로드 완료 → {tmp_path}")

    try:
        print(f"[1/3] 장부 파싱 중...")
        df = mfp.run(tmp_path, args.start, args.end, output_path=ledger_output)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    if df is None:
        return
    print(f"      → {len(df)}건 ({ledger_output})")

    expenses = df[df['입/출'] < 0].copy()
    expenses['종류'] = expenses['내용']
    print(f"      → 지출 {len(expenses)}건 증빙 서류 생성 대상")

    import ledger.hwp.image_downloader as imgd

    print("[2/3] 이미지 다운로드 중...")
    data_with_paths = imgd.run(expenses, IMAGE_DIR)

    if args.transfer_proof:
        from ledger.hwp import transfer_proof as tp
        print(f"      입금확인증 PDF 처리 중... ({args.transfer_proof})")
        proof_dir = os.path.join(IMAGE_DIR, 'transfer_proofs')
        pages = tp.pdf_to_pages(args.transfer_proof, proof_dir)
        print(f"      → {len(pages)}페이지 파싱 완료")
        matched = tp.match_transfer_proofs(pages, data_with_paths)
        print(f"      → {len(matched)}건 매칭")
        page_by_path = {p['image_path']: p for p in pages}
        for idx, proof_paths in matched.items():
            existing = data_with_paths.at[idx, 'img_paths']
            if not isinstance(existing, list):
                existing = list(existing) if existing else []
            data_with_paths.at[idx, 'img_paths'] = proof_paths + existing
            # 입금확인증 수수료 실제값으로 플래그 보정
            proof_page = page_by_path.get(proof_paths[0])
            if proof_page is not None:
                data_with_paths.at[idx, '이체수수료'] = proof_page['fee'] > 0

    print(f"[3/3] HWP 생성 중... ({hwp_output})")
    try:
        if sys.platform == 'win32':
            import ledger.hwp.hwp_generator as hwpg
        else:
            import ledger.hwp.hwp_generator_xml as hwpg
        hwpg.run(data_with_paths, hwp_template, hwp_output)
    except ValueError as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)
    print("완료")


if __name__ == '__main__':
    main()
