from excel_parser import parse
import image_downloader as imgd
import hwp_generator as hwpg
import membership_fee_parser as mfp

EXCEL_FILE_PATH = 'ledger.xlsx'
HWP_FILE_PATH = 'evid_format.hwpx'
HWP_OUTPUT_PATH = 'output.hwpx'
IMAGE_DIR = 'receipt_images'
HEADER_ROW = 0
DATA_START_ROW = 1

MEMBERSHIP_SOURCE = '재학생 회비 관리 문서_20260225.xlsx'
MEMBERSHIP_START = '2025-11'
MEMBERSHIP_END = '2026-02'

def main():
    mfp.run(MEMBERSHIP_SOURCE, MEMBERSHIP_START, MEMBERSHIP_END)

    data = parse(EXCEL_FILE_PATH, HEADER_ROW, DATA_START_ROW)
    if data.empty:
        print("empty data")
        return

    data_with_paths = imgd.run(data, IMAGE_DIR)
    hwpg.run(data_with_paths, HWP_FILE_PATH, HWP_OUTPUT_PATH)
    print('done')

if __name__ == '__main__':
    main()
