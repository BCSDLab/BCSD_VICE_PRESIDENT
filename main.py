from excel_parser import parse
import image_downloader as imgd
import hwp_generator as hwpg

EXCEL_FILE_PATH = 'ledger.xlsx'
HWP_FILE_PATH = 'evid_format.hwpx'
HWP_OUTPUT_PATH = 'output.hwpx'
IMAGE_DIR = 'receipt_images'
HEADER_ROW = 0
DATA_START_ROW = 1

def main():
    data = parse(EXCEL_FILE_PATH, HEADER_ROW, DATA_START_ROW)
    if data.empty:
        print("empty data")
        return
    
    data_with_paths = imgd.run(data, IMAGE_DIR)
    hwpg.run(data_with_paths, HWP_FILE_PATH, HWP_OUTPUT_PATH)
    print('done')

if __name__ == '__main__':
    main()
