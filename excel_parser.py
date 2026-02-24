import pandas as pd
import openpyxl
from utils import convert_to_doc_url

def _load(file, header):
    return pd.read_excel(file, header=header)

def _get_links(file, num_rows, start_row):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active
    links = []

    for i in range(start_row + 1, start_row + 1 + num_rows):
        cell = sheet[f'B{i}']
        links.append(convert_to_doc_url(cell.hyperlink.target) if cell.hyperlink else None)

    return links

def _filter(df):
    w_df = df[df['금액'] < 0].copy()
    w_df['금액'] = w_df['금액'].abs()
    return w_df

def parse(file, header, start_row):
    df = _load(file, header)
    df['링크'] = _get_links(file, df.shape[0], start_row)
    return _filter(df)
