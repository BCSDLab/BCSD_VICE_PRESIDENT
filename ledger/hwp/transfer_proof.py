"""입금확인증 PDF → 이미지 변환 및 장부 레코드 매칭 모듈

pdftotext / pdftoppm (poppler) 시스템 도구를 사용한다.
"""

import os
import re
import subprocess


def _parse_page_text(text: str) -> dict | None:
    """입금확인증 페이지 텍스트에서 거래일시·입금금액·수수료 추출.

    반환: {date: 'YYYYMMDD', amount: int, fee: int, total: int}
    또는 파싱 실패 시 None.
    """
    date_m  = re.search(r'거래일시\s+(\d{4}\.\d{2}\.\d{2})', text)
    amt_m   = re.search(r'입금금액\s+([\d,]+)원', text)
    fee_m   = re.search(r'수수료\s+([\d,]+)원', text)

    if not date_m or not amt_m:
        return None

    date_str = date_m.group(1).replace('.', '')          # '20251104'
    amount   = int(amt_m.group(1).replace(',', ''))
    fee      = int(fee_m.group(1).replace(',', '')) if fee_m else 0
    return {'date': date_str, 'amount': amount, 'fee': fee, 'total': amount + fee}


def pdf_to_pages(pdf_path: str, out_dir: str, prefix: str = 'transfer') -> list[dict]:
    """PDF 각 페이지를 PNG 이미지로 변환하고 파싱 결과와 함께 반환.

    pdftoppm / pdftotext (poppler) 필요.

    반환: [{date, amount, fee, total, image_path}, ...]
    """
    os.makedirs(out_dir, exist_ok=True)

    # 페이지별 텍스트 추출 (Form Feed 0x0c 가 페이지 구분자)
    result = subprocess.run(
        ['pdftotext', '-layout', pdf_path, '-'],
        capture_output=True, text=True, check=True,
    )
    page_texts = result.stdout.split('\x0c')

    # pdftoppm으로 페이지별 PNG 생성 (150 DPI)
    img_prefix = os.path.join(out_dir, prefix)
    subprocess.run(
        ['pdftoppm', '-r', '150', '-png', pdf_path, img_prefix],
        check=True, capture_output=True,
    )

    # 생성된 이미지 파일 정렬 (pdftoppm: prefix-1.png, prefix-2.png, ...)
    base = os.path.basename(prefix)
    def _page_num(p: str) -> int:
        m = re.search(r'-(\d+)\.png$', p)
        return int(m.group(1)) if m else 0

    img_files = sorted(
        [os.path.join(out_dir, f) for f in os.listdir(out_dir)
         if f.startswith(base + '-') and f.endswith('.png')],
        key=_page_num,
    )

    pages = []
    for i, text in enumerate(page_texts):
        if i >= len(img_files):
            break
        info = _parse_page_text(text)
        if info is None:
            print(f"  [입금확인증] 페이지 {i + 1} 파싱 실패 — 건너뜀")
            continue
        info['image_path'] = img_files[i]
        pages.append(info)

    return pages


def match_transfer_proofs(pages: list[dict], df) -> dict:
    """입금확인증 페이지 목록을 DataFrame 레코드에 매칭.

    매칭 기준 (우선순위 순):
      1. date(YYYYMMDD) + total(입금금액 + 수수료) == abs(입/출)
      2. total만 일치 (날짜 정보가 없을 경우 폴백)

    한 페이지는 한 레코드에만 매칭된다 (먼저 매칭된 레코드 우선).

    반환: {df_index: [image_path, ...]}
    """
    matched: dict[int, list[str]] = {}
    used: set[int] = set()

    def _tx_date(row) -> str | None:
        m = re.match(r'(\d{4})\.(\d{2})\.(\d{2})', str(row['날짜']))
        return m.group(1) + m.group(2) + m.group(3) if m else None

    # 1차: date + total 매칭
    for idx, row in df.iterrows():
        tx_amount = abs(int(row['입/출']))
        tx_date   = _tx_date(row)
        for pi, page in enumerate(pages):
            if pi in used:
                continue
            if page['total'] != tx_amount:
                continue
            if tx_date and page['date'] != tx_date:
                continue
            matched.setdefault(idx, []).append(page['image_path'])
            used.add(pi)
            break

    # 2차: 날짜 미매칭 레코드에 대해 total만으로 재시도
    for idx, row in df.iterrows():
        if idx in matched:
            continue
        tx_amount = abs(int(row['입/출']))
        for pi, page in enumerate(pages):
            if pi in used:
                continue
            if page['total'] != tx_amount:
                continue
            matched.setdefault(idx, []).append(page['image_path'])
            used.add(pi)
            break

    unmatched_pages = [p for i, p in enumerate(pages) if i not in used]
    if unmatched_pages:
        print(f"  [입금확인증] 매칭되지 않은 페이지 {len(unmatched_pages)}개")

    return matched