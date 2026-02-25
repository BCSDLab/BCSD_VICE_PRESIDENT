"""
HWPX 증빙 서류 생성기 — XML 직접 조작 방식 (pyhwpx / Windows 불필요)

HWPX는 ZIP + XML 구조이므로, 템플릿 파일을 언패킹하여
section0.xml을 파싱·수정한 뒤 다시 패킹한다.
"""

import os
import random
import zipfile
from lxml import etree
from PIL import Image

from ledger.hwp.image_packer import Img, _layout

# ── HWPX XML 네임스페이스 ──────────────────────────────────────────────
HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
HC = 'http://www.hancom.co.kr/hwpml/2011/core'

# ── 단위 변환 ──────────────────────────────────────────────────────────
# 1 inch = 7200 HWP unit = 25.4 mm  →  1 mm ≈ 283.465 HWP unit
HWP_PER_MM = 7200 / 25.4

# ── 템플릿에서 읽은 셀 크기 (HWP unit) ───────────────────────────────
CELL_W      = 51877   # 표 전체 너비
TITLE_ROW_H = 4117    # 제목 행 높이
IMG_CELL_H  = 25332   # 이미지 행 높이
MARGIN_LR   = 510     # 좌우 셀 여백
MARGIN_TB   = 141     # 상하 셀 여백

# 이미지가 들어갈 실제 내부 너비·높이
IMG_W_HWP = CELL_W - 2 * MARGIN_LR    # 50857
IMG_H_HWP = IMG_CELL_H - 2 * MARGIN_TB  # 25050
IMG_W_MM  = IMG_W_HWP / HWP_PER_MM    # ≈ 179.4 mm
IMG_H_MM  = IMG_H_HWP / HWP_PER_MM    # ≈ 88.4 mm


# ── 헬퍼 함수 ─────────────────────────────────────────────────────────

def _rand_id() -> int:
    return random.randint(100_000_000, 2_000_000_000)


def _max_binary_idx(zip_files: dict) -> int:
    """BinData/imageN.* 중 최대 N 반환."""
    idx = 0
    for name in zip_files:
        if name.startswith('BinData/image'):
            try:
                n = int(name.split('/image')[1].split('.')[0])
                idx = max(idx, n)
            except ValueError:
                pass
    return idx


def _max_z_order(root: etree._Element) -> int:
    return max(
        (int(e.get('zOrder', 0)) for e in root.iter() if e.get('zOrder') is not None),
        default=0,
    )


def _find_expense_ps(root: etree._Element) -> list:
    """제목+이미지 표(2행 1열)를 담은 직계 hp:p 목록 반환."""
    result = []
    for p in root:
        if p.tag != f'{{{HP}}}p':
            continue
        for run in p:
            if run.tag != f'{{{HP}}}run':
                continue
            for tbl in run:
                if (tbl.tag == f'{{{HP}}}tbl'
                        and tbl.get('rowCnt') == '2'
                        and tbl.get('colCnt') == '1'):
                    result.append(p)
                    break
    return result


def _hp(parent: etree._Element, tag: str, attribs: dict | None = None) -> etree._Element:
    return etree.SubElement(parent, f'{{{HP}}}{tag}', attribs or {})


# ── 핵심 XML 빌더 ──────────────────────────────────────────────────────

def _build_pic(binary_id: str, img: Img, disp_w: int, disp_h: int, z: int) -> etree._Element:
    """인라인(treatAsChar=1) 이미지 요소를 만든다."""
    with Image.open(img.path) as im:
        raw_dpi = im.info.get('dpi', (96, 96))
    dpi_x = float(raw_dpi[0]) or 96
    dpi_y = float(raw_dpi[1]) or 96

    org_w = round(img.w / dpi_x * 7200)
    org_h = round(img.h / dpi_y * 7200)
    sx = disp_w / org_w if org_w else 1.0
    sy = disp_h / org_h if org_h else 1.0

    pic = etree.Element(f'{{{HP}}}pic', {
        'id':            str(_rand_id()),
        'zOrder':        str(z),
        'numberingType': 'PICTURE',
        'textWrap':      'TOP_AND_BOTTOM',
        'textFlow':      'BOTH_SIDES',
        'lock':          '0',
        'dropcapstyle':  'None',
        'href':          '',
        'groupLevel':    '0',
        'instid':        str(_rand_id()),
        'reverse':       '0',
    })

    _hp(pic, 'offset',  {'x': '0', 'y': '0'})
    _hp(pic, 'orgSz',   {'width': str(org_w), 'height': str(org_h)})
    _hp(pic, 'curSz',   {'width': str(disp_w), 'height': str(disp_h)})
    _hp(pic, 'flip',    {'horizontal': '0', 'vertical': '0'})
    _hp(pic, 'rotationInfo', {
        'angle': '0', 'centerX': str(disp_w // 2), 'centerY': str(disp_h // 2),
        'rotateimage': '1',
    })

    ri = _hp(pic, 'renderingInfo')
    etree.SubElement(ri, f'{{{HC}}}transMatrix',
                     {'e1': '1', 'e2': '0', 'e3': '0', 'e4': '0', 'e5': '1', 'e6': '0'})
    etree.SubElement(ri, f'{{{HC}}}scaMatrix',
                     {'e1': f'{sx:.6f}', 'e2': '0', 'e3': '0',
                      'e4': '0', 'e5': f'{sy:.6f}', 'e6': '0'})
    etree.SubElement(ri, f'{{{HC}}}rotMatrix',
                     {'e1': '1', 'e2': '0', 'e3': '0', 'e4': '0', 'e5': '1', 'e6': '0'})

    img_rect = _hp(pic, 'imgRect')
    for tag, x, y in [('pt0', 0, 0), ('pt1', org_w, 0), ('pt2', org_w, org_h), ('pt3', 0, org_h)]:
        etree.SubElement(img_rect, f'{{{HC}}}{tag}', {'x': str(x), 'y': str(y)})

    _hp(pic, 'imgClip',  {'left': '0', 'right': '0', 'top': '0', 'bottom': '0'})
    _hp(pic, 'inMargin', {'left': '0', 'right': '0', 'top': '0', 'bottom': '0'})
    etree.SubElement(pic, f'{{{HC}}}img', {
        'binaryItemIDRef': binary_id,
        'bright': '0', 'contrast': '0', 'effect': 'REAL_PIC', 'alpha': '0',
    })
    _hp(pic, 'effects')
    _hp(pic, 'sz', {
        'width': str(disp_w), 'widthRelTo': 'ABSOLUTE',
        'height': str(disp_h), 'heightRelTo': 'ABSOLUTE', 'protect': '0',
    })
    _hp(pic, 'pos', {
        'treatAsChar': '1', 'affectLSpacing': '0', 'flowWithText': '1',
        'allowOverlap': '0', 'holdAnchorAndSO': '0',
        'vertRelTo': 'PARA', 'horzRelTo': 'COLUMN',
        'vertAlign': 'TOP', 'horzAlign': 'LEFT',
        'vertOffset': '0', 'horzOffset': '0',
    })
    _hp(pic, 'outMargin', {'left': '0', 'right': '0', 'top': '0', 'bottom': '0'})
    _hp(pic, 'shapeComment')
    return pic


def _build_table(title: str, img_rows: list, z: int) -> tuple:
    """
    2행 1열 증빙 표 생성.
    img_rows: 행 리스트, 각 행 = [(binary_id, Img, disp_w_hwp, disp_h_hwp), ...]
    반환: (hp:tbl 요소, 다음 z 값)
    """
    tbl = etree.Element(f'{{{HP}}}tbl', {
        'id':              str(_rand_id()),
        'zOrder':          str(z),
        'numberingType':   'TABLE',
        'textWrap':        'TOP_AND_BOTTOM',
        'textFlow':        'BOTH_SIDES',
        'lock':            '0',
        'dropcapstyle':    'None',
        'pageBreak':       'CELL',
        'repeatHeader':    '1',
        'rowCnt':          '2',
        'colCnt':          '1',
        'cellSpacing':     '0',
        'borderFillIDRef': '3',
        'noAdjust':        '0',
    })
    z += 1

    total_h = TITLE_ROW_H + IMG_CELL_H
    _hp(tbl, 'sz',  {'width': str(CELL_W), 'widthRelTo': 'ABSOLUTE',
                     'height': str(total_h), 'heightRelTo': 'ABSOLUTE', 'protect': '0'})
    _hp(tbl, 'pos', {
        'treatAsChar': '0', 'affectLSpacing': '0', 'flowWithText': '1',
        'allowOverlap': '0', 'holdAnchorAndSO': '0',
        'vertRelTo': 'PARA', 'horzRelTo': 'COLUMN',
        'vertAlign': 'TOP', 'horzAlign': 'LEFT', 'vertOffset': '0', 'horzOffset': '0',
    })
    _hp(tbl, 'outMargin', {'left': '283', 'right': '283', 'top': '283', 'bottom': '283'})
    _hp(tbl, 'inMargin',  {'left': '510', 'right': '510', 'top': '141', 'bottom': '141'})

    # ── 행 1: 제목 ────────────────────────────────────────────────────
    tr1 = _hp(tbl, 'tr')
    tc1 = _hp(tr1, 'tc', {'name': '', 'header': '0', 'hasMargin': '0',
                           'protect': '0', 'editable': '0', 'dirty': '0',
                           'borderFillIDRef': '3'})
    sl1 = _hp(tc1, 'subList', {
        'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
        'vertAlign': 'CENTER', 'linkListIDRef': '0', 'linkListNextIDRef': '0',
        'textWidth': '0', 'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
    })
    p1  = _hp(sl1, 'p', {'id': '2147483648', 'paraPrIDRef': '22', 'styleIDRef': '22',
                          'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
    r1  = _hp(p1, 'run', {'charPrIDRef': '13'})
    _hp(r1, 't').text = title
    lsa1 = _hp(p1, 'linesegarray')
    _hp(lsa1, 'lineseg', {
        'textpos': '0', 'vertpos': '0', 'vertsize': '1100', 'textheight': '1100',
        'baseline': '550', 'spacing': '0', 'horzpos': '0', 'horzsize': '50856', 'flags': '393216',
    })
    _hp(tc1, 'cellAddr',   {'colAddr': '0', 'rowAddr': '0'})
    _hp(tc1, 'cellSpan',   {'colSpan': '1', 'rowSpan': '1'})
    _hp(tc1, 'cellSz',     {'width': str(CELL_W), 'height': str(TITLE_ROW_H)})
    _hp(tc1, 'cellMargin', {'left': '510', 'right': '510', 'top': '141', 'bottom': '141'})

    # ── 행 2: 이미지 ──────────────────────────────────────────────────
    tr2 = _hp(tbl, 'tr')
    tc2 = _hp(tr2, 'tc', {'name': '', 'header': '0', 'hasMargin': '0',
                           'protect': '0', 'editable': '0', 'dirty': '0',
                           'borderFillIDRef': '3'})
    sl2 = _hp(tc2, 'subList', {
        'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
        'vertAlign': 'CENTER', 'linkListIDRef': '0', 'linkListNextIDRef': '0',
        'textWidth': '0', 'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
    })

    if img_rows:
        for row in img_rows:
            p = _hp(sl2, 'p', {'id': '0', 'paraPrIDRef': '20', 'styleIDRef': '0',
                                'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
            run = _hp(p, 'run', {'charPrIDRef': '14'})
            max_h = 0
            for binary_id, img, disp_w, disp_h in row:
                run.append(_build_pic(binary_id, img, disp_w, disp_h, z))
                z += 1
                max_h = max(max_h, disp_h)
            lsa = _hp(p, 'linesegarray')
            _hp(lsa, 'lineseg', {
                'textpos': '0', 'vertpos': '0',
                'vertsize': str(max_h), 'textheight': str(max_h),
                'baseline': str(round(max_h * 0.85)), 'spacing': '600',
                'horzpos': '0', 'horzsize': '50856', 'flags': '393216',
            })
    else:
        p = _hp(sl2, 'p', {'id': '0', 'paraPrIDRef': '20', 'styleIDRef': '0',
                            'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
        _hp(p, 'run', {'charPrIDRef': '14'})
        lsa = _hp(p, 'linesegarray')
        _hp(lsa, 'lineseg', {'textpos': '0', 'vertpos': '0', 'vertsize': '1000',
                              'textheight': '1000', 'baseline': '850', 'spacing': '600',
                              'horzpos': '0', 'horzsize': '50856', 'flags': '393216'})

    _hp(tc2, 'cellAddr',   {'colAddr': '0', 'rowAddr': '1'})
    _hp(tc2, 'cellSpan',   {'colSpan': '1', 'rowSpan': '1'})
    _hp(tc2, 'cellSz',     {'width': str(CELL_W), 'height': str(IMG_CELL_H)})
    _hp(tc2, 'cellMargin', {'left': '510', 'right': '510', 'top': '141', 'bottom': '141'})

    return tbl, z


# ── 공개 진입점 ────────────────────────────────────────────────────────

def run(data, t_path: str, o_path: str):
    """
    HWPX 증빙 서류 생성 (XML 직접 조작, pyhwpx 불필요).

    data      : image_downloader.run() 반환 DataFrame (img_paths 컬럼 포함)
    t_path    : 템플릿 .hwpx 파일 경로
    o_path    : 출력 .hwpx 파일 경로
    """
    # 1. 템플릿 ZIP 읽기 (.hwp 바이너리는 지원 불가)
    if not zipfile.is_zipfile(t_path):
        ext = os.path.splitext(t_path)[1]
        raise ValueError(
            f"템플릿 파일이 HWPX 형식이 아닙니다 ({ext}).\n"
            "HWP 바이너리 포맷은 지원하지 않습니다. "
            "한글에서 '다른 이름으로 저장 → .hwpx'로 변환 후 다시 시도하세요."
        )

    with zipfile.ZipFile(t_path, 'r') as zin:
        zip_files = {name: zin.read(name) for name in zin.namelist()}

    # 2. section0.xml 파싱
    root = etree.fromstring(zip_files['Contents/section0.xml'])

    # 3. 기존 증빙 표 단락 제거 (삽입 위치 기억)
    expense_ps = _find_expense_ps(root)
    insert_idx = list(root).index(expense_ps[0]) if expense_ps else len(list(root))
    for p in expense_ps:
        root.remove(p)

    # 4. 카운터 초기화
    bin_counter = _max_binary_idx(zip_files) + 1
    z           = _max_z_order(root) + 1
    new_binaries: dict[str, bytes] = {}

    # 5. 지출 행마다 표 생성
    for data_idx, row in data.iterrows():
        title     = f'{data_idx + 1}. {row["종류"]}'
        img_paths = row.get('img_paths', []) or []

        img_rows: list[list] = []
        if img_paths:
            imgs   = [Img(p) for p in img_paths]
            layout = _layout(imgs, IMG_W_MM, IMG_H_MM)
            if layout and layout.items:
                rows_n, cols_n = layout.grid
                for r_idx in range(rows_n):
                    row_items = []
                    for c_idx in range(cols_n):
                        i = r_idx * cols_n + c_idx
                        if i >= len(layout.items):
                            break
                        item = layout.items[i]
                        img  = imgs[item.idx]

                        ext = os.path.splitext(img.path)[1].lstrip('.').lower() or 'png'
                        bid = f'image{bin_counter}'
                        bin_counter += 1

                        with open(img.path, 'rb') as f:
                            new_binaries[f'BinData/{bid}.{ext}'] = f.read()

                        disp_w = round(item.size[0] * HWP_PER_MM)
                        disp_h = round(item.size[1] * HWP_PER_MM)
                        row_items.append((bid, img, disp_w, disp_h))
                    if row_items:
                        img_rows.append(row_items)
            else:
                print(f"  [{data_idx + 1}] 레이아웃 계산 실패 — 이미지 셀 비워둠")
        else:
            print(f"  [{data_idx + 1}] 증빙 자료 누락 — 확인 필요")

        tbl_elem, z = _build_table(title, img_rows, z)

        # hp:p > hp:run > hp:tbl 구조로 래핑 (템플릿 패턴과 동일)
        p_wrap = etree.Element(f'{{{HP}}}p', {
            'id': '0', 'paraPrIDRef': '20', 'styleIDRef': '0',
            'pageBreak': '0', 'columnBreak': '0', 'merged': '0',
        })
        run_wrap = etree.SubElement(p_wrap, f'{{{HP}}}run', {'charPrIDRef': '7'})
        run_wrap.append(tbl_elem)
        etree.SubElement(run_wrap, f'{{{HP}}}t')
        lsa = etree.SubElement(p_wrap, f'{{{HP}}}linesegarray')
        etree.SubElement(lsa, f'{{{HP}}}lineseg', {
            'textpos': '0', 'vertpos': '0', 'vertsize': '1700', 'textheight': '1700',
            'baseline': '1445', 'spacing': '1020', 'horzpos': '0', 'horzsize': '0',
            'flags': '393216',
        })

        root.insert(insert_idx, p_wrap)
        insert_idx += 1

    # 6. XML 직렬화
    zip_files['Contents/section0.xml'] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True,
    )
    for path, content in new_binaries.items():
        zip_files[path] = content

    # 7. HWPX 패킹
    os.makedirs(os.path.dirname(o_path) or '.', exist_ok=True)
    with zipfile.ZipFile(o_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, content in zip_files.items():
            zout.writestr(name, content)

    print(f"HWPX 생성 완료: {o_path}")