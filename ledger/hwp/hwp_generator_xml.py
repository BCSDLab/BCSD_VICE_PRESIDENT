"""
HWPX 증빙 서류 생성기 — XML 직접 조작 방식 (pyhwpx / Windows 불필요)

HWPX는 ZIP + XML 구조이므로, 템플릿 파일을 언패킹하여
section0.xml을 파싱·수정한 뒤 다시 패킹한다.
"""

import itertools
import os
import zipfile
from dataclasses import dataclass, field
from lxml import etree
from PIL import Image, UnidentifiedImageError
from PIL.Image import DecompressionBombError

from ledger.hwp.image_packer import Img, _layout

# ── HWPX XML 네임스페이스 ──────────────────────────────────────────────
HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
HC = 'http://www.hancom.co.kr/hwpml/2011/core'

# ── 단위 변환 ──────────────────────────────────────────────────────────
# 1 inch = 7200 HWP unit = 25.4 mm  →  1 mm ≈ 283.465 HWP unit
HWP_PER_MM = 7200 / 25.4


@dataclass
class _TemplateParams:
    """템플릿 section0.xml에서 동적으로 추출한 레이아웃·스타일 파라미터.

    추출에 실패한 항목은 기존 템플릿을 분석해 측정한 기본값을 유지한다.
    """
    # 셀 크기 (HWP unit)
    cell_w:      int = 51877
    title_row_h: int = 4117
    img_cell_h:  int = 25332
    # 셀 여백
    margin_lr: int = 510
    margin_tb: int = 141
    # 표 outMargin (sz.height 보정용)
    out_margin: int = 283
    # 표 pos.vertOffset
    vert_offset: int = 434
    # 스타일 ID
    border_fill_id: str = '3'
    title_para_pr:  str = '22'
    title_style_id: str = '22'
    title_char_pr:  str = '13'
    img_para_pr:    str = '20'
    img_style_id:   str = '0'
    img_char_pr:    str = '14'
    wrap_para_pr:   str = '20'
    wrap_char_pr:   str = '7'
    # 제목 행 lineseg
    title_lineseg: dict = field(default_factory=lambda: {
        'textpos': '0', 'vertpos': '0', 'vertsize': '1100', 'textheight': '1100',
        'baseline': '550', 'spacing': '0', 'horzpos': '0', 'horzsize': '50856',
        'flags': '393216',
    })
    # 래퍼 단락 lineseg
    wrap_lineseg: dict = field(default_factory=lambda: {
        'textpos': '0', 'vertpos': '0', 'vertsize': '1700', 'textheight': '1700',
        'baseline': '1445', 'spacing': '1020', 'horzpos': '0', 'horzsize': '0',
        'flags': '393216',
    })

    @property
    def img_w_mm(self) -> float:
        """이미지 레이아웃 기준 너비 (mm) — 셀 전체 영역."""
        return self.cell_w / HWP_PER_MM

    @property
    def img_h_mm(self) -> float:
        """이미지 레이아웃 기준 높이 (mm) — 셀 전체 영역."""
        return self.img_cell_h / HWP_PER_MM


# ── 헬퍼 함수 ─────────────────────────────────────────────────────────

# run() 시작 시 리셋되는 순차 ID 카운터 (다중 호출 간 결정론적 ID 보장)
_id_counter = itertools.count(100_000_000)

def _next_id() -> int:
    return next(_id_counter)


def _reset_id_counter() -> None:
    global _id_counter
    _id_counter = itertools.count(100_000_000)


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


_MIME = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'bmp': 'image/bmp', 'gif': 'image/gif',  'webp': 'image/webp',
}

def _update_content_hpf(hpf_bytes: bytes, new_binaries: dict) -> bytes:
    """content.hpf의 opf:manifest에 신규 바이너리 항목을 추가한다."""
    root = etree.fromstring(hpf_bytes)
    # 네임스페이스를 루트 태그에서 동적으로 추출 (하드코딩 오류 방지)
    opf_ns = root.tag.split('}')[0].lstrip('{') if '}' in root.tag else ''
    manifest = root.find(f'{{{opf_ns}}}manifest') if opf_ns else root.find('manifest')
    if manifest is None:
        raise ValueError(
            "content.hpf에서 manifest를 찾지 못했습니다. "
            "템플릿의 OPF 구조/네임스페이스를 확인하세요."
        )

    for bin_path in new_binaries:               # 'BinData/image4.png'
        fname = bin_path.split('/')[-1]          # 'image4.png'
        bid   = fname.rsplit('.', 1)[0]          # 'image4'
        ext   = fname.rsplit('.', 1)[-1].lower() # 'png'
        etree.SubElement(manifest, f'{{{opf_ns}}}item', {
            'id':         bid,
            'href':       bin_path,
            'media-type': _MIME.get(ext, 'application/octet-stream'),
            'isEmbeded':  '1',
        })
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


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


def _read_template_params(expense_ps: list) -> _TemplateParams:
    """증빙 표 템플릿에서 레이아웃·스타일 파라미터를 동적으로 추출한다.

    추출에 실패한 항목은 _TemplateParams 기본값을 유지한다.
    """
    p = _TemplateParams()
    if not expense_ps:
        return p

    wrap_p = expense_ps[0]

    # ── 래퍼 단락 스타일 ──────────────────────────────────────────────
    p.wrap_para_pr = wrap_p.get('paraPrIDRef', p.wrap_para_pr)
    wrap_run = wrap_p.find(f'{{{HP}}}run')
    if wrap_run is not None:
        p.wrap_char_pr = wrap_run.get('charPrIDRef', p.wrap_char_pr)
        wrap_lsa = wrap_p.find(f'{{{HP}}}linesegarray')
        if wrap_lsa is not None and len(wrap_lsa) > 0:
            p.wrap_lineseg = dict(wrap_lsa[0].attrib)

    # ── 표 속성 ───────────────────────────────────────────────────────
    tbl = next(
        (t for r in (wrap_run,) if r is not None
         for t in r if t.tag == f'{{{HP}}}tbl' and t.get('colCnt') == '1'),
        None,
    )
    if tbl is None:
        return p

    p.border_fill_id = tbl.get('borderFillIDRef', p.border_fill_id)

    tbl_pos = tbl.find(f'{{{HP}}}pos')
    if tbl_pos is not None:
        p.vert_offset = int(tbl_pos.get('vertOffset', p.vert_offset))

    tbl_out = tbl.find(f'{{{HP}}}outMargin')
    if tbl_out is not None:
        p.out_margin = int(tbl_out.get('bottom', p.out_margin))

    # ── 행별 셀 크기·스타일 ───────────────────────────────────────────
    rows = tbl.findall(f'{{{HP}}}tr')
    if len(rows) < 2:
        return p

    def _from_tc(tc: etree._Element | None, is_title: bool) -> None:
        if tc is None:
            return
        csz = tc.find(f'{{{HP}}}cellSz')
        if csz is not None:
            p.cell_w = int(csz.get('width', p.cell_w))
            if is_title:
                p.title_row_h = int(csz.get('height', p.title_row_h))
            else:
                p.img_cell_h = int(csz.get('height', p.img_cell_h))
        cm = tc.find(f'{{{HP}}}cellMargin')
        if cm is not None:
            p.margin_lr = int(cm.get('left', p.margin_lr))
            p.margin_tb = int(cm.get('top',  p.margin_tb))
        sl = tc.find(f'{{{HP}}}subList')
        if sl is None:
            return
        para = sl.find(f'{{{HP}}}p')
        if para is None:
            return
        run = para.find(f'{{{HP}}}run')
        if is_title:
            p.title_para_pr  = para.get('paraPrIDRef', p.title_para_pr)
            p.title_style_id = para.get('styleIDRef',  p.title_style_id)
            if run is not None:
                p.title_char_pr = run.get('charPrIDRef', p.title_char_pr)
            lsa = para.find(f'{{{HP}}}linesegarray')
            if lsa is not None and len(lsa) > 0:
                p.title_lineseg = dict(lsa[0].attrib)
        else:
            p.img_para_pr  = para.get('paraPrIDRef', p.img_para_pr)
            p.img_style_id = para.get('styleIDRef',  p.img_style_id)
            if run is not None:
                p.img_char_pr = run.get('charPrIDRef', p.img_char_pr)

    _from_tc(rows[0].find(f'{{{HP}}}tc'), is_title=True)
    _from_tc(rows[1].find(f'{{{HP}}}tc'), is_title=False)
    return p


def _hp(parent: etree._Element, tag: str, attribs: dict | None = None) -> etree._Element:
    return etree.SubElement(parent, f'{{{HP}}}{tag}', attribs or {})


# ── 핵심 XML 빌더 ──────────────────────────────────────────────────────

def _build_pic(binary_id: str, img: Img, disp_w: int, disp_h: int, z: int) -> etree._Element:
    """인라인(treatAsChar=1) 이미지 요소를 만든다."""
    # 물리적 원본 크기: DPI 없으면 96 DPI 가정 (pyhwpx 기본값)
    with Image.open(img.path) as im:
        raw_dpi = im.info.get('dpi', (96, 96))
    dpi_x = float(raw_dpi[0]) or 96
    dpi_y = float(raw_dpi[1]) or 96
    org_w = round(img.w / dpi_x * 7200)  # 물리적 크기 (HWP unit) → imgClip/imgDim 용
    org_h = round(img.h / dpi_y * 7200)

    pic = etree.Element(f'{{{HP}}}pic', {
        'id':            str(_next_id()),
        'zOrder':        str(z),
        'numberingType': 'PICTURE',
        'textWrap':      'TOP_AND_BOTTOM',
        'textFlow':      'BOTH_SIDES',
        'lock':          '0',
        'dropcapstyle':  'None',
        'href':          '',
        'groupLevel':    '0',
        'instid':        str(_next_id()),
        'reverse':       '0',
    })

    _hp(pic, 'offset',  {'x': '0', 'y': '0'})
    # orgSz = 표시 크기 (pyhwpx 방식: 물리적 크기가 아닌 렌더링 크기)
    _hp(pic, 'orgSz',   {'width': str(disp_w), 'height': str(disp_h)})
    # curSz = 0,0 → "sz와 동일" 의미 (pyhwpx 동작)
    _hp(pic, 'curSz',   {'width': '0', 'height': '0'})
    _hp(pic, 'flip',    {'horizontal': '0', 'vertical': '0'})
    _hp(pic, 'rotationInfo', {
        'angle': '0', 'centerX': str(disp_w // 2), 'centerY': str(disp_h // 2),
        'rotateimage': '1',
    })

    ri = _hp(pic, 'renderingInfo')
    etree.SubElement(ri, f'{{{HC}}}transMatrix',
                     {'e1': '1', 'e2': '0', 'e3': '0', 'e4': '0', 'e5': '1', 'e6': '0'})
    # scaMatrix = identity (orgSz이 이미 표시 크기이므로 배율 불필요)
    etree.SubElement(ri, f'{{{HC}}}scaMatrix',
                     {'e1': '1', 'e2': '0', 'e3': '0', 'e4': '0', 'e5': '1', 'e6': '0'})
    etree.SubElement(ri, f'{{{HC}}}rotMatrix',
                     {'e1': '1', 'e2': '0', 'e3': '0', 'e4': '0', 'e5': '1', 'e6': '0'})

    # img는 imgRect 앞에 위치 (pyhwpx 요소 순서)
    etree.SubElement(pic, f'{{{HC}}}img', {
        'binaryItemIDRef': binary_id,
        'bright': '0', 'contrast': '0', 'effect': 'REAL_PIC', 'alpha': '0',
    })

    # imgRect 좌표는 orgSz(= 표시 크기) 기준
    img_rect = _hp(pic, 'imgRect')
    for tag, x, y in [('pt0', 0, 0), ('pt1', disp_w, 0), ('pt2', disp_w, disp_h), ('pt3', 0, disp_h)]:
        etree.SubElement(img_rect, f'{{{HC}}}{tag}', {'x': str(x), 'y': str(y)})

    # imgClip right/bottom = 물리적 원본 크기 (전체 이미지 사용, 크롭 없음)
    _hp(pic, 'imgClip',  {'left': '0', 'right': str(org_w), 'top': '0', 'bottom': str(org_h)})
    _hp(pic, 'inMargin', {'left': '0', 'right': '0', 'top': '0', 'bottom': '0'})
    # imgDim = 물리적 원본 크기 (pyhwpx가 생성하는 메타데이터)
    _hp(pic, 'imgDim',   {'dimwidth': str(org_w), 'dimheight': str(org_h)})
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


def _build_table(title: str, img_rows: list, z: int, params: _TemplateParams) -> tuple:
    """
    2행 1열 증빙 표 생성.
    img_rows: 행 리스트, 각 행 = [(binary_id, Img, disp_w_hwp, disp_h_hwp), ...]
    반환: (hp:tbl 요소, 다음 z 값)
    """
    m = str(params.margin_lr)
    mt = str(params.margin_tb)
    bf = params.border_fill_id

    tbl = etree.Element(f'{{{HP}}}tbl', {
        'id':              str(_next_id()),
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
        'borderFillIDRef': bf,
        'noAdjust':        '0',
    })
    z += 1

    om = str(params.out_margin)
    total_h = params.title_row_h + params.img_cell_h + params.out_margin
    _hp(tbl, 'sz',  {'width': str(params.cell_w), 'widthRelTo': 'ABSOLUTE',
                     'height': str(total_h), 'heightRelTo': 'ABSOLUTE', 'protect': '0'})
    _hp(tbl, 'pos', {
        'treatAsChar': '1', 'affectLSpacing': '0', 'flowWithText': '1',
        'allowOverlap': '0', 'holdAnchorAndSO': '0',
        'vertRelTo': 'PARA', 'horzRelTo': 'COLUMN',
        'vertAlign': 'TOP', 'horzAlign': 'LEFT',
        'vertOffset': str(params.vert_offset), 'horzOffset': '0',
    })
    _hp(tbl, 'outMargin', {'left': om, 'right': om, 'top': om, 'bottom': om})
    _hp(tbl, 'inMargin',  {'left': m, 'right': m, 'top': mt, 'bottom': mt})

    # ── 행 1: 제목 ────────────────────────────────────────────────────
    tr1 = _hp(tbl, 'tr')
    tc1 = _hp(tr1, 'tc', {'name': '', 'header': '0', 'hasMargin': '0',
                           'protect': '0', 'editable': '0', 'dirty': '0',
                           'borderFillIDRef': bf})
    sl1 = _hp(tc1, 'subList', {
        'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
        'vertAlign': 'CENTER', 'linkListIDRef': '0', 'linkListNextIDRef': '0',
        'textWidth': '0', 'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
    })
    p1  = _hp(sl1, 'p', {'id': str(_next_id()),
                          'paraPrIDRef': params.title_para_pr,
                          'styleIDRef':  params.title_style_id,
                          'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
    r1  = _hp(p1, 'run', {'charPrIDRef': params.title_char_pr})
    _hp(r1, 't').text = title
    lsa1 = _hp(p1, 'linesegarray')
    _hp(lsa1, 'lineseg', params.title_lineseg)
    _hp(tc1, 'cellAddr',   {'colAddr': '0', 'rowAddr': '0'})
    _hp(tc1, 'cellSpan',   {'colSpan': '1', 'rowSpan': '1'})
    _hp(tc1, 'cellSz',     {'width': str(params.cell_w), 'height': str(params.title_row_h)})
    _hp(tc1, 'cellMargin', {'left': m, 'right': m, 'top': mt, 'bottom': mt})

    # ── 행 2: 이미지 ──────────────────────────────────────────────────
    tr2 = _hp(tbl, 'tr')
    tc2 = _hp(tr2, 'tc', {'name': '', 'header': '0', 'hasMargin': '0',
                           'protect': '0', 'editable': '0', 'dirty': '0',
                           'borderFillIDRef': bf})
    sl2 = _hp(tc2, 'subList', {
        'id': '', 'textDirection': 'HORIZONTAL', 'lineWrap': 'BREAK',
        'vertAlign': 'CENTER', 'linkListIDRef': '0', 'linkListNextIDRef': '0',
        'textWidth': '0', 'textHeight': '0', 'hasTextRef': '0', 'hasNumRef': '0',
    })

    img_horzsize = params.title_lineseg.get('horzsize', '50856')
    if img_rows:
        for row in img_rows:
            p = _hp(sl2, 'p', {'id': '0',
                                'paraPrIDRef': params.img_para_pr,
                                'styleIDRef':  params.img_style_id,
                                'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
            run = _hp(p, 'run', {'charPrIDRef': params.img_char_pr})
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
                'horzpos': '0', 'horzsize': img_horzsize, 'flags': '393216',
            })
    else:
        p = _hp(sl2, 'p', {'id': '0',
                            'paraPrIDRef': params.img_para_pr,
                            'styleIDRef':  params.img_style_id,
                            'pageBreak': '0', 'columnBreak': '0', 'merged': '0'})
        _hp(p, 'run', {'charPrIDRef': params.img_char_pr})
        lsa = _hp(p, 'linesegarray')
        _hp(lsa, 'lineseg', {'textpos': '0', 'vertpos': '0', 'vertsize': '1000',
                              'textheight': '1000', 'baseline': '850', 'spacing': '600',
                              'horzpos': '0', 'horzsize': img_horzsize, 'flags': '393216'})

    _hp(tc2, 'cellAddr',   {'colAddr': '0', 'rowAddr': '1'})
    _hp(tc2, 'cellSpan',   {'colSpan': '1', 'rowSpan': '1'})
    _hp(tc2, 'cellSz',     {'width': str(params.cell_w), 'height': str(params.img_cell_h)})
    _hp(tc2, 'cellMargin', {'left': m, 'right': m, 'top': mt, 'bottom': mt})

    return tbl, z


# ── 공개 진입점 ────────────────────────────────────────────────────────

def run(data, t_path: str, o_path: str):
    """
    HWPX 증빙 서류 생성 (XML 직접 조작, pyhwpx 불필요).

    data      : image_downloader.run() 반환 DataFrame (img_paths 컬럼 포함)
    t_path    : 템플릿 .hwpx 파일 경로
    o_path    : 출력 .hwpx 파일 경로
    """
    # 0. ID 카운터 초기화 (다중 호출 시 결정론적 ID 보장)
    _reset_id_counter()

    # 1. 템플릿 ZIP 읽기 (.hwp 바이너리는 지원 불가)
    if not zipfile.is_zipfile(t_path):
        ext = os.path.splitext(t_path)[1]
        raise ValueError(
            f"템플릿 파일이 HWPX 형식이 아닙니다 ({ext}).\n"
            "HWP 바이너리 포맷은 지원하지 않습니다. "
            "한글에서 '다른 이름으로 저장 → .hwpx'로 변환 후 다시 시도하세요."
        )

    with zipfile.ZipFile(t_path, 'r') as zin:
        # ZipInfo 보존: 원본 압축 방식을 출력 시 그대로 유지
        zip_infos = {info.filename: info for info in zin.infolist()}
        zip_files = {name: zin.read(name) for name in zin.namelist()}

    # 2. section0.xml / content.hpf 존재 여부 확인
    SEC_KEY = 'Contents/section0.xml'
    HPF_KEY = 'Contents/content.hpf'
    if SEC_KEY not in zip_files:
        raise ValueError(f"템플릿 HWPX에 '{SEC_KEY}'가 없습니다. 올바른 HWPX 파일인지 확인하세요.")
    if HPF_KEY not in zip_files:
        raise ValueError(f"템플릿 HWPX에 '{HPF_KEY}'가 없습니다. 올바른 HWPX 파일인지 확인하세요.")

    # 3. section0.xml 파싱
    root = etree.fromstring(zip_files[SEC_KEY])

    # 4. 기존 증빙 표 단락 제거 (삽입 위치 기억)
    expense_ps = _find_expense_ps(root)
    params     = _read_template_params(expense_ps)
    insert_idx = list(root).index(expense_ps[0]) if expense_ps else len(list(root))
    for p in expense_ps:
        root.remove(p)

    # 5. 카운터 초기화
    bin_counter = _max_binary_idx(zip_files) + 1
    z           = _max_z_order(root) + 1
    new_binaries: dict[str, bytes] = {}

    # 6. 지출 행마다 표 생성 (data_idx = 장부 전체 기준 0-based 인덱스 → +1이 장부 순번)
    for data_idx, row in data.iterrows():
        title     = f'{data_idx + 1}. {row["종류"]}'
        raw = row.get('img_paths', [])
        if isinstance(raw, list):
            img_paths = raw
        elif raw is None:
            img_paths = []
        elif isinstance(raw, str):
            img_paths = [raw]
        else:
            try:
                img_paths = list(raw)
            except TypeError:
                img_paths = []

        img_rows: list[list] = []
        if img_paths:
            imgs = []
            for p in img_paths:
                try:
                    imgs.append(Img(p))
                except (OSError, UnidentifiedImageError, DecompressionBombError) as e:
                    print(f"  [{data_idx + 1}] 이미지 로드 실패: {p} ({e}) — 건너뜀")

            if not imgs:
                print(f"  [{data_idx + 1}] 유효한 이미지가 없어 이미지 셀 비워둠")
                layout = None
            else:
                layout = _layout(imgs, params.img_w_mm, params.img_h_mm)

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

                        try:
                            with open(img.path, 'rb') as f:
                                new_binaries[f'BinData/{bid}.{ext}'] = f.read()
                        except OSError as e:
                            print(f"  [{data_idx + 1}] 이미지 읽기 실패: {img.path} ({e}) — 건너뜀")
                            bin_counter -= 1
                            continue

                        disp_w = round(item.size[0] * HWP_PER_MM)
                        disp_h = round(item.size[1] * HWP_PER_MM)
                        row_items.append((bid, img, disp_w, disp_h))
                    if row_items:
                        img_rows.append(row_items)
            elif imgs:
                print(f"  [{data_idx + 1}] 레이아웃 계산 실패 — 이미지 셀 비워둠")
        else:
            print(f"  [{data_idx + 1}] 증빙 자료 누락 — 확인 필요")

        tbl_elem, z = _build_table(title, img_rows, z, params)

        # hp:p > hp:run > hp:tbl 구조로 래핑 (템플릿 패턴과 동일)
        p_wrap = etree.Element(f'{{{HP}}}p', {
            'id': '0', 'paraPrIDRef': params.wrap_para_pr, 'styleIDRef': '0',
            'pageBreak': '0', 'columnBreak': '0', 'merged': '0',
        })
        run_wrap = etree.SubElement(p_wrap, f'{{{HP}}}run', {'charPrIDRef': params.wrap_char_pr})
        run_wrap.append(tbl_elem)
        etree.SubElement(run_wrap, f'{{{HP}}}t')
        lsa = etree.SubElement(p_wrap, f'{{{HP}}}linesegarray')
        etree.SubElement(lsa, f'{{{HP}}}lineseg', params.wrap_lineseg)

        root.insert(insert_idx, p_wrap)
        insert_idx += 1

    # 7. XML 직렬화
    zip_files[SEC_KEY] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True,
    )
    for path, content in new_binaries.items():
        zip_files[path] = content

    # 8. content.hpf 매니페스트에 신규 이미지 등록
    if new_binaries:
        zip_files[HPF_KEY] = _update_content_hpf(zip_files[HPF_KEY], new_binaries)

    # 9. HWPX 패킹 (원본 ZipInfo 재사용 → 타임스탬프·외부속성 보존)
    os.makedirs(os.path.dirname(o_path) or '.', exist_ok=True)
    with zipfile.ZipFile(o_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, content in zip_files.items():
            orig = zip_infos.get(name)
            if orig:
                zout.writestr(orig, content)
            else:
                zout.writestr(name, content, compress_type=zipfile.ZIP_DEFLATED)

    print(f"HWPX 생성 완료: {o_path}")