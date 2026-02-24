from math import ceil
from PIL import Image, ImageOps

class Img:
    def __init__(self, path: str):
        self.path = path
        ImageOps.exif_transpose(Image.open(path)).save(path, exif=b"")
        self.w, self.h = Image.open(self.path).size

class LayoutItem:
    def __init__(
            self,
            idx: int,
            path: str,
            size: tuple[float, float],
        ):
        self.idx = idx      # 원본 리스트에서의 인덱스
        self.path = path    # 이미지 파일 경로
        self.size = size    # 배율이 적용된 크기 (너비, 높이)

class LayoutResult:
    def __init__(
            self,
            area: float = -1,
            items: list[LayoutItem] | None = None,
            grid: tuple | None = None,
        ):
        self.area = area            # 이미지들의 총 면적
        self.items = items or []    # 개별 이미지 정보 리스트
        self.grid = grid            # 그리드 구조 (행, 열)

def _layout(imgs: list[Img], cw: float, ch: float) -> LayoutResult | None:
    n = len(imgs)
    if n == 0: return None
    
    best = LayoutResult()
    
    for cols in range(1, min(n + 1, 6)):
        rows = ceil(n / cols)
        cell_w, cell_h = cw / cols, ch / rows
        
        # 모든 이미지의 최대 스케일 계산
        scales = [min(cell_w / img.w, cell_h / img.h) for img in imgs]
        
        median_scale = sorted(scales)[n // 2]
        
        items = []
        for i, img in enumerate(imgs):
            scale = min(median_scale, cell_w / img.w, cell_h / img.h)
            w, h = img.w * scale, img.h * scale
            x, y = (i % cols) * cell_w, (i // cols) * cell_h
            items.append(LayoutItem(i, img.path, (w, h)))
        
        area = sum(item.size[0] * item.size[1] for item in items)
        if len(items) == n and area > best.area:
            best = LayoutResult(area, items, (rows, cols))
    
    return best if len(best.items) == n else None


def _get_cell(hwp):
    a = hwp.CreateAction('TablePropertyDialog')
    s = a.CreateSet()
    p = s.CreateItemSet('ShapeTableCell', 'Cell')
    a.GetDefault(s)

    return hwp.HwpUnitToMili(p.Item('Width')), hwp.HwpUnitToMili(p.Item('Height'))


def _pack(hwp, layout: LayoutResult):
    if not layout.items or not layout.grid: return
    
    rows, cols = layout.grid
    for r in range(rows):
        for c in range(cols):
            idx = r * cols + c
            if idx >= len(layout.items): break

            item = layout.items[idx]
            hwp.insert_picture(
                item.path,
                sizeoption=1,
                width=item.size[0],
                height=item.size[1],
                treat_as_char=True
            )

        if r < rows - 1 and (r + 1) * cols < len(layout.items):
            hwp.BreakPara()


def pack(hwp, paths):
    if not hwp.ParentCtrl or hwp.ParentCtrl.CtrlID != 'tbl': return False
    cw, ch = _get_cell(hwp)

    imgs = [Img(path) for path in paths]
    if not imgs: return False
    
    layout = _layout(imgs, cw, ch)
    if not layout: return False

    _pack(hwp, layout)
    return True
