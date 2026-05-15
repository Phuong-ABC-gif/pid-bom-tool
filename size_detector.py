"""
size_detector.py  —  ANH MINH 2021
Xác định size thiết bị theo annotation gần nhất trong DXF.

Các format được nhận diện:
  DN format   : DN50  DN 50  dn100
  Double quote: 2"  1.5"  1/2"  3/4"  1 1/2"  2 1/2"
  Two singles : 2''  1.5''  1/2''  1 1/2''
  Phi/Ø       : Ø50  ø25  Φ100
  mm suffix   : 50mm  50 mm

Kết quả trả về chuỗi chuẩn: 'DN50' hoặc '1.5"' hoặc '1/2"'
"""
import math
import re

MAX_DIST = 300   # đơn vị DXF — khoảng cách tối đa annotation → INSERT

# ── Pattern (ưu tiên từ cao → thấp) ──────────────────────────────────────
_INCH = r'(?:"|\'\')'                              # " hoặc ''

PAT_DN       = re.compile(r'DN\s*(\d+(?:\.\d+)?)', re.IGNORECASE)
PAT_COMPOUND = re.compile(rf'(\d+)\s+(\d+)/(\d+)\s*{_INCH}')   # 1 1/2" | 1 1/2''
PAT_FRACTION = re.compile(rf'(\d+)/(\d+)\s*{_INCH}')            # 1/2"   | 1/2''
PAT_DECIMAL  = re.compile(rf'(\d+(?:\.\d+)?)\s*{_INCH}')        # 1.5"   | 2''
PAT_PHI      = re.compile(r'[ØøΦφ]\s*(\d+(?:\.\d+)?)')          # Ø50
PAT_MM       = re.compile(r'(\d+(?:\.\d+)?)\s*mm', re.IGNORECASE)  # 50mm


def _clean_mtext(raw: str) -> str:
    """Xóa MTEXT formatting codes (\\P, \\H, \\W, {{...}}, v.v.)."""
    raw = re.sub(r'\\[A-Za-z][^;]*;', '', raw)
    raw = re.sub(r'\{[^}]*\}', '', raw)
    raw = raw.replace('\\P', ' ').replace('\\n', ' ')
    return raw.strip()


def parse_size(text: str) -> str | None:
    """
    Nhận chuỗi văn bản → trả về chuỗi size chuẩn hoặc None nếu không nhận ra.

    Thứ tự ưu tiên:
      1. DN  → 'DN50'
      2. Compound fraction inch  → '1.5"'
      3. Simple fraction inch    → '1/2"'
      4. Decimal / integer inch  → '2"' / '1.5"'
      5. Ø / Φ (mm)             → 'DN50'
      6. mm suffix               → 'DN50'
    """
    s = _clean_mtext(text)

    # 1. DN format
    m = PAT_DN.search(s)
    if m:
        return f'DN{m.group(1)}'

    # 2. Compound fraction: 1 1/2" hoặc 1 1/2''
    m = PAT_COMPOUND.search(s)
    if m:
        whole = int(m.group(1))
        num   = int(m.group(2))
        den   = int(m.group(3))
        val   = whole + num / den
        # Giữ dạng thập phân gọn: 1.5" thay vì 1.5000"
        val_str = str(int(val)) if val == int(val) else f'{val:.4g}'
        return f'{val_str}"'

    # 3. Simple fraction: 1/2" hoặc 3/4''
    m = PAT_FRACTION.search(s)
    if m:
        return f'{m.group(1)}/{m.group(2)}"'

    # 4. Decimal/integer: 1.5" hoặc 2''
    m = PAT_DECIMAL.search(s)
    if m:
        return f'{m.group(1)}"'

    # 5. Ø / Φ symbol (mm → DN)
    m = PAT_PHI.search(s)
    if m:
        return f'DN{int(float(m.group(1)))}'

    # 6. mm suffix
    m = PAT_MM.search(s)
    if m:
        return f'DN{int(float(m.group(1)))}'

    return None


# ── Thu thập annotation từ modelspace ────────────────────────────────────

def collect_size_annotations(msp, wcs2ucs=None) -> list:
    """
    Quét TEXT/MTEXT trong modelspace, trả về list:
      [{'x': ..., 'y': ..., 'dn': '1.5"' | 'DN50', 'raw': ...}, ...]
    Tọa độ đã chuyển sang UCS.
    """
    from ezdxf.math import Matrix44 as _M44, Vec3
    if wcs2ucs is None:
        wcs2ucs = _M44()

    def to_ucs(x, y):
        p = wcs2ucs.transform(Vec3(x, y, 0))
        return p.x, p.y

    annotations = []
    for e in msp:
        if e.dxftype() not in ("TEXT", "MTEXT"):
            continue
        try:
            raw = e.text if e.dxftype() == "MTEXT" else e.dxf.text
            pos = e.dxf.insert
            ux, uy = to_ucs(pos.x, pos.y)

            size = parse_size(raw)
            # Debug: log tất cả TEXT/MTEXT và kết quả parse
            import sys
            print(f"[SIZE_DEBUG] type={e.dxftype()} raw={repr(raw)[:80]} → parse={repr(size)}", file=sys.stderr)
            if size:
                annotations.append({"x": ux, "y": uy, "dn": size, "raw": raw})
        except Exception:
            continue
    return annotations


def nearest_size(eq_x: float, eq_y: float, annotations: list) -> str:
    """Trả về size của annotation gần nhất. Nếu vượt MAX_DIST → '?'."""
    best, best_d = None, float("inf")
    for ann in annotations:
        d = math.hypot(ann["x"] - eq_x, ann["y"] - eq_y)
        if d < best_d:
            best_d = d
            best = ann
    if best and best_d <= MAX_DIST:
        return best["dn"]
    return "?"
