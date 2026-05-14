"""
size_detector.py
Kỹ năng 1 — Xác định size thiết bị theo vị trí annotation gần nhất.
Skill ANH MINH 2021.
"""
import math
import re
import ezdxf

MAX_DIST = 300  # đơn vị DXF


def _clean_text(raw: str) -> str:
    """Xóa MTEXT formatting codes."""
    raw = re.sub(r"\\[A-Za-z][^;]*;", "", raw)
    raw = re.sub(r"\{[^}]*\}", "", raw)
    return raw.strip()


def collect_size_annotations(msp, wcs2ucs=None) -> list:
    """Thu thập tất cả TEXT/MTEXT có chứa DN hoặc inch annotation (tọa độ UCS)."""
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
            raw = _clean_text(raw)
            pos = e.dxf.insert
            ux, uy = to_ucs(pos.x, pos.y)

            m_dn = re.search(r"DN\s*(\d+)", raw, re.IGNORECASE)
            if m_dn:
                annotations.append({"x": ux, "y": uy, "dn": f"DN{m_dn.group(1)}", "raw": raw})
                continue

            m_inch = re.search(r'(\d+(?:[./]\d+)?)\s*"', raw)
            if m_inch:
                annotations.append({"x": ux, "y": uy, "dn": f'{m_inch.group(1)}"', "raw": raw})
        except Exception:
            continue
    return annotations


def nearest_size(eq_x: float, eq_y: float, annotations: list) -> str:
    """Trả về chuỗi size gần nhất. Nếu vượt MAX_DIST → '?'."""
    best, best_d = None, float("inf")
    for ann in annotations:
        d = math.hypot(ann["x"] - eq_x, ann["y"] - eq_y)
        if d < best_d:
            best_d = d
            best = ann
    if best and best_d <= MAX_DIST:
        return best["dn"]
    return "?"
