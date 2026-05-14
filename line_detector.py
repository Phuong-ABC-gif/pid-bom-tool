"""
line_detector.py
Xác định loại đường ống dựa vào màu sắc (ACI index hoặc RGB true color).
Quy tắc 1 — Skill ANH MINH 2021
"""
import math
import re
import ezdxf

# ── Màu ACI chuẩn (AutoCAD Color Index) ───────────────────────────────────
ACI_RGB = {
    1:  (255, 0,   0),    # Red
    2:  (255, 255, 0),    # Yellow
    3:  (0,   255, 0),    # Green
    4:  (0,   255, 255),  # Cyan
    5:  (0,   0,   255),  # Blue
    6:  (255, 0,   255),  # Magenta
    7:  (255, 255, 255),  # White
    30: (255, 127, 0),    # Orange-red  (Steam index 30)
}

# ── Bảng nhận diện line theo màu ──────────────────────────────────────────
# Mỗi entry: (rgb_tuple, tolerance, line_type, group, label_vn)
LINE_COLOR_RULES = [
    ((255, 0,   255), 30, "CIP",              "sanitary", "CIP (Magenta)"),
    ((255, 0,   0),   30, "Product",          "sanitary", "Product (Red)"),
    ((0,   255, 255), 30, "RO/Process Water", "sanitary", "RO/Process Water (Cyan)"),
    ((0,   0,   255), 30, "Compressed Air",   "industry", "Compressed Air (Blue)"),
    ((255, 127, 0),   40, "Steam/Condensate", "industry", "Steam & Condensate (ACI-30)"),
    ((0,   166, 166), 25, "Ice Water",        "industry", "Ice Water (RGB 0,166,166)"),
    ((0,   255, 0),   30, "Cooling Water",    "industry", "Cooling Water (Green)"),
    ((2,   237, 233), 25, "City/Soft Water",  "industry", "City/Soft Water (RGB 2,237,233)"),
]

UNKNOWN_LINE = {"line_type": "Unknown", "group": "unknown", "label": "Unknown"}


def _rgb_distance(a, b):
    return math.sqrt(sum((x - y) ** 2 for x, y in zip(a, b)))


def aci_to_rgb(aci: int):
    """Trả về RGB tuple từ ACI index (chỉ hỗ trợ các màu đã map)."""
    return ACI_RGB.get(aci)


def color_to_line(rgb: tuple) -> dict:
    """So khớp RGB → line type."""
    if rgb is None:
        return UNKNOWN_LINE
    best, best_d = None, float("inf")
    for rule_rgb, tol, lt, grp, lbl in LINE_COLOR_RULES:
        d = _rgb_distance(rgb, rule_rgb)
        if d < best_d:
            best_d = d
            best = (lt, grp, lbl, tol)
    if best and best_d <= best[3]:
        return {"line_type": best[0], "group": best[1], "label": best[2]}
    return UNKNOWN_LINE


def get_entity_color(entity, layer_colors: dict):
    """Lấy màu thực sự của entity (ưu tiên true_color → ACI → BYLAYER)."""
    # True color (RGB 24-bit)
    if hasattr(entity.dxf, "true_color") and entity.dxf.hasattr("true_color"):
        tc = entity.dxf.true_color
        r = (tc >> 16) & 0xFF
        g = (tc >> 8)  & 0xFF
        b = tc & 0xFF
        return (r, g, b)
    # ACI color
    color_val = None
    if entity.dxf.hasattr("color"):
        color_val = entity.dxf.color
    if color_val is not None and color_val not in (0, 256):  # 0=BYBLOCK 256=BYLAYER
        return aci_to_rgb(color_val)
    # BYLAYER
    layer_name = entity.dxf.layer if entity.dxf.hasattr("layer") else "0"
    return layer_colors.get(layer_name)


def build_layer_colors(doc) -> dict:
    """Xây dựng dict {layer_name: rgb_tuple} từ layer table."""
    lc = {}
    for layer in doc.layers:
        name = layer.dxf.name
        if layer.dxf.hasattr("true_color"):
            tc = layer.dxf.true_color
            lc[name] = ((tc >> 16) & 0xFF, (tc >> 8) & 0xFF, tc & 0xFF)
        elif layer.dxf.hasattr("color"):
            lc[name] = aci_to_rgb(layer.dxf.color)
    return lc


# ── Tập các entity đường ống ──────────────────────────────────────────────
PIPE_TYPES = {"LINE", "LWPOLYLINE", "POLYLINE", "SPLINE", "ARC"}

# Block bỏ qua (không phải thiết bị)
IGNORE_BLOCKS = {
    "flow arrow", "linenamettpleft-1", "linenametpleft-1", "linenametpleft",
    "pid", "objects", "title block", "titleblock", "north arrow", "revision",
    "border", "frame",
}


def collect_pipes(msp, layer_colors: dict) -> list:
    """Thu thập tất cả đoạn đường ống với màu và tọa độ trung điểm."""
    pipes = []
    for e in msp:
        if e.dxftype() not in PIPE_TYPES:
            continue
        rgb = get_entity_color(e, layer_colors)
        line_info = color_to_line(rgb)
        if line_info["line_type"] == "Unknown":
            continue
        # Lấy tọa độ đại diện
        try:
            if e.dxftype() == "LINE":
                s, end = e.dxf.start, e.dxf.end
                cx, cy = (s.x + end.x) / 2, (s.y + end.y) / 2
            elif e.dxftype() == "LWPOLYLINE":
                pts = list(e.get_points())
                cx = sum(p[0] for p in pts) / len(pts)
                cy = sum(p[1] for p in pts) / len(pts)
            else:
                cx, cy = e.dxf.insert.x, e.dxf.insert.y
        except Exception:
            continue
        pipes.append({"x": cx, "y": cy, "rgb": rgb, "line_info": line_info})
    return pipes


def detect_line_for_insert(insert_x: float, insert_y: float, pipes: list,
                            max_dist: float = 500) -> dict:
    """Tìm đường ống gần nhất với INSERT và trả về line_info."""
    best, best_d = None, float("inf")
    for p in pipes:
        d = math.hypot(p["x"] - insert_x, p["y"] - insert_y)
        if d < best_d:
            best_d = d
            best = p
    if best and best_d <= max_dist:
        return best["line_info"]
    return UNKNOWN_LINE
