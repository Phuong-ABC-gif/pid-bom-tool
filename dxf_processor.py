"""
dxf_processor.py
Orchestrator: đọc DXF → trích INSERT → xác định line + size → tra spec → trả list items.

Ghi chú tọa độ:
  AutoCAD hiển thị tọa độ theo UCS (User Coordinate System).
  ezdxf đọc tọa độ theo WCS (World Coordinate System).
  Nếu bản vẽ dùng custom UCS (origin offset, trục lật...) thì 2 hệ lệch nhau
  → dẫn đến dấu bị đảo hoặc tọa độ sai so với màn hình AutoCAD.
  Fix: đọc UCS từ DXF header ($UCSORG / $UCSXDIR / $UCSYDIR) rồi
  transform WCS → UCS để tọa độ khớp với màn hình.
"""
import re
import ezdxf
from ezdxf.math import Vec3, Matrix44

from line_detector import (
    build_layer_colors, collect_pipes, detect_line_for_insert, IGNORE_BLOCKS,
)
from block_library import lookup_display_name, get_spec, SANITARY_LINE_TYPES
from size_detector import collect_size_annotations, nearest_size


def _dn_num(dn_str: str) -> int:
    m = re.search(r"\d+", str(dn_str))
    return int(m.group()) if m else -1


def build_wcs_to_ucs(doc) -> Matrix44:
    """
    Đọc UCS hiện hành từ DXF header và trả về ma trận chuyển WCS → UCS.
    Nếu không có UCS header (bản vẽ dùng WCS mặc định) → trả về identity.
    """
    hdr = doc.header
    origin = Vec3(hdr.get("$UCSORG",  (0.0, 0.0, 0.0)))
    x_axis = Vec3(hdr.get("$UCSXDIR", (1.0, 0.0, 0.0)))
    y_axis = Vec3(hdr.get("$UCSYDIR", (0.0, 1.0, 0.0)))

    # Kiểm tra UCS có phải WCS mặc định không
    is_default = (
        origin.isclose((0, 0, 0), abs_tol=1e-6)
        and x_axis.isclose((1, 0, 0), abs_tol=1e-6)
        and y_axis.isclose((0, 1, 0), abs_tol=1e-6)
    )
    if is_default:
        return Matrix44()  # identity — không cần transform

    # Xây ma trận rotation từ trục UCS (UCS axes in WCS)
    z_axis = x_axis.cross(y_axis).normalize()
    # Ma trận chuyển WCS point P_wcs → P_ucs:
    #   P_ucs = R^T * (P_wcs - origin)
    # R^T rows = [x_axis, y_axis, z_axis]
    # Matrix44 nhận flat list 16 phần tử (row-major)
    rot_inv = Matrix44([
        x_axis.x, x_axis.y, x_axis.z, 0.0,
        y_axis.x, y_axis.y, y_axis.z, 0.0,
        z_axis.x, z_axis.y, z_axis.z, 0.0,
        0.0,      0.0,      0.0,      1.0,
    ])
    translate = Matrix44.translate(-origin.x, -origin.y, -origin.z)
    return rot_inv @ translate


def wcs_to_ucs(pt_wcs: tuple, m: Matrix44) -> tuple:
    """Chuyển 1 điểm WCS sang UCS."""
    p = m.transform(Vec3(pt_wcs))
    return (p.x, p.y)


def process_dxf(dxf_path: str, progress_cb=None) -> tuple:
    """
    Đọc file DXF và trả về:
      items  : list of dict (mỗi INSERT thành 1 item)
      warnings: list of str (block không khớp thư viện, size=?)
    """
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    if progress_cb:
        progress_cb("Đang đọc UCS & layer colors...", 0.05)

    # Ma trận WCS → UCS (fix tọa độ khớp với màn hình AutoCAD)
    wcs2ucs = build_wcs_to_ucs(doc)

    layer_colors = build_layer_colors(doc)

    if progress_cb:
        progress_cb("Đang thu thập đường ống...", 0.15)

    pipes = collect_pipes(msp, layer_colors, wcs2ucs)

    if progress_cb:
        progress_cb("Đang thu thập annotation size...", 0.25)

    annotations = collect_size_annotations(msp, wcs2ucs)

    if progress_cb:
        progress_cb("Đang xử lý INSERT blocks...", 0.40)

    items = []
    warnings = []

    # Thu thập tất cả INSERT
    inserts = []
    for e in msp:
        if e.dxftype() != "INSERT":
            continue
        block_name = e.dxf.name or ""
        if block_name.lower() in IGNORE_BLOCKS:
            continue
        if not block_name.strip():
            continue
        # Đọc tọa độ WCS từ dxf.insert rồi chuyển sang UCS
        raw = e.dxf.insert
        wx, wy = wcs_to_ucs((raw.x, raw.y, 0.0), wcs2ucs)
        inserts.append({"block_name": block_name, "x": wx, "y": wy})

    total = len(inserts)
    for i, ins in enumerate(inserts):
        if progress_cb and i % max(1, total // 20) == 0:
            pct = 0.40 + 0.50 * (i / max(total, 1))
            progress_cb(f"Xử lý block {i+1}/{total}: {ins['block_name']}", pct)

        block_name = ins["block_name"]
        x, y = ins["x"], ins["y"]

        # 1. Xác định line (màu sắc đường ống gần nhất)
        line_info = detect_line_for_insert(x, y, pipes)

        # 2. Display name từ block name
        display_name = lookup_display_name(block_name)
        if display_name == block_name:
            warnings.append(f"Block không có trong thư viện (vẫn đưa vào BOM): «{block_name}»")

        # 3. Size gần nhất
        size = nearest_size(x, y, annotations)
        if size == "?":
            warnings.append(f"Không xác định được size cho: «{block_name}» tại ({x:.0f},{y:.0f})")

        # 4. Tra spec
        spec = get_spec(display_name, line_info["line_type"], size)

        items.append({
            "block_name":   block_name,
            "display_name": display_name,
            "line_type":    line_info["line_type"],
            "line_label":   line_info["label"],
            "group":        line_info["group"],
            "size":         size,
            "chung_loai":   spec.get("chung_loai", "-"),
            "vat_lieu":     spec.get("vat_lieu", "-"),
            "tieu_chuan":   spec.get("tieu_chuan", "-"),
            "don_vi":       spec.get("don_vi", "pcs"),
            "x": x, "y": y,
        })

    if progress_cb:
        progress_cb("Hoàn tất xử lý DXF.", 1.0)

    return items, warnings
