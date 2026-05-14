"""
dxf_processor.py
Orchestrator: đọc DXF → trích INSERT → xác định line + size → tra spec → trả list items.
"""
import re
import ezdxf

from line_detector import (
    build_layer_colors, collect_pipes, detect_line_for_insert, IGNORE_BLOCKS,
)
from block_library import lookup_display_name, get_spec, SANITARY_LINE_TYPES
from size_detector import collect_size_annotations, nearest_size


def _dn_num(dn_str: str) -> int:
    m = re.search(r"\d+", str(dn_str))
    return int(m.group()) if m else -1


def process_dxf(dxf_path: str, progress_cb=None) -> tuple:
    """
    Đọc file DXF và trả về:
      items  : list of dict (mỗi INSERT thành 1 item)
      warnings: list of str (block không khớp thư viện, size=?)
    """
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    if progress_cb:
        progress_cb("Đang đọc layer colors...", 0.05)

    layer_colors = build_layer_colors(doc)

    if progress_cb:
        progress_cb("Đang thu thập đường ống...", 0.15)

    pipes = collect_pipes(msp, layer_colors)

    if progress_cb:
        progress_cb("Đang thu thập annotation size...", 0.25)

    annotations = collect_size_annotations(msp)

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
        pos = e.dxf.insert
        inserts.append({"block_name": block_name, "x": pos.x, "y": pos.y})

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
