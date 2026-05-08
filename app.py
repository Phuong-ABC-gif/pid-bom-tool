"""
P&ID BOM Extractor — SPX & Tetra Pak (ANH MINH 2021)
=====================================================
• Import file DXF
• Chọn skill (SPX / Tetra Pak)
• Đọc block name trực tiếp → tra cứu database nội bộ (KHÔNG dùng AI, không tốn cost)
• Xác định size theo annotation gần nhất (proximity detection)
• Xuất BOM dạng Excel (.xlsx) chuẩn ANH MINH

Run: streamlit run app.py
"""

import io
import math
import re
import tempfile
from collections import defaultdict
from pathlib import Path

import ezdxf
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from pid_database import DATABASES, is_ignore_block, lookup_block

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="P&ID BOM Extractor",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# STYLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title  { font-size:2rem; font-weight:700; color:#1a5276; margin-bottom:0; }
    .sub-title   { font-size:1rem; color:#5d6d7e; margin-top:0; }
    .section-box { background:#eaf4fb; border-left:4px solid #2e86c1;
                   padding:8px 14px; border-radius:4px; margin:8px 0; }
    .warn-box    { background:#fef9e7; border-left:4px solid #f0b429;
                   padding:8px 14px; border-radius:4px; margin:8px 0; }
    .ok-box      { background:#eafaf1; border-left:4px solid #27ae60;
                   padding:8px 14px; border-radius:4px; margin:8px 0; }
    .stat-num    { font-size:1.8rem; font-weight:700; color:#2e86c1; }
    .stat-label  { font-size:0.85rem; color:#7f8c8d; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="main-title">🔧 P&ID BOM Extractor</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">SPX & Tetra Pak — ANH MINH 2021 | Đọc block name trực tiếp, không dùng AI</p>',
            unsafe_allow_html=True)
st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Cài đặt")

    skill_choice = st.selectbox(
        "📚 Chọn Skill / Tiêu chuẩn",
        list(DATABASES.keys()),
        help="SPX dùng model SV41/42/43/44. Tetra Pak dùng SSV 200/210/220/300 không kèm model."
    )

    st.markdown("---")
    st.subheader("🔍 Proximity Detection")
    max_dist = st.slider(
        "Khoảng cách tối đa tìm size (đơn vị DXF)",
        min_value=50, max_value=1000, value=300, step=50,
        help="Nếu annotation size xa hơn ngưỡng này, size sẽ được đánh dấu '?'."
    )

    st.markdown("---")
    st.subheader("📋 Tên dự án")
    project_name = st.text_input("Tên dự án", value="ANH MINH 2021",
                                  placeholder="Nhập tên dự án...")
    drawing_no   = st.text_input("Số bản vẽ",  value="AM-SB-11/2021",
                                  placeholder="Số bản vẽ...")

    st.markdown("---")
    st.caption("v1.0 | Không dùng AI | Không tốn cost")


# ─────────────────────────────────────────────────────────────────────────────
# DXF PARSING FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────
SIZE_PATTERN = re.compile(
    r'DN\s*(\d+)'          # DN50, DN 50
    r'|(\d+(?:\.\d+)?)\s*"'  # 2", 1.5"
    r'|(\d+(?:\.\d+)?)\s*inch',  # 2 inch
    re.IGNORECASE
)

def normalize_size(raw_text: str) -> str:
    """Chuẩn hóa text size → dạng 'DN50' hoặc '2"'."""
    m = SIZE_PATTERN.search(raw_text)
    if not m:
        return None
    if m.group(1):              # DN format
        return f"DN{m.group(1)}"
    elif m.group(2):            # inch format
        return f'{m.group(2)}"'
    elif m.group(3):
        return f'{m.group(3)}"'
    return None


def parse_dxf(file_bytes: bytes, skill_key: str, max_dist: float):
    """
    Đọc DXF từ bytes.
    Trả về:
        equipment_list: list of dict {block_name, x, y, size, info}
        unknown_blocks: list of dict {block_name, count}
        size_annotations: list of dict {text, x, y, dn}
        errors: list of str
    """
    errors = []
    equipment_list = []
    size_annotations = []

    # ── Đọc file ─────────────────────────────────────────────────────────────
    try:
        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name

        doc = ezdxf.readfile(tmp_path)
        msp = doc.modelspace()
    except Exception as e:
        errors.append(f"Lỗi đọc DXF: {e}")
        return [], {}, [], errors

    # ── Quét entities ─────────────────────────────────────────────────────────
    raw_equipment = []  # {name, x, y}

    for entity in msp:
        etype = entity.dxftype()

        if etype == "INSERT":
            try:
                name = entity.dxf.name
                if is_ignore_block(name):
                    continue
                pos = entity.dxf.insert
                raw_equipment.append({"name": name, "x": float(pos.x), "y": float(pos.y)})
            except Exception:
                pass

        elif etype in ("TEXT", "MTEXT"):
            try:
                if etype == "MTEXT":
                    raw_text = entity.text
                    # strip MTEXT formatting codes
                    raw_text = re.sub(r'\{\\[^}]*\}|\\[A-Za-z]+;?', '', raw_text)
                else:
                    raw_text = entity.dxf.text

                pos = entity.dxf.insert
                dn = normalize_size(raw_text)
                if dn:
                    size_annotations.append({
                        "text": raw_text.strip(),
                        "dn":   dn,
                        "x":    float(pos.x),
                        "y":    float(pos.y),
                    })
            except Exception:
                pass

    # ── Proximity: gán size cho từng thiết bị ─────────────────────────────────
    unknown_blocks = defaultdict(int)

    for eq in raw_equipment:
        info = lookup_block(eq["name"], skill_key)
        size = find_nearest_size(eq["x"], eq["y"], size_annotations, max_dist)
        if info:
            equipment_list.append({
                "block_name": eq["name"],
                "x": eq["x"],
                "y": eq["y"],
                "size":  size,
                "info":  info,
                "is_unknown": False,
            })
        else:
            # Fallback: dùng tên block làm tên thiết bị, vào section "unknown"
            unknown_blocks[eq["name"]] += 1
            equipment_list.append({
                "block_name": eq["name"],
                "x": eq["x"],
                "y": eq["y"],
                "size": size,
                "info": {
                    "desc":       eq["name"],   # tên block = tên thiết bị
                    "code":       eq["name"],
                    "material":   "",
                    "standard":   "",
                    "unit":       "EA",
                    "section":    "unknown",
                    "subsection": "",
                },
                "is_unknown": True,
            })

    return equipment_list, dict(unknown_blocks), size_annotations, errors


def find_nearest_size(ex: float, ey: float, size_texts: list, max_dist: float) -> str:
    """Tìm annotation size gần nhất trong bán kính max_dist."""
    best, best_d = None, float("inf")
    for t in size_texts:
        d = math.hypot(t["x"] - ex, t["y"] - ey)
        if d < best_d:
            best_d, best = d, t
    if best and best_d <= max_dist:
        return best["dn"]
    return "?"


# ─────────────────────────────────────────────────────────────────────────────
# BOM BUILDER
# ─────────────────────────────────────────────────────────────────────────────
SECTION_ORDER = ["sanitary", "utility", "steam", "instrument", "unknown"]
SUBSECTION_LABELS = {
    "process":    "Đường Product / CIP / Process Water",
    "chiller":    "Đường Chiller (Ice Water)",
    "cooling":    "Đường Cooling Water",
    "steam":      "Đường Steam",
    "":           "Thiết bị khác",
}
SECTION_LABELS = {
    "sanitary":   "Thiết bị Vi Sinh (Sanitary)",
    "utility":    "Thiết bị Công Nghiệp (Utility)",
    "steam":      "Thiết bị Đường Hơi (Steam)",
    "instrument": "Thiết bị Đo Lường (Instrument)",
    "unknown":    "Thiết bị Chưa Phân Loại (Cần Kiểm Tra)",
}


def build_bom(equipment_list: list) -> dict:
    """
    Nhóm thiết bị theo section → subsection → (desc, code, material, standard, unit, size).
    Trả về nested dict.
    """
    # Nhóm: section → subsection → (desc, code, mat, std, unit) → {size: count}
    groups = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(int))))

    for eq in equipment_list:
        info = eq["info"]
        key = (
            info["desc"],
            info["code"],
            info["material"],
            info["standard"],
            info["unit"],
        )
        sec    = info.get("section", "utility")
        subsec = info.get("subsection", "")
        # steam section: utility items with subsection='steam' stay in utility
        # but we keep them distinguishable
        size = eq["size"] or "?"
        groups[sec][subsec][key][size] += 1

    return groups


def flatten_bom(groups: dict) -> list:
    """
    Chuyển groups → list of row dicts sẵn sàng ghi Excel.
    Mỗi row: stt, mo_ta, chung_loai, vat_lieu, kthuoc1, kthuoc2, tieu_chuan, xuat_xu, don_vi, so_luong
    Cộng thêm các row header (section, subsection).
    """
    rows = []
    section_num = 0
    ROMAN = ["I", "II", "III", "IV", "V", "VI"]

    for sec in SECTION_ORDER:
        if sec not in groups:
            continue
        section_num += 1
        rom = ROMAN[section_num - 1] if section_num <= len(ROMAN) else str(section_num)
        rows.append({
            "stt": rom,
            "mo_ta": f"PHẦN {SECTION_LABELS.get(sec, sec).upper()}",
            "row_type": "section",
        })

        subsec_num = 0
        for subsec in ["process", "chiller", "cooling", "steam", ""]:
            if subsec not in groups[sec]:
                continue
            subsec_num += 1
            if subsec:  # only add sub-header if there's a named subsection
                rows.append({
                    "stt": str(subsec_num),
                    "mo_ta": SUBSECTION_LABELS.get(subsec, subsec),
                    "row_type": "subsection",
                })

            # Sort items by desc
            items = sorted(groups[sec][subsec].items(), key=lambda x: x[0][0])
            item_letter_idx = 0
            LETTERS = "abcdefghijklmnopqrstuvwxyz"

            for key, sizes in items:
                if not isinstance(key, tuple) or len(key) != 5:
                    continue
                desc, code, mat, std, unit = key
                for size, count in sorted(sizes.items(), key=lambda x: x[0]):
                    ltr = LETTERS[item_letter_idx % 26] if item_letter_idx < 26 else \
                          LETTERS[item_letter_idx // 26 - 1] + LETTERS[item_letter_idx % 26]
                    rows.append({
                        "stt":        ltr,
                        "mo_ta":      desc,
                        "chung_loai": code,
                        "vat_lieu":   mat,
                        "kthuoc1":    size if size != "?" else "",
                        "kthuoc2":    "",
                        "tieu_chuan": std,
                        "xuat_xu":    "",
                        "don_vi":     unit,
                        "so_luong":   count,
                        "row_type":   "data",
                        "size_unknown": size == "?",
                    })
                    item_letter_idx += 1

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
# Colors from skill spec
CLR_HEADER   = "00CCFF"   # Column header
CLR_SECTION  = "00FFFF"   # Section I / II / ...
CLR_SUBSEC   = "E0FFFF"   # Subsection 1 / 2 / ...
CLR_DATA     = "FFFFFF"   # Data row
CLR_UNKNOWN  = "FFF3CD"   # Row with unknown size (warning)

COLS = ["STT", "MÔ TẢ", "CHỦNG LOẠI", "VẬT LIỆU",
        "K.THƯỚC 1", "K.THƯỚC 2", "TIÊU CHUẨN", "XUẤT XỨ", "ĐƠN VỊ", "SỐ LƯỢNG"]
COL_WIDTHS = [8, 50, 20, 20, 12, 12, 14, 12, 10, 12]


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, italic=False, size=10) -> Font:
    return Font(name="Arial", bold=bold, italic=italic, size=size)


def export_excel(bom_rows: list, project_name: str, drawing_no: str, skill_key: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"

    # ── Title rows ────────────────────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    ws["A1"] = f"BẢNG KÊ VẬT TƯ (BOM) — {skill_key.upper()}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="1A5276")
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = f"Dự án: {project_name}    |    Bản vẽ: {drawing_no}"
    ws["A2"].font = Font(name="Arial", italic=True, size=10, color="5D6D7E")
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.append([])  # blank row

    # ── Column headers (row 4) ────────────────────────────────────────────────
    header_row = 4
    for ci, (col_name, width) in enumerate(zip(COLS, COL_WIDTHS), start=1):
        cell = ws.cell(row=header_row, column=ci, value=col_name)
        cell.fill  = _fill(CLR_HEADER)
        cell.font  = _font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.row_dimensions[header_row].height = 22

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_dict in bom_rows:
        rtype = row_dict.get("row_type", "data")
        r = ws.max_row + 1

        if rtype == "section":
            ws.cell(r, 1, row_dict["stt"])
            ws.cell(r, 2, row_dict["mo_ta"])
            sec_clr = "FFE0B2" if "Chưa Phân Loại" in str(row_dict.get("mo_ta","")) else CLR_SECTION
            for ci in range(1, 11):
                c = ws.cell(r, ci)
                c.fill = _fill(sec_clr)
                c.font = _font(bold=True)
                c.alignment = Alignment(horizontal="left" if ci == 2 else "center",
                                        vertical="center")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)

        elif rtype == "subsection":
            ws.cell(r, 1, row_dict["stt"])
            ws.cell(r, 2, row_dict["mo_ta"])
            for ci in range(1, 11):
                c = ws.cell(r, ci)
                c.fill = _fill(CLR_SUBSEC)
                c.font = _font(italic=True)
                c.alignment = Alignment(horizontal="left" if ci == 2 else "center",
                                        vertical="center")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)

        else:  # data
            values = [
                row_dict.get("stt",        ""),
                row_dict.get("mo_ta",      ""),
                row_dict.get("chung_loai", ""),
                row_dict.get("vat_lieu",   ""),
                row_dict.get("kthuoc1",    ""),
                row_dict.get("kthuoc2",    ""),
                row_dict.get("tieu_chuan", ""),
                row_dict.get("xuat_xu",    ""),
                row_dict.get("don_vi",     ""),
                row_dict.get("so_luong",   0),
            ]
            clr = CLR_UNKNOWN if row_dict.get("size_unknown") else CLR_DATA
            for ci, val in enumerate(values, start=1):
                c = ws.cell(r, ci, val)
                c.fill = _fill(clr)
                c.font = _font()
                c.alignment = Alignment(
                    horizontal="left" if ci == 2 else "center",
                    vertical="center",
                    wrap_text=(ci == 2),
                )

        ws.row_dimensions[r].height = 18

    # ── Thin border for all data ───────────────────────────────────────────────
    from openpyxl.styles import Border, Side
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = border

    # ── Freeze header ─────────────────────────────────────────────────────────
    ws.freeze_panes = ws.cell(header_row + 1, 1)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# VERIFICATION
# ─────────────────────────────────────────────────────────────────────────────
def verify_bom(equipment_list: list, bom_rows: list) -> list:
    """
    Tự kiểm tra BOM theo skill — trả về list lỗi.
    """
    errors = []

    # Build DXF summary: desc → size → count
    dxf_summary = defaultdict(lambda: defaultdict(int))
    for eq in equipment_list:
        dxf_summary[eq["info"]["desc"]][eq["size"]] += 1

    # Build BOM summary
    bom_summary = defaultdict(lambda: defaultdict(int))
    for row in bom_rows:
        if row.get("row_type") != "data":
            continue
        desc = row.get("mo_ta", "")
        size = row.get("kthuoc1") or "?"
        qty  = row.get("so_luong", 0)
        bom_summary[desc][size] += qty

    # Check 1: tổng số lượng
    for desc, sizes in dxf_summary.items():
        total_dxf = sum(sizes.values())
        total_bom = sum(bom_summary.get(desc, {}).values())
        if total_bom != total_dxf:
            errors.append(f"❌ TỔNG: {desc} → DXF={total_dxf}, BOM={total_bom}")

    # Check 2: size = "?"
    for desc, sizes in dxf_summary.items():
        if "?" in sizes:
            errors.append(f"⚠️ UNKNOWN SIZE: {desc} × {sizes['?']} EA — cần kiểm tra bản vẽ")

    return errors


# ─────────────────────────────────────────────────────────────────────────────
# MAIN UI
# ─────────────────────────────────────────────────────────────────────────────
col_up, col_info = st.columns([2, 1])

with col_up:
    st.subheader("📂 Upload File DXF")
    uploaded = st.file_uploader(
        "Kéo thả hoặc chọn file .dxf",
        type=["dxf"],
        help="File DXF P&ID theo chuẩn ANH MINH 2021"
    )

with col_info:
    st.subheader("ℹ️ Quy trình")
    st.markdown("""
1. Upload file **DXF**
2. Chọn **Skill** (SPX / Tetra Pak)
3. App đọc **block name** trực tiếp
4. Tìm **size** theo annotation gần nhất
5. Xuất **BOM Excel** chuẩn ANH MINH
    """)


if uploaded:
    file_bytes = uploaded.read()
    file_name  = Path(uploaded.name).stem

    with st.spinner("⏳ Đang phân tích file DXF..."):
        equipment_list, unknown_blocks, size_annotations, parse_errors = parse_dxf(
            file_bytes, skill_choice, max_dist
        )

    # ── Parse errors ──────────────────────────────────────────────────────────
    if parse_errors:
        for e in parse_errors:
            st.error(e)
        st.stop()

    # ── Stats ─────────────────────────────────────────────────────────────────
    st.markdown("---")
    total_equip   = len(equipment_list)
    total_unknown = sum(unknown_blocks.values())
    total_sizes   = len(size_annotations)
    size_missing  = sum(1 for eq in equipment_list if eq["size"] == "?")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-num">{total_equip}</div>'
                    f'<div class="stat-label">Thiết bị nhận diện</div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-num" style="color:#e74c3c">{total_unknown}</div>'
                    f'<div class="stat-label">Block chưa nhận diện</div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="stat-num">{total_sizes}</div>'
                    f'<div class="stat-label">Annotation size tìm thấy</div>', unsafe_allow_html=True)
    with c4:
        clr = "#e74c3c" if size_missing else "#27ae60"
        st.markdown(f'<div class="stat-num" style="color:{clr}">{size_missing}</div>'
                    f'<div class="stat-label">Thiết bị chưa có size</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_bom, tab_equip, tab_unknown, tab_size = st.tabs([
        "📋 BOM Preview", "🔩 Thiết bị nhận diện", "❓ Block chưa nhận diện", "📏 Size annotations"
    ])

    # ── Build BOM ─────────────────────────────────────────────────────────────
    groups   = build_bom(equipment_list)
    bom_rows = flatten_bom(groups)

    with tab_bom:
        if not bom_rows:
            st.warning("Không có thiết bị nào được nhận diện. Kiểm tra lại file DXF và skill đã chọn.")
        else:
            # Preview table
            preview_data = []
            for row in bom_rows:
                rtype = row.get("row_type", "data")
                if rtype == "section":
                    preview_data.append({
                        "STT": row["stt"], "MÔ TẢ": f"▸ {row['mo_ta']}",
                        "CHỦNG LOẠI":"","VẬT LIỆU":"","K.THƯỚC 1":"",
                        "TIÊU CHUẨN":"","ĐƠN VỊ":"","SỐ LƯỢNG":"",
                    })
                elif rtype == "subsection":
                    preview_data.append({
                        "STT": row["stt"], "MÔ TẢ": f"  → {row['mo_ta']}",
                        "CHỦNG LOẠI":"","VẬT LIỆU":"","K.THƯỚC 1":"",
                        "TIÊU CHUẨN":"","ĐƠN VỊ":"","SỐ LƯỢNG":"",
                    })
                else:
                    sz = row.get("kthuoc1","") or "⚠️ ?"
                    preview_data.append({
                        "STT":          row.get("stt",""),
                        "MÔ TẢ":        row.get("mo_ta",""),
                        "CHỦNG LOẠI":   row.get("chung_loai",""),
                        "VẬT LIỆU":     row.get("vat_lieu",""),
                        "K.THƯỚC 1":    sz,
                        "TIÊU CHUẨN":   row.get("tieu_chuan",""),
                        "ĐƠN VỊ":       row.get("don_vi",""),
                        "SỐ LƯỢNG":     row.get("so_luong",""),
                    })

            df_preview = pd.DataFrame(preview_data)
            st.dataframe(df_preview, use_container_width=True, height=450,
                         hide_index=True)

            # Verification
            st.subheader("✅ Kiểm tra BOM")
            verify_errors = verify_bom(equipment_list, bom_rows)
            if verify_errors:
                for ve in verify_errors:
                    if ve.startswith("❌"):
                        st.error(ve)
                    else:
                        st.warning(ve)
            else:
                total_items = sum(row.get("so_luong", 0) for row in bom_rows
                                  if row.get("row_type") == "data")
                st.success(f"✅ BOM kiểm tra PASS — {len(groups)} nhóm, tổng {total_items} EA")

            # Export button
            st.subheader("💾 Xuất BOM Excel")
            excel_bytes = export_excel(bom_rows, project_name, drawing_no, skill_choice)
            output_name = f"BOM_{file_name}_{skill_choice.split()[0]}.xlsx"

            st.download_button(
                label="⬇️  Tải xuống BOM Excel (.xlsx)",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

    with tab_equip:
        if equipment_list:
            rows_display = []
            for eq in equipment_list:
                info = eq["info"]
                rows_display.append({
                    "Block Name":   eq["block_name"],
                    "Mô tả":        info["desc"],
                    "Mã":           info["code"],
                    "Vật liệu":     info["material"],
                    "Tiêu chuẩn":   info["standard"],
                    "Section":      info["section"],
                    "Size":         eq["size"],
                    "X":            round(eq["x"], 1),
                    "Y":            round(eq["y"], 1),
                })
            df_eq = pd.DataFrame(rows_display)
            st.dataframe(df_eq, use_container_width=True, height=500, hide_index=True)
            st.caption(f"Tổng: {len(rows_display)} thiết bị nhận diện từ {len(set(e['block_name'] for e in equipment_list))} loại block")
        else:
            st.info("Không có thiết bị nào được nhận diện.")

    with tab_unknown:
        if unknown_blocks:
            st.markdown('<div class="warn-box">⚠️ Các block dưới đây chưa có trong database — '
                        '<b>đã tự động đưa vào BOM</b> dùng tên block làm tên thiết bị. '
                        'Cập nhật <code>pid_database.py</code> để phân loại chính xác hơn.</div>',
                        unsafe_allow_html=True)
            df_unk = pd.DataFrame([
                {"Block Name": k, "Số lần xuất hiện": v, "Xử lý": "✅ Đã đưa vào BOM (tên block = tên thiết bị)"}
                for k, v in sorted(unknown_blocks.items(), key=lambda x: -x[1])
            ])
            st.dataframe(df_unk, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Tất cả block đều được nhận diện!")

    with tab_size:
        st.subheader(f"📏 {len(size_annotations)} size annotation tìm thấy trong DXF")
        if size_annotations:
            df_sz = pd.DataFrame([
                {"DN/Size": a["dn"], "Text gốc": a["text"],
                 "X": round(a["x"],1), "Y": round(a["y"],1)}
                for a in size_annotations
            ])
            st.dataframe(df_sz, use_container_width=True, height=400, hide_index=True)

            # Size distribution
            dn_counts = defaultdict(int)
            for eq in equipment_list:
                dn_counts[eq["size"]] += 1
            if dn_counts:
                st.subheader("📊 Phân bố size thiết bị")
                df_dist = pd.DataFrame([
                    {"Size": k, "Số thiết bị": v}
                    for k, v in sorted(dn_counts.items())
                ])
                st.bar_chart(df_dist.set_index("Size"))
        else:
            st.warning("Không tìm thấy annotation size nào trong DXF. "
                       "Kiểm tra text trong file có chứa 'DN50', '2\"',... không.")

else:
    # ── Welcome screen ────────────────────────────────────────────────────────
    st.markdown("""
    <div class="section-box">
    <b>🚀 Hướng dẫn sử dụng:</b><br>
    1. Chọn <b>Skill</b> phù hợp ở thanh bên (SPX hoặc Tetra Pak)<br>
    2. Upload file <b>.dxf</b> P&ID<br>
    3. App tự động đọc <b>tên block</b> → tra cứu database → tìm <b>size</b> gần nhất<br>
    4. Xem preview BOM và tải xuống file <b>Excel</b>
    </div>
    """, unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("""
        **📚 SPX Type (ANH MINH 2021)**
        - SSV 200 (SV41-L) / 210 (SV43-LL) / 220 (SV44-TL) / 300 (SV42-T)
        - Mixproof Valve (MPV)
        - Butterfly Valve (BFV-M/A/TK)
        - Sampling Valve (SVA/SVM)
        - Centrifugal Pump W+ / WS+
        - Lobe Pump DW+, Vane Pump, Blower...
        - Utility: Globe/Ball/Diaphragm/Angle Seat Valve
        """)
    with col_b:
        st.markdown("""
        **📚 Tetra Pak Type (ANH MINH 2021)**
        - SSV 200 / 210 / 220 / 300 (không kèm model SV41...)
        - **Leakage Valve NC/NO** (đặc thù Tetra Pak)
        - Aseptic Sampling Valve / Sanitary Sampling Valve
        - **Screw Pump / Utility Pump / Diaphragm Pump** (đặc thù)
        - **Bag Filter / Angle Filter** (đặc thù)
        - Utility: Globe/Ball/BFV/Diaphragm Valve
        """)

    st.info("💡 App không gọi AI, không tốn cost. Toàn bộ tra cứu từ database nội bộ.")
