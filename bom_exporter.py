"""
bom_exporter.py
Xuất BOM đúng chuẩn ANHMINH (FORM_BOM_STANDARD).
Skill ANH MINH 2021.

Màu sắc:
  Major section (A/B/C): nền CYAN,   chữ đỏ, bold
  Sub-section  (I/II/III): nền VÀNG, chữ đỏ, bold
  Data rows (thiết bị thực tế): KHÔNG nền, chữ đen, KHÔNG bold
"""
import io
import re
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Hằng số ───────────────────────────────────────────────────────────────
BG_CYAN   = PatternFill("solid", start_color="00FFFF", end_color="00FFFF")
BG_YELLOW = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
BG_NONE   = PatternFill(fill_type=None)
FG_RED    = "FF0000"
FG_BLACK  = "000000"

COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
COL_WIDTHS = [5.57, 47.71, 15.14, 16.14, 11.29, 12.00, 10.71, 15.43, 6.71, 9.71]
COL_HEADERS = ["STT", "MÔ TẢ", "CHỦNG LOẠI", "VẬT LIỆU",
               "K.THƯỚC 1", "K.THƯỚC 2", "TIÊU CHUẨN", "XUẤT XỨ", "ĐƠN VỊ", "SỐ LƯỢNG"]

ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
         "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"]

SECTION_ORDER = [
    ("Product",          "ĐƯỜNG PRODUCT (SANITARY)"),
    ("CIP",              "ĐƯỜNG CIP (SANITARY)"),
    ("RO/Process Water", "ĐƯỜNG RO / PROCESS WATER (SANITARY)"),
    ("Ice Water",        "ĐƯỜNG ICE WATER (CHILLER)"),
    ("Cooling Water",    "ĐƯỜNG COOLING WATER"),
    ("Steam/Condensate", "ĐƯỜNG STEAM & CONDENSATE"),
    ("Compressed Air",   "ĐƯỜNG COMPRESSED AIR"),
    ("City/Soft Water",  "ĐƯỜNG CITY WATER / SOFT WATER"),
    ("Unknown",          "CHƯA XÁC ĐỊNH LINE"),
]


def _thin():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _apply(ws, row: int, values: list, bg, font_color: str, bold: bool,
           row_height: float = 16):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = bg
        c.font = Font(name="Arial", bold=bold, size=11, color=font_color)
        c.border = _thin()
        c.alignment = Alignment(
            horizontal="left" if col == 2 else "center",
            vertical="center", wrap_text=True,
        )
    ws.row_dimensions[row].height = row_height


def write_major(ws, row: int, stt: str, mo_ta: str) -> int:
    _apply(ws, row, [stt, mo_ta, "", "", "", "", "", "", "SET", 1],
           BG_CYAN, FG_RED, True)
    return row + 1


def write_sub(ws, row: int, stt: str, mo_ta: str) -> int:
    _apply(ws, row, [stt, mo_ta, "", "", "", "", "", "", "SET", 1],
           BG_YELLOW, FG_RED, True)
    return row + 1


def write_data(ws, row: int, stt, mo_ta, chung_loai, vat_lieu,
               kt1, kt2, tieu_chuan, xuat_xu, don_vi, so_luong) -> int:
    _apply(ws, row,
           [stt, mo_ta, chung_loai, vat_lieu, kt1, kt2,
            tieu_chuan, xuat_xu, don_vi, so_luong],
           BG_NONE, FG_BLACK, False)
    return row + 1


def _build_header(ws):
    """Dòng 1–15: Header công ty + tiêu đề BOM."""
    # Merge A1:J1 — tên công ty
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = "ANHMINH TECHNOLOGY TRADING COMPANY LIMITED"
    c.font = Font(name="Arial", bold=True, size=16, color="000000")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:J2")
    ws["A2"].value = "Address: ..."
    ws["A2"].font = Font(name="Arial", size=10)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:J3")
    ws["A3"].value = f"BÁO GIÁ / QUOTATION — {date.today().strftime('%d/%m/%Y')}"
    ws["A3"].font = Font(name="Arial", bold=True, size=14)
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 22

    # Rows 4-9: bỏ trống
    for r in range(4, 10):
        ws.row_dimensions[r].height = 8

    # Rows 10-11: Column header (merge 2 dòng)
    for r in range(10, 12):
        for col, (hdr, w) in enumerate(zip(COL_HEADERS, COL_WIDTHS), 1):
            c = ws.cell(row=r, column=col)
            if r == 10:
                c.value = hdr
            c.fill = BG_CYAN
            c.font = Font(name="Arial", bold=True, size=11, color=FG_RED)
            c.border = _thin()
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        ws.row_dimensions[r].height = 18
    # Merge header cells vertically
    for col in range(1, 11):
        ws.merge_cells(
            start_row=10, start_column=col,
            end_row=11, end_column=col,
        )

    # Set column widths
    for col, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    return 12  # Dòng data bắt đầu từ 12


def _build_footer(ws, row: int):
    for label in ["TOTAL (VND)", "VAT 10% (VND)", "GRAND TOTAL (VND)"]:
        ws.merge_cells(f"A{row}:J{row}")
        c = ws[f"A{row}"]
        c.value = label
        c.font = Font(name="Arial", bold=True, size=11)
        c.border = _thin()
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 16
        row += 1

    ws.merge_cells(f"A{row}:J{row}")
    ws[f"A{row}"].value = "* Ghi chú / Notes:"
    ws[f"A{row}"].font = Font(name="Arial", bold=True, size=11, color=FG_RED)
    row += 1
    notes = [
        "- Giá này không bao gồm VAT.",
        "- Thời gian giao hàng / Lead time: 7–10 tuần.",
        "- Điều kiện thanh toán / Payment term: 30% tạm ứng – 40% tập kết – 30% sau khi hoàn thành.",
        "- Thời hạn chào giá / Validity: 30 ngày kể từ ngày gửi báo giá.",
        "Trân trọng kính chào / Best regards",
    ]
    for note in notes:
        ws.merge_cells(f"A{row}:J{row}")
        ws[f"A{row}"].value = note
        ws[f"A{row}"].font = Font(name="Arial", size=10)
        row += 1


def _size_sort_key(size_str: str) -> float:
    """
    Chuyển chuỗi size → số để sort giảm dần (lớn → nhỏ).
    DN50 → 50.0 | 2" → 50.8 (2*25.4) | 1.5" → 38.1 | ? → -1
    """
    s = str(size_str).strip()
    if s == "?":
        return -1.0
    # DN format
    m = re.search(r"DN\s*(\d+(?:\.\d+)?)", s, re.IGNORECASE)
    if m:
        return float(m.group(1))
    # Inch format: 1", 1.5", 1/2"
    m2 = re.search(r"(\d+)/(\d+)", s)
    if m2:
        return float(m2.group(1)) / float(m2.group(2)) * 25.4
    m3 = re.search(r"(\d+(?:\.\d+)?)", s)
    if m3:
        return float(m3.group(1)) * 25.4
    return -1.0


def export_bom(items: list) -> bytes:
    """
    items: list of dict {
        block_name, display_name, line_type, line_label, group,
        size, chung_loai, vat_lieu, tieu_chuan, don_vi
    }
    Trả về bytes của file .xlsx.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"
    ws.sheet_view.showGridLines = False

    r = _build_header(ws)

    # Nhóm items theo line_type
    by_line = defaultdict(list)
    for item in items:
        by_line[item["line_type"]].append(item)

    # Đếm số lượng theo (display_name, size, line_type)
    summary = defaultdict(int)
    for item in items:
        key = (item["display_name"], item["size"], item["line_type"],
               item["chung_loai"], item["vat_lieu"], item["tieu_chuan"],
               item["don_vi"], item["line_label"])
        summary[key] += 1

    # Tái tổ chức: grouped_summary[line_type] = [(key, count), ...]
    # Sắp xếp: theo tên thiết bị (asc) → size giảm dần (lớn → nhỏ)
    grouped = defaultdict(list)
    for key, count in summary.items():
        lt = key[2]
        grouped[lt].append((key, count))

    for lt in grouped:
        grouped[lt].sort(key=lambda kc: (
            kc[0][0],                        # display_name asc
            -_size_sort_key(kc[0][1]),       # size desc (âm để sort tăng = giảm thực)
        ))

    major_idx = 0
    roman_idx = 0
    major_letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    r = write_major(ws, r, major_letters[major_idx], "PHẦN CƠ KHÍ / MECHANICAL")
    major_idx += 1

    for line_type, section_label in SECTION_ORDER:
        rows_in_section = grouped.get(line_type, [])
        if not rows_in_section:
            continue

        sub_label = section_label
        r = write_sub(ws, r, ROMAN[roman_idx], sub_label)
        roman_idx += 1

        for idx, (key, count) in enumerate(rows_in_section, 1):
            display_name, size, lt, chung_loai, vat_lieu, tieu_chuan, don_vi, lbl = key
            r = write_data(
                ws, r,
                stt=str(idx),
                mo_ta=display_name,
                chung_loai=chung_loai,
                vat_lieu=vat_lieu,
                kt1=size,
                kt2="-",
                tieu_chuan=tieu_chuan,
                xuat_xu="",
                don_vi=don_vi,
                so_luong=count,
            )

    _build_footer(ws, r + 1)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
