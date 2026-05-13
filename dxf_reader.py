"""
dxf_reader.py — Đọc file DXF P&ID (SPX & Tetra Pak)
Kỹ năng:
  0. Word-order-insensitive block name matching
  1. Proximity-based size detection (nearest DN annotation)
  2. Linename TP-left → line type → spec group
"""

import math
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

import ezdxf

from pid_spec import (
    classify_line, should_ignore, is_linename_block,
    get_spec, is_thread_end, parse_dn,
)

MAX_DIST = 500      # ngưỡng proximity (đơn vị DXF)
DN_PATTERN = re.compile(r'DN\s*\d+|\d+\s*"|\b\d+\s*(?:mm|MM)\b')


# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class LineInfo:
    line_id: str          # VD: IW-001
    line_type: str        # SANITARY / ICE_WATER / COOLING / STEAM / AIR / UNKNOWN
    x: float = 0.0
    y: float = 0.0


@dataclass
class Equipment:
    block_name: str
    x: float
    y: float
    size: str = "?"
    line_id: str = ""
    line_type: str = "UNKNOWN"
    spec: dict = field(default_factory=dict)
    conn_type: str = ""   # Thread end / Flange PN16 / Flange JIS10K / SMS / ...


@dataclass
class DXFResult:
    lines: list[LineInfo]
    equipment: list[Equipment]
    warnings: list[str]


# ── Main reader ────────────────────────────────────────────────────────────────

def read_dxf(filepath: str) -> DXFResult:
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()

    lines: list[LineInfo] = []
    raw_equip: list[dict] = []
    size_texts: list[dict] = []   # {dn, x, y}
    warnings: list[str] = []

    # ── Pass 1: Scan tất cả entities ──────────────────────────────────────────
    for e in msp:
        t = e.dxftype()

        if t == "INSERT":
            name = e.dxf.name or ""
            pos = e.dxf.insert
            px, py = float(pos.x), float(pos.y)

            if is_linename_block(name):
                # Đọc attribute để lấy line ID
                line_id = _extract_linename_attr(e) or _extract_linename_text(name)
                if line_id:
                    lt = classify_line(line_id)
                    lines.append(LineInfo(line_id=line_id, line_type=lt, x=px, y=py))

            elif not should_ignore(name):
                raw_equip.append({"name": name, "x": px, "y": py})

        elif t in ("TEXT", "MTEXT"):
            raw_text = _get_text(e)
            if not raw_text:
                continue
            try:
                pos = e.dxf.insert
                px2, py2 = float(pos.x), float(pos.y)
            except Exception:
                continue
            # Thu thập annotation kích thước
            m = re.search(r'DN\s*(\d+)', raw_text.upper())
            if m:
                size_texts.append({"dn": f"DN{m.group(1)}", "x": px2, "y": py2})
            else:
                m2 = re.search(r'(\d+(?:\.\d+)?)\s*"', raw_text)
                if m2:
                    size_texts.append({"dn": m2.group(0).strip(), "x": px2, "y": py2})

        # ATTRIB entities (attribute text trực tiếp)
        elif t == "ATTRIB":
            raw_text = e.dxf.text or ""
            try:
                pos = e.dxf.insert
                px2, py2 = float(pos.x), float(pos.y)
            except Exception:
                continue
            m = re.search(r'DN\s*(\d+)', raw_text.upper())
            if m:
                size_texts.append({"dn": f"DN{m.group(1)}", "x": px2, "y": py2})

    # ── Pass 2: Assign size → proximity nearest DN ────────────────────────────
    # ── Pass 3: Assign line → nearest linename block ──────────────────────────
    equipment: list[Equipment] = []

    for eq in raw_equip:
        ex, ey = eq["x"], eq["y"]

        # Size
        dn_str, dist = _nearest_dn(ex, ey, size_texts)
        if dist is None:
            warnings.append(f"⚠️ {eq['name']} tại ({ex:.0f},{ey:.0f}): không tìm được size → cần kiểm tra")

        # Line
        nearest_line = _nearest_line(ex, ey, lines)
        line_id   = nearest_line.line_id   if nearest_line else ""
        line_type = nearest_line.line_type if nearest_line else "UNKNOWN"

        # Spec
        spec = get_spec(line_type, eq["name"], dn_str)
        conn_type = _resolve_conn(line_type, dn_str, spec)

        equipment.append(Equipment(
            block_name=eq["name"],
            x=ex, y=ey,
            size=dn_str,
            line_id=line_id,
            line_type=line_type,
            spec=spec,
            conn_type=conn_type,
        ))

    return DXFResult(lines=lines, equipment=equipment, warnings=warnings)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _extract_linename_attr(insert_entity) -> Optional[str]:
    """Đọc ATTRIB đính kèm INSERT block để lấy line ID."""
    try:
        for attrib in insert_entity.attribs:
            text = (attrib.dxf.text or "").strip()
            if text and len(text) <= 30:
                # Kiểm tra có dạng PREFIX-NUMBER không
                if re.match(r"[A-Z]{1,6}[-_]\d+", text.upper()):
                    return text.upper()
                if len(text) >= 3:
                    return text.upper()
    except Exception:
        pass
    return None


def _extract_linename_text(block_name: str) -> Optional[str]:
    """Fallback: thử parse tên block nếu chứa line ID."""
    m = re.search(r"([A-Z]{1,6}[-_]\d+)", block_name.upper())
    return m.group(1) if m else None


def _get_text(entity) -> str:
    try:
        if entity.dxftype() == "MTEXT":
            return entity.text or ""
        return entity.dxf.text or ""
    except Exception:
        return ""


def _nearest_dn(ex: float, ey: float, candidates: list, max_dist: float = MAX_DIST):
    best, best_d = None, float("inf")
    for t in candidates:
        d = math.hypot(t["x"] - ex, t["y"] - ey)
        if d < best_d:
            best_d, best = d, t
    if best and best_d <= max_dist:
        return best["dn"], round(best_d, 1)
    return "?", None


def _nearest_line(ex: float, ey: float, lines: list[LineInfo]) -> Optional[LineInfo]:
    """Gán thiết bị vào line gần nhất (không giới hạn khoảng cách)."""
    if not lines:
        return None
    return min(lines, key=lambda l: math.hypot(l.x - ex, l.y - ey))


def _resolve_conn(line_type: str, size_str: str, spec: dict) -> str:
    """Xác định kiểu kết nối hiển thị."""
    tc = spec.get("tieu_chuan", "")
    if "Thread" in tc:
        return "Thread end"
    if "JIS10K" in tc:
        return "Flange JIS10K"
    if "PN16" in tc:
        return "Flange PN16"
    if "SMS" in tc or "DIN" in tc:
        return "Clamp/SMS"
    if line_type == "SANITARY":
        return "Clamp/SMS"
    if is_thread_end(size_str):
        return "Thread end"
    return tc or "-"


# ── Aggregate BOM ──────────────────────────────────────────────────────────────

def aggregate_bom(equipment: list[Equipment]) -> list[dict]:
    """
    Gom nhóm thiết bị giống nhau (cùng block_name + size + line_type) → đếm SL.
    Trả về list[dict] đã sort theo line_type, block_name, size.
    """
    counts: dict[tuple, int] = defaultdict(int)
    meta: dict[tuple, dict] = {}

    for eq in equipment:
        key = (eq.line_type, eq.line_id[:3] if eq.line_id else "??",
               eq.block_name, eq.size)
        counts[key] += 1
        if key not in meta:
            meta[key] = {
                "line_type": eq.line_type,
                "line_id": eq.line_id,
                "block_name": eq.block_name,
                "size": eq.size,
                "conn_type": eq.conn_type,
                **eq.spec,
            }

    rows = []
    for key, sl in counts.items():
        row = dict(meta[key])
        row["sl"] = sl
        rows.append(row)

    # Sort
    order = {"SANITARY": 0, "ICE_WATER": 1, "COOLING": 2, "STEAM": 3, "AIR": 4, "UNKNOWN": 9}
    rows.sort(key=lambda r: (order.get(r["line_type"], 9), r["block_name"], r["size"]))
    return rows
