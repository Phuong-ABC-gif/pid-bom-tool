"""
dxf_reader.py — Đọc file DXF P&ID (SPX & Tetra Pak) v3
Thuật toán nhận diện loại đường ống:
  1. Thu thập tất cả LINE / LWPOLYLINE / POLYLINE thành segments
  2. Thu thập tất cả TEXT / MTEXT
  3. Với mỗi thiết bị (INSERT):
       a. Tìm segment chạm vào thiết bị (endpoint trong TOUCH_TOL)
       b. BFS theo segment kết nối (endpoint chung trong SNAP_TOL)
       c. Tìm TEXT gần bất kỳ segment nào đã duyệt (TEXT_SEG_DIST)
       d. Score text → chọn tên line tốt nhất → classify_line_from_text()
  4. Size: proximity nearest DN annotation
  5. KHÔNG dùng block linename TP-left nữa
"""

import math
import re
from collections import defaultdict, deque
from dataclasses import dataclass, field
from typing import Optional

import ezdxf

from pid_spec import (
    classify_line, classify_line_from_text,
    should_ignore, get_spec, is_thread_end,
)

# ── Ngưỡng ────────────────────────────────────────────────────────────────────
TOUCH_TOL   = 80    # segment endpoint cách thiết bị bao nhiêu units thì coi là kết nối
SNAP_TOL    = 30    # 2 endpoint coi là chung (connected)
TEXT_SEG    = 200   # text cách segment bao nhiêu units thì lấy vào candidate
BFS_HOPS    = 20    # max số hop BFS
SIZE_DIST   = 500   # ngưỡng proximity tìm size


# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class Segment:
    sx: float; sy: float
    ex: float; ey: float
    layer: str = ""
    color: int = 256   # 256 = BYLAYER


@dataclass
class Equipment:
    block_name: str
    x: float; y: float
    size: str       = "?"
    line_id: str    = ""
    line_type: str  = "UNKNOWN"
    spec: dict      = field(default_factory=dict)
    conn_type: str  = ""


@dataclass
class DXFResult:
    equipment: list
    warnings:  list


# ── Main ───────────────────────────────────────────────────────────────────────

def read_dxf(filepath: str) -> DXFResult:
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()

    segments:   list[Segment] = []
    all_texts:  list[dict]    = []   # {text, x, y}
    size_texts: list[dict]    = []   # {dn, x, y}
    raw_equip:  list[dict]    = []
    warnings:   list[str]     = []

    # ── Pass 1: Thu thập tất cả entities ──────────────────────────────────────
    for e in msp:
        t = e.dxftype()

        # Segments
        if t == "LINE":
            try:
                s = e.dxf.start; en = e.dxf.end
                segments.append(Segment(
                    sx=s.x, sy=s.y, ex=en.x, ey=en.y,
                    layer=e.dxf.layer or "",
                    color=e.dxf.get("color", 256),
                ))
            except Exception:
                pass

        elif t == "LWPOLYLINE":
            try:
                pts = list(e.get_points())   # [(x,y,…), …]
                lyr = e.dxf.layer or ""
                col = e.dxf.get("color", 256)
                for i in range(len(pts) - 1):
                    segments.append(Segment(
                        sx=pts[i][0], sy=pts[i][1],
                        ex=pts[i+1][0], ey=pts[i+1][1],
                        layer=lyr, color=col,
                    ))
                if e.is_closed and len(pts) > 1:
                    segments.append(Segment(
                        sx=pts[-1][0], sy=pts[-1][1],
                        ex=pts[0][0], ey=pts[0][1],
                        layer=lyr, color=col,
                    ))
            except Exception:
                pass

        elif t == "POLYLINE":
            try:
                verts = list(e.vertices)
                lyr = e.dxf.layer or ""
                col = e.dxf.get("color", 256)
                for i in range(len(verts) - 1):
                    p1 = verts[i].dxf.location
                    p2 = verts[i+1].dxf.location
                    segments.append(Segment(
                        sx=p1.x, sy=p1.y, ex=p2.x, ey=p2.y,
                        layer=lyr, color=col,
                    ))
            except Exception:
                pass

        # Text
        elif t in ("TEXT", "MTEXT", "ATTRIB"):
            raw = _get_text(e)
            if not raw:
                continue
            try:
                pos = e.dxf.insert
                px, py = float(pos.x), float(pos.y)
            except Exception:
                continue
            clean = _strip_mtext(raw).strip()
            if clean:
                all_texts.append({"text": clean, "x": px, "y": py})

        # Equipment
        elif t == "INSERT":
            name = e.dxf.name or ""
            if should_ignore(name):
                continue
            pos = e.dxf.insert
            raw_equip.append({"name": name,
                               "x": float(pos.x),
                               "y": float(pos.y)})

    # ── Pass 2: Tách size texts khỏi all_texts ─────────────────────────────────
    for tx in all_texts:
        raw = tx["text"].upper()
        m = re.search(r'DN\s*(\d+)', raw)
        if m:
            size_texts.append({"dn": f"DN{m.group(1)}", "x": tx["x"], "y": tx["y"]})
            continue
        m2 = re.search(r'(\d+(?:\.\d+)?)\s*"', tx["text"])
        if m2:
            size_texts.append({"dn": m2.group(0).strip(), "x": tx["x"], "y": tx["y"]})

    # ── Pass 3: Build adjacency index (endpoint grid) ─────────────────────────
    # adj[seg_idx] = list of seg_idx kết nối (chung endpoint)
    adj = _build_adjacency(segments, SNAP_TOL)

    # ── Pass 4: Với mỗi thiết bị → BFS → tìm text → classify ─────────────────
    equipment: list[Equipment] = []

    for eq in raw_equip:
        ex, ey = eq["x"], eq["y"]

        # 4a. Tìm segments chạm thiết bị
        seed_segs = _touching_segments(ex, ey, segments, TOUCH_TOL)

        # 4b. BFS mở rộng
        visited = _bfs(seed_segs, adj, BFS_HOPS)

        # 4c. Tìm text gần các segment đã duyệt
        candidate_texts = _texts_near_segments(
            [segments[i] for i in visited], all_texts, TEXT_SEG
        )

        # 4d. Classify line type từ text candidates
        line_id, line_type = _best_line_from_texts(candidate_texts)

        # 4e. Fallback: dùng layer / color của segment chạm thiết bị
        if line_type == "UNKNOWN" and seed_segs:
            seg0 = segments[seed_segs[0]]
            line_type = _classify_by_layer_color(seg0.layer, seg0.color)
            line_id   = line_type  # không có tên cụ thể

        # 4f. Size
        dn_str, dist = _nearest_dn(ex, ey, size_texts, SIZE_DIST)
        if dist is None:
            warnings.append(
                f"⚠ {eq['name']} ({ex:.0f},{ey:.0f}): không xác định được size"
            )

        # 4g. Spec
        spec      = get_spec(line_type, eq["name"], dn_str)
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

    return DXFResult(equipment=equipment, warnings=warnings)


# ── Segment helpers ────────────────────────────────────────────────────────────

def _build_adjacency(segs: list[Segment], tol: float) -> dict[int, list[int]]:
    """
    Hai segment kết nối nếu endpoint của seg A gần endpoint của seg B (< tol).
    Dùng grid bucket để O(n) thay vì O(n²).
    """
    bucket_size = tol * 2
    grid: dict[tuple, list[tuple]] = defaultdict(list)

    def cell(x, y):
        return (int(x // bucket_size), int(y // bucket_size))

    # Đăng ký tất cả endpoints vào grid
    for i, s in enumerate(segs):
        for (px, py) in [(s.sx, s.sy), (s.ex, s.ey)]:
            grid[cell(px, py)].append((px, py, i))

    adj: dict[int, list[int]] = defaultdict(list)

    def neighbors(x, y):
        cx, cy = cell(x, y)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                yield from grid.get((cx+dx, cy+dy), [])

    for i, s in enumerate(segs):
        for (px, py) in [(s.sx, s.sy), (s.ex, s.ey)]:
            for (qx, qy, j) in neighbors(px, py):
                if j != i and math.hypot(px - qx, py - qy) <= tol:
                    adj[i].append(j)

    return adj


def _touching_segments(ex: float, ey: float,
                        segs: list[Segment], tol: float) -> list[int]:
    """Trả về index các segment có endpoint gần (ex, ey) trong ngưỡng tol."""
    result = []
    for i, s in enumerate(segs):
        if (math.hypot(s.sx - ex, s.sy - ey) <= tol or
                math.hypot(s.ex - ex, s.ey - ey) <= tol):
            result.append(i)
    return result


def _bfs(seeds: list[int], adj: dict, max_hops: int) -> set[int]:
    """BFS từ seed segments theo adjacency graph."""
    visited = set(seeds)
    queue   = deque((s, 0) for s in seeds)
    while queue:
        idx, hop = queue.popleft()
        if hop >= max_hops:
            continue
        for nb in adj.get(idx, []):
            if nb not in visited:
                visited.add(nb)
                queue.append((nb, hop + 1))
    return visited


def _pt_to_seg_dist(px: float, py: float,
                    sx: float, sy: float,
                    ex: float, ey: float) -> float:
    """Khoảng cách từ điểm (px,py) đến đoạn thẳng (sx,sy)-(ex,ey)."""
    dx, dy = ex - sx, ey - sy
    if dx == 0 and dy == 0:
        return math.hypot(px - sx, py - sy)
    t = max(0.0, min(1.0, ((px - sx)*dx + (py - sy)*dy) / (dx*dx + dy*dy)))
    return math.hypot(px - (sx + t*dx), py - (sy + t*dy))


def _texts_near_segments(segs: list[Segment],
                          texts: list[dict], max_d: float) -> list[dict]:
    """Trả về texts có khoảng cách đến bất kỳ segment nào <= max_d, kèm min_dist."""
    result = []
    for tx in texts:
        best_d = float("inf")
        for s in segs:
            d = _pt_to_seg_dist(tx["x"], tx["y"], s.sx, s.sy, s.ex, s.ey)
            if d < best_d:
                best_d = d
        if best_d <= max_d:
            result.append({**tx, "seg_dist": best_d})
    return result


# ── Line type từ text ──────────────────────────────────────────────────────────

def _best_line_from_texts(candidates: list[dict]) -> tuple[str, str]:
    """
    Trả về (line_id, line_type) tốt nhất từ danh sách text candidates.
    Ưu tiên text chứa line ID có prefix rõ ràng, sau đó keyword.
    """
    if not candidates:
        return "", "UNKNOWN"

    # Sort theo seg_dist ASC
    candidates = sorted(candidates, key=lambda x: x.get("seg_dist", 9999))

    best_id, best_type, best_score = "", "UNKNOWN", -1

    for tx in candidates:
        raw = tx["text"]
        lt, lid, score = classify_line_from_text(raw)
        if score > best_score:
            best_score = score
            best_type  = lt
            best_id    = lid
        if score >= 10:   # line ID rõ ràng → dừng
            break

    return best_id, best_type


# ── Layer / Color fallback ─────────────────────────────────────────────────────
# ACI color → line type hint (theo quy ước P&ID ANHMINH)
_COLOR_MAP = {
    1:  "SANITARY",   # đỏ → Process
    6:  "SANITARY",   # magenta → CIP
    4:  "ICE_WATER",  # cyan → Water
    3:  "SANITARY",   # xanh lá → Signal / sanitary
    30: "STEAM",      # cam → Steam
    40: "STEAM",
    5:  "AIR",        # xanh dương → Air
    2:  "COOLING",    # vàng
}

_LAYER_KEYWORDS = {
    "PROCESS": "SANITARY", "PRODUCT": "SANITARY", "CIP": "SANITARY",
    "WATER": "ICE_WATER",  "ICE": "ICE_WATER",    "CHILLER": "ICE_WATER",
    "COOLING": "COOLING",  "COOL": "COOLING",
    "STEAM": "STEAM",      "STM": "STEAM",         "CONDENSATE": "STEAM",
    "AIR": "AIR",          "INSTRUMENT": "AIR",
}


def _classify_by_layer_color(layer: str, color: int) -> str:
    # Thử layer name trước
    lu = layer.upper()
    for kw, lt in _LAYER_KEYWORDS.items():
        if kw in lu:
            return lt
    # Thử color
    return _COLOR_MAP.get(color, "UNKNOWN")


# ── Proximity size ─────────────────────────────────────────────────────────────

def _nearest_dn(ex: float, ey: float,
                candidates: list, max_dist: float):
    best, best_d = None, float("inf")
    for t in candidates:
        d = math.hypot(t["x"] - ex, t["y"] - ey)
        if d < best_d:
            best_d, best = d, t
    if best and best_d <= max_dist:
        return best["dn"], round(best_d, 1)
    return "?", None


# ── Connection type ────────────────────────────────────────────────────────────

def _resolve_conn(line_type: str, size_str: str, spec: dict) -> str:
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


# ── Text helpers ───────────────────────────────────────────────────────────────

def _get_text(entity) -> str:
    try:
        if entity.dxftype() == "MTEXT":
            return entity.text or ""
        return entity.dxf.text or ""
    except Exception:
        return ""


def _strip_mtext(text: str) -> str:
    """Loại bỏ MTEXT formatting codes."""
    t = re.sub(r'\{\\[^}]*\}', '', text)          # {\\fFont|b0;...}
    t = re.sub(r'\\[A-Za-z][^;\\{} ]*;?', '', t)  # \\P \\H2; \\W0.8; etc.
    t = re.sub(r'[{}]', '', t)
    t = re.sub(r'\s+', ' ', t)
    return t.strip()


# ── Aggregate BOM ──────────────────────────────────────────────────────────────

def aggregate_bom(equipment: list[Equipment]) -> list[dict]:
    counts: dict[tuple, int]  = defaultdict(int)
    meta:   dict[tuple, dict] = {}

    for eq in equipment:
        key = (eq.line_type, eq.block_name, eq.size)
        counts[key] += 1
        if key not in meta:
            meta[key] = {
                "line_type":  eq.line_type,
                "line_id":    eq.line_id,
                "block_name": eq.block_name,
                "size":       eq.size,
                "conn_type":  eq.conn_type,
                **eq.spec,
            }

    rows = []
    for key, sl in counts.items():
        row = dict(meta[key])
        row["sl"] = sl
        rows.append(row)

    ORDER = {"SANITARY": 0, "ICE_WATER": 1, "COOLING": 2,
             "STEAM": 3, "AIR": 4, "UNKNOWN": 9}
    rows.sort(key=lambda r: (ORDER.get(r["line_type"], 9),
                              r["block_name"], r["size"]))
    return rows
