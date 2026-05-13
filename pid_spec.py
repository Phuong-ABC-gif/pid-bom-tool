"""
pid_spec.py — Spec database theo chuẩn ANHMINH (SPX & Tetra Pak)
Quy tắc:
  - Xác định line type từ prefix tên line (block linename TP-left)
  - Áp spec vật liệu/tiêu chuẩn theo line type + equipment name + size
  - Size < DN50 → Thread end (chỉ áp cho thiết bị công nghiệp)
"""

import re

# ── 1. PHÂN LOẠI LINE THEO PREFIX ─────────────────────────────────────────────

LINE_GROUPS = {
    "SANITARY": [
        "PRD", "PRODUCT", "CIP", "PW", "SS", "RO"
    ],
    "ICE_WATER": [
        "IW", "IWR",
    ],
    "COOLING": [
        "CW", "TW", "TWR", "HW", "HWR",
    ],
    "STEAM": [
        "STM", "S-", "STEAM", "C-", "CON", "COND",
    ],
    "AIR": [
        "IA", "SA", "CA",
    ],
}

# Flatten để tìm nhanh
_WATER_LINES = LINE_GROUPS["ICE_WATER"] + LINE_GROUPS["COOLING"]


def classify_line(line_name: str) -> str:
    """
    Trả về: 'SANITARY' | 'ICE_WATER' | 'COOLING' | 'STEAM' | 'AIR' | 'UNKNOWN'
    Ví dụ: 'IW-001' → 'ICE_WATER', 'PRD-003' → 'SANITARY'
    """
    if not line_name:
        return "UNKNOWN"
    ln = line_name.upper().strip()
    for group, prefixes in LINE_GROUPS.items():
        for p in prefixes:
            if ln.startswith(p.upper()):
                return group
    return "UNKNOWN"


def get_line_category(line_type: str) -> str:
    """Trả về nhóm đề mục BOM."""
    return {
        "SANITARY":  "Đường Product / CIP / Process Water",
        "ICE_WATER": "Đường Chiller (Ice Water)",
        "COOLING":   "Đường Cooling Water",
        "STEAM":     "Đường Steam / Condensate",
        "AIR":       "Đường Air (Khí nén)",
        "UNKNOWN":   "Chưa xác định",
    }.get(line_type, "Chưa xác định")


# ── 2. PARSE SIZE ──────────────────────────────────────────────────────────────

def parse_dn(size_str: str) -> int:
    """
    'DN50' → 50, '50' → 50, '2"' → 50, '1.5"' → 40
    Trả về -1 nếu không parse được.
    """
    if not size_str or size_str.strip() == "?":
        return -1
    s = size_str.strip().upper()
    # DN notation
    m = re.search(r"DN\s*(\d+)", s)
    if m:
        return int(m.group(1))
    # Inch → DN
    m = re.search(r'(\d+(?:\.\d+)?)\s*"', s)
    if m:
        inch_map = {0.5: 15, 0.75: 20, 1.0: 25, 1.25: 32, 1.5: 40, 2.0: 50,
                    2.5: 65, 3.0: 80, 4.0: 100, 5.0: 125, 6.0: 150}
        v = float(m.group(1))
        return inch_map.get(v, int(v * 25))
    # Số thuần
    m = re.search(r"(\d+)", s)
    if m:
        return int(m.group(1))
    return -1


def is_thread_end(size_str: str) -> bool:
    """True nếu size < DN50 → dùng Thread end cho thiết bị công nghiệp."""
    dn = parse_dn(size_str)
    return 0 < dn < 50


# ── 3. SPEC DATABASE ───────────────────────────────────────────────────────────

# Key: (line_type, equip_key)  →  dict spec
# equip_key dạng lowercase, dùng is_match() để tìm

SANITARY_SPEC = {
    # format: chung_loai, vat_lieu, tieu_chuan, don_vi
    "pipe":                         ("Weld",           "SS316L", "SMS",      "m"),
    "elbow 90":                     ("Weld",           "SS316L", "SMS",      "pcs"),
    "tee":                          ("Weld",           "SS316L", "SMS",      "pcs"),
    "reduced tee":                  ("Weld",           "SS316L", "SMS",      "pcs"),
    "reducer":                      ("Weld",           "SS316L", "SMS",      "pcs"),
    "butterfly valve thinktop":     ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "butterfly valve actuator":     ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "butterfly valve manual":       ("Manual",         "SS316L", "SMS",      "pcs"),
    "butterfly valve":              ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "non return valve":             ("-",              "SS316L", "SMS",      "pcs"),
    "non-return valve":             ("-",              "SS316L", "SMS",      "pcs"),
    "nrv":                          ("-",              "SS316L", "SMS",      "pcs"),
    "ssv 200 manual":               ("Manual",         "SS316L", "SMS",      "pcs"),
    "ssv 200 actuator":             ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 200 thinktop":             ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "ssv 200":                      ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 210 manual":               ("Manual",         "SS316L", "SMS",      "pcs"),
    "ssv 210 actuator":             ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 210 thinktop":             ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "ssv 210":                      ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 220 manual":               ("Manual",         "SS316L", "SMS",      "pcs"),
    "ssv 220 actuator":             ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 220 thinktop":             ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "ssv 220":                      ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 300 manual":               ("Manual",         "SS316L", "SMS",      "pcs"),
    "ssv 300 actuator":             ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "ssv 300 thinktop":             ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "ssv 300":                      ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "single seat valve 200":        ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "single seat valve 210":        ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "single seat valve 220":        ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "single seat valve 300":        ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "mixproof valve actuator":      ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "mixproof valve thinktop":      ("Thinktop",       "SS316L", "SMS",      "pcs"),
    "mpv":                          ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "leakage valve nc":             ("Actuator NC",    "SS316L", "SMS",      "pcs"),
    "leakage valve no":             ("Actuator NO",    "SS316L", "SMS",      "pcs"),
    "sampling valve actuator":      ("Actuator",       "SS316L", "SMS",      "pcs"),
    "sampling valve manual":        ("Manual",         "SS316L", "SMS",      "pcs"),
    "aseptic sampling valve":       ("-",              "SS316L", "SMS",      "pcs"),
    "sanitary sampling valve":      ("-",              "SS316L", "SMS",      "pcs"),
    "sva":                          ("Actuator",       "SS316L", "SMS",      "pcs"),
    "svm":                          ("Manual",         "SS316L", "SMS",      "pcs"),
    "sight glass":                  ("-",              "SS316L", "SMS",      "set"),
    "sg":                           ("-",              "SS316L", "SMS",      "set"),
    "bag filter":                   ("-",              "SS316L", "SMS",      "pcs"),
    "inline filter":                ("-",              "SS316L", "SMS",      "pcs"),
    "angle filter":                 ("-",              "SS316L", "SMS",      "pcs"),
    "constant pressure valve":      ("-",              "SS316L", "SMS",      "pcs"),
    "cpv":                          ("-",              "SS316L", "SMS",      "pcs"),
    "centrifugal pump":             ("-",              "SS316L", "SMS",      "pcs"),
    "lobe pump":                    ("-",              "SS316L", "SMS",      "pcs"),
    "plate heat exchanger":         ("-",              "SS316L", "SMS",      "set"),
    "tubular heat exchanger":       ("-",              "SS316L", "SMS",      "set"),
}

# Đường Chiller / Cooling — phân theo size (DN50+ Flange, <DN50 Thread end)
def _water_spec(equip_key: str, size_str: str) -> dict:
    """Trả về spec cho đường nước lạnh/cooling theo size."""
    thread = is_thread_end(size_str)
    specs = {
        "pipe":                  ("Weld",              "Galv",                "BS1387M", "m"),
        "elbow":                 ("Weld",              "Galv",                "SCH20",   "pcs"),
        "tee":                   ("Weld",              "Galv",                "SCH20",   "pcs"),
        "reducer":               ("Weld",              "Galv",                "SCH20",   "pcs"),
        "flange":                ("Weld",              "Galv",                "JIS10K",  "pcs"),
        "butterfly valve":       ("Manual",            "Cast Iron/Disc SS304","JIS10K",  "pcs"),
        "butterfly valve manual":("Manual",            "Cast Iron/Disc SS304","JIS10K",  "pcs"),
        "butterfly valve actuator":("Actuator NC",     "Cast Iron/Disc SS304","JIS10K",  "pcs"),
        "ball valve":            ("Manual",            "Brass",               "Thread end","pcs"),
        "ball valve manual":     ("Manual",            "Brass",               "Thread end","pcs"),
        "ball valve actuator":   ("Actuator NC",       "Brass",               "Thread end","pcs"),
        "check valve":           ("-",                 "Cast Iron" if not thread else "Brass",
                                                        "PN16" if not thread else "Thread end","pcs"),
        "y strainer":            ("-",                 "Cast Iron",
                                                        "JIS10K" if not thread else "Thread end","pcs"),
        "pressure safety valve": ("-",                 "Brass",               "Thread end","pcs"),
    }
    # Butterfly valve: DN50+ Flange JIS10K, <DN50 → Ball Valve (không dùng GV)
    ek = equip_key.lower()
    if ("butterfly" in ek or "bfv" in ek) and thread:
        # < DN50, dùng Ball Valve thay thế
        return ("Manual", "Brass", "Thread end", "pcs")
    s = specs.get(ek)
    if s:
        return s
    # Fallback tìm partial match
    for k, v in specs.items():
        if k in ek or ek in k:
            return v
    return ("-", "Galv", "-", "pcs")


STEAM_SPEC_FLANGE = {
    "globe valve":               ("Manual",             "Cast Iron", "PN16",       "pcs"),
    "globe valve manual":        ("Manual",             "Cast Iron", "PN16",       "pcs"),
    "globe valve actuator":      ("Actuator NC",        "Cast Iron", "PN16",       "pcs"),
    "globe valve modulating":    ("Pneumatic Modulating","Cast Iron","PN16",       "pcs"),
    "y strainer":                ("-",                  "Cast Iron", "PN16",       "pcs"),
    "float steam trap":          ("-",                  "Cast Iron", "PN16",       "pcs"),
    "pressure reducing valve":   ("-",                  "Cast Iron", "PN16",       "pcs"),
    "pressure safety valve":     ("-",                  "Cast Iron", "PN16",       "pcs"),
    "check valve":               ("-",                  "SS304",     "PN16",       "pcs"),
}

STEAM_SPEC_THREAD = {
    "globe valve":               ("Manual",             "Cast Iron", "Thread end", "pcs"),
    "globe valve manual":        ("Manual",             "Cast Iron", "Thread end", "pcs"),
    "globe valve actuator":      ("Actuator NC",        "Cast Iron", "Thread end", "pcs"),
    "globe valve modulating":    ("Pneumatic Modulating","Cast Iron","Thread end", "pcs"),
    "angle piston valve":        ("Actuator NC",        "SS304",     "Thread end", "pcs"),
    "asv":                       ("Actuator NC",        "SS304",     "Thread end", "pcs"),
    "y strainer":                ("-",                  "Cast Iron", "Thread end", "pcs"),
    "float steam trap":          ("-",                  "Cast Iron", "Thread end", "pcs"),
    "thermal dynamic steam trap":("-",                  "Cast Iron", "Thread end", "pcs"),
    "bẫy đồng tiền":             ("-",                  "Cast Iron", "Thread end", "pcs"),
    "pressure safety valve":     ("-",                  "Brass",     "Thread end", "pcs"),
}

STEAM_PIPE_SPEC = {
    "pipe":        ("Weld", "Carbon Steel", "SCH40", "m"),
    "elbow":       ("Weld", "Carbon Steel", "SCH40", "pcs"),
    "tee":         ("Weld", "Carbon Steel", "SCH40", "pcs"),
    "reducer":     ("Weld", "Carbon Steel", "SCH40", "pcs"),
    "flange":      ("Weld", "Carbon Steel", "PN16",  "pcs"),
}


def get_spec(line_type: str, block_name: str, size_str: str) -> dict:
    """
    Trả về dict: {mo_ta, chung_loai, vat_lieu, tieu_chuan, don_vi}
    Áp dụng đúng quy tắc Thread end cho size < DN50.
    """
    ek = _normalize_key(block_name)
    thread = is_thread_end(size_str)

    if line_type == "SANITARY":
        # Tìm trong SANITARY_SPEC (longest match)
        spec = _find_spec(ek, SANITARY_SPEC)
        if spec:
            return _pack(spec)
        return _pack(("-", "SS316L", "SMS", "pcs"))

    elif line_type in ("ICE_WATER", "COOLING"):
        s = _water_spec(ek, size_str)
        return _pack(s)

    elif line_type == "STEAM":
        # Pipe/fitting
        for k, v in STEAM_PIPE_SPEC.items():
            if k in ek:
                return _pack(v)
        # Van
        if thread:
            spec = _find_spec(ek, STEAM_SPEC_THREAD)
        else:
            spec = _find_spec(ek, STEAM_SPEC_FLANGE)
        return _pack(spec) if spec else _pack(("-", "Cast Iron", "Thread end" if thread else "PN16", "pcs"))

    else:
        return _pack(("-", "-", "-", "pcs"))


def _normalize_key(name: str) -> str:
    s = name.lower()
    s = re.sub(r"[-_]", " ", s)
    return s.strip()


def _find_spec(ek: str, table: dict):
    # Exact match first
    if ek in table:
        return table[ek]
    # Longest substring match
    best, best_len = None, 0
    for k, v in table.items():
        if k in ek or ek in k:
            if len(k) > best_len:
                best, best_len = v, len(k)
    return best


def _pack(spec_tuple) -> dict:
    if len(spec_tuple) == 4:
        cl, vl, tc, dv = spec_tuple
        return {"chung_loai": cl, "vat_lieu": vl, "tieu_chuan": tc, "don_vi": dv}
    return {"chung_loai": "-", "vat_lieu": "-", "tieu_chuan": "-", "don_vi": "pcs"}


# ── 4. IGNORE LIST (không phải thiết bị) ──────────────────────────────────────

IGNORE_BLOCKS = {
    "flow arrow", "linenametpleft", "linename tp left", "linename-tp-left",
    "linenametpleft-1", "linenametpleft1", "linenametpleft-2", "linenametpleft2",
    "pid", "objects", "title block", "north arrow", "revision", "border",
    "flow direction", "flowarrow", "flow_arrow",
}


def should_ignore(block_name: str) -> bool:
    return _normalize_key(block_name) in IGNORE_BLOCKS


# ── 5. WORD-ORDER-INSENSITIVE MATCHING ────────────────────────────────────────

def normalize_words(name: str) -> set:
    s = name.lower()
    s = re.sub(r"[-_/]", " ", s)
    return set(s.split())


def is_match(block_name: str, library_name: str) -> bool:
    b = normalize_words(block_name)
    l = normalize_words(library_name)
    return b == l or b.issubset(l) or l.issubset(b)


# ── 6. LINENAME BLOCK PATTERNS ────────────────────────────────────────────────

LINENAME_PATTERNS = [
    "linenametpleft", "linename tp left", "linename-tp-left",
    "linenametpleft-1", "linename_tp_left", "tpleft", "tp-left",
    "linename", "line name",
]


def is_linename_block(block_name: str) -> bool:
    """True nếu block là block linename TP-left."""
    nk = _normalize_key(block_name)
    for p in LINENAME_PATTERNS:
        if p in nk:
            return True
    return False
