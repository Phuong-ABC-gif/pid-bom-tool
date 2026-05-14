"""
block_library.py
Thư viện block name → tên thiết bị + spec theo line type.
Tuân thủ skill ANH MINH 2021 (SPX + Tetra Pak).
"""
import re

# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZE & MATCH
# ══════════════════════════════════════════════════════════════════════════════

def normalize(name: str) -> set:
    name = name.lower()
    name = re.sub(r"[-_/]", " ", name)
    return set(name.split())


def is_match(block_name: str, lib_key: str) -> bool:
    b = normalize(block_name)
    l = normalize(lib_key)
    return b == l or b.issubset(l) or l.issubset(b)


# ══════════════════════════════════════════════════════════════════════════════
# BLOCK NAME → DISPLAY NAME  (word-order-insensitive)
# ══════════════════════════════════════════════════════════════════════════════

BLOCK_DISPLAY_MAP = {
    # ── Sanitary valves ────────────────────────────────────────────────────
    "ssv 200 manual":          "Single Seat Valve 200 Manual",
    "ssv 200 actuator":        "Single Seat Valve 200 Actuator NC",
    "ssv 200 thinktop":        "Single Seat Valve 200 Thinktop",
    "ssv 210 manual":          "Single Seat Valve 210 Manual",
    "ssv 210 actuator":        "Single Seat Valve 210 Actuator NC",
    "ssv 210 thinktop":        "Single Seat Valve 210 Thinktop",
    "ssv 220 manual":          "Single Seat Valve 220 Manual",
    "ssv 220 actuator":        "Single Seat Valve 220 Actuator NC",
    "ssv 220 thinktop":        "Single Seat Valve 220 Thinktop",
    "ssv 300 manual":          "Single Seat Valve 300 Manual",
    "ssv 300 actuator":        "Single Seat Valve 300 Actuator NC",
    "ssv 300 thinktop":        "Single Seat Valve 300 Thinktop",
    "sv41 manual":             "Single Seat Valve SV41 Manual",
    "sv41 actuator":           "Single Seat Valve SV41 Actuator NC",
    "sv41 thinktop":           "Single Seat Valve SV41 Thinktop",
    "sv43 manual":             "Single Seat Valve SV43 Manual",
    "sv43 actuator":           "Single Seat Valve SV43 Actuator NC",
    "sv43 thinktop":           "Single Seat Valve SV43 Thinktop",
    "sv44 manual":             "Single Seat Valve SV44 Manual",
    "sv44 actuator":           "Single Seat Valve SV44 Actuator NC",
    "sv44 thinktop":           "Single Seat Valve SV44 Thinktop",
    "double seal valve 200":   "Double Seal Valve 200",
    "double seal valve 210":   "Double Seal Valve 210",
    "double seal valve 220":   "Double Seal Valve 220",
    "double seal valve 300":   "Double Seal Valve 300",
    "mixproof valve actuator": "Mixproof Valve (MPV) Actuator NC",
    "mixproof valve thinktop": "Mixproof Valve (MPV) Thinktop",
    "mpv actuator":            "Mixproof Valve (MPV) Actuator NC",
    "mpv thinktop":            "Mixproof Valve (MPV) Thinktop",
    "butterfly valve manual":  "Butterfly Valve Manual",
    "butterfly valve actuator":"Butterfly Valve Actuator NC",
    "butterfly valve thinktop":"Butterfly Valve Thinktop",
    "bfv manual":              "Butterfly Valve Manual",
    "bfv actuator":            "Butterfly Valve Actuator NC",
    "bfv thinktop":            "Butterfly Valve Thinktop",
    "leakage valve nc":        "Leakage Valve NC Actuator",
    "leakage valve no":        "Leakage Valve NO Actuator",
    "lv nc":                   "Leakage Valve NC Actuator",
    "lv no":                   "Leakage Valve NO Actuator",
    "non return valve":        "Non-Return Valve (NRV)",
    "nrv":                     "Non-Return Valve (NRV)",
    "sampling valve manual":   "Sampling Valve Manual",
    "sampling valve actuator": "Sampling Valve Actuator",
    "svm":                     "Sampling Valve Manual",
    "sva":                     "Sampling Valve Actuator",
    "aseptic sampling valve":  "Aseptic Sampling Valve",
    "sanitary sampling valve": "Sanitary Sampling Valve",
    "sight glass":             "Sight Glass",
    "sg":                      "Sight Glass",
    "constant pressure valve": "Constant Pressure Valve (CPV)",
    "cpv":                     "Constant Pressure Valve (CPV)",
    "modulating valve rg41":   "Modulating Valve RG41",
    "modulating valve rg42":   "Modulating Valve RG42",
    "bag filter":              "Bag Filter",
    "inline filter":           "Inline Filter (177 micron)",
    "angle filter":            "Angle Filter (177 micron)",
    "sterile filter":          "Sterile Filter",
    # ── Sanitary pumps ────────────────────────────────────────────────────
    "centrifugal pump":        "Product Pump – Centrifugal",
    "centrifugal pump w+":     "Centrifugal Pump W+",
    "centrifugal pump ws+":    "Centrifugal Pump WS+",
    "lobe pump":               "Product Pump – Lobe Pump",
    "lobe pump dw+":           "Lobe Pump DW+",
    "cip pump":                "CIP Return Self-Priming Pump",
    "diaphragm pump":          "Chemical Pump – Diaphragm",
    "peristaltic pump":        "Peristaltic Pump",
    "screw pump":              "Screw Pump",
    # ── Equipment ─────────────────────────────────────────────────────────
    "plate heat exchanger":    "Plate Heat Exchanger (PHE)",
    "phe":                     "Plate Heat Exchanger (PHE)",
    "tubular heat exchanger":  "Tubular Heat Exchanger (THE)",
    "the":                     "Tubular Heat Exchanger (THE)",
    "union":                   "Union/Coupling",
    "tri clamp":               "Tri-Clamp/Ferrule",
    "tc":                      "Tri-Clamp/Ferrule",
    "flange":                  "Flange",
    "flexible hose":           "Flexible Hose",
    # ── Industry valves ───────────────────────────────────────────────────
    "globe valve manual":      "Globe Valve Manual",
    "globe valve actuator":    "Globe Valve Actuator NC",
    "globe valve solenoid":    "Globe Valve Solenoid",
    "globe valve modulating":  "Globe Valve Pneumatic Modulating",
    "gv manual":               "Globe Valve Manual",
    "gv actuator":             "Globe Valve Actuator NC",
    "ball valve manual":       "Ball Valve Manual",
    "ball valve actuator":     "Ball Valve Actuator NC",
    "ball valve solenoid":     "Ball Valve Solenoid",
    "bv manual":               "Ball Valve Manual",
    "bv actuator":             "Ball Valve Actuator NC",
    "butterfly valve utility manual": "Butterfly Valve Utility Manual",
    "butterfly valve utility actuator":"Butterfly Valve Utility Actuator NC",
    "bfv um":                  "Butterfly Valve Utility Manual",
    "bfv ua":                  "Butterfly Valve Utility Actuator NC",
    "diaphragm valve manual":  "Diaphragm Valve Manual",
    "diaphragm valve actuator":"Diaphragm Valve Actuator NC",
    "dv manual":               "Diaphragm Valve Manual",
    "angle seat valve manual": "Angle Seat Valve Manual",
    "angle seat valve actuator":"Angle Seat Valve Actuator NC",
    "angle seat valve solenoid":"Angle Seat Valve Solenoid",
    "asv manual":              "Angle Seat Valve Manual",
    "asv actuator":            "Angle Seat Valve Actuator NC",
    "check valve":             "Check Valve (CV)",
    "cv":                      "Check Valve (CV)",
    "y strainer":              "Y Strainer",
    "ys":                      "Y Strainer",
    "steam trap":              "Float Steam Trap",
    "st":                      "Float Steam Trap",
    "float steam trap":        "Float Steam Trap",
    "thermal steam trap":      "Thermal Dynamic Steam Trap",
    "pressure reducing valve": "Pressure Reducing Valve (PRV)",
    "prv":                     "Pressure Reducing Valve",
    "pressure relief valve":   "Pressure Relief Valve",
    "air relief valve":        "Air Relief Valve (ARV)",
    "arv":                     "Air Relief Valve (ARV)",
    "vacuum relief valve":     "Vacuum Relief Valve (VRV)",
    "vrv":                     "Vacuum Relief Valve (VRV)",
    "angle piston valve":      "Angle Piston Valve",
    "solenoid valve":          "Solenoid Valve 24V",
    "expansion vessel":        "Expansion Vessel",
    "ev":                      "Expansion Vessel",
    "syphon":                  "Syphon",
    "silencer":                "Silencer",
    # ── Instruments ───────────────────────────────────────────────────────
    "flow meter":              "Flow Meter",
    "fm":                      "Flow Meter",
    "flow switch":             "Flow Switch",
    "temperature indicator":   "Temperature Indicator",
    "temperature transmitter": "Temperature Transmitter",
    "pressure indicator":      "Pressure Indicator",
    "pressure transmitter":    "Pressure Transmitter",
    "level switch":            "Level Switch",
    "conductivity meter":      "Conductivity Meter",
    "proximity switch":        "Proximity Switch",
}


def lookup_display_name(block_name: str) -> str:
    """Trả về display name từ block name. Nếu không tìm thấy → trả lại block_name gốc."""
    for lib_key, display_name in BLOCK_DISPLAY_MAP.items():
        if is_match(block_name, lib_key):
            return display_name
    return block_name  # Không có trong thư viện → giữ nguyên để vẫn vào BOM


# ══════════════════════════════════════════════════════════════════════════════
# SPEC LOOKUP theo line_type + display_name
# ══════════════════════════════════════════════════════════════════════════════

def _dn_val(dn_str: str) -> int:
    """Trích số từ chuỗi DN như 'DN50' → 50. Trả -1 nếu không parse được."""
    m = re.search(r"\d+", str(dn_str))
    return int(m.group()) if m else -1


def apply_size_rule(spec: dict, dn_str: str, is_sanitary: bool) -> dict:
    """
    Quy tắc 2: size < DN50 → Thread end cho thiết bị công nghiệp.
    Sanitary không áp dụng.
    """
    if is_sanitary:
        return spec
    dn = _dn_val(dn_str)
    if dn > 0 and dn < 50:
        spec = dict(spec)
        spec["tieu_chuan"] = "Thread end"
        # Xóa "Flange RF" trong mô tả nếu có
        spec["mo_ta"] = re.sub(r",?\s*Flange RF", "", spec.get("mo_ta", ""), flags=re.IGNORECASE).strip()
    return spec


# ── Spec tables theo line_type ─────────────────────────────────────────────
# Format: key (lowercase keywords) → {chung_loai, vat_lieu, tieu_chuan, don_vi}

_SANITARY_SPEC = {
    "butterfly valve manual":       {"chung_loai": "Manual",      "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "butterfly valve actuator":     {"chung_loai": "Actuator NC", "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "butterfly valve thinktop":     {"chung_loai": "Thinktop",    "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "leakage valve":                {"chung_loai": "Actuator NC", "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "single seat valve":            {"chung_loai": "Actuator NC", "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "single seat valve manual":     {"chung_loai": "Manual",      "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "single seat valve thinktop":   {"chung_loai": "Thinktop",    "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "mixproof valve":               {"chung_loai": "Thinktop",    "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "double seal valve":            {"chung_loai": "Thinktop",    "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "non-return valve":             {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "sampling valve manual":        {"chung_loai": "Manual",      "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "sampling valve actuator":      {"chung_loai": "Actuator NC", "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "aseptic sampling valve":       {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "sanitary sampling valve":      {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "sight glass":                  {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "set"},
    "constant pressure valve":      {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "modulating valve":             {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "bag filter":                   {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "inline filter":                {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "angle filter":                 {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "sterile filter":               {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "centrifugal pump":             {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "lobe pump":                    {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "cip pump":                     {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "pcs"},
    "diaphragm pump":               {"chung_loai": "-",           "vat_lieu": "PP",     "tieu_chuan": "-",   "don_vi": "pcs"},
    "plate heat exchanger":         {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "set"},
    "tubular heat exchanger":       {"chung_loai": "-",           "vat_lieu": "SS316L", "tieu_chuan": "SMS", "don_vi": "set"},
}

_ICE_COOLING_SPEC = {
    "butterfly valve manual":       {"chung_loai": "Manual",      "vat_lieu": "Cast Iron / Disc SS304", "tieu_chuan": "JIS10K",    "don_vi": "pcs"},
    "butterfly valve actuator":     {"chung_loai": "Actuator NC", "vat_lieu": "Cast Iron / Disc SS304", "tieu_chuan": "JIS10K",    "don_vi": "pcs"},
    "butterfly valve modulating":   {"chung_loai": "Pneumatic modulating","vat_lieu": "Cast Iron / Disc SS304","tieu_chuan": "JIS10K","don_vi": "pcs"},
    "ball valve manual":            {"chung_loai": "Manual",      "vat_lieu": "Brass",       "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "ball valve actuator":          {"chung_loai": "Actuator NC", "vat_lieu": "Brass",       "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "check valve":                  {"chung_loai": "-",           "vat_lieu": "Cast Iron",   "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "y strainer":                   {"chung_loai": "-",           "vat_lieu": "Cast Iron",   "tieu_chuan": "JIS10K",     "don_vi": "pcs"},
    "pressure safety valve":        {"chung_loai": "-",           "vat_lieu": "Brass",       "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "expansion vessel":             {"chung_loai": "-",           "vat_lieu": "SS304",       "tieu_chuan": "Industry",   "don_vi": "set"},
}

_CITY_AIR_SPEC = {
    "butterfly valve manual":       {"chung_loai": "Manual",      "vat_lieu": "SS304", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "butterfly valve actuator":     {"chung_loai": "Actuator NC", "vat_lieu": "SS304", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "ball valve manual":            {"chung_loai": "Manual",      "vat_lieu": "SS304", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "ball valve actuator":          {"chung_loai": "Actuator NC", "vat_lieu": "SS304", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "check valve":                  {"chung_loai": "-",           "vat_lieu": "SS304", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "solenoid valve":               {"chung_loai": "-",           "vat_lieu": "SS304", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "angle piston valve":           {"chung_loai": "Actuator NC", "vat_lieu": "SS304", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "pressure safety valve":        {"chung_loai": "-",           "vat_lieu": "Brass", "tieu_chuan": "Thread end", "don_vi": "pcs"},
}

_STEAM_SPEC = {
    "globe valve manual":           {"chung_loai": "Manual",              "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "globe valve actuator":         {"chung_loai": "Actuator NC",         "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "globe valve modulating":       {"chung_loai": "Pneumatic Modulating","vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "globe valve":                  {"chung_loai": "Manual",              "vat_lieu": "Cast Iron", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "y strainer":                   {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "float steam trap":             {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "steam trap":                   {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "thermal steam trap":           {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "check valve":                  {"chung_loai": "-",                   "vat_lieu": "SS304",     "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "angle piston valve":           {"chung_loai": "Actuator NC",         "vat_lieu": "SS304",     "tieu_chuan": "Thread end", "don_vi": "pcs"},
    "pressure reducing valve":      {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "pressure safety valve":        {"chung_loai": "-",                   "vat_lieu": "Cast Iron", "tieu_chuan": "PN16",       "don_vi": "pcs"},
    "air filter regulator":         {"chung_loai": "-",                   "vat_lieu": "SS304",     "tieu_chuan": "PN16",       "don_vi": "pcs"},
}

_LINE_SPEC_MAP = {
    "Product":          _SANITARY_SPEC,
    "CIP":              _SANITARY_SPEC,
    "RO/Process Water": _SANITARY_SPEC,
    "Ice Water":        _ICE_COOLING_SPEC,
    "Cooling Water":    _ICE_COOLING_SPEC,
    "Compressed Air":   _CITY_AIR_SPEC,
    "City/Soft Water":  _CITY_AIR_SPEC,
    "Steam/Condensate": _STEAM_SPEC,
}

SANITARY_LINE_TYPES = {"Product", "CIP", "RO/Process Water"}
DEFAULT_SPEC = {"chung_loai": "-", "vat_lieu": "-", "tieu_chuan": "-", "don_vi": "pcs"}


def get_spec(display_name: str, line_type: str, dn_str: str) -> dict:
    """
    Tra cứu spec cho thiết bị theo line_type.
    Áp dụng quy tắc size < DN50 → Thread end (thiết bị công nghiệp).
    """
    table = _LINE_SPEC_MAP.get(line_type, {})
    dn = normalize(display_name)
    best_spec = None
    best_score = 0

    for key, spec in table.items():
        key_words = normalize(key)
        score = len(dn & key_words)
        if score > best_score:
            best_score = score
            best_spec = dict(spec)

    if best_spec is None:
        best_spec = dict(DEFAULT_SPEC)

    is_sanitary = line_type in SANITARY_LINE_TYPES
    best_spec["mo_ta"] = display_name
    return apply_size_rule(best_spec, dn_str, is_sanitary)
