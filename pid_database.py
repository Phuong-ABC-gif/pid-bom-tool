# pid_database.py
# Block name → Equipment info mapping for SPX and Tetra Pak (ANH MINH 2021)
# Không dùng AI — tra cứu thuần túy từ tên block trong DXF
# Priority: kiểm tra pattern theo thứ tự từ trên xuống; match đầu tiên thắng.

import re

# ─────────────────────────────────────────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────────────────────────────────────────
def _eq(desc, code, material, standard, unit, section, subsection=""):
    """Tạo dict thông tin thiết bị."""
    return {
        "desc": desc,
        "code": code,
        "material": material,
        "standard": standard,
        "unit": unit,
        "section": section,        # 'sanitary' | 'utility' | 'steam' | 'instrument'
        "subsection": subsection,  # 'process' | 'chiller' | 'cooling' | 'steam' | ''
    }

# ─────────────────────────────────────────────────────────────────────────────
# IGNORE BLOCKS (không phải thiết bị)
# ─────────────────────────────────────────────────────────────────────────────
IGNORE_PATTERNS = re.compile(
    r'flow.?arrow|linenameTPleft|^PID$|^OBJECTS$|^title|^border|^north|'
    r'revision|^frame|legend|^tag|^note|^text|^line|^pipe.?label|^break',
    re.IGNORECASE
)

def is_ignore_block(name: str) -> bool:
    if not name or name.strip() == "":
        return True
    return bool(IGNORE_PATTERNS.search(name))


# ─────────────────────────────────────────────────────────────────────────────
# SPX DATABASE  (patterns checked top→bottom, first match wins)
# ─────────────────────────────────────────────────────────────────────────────
_SPX_RAW = [
    # ── Single Seat Valve 200 (SV41-L) ───────────────────────────────────────
    (r'SSV.?200.*THINKTOP|SSV.?200.*THIN.TOP',
     _eq("Single Seat Valve 200 SV41-L Thinktop",  "SSV-200-TK", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?200.*ACTUATOR|SSV.?200.*AIR.ACT|SSV.?200.*NC.*N.AIR|SSV.?200.*NO.*N.AIR',
     _eq("Single Seat Valve 200 SV41-L Actuator",  "SSV-200-A",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?200.*MANUAL|SSV.?200.*M\b',
     _eq("Single Seat Valve 200 SV41-L Manual",    "SSV-200-M",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bSSV.?200\b',
     _eq("Single Seat Valve 200 SV41-L",           "SSV-200",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Single Seat Valve 210 (SV43-LL) ──────────────────────────────────────
    (r'SSV.?210.*THINKTOP|SSV.?210.*THIN.TOP',
     _eq("Single Seat Valve 210 SV43-LL Thinktop", "SSV-210-TK", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?210.*ACTUATOR|SSV.?210.*AIR.ACT|SSV.?210.*NC.*N.AIR|SSV.?210.*NO.*N.AIR',
     _eq("Single Seat Valve 210 SV43-LL Actuator", "SSV-210-A",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?210.*MANUAL',
     _eq("Single Seat Valve 210 SV43-LL Manual",   "SSV-210-M",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bSSV.?210\b',
     _eq("Single Seat Valve 210 SV43-LL",          "SSV-210",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Single Seat Valve 220 (SV44-TL) ──────────────────────────────────────
    (r'SSV.?220.*THINKTOP|SSV.?220.*THIN.TOP',
     _eq("Single Seat Valve 220 SV44-TL Thinktop", "SSV-220-TK", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?220.*ACTUATOR|SSV.?220.*AIR.ACT',
     _eq("Single Seat Valve 220 SV44-TL Actuator", "SSV-220-A",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?220.*MANUAL',
     _eq("Single Seat Valve 220 SV44-TL Manual",   "SSV-220-M",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bSSV.?220\b',
     _eq("Single Seat Valve 220 SV44-TL",          "SSV-220",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Single Seat Valve 300 (SV42-T) ───────────────────────────────────────
    (r'SSV.?300.*THINKTOP|SSV.?300.*THIN.TOP',
     _eq("Single Seat Valve 300 SV42-T Thinktop",  "SSV-300-TK", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?300.*ACTUATOR|SSV.?300.*AIR.ACT',
     _eq("Single Seat Valve 300 SV42-T Actuator",  "SSV-300-A",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SSV.?300.*MANUAL',
     _eq("Single Seat Valve 300 SV42-T Manual",    "SSV-300-M",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bSSV.?300\b',
     _eq("Single Seat Valve 300 SV42-T",           "SSV-300",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Mixproof Valve ────────────────────────────────────────────────────────
    (r'MIXPROOF.*THINKTOP|MPV.?TK\b',
     _eq("Mixproof Valve Thinktop",  "MPV-TK", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'MIXPROOF.*ACTUATOR|MIXPROOF.*AIR|MPV.?A\b',
     _eq("Mixproof Valve Actuator",  "MPV-A",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bMIXPROOF\b',
     _eq("Mixproof Valve",           "MPV",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Butterfly Valve Utility (khai trước để tránh match nhầm sang sanitary)
    (r'\bBFV.?UA\b|BUTTERFLY.*UTIL.*ACT',
     _eq("Butterfly Valve Actuator (Utility)", "BFV-UA", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    (r'\bBFV.?UM\b|BUTTERFLY.*UTIL.*MAN',
     _eq("Butterfly Valve Manual (Utility)",   "BFV-UM", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    # ── Butterfly Valve (Sanitary) ────────────────────────────────────────────
    (r'\bBFV.?TK\b|BUTTERFLY.*THINKTOP|BUTTERFLY.*THIN.TOP',
     _eq("Butterfly Valve Thinktop (Sanitary)", "BFV-TK", "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bBFV.?A\b|BUTTERFLY.*ACTUATOR|BUTTERFLY.*AIR.*ACT',
     _eq("Butterfly Valve Actuator (Sanitary)", "BFV-A",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bBFV.?M\b|BUTTERFLY.*MANUAL|MANUAL.*BUTTERFLY',
     _eq("Butterfly Valve Manual (Sanitary)",   "BFV-M",  "SS316L", "SMS", "EA", "sanitary", "process")),

    # ── Sanitary Component ────────────────────────────────────────────────────
    (r'\bNRV\b|NON.?RETURN.?VALVE',
     _eq("Non-Return Valve",           "NRV",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bSG\b|SIGHT.?GLASS',
     _eq("Sight Glass",                "SG",   "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'SAMPLING.*ACTUATOR|SVA\b',
     _eq("Sampling Valve Actuator",    "SVA",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'SAMPLING.*MANUAL|SVM\b',
     _eq("Sampling Valve Manual",      "SVM",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bCPV\b|CONSTANT.?PRESSURE',
     _eq("Constant Pressure Valve",    "CPV",  "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Pumps (Sanitary) ──────────────────────────────────────────────────────
    (r'CENTRIFUGAL.*WS\+|CP.?WS\+',
     _eq("Centrifugal Pump WS+",  "CP-WS+", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'CENTRIFUGAL.*W\+|CP.?W\+',
     _eq("Centrifugal Pump W+",   "CP-W+",  "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'LOBE.*DW\+|LP.?DW\+',
     _eq("Lobe Pump DW+",         "LP-DW+", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bLRP\b|LIQUID.?RING.?PUMP',
     _eq("Liquid Ring Pump",      "LRP",    "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bVAC\b|VACUUM.?PUMP',
     _eq("Vacuum Pump",           "VAC",    "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bVP\b|VANE.?PUMP',
     _eq("Vane Pump",             "VP",     "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bCOMP\b|COMPRESSOR',
     _eq("Compressor",            "COMP",   "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bPRP\b|PERISTALTIC',
     _eq("Peristaltic Pump",      "PRP",    "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bBLW\b|BLOWER|FAN\b',
     _eq("Blower / Fan",          "BLW",    "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Equipment / Fitting ───────────────────────────────────────────────────
    (r'\bUC\b|UNION|COUPLING',
     _eq("Union/Coupling",   "UC",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bTC\b|TRI.?CLAMP|FERRULE',
     _eq("Tri-Clamp/Ferrule","TC",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bFL\b|FLANGE',
     _eq("Flange",           "FL",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bOR\b|ORIFICE',
     _eq("Orifice",          "OR",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bFH\b|FLEXIBLE.?HOSE',
     _eq("Flexible Hose",    "FH",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bCD\b|CLEANING.?DEVICE',
     _eq("Cleaning Device",  "CD",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bSF\b|STERILE.?FILTER',
     _eq("Sterile Filter",   "SF",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bFSH\b|INLINE.?FILTER.*FSH',
     _eq("Inline Filter FSH","FSH", "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bFSR\b|INLINE.?FILTER.*FSR',
     _eq("Inline Filter FSR","FSR", "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bPHE\b|PLATE.?HEAT.?EX',
     _eq("Plate Heat Exchanger",   "PHE", "SS316L", "SPX", "EA", "sanitary", "process")),
    (r'\bTHE\b|TUBULAR.?HEAT.?EX',
     _eq("Tubular Heat Exchanger", "THE", "SS316L", "SPX", "EA", "sanitary", "process")),

    # ── Utility Valves ────────────────────────────────────────────────────────
    (r'\bGV.?A\b|GLOBE.*ACTUATOR',
     _eq("Globe Valve Actuator",       "GV-A",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bGV.?M\b|GLOBE.*MANUAL',
     _eq("Globe Valve Manual",         "GV-M",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bGV.?S\b|GLOBE.*SOLENOID',
     _eq("Globe Valve Solenoid",       "GV-S",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bGV.?LC\b|GLOBE.*MODULATING|GLOBE.*LINEAR',
     _eq("Globe Valve Modulating",     "GV-LC",  "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?A\b|BALL.*ACTUATOR',
     _eq("Ball Valve Actuator",        "BV-A",   "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?M\b|BALL.*MANUAL|MANUAL.*BALL',
     _eq("Ball Valve Manual",          "BV-M",   "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?S\b|BALL.*SOLENOID',
     _eq("Ball Valve Solenoid",        "BV-S",   "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?LC\b|BALL.*MODULATING|BALL.*LINEAR',
     _eq("Ball Valve Modulating",      "BV-LC",  "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBFV.?UA\b|BUTTERFLY.*UTIL.*ACT',
     _eq("Butterfly Valve Actuator (Utility)", "BFV-UA", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    (r'\bBFV.?UM\b|BUTTERFLY.*UTIL.*MAN',
     _eq("Butterfly Valve Manual (Utility)",   "BFV-UM", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    (r'\bDV.?A\b|DIAPHRAGM.*ACTUATOR',
     _eq("Diaphragm Valve Actuator",   "DV-A",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bDV.?M\b|DIAPHRAGM.*MANUAL',
     _eq("Diaphragm Valve Manual",     "DV-M",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bDV.?S\b|DIAPHRAGM.*SOLENOID',
     _eq("Diaphragm Valve Solenoid",   "DV-S",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bASV.?A\b|ANGLE.?SEAT.*ACTUATOR',
     _eq("Angle Seat Valve Actuator",  "ASV-A",  "SS304",     "Industry", "EA", "utility", "steam")),
    (r'\bASV.?M\b|ANGLE.?SEAT.*MANUAL',
     _eq("Angle Seat Valve Manual",    "ASV-M",  "SS304",     "Industry", "EA", "utility", "steam")),
    (r'\bASV.?S\b|ANGLE.?SEAT.*SOLENOID',
     _eq("Angle Seat Valve Solenoid",  "ASV-S",  "SS304",     "Industry", "EA", "utility", "steam")),
    (r'\bAV.?M\b|ANGLE.?VALVE.*MANUAL',
     _eq("Angle Valve Manual",         "AV-M",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bCV\b|CHECK.?VALVE',
     _eq("Check Valve",                "CV",     "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bCSTVA?\b|CONSTANT.?VALVE',
     _eq("Constant Valve",             "CSTV",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bRPV\b|REDUCING.?PRESSURE',
     _eq("Reducing Pressure Valve",    "RPV",    "Cast Iron", "PN16",     "EA", "utility", "steam")),
    (r'\bARV\b|AIR.?RELIEF',
     _eq("Air Relief Valve",           "ARV",    "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bVRV\b|VACUUM.?RELIEF',
     _eq("Vacuum Relief Valve",        "VRV",    "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bPRV\b|PRESSURE.?RELIEF|PRESSURE.?SAFETY|PSV\b',
     _eq("Pressure Relief Valve",      "PRV",    "Cast Iron", "PN16",     "EA", "utility", "steam")),

    # ── Utility Components ────────────────────────────────────────────────────
    (r'\bEV\b|EXPANSION.?VESSEL',
     _eq("Expansion Vessel",  "EV",    "Carbon Steel", "Industry", "EA", "utility", "chiller")),
    (r'\bST\b|STEAM.?TRAP',
     _eq("Steam Trap",        "ST",    "Cast Iron",    "PN16",     "EA", "utility", "steam")),
    (r'\bYS\b|Y.?STRAINER',
     _eq("Y Strainer",        "YS",    "Cast Iron",    "Industry", "EA", "utility", "chiller")),
    (r'\bSY180\b|SYPHON.?180',
     _eq("Syphon 180°",       "SY180", "SS304",        "Industry", "EA", "utility", "steam")),
    (r'\bSY90\b|SYPHON.?90',
     _eq("Syphon 90°",        "SY90",  "SS304",        "Industry", "EA", "utility", "steam")),
    (r'\bSIL\b|SILENCER',
     _eq("Silencer",          "SIL",   "SS304",        "Industry", "EA", "utility", "steam")),
    (r'\bDP\b|DRAIN.?POINT',
     _eq("Drain Point",       "DP",    "SS304",        "Industry", "EA", "utility", "chiller")),
    (r'\bEJ\b|EXPANSION.?JOINT',
     _eq("Expansion Joint",   "EJ",    "Carbon Steel", "Industry", "EA", "utility", "chiller")),

    # ── Instruments (Sanitary) ────────────────────────────────────────────────
    (r'\bLSH\b|LEVEL.?SWITCH.?HIGH',
     _eq("Level Switch High",      "LSH",   "", "", "EA", "instrument", "")),
    (r'\bLSM\b|LEVEL.?SWITCH.?MED',
     _eq("Level Switch Medium",    "LSM",   "", "", "EA", "instrument", "")),
    (r'\bLSL\b|LEVEL.?SWITCH.?LOW',
     _eq("Level Switch Low",       "LSL",   "", "", "EA", "instrument", "")),
    (r'\bTT\b|TEMPERATURE.?TRANSMITTER',
     _eq("Temperature Transmitter","TT",    "", "", "EA", "instrument", "")),
    (r'\bTI\b|TEMPERATURE.?INDICATOR',
     _eq("Temperature Indicator",  "TI",    "", "", "EA", "instrument", "")),
    (r'\bPT\b|PRESSURE.?TRANSMITTER',
     _eq("Pressure Transmitter",   "PT",    "", "", "EA", "instrument", "")),
    (r'\bPI\b|PRESSURE.?INDICATOR',
     _eq("Pressure Indicator",     "PI",    "", "", "EA", "instrument", "")),
    (r'\bpHT\b|PH.?TRANSMITTER',
     _eq("pH Transmitter",         "pHT",   "", "", "EA", "instrument", "")),
    (r'\bBT\b|BRIX.?TRANSMITTER',
     _eq("Brix Transmitter",       "BT",    "", "", "EA", "instrument", "")),
    (r'\bPS.?PROX\b|PROXIMITY.?SWITCH',
     _eq("Proximity Switch",       "PS-prox","", "", "EA", "instrument", "")),
    (r'\bCS\b|CONDUCTIVITY',
     _eq("Conductivity Sensor",    "CS",    "", "", "EA", "instrument", "")),
    (r'\bSC\b|SPEED.?CONTROL',
     _eq("Speed Control",          "SC",    "", "", "EA", "instrument", "")),
    (r'\bFM\b|FLOW.?METER',
     _eq("Flow Meter",             "FM",    "", "", "EA", "instrument", "")),
    (r'\bFS\b|FLOW.?SWITCH',
     _eq("Flow Switch",            "FS",    "", "", "EA", "instrument", "")),
    (r'\bLS\b|LEVEL.?SWITCH\b',
     _eq("Level Switch",           "LS",    "", "", "EA", "instrument", "")),
    (r'\bPSw\b|PRESSURE.?SWITCH',
     _eq("Pressure Switch",        "PSw",   "", "", "EA", "instrument", "")),
    (r'\bPS.?V\b|PROXIMITY.?SWITCH.*VALVE',
     _eq("Proximity Switch – Valve","PS-V", "", "", "EA", "instrument", "")),
]

# ─────────────────────────────────────────────────────────────────────────────
# TETRA PAK DATABASE
# ─────────────────────────────────────────────────────────────────────────────
_TPV_RAW = [
    # ── Single Seat Valve 200 ─────────────────────────────────────────────────
    (r'SSV.?200.*THINKTOP|SSV.?200.*THIN.TOP',
     _eq("Single Seat Valve 200 Thinktop",  "SSV-200-TK", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?200.*ACTUATOR|SSV.?200.*AIR.ACT',
     _eq("Single Seat Valve 200 Actuator",  "SSV-200-A",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?200.*MANUAL',
     _eq("Single Seat Valve 200 Manual",    "SSV-200-M",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bSSV.?200\b',
     _eq("Single Seat Valve 200",           "SSV-200",    "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Single Seat Valve 210 ─────────────────────────────────────────────────
    (r'SSV.?210.*THINKTOP|SSV.?210.*THIN.TOP',
     _eq("Single Seat Valve 210 Thinktop",  "SSV-210-TK", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?210.*ACTUATOR|SSV.?210.*AIR.ACT',
     _eq("Single Seat Valve 210 Actuator",  "SSV-210-A",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?210.*MANUAL',
     _eq("Single Seat Valve 210 Manual",    "SSV-210-M",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bSSV.?210\b',
     _eq("Single Seat Valve 210",           "SSV-210",    "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Single Seat Valve 220 ─────────────────────────────────────────────────
    (r'SSV.?220.*THINKTOP|SSV.?220.*THIN.TOP',
     _eq("Single Seat Valve 220 Thinktop",  "SSV-220-TK", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?220.*ACTUATOR',
     _eq("Single Seat Valve 220 Actuator",  "SSV-220-A",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?220.*MANUAL',
     _eq("Single Seat Valve 220 Manual",    "SSV-220-M",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bSSV.?220\b',
     _eq("Single Seat Valve 220",           "SSV-220",    "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Single Seat Valve 300 ─────────────────────────────────────────────────
    (r'SSV.?300.*THINKTOP|SSV.?300.*THIN.TOP',
     _eq("Single Seat Valve 300 Thinktop",  "SSV-300-TK", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?300.*ACTUATOR',
     _eq("Single Seat Valve 300 Actuator",  "SSV-300-A",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?300.*MANUAL',
     _eq("Single Seat Valve 300 Manual",    "SSV-300-M",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bSSV.?300\b',
     _eq("Single Seat Valve 300",           "SSV-300",    "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Leakage Valve (Tetra Pak đặc thù) ────────────────────────────────────
    (r'\bLV.?NC\b|LEAKAGE.*NC',
     _eq("Leakage Valve NC Actuator",  "LV-NC", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bLV.?NO\b|LEAKAGE.*NO',
     _eq("Leakage Valve NO Actuator",  "LV-NO", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Butterfly Valve (Sanitary) ────────────────────────────────────────────
    (r'\bBFV.?TK\b|BUTTERFLY.*THINKTOP',
     _eq("Butterfly Valve Thinktop (Sanitary)", "BFV-TK", "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bBFV.?A\b(?!.*UTIL)|BUTTERFLY.*ACTUATOR(?!.*UTIL)',
     _eq("Butterfly Valve Actuator (Sanitary)", "BFV-A",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bBFV.?M\b|BUTTERFLY.*MANUAL|MANUAL.*BUTTERFLY',
     _eq("Butterfly Valve Manual (Sanitary)",   "BFV-M",  "SS316L", "SMS", "EA", "sanitary", "process")),

    # ── Constant Pressure Valve ───────────────────────────────────────────────
    (r'\bCPV\b|CONSTANT.?PRESSURE',
     _eq("Constant Pressure Valve", "CPV", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Mixproof Valve ────────────────────────────────────────────────────────
    (r'MIXPROOF.*THINKTOP|MPV.?TK\b',
     _eq("Mixproof Valve Thinktop", "MPV-TK", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'MIXPROOF.*ACTUATOR|MIXPROOF.*AIR|MPV.?A\b',
     _eq("Mixproof Valve Actuator", "MPV-A",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bMIXPROOF\b',
     _eq("Mixproof Valve",          "MPV",    "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Sanitary Component (Tetra Pak) ────────────────────────────────────────
    (r'\bNRV\b|NON.?RETURN.?VALVE',
     _eq("Non-Return Valve",         "NRV",     "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bSG\b|SIGHT.?GLASS',
     _eq("Sight Glass",              "SG",      "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bASV\b|ASEPTIC.?SAMPLING',
     _eq("Aseptic Sampling Valve",   "ASV",     "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'SSV.?SAMP|SANITARY.?SAMPLING',
     _eq("Sanitary Sampling Valve",  "SSV-samp","SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Pumps (Tetra Pak) ─────────────────────────────────────────────────────
    (r'\bCP\b|CENTRIFUGAL.?PUMP',
     _eq("Centrifugal Pump",   "CP",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bLP\b|LOBE.?PUMP',
     _eq("Lobe Pump",          "LP",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bSP\b|SCREW.?PUMP',
     _eq("Screw Pump",         "SP",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bUP\b|UTILITY.?PUMP',
     _eq("Utility Pump",       "UP",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bDP\b|DIAPHRAGM.?PUMP',
     _eq("Diaphragm Pump",     "DP",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bLRP\b|LIQUID.?RING',
     _eq("Liquid Ring Pump",   "LRP", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bVAC\b|VACUUM.?PUMP',
     _eq("Vacuum Pump",        "VAC", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Equipment (Tetra Pak) ─────────────────────────────────────────────────
    (r'\bUC\b|UNION|COUPLING',
     _eq("Union/Coupling",         "UC",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bTC\b|TRI.?CLAMP|FERRULE',
     _eq("Tri-Clamp/Ferrule",      "TC",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bFL\b|FLANGE',
     _eq("Flange",                 "FL",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bOR\b|ORIFICE',
     _eq("Orifice",                "OR",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bCD\b|CLEANING.?DEVICE',
     _eq("Cleaning Device",        "CD",  "SS316L", "SMS", "EA", "sanitary", "process")),
    (r'\bBF\b|BAG.?FILTER',
     _eq("Bag Filter",             "BF",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bILF\b|INLINE.?FILTER',
     _eq("Inline Filter",          "ILF", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bAF\b|ANGLE.?FILTER',
     _eq("Angle Filter",           "AF",  "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bPHE\b|PLATE.?HEAT',
     _eq("Plate Heat Exchanger",   "PHE", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),
    (r'\bTHE\b|TUBULAR.?HEAT',
     _eq("Tubular Heat Exchanger", "THE", "SS316L", "Tetra Pak", "EA", "sanitary", "process")),

    # ── Utility Valves (same as SPX) ──────────────────────────────────────────
    (r'\bGV.?A\b|GLOBE.*ACTUATOR',
     _eq("Globe Valve Actuator",     "GV-A",  "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bGV.?M\b|GLOBE.*MANUAL',
     _eq("Globe Valve Manual",       "GV-M",  "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?A\b|BALL.*ACTUATOR',
     _eq("Ball Valve Actuator",      "BV-A",  "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBV.?M\b|BALL.*MANUAL|MANUAL.*BALL',
     _eq("Ball Valve Manual",        "BV-M",  "Brass",     "Industry", "EA", "utility", "chiller")),
    (r'\bBFV.?UA\b|BUTTERFLY.*UTIL.*ACT',
     _eq("Butterfly Valve Actuator (Utility)", "BFV-UA", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    (r'\bBFV.?UM\b|BUTTERFLY.*UTIL.*MAN',
     _eq("Butterfly Valve Manual (Utility)",   "BFV-UM", "Cast Iron/Disc SS304", "JIS10K", "EA", "utility", "chiller")),
    (r'\bDV.?A\b|DIAPHRAGM.*ACTUATOR',
     _eq("Diaphragm Valve Actuator", "DV-A",  "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bDV.?M\b|DIAPHRAGM.*MANUAL',
     _eq("Diaphragm Valve Manual",   "DV-M",  "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bASV.?A\b|ANGLE.?SEAT.*ACT',
     _eq("Angle Seat Valve Actuator","ASV-A",  "SS304",    "Industry", "EA", "utility", "steam")),
    (r'\bCV\b|CHECK.?VALVE',
     _eq("Check Valve",              "CV",    "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bRPV\b|REDUCING.?PRESSURE',
     _eq("Reducing Pressure Valve",  "RPV",   "Cast Iron", "PN16",     "EA", "utility", "steam")),
    (r'\bARV\b|AIR.?RELIEF',
     _eq("Air Relief Valve",         "ARV",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bVRV\b|VACUUM.?RELIEF',
     _eq("Vacuum Relief Valve",      "VRV",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bPRV\b|PRESSURE.?RELIEF|PSV\b',
     _eq("Pressure Relief Valve",    "PRV",   "Cast Iron", "PN16",     "EA", "utility", "steam")),

    # ── Utility Components ────────────────────────────────────────────────────
    (r'\bST\b|STEAM.?TRAP',
     _eq("Steam Trap",   "ST",   "Cast Iron", "PN16",     "EA", "utility", "steam")),
    (r'\bYS\b|Y.?STRAINER',
     _eq("Y Strainer",   "YS",   "Cast Iron", "Industry", "EA", "utility", "chiller")),
    (r'\bSY180\b|SYPHON.?180',
     _eq("Syphon 180°",  "SY180","SS304",     "Industry", "EA", "utility", "steam")),
    (r'\bSY90\b|SYPHON.?90',
     _eq("Syphon 90°",   "SY90", "SS304",     "Industry", "EA", "utility", "steam")),
    (r'\bEV\b|EXPANSION.?VESSEL',
     _eq("Expansion Vessel","EV","Carbon Steel","Industry","EA","utility","chiller")),
    (r'\bEJ\b|EXPANSION.?JOINT',
     _eq("Expansion Joint","EJ", "Carbon Steel","Industry","EA","utility","chiller")),

    # ── Instruments ───────────────────────────────────────────────────────────
    (r'\bLSH\b', _eq("Level Switch High",      "LSH",  "","","EA","instrument","")),
    (r'\bLSM\b', _eq("Level Switch Medium",    "LSM",  "","","EA","instrument","")),
    (r'\bLSL\b', _eq("Level Switch Low",       "LSL",  "","","EA","instrument","")),
    (r'\bTT\b',  _eq("Temperature Transmitter","TT",   "","","EA","instrument","")),
    (r'\bTI\b',  _eq("Temperature Indicator",  "TI",   "","","EA","instrument","")),
    (r'\bPT\b',  _eq("Pressure Transmitter",   "PT",   "","","EA","instrument","")),
    (r'\bPI\b',  _eq("Pressure Indicator",     "PI",   "","","EA","instrument","")),
    (r'\bpHT\b', _eq("pH Transmitter",         "pHT",  "","","EA","instrument","")),
    (r'\bBT\b',  _eq("Brix Transmitter",       "BT",   "","","EA","instrument","")),
    (r'\bCS\b',  _eq("Conductivity Sensor",    "CS",   "","","EA","instrument","")),
    (r'\bFM\b',  _eq("Flow Meter",             "FM",   "","","EA","instrument","")),
    (r'\bFS\b',  _eq("Flow Switch",            "FS",   "","","EA","instrument","")),
]


# ─────────────────────────────────────────────────────────────────────────────
# COMPILE PATTERNS
# ─────────────────────────────────────────────────────────────────────────────
def _compile(raw_list):
    return [(re.compile(p, re.IGNORECASE), info) for p, info in raw_list]

SPX_DB  = _compile(_SPX_RAW)
TPV_DB  = _compile(_TPV_RAW)

DATABASES = {
    "SPX (ANH MINH 2021)":      SPX_DB,
    "Tetra Pak (ANH MINH 2021)": TPV_DB,
}


# ─────────────────────────────────────────────────────────────────────────────
# LOOKUP FUNCTION
# ─────────────────────────────────────────────────────────────────────────────
def lookup_block(block_name: str, db_key: str) -> dict | None:
    """
    Tra cứu tên block → thông tin thiết bị.
    Trả về dict hoặc None nếu không tìm thấy.
    """
    db = DATABASES.get(db_key, SPX_DB)
    name = block_name.strip()
    for pattern, info in db:
        if pattern.search(name):
            return info
    return None
