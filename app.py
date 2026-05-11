"""
P&ID BOM Generator — SPX + Tetra Pak (ANH MINH 2021)
Tuân thủ: pid-reader-spx + pid-reader-tetrapak SKILL.md
Tác giả: Tự động sinh từ skill rules
"""

import streamlit as st
import ezdxf
import math
import re
import io
from collections import defaultdict
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ══════════════════════════════════════════════════════════════════════════════
# 1. CONSTANTS & LOOKUP TABLES
# ══════════════════════════════════════════════════════════════════════════════

# Blocks chắc chắn không phải thiết bị → bỏ qua khi scan INSERT
IGNORE_BLOCK_KEYWORDS = {
    'linenamettpleft', 'linenametpvleft', 'linenametpleft',
    'linename', 'flow arrow', 'flowarrow', 'pid', 'north arrow',
    'title block', 'titleblock', 'revision', 'border', 'frame',
    'legend', 'notes', 'objects', 'viewport'
}

# Tên block chứa keyword nào → block đó là LINENAME (cần đọc attribute)
LINENAME_BLOCK_KEYWORDS = {
    'linename', 'linenamettpleft', 'linenametpvleft', 'linenametpleft',
    'tpleft', 'tpvleft', 'line name', 'linenum'
}

# Prefix line name → loại đường
LINE_PREFIX_MAP = {
    'PRD': 'sanitary', 'CIP': 'sanitary', 'PW':  'sanitary',
    'WW':  'sanitary', 'SW':  'sanitary',
    'IW':  'ice_water', 'IWR': 'ice_water',
    'CW':  'cooling',  'TWR': 'cooling',  'TW':  'cooling',
    'HW':  'hot_water', 'HWR': 'hot_water',
    'STM': 'steam',    'S':   'steam',    'SS':  'steam',
    'C':   'condensate', 'CN': 'condensate',
    'IA':  'air',      'SA':  'air',      'CA':  'air', 'PA': 'air',
    'RO':  'water',    'CW2': 'water',
}

LINE_TYPE_LABELS = {
    'sanitary':   '1. Đường Product / CIP / Process Water (Vi sinh)',
    'ice_water':  '2. Đường Ice Water – Chiller (Công nghiệp)',
    'cooling':    '3. Đường Cooling Water (Công nghiệp)',
    'hot_water':  '3b. Đường Hot Water (Công nghiệp)',
    'steam':      '4. Đường Steam (Công nghiệp)',
    'condensate': '4b. Đường Condensate (Công nghiệp)',
    'air':        '5. Đường Air (Khí nén)',
    'water':      '6. Đường RO / City Water',
    'unknown':    '0. Chưa xác định loại đường',
}

# Thư viện thiết bị: key = tên chuẩn hóa (lowercase, space)
# value = dict {desc, act, category}
# category: 'sanitary' hoặc 'utility'
EQUIPMENT_LIBRARY = {
    # ── SSV Tetra Pak ─────────────────────────────────────────────────────
    'ssv 200 manual':   {'desc': 'Single Seat Valve 200',        'act': 'Manual',        'cat': 'sanitary'},
    'ssv 200 actuator': {'desc': 'Single Seat Valve 200',        'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 200 thinktop': {'desc': 'Single Seat Valve 200',        'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 210 manual':   {'desc': 'Single Seat Valve 210',        'act': 'Manual',        'cat': 'sanitary'},
    'ssv 210 actuator': {'desc': 'Single Seat Valve 210',        'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 210 thinktop': {'desc': 'Single Seat Valve 210',        'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 220 manual':   {'desc': 'Single Seat Valve 220',        'act': 'Manual',        'cat': 'sanitary'},
    'ssv 220 actuator': {'desc': 'Single Seat Valve 220',        'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 220 thinktop': {'desc': 'Single Seat Valve 220',        'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 300 manual':   {'desc': 'Single Seat Valve 300',        'act': 'Manual',        'cat': 'sanitary'},
    'ssv 300 actuator': {'desc': 'Single Seat Valve 300',        'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 300 thinktop': {'desc': 'Single Seat Valve 300',        'act': 'Thinktop',      'cat': 'sanitary'},
    # ── SSV SPX (có model SV41/42/43/44) ─────────────────────────────────
    'ssv 200 sv41 l manual':   {'desc': 'SSV 200 SV41-L',  'act': 'Manual',        'cat': 'sanitary'},
    'ssv 200 sv41 l actuator': {'desc': 'SSV 200 SV41-L',  'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 200 sv41 l thinktop': {'desc': 'SSV 200 SV41-L',  'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 210 sv43 ll manual':  {'desc': 'SSV 210 SV43-LL', 'act': 'Manual',        'cat': 'sanitary'},
    'ssv 210 sv43 ll actuator':{'desc': 'SSV 210 SV43-LL', 'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 210 sv43 ll thinktop':{'desc': 'SSV 210 SV43-LL', 'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 220 sv44 tl manual':  {'desc': 'SSV 220 SV44-TL', 'act': 'Manual',        'cat': 'sanitary'},
    'ssv 220 sv44 tl actuator':{'desc': 'SSV 220 SV44-TL', 'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 220 sv44 tl thinktop':{'desc': 'SSV 220 SV44-TL', 'act': 'Thinktop',      'cat': 'sanitary'},
    'ssv 300 sv42 t manual':   {'desc': 'SSV 300 SV42-T',  'act': 'Manual',        'cat': 'sanitary'},
    'ssv 300 sv42 t actuator': {'desc': 'SSV 300 SV42-T',  'act': 'Actuator NC',   'cat': 'sanitary'},
    'ssv 300 sv42 t thinktop': {'desc': 'SSV 300 SV42-T',  'act': 'Thinktop',      'cat': 'sanitary'},
    # ── Shorthand SSV SPX ─────────────────────────────────────────────────
    'ssv 200 sv41': {'desc': 'SSV 200 SV41-L', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'ssv 210 sv43': {'desc': 'SSV 210 SV43-LL','act': 'Actuator NC', 'cat': 'sanitary'},
    'ssv 220 sv44': {'desc': 'SSV 220 SV44-TL','act': 'Actuator NC', 'cat': 'sanitary'},
    'ssv 300 sv42': {'desc': 'SSV 300 SV42-T', 'act': 'Actuator NC', 'cat': 'sanitary'},
    # ── MPV ───────────────────────────────────────────────────────────────
    'mpv actuator': {'desc': 'Mixproof Valve',  'act': 'Actuator NC', 'cat': 'sanitary'},
    'mpv thinktop': {'desc': 'Mixproof Valve',  'act': 'Thinktop',    'cat': 'sanitary'},
    'mpv a':        {'desc': 'Mixproof Valve',  'act': 'Actuator NC', 'cat': 'sanitary'},
    'mpv tk':       {'desc': 'Mixproof Valve',  'act': 'Thinktop',    'cat': 'sanitary'},
    # ── BFV Sanitary ──────────────────────────────────────────────────────
    'bfv manual':      {'desc': 'Butterfly Valve', 'act': 'Manual',      'cat': 'sanitary'},
    'bfv actuator':    {'desc': 'Butterfly Valve', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'bfv thinktop':    {'desc': 'Butterfly Valve', 'act': 'Thinktop',    'cat': 'sanitary'},
    'bfv m':           {'desc': 'Butterfly Valve', 'act': 'Manual',      'cat': 'sanitary'},
    'bfv a':           {'desc': 'Butterfly Valve', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'bfv tk':          {'desc': 'Butterfly Valve', 'act': 'Thinktop',    'cat': 'sanitary'},
    'butterfly valve manual':      {'desc': 'Butterfly Valve', 'act': 'Manual',      'cat': 'sanitary'},
    'butterfly valve actuator':    {'desc': 'Butterfly Valve', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'butterfly valve thinktop':    {'desc': 'Butterfly Valve', 'act': 'Thinktop',    'cat': 'sanitary'},
    # ── Leakage Valve (Tetra Pak đặc thù) ────────────────────────────────
    'lv nc':             {'desc': 'Leakage Valve NC', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'lv no':             {'desc': 'Leakage Valve NO', 'act': 'Actuator NO', 'cat': 'sanitary'},
    'leakage valve nc':  {'desc': 'Leakage Valve NC', 'act': 'Actuator NC', 'cat': 'sanitary'},
    'leakage valve no':  {'desc': 'Leakage Valve NO', 'act': 'Actuator NO', 'cat': 'sanitary'},
    # ── Sanitary Components ───────────────────────────────────────────────
    'nrv':              {'desc': 'Non-Return Valve',     'act': '—', 'cat': 'sanitary'},
    'non return valve': {'desc': 'Non-Return Valve',     'act': '—', 'cat': 'sanitary'},
    'sg':               {'desc': 'Sight Glass',          'act': '—', 'cat': 'sanitary'},
    'sight glass':      {'desc': 'Sight Glass',          'act': '—', 'cat': 'sanitary'},
    'cpv':              {'desc': 'Constant Pressure Valve', 'act': '—', 'cat': 'sanitary'},
    'asv':              {'desc': 'Aseptic Sampling Valve',  'act': '—', 'cat': 'sanitary'},
    'sva':              {'desc': 'Sampling Valve',       'act': 'Actuator', 'cat': 'sanitary'},
    'svm':              {'desc': 'Sampling Valve',       'act': 'Manual',   'cat': 'sanitary'},
    'sampling valve':   {'desc': 'Sampling Valve',       'act': '—',        'cat': 'sanitary'},
    'ssv samp':         {'desc': 'Sanitary Sampling Valve','act': '—',      'cat': 'sanitary'},
    # ── Pumps (Sanitary) ──────────────────────────────────────────────────
    'cp':      {'desc': 'Centrifugal Pump',   'act': '—', 'cat': 'sanitary'},
    'cp w+':   {'desc': 'Centrifugal Pump W+','act': '—', 'cat': 'sanitary'},
    'cp ws+':  {'desc': 'Centrifugal Pump WS+','act': '—','cat': 'sanitary'},
    'lp':      {'desc': 'Lobe Pump',          'act': '—', 'cat': 'sanitary'},
    'lp dw+':  {'desc': 'Lobe Pump DW+',      'act': '—', 'cat': 'sanitary'},
    'sp':      {'desc': 'Screw Pump',         'act': '—', 'cat': 'sanitary'},
    'dp pump': {'desc': 'Diaphragm Pump',     'act': '—', 'cat': 'sanitary'},
    'up':      {'desc': 'Utility Pump',       'act': '—', 'cat': 'utility'},
    'vac':     {'desc': 'Vacuum Pump',        'act': '—', 'cat': 'utility'},
    'lrp':     {'desc': 'Liquid Ring Pump',   'act': '—', 'cat': 'sanitary'},
    # ── Filters / Equipment ───────────────────────────────────────────────
    'bf':           {'desc': 'Bag Filter',        'act': '—', 'cat': 'sanitary'},
    'bag filter':   {'desc': 'Bag Filter',        'act': '—', 'cat': 'sanitary'},
    'ilf':          {'desc': 'Inline Filter',     'act': '—', 'cat': 'sanitary'},
    'af':           {'desc': 'Angle Filter',      'act': '—', 'cat': 'sanitary'},
    'sf':           {'desc': 'Sterile Filter',    'act': '—', 'cat': 'sanitary'},
    'fsh':          {'desc': 'Inline Filter FSH', 'act': '—', 'cat': 'sanitary'},
    'fsr':          {'desc': 'Inline Filter FSR', 'act': '—', 'cat': 'sanitary'},
    'phe':          {'desc': 'Plate Heat Exchanger',   'act': '—', 'cat': 'sanitary'},
    'the':          {'desc': 'Tubular Heat Exchanger', 'act': '—', 'cat': 'sanitary'},
    # ── Utility Valves ────────────────────────────────────────────────────
    'gv m':              {'desc': 'Globe Valve', 'act': 'Manual',              'cat': 'utility'},
    'gv a':              {'desc': 'Globe Valve', 'act': 'Actuator NC',         'cat': 'utility'},
    'gv s':              {'desc': 'Globe Valve', 'act': 'Solenoid',            'cat': 'utility'},
    'gv lc':             {'desc': 'Globe Valve', 'act': 'Pneumatic Modulating','cat': 'utility'},
    'globe valve manual':        {'desc': 'Globe Valve', 'act': 'Manual',              'cat': 'utility'},
    'globe valve actuator':      {'desc': 'Globe Valve', 'act': 'Actuator NC',         'cat': 'utility'},
    'globe valve modulating':    {'desc': 'Globe Valve', 'act': 'Pneumatic Modulating','cat': 'utility'},
    'globe valve':               {'desc': 'Globe Valve', 'act': 'Manual',              'cat': 'utility'},
    'bv m':              {'desc': 'Ball Valve',  'act': 'Manual',              'cat': 'utility'},
    'bv a':              {'desc': 'Ball Valve',  'act': 'Actuator NC',         'cat': 'utility'},
    'bv s':              {'desc': 'Ball Valve',  'act': 'Solenoid',            'cat': 'utility'},
    'ball valve manual':  {'desc': 'Ball Valve', 'act': 'Manual',              'cat': 'utility'},
    'ball valve actuator':{'desc': 'Ball Valve', 'act': 'Actuator NC',         'cat': 'utility'},
    'ball valve':         {'desc': 'Ball Valve', 'act': 'Manual',              'cat': 'utility'},
    'bfv um':            {'desc': 'Butterfly Valve (Utility)', 'act': 'Manual',       'cat': 'utility'},
    'bfv ua':            {'desc': 'Butterfly Valve (Utility)', 'act': 'Actuator NC',  'cat': 'utility'},
    'butterfly valve utility manual':  {'desc': 'Butterfly Valve (Utility)', 'act': 'Manual',      'cat': 'utility'},
    'butterfly valve utility actuator':{'desc': 'Butterfly Valve (Utility)', 'act': 'Actuator NC', 'cat': 'utility'},
    'cv':                {'desc': 'Check Valve', 'act': '—', 'cat': 'utility'},
    'check valve':       {'desc': 'Check Valve', 'act': '—', 'cat': 'utility'},
    'asv a':             {'desc': 'Angle Seat Valve', 'act': 'Actuator NC',    'cat': 'utility'},
    'asv s':             {'desc': 'Angle Seat Valve', 'act': 'Solenoid',       'cat': 'utility'},
    'angle seat valve':  {'desc': 'Angle Seat Valve', 'act': 'Actuator NC',    'cat': 'utility'},
    'angle piston valve':{'desc': 'Angle Piston Valve','act': 'Actuator NC',   'cat': 'utility'},
    'dv m':              {'desc': 'Diaphragm Valve', 'act': 'Manual',          'cat': 'utility'},
    'dv a':              {'desc': 'Diaphragm Valve', 'act': 'Actuator NC',     'cat': 'utility'},
    # ── Utility Components ────────────────────────────────────────────────
    'ys':                {'desc': 'Y Strainer',           'act': '—', 'cat': 'utility'},
    'y strainer':        {'desc': 'Y Strainer',           'act': '—', 'cat': 'utility'},
    'st':                {'desc': 'Steam Trap',           'act': '—', 'cat': 'utility'},
    'steam trap':        {'desc': 'Steam Trap',           'act': '—', 'cat': 'utility'},
    'float steam trap':  {'desc': 'Float Steam Trap',     'act': '—', 'cat': 'utility'},
    'thermal dynamic steam trap': {'desc': 'Thermal Dynamic Steam Trap (Bẫy đồng tiền)', 'act': '—', 'cat': 'utility'},
    'rpv':               {'desc': 'Pressure Reducing Valve','act': '—', 'cat': 'utility'},
    'pressure reducing valve': {'desc': 'Pressure Reducing Valve','act': '—','cat': 'utility'},
    'prv':               {'desc': 'Pressure Relief Valve', 'act': '—', 'cat': 'utility'},
    'pressure safety valve': {'desc': 'Pressure Safety Valve','act': '—', 'cat': 'utility'},
    'arv':               {'desc': 'Air Relief Valve',     'act': '—', 'cat': 'utility'},
    'vrv':               {'desc': 'Vacuum Relief Valve',  'act': '—', 'cat': 'utility'},
    'ev':                {'desc': 'Expansion Vessel',     'act': '—', 'cat': 'utility'},
    'ej':                {'desc': 'Expansion Joint',      'act': '—', 'cat': 'utility'},
    'sil':               {'desc': 'Silencer',             'act': '—', 'cat': 'utility'},
    'syphon':            {'desc': 'Syphon',               'act': '—', 'cat': 'utility'},
    'sy180':             {'desc': 'Syphon 180°',          'act': '—', 'cat': 'utility'},
    'sy90':              {'desc': 'Syphon 90°',           'act': '—', 'cat': 'utility'},
}

# ══════════════════════════════════════════════════════════════════════════════
# 2. CORE PARSING FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def normalize_block_name(name: str) -> set:
    """Chuẩn hóa block name → set of words (lowercase, replace - _ with space)."""
    name = name.lower().strip()
    name = re.sub(r'[-_/]', ' ', name)
    return set(name.split())


def is_linename_block(block_name: str) -> bool:
    """Kiểm tra xem block có phải là linename block không."""
    norm = block_name.lower().replace('-', '').replace('_', '')
    for kw in LINENAME_BLOCK_KEYWORDS:
        kw_norm = kw.replace(' ', '').replace('-', '').replace('_', '')
        if kw_norm in norm:
            return True
    return False


def is_ignore_block(block_name: str) -> bool:
    """True nếu block nằm trong danh sách bỏ qua."""
    norm_words = normalize_block_name(block_name)
    for kw in IGNORE_BLOCK_KEYWORDS:
        kw_words = set(kw.split())
        if kw_words.issubset(norm_words) or norm_words.issubset(kw_words):
            return True
    return False


def match_equipment(block_name: str) -> dict | None:
    """
    Word-order-insensitive matching: so tên block với thư viện.
    Trả về equipment info dict hoặc None nếu không tìm thấy.
    """
    b_words = normalize_block_name(block_name)
    best_match = None
    best_score = 0
    for lib_key, lib_info in EQUIPMENT_LIBRARY.items():
        l_words = set(lib_key.split())
        if b_words == l_words or b_words.issubset(l_words) or l_words.issubset(b_words):
            # Ưu tiên match dài hơn (nhiều từ trùng hơn)
            score = len(b_words & l_words)
            if score > best_score:
                best_score = score
                best_match = {**lib_info, 'matched_key': lib_key}
    return best_match


def get_line_type(line_name: str) -> str:
    """Xác định loại đường từ tên line (vd: IW-001 → ice_water)."""
    if not line_name:
        return 'unknown'
    prefix = re.split(r'[-_\d]', line_name.strip())[0].upper()
    return LINE_PREFIX_MAP.get(prefix, 'unknown')


def nearest_annotation(ex, ey, candidates, max_dist=300):
    """Tìm annotation gần nhất trong max_dist đơn vị. Trả về dn_str hoặc '?'."""
    best, best_d = None, float('inf')
    for t in candidates:
        d = math.hypot(t['x'] - ex, t['y'] - ey)
        if d < best_d:
            best_d, best = d, t
    if best and best_d <= max_dist:
        return best['dn'], round(best_d, 1)
    return '?', None


def get_dn_number(dn_str: str):
    """'DN50' → 50, '2"' → ~50 (approx), '?' → None."""
    if not dn_str or dn_str == '?':
        return None
    m = re.search(r'DN\s*(\d+)', dn_str, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # Inch conversion
    inch_m = re.search(r'(\d+(?:\.\d+)?)\s*["\']', dn_str)
    if inch_m:
        inch_map = {'0.5': 15, '0.75': 20, '1': 25, '1.25': 32, '1.5': 40,
                    '2': 50, '2.5': 65, '3': 80, '4': 100, '6': 150, '8': 200}
        return inch_map.get(inch_m.group(1), None)
    return None


def is_small_size(dn_str: str) -> bool:
    """True nếu size < DN50 → dùng Thread end."""
    n = get_dn_number(dn_str)
    if n is None:
        return False
    return n < 50


# ══════════════════════════════════════════════════════════════════════════════
# 3. SPEC LOOKUP — Quy tắc từ skill
# ══════════════════════════════════════════════════════════════════════════════

def get_spec(equip_info: dict, line_type: str, dn_str: str) -> dict:
    """
    Trả về spec: {material, standard, connection, warning}
    Tuân thủ:
      - Quy tắc 1: Xác định line trước, áp spec sau
      - Quy tắc 2: Size < DN50 → Thread end (chỉ cho thiết bị công nghiệp)
    """
    cat = equip_info.get('cat', 'utility')
    desc = equip_info.get('desc', '')
    small = is_small_size(dn_str)
    warn = ''

    # ── Sanitary (Vi sinh) ─────────────────────────────────────────────────
    # Áp dụng khi: cat='sanitary' HOẶC line_type='sanitary'
    # Quy tắc 2 KHÔNG áp dụng → dùng Clamp/SMS theo size riêng
    if cat == 'sanitary' or line_type == 'sanitary':
        return {'material': 'SS316L', 'standard': 'SMS',
                'connection': 'SMS', 'warning': ''}

    # ── Từ đây: thiết bị công nghiệp (utility) ────────────────────────────

    # ── Ice Water / Cooling Water ──────────────────────────────────────────
    if line_type in ('ice_water', 'cooling', 'hot_water', 'water'):
        if 'Globe Valve' in desc:
            return {'material': '—', 'standard': '—', 'connection': '—',
                    'warning': '⚠️ KHÔNG dùng Globe Valve trên đường nước!'}

        if small:  # DN < 50 → Thread end
            if 'Ball Valve' in desc:
                mat, std, conn = 'Brass', 'Thread end', 'Thread end'
            elif 'Check Valve' in desc:
                mat, std, conn = 'Brass', 'Thread end', 'Thread end'
            elif 'Y Strainer' in desc:
                mat, std, conn = 'Cast Iron', 'Thread end', 'Thread end'
            elif 'Pressure Safety' in desc:
                mat, std, conn = 'Brass', 'Thread end', 'Thread end'
            else:
                mat, std, conn = 'Brass', 'Thread end', 'Thread end'
        else:  # DN ≥ 50 → Flange JIS10K
            if 'Butterfly' in desc:
                mat, std, conn = 'Cast Iron / Disc SS304', 'JIS10K', 'Flange RF'
            elif 'Check Valve' in desc:
                mat, std, conn = 'Cast Iron', 'PN16', 'Flange RF'
            elif 'Y Strainer' in desc:
                mat, std, conn = 'Cast Iron', 'JIS10K', 'Flange RF'
            else:
                mat, std, conn = 'Cast Iron / Disc SS304', 'JIS10K', 'Flange RF'
        return {'material': mat, 'standard': std, 'connection': conn, 'warning': warn}

    # ── Steam / Condensate ─────────────────────────────────────────────────
    if line_type in ('steam', 'condensate'):
        if 'Globe Valve' in desc:
            if small:
                return {'material': 'Cast Iron', 'standard': 'Thread end',
                        'connection': 'Thread end', 'warning': ''}
            else:
                return {'material': 'Cast Iron', 'standard': 'PN16',
                        'connection': 'Flange RF', 'warning': ''}

        if 'Angle Piston' in desc or 'Angle Seat' in desc:
            # Angle piston valve luôn dùng Thread end (SS304)
            return {'material': 'SS304', 'standard': 'Thread end',
                    'connection': 'Thread end', 'warning': ''}

        if 'Y Strainer' in desc:
            if small:
                return {'material': 'Cast Iron', 'standard': 'Thread end',
                        'connection': 'Thread end', 'warning': ''}
            else:
                return {'material': 'Cast Iron', 'standard': 'PN16',
                        'connection': 'Flange RF', 'warning': ''}

        if 'Steam Trap' in desc:
            if 'Thermal Dynamic' in desc or small:
                return {'material': 'Cast Iron', 'standard': 'Thread end',
                        'connection': 'Thread end', 'warning': ''}
            else:
                return {'material': 'Cast Iron', 'standard': 'PN16',
                        'connection': 'Flange RF', 'warning': ''}

        if 'Pressure Reducing' in desc:
            return {'material': 'Cast Iron', 'standard': 'PN16',
                    'connection': 'Flange RF', 'warning': ''}

        if 'Safety' in desc or 'Relief' in desc:
            if small:
                return {'material': 'Brass', 'standard': 'Thread end',
                        'connection': 'Thread end', 'warning': ''}
            else:
                return {'material': 'Cast Iron', 'standard': 'PN16',
                        'connection': 'Flange RF', 'warning': ''}

        if 'Check Valve' in desc:
            if small:
                return {'material': 'SS304', 'standard': 'Thread end',
                        'connection': 'Thread end', 'warning': ''}
            else:
                return {'material': 'SS304', 'standard': 'PN16',
                        'connection': 'Flange RF', 'warning': ''}

        # Default steam valve
        if small:
            return {'material': 'Cast Iron', 'standard': 'Thread end',
                    'connection': 'Thread end', 'warning': ''}
        else:
            return {'material': 'Cast Iron', 'standard': 'PN16',
                    'connection': 'Flange RF', 'warning': ''}

    # ── Air / Instrument Air ───────────────────────────────────────────────
    if line_type == 'air':
        if small:
            return {'material': 'SS304', 'standard': 'Thread end',
                    'connection': 'Thread end', 'warning': ''}
        else:
            return {'material': 'SS304', 'standard': 'PN16',
                    'connection': 'Flange RF', 'warning': ''}

    # ── Unknown / Fallback ─────────────────────────────────────────────────
    return {'material': '—', 'standard': '—', 'connection': '—',
            'warning': '⚠️ Chưa xác định loại đường — kiểm tra linename block'}


# ══════════════════════════════════════════════════════════════════════════════
# 4. DXF READER
# ══════════════════════════════════════════════════════════════════════════════

def read_attribs(insert_entity) -> list[str]:
    """Đọc tất cả attribute text trong một INSERT entity."""
    texts = []
    try:
        for attrib in insert_entity.attribs:
            t = attrib.dxf.text.strip()
            if t:
                texts.append(t)
    except Exception:
        pass
    return texts


def parse_dxf(file_bytes: bytes) -> dict:
    """
    Parse DXF → trả về:
      {
        'linenames': [{'name': str, 'line_type': str, 'x': float, 'y': float}],
        'equipment': [{'block_name': str, 'x': float, 'y': float,
                       'equip_info': dict|None, 'dn': str, 'line_name': str,
                       'line_type': str, 'spec': dict}],
        'size_annotations': [{'dn': str, 'x': float, 'y': float}],
        'unknown_blocks': [str],
        'errors': [str],
      }
    """
    errors = []
    linenames = []
    equipment_raw = []
    size_annotations = []
    unknown_blocks = []

    try:
        stream = io.BytesIO(file_bytes)
        doc = ezdxf.read(stream)
        msp = doc.modelspace()
    except Exception as e:
        return {'linenames': [], 'equipment': [], 'size_annotations': [],
                'unknown_blocks': [], 'errors': [f'Lỗi đọc DXF: {e}']}

    # ── Pass 1: Thu thập tất cả entities ──────────────────────────────────
    for entity in msp:
        etype = entity.dxftype()

        if etype == 'INSERT':
            block_name = entity.dxf.name
            try:
                pos = entity.dxf.insert
                x, y = float(pos.x), float(pos.y)
            except Exception:
                continue

            # Linename block?
            if is_linename_block(block_name):
                attribs = read_attribs(entity)
                line_text = ' '.join(attribs).strip()
                # Lọc lấy tên line thực sự (vd: IW-001, PRD-003)
                line_codes = re.findall(r'[A-Z]{1,5}-?\d{3,}|[A-Z]{2,5}\d*', line_text)
                code = line_codes[0] if line_codes else line_text or block_name
                lt = get_line_type(code)
                linenames.append({'name': code, 'line_type': lt, 'x': x, 'y': y})
                continue

            # Equipment block
            if is_ignore_block(block_name):
                continue

            equipment_raw.append({'block_name': block_name, 'x': x, 'y': y})

        elif etype in ('TEXT', 'MTEXT'):
            try:
                raw_text = entity.text if etype == 'MTEXT' else entity.dxf.text
                pos = entity.dxf.insert
                x, y = float(pos.x), float(pos.y)
                # Tìm DN annotation
                dn_match = re.search(r'DN\s*\d+', raw_text, re.IGNORECASE)
                if dn_match:
                    size_annotations.append({
                        'dn': dn_match.group().upper().replace(' ', ''),
                        'x': x, 'y': y
                    })
                # Tìm inch annotation
                inch_match = re.search(r'\d+\.?\d*\s*["\']', raw_text)
                if inch_match and not dn_match:
                    size_annotations.append({
                        'dn': inch_match.group().strip(),
                        'x': x, 'y': y
                    })
            except Exception:
                pass

    # ── Pass 2: Gán line và size cho từng thiết bị ─────────────────────────
    equipment_out = []
    for eq in equipment_raw:
        # Tìm size gần nhất (Kỹ năng 1 — Proximity Detection)
        dn, dn_dist = nearest_annotation(eq['x'], eq['y'], size_annotations, max_dist=300)

        # Tìm linename gần nhất theo trục Y (cùng hàng ngang trên P&ID)
        # Ưu tiên linename ở bên TRÁI (x nhỏ hơn) và cùng y
        best_line = None
        best_line_dist = float('inf')
        for ln in linenames:
            # Khoảng cách theo y (cùng đường ống nằm ngang)
            y_dist = abs(ln['y'] - eq['y'])
            x_dist = abs(ln['x'] - eq['x'])
            dist = math.hypot(x_dist * 0.3, y_dist)  # ưu tiên y-alignment
            if dist < best_line_dist:
                best_line_dist = dist
                best_line = ln

        line_name = best_line['name'] if best_line else 'UNKNOWN'
        line_type = best_line['line_type'] if best_line else 'unknown'

        # Match equipment
        equip_info = match_equipment(eq['block_name'])
        if equip_info is None:
            unknown_blocks.append(eq['block_name'])

        # Lấy spec
        spec = {}
        if equip_info:
            spec = get_spec(equip_info, line_type, dn)

        equipment_out.append({
            'block_name': eq['block_name'],
            'x': eq['x'], 'y': eq['y'],
            'equip_info': equip_info,
            'dn': dn,
            'dn_dist': dn_dist,
            'line_name': line_name,
            'line_type': line_type,
            'spec': spec,
        })

    return {
        'linenames': linenames,
        'equipment': equipment_out,
        'size_annotations': size_annotations,
        'unknown_blocks': list(set(unknown_blocks)),
        'errors': errors,
    }


# ══════════════════════════════════════════════════════════════════════════════
# 5. POST-EXPORT VERIFICATION (Kỹ năng 2)
# ══════════════════════════════════════════════════════════════════════════════

def verify_bom(equipment_list: list, bom_rows: list) -> list[str]:
    """
    So sánh số lượng thiết bị trong DXF với BOM.
    Trả về list error strings.
    """
    errors = []
    # Build summary từ equipment_list
    summary = defaultdict(lambda: defaultdict(int))
    for eq in equipment_list:
        if eq['equip_info']:
            desc = eq['equip_info']['desc']
            summary[desc][eq['dn']] += 1

    # Kiểm tra 1: Tổng số lượng
    for desc, sizes in summary.items():
        total_dxf = sum(sizes.values())
        total_bom = sum(r['sl'] for r in bom_rows if desc in r.get('mo_ta', ''))
        if total_bom != total_dxf:
            errors.append(f'❌ TỔNG: {desc} → DXF={total_dxf}, BOM={total_bom}')

    # Kiểm tra 2: Số lượng theo size
    for desc, sizes in summary.items():
        for dn, count_dxf in sizes.items():
            count_bom = sum(r['sl'] for r in bom_rows
                            if desc in r.get('mo_ta', '') and r.get('kt1') == dn)
            if count_bom != count_dxf:
                errors.append(f'❌ SIZE: {desc} {dn} → DXF={count_dxf}, BOM={count_bom}')

    # Kiểm tra 3: Unknown size
    for desc, sizes in summary.items():
        if '?' in sizes:
            errors.append(f'⚠️  UNKNOWN SIZE: {desc} × {sizes["?"]} EA — cần kiểm tra bản vẽ')

    return errors


# ══════════════════════════════════════════════════════════════════════════════
# 6. BOM EXCEL GENERATOR — FORM_BOM_STANDARD
# ══════════════════════════════════════════════════════════════════════════════

def thin_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

BG_CYAN   = PatternFill('solid', fgColor='00FFFF')
BG_YELLOW = PatternFill('solid', fgColor='FFFF00')
BG_NONE   = PatternFill(fill_type=None)
FG_RED    = 'FF0000'
FG_BLACK  = '000000'


def _cell(ws, row, col, value='', bg=None, bold=False,
          color='000000', size=11, halign='center', valign='center'):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name='Arial', bold=bold, size=size, color=color)
    c.fill = bg if bg else BG_NONE
    c.border = thin_border()
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=True)
    return c


def write_company_header(ws):
    """Rows 1–15: Header công ty ANH MINH."""
    today = date.today().strftime('%d/%m/%Y')

    def merge_row(r, text, font_size=11, bold=False):
        ws.merge_cells(f'A{r}:J{r}')
        c = ws[f'A{r}']
        c.value = text
        c.font = Font(name='Arial', bold=bold, size=font_size)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = thin_border()
        ws.row_dimensions[r].height = max(15, font_size * 1.8)

    merge_row(1,  'ANH MINH', 24, True)
    merge_row(2,  'TECHNOLOGY TRADING COMPANY LIMITED', 12, True)
    merge_row(3,  'Add: 179/2 Nguyen Van Troi, Ward 11, Phu Nhuan Dist., Ho Chi Minh City, Vietnam')
    merge_row(4,  'Tel: (028) 38 44 77 69  |  Fax: (028) 38 44 77 70')
    merge_row(5,  'Email: info@anhminh.com.vn  |  Tax Code: 0303 657 602')
    merge_row(6,  'FACSIMILE TRANSMISSION', 14, True)

    # Rows 7–11: Metadata
    meta = [
        (7,  f'Ngày/Date: {today}',         'Gởi từ/From: ANH MINH'),
        (8,  'Số trang/Pages:',              'BG số/Quote No:'),
        (9,  'Dự án/Project:',               ''),
        (10, 'Chú ý/Attention:',             ''),
        (11, 'Điện thoại/Phone:',            ''),
    ]
    for r, left_txt, right_txt in meta:
        ws.merge_cells(f'A{r}:E{r}')
        ws.merge_cells(f'F{r}:J{r}')
        for col_range, txt in [('A', left_txt), ('F', right_txt)]:
            c = ws[f'{col_range}{r}']
            c.value = txt
            c.font = Font(name='Arial', size=11)
            c.border = thin_border()
            ws.row_dimensions[r].height = 15

    merge_row(12, 'BÁO GIÁ / QUOTATION', 18, True)
    merge_row(13, 'Kính gởi quý khách hàng, chúng tôi xin trân trọng báo giá như sau:')
    merge_row(14, 'A  PHẠM VI CÔNG VIỆC / SCOPE OF WORK: Cung cấp thiết bị theo danh mục đính kèm.')
    merge_row(15, 'B  ĐƠN GIÁ / PRICE: Xem chi tiết bên dưới.')


def write_column_header(ws, start_row=16):
    """Rows 16–18: Header cột."""
    col_widths = [5.57, 47.71, 15.14, 16.14, 11.29, 12.00, 10.71, 15.43, 6.71, 9.71]
    col_names  = ['STT', 'MÔ TẢ', 'CHỦNG LOẠI', 'VẬT LIỆU',
                  'K.THƯỚC 1', 'K.THƯỚC 2', 'TIÊU CHUẨN', 'XUẤT XỨ', 'ĐƠN VỊ', 'SỐ LƯỢNG']
    from openpyxl.utils import get_column_letter
    for col_idx, (name, width) in enumerate(zip(col_names, col_widths), 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        for r in range(start_row, start_row + 3):
            c = ws.cell(row=r, column=col_idx)
            c.fill = BG_CYAN
            c.font = Font(name='Arial', bold=True, size=11, color=FG_RED)
            c.border = thin_border()
            c.alignment = Alignment(horizontal='left' if col_idx == 2 else 'center',
                                    vertical='center', wrap_text=True)
        ws.merge_cells(start_row=start_row, start_column=col_idx,
                       end_row=start_row + 2, end_column=col_idx)
        ws.cell(row=start_row, column=col_idx).value = name
        ws.row_dimensions[start_row].height = 40


def write_section(ws, row, stt, mo_ta, level='major'):
    bg = BG_CYAN if level == 'major' else BG_YELLOW
    vals = [stt, mo_ta, '', '', '', '', '', '', '', '']
    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = bg
        c.font = Font(name='Arial', bold=True, size=11, color=FG_RED)
        c.border = thin_border()
        c.alignment = Alignment(
            horizontal='left' if col == 2 else 'center',
            vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 18


def write_data_row(ws, row, stt, mo_ta, chung_loai, vat_lieu,
                   kt1, kt2, tieu_chuan, xuat_xu, dv, sl, warn=''):
    vals = [stt, mo_ta, chung_loai, vat_lieu, kt1, kt2, tieu_chuan, xuat_xu, dv, sl]
    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = BG_NONE
        # Cảnh báo → chữ đỏ
        color = FG_RED if (warn and col == 2) else FG_BLACK
        c.font = Font(name='Arial', bold=False, size=11, color=color)
        c.border = thin_border()
        c.alignment = Alignment(
            horizontal='left' if col == 2 else 'center',
            vertical='center', wrap_text=True)
    # Nếu có warning → ghi vào cột MÔ TẢ
    if warn:
        ws.cell(row=row, column=2).value = f'{mo_ta}  {warn}'
    ws.row_dimensions[row].height = 15


def write_footer(ws, row):
    """Footer: TOTAL, VAT, GRAND TOTAL, notes, conditions."""
    footer_lines = [
        ('TOTAL (VND)', True),
        ('VAT 10% (VND)', True),
        ('GRAND TOTAL (VND)', True),
        ('* Ghi chú / Notes:', True),
        ('- Giá trên chưa bao gồm thuế VAT 10%', False),
        ('- Không bao gồm chi phí vận chuyển, lắp đặt', False),
        ('C  ĐIỀU KIỆN BÁN HÀNG / SALES CONDITION:', True),
        ('-  Địa điểm giao hàng / Delivery: Công trình', False),
        ('-  Thời gian giao hàng / Lead time: 7-10 tuần kể từ ngày đặt hàng', False),
        ('-  Thanh toán / Payment: 30% tạm ứng – 40% tập kết – 30% sau hoàn thành', False),
        ('-  Hiệu lực báo giá / Validity: 30 ngày kể từ ngày gửi báo giá', False),
        ('Chân thành cảm ơn sự quan tâm của Quý khách hàng!', False),
        ('Trân trọng kính chào / Best regards,', False),
    ]
    for line, bold in footer_lines:
        ws.merge_cells(f'A{row}:J{row}')
        c = ws[f'A{row}']
        c.value = line
        c.font = Font(name='Arial', bold=bold, size=11,
                      color=(FG_RED if bold else FG_BLACK))
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c.border = thin_border()
        ws.row_dimensions[row].height = 15
        row += 1
    return row


def build_bom_rows(equipment_list: list) -> tuple[list, list]:
    """
    Nhóm thiết bị theo line_type → tạo danh sách bom_rows.
    Trả về (bom_rows, summary_for_verify).
    """
    # Gom theo: line_type → {(desc, act, dn): count}
    groups = defaultdict(lambda: defaultdict(int))
    for eq in equipment_list:
        if eq['equip_info'] is None:
            continue
        info = eq['equip_info']
        line_type = eq['line_type']
        desc = info['desc']
        act  = info['act']
        dn   = eq['dn']
        spec = eq['spec']
        key  = (desc, act, dn, spec.get('material',''), spec.get('standard',''),
                spec.get('connection',''), spec.get('warning',''))
        groups[line_type][key] += 1

    bom_rows = []
    section_order = ['sanitary', 'ice_water', 'cooling', 'hot_water',
                     'steam', 'condensate', 'air', 'water', 'unknown']
    for lt in section_order:
        if lt not in groups:
            continue
        for key, count in sorted(groups[lt].items(), key=lambda x: x[0][0]):
            desc, act, dn, mat, std, conn, warn = key
            bom_rows.append({
                'line_type': lt,
                'mo_ta': desc,
                'chung_loai': act,
                'vat_lieu': mat,
                'kt1': dn if dn != '?' else '?',
                'kt2': '—',
                'tieu_chuan': conn,
                'xuat_xu': '',
                'dv': 'pcs',
                'sl': count,
                'warn': warn,
            })
    return bom_rows


def generate_excel(equipment_list: list, project_name: str = '') -> bytes:
    """Tạo file Excel BOM đúng chuẩn FORM_BOM_STANDARD."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BOQ'
    ws.sheet_view.showGridLines = False

    # Headers
    write_company_header(ws)
    write_column_header(ws, start_row=16)

    bom_rows = build_bom_rows(equipment_list)

    # Ghi dữ liệu bắt đầu từ row 19
    r = 19
    write_section(ws, r, 'A', 'PHẦN CƠ KHÍ', level='major')
    r += 1

    section_order = ['sanitary', 'ice_water', 'cooling', 'hot_water',
                     'steam', 'condensate', 'air', 'water', 'unknown']
    sub_idx = {lt: i + 1 for i, lt in enumerate(section_order)}
    roman  = ['I','II','III','IV','V','VI','VII','VIII','IX','X']

    current_lt = None
    item_num = 0
    for row_data in bom_rows:
        lt = row_data['line_type']
        if lt != current_lt:
            current_lt = lt
            item_num = 0
            sub_roman = roman[sub_idx.get(lt, 0) - 1] if sub_idx.get(lt, 0) <= len(roman) else '?'
            label = LINE_TYPE_LABELS.get(lt, lt)
            write_section(ws, r, sub_roman, label.split('. ', 1)[-1], level='sub')
            r += 1

        item_num += 1
        write_data_row(
            ws, r,
            stt=str(item_num),
            mo_ta=row_data['mo_ta'],
            chung_loai=row_data['chung_loai'],
            vat_lieu=row_data['vat_lieu'],
            kt1=row_data['kt1'],
            kt2=row_data['kt2'],
            tieu_chuan=row_data['tieu_chuan'],
            xuat_xu=row_data['xuat_xu'],
            dv=row_data['dv'],
            sl=row_data['sl'],
            warn=row_data['warn'],
        )
        r += 1

    write_footer(ws, r)

    # Xuất ra bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
# 7. STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title='P&ID BOM Generator — ANH MINH',
    page_icon='🔧',
    layout='wide',
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp { background: #f7f9fc; }
  .block-label { font-weight:600; color:#1a237e; }
  .warn-box { background:#fff3cd; border-left:4px solid #ffc107;
              padding:8px 12px; border-radius:4px; margin:4px 0; font-size:13px; }
  .ok-box   { background:#d4edda; border-left:4px solid #28a745;
              padding:8px 12px; border-radius:4px; margin:4px 0; font-size:13px; }
  .err-box  { background:#f8d7da; border-left:4px solid #dc3545;
              padding:8px 12px; border-radius:4px; margin:4px 0; font-size:13px; }
  .info-box { background:#d1ecf1; border-left:4px solid #17a2b8;
              padding:8px 12px; border-radius:4px; margin:4px 0; font-size:13px; }
  h1 { color:#1a237e !important; }
  h2 { color:#283593 !important; }
  h3 { color:#3949ab !important; }
</style>
""", unsafe_allow_html=True)

st.title('🔧 P&ID BOM Generator — ANH MINH')
st.caption('Tuân thủ: pid-reader-SPX + pid-reader-TPV | ANH MINH 2021')
st.markdown('---')

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header('⚙️ Cài đặt')
    project_name = st.text_input('Tên dự án', placeholder='ANHMINH-2025-XXX')
    max_dist = st.slider('Ngưỡng proximity (max_dist)', 50, 1000, 300, 50,
                         help='Khoảng cách tối đa (DXF units) để gán size cho thiết bị')
    show_unknown = st.checkbox('Hiển thị block chưa nhận diện', value=True)
    show_linenames = st.checkbox('Hiển thị linename blocks', value=True)

    st.markdown('---')
    st.subheader('📋 Quy tắc áp dụng')
    st.markdown("""
**Quy tắc 1 — Xác định Line trước**
- Đọc block `linename TPV-left`
- PRD/CIP/PW → Vi sinh (SS316L/SMS)
- IW/IWR → Chiller (Cast Iron / Brass)
- CW/TWR → Cooling (Cast Iron / Brass)
- STM/S → Steam (Cast Iron / Carbon Steel)

**Quy tắc 2 — Size < DN50 → Thread end**
- DN < 50 → Thread end
- DN ≥ 50 → Flange (PN16 / JIS10K)
- *Chỉ áp dụng thiết bị công nghiệp*
    """)

# ── Main ──────────────────────────────────────────────────────────────────────
st.subheader('📂 Upload file DXF')
uploaded = st.file_uploader(
    'Chọn file P&ID (.dxf)',
    type=['dxf'],
    help='File DXF bản vẽ P&ID theo chuẩn ANH MINH 2021 (SPX hoặc Tetra Pak)'
)

if uploaded is None:
    st.markdown("""
<div class="info-box">
ℹ️ Vui lòng upload file DXF P&ID.<br>
App sẽ tự động:<br>
&nbsp;1. Đọc <b>block linename TPV-left</b> → xác định loại đường<br>
&nbsp;2. Đọc <b>block name</b> thiết bị → nhận diện loại van/pump<br>
&nbsp;3. Đọc <b>size</b> gần nhất (proximity detection, max_dist=300)<br>
&nbsp;4. Áp spec: DN&lt;50 → Thread end; DN≥50 → Flange<br>
&nbsp;5. Xuất BOM chuẩn <b>FORM_BOM_STANDARD</b> (.xlsx)
</div>
""", unsafe_allow_html=True)
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
with st.spinner('🔍 Đang đọc file DXF...'):
    file_bytes = uploaded.read()
    result = parse_dxf(file_bytes)

# ── Errors ────────────────────────────────────────────────────────────────────
if result['errors']:
    for e in result['errors']:
        st.markdown(f'<div class="err-box">❌ {e}</div>', unsafe_allow_html=True)
    st.stop()

linenames = result['linenames']
equipment = result['equipment']
size_ann  = result['size_annotations']
unknowns  = result['unknown_blocks']

# ── Summary cards ─────────────────────────────────────────────────────────────
col1, col2, col3, col4 = st.columns(4)
col1.metric('Line names đọc được', len(linenames))
col2.metric('Thiết bị tìm thấy', len(equipment))
col3.metric('Size annotations', len(size_ann))
col4.metric('Block chưa nhận diện', len(unknowns))

st.markdown('---')

# ── Tab layout ────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(
    ['📋 Line Names', '⚙️ Thiết bị', '⚠️ Kiểm tra', '📥 Xuất BOM']
)

# ── Tab 1: Linenames ──────────────────────────────────────────────────────────
with tab1:
    st.subheader('Block Linename TPV-left / TP-left')
    if not linenames:
        st.markdown("""
<div class="warn-box">
⚠️ Không tìm thấy block linename trong file DXF.<br>
Kiểm tra tên block: app tìm các block chứa keyword
<code>linename</code>, <code>tpleft</code>, <code>tpvleft</code>.
</div>
""", unsafe_allow_html=True)
    else:
        import pandas as pd
        df_ln = pd.DataFrame(linenames)
        df_ln['line_type_label'] = df_ln['line_type'].map(
            lambda lt: LINE_TYPE_LABELS.get(lt, lt))
        df_ln = df_ln[['name', 'line_type_label', 'x', 'y']].rename(columns={
            'name': 'Tên line', 'line_type_label': 'Loại đường',
            'x': 'X', 'y': 'Y'
        })
        st.dataframe(df_ln, use_container_width=True, height=400)

# ── Tab 2: Equipment ──────────────────────────────────────────────────────────
with tab2:
    st.subheader('Danh sách thiết bị được nhận diện')

    # Filters
    col_a, col_b = st.columns(2)
    filter_lt = col_a.multiselect(
        'Lọc theo loại đường',
        options=list(LINE_TYPE_LABELS.keys()),
        format_func=lambda x: LINE_TYPE_LABELS.get(x, x)
    )
    filter_known = col_b.radio(
        'Hiển thị',
        ['Tất cả', 'Đã nhận diện', 'Chưa nhận diện'],
        horizontal=True
    )

    import pandas as pd
    rows = []
    for eq in equipment:
        info = eq['equip_info']
        spec = eq['spec']
        known = info is not None
        if filter_lt and eq['line_type'] not in filter_lt:
            continue
        if filter_known == 'Đã nhận diện' and not known:
            continue
        if filter_known == 'Chưa nhận diện' and known:
            continue

        dn_flag = '✅' if eq['dn'] != '?' else '❓'
        warn = spec.get('warning', '') if spec else ''
        rows.append({
            'Block Name': eq['block_name'],
            'Thiết bị': info['desc'] if info else '— UNKNOWN —',
            'Chủng loại': info['act'] if info else '',
            'Loại đường': LINE_TYPE_LABELS.get(eq['line_type'], eq['line_type']),
            'Line code': eq['line_name'],
            f'Size {dn_flag}': eq['dn'],
            'Vật liệu': spec.get('material', '') if spec else '',
            'Tiêu chuẩn': spec.get('connection', '') if spec else '',
            '⚠️': '⚠️' if warn else '',
        })

    if rows:
        df_eq = pd.DataFrame(rows)
        st.dataframe(df_eq, use_container_width=True, height=500)
        st.caption(f'Tổng: {len(rows)} thiết bị hiển thị')
    else:
        st.info('Không có thiết bị nào thỏa điều kiện lọc.')

    # Unknown blocks
    if show_unknown and unknowns:
        with st.expander(f'🔍 {len(unknowns)} block chưa nhận diện'):
            st.markdown('Các tên block dưới đây không khớp với thư viện thiết bị:')
            for ub in sorted(unknowns):
                st.code(ub)
            st.markdown("""
<div class="info-box">
💡 Nếu block là thiết bị thực sự, hãy thêm vào <code>EQUIPMENT_LIBRARY</code>
hoặc liên hệ để cập nhật skill.
</div>
""", unsafe_allow_html=True)

# ── Tab 3: Verification ───────────────────────────────────────────────────────
with tab3:
    st.subheader('🔍 Kỹ năng 2 — Tự kiểm tra BOM (Post-export Verification)')

    bom_rows_preview = build_bom_rows(equipment)
    verify_errors = verify_bom(equipment, bom_rows_preview)

    if not verify_errors:
        st.markdown(f"""
<div class="ok-box">
✅ BOM đã kiểm tra: {len(set(eq.get('equip_info', {}).get('desc','') for eq in equipment if eq['equip_info']))} loại thiết bị,
tổng {len([eq for eq in equipment if eq['equip_info']])} EA — <b>PASS</b>
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="err-box">❌ Phát hiện {len(verify_errors)} lỗi — cần kiểm tra:</div>',
                    unsafe_allow_html=True)
        for err in verify_errors:
            icon = '❌' if err.startswith('❌') else '⚠️'
            css_class = 'err-box' if icon == '❌' else 'warn-box'
            st.markdown(f'<div class="{css_class}">{err}</div>', unsafe_allow_html=True)

    # Warnings về line type
    unknown_line_eqs = [eq for eq in equipment if eq['line_type'] == 'unknown' and eq['equip_info']]
    if unknown_line_eqs:
        st.markdown('---')
        st.markdown(f'<div class="warn-box">⚠️ {len(unknown_line_eqs)} thiết bị chưa xác định được loại đường — spec có thể sai.</div>',
                    unsafe_allow_html=True)

    # Globe valve warnings
    globe_water_warn = [
        eq for eq in equipment
        if eq['spec'] and '⚠️' in eq['spec'].get('warning', '')
    ]
    if globe_water_warn:
        st.markdown('---')
        st.warning(f'⚠️ Phát hiện {len(globe_water_warn)} trường hợp vi phạm quy tắc đường nước (Globe Valve trên đường IW/CW):')
        for eq in globe_water_warn:
            st.markdown(f'- **{eq["block_name"]}** tại ({eq["x"]:.0f}, {eq["y"]:.0f}) trên line `{eq["line_name"]}`')

    # Stats
    st.markdown('---')
    st.subheader('📊 Thống kê nhanh')
    import pandas as pd

    summary_data = defaultdict(lambda: defaultdict(int))
    for eq in equipment:
        if eq['equip_info']:
            lt_label = LINE_TYPE_LABELS.get(eq['line_type'], eq['line_type'])
            dn = eq['dn']
            summary_data[lt_label][dn] += 1

    if summary_data:
        for lt_label, sizes in summary_data.items():
            total = sum(sizes.values())
            unknown_sz = sizes.get('?', 0)
            pct_ok = int((total - unknown_sz) / total * 100) if total else 0
            st.markdown(f'**{lt_label}** — {total} thiết bị | Size OK: {pct_ok}%')
            size_df = pd.DataFrame(
                [{'Size': k, 'Số lượng': v} for k, v in sorted(sizes.items())]
            )
            st.dataframe(size_df, use_container_width=False, height=None)


# ── Tab 4: Export BOM ─────────────────────────────────────────────────────────
with tab4:
    st.subheader('📥 Xuất BOM — FORM_BOM_STANDARD')

    # Preview
    bom_rows_final = build_bom_rows(equipment)
    if not bom_rows_final:
        st.warning('Không có dữ liệu thiết bị để xuất BOM. Kiểm tra lại file DXF.')
    else:
        import pandas as pd
        df_bom = pd.DataFrame([{
            'STT': '',
            'Loại đường': LINE_TYPE_LABELS.get(r['line_type'], r['line_type']),
            'Mô tả': r['mo_ta'],
            'Chủng loại': r['chung_loai'],
            'Vật liệu': r['vat_lieu'],
            'K.Thước 1': r['kt1'],
            'Tiêu chuẩn': r['tieu_chuan'],
            'ĐV': r['dv'],
            'SL': r['sl'],
            '⚠️': r['warn'][:30] if r['warn'] else '',
        } for r in bom_rows_final])

        st.markdown('**Preview BOM:**')
        st.dataframe(df_bom, use_container_width=True, height=400)
        st.caption(f'Tổng: {len(bom_rows_final)} dòng vật tư | '
                   f'{sum(r["sl"] for r in bom_rows_final)} EA')

        # Check có unknown size không
        unk_sz = sum(1 for r in bom_rows_final if r['kt1'] == '?')
        if unk_sz:
            st.markdown(f"""
<div class="warn-box">
⚠️ {unk_sz} dòng có size = '?' — không xác định được từ bản vẽ.<br>
Ô K.THƯỚC sẽ để trống — cần kiểm tra thủ công trước khi gửi báo giá.
</div>
""", unsafe_allow_html=True)

        st.markdown('---')
        # Export button
        xlsx_bytes = generate_excel(equipment, project_name)
        fname = f'BOM_{project_name or "output"}_{date.today().strftime("%Y%m%d")}.xlsx'

        st.download_button(
            label='⬇️  Tải xuống BOM Excel (FORM_BOM_STANDARD)',
            data=xlsx_bytes,
            file_name=fname,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
        )

        st.markdown("""
<div class="info-box">
📌 File Excel tuân thủ chuẩn <b>FORM_BOM_STANDARD</b> của ANH MINH:<br>
&nbsp;• Rows 1–15: Header công ty<br>
&nbsp;• Rows 16–18: Header cột (nền Cyan, chữ Đỏ)<br>
&nbsp;• Row 19+: Dữ liệu theo đề mục (Major=Cyan / Sub=Yellow / Item=White)<br>
&nbsp;• Footer: TOTAL / VAT 10% / GRAND TOTAL + Điều kiện bán hàng
</div>
""", unsafe_allow_html=True)

st.markdown('---')
st.caption('P&ID BOM Generator v1.0 | Tuân thủ pid-reader-SPX + pid-reader-TPV | ANH MINH 2021')
