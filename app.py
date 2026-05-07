import streamlit as st
import anthropic
import re
import json
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="P&ID BOM Tool – ANH MINH",
    page_icon="⚙️",
    layout="wide",
)

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&display=swap');
  
  html, body, [class*="css"] { font-family: 'JetBrains Mono', monospace; }

  .stApp { background: #0d1117; color: #e6edf3; }
  
  header[data-testid="stHeader"] { background: #0d1f3c; border-bottom: 1px solid #1e3a5f; }

  .main-title {
    background: linear-gradient(135deg, #0d1f3c 0%, #0a2744 100%);
    border: 1px solid #1e3a5f;
    border-radius: 12px;
    padding: 20px 28px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 16px;
  }

  .skill-badge {
    display: inline-block;
    padding: 4px 14px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.08em;
  }
  .spx-badge { background: rgba(0,170,255,0.15); color: #00aaff; border: 1px solid #1e5a8a; }
  .tpv-badge { background: rgba(0,220,130,0.15); color: #00dc82; border: 1px solid #0e5a3a; }

  .block-chip {
    display: inline-block;
    background: #091f38;
    border: 1px solid #1a3a5a;
    border-radius: 4px;
    padding: 2px 10px;
    font-size: 11px;
    color: #7fb3d0;
    margin: 2px;
  }
  .count-chip {
    background: #0066cc;
    border-radius: 3px;
    padding: 1px 5px;
    font-size: 10px;
    color: #fff;
    margin-left: 4px;
  }

  div[data-testid="stFileUploader"] {
    border: 2px dashed #1e3a5f;
    border-radius: 12px;
    background: rgba(13,31,60,0.4);
    padding: 8px;
  }

  .stButton > button {
    background: linear-gradient(135deg, #0066cc, #0044aa) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-weight: 700 !important;
    letter-spacing: 0.06em !important;
    padding: 12px 28px !important;
    width: 100% !important;
    font-size: 15px !important;
    box-shadow: 0 4px 20px rgba(0,102,204,0.4) !important;
  }
  .stButton > button:hover {
    background: linear-gradient(135deg, #0088ff, #0055cc) !important;
    box-shadow: 0 6px 28px rgba(0,170,255,0.5) !important;
  }

  .stRadio > div { flex-direction: row !important; gap: 12px; }
  .stRadio label { color: #7fb3d0 !important; }

  div[data-testid="stDataFrame"] { border: 1px solid #1e3a5f; border-radius: 10px; }

  .section-row { background: #00ffff; color: #000; font-weight: 700; }
  .sub-row     { background: #e0ffff; color: #2a6a6a; font-style: italic; }
  .data-row    { background: #0d1f3c; color: #e6edf3; }

  .bom-table { width: 100%; border-collapse: collapse; font-size: 12px; font-family: 'JetBrains Mono', monospace; }
  .bom-table th { background: #00aaff; color: #000; padding: 8px 12px; text-align: left; font-weight: 700; letter-spacing: 0.05em; white-space: nowrap; }
  .bom-table .sec  { background: #00ffff !important; color: #000; font-weight: 700; }
  .bom-table .sub  { background: #caf0f8 !important; color: #2a6a6a; font-style: italic; }
  .bom-table .dat  { background: #0d1f3c; color: #e6edf3; }
  .bom-table .dat:hover { background: #1a3a5a; }
  .bom-table td { padding: 6px 12px; border-bottom: 1px solid #1a2a3a; }
  .qty { color: #fff; font-weight: 700; text-align: center; }
  .mat { color: #00dc82; }
  .sz  { color: #ffaa44; }
  .tag { color: #aaccee; }
</style>
""", unsafe_allow_html=True)

# ─── SPX Prompts ──────────────────────────────────────────────────────────────
SPX_SYSTEM = """Bạn là chuyên gia đọc P&ID theo tiêu chuẩn SPX (ANH MINH 2021).
Nhiệm vụ: Nhận danh sách block name từ file DXF và xuất BOM (Bill of Materials).

MAPPING BLOCK → THIẾT BỊ (SPX):
- SSV-200* / SSV200* / SV41* → Single Seat Valve 200 (SV41-L), SS316L, SMS/DIN
- SSV-210* / SSV210* / SV43* → Single Seat Valve 210 (SV43-LL), SS316L, SMS/DIN
- SSV-220* / SSV220* / SV44* → Single Seat Valve 220 (SV44-TL), SS316L, SMS/DIN
- SSV-300* / SSV300* / SV42* → Single Seat Valve 300 (SV42-T), SS316L, SMS/DIN
- MIXPROOF* / MPV* → Mixproof Valve, SS316L, SMS/DIN
- BUTTERFLY*MANUAL* / BFV-M* → Butterfly Valve Manual, SS316L, SMS
- BUTTERFLY*ACTUATOR* / BFV-A* → Butterfly Valve Actuator, SS316L, SMS
- BUTTERFLY*THINKTOP* / BFV-TK* → Butterfly Valve Thinktop, SS316L, SMS
- CENTRIFUGAL*W+* / CP-W+* → Centrifugal Pump W+, SS316L, SPX
- CENTRIFUGAL*WS+* / CP-WS+* → Centrifugal Pump WS+, SS316L, SPX
- LOBE*DW+* / LP-DW+* → Lobe Pump DW+, SS316L, SPX
- SAMPLING*MANUAL* / SVM* → Sampling Valve Manual, SS316L, SMS
- SAMPLING*ACTUATOR* / SVA* → Sampling Valve Actuator, SS316L, SMS
- NRV* → Non-Return Valve, SS316L, SMS
- SG* / SIGHT* → Sight Glass, SS316L, SMS
- CPV* / CONSTANT* → Constant Pressure Valve, SS316L, SMS
- STERILE*FILTER* / SF* → Sterile Filter, SS316L, SPX
- PHE* / PLATE*HEAT* → Plate Heat Exchanger, SS316L, —
- GV* / GLOBE* → Globe Valve, Cast Iron, Thread/PN16
- BV* / BALL*VALVE* → Ball Valve, Brass, Thread end
- CHECK* / CV* → Check Valve, Cast Iron, PN16
- STRAINER* → Y Strainer, Cast Iron, JIS10K
- STEAM*TRAP* → Float Steam Trap, Cast Iron, Thread end
- LSH* / LSL* / LSM* / TI* / TT* / PI* / PT* / pHT* / BT* / FM* / FS* / CS* → Instrument (PHẦN INSTRUMENT)

PHÂN LOẠI SECTION (BOM chuẩn):
- I. PHẦN CƠ KHÍ → 1. Đường Product/CIP/Process Water (vi sinh), 2. Đường Chiller, 3. Đường Cooling, 4. Đường Steam
- II. PHẦN INSTRUMENT

Trả về JSON với format:
{
  "sections": [
    {
      "stt": "I",
      "mo_ta": "PHẦN CƠ KHÍ",
      "type": "section",
      "subsections": [
        {
          "stt": "1",
          "mo_ta": "Đường Product / CIP / Process Water",
          "type": "sub",
          "items": [
            {
              "stt": "a",
              "mo_ta": "Single Seat Valve 210",
              "chung_loai": "SSV-210 SV43-LL Actuator",
              "vat_lieu": "SS316L",
              "kt1": "2\"",
              "kt2": "",
              "tieu_chuan": "SMS",
              "xuat_xu": "",
              "don_vi": "EA",
              "so_luong": 3
            }
          ]
        }
      ]
    }
  ]
}

Chỉ trả về JSON thuần túy, không có markdown hay giải thích thêm."""

TPV_SYSTEM = """Bạn là chuyên gia đọc P&ID theo tiêu chuẩn Tetra Pak (ANH MINH 2021).
Nhiệm vụ: Nhận danh sách block name từ file DXF và xuất BOM (Bill of Materials).

MAPPING BLOCK → THIẾT BỊ (Tetra Pak):
- SSV-200* / SSV200* → Single Seat Valve 200, SS316L, SMS/DIN
- SSV-210* / SSV210* → Single Seat Valve 210, SS316L, SMS/DIN
- SSV-220* / SSV220* → Single Seat Valve 220, SS316L, SMS/DIN
- SSV-300* / SSV300* → Single Seat Valve 300, SS316L, SMS/DIN
- MIXPROOF* / MPV* → Mixproof Valve, SS316L, SMS/DIN
- BUTTERFLY*MANUAL* / BFV-M* → Butterfly Valve Manual, SS316L, SMS
- BUTTERFLY*ACTUATOR* / BFV-A* → Butterfly Valve Actuator, SS316L, SMS
- BUTTERFLY*THINKTOP* / BFV-TK* → Butterfly Valve Thinktop, SS316L, SMS
- LEAKAGE*NC* / LV-NC* → Leakage Valve NC Actuator, SS316L, DIN (Tetra Pak)
- LEAKAGE*NO* / LV-NO* → Leakage Valve NO Actuator, SS316L, DIN (Tetra Pak)
- CENTRIFUGAL* / CP* → Centrifugal Pump, SS316L, —
- LOBE* / LP* → Lobe Pump, SS316L, —
- SCREW*PUMP* / SP* → Screw Pump, SS316L, — (Tetra Pak)
- UTILITY*PUMP* / UP* → Utility Pump, SS316L, — (Tetra Pak)
- DIAPHRAGM*PUMP* / DP* → Diaphragm Pump, SS316L, — (Tetra Pak)
- ASEPTIC*SAMPLING* / ASV* → Aseptic Sampling Valve, SS316L, SMS
- SANITARY*SAMPLING* / SSV-samp* → Sanitary Sampling Valve, SS316L, SMS
- NRV* → Non-Return Valve, SS316L, SMS
- SG* / SIGHT* → Sight Glass, SS316L, SMS
- CPV* → Constant Pressure Valve, SS316L, SMS
- BAG*FILTER* / BF* → Bag Filter, SS316L, —
- ANGLE*FILTER* / AF* → Angle Filter, SS316L, — (Tetra Pak)
- PHE* → Plate Heat Exchanger, SS316L, —
- GV* / GLOBE* → Globe Valve, Cast Iron, Thread/PN16
- BV* / BALL* → Ball Valve, Brass, Thread end
- CHECK* / CV* → Check Valve, Cast Iron, PN16
- STRAINER* → Y Strainer, Cast Iron, JIS10K
- STEAM*TRAP* → Float Steam Trap, Cast Iron, Thread end
- LSH* / LSL* / LSM* / TI* / TT* / PI* / PT* / pHT* / BT* / FM* / FS* / CS* → Instrument (PHẦN INSTRUMENT)

PHÂN LOẠI SECTION (BOM chuẩn):
- I. PHẦN CƠ KHÍ → 1. Đường Product/CIP/Process Water (vi sinh), 2. Đường Chiller, 3. Đường Cooling, 4. Đường Steam
- II. PHẦN INSTRUMENT

Trả về JSON với format chính xác như sau (chỉ JSON thuần):
{
  "sections": [
    {
      "stt": "I",
      "mo_ta": "PHẦN CƠ KHÍ",
      "type": "section",
      "subsections": [
        {
          "stt": "1",
          "mo_ta": "Đường Product / CIP / Process Water",
          "type": "sub",
          "items": [
            {
              "stt": "a",
              "mo_ta": "Single Seat Valve 210",
              "chung_loai": "SSV-210 Actuator",
              "vat_lieu": "SS316L",
              "kt1": "2\"",
              "kt2": "",
              "tieu_chuan": "SMS",
              "xuat_xu": "",
              "don_vi": "EA",
              "so_luong": 3
            }
          ]
        }
      ]
    }
  ]
}"""

# ─── DXF Parser ───────────────────────────────────────────────────────────────
def parse_dxf(text):
    lines = text.splitlines()
    inserts = []
    texts_found = []
    i = 0
    while i < len(lines) - 1:
        code = lines[i].strip()
        val  = lines[i+1].strip()
        if code == "0" and val == "INSERT":
            i += 2
            block_name = ""
            layer = ""
            while i < len(lines) - 1:
                c = lines[i].strip()
                v = lines[i+1].strip()
                if c == "0":
                    break
                if c == "2":
                    block_name = v
                if c == "8":
                    layer = v
                i += 2
            if block_name:
                inserts.append({"block": block_name.upper(), "layer": layer})
            continue
        if code == "0" and val in ("TEXT", "MTEXT"):
            i += 2
            txt = ""
            while i < len(lines) - 1:
                c = lines[i].strip()
                v = lines[i+1].strip()
                if c == "0":
                    break
                if c in ("1", "3"):
                    txt += v + " "
                i += 2
            if txt.strip():
                texts_found.append(txt.strip())
            continue
        i += 2

    block_count = {}
    for ins in inserts:
        k = ins["block"]
        block_count[k] = block_count.get(k, 0) + 1

    return block_count, texts_found, len(inserts)

# ─── Excel Export ─────────────────────────────────────────────────────────────
def create_excel(bom_data, skill):
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"

    # Fills
    fill_header  = PatternFill("solid", fgColor="00CCFF")
    fill_section = PatternFill("solid", fgColor="00FFFF")
    fill_sub     = PatternFill("solid", fgColor="E0FFFF")
    fill_data    = PatternFill("solid", fgColor="FFFFFF")

    font_header  = Font(name="Arial", bold=True, size=10, color="000000")
    font_section = Font(name="Arial", bold=True, size=10, color="000000")
    font_sub     = Font(name="Arial", italic=True, size=9, color="2A6A6A")
    font_data    = Font(name="Arial", size=9, color="000000")

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Column widths
    col_widths = [6, 42, 30, 16, 12, 12, 14, 12, 8, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Header row
    headers = ["STT", "MÔ TẢ", "CHỦNG LOẠI", "VẬT LIỆU", "K.THƯỚC 1", "K.THƯỚC 2", "TIÊU CHUẨN", "XUẤT XỨ", "ĐƠN VỊ", "SỐ LƯỢNG"]
    ws.append(headers)
    for cell in ws[ws.max_row]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[ws.max_row].height = 22

    # Data rows
    for section in bom_data.get("sections", []):
        ws.append([section["stt"], section["mo_ta"], "", "", "", "", "", "", "", ""])
        r = ws.max_row
        for col in range(1, 11):
            cell = ws.cell(row=r, column=col)
            cell.fill = fill_section
            cell.font = font_section
            cell.border = border
            if col == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        for sub in section.get("subsections", []):
            ws.append([sub["stt"], sub["mo_ta"], "", "", "", "", "", "", "", ""])
            r = ws.max_row
            for col in range(1, 11):
                cell = ws.cell(row=r, column=col)
                cell.fill = fill_sub
                cell.font = font_sub
                cell.border = border

            for item in sub.get("items", []):
                ws.append([
                    item.get("stt",""), item.get("mo_ta",""), item.get("chung_loai",""),
                    item.get("vat_lieu",""), item.get("kt1",""), item.get("kt2",""),
                    item.get("tieu_chuan",""), item.get("xuat_xu",""),
                    item.get("don_vi",""), item.get("so_luong",""),
                ])
                r = ws.max_row
                for col in range(1, 11):
                    cell = ws.cell(row=r, column=col)
                    cell.fill = fill_data
                    cell.font = font_data
                    cell.border = border
                    if col == 10:
                        cell.alignment = Alignment(horizontal="center")
                ws.row_dimensions[r].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── AI Call ─────────────────────────────────────────────────────────────────
def call_claude(block_count, texts, total, file_name, skill, api_key):
    client = anthropic.Anthropic(api_key=api_key)

    block_list = "\n".join(
        f"  {name} × {cnt}"
        for name, cnt in sorted(block_count.items(), key=lambda x: -x[1])
    )
    user_msg = f"""File DXF: {file_name}
Tổng số INSERT: {total}

DANH SÁCH BLOCK (tên × số lượng):
{block_list}

{f"TEXT ghi chú trong bản vẽ:{chr(10)}{' | '.join(texts[:40])}" if texts else ""}

Hãy tạo BOM đầy đủ. Phân loại vi sinh và công nghiệp đúng section. 
Nếu block name chứa kích thước (1", 1.5", 2", DN50...) điền vào kt1.
Chỉ trả về JSON."""

    system = SPX_SYSTEM if skill == "SPX" else TPV_SYSTEM

    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user_msg}],
    )
    raw = resp.content[0].text
    clean = re.sub(r"```json|```", "", raw).strip()
    return json.loads(clean)

# ─── UI ───────────────────────────────────────────────────────────────────────
# Header
st.markdown("""
<div style="background:linear-gradient(135deg,#0d1f3c,#0a2744);border:1px solid #1e3a5f;
border-radius:12px;padding:20px 28px;margin-bottom:24px;">
  <div style="font-size:22px;font-weight:700;color:#00aaff;letter-spacing:0.04em;">⚙  P&ID BOM TOOL</div>
  <div style="font-size:11px;color:#4a7fa5;letter-spacing:0.12em;margin-top:4px;">
    ANH MINH 2021 &nbsp;·&nbsp; DXF → BILL OF MATERIALS &nbsp;·&nbsp; SPX / TETRA PAK
  </div>
</div>
""", unsafe_allow_html=True)

col_left, col_right = st.columns([2, 1])

with col_left:
    # API Key
    api_key = st.text_input(
        "🔑 Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Lấy tại https://console.anthropic.com — Key chỉ dùng trong session này, không lưu.",
    )

    # Skill
    skill = st.radio("📐 Chọn tiêu chuẩn P&ID:", ["SPX", "TPV (Tetra Pak)"], horizontal=True)
    skill_key = "SPX" if skill == "SPX" else "TPV"

    badge_class = "spx-badge" if skill_key == "SPX" else "tpv-badge"
    badge_text = "SPX: SV41/SV43/SV44/SV42 · W+/WS+/DW+" if skill_key == "SPX" else "Tetra Pak: LV-NC/NO · Screw/Diaphragm/Utility Pump"
    st.markdown(f'<div class="skill-badge {badge_class}">{badge_text}</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Upload
    uploaded = st.file_uploader("📁 Upload file DXF", type=["dxf"], label_visibility="collapsed")

with col_right:
    st.markdown("""
    <div style="background:#0d1f3c;border:1px solid #1e3a5f;border-radius:10px;padding:16px;font-size:12px;color:#7fb3d0;">
    <div style="color:#00aaff;font-weight:700;margin-bottom:10px;">📋 HƯỚNG DẪN</div>
    <b style="color:#e6edf3">1.</b> Nhập API Key<br>
    <b style="color:#e6edf3">2.</b> Chọn chuẩn SPX hoặc TPV<br>
    <b style="color:#e6edf3">3.</b> Upload file .dxf<br>
    <b style="color:#e6edf3">4.</b> Nhấn <b style="color:#00aaff">Phân tích & Tạo BOM</b><br>
    <b style="color:#e6edf3">5.</b> Xem kết quả & Xuất Excel<br>
    <hr style="border-color:#1e3a5f;margin:12px 0">
    <div style="color:#4a7fa5;font-size:11px;">
    File DXF cần có block INSERT entity.<br>
    App đọc tên block → AI map → BOM.<br>
    Excel xuất đúng 10 cột chuẩn.
    </div>
    </div>
    """, unsafe_allow_html=True)

# Process uploaded file
if uploaded:
    try:
        text = uploaded.read().decode("utf-8", errors="replace")
        block_count, texts, total = parse_dxf(text)

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 Tổng INSERT", total)
        c2.metric("🔷 Block types", len(block_count))
        c3.metric("📝 Text entities", len(texts))

        # Block chips
        if block_count:
            chips = " ".join(
                f'<span class="block-chip">{name}<span class="count-chip">×{cnt}</span></span>'
                for name, cnt in sorted(block_count.items(), key=lambda x: -x[1])[:50]
            )
            st.markdown(f"""
            <div style="background:#0d1f3c;border:1px solid #1e3a5f;border-radius:10px;padding:14px 16px;margin:12px 0;">
              <div style="font-size:11px;color:#4a7fa5;letter-spacing:0.1em;margin-bottom:8px;">BLOCK NAMES TÌM THẤY TRONG DXF</div>
              {chips}
              {'<span style="color:#4a7fa5;font-size:11px;margin-left:6px;">+' + str(len(block_count)-50) + ' khác</span>' if len(block_count) > 50 else ''}
            </div>
            """, unsafe_allow_html=True)

        # Generate button
        if st.button("▶  PHÂN TÍCH & TẠO BOM"):
            if not api_key:
                st.error("⚠ Vui lòng nhập Anthropic API Key")
            elif not block_count:
                st.error("⚠ Không tìm thấy INSERT entity trong file DXF")
            else:
                with st.spinner("⏳ Đang phân tích và tạo BOM..."):
                    try:
                        bom = call_claude(block_count, texts, total, uploaded.name, skill_key, api_key)
                        st.session_state["bom"] = bom
                        st.session_state["file_name"] = uploaded.name
                        st.session_state["skill"] = skill_key
                        st.success(f"✅ BOM đã tạo thành công!")
                    except json.JSONDecodeError:
                        st.error("❌ Lỗi parse JSON từ AI. Thử lại.")
                    except Exception as e:
                        st.error(f"❌ Lỗi: {e}")

    except Exception as e:
        st.error(f"❌ Không đọc được file DXF: {e}")

# Show BOM
if "bom" in st.session_state:
    bom = st.session_state["bom"]
    st.markdown("---")

    # Count items
    total_items = sum(
        len(sub.get("items", []))
        for s in bom.get("sections", [])
        for sub in s.get("subsections", [])
    )

    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.markdown(f'<span style="color:#00dc82;font-weight:700;font-size:16px;">✓ BOM đã tạo</span>'
                    f'<span style="color:#4a7fa5;font-size:13px;margin-left:12px;">{total_items} dòng thiết bị</span>',
                    unsafe_allow_html=True)
    with col_btn:
        excel_buf = create_excel(bom, st.session_state.get("skill", "SPX"))
        fn = st.session_state.get("file_name", "output").replace(".dxf", "")
        st.download_button(
            label="⬇ XUẤT EXCEL",
            data=excel_buf,
            file_name=f"BOM_{st.session_state.get('skill','SPX')}_{fn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # Render BOM table
    rows_html = ""
    for section in bom.get("sections", []):
        rows_html += f'<tr class="sec"><td>{section["stt"]}</td><td colspan="9"><b>{section["mo_ta"]}</b></td></tr>'
        for sub in section.get("subsections", []):
            rows_html += f'<tr class="sub"><td>{sub["stt"]}</td><td colspan="9">{sub["mo_ta"]}</td></tr>'
            for item in sub.get("items", []):
                rows_html += (
                    f'<tr class="dat">'
                    f'<td>{item.get("stt","")}</td>'
                    f'<td>{item.get("mo_ta","")}</td>'
                    f'<td class="tag">{item.get("chung_loai","")}</td>'
                    f'<td class="mat">{item.get("vat_lieu","")}</td>'
                    f'<td class="sz">{item.get("kt1","")}</td>'
                    f'<td class="sz">{item.get("kt2","")}</td>'
                    f'<td>{item.get("tieu_chuan","")}</td>'
                    f'<td>{item.get("xuat_xu","")}</td>'
                    f'<td>{item.get("don_vi","")}</td>'
                    f'<td class="qty">{item.get("so_luong","")}</td>'
                    f'</tr>'
                )

    st.markdown(f"""
    <div style="overflow-x:auto;border-radius:10px;border:1px solid #1e3a5f;">
    <table class="bom-table">
      <thead><tr>
        <th>STT</th><th>MÔ TẢ</th><th>CHỦNG LOẠI</th><th>VẬT LIỆU</th>
        <th>K.THƯỚC 1</th><th>K.THƯỚC 2</th><th>TIÊU CHUẨN</th><th>XUẤT XỨ</th>
        <th>ĐƠN VỊ</th><th>SỐ LƯỢNG</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div style="text-align:center;color:#2a4a6a;font-size:11px;margin-top:48px;letter-spacing:0.08em;">
ANH MINH P&ID BOM TOOL · SPX / TETRA PAK · v1.0
</div>
""", unsafe_allow_html=True)
