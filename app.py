"""
app.py — Streamlit P&ID BOM Generator
Hỗ trợ: SPX Type & Tetra Pak Type (ANHMINH 2021)
Quy trình:
  1. Upload DXF → đọc block linename TP-left → xác định line type
  2. Đọc block name thiết bị (word-order-insensitive)
  3. Proximity detect size → áp spec
  4. Size < DN50 → Thread end (công nghiệp)
  5. Xuất BOM Excel chuẩn ANHMINH
"""

import io
import tempfile
import os

import streamlit as st
import pandas as pd

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="P&ID BOM Generator – ANHMINH",
    page_icon="📐",
    layout="wide",
)

# ── CSS tối giản ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
thead tr th { background-color: #00FFFF !important; color: #CC0000 !important; font-weight: bold; }
.stAlert { font-size: 0.9rem; }
.block-label { color: #666; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)

# ── Title ──────────────────────────────────────────────────────────────────────
st.title("📐 P&ID BOM Generator")
st.caption("SPX Type & Tetra Pak Type | Chuẩn ANHMINH 2021")

# ── Sidebar: project info ──────────────────────────────────────────────────────
with st.sidebar:
    st.header("Thông tin dự án")
    project_name = st.text_input("Tên dự án", placeholder="VD: NHÀ MÁY SỮA ABC")
    doc_no       = st.text_input("Số BG", placeholder="VD: BG-2024-001")
    pid_type     = st.selectbox("Loại P&ID", ["Auto (SPX + Tetra Pak)", "SPX Only", "Tetra Pak Only"])
    max_dist     = st.slider("Ngưỡng proximity size (DXF units)", 100, 1000, 500, 50)
    st.markdown("---")
    st.markdown("""
**Quy tắc áp spec:**
- Xác định line từ block `linename TP-left`
- Size < DN50 → **Thread end** (công nghiệp)
- Vi sinh: **SS316L / SMS**
- Water: **Brass (ren) / Cast Iron JIS10K (bích)**
- Steam: **Cast Iron PN16 (bích) / Thread end (ren)**
    """)

# ── Main: Upload ───────────────────────────────────────────────────────────────
st.subheader("1. Upload file DXF")
uploaded = st.file_uploader("Chọn file .dxf", type=["dxf"], accept_multiple_files=False)

if uploaded is None:
    st.info("⬆️ Upload file DXF P&ID để bắt đầu.")
    st.stop()

# ── Process DXF ───────────────────────────────────────────────────────────────
with st.spinner("Đang đọc file DXF..."):
    # Ghi tạm ra disk vì ezdxf cần filepath
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        # Import ở đây để lỗi import hiện rõ ràng
        from dxf_reader import read_dxf, aggregate_bom
        from bom_generator import generate_bom
        from pid_spec import get_line_category

        result = read_dxf(tmp_path)
    except ImportError as e:
        st.error(f"Thiếu thư viện: {e}. Chạy `pip install -r requirements.txt`")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi đọc DXF: {e}")
        st.stop()
    finally:
        os.unlink(tmp_path)

# ── Lines detected ─────────────────────────────────────────────────────────────
st.subheader("2. Lines phát hiện")
if not result.lines:
    st.warning("Không tìm thấy block `linename TP-left`. Tất cả thiết bị sẽ được gán là UNKNOWN.")
else:
    df_lines = pd.DataFrame([
        {
            "Line ID": l.line_id,
            "Loại đường": l.line_type,
            "Nhóm BOM": get_line_category(l.line_type),
            "X": round(l.x, 1),
            "Y": round(l.y, 1),
        }
        for l in result.lines
    ])
    # Màu theo loại
    def color_line(val):
        c = {
            "SANITARY":  "#d4edda",
            "ICE_WATER": "#cce5ff",
            "COOLING":   "#d1ecf1",
            "STEAM":     "#fff3cd",
            "AIR":       "#e2e3e5",
        }
        return f"background-color: {c.get(val, '#fff')}"

    st.dataframe(
        df_lines.style.applymap(color_line, subset=["Loại đường"]),
        use_container_width=True, height=200
    )

# ── Equipment detected ─────────────────────────────────────────────────────────
st.subheader("3. Thiết bị phát hiện")
if not result.equipment:
    st.warning("Không tìm thấy thiết bị nào trong file DXF.")
    st.stop()

# Warnings
if result.warnings:
    with st.expander(f"⚠️ {len(result.warnings)} cảnh báo cần kiểm tra"):
        for w in result.warnings:
            st.warning(w)

df_eq = pd.DataFrame([
    {
        "Block Name":  eq.block_name,
        "Size":        eq.size,
        "Line":        eq.line_id or "—",
        "Loại đường":  eq.line_type,
        "Vật liệu":    eq.spec.get("vat_lieu", ""),
        "Tiêu chuẩn":  eq.spec.get("tieu_chuan", ""),
        "Kết nối":     eq.conn_type,
        "X": round(eq.x, 0),
        "Y": round(eq.y, 0),
    }
    for eq in result.equipment
])

# Filter
col1, col2 = st.columns(2)
with col1:
    lt_filter = st.multiselect(
        "Lọc theo loại đường",
        options=df_eq["Loại đường"].unique().tolist(),
        default=df_eq["Loại đường"].unique().tolist(),
    )
with col2:
    search = st.text_input("Tìm block name", placeholder="SSV, Globe, Ball...")

df_show = df_eq[df_eq["Loại đường"].isin(lt_filter)]
if search:
    df_show = df_show[df_show["Block Name"].str.contains(search, case=False, na=False)]

st.dataframe(df_show, use_container_width=True, height=350)
st.caption(f"Tổng: **{len(df_show)}** thiết bị (sau lọc) / {len(df_eq)} tổng cộng")

# ── BOM aggregated ─────────────────────────────────────────────────────────────
st.subheader("4. BOM tổng hợp")
bom_rows = aggregate_bom(result.equipment)

if not bom_rows:
    st.warning("Không có dữ liệu BOM.")
    st.stop()

df_bom = pd.DataFrame([
    {
        "Loại đường":  r.get("line_type", ""),
        "MÔ TẢ":       r.get("block_name", ""),
        "CHỦNG LOẠI":  r.get("chung_loai", ""),
        "VẬT LIỆU":    r.get("vat_lieu", ""),
        "K.THƯỚC":     r.get("size", "?"),
        "TIÊU CHUẨN":  r.get("tieu_chuan", ""),
        "Kết nối":     r.get("conn_type", ""),
        "ĐƠN VỊ":      r.get("don_vi", "pcs"),
        "SL":          r.get("sl", 1),
    }
    for r in bom_rows
])

def highlight_bom_row(row):
    lt = row["Loại đường"]
    colors = {
        "SANITARY":  "#d4edda",
        "ICE_WATER": "#cce5ff",
        "COOLING":   "#d1ecf1",
        "STEAM":     "#fff3cd",
        "AIR":       "#e2e3e5",
    }
    bg = colors.get(lt, "#fff")
    return [f"background-color: {bg}"] * len(row)

st.dataframe(
    df_bom.style.apply(highlight_bom_row, axis=1),
    use_container_width=True,
    height=400
)

# Legend
st.markdown("""
<div style='font-size:0.8rem; display:flex; gap:16px; flex-wrap:wrap; margin-top:4px'>
  <span style='background:#d4edda; padding:2px 8px; border-radius:4px'>■ SANITARY</span>
  <span style='background:#cce5ff; padding:2px 8px; border-radius:4px'>■ ICE WATER</span>
  <span style='background:#d1ecf1; padding:2px 8px; border-radius:4px'>■ COOLING</span>
  <span style='background:#fff3cd; padding:2px 8px; border-radius:4px'>■ STEAM</span>
  <span style='background:#e2e3e5; padding:2px 8px; border-radius:4px'>■ AIR</span>
</div>
""", unsafe_allow_html=True)

# ── Export BOM ─────────────────────────────────────────────────────────────────
st.subheader("5. Xuất BOM Excel")

if st.button("🗂️ Tạo file BOM (.xlsx)", type="primary"):
    with st.spinner("Đang tạo BOM Excel..."):
        xlsx_bytes = generate_bom(
            bom_rows=bom_rows,
            project=project_name,
            doc_no=doc_no,
            warnings=result.warnings,
        )
    fname = f"BOM_{project_name or 'PID'}_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="⬇️ Tải BOM Excel",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.success(f"✅ BOM đã tạo: {len(bom_rows)} loại thiết bị | {sum(r.get('sl',1) for r in bom_rows)} tổng số lượng")

    # Post-export verification
    st.markdown("**Kiểm tra BOM:**")
    equip_total = len(result.equipment)
    bom_total   = sum(r.get("sl", 1) for r in bom_rows)
    if equip_total == bom_total:
        st.success(f"✅ PASS — {equip_total} thiết bị DXF = {bom_total} tổng SL trong BOM")
    else:
        st.error(f"❌ MISMATCH — DXF: {equip_total} EA | BOM: {bom_total} EA — Kiểm tra lại!")

    unk_size = [r for r in bom_rows if r.get("size") == "?"]
    if unk_size:
        st.warning(f"⚠️ {len(unk_size)} dòng chưa xác định size — cần kiểm tra bản vẽ.")
    unk_line = [r for r in bom_rows if r.get("line_type") == "UNKNOWN"]
    if unk_line:
        st.warning(f"⚠️ {len(unk_line)} thiết bị chưa xác định line type.")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("P&ID BOM Generator · ANHMINH 2021 · SPX & Tetra Pak")
