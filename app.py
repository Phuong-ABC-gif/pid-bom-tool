"""
app.py — PID BOM Generator (ANH MINH 2021)
FIX: items lưu session_state → click row hiện tọa độ đúng. Bỏ tiêu đề lặp.
"""
import os
import tempfile
import pandas as pd
import streamlit as st

from dxf_processor import process_dxf
from bom_exporter import export_bom

st.set_page_config(page_title="P&ID BOM Generator — ANH MINH", page_icon="⚙️", layout="wide")

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ P&ID BOM Generator")
    st.caption("ANH MINH 2021 — SPX & Tetra Pak")
    st.divider()
    st.markdown("### 🎨 Line theo màu sắc")
    for icon, line, grp in [
        ("🟣 Magenta",        "CIP",                "Sanitary"),
        ("🔴 Red",            "Product",            "Sanitary"),
        ("🩵 Cyan",           "RO / Process Water", "Sanitary"),
        ("🔵 Blue",           "Compressed Air",     "Industry"),
        ("🟠 ACI-30",         "Steam & Condensate", "Industry"),
        ("🟦 RGB(0,166,166)", "Ice Water",          "Industry"),
        ("🟢 Green",          "Cooling Water",      "Industry"),
        ("🩶 RGB(2,237,233)", "City / Soft Water",  "Industry"),
    ]:
        st.markdown(f"{icon} **{line}** {'🧬' if grp=='Sanitary' else '🏭'}")
    st.divider()
    st.markdown("### 📏 Quy tắc size")
    st.info("Size **< DN50** → **Thread end**\n\nSize **≥ DN50** → **Flange**")

# ── Title ──────────────────────────────────────────────────────────────────
st.title("⚙️ P&ID BOM Generator")
st.markdown("Upload **DXF** → nhận diện thiết bị, line, size → xuất **BOM Excel** chuẩn ANH MINH.")

# ── Upload ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📁 Chọn file DXF", type=["dxf"])

# Reset session khi đổi file
if uploaded:
    fid = f"{uploaded.name}_{uploaded.size}"
    if st.session_state.get("_fid") != fid:
        st.session_state["_fid"] = fid
        st.session_state.pop("items", None)
        st.session_state.pop("warnings", None)

# ── Nút xử lý ─────────────────────────────────────────────────────────────
if uploaded:
    st.success(f"✅ **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")
    c1, _ = st.columns([1, 3])
    with c1:
        run_btn = st.button("🚀 Xử lý & Xuất BOM", type="primary", use_container_width=True)

    # PHASE 1 — chỉ chạy khi bấm nút
    if run_btn:
        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        _s = st.empty()
        _p = st.progress(0.0)

        def _cb(msg, pct):
            _s.markdown(f"⏳ {msg}")
            _p.progress(min(pct, 1.0))

        try:
            items, warnings = process_dxf(tmp_path, _cb)
            st.session_state["items"]    = items
            st.session_state["warnings"] = warnings
        except Exception as err:
            st.error(f"❌ {err}")
            os.unlink(tmp_path)
            st.stop()
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

        _p.empty(); _s.empty()

# PHASE 2 — render mỗi lần rerun (kể cả khi click row trong bảng)
if "items" in st.session_state:
    items    = st.session_state["items"]
    warnings = st.session_state.get("warnings", [])

    if not items:
        st.warning("⚠️ Không tìm thấy INSERT block nào.")
        st.stop()

    if warnings:
        with st.expander(f"⚠️ {len(warnings)} cảnh báo", expanded=False):
            for w in warnings:
                st.markdown(f"- {w}")

    # Build DataFrame (1 lần)
    df = (
        pd.DataFrame(items)
        [["block_name","display_name","line_label","group",
          "size","chung_loai","vat_lieu","tieu_chuan","don_vi","x","y"]]
        .rename(columns={
            "block_name":  "Block Name",
            "display_name":"Tên thiết bị",
            "line_label":  "Line",
            "group":       "Nhóm",
            "size":        "Size",
            "chung_loai":  "Chủng loại",
            "vat_lieu":    "Vật liệu",
            "tieu_chuan":  "Tiêu chuẩn",
            "don_vi":      "Đơn vị",
            "x":           "X",
            "y":           "Y",
        })
    )

    # Tiêu đề — 1 dòng duy nhất
    st.markdown(f"### 📋 Kết quả: **{len(items)} thiết bị** tìm thấy")
    st.caption("💡 Click vào một dòng để xem tọa độ trên bản vẽ")

    # Placeholder tọa độ khai báo TRƯỚC bảng → hiện phía trên sau rerun
    coord_box = st.empty()

    # Bảng preview — on_select="rerun" trigger rerun khi click row
    sel = st.dataframe(
        df,
        use_container_width=True,
        height=440,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "Block Name":   st.column_config.TextColumn("Block Name",   width="medium"),
            "Tên thiết bị": st.column_config.TextColumn("Tên thiết bị", width="large"),
            "Line":         st.column_config.TextColumn("Line",         width="medium"),
            "Nhóm":         st.column_config.TextColumn("Nhóm",         width="small"),
            "Size":         st.column_config.TextColumn("Size",         width="small"),
            "Chủng loại":   st.column_config.TextColumn("Chủng loại",   width="medium"),
            "Vật liệu":     st.column_config.TextColumn("Vật liệu",     width="medium"),
            "Tiêu chuẩn":   st.column_config.TextColumn("Tiêu chuẩn",   width="small"),
            "Đơn vị":       st.column_config.TextColumn("Đơn vị",       width="small"),
            "X": st.column_config.NumberColumn("X (UCS)", format="%.3f", width="small"),
            "Y": st.column_config.NumberColumn("Y (UCS)", format="%.3f", width="small"),
        },
    )

    # Đọc selection → điền coord_box (placeholder đã khai báo trước bảng)
    rows = sel.selection.rows if (sel and sel.selection) else []
    if rows:
        r = df.iloc[rows[0]]
        coord_box.info(
            f"📍 &nbsp;**{r['Tên thiết bị']}** &nbsp;·&nbsp; "
            f"Block: `{r['Block Name']}` &nbsp;·&nbsp; "
            f"Line: {r['Line']} &nbsp;·&nbsp; Size: **{r['Size']}**\n\n"
            f"**Tọa độ UCS:** &nbsp;&nbsp;"
            f"X = `{r['X']:.3f}` &nbsp;&nbsp;|&nbsp;&nbsp; Y = `{r['Y']:.3f}`"
        )
    else:
        coord_box.empty()

    # Thống kê
    with st.expander("📊 Thống kê theo Line"):
        stats = (df.groupby(["Line","Nhóm"]).size()
                   .reset_index(name="Số lượng")
                   .sort_values("Số lượng", ascending=False))
        st.dataframe(stats, use_container_width=True, hide_index=True)

    # Xuất BOM
    st.divider()
    st.markdown("### 📥 Tải BOM Excel")
    try:
        bom_bytes = export_bom(items)
        fname = (uploaded.name if uploaded else "output").replace(".dxf","_BOM.xlsx")
        st.download_button(
            "⬇️ Tải BOM Excel (.xlsx)", bom_bytes, fname,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.success("✅ BOM sẵn sàng.")
    except Exception as err:
        st.error(f"❌ {err}")

    # Edit thủ công
    st.divider()
    with st.expander("✏️ Chỉnh sửa thủ công trước khi xuất lại"):
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        if st.button("🔄 Xuất lại BOM"):
            ei = []
            for _, row in edited.iterrows():
                ei.append({
                    "block_name":   row.get("Block Name",""),
                    "display_name": row.get("Tên thiết bị",""),
                    "line_type":    str(row.get("Line","Unknown")).split("(")[0].strip(),
                    "line_label":   row.get("Line","Unknown"),
                    "group":        row.get("Nhóm","unknown"),
                    "size":         row.get("Size","?"),
                    "chung_loai":   row.get("Chủng loại","-"),
                    "vat_lieu":     row.get("Vật liệu","-"),
                    "tieu_chuan":   row.get("Tiêu chuẩn","-"),
                    "don_vi":       row.get("Đơn vị","pcs"),
                })
            try:
                nb = export_bom(ei)
                efname = (uploaded.name if uploaded else "output").replace(".dxf","_BOM_edited.xlsx")
                st.download_button("⬇️ Tải BOM đã chỉnh sửa", nb, efname,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as err:
                st.error(f"❌ {err}")

else:
    st.info("👆 Upload file DXF rồi bấm **Xử lý & Xuất BOM** để bắt đầu.")
    st.markdown("""
**Quy trình:**
1. Upload file **DXF** từ AutoCAD P&ID
2. App tự động nhận diện line (màu), block name, size, áp spec
3. Click dòng trong bảng → xem **tọa độ UCS**
4. Tải **BOM Excel** chuẩn ANH MINH
""")
