"""
app.py — PID BOM Generator (ANH MINH 2021)
Streamlit app: upload DXF → xuất BOM Excel
"""
import os
import tempfile
import pandas as pd
import streamlit as st

from dxf_processor import process_dxf
from bom_exporter import export_bom

# ── Config ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="P&ID BOM Generator — ANH MINH",
    page_icon="⚙️",
    layout="wide",
)

# ── Sidebar: legend màu ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ P&ID BOM Generator")
    st.caption("ANH MINH 2021 — SPX & Tetra Pak")
    st.divider()
    st.markdown("### 🎨 Nhận diện Line theo màu")
    legend = [
        ("🟣 Magenta",          "CIP",                  "Sanitary"),
        ("🔴 Red",              "Product",              "Sanitary"),
        ("🩵 Cyan",             "RO / Process Water",   "Sanitary"),
        ("🔵 Blue",             "Compressed Air",       "Industry"),
        ("🟠 ACI-30",           "Steam & Condensate",   "Industry"),
        ("🟦 RGB(0,166,166)",   "Ice Water",            "Industry"),
        ("🟢 Green",            "Cooling Water",        "Industry"),
        ("🩶 RGB(2,237,233)",   "City / Soft Water",    "Industry"),
    ]
    for icon, line, grp in legend:
        badge = "🧬" if grp == "Sanitary" else "🏭"
        st.markdown(f"{icon} **{line}** {badge}")
    st.divider()
    st.markdown("### 📏 Quy tắc size")
    st.info("Size **< DN50** → **Thread end** (thiết bị công nghiệp)\n\n"
            "Size **≥ DN50** → **Flange** (PN16 / JIS10K)")
    st.markdown("### 📛 Block name")
    st.info("Đọc theo tên block trong DXF.\n"
            "Block không có trong thư viện vẫn được đưa vào BOM với cảnh báo.")


# ── Main ───────────────────────────────────────────────────────────────────
st.title("⚙️ P&ID BOM Generator")
st.markdown("Upload file **DXF** → tự động nhận diện thiết bị, line type, size → xuất **BOM Excel** chuẩn ANH MINH.")

uploaded = st.file_uploader("📁 Chọn file DXF", type=["dxf"])

if uploaded:
    st.success(f"✅ Đã nhận file: **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")

    col1, col2 = st.columns([1, 3])
    with col1:
        run_btn = st.button("🚀 Xử lý & Xuất BOM", type="primary", use_container_width=True)

    if run_btn:
        # Lưu file tạm
        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        status_text = st.empty()
        progress_bar = st.progress(0.0)

        def progress_cb(msg: str, pct: float):
            status_text.markdown(f"⏳ {msg}")
            progress_bar.progress(min(pct, 1.0))

        try:
            items, warnings = process_dxf(tmp_path, progress_cb)
        except Exception as err:
            st.error(f"❌ Lỗi khi đọc DXF: {err}")
            os.unlink(tmp_path)
            st.stop()
        finally:
            os.unlink(tmp_path)

        progress_bar.empty()
        status_text.empty()

        if not items:
            st.warning("⚠️ Không tìm thấy INSERT block nào trong file DXF.")
            st.stop()

        # ── Cảnh báo ──────────────────────────────────────────────────────
        if warnings:
            with st.expander(f"⚠️ {len(warnings)} cảnh báo (bấm để xem)", expanded=False):
                for w in warnings:
                    st.markdown(f"- {w}")

        # ── Bảng preview ──────────────────────────────────────────────────
        st.markdown(f"### 📋 Kết quả: **{len(items)} thiết bị** tìm thấy")

        df = pd.DataFrame(items)[
            ["block_name", "display_name", "line_label", "group",
             "size", "chung_loai", "vat_lieu", "tieu_chuan", "don_vi"]
        ].rename(columns={
            "block_name":   "Block Name",
            "display_name": "Tên thiết bị",
            "line_label":   "Line",
            "group":        "Nhóm",
            "size":         "Size",
            "chung_loai":   "Chủng loại",
            "vat_lieu":     "Vật liệu",
            "tieu_chuan":   "Tiêu chuẩn",
            "don_vi":       "Đơn vị",
        })

        # Tô màu theo nhóm
        def color_row(row):
            if row["Nhóm"] == "sanitary":
                return ["background-color: #e8f8e8"] * len(row)
            elif row["Nhóm"] == "industry":
                return ["background-color: #e8eeff"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(color_row, axis=1),
            use_container_width=True,
            height=420,
        )

        # ── Thống kê theo Line ────────────────────────────────────────────
        with st.expander("📊 Thống kê theo Line"):
            stats = df.groupby(["Line", "Nhóm"]).size().reset_index(name="Số lượng")
            st.dataframe(stats, use_container_width=True)

        # ── Xuất BOM Excel ────────────────────────────────────────────────
        st.divider()
        st.markdown("### 📥 Tải BOM Excel")

        try:
            bom_bytes = export_bom(items)
            fname = uploaded.name.replace(".dxf", "_BOM.xlsx")
            st.download_button(
                label="⬇️ Tải BOM Excel (.xlsx)",
                data=bom_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            st.success("✅ BOM sẵn sàng. Bấm nút trên để tải về.")
        except Exception as err:
            st.error(f"❌ Lỗi khi xuất BOM: {err}")

        # ── Edit thủ công ─────────────────────────────────────────────────
        st.divider()
        with st.expander("✏️ Chỉnh sửa thủ công trước khi xuất lại"):
            st.caption("Chỉnh sửa bảng bên dưới, sau đó bấm 'Xuất lại BOM' để tạo file mới.")
            edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

            if st.button("🔄 Xuất lại BOM từ bảng đã chỉnh sửa"):
                # Chuyển đổi df ngược lại thành items
                edited_items = []
                for _, row in edited_df.iterrows():
                    edited_items.append({
                        "block_name":   row.get("Block Name", ""),
                        "display_name": row.get("Tên thiết bị", ""),
                        "line_type":    row.get("Line", "Unknown").split("(")[0].strip(),
                        "line_label":   row.get("Line", "Unknown"),
                        "group":        row.get("Nhóm", "unknown"),
                        "size":         row.get("Size", "?"),
                        "chung_loai":   row.get("Chủng loại", "-"),
                        "vat_lieu":     row.get("Vật liệu", "-"),
                        "tieu_chuan":   row.get("Tiêu chuẩn", "-"),
                        "don_vi":       row.get("Đơn vị", "pcs"),
                    })
                try:
                    new_bom = export_bom(edited_items)
                    st.download_button(
                        label="⬇️ Tải BOM đã chỉnh sửa",
                        data=new_bom,
                        file_name=uploaded.name.replace(".dxf", "_BOM_edited.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as err:
                    st.error(f"❌ {err}")

else:
    # Màn hình chào
    st.info("👆 Upload file DXF để bắt đầu xử lý.")
    st.markdown("""
    **Quy trình:**
    1. Upload file **DXF** từ AutoCAD P&ID
    2. App tự động:
       - Nhận diện line type theo **màu sắc đường ống**
       - Đọc tên thiết bị từ **block name**
       - Xác định **size** từ annotation gần nhất
       - Áp **spec** đúng theo line (Sanitary / Industry)
       - Áp quy tắc **size < DN50 → Thread end**
    3. Preview bảng kết quả
    4. Tải file **BOM Excel** chuẩn ANH MINH
    """)
