# P&ID BOM Generator — ANH MINH 2021

Streamlit app tự động đọc file DXF (P&ID chuẩn SPX / Tetra Pak ANH MINH)
và xuất BOM Excel đúng chuẩn FORM_BOM_STANDARD.

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy app

```bash
streamlit run app.py
```

---

## Cấu trúc file

| File | Chức năng |
|---|---|
| `app.py` | Giao diện Streamlit chính |
| `dxf_processor.py` | Orchestrator: đọc DXF → trả items |
| `line_detector.py` | Nhận diện line type từ màu sắc đường ống |
| `block_library.py` | Block name matching + tra spec theo line |
| `size_detector.py` | Proximity-based size detection (DN từ annotation) |
| `bom_exporter.py` | Xuất BOM Excel đúng chuẩn ANH MINH |

---

## Quy tắc áp dụng

### Quy tắc 1 — Nhận diện Line theo màu sắc

| Màu | Line | Nhóm |
|---|---|---|
| Magenta | CIP | Sanitary |
| Red | Product | Sanitary |
| Cyan | RO / Process Water | Sanitary |
| Blue | Compressed Air | Industry |
| ACI-30 (cam đỏ) | Steam & Condensate | Industry |
| RGB(0,166,166) | Ice Water | Industry |
| Green | Cooling Water | Industry |
| RGB(2,237,233) | City Water / Soft Water | Industry |

### Quy tắc 2 — Size < DN50 → Thread end

- Áp dụng cho thiết bị **công nghiệp** (Industry)
- Size ≥ DN50 → Flange (PN16 / JIS10K)
- Size < DN50 → Thread end
- Thiết bị **vi sinh** (Sanitary): không áp dụng — dùng Clamp/SMS/DIN

### Block name matching

- Word-order-insensitive: `MANUAL BUTTERFLY VALVE` = `BUTTERFLY VALVE MANUAL`
- Nếu không có trong thư viện → vẫn đưa vào BOM, hiển thị cảnh báo

### BOM Formatting

| Row type | Nền | Chữ | Style |
|---|---|---|---|
| Major (A/B/C) | Cyan `#00FFFF` | Đỏ | Bold |
| Sub-section (I/II/III) | Vàng `#FFFF00` | Đỏ | Bold |
| Dữ liệu thiết bị | Không tô | Đen | Không bold |
