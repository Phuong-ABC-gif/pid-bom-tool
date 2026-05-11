# P&ID BOM Extractor — ANH MINH 2021
## SPX / Tetra Pak — Rule-based, No AI Cost

---

## 🚀 Deploy lên Streamlit Cloud

### Bước 1 — Tạo repository GitHub
1. Tạo repo mới trên GitHub (public hoặc private)
2. Upload 2 file vào repo:
   - `app.py`
   - `requirements.txt`

### Bước 2 — Đăng nhập Streamlit Cloud
- Vào https://share.streamlit.io
- Đăng nhập bằng tài khoản GitHub

### Bước 3 — Deploy
- Click **"New app"**
- Chọn repo vừa tạo
- Main file path: `app.py`
- Click **"Deploy!"**

---

## 🖥️ Chạy local (test trước khi deploy)

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## ⚙️ Cách hoạt động

### Không dùng AI, không tốn cost:
1. **Đọc DXF** → extract tất cả `INSERT entity` (block name + tọa độ)
2. **Nhận diện thiết bị** → so khớp tên block với thư viện (word-order-insensitive)
3. **Xác định size** → tìm chú thích `DN` gần thiết bị nhất (proximity detection, ngưỡng 300 đơn vị)
4. **Xuất BOM** → file Excel chuẩn ANHMINH (FORM_BOM_STANDARD)

### Quy tắc so khớp block name:
- Không phân biệt thứ tự từ: `BUTTERFLY VALVE MANUAL` = `MANUAL BUTTERFLY VALVE`
- Không phân biệt dấu `-` và `_`: `SSV-210` = `SSV 210`
- Không phân biệt hoa/thường

### Ngưỡng proximity:
- Mặc định: 300 đơn vị DXF
- Size = annotation DN gần thiết bị nhất
- Nếu không tìm thấy → ghi `?` và cảnh báo

---

## 📋 Thư viện thiết bị

### SPX Type:
- SSV 200/210/220/300 (SV41/43/44/42) — Manual/Actuator/Thinktop
- Mixproof Valve — Actuator/Thinktop
- Butterfly Valve — Manual/Actuator/Thinktop
- Centrifugal Pump W+/WS+, Lobe Pump DW+
- NRV, SG, SVA, SVM, CPV, SF, FSH, FSR
- Utility Valves: GV, BV, BFV-U, DV, ASV, CV, PRV, ARV, VRV, RPV
- Instruments: FM, TT, TI, PT, PI, LS, CS, FS

### Tetra Pak Type (SPX + thêm):
- Leakage Valve NC/NO Actuator (**TP đặc thù**)
- Screw Pump, Utility Pump, Diaphragm Pump (**TP đặc thù**)
- Aseptic Sampling Valve, Sanitary Sampling Valve (**TP đặc thù**)
- Bag Filter, Inline Filter, Angle Filter (**TP đặc thù**)

---

## 📂 Cấu trúc file output BOM (Excel)

```
Rows 1-6:   Header công ty ANHMINH
Rows 7-11:  Ngày, BG số, Dự án
Row 12:     BÁO GIÁ / QUOTATION
Rows 13-15: Phạm vi công việc
Rows 16-18: Header cột (STT, MÔ TẢ, CHỦNG LOẠI, VẬT LIỆU, K.THƯỚC 1/2, TIÊU CHUẨN, XUẤT XỨ, ĐƠN VỊ, SỐ LƯỢNG)
Rows 19+:   Dữ liệu BOM (phân nhóm theo đường)
Footer:     TOTAL / VAT / GRAND TOTAL + Điều kiện bán hàng
```

**Màu sắc chuẩn ANHMINH:**
- Nền `#00FFFF` (cyan) + chữ `#FF0000` (đỏ) → Section header
- Nền `#FFFF00` (vàng) + chữ `#FF0000` (đỏ) → Item data

---

## 🔧 Thêm thiết bị vào thư viện

Mở `app.py`, tìm dict `SPX_LIBRARY` hoặc `TPK_OVERRIDES` và thêm entry:

```python
"tên block lowercase": {
    "mo_ta":      "Mô tả đầy đủ hiển thị trong BOM",
    "chung_loai": "Manual / Actuator NC / Thinktop / -",
    "vat_lieu":   "SS316L / Cast Iron / Brass / ...",
    "tieu_chuan": "SMS / JIS10K / PN16 / Thread end",
    "don_vi":     "pcs / set / m",
    "section":    "sanitary / ice_water / steam / air / instrument / equipment",
},
```
