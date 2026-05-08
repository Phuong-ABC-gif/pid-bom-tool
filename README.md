# P&ID BOM Extractor — SPX & Tetra Pak

## Mô tả
App Streamlit đọc file DXF P&ID, tra cứu block name trực tiếp theo database nội bộ (không dùng AI, không tốn cost), xác định size theo proximity detection, xuất BOM dạng Excel chuẩn ANH MINH 2021.

## Cài đặt

```bash
# 1. Tạo môi trường ảo (khuyến nghị)
python -m venv venv
source venv/bin/activate          # Linux/Mac
# hoặc: venv\Scripts\activate     # Windows

# 2. Cài thư viện
pip install -r requirements.txt

# 3. Chạy app
streamlit run app.py
```

App sẽ mở trên trình duyệt tại `http://localhost:8501`

---

## Cấu trúc file

```
pid_bom_app/
├── app.py              ← Streamlit app chính
├── pid_database.py     ← Database block name → thiết bị (SPX & Tetra Pak)
├── requirements.txt    ← Thư viện cần cài
└── README.md           ← Hướng dẫn này
```

---

## Quy trình hoạt động

1. **Upload DXF** — đọc tất cả INSERT entities trong model space
2. **Lọc block** — bỏ qua Flow Arrow, LineNameTPLeft, PID, border, v.v.
3. **Tra cứu block name** — so sánh với database theo regex (ưu tiên match cụ thể nhất)
4. **Proximity detection** — với mỗi thiết bị, tìm annotation DN/inch gần nhất trong bán kính max_dist (mặc định 300 đơn vị DXF)
5. **Build BOM** — nhóm theo Section I (Cơ khí) / Section II (Instrument), phân theo đường Process/Chiller/Cooling/Steam
6. **Verification** — tự kiểm tra tổng số lượng, size unknown
7. **Export Excel** — màu sắc chuẩn ANH MINH (header cyan, section xanh, data trắng)

---

## Thêm block name mới

Mở `pid_database.py`, thêm vào `_SPX_RAW` hoặc `_TPV_RAW`:

```python
(r'MY_BLOCK_PATTERN',
 _eq("Tên thiết bị đầy đủ", "MÃ", "Vật liệu", "Tiêu chuẩn", "EA", "sanitary", "process")),
```

**Section values:** `sanitary` | `utility` | `instrument`  
**Subsection values:** `process` | `chiller` | `cooling` | `steam` | `""`

---

## Màu sắc BOM Excel

| Loại dòng | Màu nền |
|---|---|
| Column header | #00CCFF (xanh dương nhạt) |
| Section I / II | #00FFFF (cyan) |
| Sub-header 1/2/3 | #E0FFFF (trắng xanh nhạt) |
| Data row | #FFFFFF (trắng) |
| Data row (size ?) | #FFF3CD (vàng nhạt — cảnh báo) |

---

## Skill được tích hợp

- **SPX (ANH MINH 2021):** SSV SV41/43/44/42, MPV, BFV, SVA/SVM, CPV, CP-W+/WS+, LP-DW+, GV/BV/DV/ASV...
- **Tetra Pak (ANH MINH 2021):** SSV 200/210/220/300, LV-NC/NO, ASV, SSV-samp, CP, LP, SP, UP, DP, BF, AF, LV...
