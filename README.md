# P&ID BOM Tool – ANH MINH

Upload file DXF → AI phân tích block names → Xuất BOM Excel  
Hỗ trợ: **SPX** (SV41/SV43/SV44/SV42) và **Tetra Pak** (LV-NC/NO, Screw/Diaphragm Pump)

## Deploy lên Streamlit Community Cloud (miễn phí)

### Bước 1 – Tạo GitHub repo
1. Vào https://github.com → New repository → đặt tên `pid-bom-tool`
2. Upload 3 file: `app.py`, `requirements.txt`, `README.md`

### Bước 2 – Deploy Streamlit
1. Vào https://share.streamlit.io → Sign in bằng GitHub
2. Click **New app** → chọn repo `pid-bom-tool`
3. Main file: `app.py` → Click **Deploy!**
4. Đợi ~2 phút → nhận link dạng `https://xxx.streamlit.app`

### Bước 3 – Share cho đồng nghiệp
- Copy link Streamlit → share
- Mỗi người nhập **Anthropic API Key** của mình khi dùng
- Không cần cài đặt gì thêm

## Cấu trúc file
```
pid-bom-tool/
├── app.py            ← Main Streamlit app
├── requirements.txt  ← Dependencies
└── README.md
```

## Cách dùng
1. Nhập Anthropic API Key (lấy tại https://console.anthropic.com)
2. Chọn chuẩn: SPX hoặc TPV (Tetra Pak)
3. Upload file .dxf
4. Click "Phân tích & Tạo BOM"
5. Xem kết quả → Xuất Excel

## Tech stack
- **Streamlit** – Web UI
- **Anthropic Claude** – AI phân tích P&ID
- **openpyxl** – Xuất Excel định dạng chuẩn
