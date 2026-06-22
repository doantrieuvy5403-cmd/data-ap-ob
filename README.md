# DATA - AP - OB

Hệ thống quản lý database căn hộ cho Inspired Space, bao gồm:
- Database AP Miền Nam (AP_MN)
- Database AP Miền Bắc (AP_MB)
- Dashboard thống kê và biểu đồ
- Import/Export dữ liệu Excel/CSV

## Tính năng

### 1. Quản lý dữ liệu
- Xem danh sách căn hộ theo khu vực (Miền Nam/Miền Bắc)
- Tìm kiếm và lọc theo: trạng thái, người phụ trách, thành phố, phân loại
- Thêm mới, chỉnh sửa, xóa records
- Phân trang (50 records/trang)

### 2. Dashboard
- Biểu đồ phân bố theo trạng thái
- Biểu đồ phân bố theo thành phố (Top 10)
- Biểu đồ khối lượng công việc theo người
- Biểu đồ chất lượng cơ sở vật chất
- Biểu đồ phân loại căn hộ
- So sánh Miền Nam vs Miền Bắc theo trạng thái

### 3. Import/Export
- Import dữ liệu từ file Excel (.xlsx) hoặc CSV
- Export dữ liệu ra file Excel với tên cột tiếng Việt
- Hỗ trợ nhiều người dùng cùng quản lý dữ liệu

## Cài đặt

### Yêu cầu
- Python 3.8+
- pip

### Các bước cài đặt

1. **Clone hoặc tải dự án về máy**
```bash
cd "/Users/qweasdzxcbm/Downloads/DATA - AP - OB"
```

2. **Tạo môi trường ảo (khuyến nghị)**
```bash
python3 -m venv venv
source venv/bin/activate  # Trên macOS/Linux
# hoặc
venv\Scripts\activate  # Trên Windows
```

3. **Cài đặt dependencies**
```bash
pip install -r requirements.txt
```

4. **Import dữ liệu từ Excel vào database**
```bash
python seed.py
```

Script này sẽ đọc file `Database_Inspired Space.xlsx` và import:
- Sheet `Databse AP_MN` → ~1852 records
- Sheet `Databse AP_MB` → ~731 records

5. **Chạy ứng dụng**
```bash
python app.py
```

Ứng dụng sẽ chạy tại: http://localhost:5000

## Cấu trúc dự án

```
DATA - AP - OB/
├── app.py                  # Flask application chính
├── models.py              # Database models (SQLAlchemy)
├── seed.py                # Script import dữ liệu từ Excel
├── requirements.txt       # Python dependencies
├── README.md             # Tài liệu này
├── instance/
│   └── database.db       # SQLite database (tự động tạo)
├── templates/            # HTML templates
│   ├── base.html         # Layout chính
│   ├── index.html        # Trang tổng quan
│   ├── database.html     # Trang xem dữ liệu
│   ├── add_edit.html     # Form thêm/sửa
│   ├── dashboard.html    # Dashboard với biểu đồ
│   └── import_export.html # Trang import/export
└── static/
    └── style.css         # Custom CSS
```

## Sử dụng

### Xem dữ liệu
- Truy cập trang chủ → Chọn "AP Miền Nam" hoặc "AP Miền Bắc"
- Sử dụng bộ lọc để tìm kiếm theo trạng thái, người phụ trách, thành phố, phân loại
- Click vào nút "Sửa" để chỉnh sửa hoặc "Xóa" để xóa record

### Thêm mới
- Click nút "+ Thêm mới" ở góc phải trên cùng
- Điền thông tin vào form
- Click "Thêm mới" để lưu

### Dashboard
- Truy cập menu "Dashboard" ở sidebar
- Xem 6 biểu đồ thống kê khác nhau
- Dữ liệu được cập nhật real-time từ database

### Import/Export
- **Import**: Chọn khu vực → Chọn file Excel/CSV → Click "Import dữ liệu"
- **Export**: Click nút "Export AP Miền Nam" hoặc "Export AP Miền Bắc" để tải file Excel

## API Endpoints

| Method | Endpoint | Mô tả |
|--------|----------|-------|
| GET | `/` | Trang chủ |
| GET | `/database/<region>` | Xem dữ liệu (mn/mb) |
| GET/POST | `/database/<region>/add` | Thêm mới record |
| GET/POST | `/record/<id>/edit` | Chỉnh sửa record |
| POST | `/record/<id>/delete` | Xóa record |
| GET | `/dashboard` | Dashboard |
| GET | `/api/stats` | API lấy thống kê (JSON) |
| GET | `/export/<region>` | Export dữ liệu Excel |
| POST | `/import/<region>` | Import dữ liệu từ file |
| GET | `/import-export` | Trang import/export |

## Công nghệ sử dụng

- **Backend**: Flask 3.0.0
- **Database**: SQLite + Flask-SQLAlchemy
- **Frontend**: Jinja2 templates + Tailwind CSS
- **Charts**: Chart.js 4.4.1
- **Data Processing**: pandas + openpyxl

## Deploy

### Deploy lên GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <your-github-repo-url>
git push -u origin main
```

### Deploy lên production (ví dụ: Heroku)
```bash
# Cài đặt Heroku CLI
heroku create data-ap-ob
git push heroku main
heroku open
```

## License

MIT License - Tự do sử dụng và chỉnh sửa

## Tác giả

Inspired Space Team
