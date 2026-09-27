# BankReport Mapping Hub (TT35 & CIC)

Hệ thống quản lý tập trung Từ điển Dữ liệu, Ma trận Ánh xạ Nghiệp vụ sang Kỹ thuật và Phân tích Tác động (Data Lineage) cho các báo cáo tuân thủ ngân hàng theo Thông tư 35/2015/TT-NHNN và Quyết định 573/QĐ-NHNN (CIC).

## Tính năng chính
- **Dashboard Tổng quan**: Theo dõi độ phủ mapping và số lượng chỉ tiêu của 38 mẫu biểu TT35 & CIC.
- **Tra cứu Master-Detail**: Bố cục 2 cột chuyên nghiệp, tìm kiếm tức thì theo mã chỉ tiêu, tên tiếng Việt, cột CSDL.
- **Phân tích Tác động 2 chiều (Lineage)**: Tra cứu nhanh từ bảng/cột Core ODS sang tất cả các báo cáo và thủ tục PL/SQL bị ảnh hưởng.
- **Xuất / Nhập chuẩn ngân hàng**: Xuất Ma trận Excel (.xlsx) và Tài liệu Đặc tả kỹ thuật (.docx) tự động.

## Khởi chạy Local
`ash
pip install -r requirements.txt
python app.py
`
Truy cập: http://localhost:5050

## Triển khai Render (Cloud)
Ứng dụng tương thích hoàn toàn với nền tảng Render thông qua file cấu hình ender.yaml.