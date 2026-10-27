# HỆ THỐNG THAM VẤN & BÁO CÁO HCVF & FSC 2026 — QPFL

> **CÔNG TY TNHH TRỒNG RỪNG QUY NHƠN (QPFL)**  
> **Dự án Đánh giá Rừng có Giá trị Bảo tồn Cao (HCVF) & Khắc phục Major CAR 03 (FSC 2026)**  
> **Tác giả:** Lộc Vũ Trung

---

## 🌐 ĐƯỜNG DẪN TRUY CẬP TRỰC TUYẾN

| Hạng mục | Đường dẫn truy cập | Mô tả & Chức năng |
| :--- | :--- | :--- |
| 📝 **Phiếu Tham Vấn Trực Tuyến** | [https://qpflfsc.github.io/HCVF2026/](https://qpflfsc.github.io/HCVF2026/) | Dành cho các bên liên quan, chính quyền xã, kiểm lâm và cộng đồng điền phiếu tham vấn. |
| 📊 **Dashboard Báo Cáo Tổng Hợp** | [https://qpflfsc.github.io/HCVF2026/report.html](https://qpflfsc.github.io/HCVF2026/report.html) | Báo cáo trực quan (Chart.js), phân tích tỷ lệ đồng thuận, thống kê theo xã và phân loại bên liên quan. |
| 📥 **Tải File Master Excel** | [Tải file QPFL_Tong_Hop_Ket_Qua_Tham_Van_2026.xlsx](QPFL_Tong_Hop_Ket_Qua_Tham_Van_2026.xlsx) | Bảng tính chi tiết 100% công thức Excel tự động, sẵn sàng phục vụ kiểm toán FSC. |

---

## 🔐 CƠ CHẾ BẢO MẬT & KIỂM SOÁT

Hệ thống được thiết kế hoàn toàn độc lập trên hạ tầng GitHub, giải quyết trọn vẹn bài toán bảo mật:

1. **Kiểm soát người xem (Viewer Access Control):**
   - Khi truy cập trang Báo cáo (`report.html`), hệ thống kích hoạt **Security Gate**.
   - Người xem bắt buộc phải nhập mã PIN xác thực nội bộ (**`QPFL2026`**).
   - Nếu không có mã PIN, toàn bộ biểu đồ, số liệu và chức năng tải file Excel đều bị khóa.

2. **Kiểm soát người trả lời (Respondent Verification):**
   - Trên Form tham vấn (`index.html`), người tham gia phải nhập đúng **Mã xác thực tham vấn (Survey PIN)** do QPFL cung cấp trong thư mời (**`QPFL2026`**).
   - Ngăn chặn hoàn toàn việc người ngoài spam hoặc gửi dữ liệu rác làm sai lệch kết quả đánh giá.

---

## ⚙️ CƠ CHẾ VẬN HÀNH TỰ ĐỘNG (GITHUB-NATIVE)

```
[ Người dùng điền Form ]
          │ (Nhập mã PIN QPFL2026)
          ▼
[ GitHub Issues API ]
          │ (Tạo Issue tự động kèm dữ liệu JSON)
          ▼
[ GitHub Actions Workflow ]
          │ (Kích hoạt máy chủ ảo Ubuntu)
          ├── 1. Append dữ liệu mới vào data/data.json
          ├── 2. Chạy Python update_excel.py (Author: Lộc Vũ Trung)
          ├── 3. Sinh file Master Excel (.xlsx) với công thức tự động
          └── 4. Tự động commit và đóng Issue
          ▼
[ Cập nhật tức thì vào Dashboard & File Excel tải về ]
```

---

## 📁 CẤU TRÚC THƯ MỤC

```
HCVF2026/
├── index.html                                    # Form tham vấn trực tuyến (có kiểm soát PIN)
├── report.html                                   # Dashboard báo cáo & thống kê (bảo mật PIN)
├── qpfl_logo.png                                 # Logo nhận diện thương hiệu QPFL
├── F1-1_QPFL_BanDo_TongThe_LamPhan_2026_GPS.jpg   # Bản đồ quy hoạch tổng thể lâm phận
├── QPFL_Tong_Hop_Ket_Qua_Tham_Van_2026.xlsx      # File Master Excel (Author: Lộc Vũ Trung)
├── update_excel.py                               # Script Python tự động build Excel từ JSON
├── data/
│   └── data.json                                 # Cơ sở dữ liệu JSON 13 phiếu hiện có & thống kê
└── .github/
    └── workflows/
        └── update_excel.yml                      # GitHub Action tự động hóa
```
