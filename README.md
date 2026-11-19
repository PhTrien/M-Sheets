# M-Sheets - Microsoft Excel Add-in (.xlam)

Add-in đa năng chuyên sâu cho Microsoft Excel, tối ưu hóa cho công tác hành chính văn phòng, quản lý địa chính, kế toán và xử lý dữ liệu chuẩn Việt Nam.

---

## 🌟 Cấu trúc 7 Nhóm chức năng trên Ribbon ("M-Sheets")

### 1. VĂN BẢN & HỌ TÊN
- **Hoàn Tác (Ctrl+Z)**: Khôi phục dữ liệu lên đến 30 bước trước khi can thiệp.
- **Tách Họ Tên (Ctrl+Shift+T)**: Tự động phân tách Họ + Tên đệm thành cột "Họ lót" và cột "Tên".
- **Ghép Họ Tên**: Ghép 2 cột Họ lót và Tên thành 1 cột Họ tên hoàn chỉnh.
- **Sắp Xếp VN (Ctrl+Shift+V)**: Sắp xếp bảng tính theo chuẩn tiếng Việt (Tên trước, Họ lót sau), bảo tồn số CCCD, ngày sinh và định dạng.
- **Đổi Kiểu Chữ**: CHỮ IN HOA, chữ in thường, Viết Hoa Đầu Từ (Proper Case), Viết hoa đầu câu (Sentence case).
- **Chuyển Mã Font**: TCVN3 (.VnTime) sang Unicode, VNI-Times sang Unicode.
- **Xóa Dấu VN**: Chuyển đổi văn bản có dấu thành không dấu.
- **Làm Sạch Chữ**: Xóa khoảng trắng thừa (TRIM), khoảng trắng non-breaking (`ChrW(160)`) và ký tự ẩn.

### 2. SỐ LIỆU & ĐỊA CHÍNH
- **Đọc Số Ra Chữ**:
  - Đọc Số (Thuần Túy)
  - Đọc Tiền VNĐ (Thêm Cột Chữ - Ctrl+Shift+D)
  - Đọc Diện Tích Đất ($m^2$)
  - Đọc Tiền USD ($)
- **Tính Thuế Đất**: Tự động tính Thuế TNCN (2%), Lệ phí trước bạ (0.5%), tổng nghĩa vụ tài chính và định dạng chuẩn VN.
- **Thẩm Định CCCD (Ctrl+Shift+C)**: Tự động bù số 0 đầu, kiểm tra hợp lệ 12 số, cảnh báo ô sai định dạng.
- **Trích Xuất CCCD**: Bóc tách từ số CCCD ra 3 cột: Năm sinh, Giới tính, Tỉnh/Thành phố đăng ký khai sinh.
- **Chèn Mã QR**: Tạo và chèn hình ảnh mã QR từ văn bản hoặc ô được chọn trực tiếp vào bảng tính.
- **Tách Địa Chỉ (3 Cấp)**: Tách địa chỉ thành: Cấp Cơ sở/Thôn/TDP, Cấp Xã/Phường, Cấp Quận/Huyện/Tỉnh.
- **Chuẩn Hóa SĐT**: Bù số 0 đầu, chuyển đổi +84/84 về 0, chuẩn hóa 10 số.
- **Định Dạng Số VN**: Phân cách hàng nghìn bằng dấu chấm (`1.234.567`).
- **Đánh STT**: Tự động đánh số thứ tự 1, 2, 3...

### 3. NGÀY & ĐỐI SOÁT
- **Xử Lý Ngày Tháng**:
  - Chuẩn Hóa Ngày (`dd/mm/yyyy`)
  - Tách Ngày - Tháng - Năm (3 Cột để in phôi văn bản / GCN)
  - Đọc Ngày Thành Chữ Hành Chính ("ngày ... tháng ... năm ...")
- **Đối Soát 2 Bảng**: So khớp 2 bảng dữ liệu theo cột khóa, tự động tô xanh dòng trùng khớp và tô đỏ dòng sai lệch/thiếu.
- **Đổi Số La Mã**: Chuyển đổi qua lại giữa số tự nhiên (1, 2, 3...) và số La Mã (I, II, III...).

### 4. RÀ SOÁT & ĐỊNH DẠNG
- **Soi Dòng Cột (Ctrl+Shift+H)**:
  - Bật soi cả dòng và cột (làm sáng nổi bật).
  - Chỉ soi dòng ngang.
  - Chọn bảng màu: Màu vàng nổi bật (mặc định), Màu xanh da trời, Màu xanh lá mạ, Màu cam đào.
  - Tắt chế độ soi.
- **Tô Màu Trùng**: Tìm và làm nổi bật các ô có giá trị trùng lặp.
- **Tô Màu Zebra**: Tô màu dòng xen kẽ dịu mắt.
- **Lọc Duy Nhất**: Trích xuất các giá trị độc nhất sang cột mới.
- **Xóa Định Dạng**: Xóa sạch định dạng, đưa về chuẩn Times New Roman 12.
- **Xóa Trống**: Xóa các dòng/cột trống hoàn toàn.
- **Sửa Lỗi #N/A**: Tự động bao bọc công thức bằng `IFERROR`.
- **Xử Lý Gộp Ô**:
  - Bỏ Gộp & Điền Đầy (Unmerge and Fill Down - chuẩn hóa để làm PivotTable / lọc dữ liệu).
  - Gộp Ô Cùng Giá Trị (Merge identical consecutive cells vertically).

### 5. QUẢN LÝ & BẢO MẬT
- **Gộp Sheet**: Gộp dữ liệu từ nhiều sheet về một bảng tính tổng.
- **Hiện Sheet Ẩn**: Làm hiện toàn bộ các sheet bị ẩn (kể cả VeryHidden).
- **Dọn Rác File**: Xóa Name lỗi (`#REF!`), xóa Object tàng hình, tối ưu lại UsedRange giảm dung lượng file.
- **Khóa Công Thức**: Khóa bảo vệ tất cả các ô công thức, chỉ cho phép chỉnh sửa ô nhập liệu.
- **Mở Khóa Sheet**: Mở khóa bảo vệ bảng tính nhanh chóng.
- **Bật Lưới Ô**: Bật/tắt hiển thị đường lưới ô tính.

### 6. IN ẤN & XUẤT BẢN ⭐
- **Icon máy in chuyên nghiệp**: Sử dụng `PrintPreviewAndPrint` thay cho biểu tượng vòng tròn cũ.
- **Căn Lề A4 (Ctrl+Shift+P)**: Căn lề chuẩn Nghị định 30/2020/NĐ-CP (Trái 2.5cm, Phải 1.5cm, Trên 2.0cm, Dưới 2.0cm, căn giữa trang in).
- **Vừa 1 Trang**: Tự động co chiều ngang vừa khít 1 trang in.
- **Lặp Tiêu Đề**: Thiết lập dòng tiêu đề lặp lại khi in bảng dài nhiều trang.
- **Xuất PDF**: Xuất bảng tính ra file PDF chất lượng cao.
- **Dán Word**: Chuẩn hóa viền kẻ, font Times New Roman và copy sẵn sàng paste sang Word.
- **Xuất Word Mẫu**: Fast Mail Merge - tự động lấy dữ liệu từng dòng trong Excel điền vào các thẻ `{{TenCot}}` trong tệp Word mẫu.

### 7. TRỢ GIÚP & TÁC GIẢ
- **Hướng Dẫn Sử Dụng**: Tra cứu phím tắt và tài liệu hướng dẫn.
- **Gửi Phản Hồi**: Gửi thư đóng góp ý kiến.
- **Thông Tin Tác Giả**: Tô Quang Vĩnh - VPĐKĐĐ tỉnh Khánh Hòa.

---

## 🛠️ Hướng dẫn Cài đặt & Sử dụng

1. **Cài đặt tự động**:
   - Nhấp đúp vào file `Install-Excel-Addin.bat`.
   - File `M-Sheets.xlam` sẽ được tự động sao chép vào thư mục `XLSTART` của Excel và gỡ chặn Macro.
   - Mở Microsoft Excel lên, bạn sẽ thấy ngay thẻ **"M-Sheets"** trên thanh Ribbon.

2. **Gỡ cài đặt**:
   - Nhấp đúp vào file `Uninstall-Excel-Addin.bat` để xóa add-in khỏi Excel.

3. **Tự biên dịch từ mã nguồn (`src/`)**:
   - Mở PowerShell trong thư mục `Excel` và chạy:
     ```powershell
     powershell -ExecutionPolicy Bypass -File .\Build-Excel-Addin.ps1
     ```
