# SSU CTV Assignment & Attendance

Web app tĩnh để quản lý phân công và chấm công CTV SSU ngay trên trình duyệt. Dự án được thiết kế để dễ đưa lên GitHub, dễ đọc và dễ mở rộng về sau.

## Tính năng chính

- Đọc dữ liệu từ Excel/CSV ngay trên trình duyệt.
- Phân công C1 dựa trên sheet `Lịch ctv đăng kí`.
- Ưu tiên người có tên in đậm hoặc viết hoa.
- Chấm công thủ công bằng tick checkbox.
- Hỗ trợ upload ảnh chấm công để lưu làm chứng cứ.
- Phân công DDV từ file `BCS CTV.xlsx` và file lịch phân công.
- Ghép theo ca `Sáng` / `Chiều`.
- Ưu tiên người mới hoặc người đi ít.
- Tính bảng Rank tổng.
- Xuất kết quả ra Excel chỉ với một nút bấm.

## Cấu trúc dự án

- `index.html`: Khung giao diện web.
- `style.css`: Toàn bộ phần giao diện, tông màu hồng pastel.
- `script.js`: Logic đọc file, xử lý dữ liệu, phân công, chấm công và xuất Excel.
- `README.md`: Tài liệu hướng dẫn sử dụng.

## Cách chạy

### Cách 1: Mở trực tiếp

1. Tải toàn bộ thư mục dự án về máy.
2. Mở `index.html` bằng trình duyệt hiện đại như Chrome, Edge hoặc Firefox.
3. Upload các file Excel/CSV cần thiết.
4. Bấm `Đọc và xử lý dữ liệu`.
5. Xem bảng kết quả và bấm `Xuất kết quả Excel` khi cần.

### Cách 2: Chạy qua Live Server

1. Mở thư mục dự án trong VS Code.
2. Cài extension Live Server nếu chưa có.
3. Mở `index.html` bằng Live Server để có trải nghiệm ổn định hơn khi làm việc lâu dài.

## Định dạng file đầu vào

### 1. C1 - `CTV SSU.xlsx`

- Sheet chính: `Lịch ctv đăng kí`
- Cột nên có:
  - Họ tên
  - Ca
  - Điểm hoặc trạng thái đi trực
  - Ghi chú ưu tiên nếu có

Hệ thống sẽ ưu tiên người có tên in hoa hoặc đánh dấu nổi bật nếu dữ liệu nguồn thể hiện điều đó.

### 2. DDV - `BCS CTV.xlsx`

- Sheet chính: `DDV`
- Cột nên có:
  - Họ tên
  - Ca
  - Số lần đi
  - Cố định
  - Ưu tiên

### 3. File `Phân công DDV`

- Có thể là Excel hoặc CSV.
- Nên chứa các cột tương tự:
  - Họ tên
  - Ca
  - Đã đi / Số lần đi
  - Ghi chú

## Quy tắc logic

### C1

- Đọc dữ liệu từ sheet `Lịch ctv đăng kí`.
- Ưu tiên người có tên in hoa hoặc được đánh dấu ưu tiên.
- Chấm công:
  - Đi đủ: +1 điểm
  - Không đi: -1 điểm
- Có thể tick thủ công để cập nhật trạng thái ngay trên giao diện.

### DDV

- Đọc dữ liệu từ sheet `DDV` trong `BCS CTV.xlsx`.
- Đọc danh sách từ file `Phân công DDV`.
- Chỉ ghép theo đúng ca `Sáng` hoặc `Chiều`.
- Ưu tiên người chưa đi hoặc đi ít nhất.
- Hạn chế phân công trùng lặp cho cùng một người.

### Rank tổng

- Tổng hợp dữ liệu từ C1, DDV và lịch phân công.
- Điểm Rank được tính từ:
  - Trạng thái đã đi / chưa đi
  - Mức ưu tiên
  - Số lần đi trước đó
  - Cờ cố định nếu có

## Lưu ý kỹ thuật

- Dự án chạy thuần frontend, không cần backend.
- Thư viện `xlsx` được nạp qua CDN trong `index.html`.
- Nếu muốn dùng offline hoàn toàn, có thể thay CDN bằng bản cài local sau này.
- Logic hiện tại được viết để dễ bảo trì và có thể tách tiếp thành module khi dự án lớn hơn.

## Gợi ý phát triển tiếp

- Tách dữ liệu thành tab theo ca.
- Lưu trạng thái vào LocalStorage hoặc backend.
- Thêm bộ lọc nâng cao cho từng nhóm CTV.
- Thêm export nhiều sheet Excel với style đẹp hơn.
- Tích hợp nhận diện ảnh chấm công nếu sau này có AI hoặc backend xử lý ảnh.

## License

Dùng nội bộ cho dự án SSU CTV. Nếu đưa lên GitHub công khai, nên bổ sung license phù hợp với mục đích sử dụng của nhóm.
