# Báo cáo đánh giá năng lực chuyển đổi (Stress Test Report)

- **Thời gian chạy test**: 2026-06-18 13:19:10
- **Tổng số tệp mẫu đã tạo**: 100 tệp docx
- **Số tệp chuyển đổi thành công**: 100 (100.0%)
- **Số tệp thất bại**: 0
- **Tổng thời gian xử lý**: 9.25 giây (0.15 phút)
- **Thời gian trung bình mỗi tệp**: 0.092 giây
## Đánh giá khả năng tối đa
1. **Về mặt Tốc độ**: Tốc độ xử lý file Word `.docx` cực kỳ nhanh (không qua OCR), trung bình khoảng **0.092 giây/file**. Việc xử lý hàng trăm hoặc hàng nghìn file chỉ mất vài phút.
3. **Khả năng tối đa của Streamlit (Giao diện web)**:
   - **Dung lượng file tải lên**: Giới hạn mặc định của Streamlit là **200 MB** tổng cộng cho tất cả các file tải lên cùng lúc. Với các file `.docx` mẫu chỉ nặng khoảng 17 KB/file, giới hạn 200 MB tương đương với khoảng **11,000 file** tải lên đồng thời.
   - **Giới hạn trình duyệt**: Tải lên hàng nghìn file cùng lúc có thể làm đơ trình duyệt của người dùng khi chọn file. Khuyến nghị thực tế tối ưu qua giao diện web là khoảng **100 - 300 file mỗi lượt** để đảm bảo giao diện hiển thị mượt mà.
   - **Nếu chạy script offline**: Hầu như không có giới hạn, có thể chạy hàng vạn file liên tục mà không gặp sự cố về tài nguyên máy tính.
