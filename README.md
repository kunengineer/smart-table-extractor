# 📊 Chương trình Chuyển Đổi Bảng Dữ Liệu → Excel (.xlsx)

Ứng dụng web được xây dựng bằng Python và Streamlit, hỗ trợ tự động nhận diện cấu trúc bảng từ các tệp đầu vào (PDF Text, PDF Scan, Word `.doc`/`.docx`, Hình ảnh `.png`/`.jpg`/`.jpeg`) và xuất sang tệp Excel được định dạng chuyên nghiệp, đẹp mắt.

---

## 🚀 Hướng Dẫn Chạy Chương Trình

1. **Kích hoạt môi trường ảo Python**:
   ```bash
   .venv\Scripts\activate
   ```
2. **Chạy ứng dụng Streamlit**:
   ```bash
   streamlit run app.py
   ```
   *Giao diện web sẽ tự động mở tại địa chỉ `http://localhost:8501`.*

---

## 🛠️ Cài Đặt Các Phần Mềm Phụ Trợ (Bắt buộc cho PDF Scan / Ảnh)

Để ứng dụng có thể xử lý các file **PDF quét (Scan)** hoặc **Hình ảnh**, bạn cần cài đặt hai công cụ hệ thống sau đây vào máy tính Windows của mình:

### 1. Cài đặt Tesseract OCR (Nhận diện chữ viết từ ảnh)
*   **Bước 1**: Tải xuống trình cài đặt Tesseract OCR cho Windows từ: [UB-Mannheim Tesseract Installer](https://github.com/UB-Mannheim/tesseract/wiki) (chọn bản 64-bit mới nhất).
*   **Bước 2**: Chạy file cài đặt `.exe` vừa tải xuống.
*   **Bước 3**: Bấm **Next** cho đến khi hoàn thành. Hãy giữ nguyên đường dẫn cài đặt mặc định:
    `C:\Program Files\Tesseract-OCR`
    *(Hệ thống của chúng tôi được thiết kế để tự động phát hiện Tesseract tại đường dẫn này, bạn không cần phải cấu hình biến môi trường PATH thủ công)*.

### 2. Cài đặt Poppler (Chuyển đổi trang PDF sang hình ảnh)
*   **Bước 1**: Tải xuống tệp Poppler dạng zip từ: [Poppler Windows Releases](https://github.com/oschwartz10612/poppler-windows/releases) (Tải bản zip mới nhất, ví dụ: `Release-xx.xx.x-x.zip`).
*   **Bước 2**: Giải nén tệp zip vừa tải.
*   **Bước 3**: Sao chép thư mục đã giải nén vào ổ đĩa **C:\\** và đổi tên thư mục thành **poppler** để có đường dẫn:
    `C:\poppler`
    *(Chương trình sẽ tự động dò quét các thư mục con trong C:\ có chứa chữ "poppler" để chạy ngay mà không cần cấu hình PATH)*.

---

## 🎨 Điểm Nổi Bật của Định Dạng Excel Đầu Ra
Tệp Excel kết xuất từ chương trình được thiết kế theo tiêu chuẩn thẩm mỹ cao:
*   **Font chữ hiện đại**: Sử dụng bộ font `Segoe UI` tinh tế và gọn gàng.
*   **Trình bày rõ ràng**: Cột tiêu đề (Header) được tô nền xanh dương đậm `#1F4E79`, chữ trắng in đậm nổi bật.
*   **Dễ theo dõi dòng**: Tự động tô màu nền xen kẽ (Zebra rows) với màu xanh nhạt `#F2F7FA` dịu mắt.
*   **Độ rộng cột tự động**: Căn chỉnh độ rộng theo chiều dài nội dung thực tế của cột để không bao giờ bị khuất chữ.
*   **Căn lề thông minh**: Cột số thứ tự (STT) được căn giữa, cột số (Số lượng/Đơn giá) căn phải và định dạng dấu phân cách phần nghìn, các cột văn bản căn trái.
*   **Cố định tiêu đề**: Dòng đầu tiên (Header) được đóng băng (`Freeze Panes`) để luôn hiển thị khi cuộn chuột.
