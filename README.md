# Lọc lịch sát hạch Trung tâm Nguyễn Trình

Ứng dụng Streamlit để tải thông báo PDF, lọc dòng lịch có từ khóa Nguyễn Trình, kiểm tra/chỉnh sửa và xuất Word `.doc` (RTF tương thích Word), `.docx`, PDF và Excel `.xlsx`.

## Chạy thử trên máy
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy lên Render
1. Tạo repository GitHub và đưa các file trong thư mục này lên.
2. Vào Render, chọn **New + → Blueprint**, kết nối repository có `render.yaml`.
3. Chọn deploy. Khi hoàn tất, mở URL Render và tải PDF lên.

## Độ chính xác và lưu ý
- Ưu tiên trích xuất bảng PDF bằng pdfplumber; nếu bảng không trích xuất đủ, dùng text fallback.
- Người dùng có thể sửa bảng kết quả trước khi xuất.
- Luôn đối chiếu với PDF gốc trước khi thông báo thí sinh. Với PDF scan/ảnh hoặc bảng phức tạp, cần bổ sung OCR và kiểm tra thủ công.
- `.doc` được tạo dưới dạng RTF để Word có thể mở; `.docx` cũng được cung cấp để giữ định dạng tốt hơn.
- Bản mẫu không tự gửi SMS/Zalo và không lưu hồ sơ lên cơ sở dữ liệu.


## V4 sửa lỗi theo PDF thông báo tháng 11/2026
- Dùng bảng 5 cột thực tế bằng `pdfplumber` mặc định, không dùng chiến lược kẻ dòng dễ làm lệch cột.
- Ghép ô tiếp nối qua trang (ví dụ dòng 06).
- Xuất PDF nhúng font DejaVu Sans được đóng gói cùng ứng dụng để hỗ trợ tiếng Việt trên Render.
- Có cột STT, ngày, cơ sở đào tạo, hạng thi, số lượng, địa điểm và trang nguồn.
