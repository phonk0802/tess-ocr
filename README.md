# tess-ocr
Sử dụng Tesseract-OCR trích xuất thông tin trên mẫu phiếu đăng ký sinh viên

images.rar bao gồm 30 ảnh chụp các phiếu đăng ký thông tin sinh viên.

File scanner_students.py dùng để trích xuất từ ảnh các trường thông tin: Họ tên, Mã sinh viên, Ngày sinh, Khóa, Khoa, Ngành, Điểm trung bình kỳ 1, Điểm trung bình kỳ 2, Điểm trung bình cả năm từ mẫu phiếu có trong images.rar.

result.rar bao gồm 3 file: 
- DSSV_chuẩn.xlsx là bản xlsx có thông tin chính xác.
- DSSV_đọc.xlsx là bản xlsx do Tesseract OCR đọc được.
- Diff.xlsx là bản xlsx thể hiện độ sai lệch giữa bản gốc và bản đọc được. Trong đó, các ô sai lệch sẽ đánh dấu đỏ và ghi rõ "bản gốc --> bản đọc được".
