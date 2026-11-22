# Mã hóa file VBA `.bas`

## Điều đã quan sát

Các file tải lên chứa byte không phải UTF-8 thuần. Khi đọc theo UTF-8, một số chữ tiếng Việt hiện thành ký tự lỗi; khi đọc theo Windows-1252 cũng có thể không khôi phục đúng tiếng Việt nếu file gốc được lưu bằng code page Windows khác.

Vì vậy:
- `src/originals` giữ nguyên byte gốc, không chuyển mã.
- `src/working` ban đầu cũng là bản sao byte-identical để tránh vô tình làm hỏng dữ liệu.
- Thiết lập VS Code trong dự án dùng `windows1252` làm lựa chọn ban đầu, nhưng đây chỉ là gợi ý hiển thị, không phải kết luận chắc chắn về code page gốc.

## Cách xác định đúng

1. Mở module trong Word VBA Editor trên chính máy đã tạo Macro.
2. Đối chiếu các chuỗi tiếng Việt trong cửa sổ code.
3. Nếu tiếng Việt hiển thị đúng trong Word nhưng file xuất ra sai, cần xác định code page khi export.
4. Chỉ chuyển file làm việc sang encoding khác sau khi đã lưu bản gốc và kiểm tra diff.
5. Không thay hàng loạt dấu `?`: có thể chúng vốn là ký tự thật hoặc đã mất thông tin từ trước.

## Khuyến nghị

Với mã VBA cổ điển, ưu tiên giữ nguyên encoding mà Word/VBE xuất ra. Nếu cần chuẩn hóa tiếng Việt, làm từng Module, kiểm tra trong Word sau khi Import.
