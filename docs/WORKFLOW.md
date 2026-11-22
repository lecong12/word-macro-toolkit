# Quy trình làm việc VS Code ↔ Word VBA

## A. Chỉnh sửa an toàn

1. Tạo commit Git trước khi sửa.
2. Chỉ sửa file trong `src/working`.
3. Giữ nguyên tên module trong dòng `Attribute VB_Name = "..."` trừ khi có kế hoạch đổi tên và đã rà soát toàn bộ tham chiếu.
4. Không tự ý đổi bảng mã file. Xem `ENCODING.md`.
5. Lưu file và xem Git diff để xác nhận chỉ có thay đổi dự kiến.

## B. Đưa một Module vào Word

1. Mở Word và mở VBA Editor bằng `Alt + F11`.
2. Sao lưu `Normal.dotm` trước khi thao tác.
3. Trong Project Explorer, chọn đúng project `Normal`.
4. Với module cần cập nhật, xuất bản hiện tại ra thư mục backup.
5. Xóa module cũ chỉ sau khi đã xác nhận đúng module và có bản xuất dự phòng.
6. Chọn **File → Import File...** và nhập file `.bas` từ `src/working`.
7. Chọn **Debug → Compile VBAProject**.
8. Chạy thử trên bản sao tài liệu; không thử Macro phá hủy trên tài liệu gốc.
9. Khi kiểm thử đạt, lưu Normal.dotm và tạo commit/tag Git tương ứng.

## C. Phát hành bản sao Normal.dotm

- Đóng toàn bộ cửa sổ Word trước khi sao chép file Normal.dotm.
- Lưu bản hiện tại vào `backups` trước khi thay thế.
- Không ghi đè bản duy nhất.
- Trên máy khác, sao lưu Normal.dotm hiện có trước khi triển khai.
- Kiểm tra Macro Settings/Trust Center và chính sách bảo mật của tổ chức nếu Macro bị chặn.

## D. Kiểm thử tối thiểu cho mỗi Macro

- Tài liệu trống.
- Tài liệu thông thường có nhiều đoạn.
- Tài liệu nhiều Section/Header/Footer (nếu liên quan).
- Vùng chọn rỗng và vùng chọn phức tạp (nếu liên quan).
- Chạy hai lần liên tiếp để phát hiện thao tác bị lặp.
- Hủy hộp thoại/chọn file.
- Đường dẫn có dấu tiếng Việt và khoảng trắng.
- File đầu ra đã tồn tại.
- Lỗi giữa chừng: xác nhận tài liệu và ứng dụng được đóng/khôi phục trạng thái đúng.

## E. Git gợi ý

```powershell
git init
git add .
git commit -m "chore: import original Word VBA modules"
git status
```

Mỗi thay đổi tiếp theo nên commit riêng, ví dụ:

```powershell
git add src/working/Module18.bas docs/
git commit -m "fix(module18): correct partial text matching"
```
