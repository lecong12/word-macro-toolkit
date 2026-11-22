# Word Macro Toolkit

Bộ dự án VS Code để quản lý, rà soát và phát triển 22 module VBA đang dùng trong Microsoft Word.

> **Trạng thái hiện tại:** giai đoạn nhập mã nguồn và tổ chức dự án. Chưa tự ý sửa logic VBA.  
> Thư mục `src/originals` là bản lưu nguyên gốc theo byte của các file người dùng đã tải lên. Chỉ chỉnh sửa trong `src/working`.

## Bắt đầu nhanh

1. Giải nén ZIP vào một thư mục ổn định, ví dụ `D:\\Projects\\WORD-MACRO-TOOLKIT`.
2. Mở VS Code → **File → Open Folder...** → chọn thư mục dự án.
3. Cài các extension được gợi ý trong tab Extensions. Extension chỉ hỗ trợ tô màu/cú pháp; không thay thế trình biên dịch VBA của Word.
4. Mở `src/working/Module01.bas` để xem mã. Các file `.bas` gốc có thể dùng bảng mã Windows-1252/ANSI; xem `docs/ENCODING.md` trước khi đổi encoding.
5. Dùng Source Control của VS Code để khởi tạo Git và lưu các mốc thay đổi.
6. Sau khi sửa, xuất/nhập module qua Word VBA Editor theo `docs/WORKFLOW.md`. Luôn thử trên bản sao `Normal.dotm`, không thử trực tiếp trên bản duy nhất đang dùng.

## Cấu trúc

```text
WORD-MACRO-TOOLKIT/
├── src/
│   ├── originals/   # 22 file gốc, không chỉnh sửa
│   └── working/     # bản làm việc để chỉnh sửa
├── docs/            # hướng dẫn, kiểm kê, quy trình, kiểm toán sơ bộ
├── tools/           # tiện ích kiểm tra tĩnh
├── backups/         # bản sao Normal.dotm và bản phát hành do bạn lưu
└── .vscode/         # thiết lập VS Code và extension gợi ý
```

## Quy tắc an toàn

- Không chỉnh sửa `src/originals`.
- Không ghi đè `Normal.dotm` duy nhất.
- Macro xóa/chuyển đổi hàng loạt phải có bản sao tài liệu thử nghiệm.
- Sau khi nhập vào Word: chạy **Debug → Compile VBAProject**, sau đó kiểm thử thủ công.
- Mỗi thay đổi nên là một commit Git riêng, ghi rõ module và lý do.
- Không coi kiểm tra tĩnh là xác nhận Macro chạy đúng: Word Object Model cần kiểm thử trong Word trên Windows.

## Tài liệu

- [Kiểm kê 22 Module](docs/INVENTORY.md)
- [Kiểm toán tĩnh và rủi ro](docs/STATIC-AUDIT.md)
- [Quy trình VS Code ↔ Word](docs/WORKFLOW.md)
- [Lưu ý mã hóa](docs/ENCODING.md)
- [Lộ trình nâng cấp](docs/ROADMAP.md)
