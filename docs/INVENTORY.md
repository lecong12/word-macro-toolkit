# Kiểm kê nguồn VBA

Các tên thủ tục dưới đây được trích tự động từ file `.bas` được tải lên; phần mô tả chức năng là nhận diện sơ bộ, cần xác nhận khi kiểm thử trong Word.

| Module | Thủ tục tìm thấy | Kích thước (byte) | Trạng thái |
|---|---|---:|---|
| Module01 | OpenNormalDotm | 215 | Có mã |
| Module02 | InsertPageNumbers | 642 | Có mã |
| Module03 | — | 31 | Chưa có Sub/Function |
| Module04 | ExportAutoCorrectUTF8 | 1639 | Có mã |
| Module05 | ImportAutoCorrectUTF8 | 1838 | Có mã |
| Module06 | InsertFootnote | 181 | Có mã |
| Module07 | OpenContainingFolder | 446 | Có mã |
| Module08 | ClearAutoCorrectEntries | 494 | Có mã |
| Module09 | QuoteToTextSelct | 524 | Có mã |
| Module10 | — | 32 | Chưa có Sub/Function |
| Module11 | DeleteTextWithNormalStyle | 469 | Có mã |
| Module12 | DeleteUnusedStyles | 1589 | Có mã |
| Module13 | FormatFirstTwoWords | 2451 | Có mã |
| Module14 | HighlightParagraphsWithStyle | 602 | Có mã |
| Module15 | RemoveHighlight | 441 | Có mã |
| Module16 | ConvertFootnotesToText | 318 | Có mã |
| Module17 | XoaDongTrong | 445 | Có mã |
| Module18 | highlightPartialMatches | 1673 | Có mã |
| Module19 | SplitDocumentByPageBreak | 2089 | Có mã |
| Module20 | MergeWordFiles | 1308 | Có mã |
| Module21 | ReplaceSquareWithCheckboxInTable | 1896 | Có mã |
| Module22 | ConvertDocToDocx | 2033 | Có mã |

## Ghi chú

- Có 22 file nguồn, được giữ nguyên trong `src/originals`.
- Module03 và Module10 chỉ chứa khai báo tên module, không có Sub/Function được bộ trích xuất nhận diện.
- File ban đầu có tên `Module 2.bas`; bản làm việc được chuẩn hóa tên thành `Module02.bas`, nội dung không đổi.
- Không suy diễn rằng Module trống chắc chắn không còn được sử dụng; cần đối chiếu dự án Normal.dotm hiện hành trước khi loại bỏ.
