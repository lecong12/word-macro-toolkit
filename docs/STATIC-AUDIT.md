# Kiểm toán tĩnh – bản khởi tạo

## Phạm vi và giới hạn

Đã nhận đủ 22 file `.bas`. Kiểm tra này dựa trên mã nguồn được xuất; không có phiên chạy Word/Office Object Model trong môi trường tạo dự án. Do đó, các nhận định về tương thích, kết quả thao tác tài liệu và lỗi runtime phải được xác minh trong Microsoft Word trên Windows.

## Các điểm cần ưu tiên rà soát

| Module | Thủ tục | Điểm cần kiểm tra | Mức ưu tiên |
|---|---|---|---|
| 01 | `OpenNormalDotm` | Đường dẫn cứng chứa tên tài khoản `ky`; kiểm tra file tồn tại và tránh mở nhầm bản mẫu | Cao |
| 02 | `InsertPageNumbers` | Chỉ xử lý Section 1; kiểm tra khả năng chèn trường PAGE trùng với PageNumbers.Add; chạy Macro lặp lại | Cao |
| 03 | — | Không có thủ tục được nhận diện | Thấp |
| 04 | `ExportAutoCorrectUTF8` | Luồng UTF-8, xử lý lỗi, đường dẫn và định dạng dữ liệu xuất | Trung bình |
| 05 | `ImportAutoCorrectUTF8` | Dòng xuống hàng, mục trùng, dữ liệu đầu vào không hợp lệ và báo cáo nhập | Cao |
| 06 | `InsertFootnote` | Kiểm tra vùng chọn và hành vi chèn tại vị trí con trỏ | Thấp |
| 07 | `OpenContainingFolder` | Dùng `ThisDocument.Path` trong Normal; kiểm tra tài liệu đang hoạt động và đường dẫn có khoảng trắng | Trung bình |
| 08 | `ClearAutoCorrectEntries` | Xóa hàng loạt toàn bộ AutoCorrect; cần xác nhận, sao lưu và giới hạn phạm vi | Rất cao |
| 09 | `QuoteToTextSelct` | Tên thủ tục sai chính tả; đặt Italic=False có thể làm mất định dạng ban đầu | Trung bình |
| 10 | — | Không có thủ tục được nhận diện | Thấp |
| 11 | `DeleteTextWithNormalStyle` | Xóa các đoạn Normal – có thể xóa phần lớn văn bản; thay đổi collection trong vòng lặp | Rất cao |
| 12 | `DeleteUnusedStyles` | Xác định Style được dùng chưa bao quát mọi story/range; xóa Style có thể ảnh hưởng định dạng | Rất cao |
| 13 | `FormatFirstTwoWords` | Tên không phản ánh giới hạn thực tế; kiểm tra chạy lặp và việc ghi đè định dạng | Cao |
| 14 | `HighlightParagraphsWithStyle` | Kiểm tra InputBox Cancel/rỗng, Style không tồn tại, tên Style cục bộ | Trung bình |
| 15 | `RemoveHighlight` | Chỉ xóa Highlight vàng và chỉ duyệt Paragraph; xác định phạm vi mong muốn | Trung bình |
| 16 | `ConvertFootnotesToText` | Xóa Footnote trong khi duyệt collection; kiểm tra thứ tự duyệt ngược và nội dung phức tạp | Cao |
| 17 | `XoaDongTrong` | Điều kiện `Len(Range.Text)=1` không bắt được dòng có dấu cách/Tab; cần giữ đoạn định dạng cấu trúc | Trung bình |
| 18 | `highlightPartialMatches` | Khai báo `Dim i, J As Long` khiến `i` là Variant; kiểm tra tìm chuỗi hai chiều, vị trí tô và phục hồi ScreenUpdating khi lỗi | Cao |
| 19 | `SplitDocumentByPageBreak` | Kiểm tra tài liệu chưa lưu, ranh giới Range, file đầu ra trùng tên và bảo toàn Section/Header/Footer | Rất cao |
| 20 | `MergeWordFiles` | Kiểm soát file trùng/đang mở, định dạng, Section, clipboard, hủy thao tác và lưu kết quả | Cao |
| 21 | `ReplaceSquareWithCheckboxInTable` | Giá trị `findText` hiện là dấu `?`; xác định ký tự vuông thực tế; chỉ thay đúng ký tự, không xóa cả ô | Rất cao |
| 22 | `ConvertDocToDocx` | Kiểm soát ghi đè, file tạm, lỗi mở/lưu, đóng tài liệu và giải phóng phiên Word COM | Rất cao |

## Quan hệ phụ thuộc

- Module04 và Module05 có quan hệ dữ liệu qua file AutoCorrect xuất/nhập; cần xác minh định dạng file có tương thích hai chiều.
- Module08 thao tác cùng tập hợp AutoCorrect nhưng là thao tác xóa độc lập.
- Module11–15 cùng làm việc với Paragraph/Style/Highlight; chưa đủ căn cứ nói chúng gọi trực tiếp lẫn nhau.
- Module02, 06, 09, 16–19, 21 thao tác nội dung/định dạng tài liệu.
- Module01, 07, 19, 20, 22 liên quan đường dẫn, mở, tách, gộp hoặc chuyển đổi file.
- Qua rà soát ban đầu, chưa xác lập được lời gọi trực tiếp giữa các Sub ở các module; đây không phải bằng chứng tuyệt đối về không có phụ thuộc vì cần phân tích đầy đủ mọi tham chiếu và kiểm thử trong VBA.

## Lưu ý mã hóa

Các file `.bas` chứa byte không phải UTF-8 thuần. Một số chú thích/chuỗi tiếng Việt hiển thị lỗi khi đọc bằng UTF-8. Không chạy chuyển mã hàng loạt trên bản gốc. Xem `ENCODING.md` và kiểm tra trong VBA Editor trước khi sửa chuỗi hiển thị cho người dùng.
