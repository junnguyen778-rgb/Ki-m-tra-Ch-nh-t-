# ViSpellCheck Pro - Windows 10/11

## Chức năng
- Kiểm tra chính tả tiếng Việt trong Word `.docx`.
- Kiểm tra văn bản trong Excel `.xlsx`, `.xlsm` (nhiều worksheet).
- Hiển thị vị trí, từ nghi lỗi và tối đa 5 gợi ý.
- Sửa từng lỗi, sửa các lỗi đã chọn hoặc sửa tất cả theo gợi ý đầu.
- Xuất báo cáo CSV.
- Giữ nguyên tệp gốc; lưu bản sửa với hậu tố `_checked`.
- Chạy kiểm tra offline. Chức năng tải/huấn luyện dữ liệu online chưa được tích hợp vào giao diện bản Pro này.

## Cách build
1. Trên Windows 10/11 64-bit, cài Python 3.10-3.12 (đánh dấu Add Python to PATH).
2. Giải nén toàn bộ thư mục.
3. Đặt `app.py`, `spell_engine.py`, `requirements.txt`, `ViSpellCheck.spec`, `build_exe.bat` cùng thư mục.
4. Nhấp đúp `build_exe.bat`.
5. EXE tạo tại `dist\ViSpellCheck.exe`.

## Cách dùng
1. Mở ViSpellCheck.exe.
2. Chọn Thêm tệp, chọn một hoặc nhiều tệp Word/Excel.
3. Chọn Kiểm tra.
4. Chọn một hay nhiều dòng lỗi. Có thể sửa nội dung ở ô Thay bằng rồi chọn Áp dụng cho dòng chọn.
5. Dùng Sửa lỗi đã chọn hoặc Sửa tất cả theo gợi ý đầu để áp dụng gợi ý.
6. Xuất báo cáo CSV nếu cần.

## Lưu ý kỹ thuật quan trọng
- Tệp gốc không bị ghi đè; bản sửa có tên `*_checked`.
- Word: văn bản được rà soát theo từng run để giữ định dạng. Từ bị Word chia thành nhiều run có thể không được nhận diện; một số cấu trúc nâng cao như textbox, SmartArt, chú thích và Track Changes chưa được xử lý.
- Excel: chỉ kiểm tra giá trị văn bản trong ô; bỏ qua công thức và ô không phải văn bản. Macro VBA được giữ khi xử lý `.xlsm` bằng openpyxl, nhưng nên thử trên bản sao vì một số thành phần Excel nâng cao có thể không được bảo toàn hoàn toàn.
- Thuật toán là kiểm tra âm tiết theo bộ luật và mô hình bigram khởi tạo nhỏ, không phải bộ từ điển chuẩn toàn diện. Tên riêng, thuật ngữ chuyên ngành và từ viết tắt có thể bị đánh dấu nhầm. Có thể bổ sung từ vào `~/.vi_spellcheck/user_words.txt`.
- Bản hiện tại dùng báo cáo CSV; chưa hỗ trợ sửa trực tiếp trong Microsoft Office bằng add-in.
- Nên thử nghiệm trên bản sao tài liệu trước khi dùng với hồ sơ quan trọng.


## Office Live - chạy song song với Word/Excel
- Cài Microsoft Office desktop trên Windows và cài thêm `pywin32` (đã thêm trong requirements).
- Mở Word/Excel và tài liệu cần kiểm tra trước, sau đó mở ViSpellCheck Pro và bấm **Office đang mở**.
- Chương trình kiểm tra tài liệu đang active; chế độ tự động mặc định bật, kiểm tra lại khoảng 3 giây khi nội dung thay đổi. Có thể bỏ chọn để chỉ kiểm tra khi bấm nút.
- Từ bị nghi sai được đổi màu chữ đỏ trực tiếp trong tài liệu đang mở. Màu ban đầu được khôi phục trước lần quét kế tiếp khi còn giữ được tham chiếu COM.
- Sửa lỗi: chọn các dòng cùng thuộc Word hoặc cùng thuộc Excel, bấm **Sửa lỗi chọn và lưu bản sao**. Word dùng SaveAs2 sang tên mới rồi sửa bản sao đang mở; Excel dùng SaveCopyAs rồi sửa workbook bản sao. Tệp gốc trên đĩa không bị ghi đè.
- Lưu ý: Word SaveAs2 chuyển cửa sổ tài liệu hiện tại sang bản sao có tên mới; Excel mở bản sao đã sửa. Nếu muốn giữ phiên bản gốc đang mở không thay đổi, hãy lưu công việc gốc trước khi chạy.
- Việc đổi màu chữ là thay đổi định dạng trong phiên Office. Đóng tài liệu gốc và chọn không lưu nếu không muốn lưu các dấu đỏ vào tệp nguồn. Bản sao đã sửa được lưu riêng.
- Kiểm tra tự động là polling 3 giây, không phải sự kiện gõ phím tức thời. Để tránh chậm trên workbook rất lớn, nên giới hạn UsedRange và chỉ mở workbook cần kiểm tra.
- Chỉ hỗ trợ Word/Excel desktop trên Windows; không hỗ trợ Office web.
