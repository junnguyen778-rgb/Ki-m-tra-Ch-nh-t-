# ViSpellCheck Pro - Windows 10/11

## Cách dùng nhanh (không cần thêm tệp)
1. Mở tài liệu Word và/hoặc Excel như bình thường.
2. Mở ViSpellCheck Pro, bấm **QUÉT NGAY** (khung xanh dương). Chương trình tự nhận diện mọi tài liệu đang mở, tô đỏ từ nghi sai ngay trong tài liệu và liệt kê trong bảng.
3. Bấm một dòng để xem gợi ý (các nút gợi ý bấm được) hoặc nhập từ thay thế, rồi **ÁP DỤNG** / **SỬA MỤC ĐÃ CHỌN** / **SỬA TẤT CẢ**.
4. Bản đã sửa luôn lưu thành tệp mới `_checked`; tệp gốc không bị ghi đè.
- **THEO DÕI TỰ ĐỘNG** (khung xanh lá, mặc định tắt): tự quét lại mỗi 3 giây khi nội dung đổi.
- **XUẤT BÁO CÁO CSV** và **MỞ TỆP TỪ ĐĨA…** (tùy chọn, cho tệp chưa mở) ở khung tím.
- Đóng chương trình sẽ trả lại màu chữ gốc cho các từ đã tô đỏ.

## Yêu cầu
Windows 10/11, Microsoft Office bản desktop (Word/Excel) cho tính năng quét tài liệu đang mở. Không hỗ trợ Office web.

## Cấu trúc mã
- `app.py`: giao diện (nút/khung 3D tự vẽ bằng Tkinter Canvas).
- `office_live.py`: quét, tô đỏ, sửa Word/Excel đang mở qua COM (pywin32).
- `doc_scan.py`: quét tệp .docx/.xlsx/.xlsm trên đĩa.
- `spell_engine.py`: bộ kiểm tra âm tiết + ngữ cảnh (không đổi).

## Build
- `build_exe.bat` → `dist\ViSpellCheck.exe`.
- `build_installer.bat` → `dist\ViSpellCheck-Setup.exe` (cần thêm Inno Setup 6).

## Lưu ý
- Word: quét theo từng đoạn; vị trí có thể lệch nếu đoạn chứa field/mã ẩn. Sau khi sửa, Word chuyển cửa sổ sang bản sao `_checked` (SaveAs2).
- Excel: bỏ qua ô công thức; ô được sửa sẽ mất định dạng từng ký tự bên trong ô.
- Tô đỏ là thay đổi định dạng trong phiên Office đang mở; đừng lưu tài liệu gốc khi còn dấu đỏ (hoặc đóng ViSpellCheck trước để trả màu).
- Tên riêng, thuật ngữ, từ viết tắt có thể bị đánh dấu nhầm; bổ sung vào `~/.vi_spellcheck/user_words.txt`.
