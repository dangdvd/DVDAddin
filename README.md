# DVDAddin
DVD Addin - Hỗ trợ công tác lập, in ấn hồ sơ hàng loạt tự động nhanh chóng. Bổ sung các tính năng mở rộng cho Excel.

Tác giả: Đặng Văn Đằng

Tính năng:
- Giải đề thi trắc nghiệm bằng AI Assistant.

- Xóa mật khẩu bảo vệ Sheet bị quên.

- Định dạng văn bản: in hoa, thường, viết hoa chữ cái đầu từ mỗi câu, viết hoa chữ cái đầu mỗi từ, viết hoa viết thường thông minh.

- Đánh số thứ tự theo hàng, theo cột, theo bước nhảy của số.

- Đảo văn bản qua ký tự ngăn cách (vd: dùng đảo văn bản song ngữ).

- Đổi bảng mã giữa Unicode <-> VNI Windows <-> TCVN3 không bị lỗi font.

- In nghiêng văn bản sau ký tự ngăn cách (vd: dùng in nghiêng văn bản song ngữ).

- Dịch văn bản tạo song ngữ theo google dịch, chatGPT, google Gemini.

- Liệt kê file trong folder, liệt kê file theo cây thư mục folder.

- Liệt kê Folder theo cây thư mục và tạo cây thư mục Folder theo danh sách trên excel.

- Xóa, chèn nhiều dòng cột.

- Xóa name range, cell styles rác giúp giảm dung lượng và mở file nhanh.

- Đánh ký hiệu đặc biệt, số đầu mục, chữ cái đầu mục vào đầu dòng các văn bảng được chọn.

- Thiết lập in ấn hồ sơ tài liệu hàng loạt, giãn dòng tự động, ngắt trang vùng ký.

- Lập Sheet mục lục cho các sheet trong bảng tính và tạo liên kết Hyperlink đến các sheet.

- Chèn hình ảnh hàng loạt vào và co giãn vừa ô excel kể cả ô merge.

- Chèn hình ảnh hàng loạt vào comment.

- Trộn dữ liệu từ Excel sang Word nâng cao (tương tự chức năng mail merge trong Word)

- Tạo highlight mặt bằng theo dõi tiến độ công việc thực hiện.

- Tạo danh mục nhật ký công trình từ danh mục hồ sơ nghiệm thu.

- Highlight hàng và cột được chọn để tránh nhầm lẫn khi tra cứu, nhập số liệu

- Hỏi và trả lời với trí tuệ nhân tạo, diễn giải công thức phức tạp.
  
- Quản lý và gửi email hàng loạt.

- Gộp, tách file PDF.

- Trích xuất ảnh từ file Excel, thay đổi kích thước hình ảnh hàng loạt.

- Trích xuất văn bản từ ảnh và file PDF bằng AI.

- Và một số hàm mở rộng có gợi ý cách dùng như của Microsoft: chuyển tiền số thành chữ, diễn giải khối lượng, chèn hình tự động khi tên vùng tham chiếu thay đổi, các hàm tương tự office 365,....

Phím tắt:
- F3            :  Xem vị trí các các ô có trong công thức
  
- F8            :  Sao chép định dạng ô

- F6            :  Sao chép dữ liệu bỏ qua dòng ẩn
  
- Shift + F1    :  Trả lời cho lệnh tại ô đang chọn, kết quả hiển thị ô liền kề bên dưới (cần nhập key API Gemini trước)
  
- Ctrl+j        :  Canh lề trái, giữa, phải
  
- Ctrl+m        :  Canh lề trên, giữa, dưới
  
- Ctrl+Shift+a  :  Tô màu vàng cho vùng được chọn
  
- Ctrl+Shift+c  :  Chuyển văn bản sang in thường, in hoa, in hoa đầu từ, in hoa đầu câu
  
- Ctrl+Shift+s  :  Chuyển văn bản in hoa, in thường thông minh (cần nhập key API Gemini trước)
  
- Ctrl+Shift+d  :  Di chuyển hình được chọn đến vừa khít vùng chọn
  
- Ctrl+Shift+v  :  Dịch văn bản vùng được chọn sang Tiếng Việt
  
- Ctrl+Shift+e  :  Dịch văn bản vùng được chọn sang Tiếng Anh
  
- Ctrl+Shift+m  :  Tự động giãn dòng các vùng Merge cell trong vùng được chọn
  
- Ctrl+Shift+t  :  Copy, Paste nhiều vùng chọn khác nhau
  
- Ctrl+Shift+i  :  In nghiêng văn bản sau ký tự ngăn cách
  
- Ctrl+Shift+r  :  Đảo ngược văn bản sau ký tự ngăn cách
  
- Ctrl+Shift+x  :  Copy và Paste bỏ qua vùng ẩn, áp dụng được cho ô merge có kích thước của vùng nguồn và đích khác nhau
  
**Cập nhật: 2.7.0 - 22/08/2025**
- Cải tiến lệnh Filenames List để tùy biến liệt kê tên file.
- Cải tiến lệnh Rename Files để đổi tên file ở các cột bất kỳ .
- Cải tiến nhóm lệnh Bullets để áp dụng được cho các ô có nhiều đoạn văn khác nhau.
- Cải tiến nhóm lệnh Numbering List để áp dụng được cho các ô có nhiều đoạn văn khác nhau.
- Cải tiến tính năng tự động thông báo khi có bản cập nhật mới.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.9 - 15/07/2025**
- Cải tiến lệnh Assign Direct Reference.
- Cải tiến lệnh View Image, bổ sung thiết lập tỉ lệ thu phóng (Enlarge ratio) trong Preferences.
- Cải tiến hiển thị đúng dấu thập phân trong Preferences khi hệ thống dùng dấu ","
- Bổ sung tính năng tự động thông báo có bản cập nhật mới khi phát hành các phiên bản sau 2.6.9
- Thêm lệnh Formula Error Checking để bật tắt chế độ kiểm tra công thức.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.8 - 18/06/2025**
- Thêm lệnh View Image để xem hình ảnh trong sheet đang chọn.
- Thêm lệnh Convert Word To PDF để chuyển file Word sang file PDF.
- Thêm lệnh Rename Sheets để đổi tên Sheet của Workbook hiện hành.
- Thêm lệnh Sheet Auto Number để đánh số trước tên Sheet của Workbook hiện hành.
- Thêm lệnh Create Hyperlink để tạo Hyperlink cho ô dựa vào nội dung trong ô của vùng được chọn.
- Cải tiến lệnh AutoGroup.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.7 - 16/05/2025**
- Thêm lệnh Screen Clip OCR để xuất văn bản từ ảnh chụp màn hình.
- Thêm lệnh Transcribe Audio để chuyển âm thanh sang văn bản.
- Thêm biểu mẫu Template Construction Diary cho menu lệnh Template
- Thêm hàm dvdTextSplitItem để lấy phần tử thứ n của một chuỗi sau khi đã tách theo dấu phân cách.
- Cải tiến một số lệnh.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.6 - 05/05/2025**
- Thêm lệnh OCR để chuyển hình ảnh, PDF sang văn bản.
- Thêm lệnh AI Solver để giải đề thi trắc nghiệm bằng AI Assistant.
- Cải tiến lệnh Regular Italic.
- Cải tiến lệnh Text Reverse.
- Cải tiến lệnh Translate.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.5 - 04/04/2025**
- Cải tiến lệnh Export Category.
- Cải tiến lệnh Translate.

**Cập nhật: 2.6.4 - 12/03/2025**
- Thêm lệnh Reset Note để Chỉnh lại kích thước và vị trí Note.
- Thêm lệnh Export Category để Tạo danh mục hồ sơ từ danh mục hồ sơ nghiệm thu và thi công.
- Thêm lệnh Hide Empty Rows để Ẩn dòng trống trong vùng được chọn.
- Cải tiến lệnh Translate.
- Cập nhật các model AI mới nhất.

**Cập nhật: 2.6.3 - 08/02/2025**
- Đổi tên lệnh Set Page Margins A4 thành Set Page Margins.
- Thêm chức năng lưu thiết đặt lề trang giấy để dùng cho lệnh Set Page Margins.
- Cải tiến hàm dvdExplain và dvdExplainE để diễn giải công thức (hỗ trợ hàm Product).
- Thêm thiết đặt cộng thêm độ rộng dòng cho các ô gộp (Merge cell).

**Cập nhật: 2.6.2 - 24/01/2025**
- Thêm chức năng Auto Click ghi và chạy lại hành động click chuột trái trên mà hình.
- Hỗ trợ đăng ký sử dụng tiện ích vĩnh viễn theo máy.
- Cải tiến và tối ưu.

**Cập nhật: 2.6.1 - 20/11/2024**
- Bổ sung chức năng tiếng Anh cho lệnh dvdLDate, dvdLTime, dvdLTimeDate.
- Cải tiến và tối ưu hiệu suất.

**Cập nhật: 2.6.0 - 01/11/2024**
- Cải tiến lệnh Translate để tăng tốc độ dịch văn bản.
- Phiên bản này yêu cầu nhập lại key đã cung cấp trước đó.

**Cập nhật: 2.5.9a - 29/10/2024**
- Sửa lỗi lệnh Batch Print khi in ra máy in không giãn dòng ô merge.
- Sửa lỗi lệnh Extract PDF Pages không hoạt động khi nhập trang cần tách dạng: 2-3.

**Cập nhật: 2.5.9 - 06/10/2024**
- Thêm chức năng dịch bằng ChatGPT cho lệnh Translate.
- Cập nhật các model AI của ChatGPT và Gemini trong lệnh Preferences
- Thêm chức năng đặt mật khẩu tệp PDF đính kèm email cho lệnh Send Emails.
- Thêm biểu mẫu Template Record Material cho menu lệnh Template
- Cải tiến lệnh Copy Paste Visible
- Thêm lệnh Checkbox, V Checkbox, X Checkbox để chèn ký hiệu hộp chọn cho vùng được chọn.

**Cập nhật: 2.5.8 - 22/09/2024**
- Cải tiến lệnh Selected Sheets To PDF để xuất được cho file có tên tiếng Việt và folder có tên tiếng Việt hoặc đồng bộ đám mây.
- Cải tiến lệnh Rename Files cho tên file hoặc đường dẫn folder là Tiếng Việt.
- Cải tiến lệnh Batch Print để in được cho các sheet đặt tên Tiếng Việt.
- Thêm chức năng gộp file cho lệnh Batch Print khi xuất ra PDF (không hỗ trợ Tiếng Việt có dấu).
- Cải tiến lệnh SheetList để tạo mục lục link được các sheet có tên Tiếng Việt
- Thêm lệnh UnDiacritics Vietnamese Sheets để xóa bỏ dấu Tiếng Việt và ký tự trắng tên sheet.
- Thêm lệnh UnDiacritics Vietnamese Folders để xóa dấu Tiếng Việt tên Folder và tất cả SubFolder.
- Thêm lệnh Send Emails để gửi email hàng loạt bằng outlook từ excel.
- Thêm menu lệnh Template để chèn các biểu mẫu có sẵn.
- Thay đổi phím tắt F9 (hỏi và trả lời bằng AI) thành Shift + F1 để tránh việc nhấn nhầm gây mất dữ liệu ô bên dưới.
- Cải tiến lệnh Merge to Word để khắc phục thừa ký tự trắng và canh lề sai.

**Cập nhật: 2.5.7 - 05/09/2024**
- Bổ sung lệnh Export Images để Trích xuất các hình ảnh từ file excel ra thành các file hình ảnh.
- Bổ sung lệnh Resize Images để Thay đổi kích thước ảnh trong thư mục hàng loạt.
- Bổ sung lệnh Lock Cell References để Thêm hoặc xóa khóa tham chiếu ô.
- Cải tiến lệnh Translate để khắc phục lỗi "Error : Resource has been exhausted (e.g. check quota)." khi dùng Gemini.
- Cải tiến lệnh Delete Error Name để khắc phục lỗi treo máy khi click chuột.
- Cải tiến lệnh AutoFit Merge để khắc phục lỗi treo máy khi click chuột.
- Cải tiến lệnh Add or Remove Rounding Functions cho các trường hợp các hàm round nằm trong chuỗi text.
- Cải tiến lệnh Assign Direct Reference lỗi khi đặt tên sheet có dấu cách " ", tiếng việt có dấu.

**Cập nhật: 2.5.6 - 20/08/2024**
- Bổ sung lệnh Add or Remove Rounding Functions để Thêm hoặc xóa các hàm làm tròn cho công thức.
- Cải tiến lệnh Translate để Dịch văn bản của vùng được chọn sang ngôn ngữ khác để tạo song ngữ.
- Cải tiến lệnh Add Text để Thêm văn bản nối vào văn bản hiện tại cho vùng được chọn.
- Cải tiến lệnh AutoGroup để Tự động Group các dòng hoặc cột trống, ẩn dựa vào vùng tham chiếu.

**Cập nhật: 2.5.5 - 11/08/2024**
- Bổ sung lệnh Roman Serial để Đánh đầu mục bằng bảng chữ số la mã I, II, III.
- Bổ sung lệnh Upper Letters Serial để Đánh đầu mục bằng bảng chữ cái in hoa A, B, C
- Bổ sung lệnh Lower Letters Serial để Đánh đầu mục bằng bảng chữ cái in thường a, b, c
- Thay đổi lệnh Letters List thành Lower Letters List để Tạo danh sách các chữ cái dạng a. b. c.
- Bổ sung lệnh Upper Letters List để Tạo danh sách các chữ cái dạng A. B. C.
- Bổ sung lệnh AI Grammarly để Sửa lỗi chính tả, ngữ pháp bằng AI Assistant.

**Cập nhật: 2.5.4 - 04/08/2024**
- Bổ sung lệnh Extract PDF Pages để Tách các trang của file PDF thành 1 file mới.
- Bổ sung lệnh Multilevel Numbering List để Tạo danh sách các số đa cấp dạng 1. 1.1. 1.1.1.
- Bổ sung lệnh Selected Sheets To PDF để Xuất các sheet được chọn ra tệp PDF lưu cùng vị trí file excel (F10).
- Bổ sung lệnh Hide Unselected Sheets để Ẩn tất cả các sheet không được chọn.
- Bổ sung lệnh Unhide All Sheets để Hiện tất cả các sheet đã ẩn.

**Cập nhật: 2.5.3 - 27/07/2024**
- Bổ sung lệnh Combine Selected PDFs để nối nhiều file PDF thành 1 file PDF.
- Bổ sung lệnh Print Pages để In trang chẵn, trang lẻ hoặc các trang bất kỳ cho Sheet được chọn.
- Bổ sung lệnh Set Page Margins A4 để Thiết lập canh lề khổ giấy A4 cho các sheet được chọn.
- Bổ sung lệnh Assign Direct Reference để Gán tham chiếu ô trực tiếp dựa trên giá trị tra cứu được tìm thấy.
- Cải tiến lệnh Print Mirror Margins để Xuất file PDF dùng cho in 2 mặt đối xứng dạng sách cho sheet được chọn.

**Cập nhật: 2.5.2 - 21/07/2024**
- Bổ sung lệnh Print Mirror Margins để xuất sheet hiện tại ra PDF dùng cho in 2 mặt dạng in sách.
- Bổ sung lệnh Get Shape Positions để liệt kê vị trí của Shape được chọn tên.
- Bổ sung hàm dvdMoveShape Di chuyển một hình vẽ, hình ảnh, textbox đến vị trí mới trên trang tính.
- Tách lệnh Rename Shapes thành 2 lệnh List Shapes và Rename Shapes.

**Cập nhật: 2.5.1 - 15/07/2024**
- Bổ sung tùy chọn làm tròn giá trị cần diễn giải cho hàm dvdExplain, dvdExplainE
- Bổ sung hàm dvdUniqueV để liệt kê giá trị 1 lần trong vùng được chọn. Tương tự hàm Unique trong excel 365.
- Bổ sung hàm dvdMVLookup để tìm kiếm nhiều giá trị trong một cột và trả về các giá trị tương ứng từ một cột khác.
- Bổ sung hàm dvdMCLookup để tìm kiếm nhiều giá trị trong một cột và trả về các giá trị tương ứng từ nhiều cột khác. Tương tự hàm Filter trong excel 365.
- Bổ sung lệnh Import Emails để lấy nội dung email từ phần mềm Outlook đưa vào excel.
- Thay chức năng phím tắt F4 bằng phím F8 để sao chép định dạng
- Cải thiện lệnh Batch print không giãn dòng ô merge.

**Cập nhật: 2.5.0 - 22/06/2024**
- Bổ sung tính năng diễn giải cho hàm mở rộng. Lưu ý: Mở file đã lưu trước để sử dụng tính năng này.
- Bổ sung lệnh UnDiacritics Vietnamese để xóa dấu văn bản Tiếng Việt.
- Bổ sung lệnh Diacritics Vietnamese để chuyển văn bản Tiếng Việt không dấu sang có dấu.
- Bổ sung lệnh Unprotect Sheet để xóa mật khẩu bảo vệ Sheet bị quên.
- Đổi tên hàm dvdUnMarkVi sang dvdUnDiacriticsVi
- Đổi tên hàm dvdDLookup sang dvdTableLookup
- Lệnh đánh số thứ tự Numbering bỏ cửa sổ nhập số kết thúc.

**Cập nhật: 2.4.9_a - 19/06/2024**
- Sửa lỗi lệnh Batch Print không AutoFilter khi xuất PDF.
- Đổi tên lệnh Format Text To Number sang lệnh Format Text To Auto để áp dụng cho chuỗi số, ngày giờ.
- Cải thiện lệnh Copy Paste Visible khi muốn copy bằng cách gán công thức.

**Cập nhật: 2.4.9 - 19/06/2024**
- Cải thiện hiệu suất các lệnh Insert Multiple Rows và Insert Multiple Columns.
- Bổ sung tùy chọn màu sắc cho lệnh Reading Layout.

**Cập nhật: 2.4.8 - 03/06/2024**
- Bổ sung lệnh Format Text To Number để Chuyển các số định dạng văn bản sang định dạng số.
- Bổ sung lệnh Merge Similar Cells để Gộp các ô liền kề có dữ liệu giống nhau hoặc các ô trống vào ô có dữ liệu.
- Bổ sung lệnh Merge Rows để Gộp các dòng dữ liệu có dữ liệu ở cột đầu tiên trong vùng chọn giống nhau vào 1 dòng.
- Bổ sung lệnh Delete Error Name để Xóa các Name Range bị lỗi.
- Bổ sung lệnh Delete Cell Styles để Xóa các Cell Styles người dùng tạo.
- Cải thiện code lệnh AutoFit Merge và lệnh Batch Print. 

**Cập nhật: 2.4.7 - 27/05/2024**
Bổ sung nhóm lệnh Text
- Thay đổi lệnh AddExtension thành Add Text với cải tiến thêm text phía trước hoặc sau văn bản, tùy chọn bỏ qua ô trống.
- Bổ sung lệnh Remove Extra Spaces để Xóa ký tự trắng thừa của văn bản trong vùng được chọn.
- Bổ sung lệnh Superscript Numbers Units để Chuyển các số của đơn vị hệ mét có trong vùng được chọn lên phía trên.
- Bổ sung lệnh Chemical Formulas để Chuyển các số của công thức hóa học có trong vùng được chọn xuống phía dưới.
- Cải thiện tốc độ lệnh Numbering. 

**Cập nhật: 2.4.6 - 04/05/2024**
- Cải tiến lệnh Translate để dịch cho các Font chữ thuộc các bảng mã khác nhau, giao diện chọn ngôn ngữ trực quan.
- Cải tiến các lệnh AI sử dụng dữ liệu mới nhất (đến tháng 11/2023).
- Các bổ sung nhỏ cho lệnh Regular Italic và Text Reverse.

**Cập nhật: 2.4.5 - 20/04/2024**
Bổ sung tính năng tạo link cho nút lệnh Folder Tree
Bổ sung tính năng trộn file từ Excel sang Word.
Bổ sung tính năng nhập dữ liệu nhanh từ danh sách (nằm trong Tab Data)

**Cập nhật: 2.4.4 - 09/04/2024**
Cải tiến lệnh Insert Picture, Inserture By Nane và hàm DVDPic khi chèn hình không xem được hình ảnh ở máy tính khác.
Bổ sung tính năng dịch văn bản tạo song ngữ bằng AI Gemini.
Bổ sung tính năng Chat With AI.
Bổ sung hàm ẩn dòng theo điều kiện DVDAutoHide.
_Lưu ý: Trong phiên bản này có tổ chức lại cách đặt tên biến nên cần nhập key kích hoạt đã cung cấp trước khi sử dụng hết tính năng._

**Cập nhật: 2.4.3 - 25/03/2024**
Mang lại tính năng AI Assistant lên thanh Ribbon.
Bổ sung nút lệnh AI Explain để diễn giải công thức của Microsoft Excel.
Bổ sung nút lệnh Roman List để tạo danh sách các số la mã dạng I. II. III.

**Cập nhật: 2.4.2 - 05/03/2024**
- Mang lại tính năng SmartCase để viết hoa viết thường thông minh bằng Gemini AI
- Bổ sung phím tắt F9 để trả lời câu hỏi bằng Gemini AI. Kết quả hiển thị ngay ô bên dưới câu hỏi.
- Thay đổi thuật toán giãn dòng các vùng Merge Cell
- Bổ sung nút lệnh Preferences để tùy chọn mặc định cho tiện ích **(chưa hoàn thiện)**

**Cập nhật: 2.4.1 - 18/02/2024**
- Sửa lỗi lệnh Translate không dịch một số ngôn ngữ.
- Bổ sung lệnh Download Tool để tải bản cài đặt về máy tính.
- Sử dụng phím tắt Ctrl+Shift+C để chuyển đổi chữ in hoa, in thường thay cho các lệnh Ctrl+Shift+S, Ctrl+Shift+U, Ctrl+Shift+L

**Cập nhật: 2.4.0 - 02/02/2024**
- Cải thiện hiệu suất, sửa lỗi hàm DVDPic, và lỗi không nhấn phím tắt mở rộng của Addin.
- Khắc phục các lệnh bị Windows security báo nhầm virus.

**Cập nhật: 2.3.9 - 02/02/2024**
- Thêm tính năng cập nhật phiên bản mới.
- Cải thiện hiệu suất, sửa lỗi.

**Cập nhật: 2.3.8 - 25/01/2024**
Cải thiện hiệu suất, sửa lỗi không tương thích với office 64 bit

**Cập nhật: 2.3.7 - 22/12/2023**

- Bổ sung các hàm tạo mã QR code DVDQR
- Bổ sung lệnh liệt kê tên Folder theo cây thư mục.
- Bổ sung lệnh tạo cây thư mục Folder từ danh sách trên excel

**Cập nhật: 2.3.6 - 22/11/2023**
- Bổ sung các hàm về thời gian DVDLDate, DVDLTime, DVDLTimeDate
- Bổ sung phím tắt F4 để gọi lệnh sao chép định dạng Format Painter
- Bổ sung phím tắt Ctrl+j để canh lề văn bản trái, giữa, phải, đều hai bên
- Bổ sung phím tắt Ctrl+m để canh lề văn bản trên, giữa, dưới, đều trên dưới
- Cải tiến lệnh Cons Diary - xuất danh mục nhật ký thi công từ danh mục hồ sơ

--------------------------------------------------------------------------------------
Tác giả: Đặng Văn Đằng

Tính năng:

Định dạng văn bản: in hoa, thường, viết hoa chữ cái đầu từ mỗi câu, viết hoa chữ cái đầu mỗi từ, viết hoa viết thường thông minh.

Đánh số thứ tự theo hàng, theo cột, theo bước nhảy của số.

Đảo văn bản qua ký tự ngăn cách (vd: dùng đảo văn bản song ngữ).

Đổi bảng mã giữa Unicode <-> VNI Windows <-> TCVN3 không bị lỗi font.

In nghiêng văn bản sau ký tự ngăn cách (vd: dùng in nghiêng văn bản song ngữ).

Dịch văn bản tạo song ngữ theo google dịch, chatGPT, google Bard.

Liệt kê file trong folder, liệt kê file theo cây thư mục folder.

Liệt kê Folder theo cây thư mục và tạo cây thư mục Folder theo danh sách trên excel.

Xóa, chèn nhiều dòng cột.

Xóa name range, cell styles rác giúp giảm dung lượng và mở file nhanh.

Đánh ký hiệu đặc biệt, số đầu mục, chữ cái đầu mục vào đầu dòng các văn bảng được chọn.

Thiết lập in ấn hồ sơ tài liệu hàng loạt, giãn dòng tự động, ngắt trang vùng ký.

Lập Sheet mục lục cho các sheet trong bảng tính và tạo liên kết Hyperlink đến các sheet.

Chèn hình ảnh hàng loạt vào và co giãn vừa ô excel kể cả ô merge.

Tạo highlight mặt bằng theo dõi tiến độ công việc thực hiện.

Quản lý và gửi email hàng loạt.

Gộp, tách file PDF.

Trích xuất ảnh từ file Excel, thay đổi kích thước hình ảnh hàng loạt.

Trộn file Excel sang Word nâng cao.

Khóa, gán tham chiếu nhanh các ô chứa dữ liệu tìm thấy.

Và một số hàm mở rộng: chuyển tiền số thành chữ, diễn giải khối lượng, chèn hình tự động khi tên vùng tham chiếu thay đổi, các hàm tương tự office 365,....
