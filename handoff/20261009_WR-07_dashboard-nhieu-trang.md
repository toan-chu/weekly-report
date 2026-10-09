Mã:         WR-07
Mạch:       WR — báo cáo công nợ phải thu hằng tuần
Nối tiếp:   WR-05, WR-06
Giao cho:   claude-logisticist   [COWORK-EXEC]
Trạng thái: chờ-nghiệm-thu
Vòng:       1/2

## ĐỀ BÀI

**Mục tiêu nghiệp vụ** (Chairman, 2026-10-09)
Mỗi tab trên thanh đầu trang là một màn riêng, không cuộn lên xuống. Bấm tab nào hiện tab đó.
Vào chi tiết (một Sales, một khách) thì mới hiện rõ phần của người đó, và các tab khác chỉ còn
phần liên quan (kiểu dashboard nhiều trang + drill-through như Looker Studio / Power BI).
Kèm theo: Task Scheduler chạy mà không bật cửa sổ CMD.

**Ngoài phạm vi**
- Nợ vendor theo job/loại phí; màu phải trả tăng — chờ FIN.

**Kịch bản nghiệm thu**
Chairman: test ngay trên máy Chairman.

## TODO
- [x] 6 trang: Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả · Dòng tiền; viên tím đứng ở trang đang mở
- [x] Trang chi tiết khách: tuổi nợ, ai phụ trách (bấm → trang Sales), job + trạng thái Misa, hoá đơn, S1–S6, Đối chiếu
- [x] Trang riêng Sales: phần phụ trách, quá hạn, nợ xấu, khách (bấm → trang khách), job, khách nợ xấu
- [x] Bộ lọc nằm trên thanh địa chỉ: đi theo qua các tab, nút Back của trình duyệt quay lại được
- [x] Tổng quan thêm "Cần chú ý" và "Theo Sales" để bấm xuyên
- [x] In / Lưu PDF in đủ mọi trang
- [x] Lịch chạy nền gọi thẳng pythonw.exe; máy cũ tự đổi lịch một lần khi runner chạy

## AUDIT

```
$ python dashboard/build.py                 Credit_Dashboard.html (1.34 MB)
$ python -m pytest -o addopts="" -q         55 passed, 20 skipped
Chromium headless với báo cáo W40 (bản WR-06): 0 lỗi console.
  Sales → bấm Trần Văn Hiếu → #sales?s=…; sang tab Nợ xấu chỉ còn 3 khách của Hiếu; Back quay lại đúng.
  Bấm khách trong trang Sales → trang khách LCC UPP (9 job, Đối chiếu: Khớp). In PDF: 11 trang.
```
[RUNTIME CHƯA KIỂM] Đổi lịch chạy ẩn (schtasks /change) trên Windows của Chairman.

## HISTORY

### 2026-10-09
Dựng theo yêu cầu. Lỗi thấy khi soi ảnh chụp: cột "Việc cần làm" bị canh phải (bảng HTML coi cột
cuối là số tiền) → canh theo tên cột. Thuộc tính hidden bị CSS display:flex đè → thêm luật chung.
Failure: không. Lessons: thanh địa chỉ giữ trạng thái trang là cách rẻ nhất để có nút Back.

### 2026-10-09 (cuối ngày)
FIN chốt Payable: "làm y chang mẫu của em… chỉ cần lấy số tổng, không phân tích, không cần tô".
Sheet Payable còn 5 cột (No · Code · Vendor's Name · Current · Total), chỉ vendor còn số dư; tiền đã
trả vendor giữ ở Cash Flow. Dashboard tab Phải trả: bỏ thẻ trạng thái, đọc tiền đã trả từ Cash Flow.
Review bỏ bảng "phải trả nhập bổ sung". Giả lập máy FIN bằng gói handoff/docs/sample_w40: 6 file W40
+ W39 final + file ghép → chạy 1 lần ra W40 (70 khách, 14.251.295.535; Đối chiếu 57/6/2/5),
nút 2: 68 mã đã ghép, 0 khách mới, 0 dòng sửa, Tham_chieu_job_no_cu có 1 hoá đơn (Minh Khôi 32tr).
