Cập nhật:  2026-10-09 UTC+7 — bởi claude-logisticist
Chỉ huy:   claude   [COWORK-EXEC]

## Việc đang mở
   WR-04   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2
           Full credit report: 4 file → Summary, Receivables, AR Risk, Payable, Cash Flow,
           Methodology, Review. 44 phép thử đạt. W40 thật: 61/70 khớp bản người, 97,8% tiền
           đúng nhóm. Chờ Chairman chạy trên máy mình rồi máy FIN.

   WR-03   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 0/2
           Logic SMS + bản final; phần tuổi nợ đã được WR-04 thay. Còn 3 dòng Windows chờ ký.

   WR-02   giao cho chairman   chờ-đề-bài   — có thể đóng huỷ, WR-04 đã so bản người.
   WR-01   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2 — 3 dòng Windows chờ ký.

   WR-05   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2
           Excel có dòng con job/Sales, sheet Invoices, trang bìa; Credit_Dashboard.html ở gốc
           workspace (kéo báo cáo vào, In / Lưu PDF). 48 phép thử đạt.

   WR-06   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2
           Job chưa thanh toán từ Misa (Bán hàng + Sổ chi tiết 131) + cột Đối chiếu tô màu.
           55 phép thử đạt. W40: 57 khớp, 6 đỏ, 2 cam, 5 khó đòi. Workspace test đã để sẵn 6 file.

   WR-07   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2
           Dashboard nhiều trang + bấm xuyên (khách, Sales); lịch chạy nền không bật CMD.

## Hôm nay (09/10)
   FIN thêm 2 file xuất từ đầu năm; luật "chỉ Đã thanh toán mới là đã trả"; cần lệnh báo lệch.
   Làm WR-06. Chairman mở được dashboard trên Windows. Sửa file ghép khách test: TTND VIETNGA
   → KH00416, KH00361 (lô tàu chìm) = không có SMS. Chiều: làm WR-07 (dashboard nhiều trang, chạy ẩn). FIN audit WR-06 xong, đồng ý. FIN chốt
   Payable = y chang mẫu, chỉ số tổng, không tô đỏ. Đã giả lập máy FIN với gói sample_w40: chạy sạch.

## Hôm nay (08/10, chiều)
   Chairman chạy WR-04 trên Windows: đúng số. CEO chốt tên Sales đầy đủ. Thêm lý do tự động
   "[TỰ ĐỘNG]". Làm WR-05: dòng con job trong Excel, sheet Invoices, dashboard HTML.

## Hôm nay (08/10, sáng)
   FIN chốt: Chi tiết 131 xuất từ 01/01/2022 mỗi tuần; quá hạn tính từ hôm sau ngày đến hạn;
   Payable = toàn bộ Misa (bản 114 vendor lệch vì hoá đơn nhập sau). Dựng WR-04. Bỏ nút
   3_Chay_ngay. Đổi tên file trong handoff/docs/new theo nhóm 1_INPUT / 2_SAMPLE /
   3_DOI-CHIEU / 4_CU; .gitignore chặn mọi xlsx dưới handoff/docs.

## Hôm 06/10
   WR-03: thêm SMS AR-AP, file ghép khách, bản final làm trí nhớ nợ cũ; 3 nút .bat.

## Vừa đóng
   chưa có

## Đang treo, cần Chairman
   1. Chạy thử WR-06 trên máy Chairman (Task Scheduler → Run), rồi commit, cài máy FIN.
   2. AR Risk dùng lý do của tuần trước; dashboard WR-05 có tự tính lại nhóm khi chị KTT
      sửa lý do không? (đề xuất: có)
   3. Gói handoff/docs/sample_w40 (zip) gửi FIN chạy lại W40; W41 thả 6 file mới.

## Bẫy đã biết
   - Số hoá đơn Misa đánh lại từ đầu mỗi năm: ghép phiếu thu theo số HĐ phải chọn HĐ
     gần nhất trước ngày thu, không lấy HĐ đầu tiên mang số đó.
   - File Tổng hợp và Chi tiết xuất khác giờ thì số dư lệch (W40: Thủy Ngân 308tr, GOLDEN
     GLOBE 2,7tr). Tiền theo Tổng hợp; Review liệt kê.
   - Hoá đơn nhập lùi ngày làm dư đầu kỳ Misa khác dư cuối báo cáo tuần trước: cột
     "Backdated vs last report" ở AR Risk và Payable.
   - File Misa tải về bị Windows thêm hậu tố (21), (22): luôn đọc ngày trong file.
   - Mã khách khác nhau giữa các file Misa nên so khớp theo TÊN đã chuẩn hoá.
   - Trạng thái thanh toán trên file Bán hàng là tại lúc xuất, không phải ngày cuối kỳ (WR-06).
   - Excel thật mới tính công thức; openpyxl không lưu giá trị — báo cáo ghi số, không công thức.

## Báo lỗi giả
   Tháng này: 0
