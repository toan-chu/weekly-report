Cập nhật: 2026-10-28 10:55 UTC+7 - bởi claude

## Việc đang mở
WR-12 - người làm claude - chờ-soát - vòng 1/2 - khách mới tự vào file credit term, ô Mở hợp đồng, khoá tuần đã chốt, đổi tên gọi
WR-11 - người làm claude - chờ-soát - vòng 1/2 - màu tuổi nợ (CEO), Client Master thêm cột, PA1 lý do, khách mới; sửa lỗi WR-09
WR-10 - người làm claude - chờ-soát - vòng 1/2 - đưa repo về handoff-standard v2.0
WR-09 - người làm claude - chờ-soát - vòng 1/2 - sửa theo FIN sau W42; 5/6 lỗi tự soi đã sửa ở WR-11
WR-08 - người làm claude - chờ-soát - vòng 1/2 - công thức KTT, Client Master, sheet Contracts
WR-07 - người làm claude - chờ-soát - vòng 1/2 - dashboard nhiều trang, chạy ẩn
WR-06 - người làm claude - chờ-soát - vòng 1/2 - job từ Misa, cột Đối chiếu
WR-05 - người làm claude - chờ-soát - vòng 1/2 - dòng con job/Sales, dashboard HTML
WR-04 - người làm claude - chờ-soát - vòng 1/2 - full credit report
WR-03 - người làm claude - chờ-soát - vòng 0/2 - SMS AR-AP, ghép khách, bản final
WR-02 - người làm claude - chờ-đề-bài - vòng 0/2 - so bản tay trên cùng bộ file Misa
WR-01 - người làm claude - chờ-soát - vòng 1/2 - dựng tool báo cáo tuần

## Hôm qua
- 20/10: WR-09 sửa theo FIN sau W42; chạy lại W42 khớp bản human 64/68 dòng job, 4 khách tool theo sổ, chờ FIN xác nhận.
- 22/10: commit WR-09 (5aff24e).
- 27/10: người làm tự soi lại WR-09, thấy 2 lỗi đếm "khách chưa có hợp đồng", Methodology chưa sửa, script nâng cấp Client Master bỏ cột lạ. Chưa sửa mã.
- 27/10: WR-10 đưa repo từ chuẩn v1.1 lên v2.0; mọi phiếu chuyển sang khuôn mới, chờ codex soát.
- 28/10: WR-12 bổ sung: file credit term tự lên mẫu mới mỗi lần chạy kể cả không có khách mới; README cho FIN; 87 phép thử đạt. Chờ user commit + triển khai máy FIN (GitHub Desktop Pull).
- 27/10: WR-12 (v0.7.1): khách mới tự vào file credit term (W43 thật: 5 khách, 83 dòng cũ giữ nguyên), ô Mở hợp đồng, khoá tuần đã chốt; 85 phép thử đạt. Tên gọi: file credit term; Client Master = SharePoint List sau này.
- 27/10: WR-11 (v0.7) xong, 81 phép thử đạt; chạy lại W43 thật trong máy ảo: tổng y bản cũ, màu gốc khớp bản Head FIN 61/61. Kịch bản chạy lại W43 trên máy FIN: dùng bản trong WR-12 (thay WR-11).

## Vừa đóng
chưa có

## Đang treo, cần user
0. Kỳ vọng (WR-11): với W43 thật, 58/61 khách bị nâng 1 màu (42 khách chưa ✓ hợp đồng, 26 khách ≥ 2 hoá đơn trả trễ / 6 tháng). Head FIN chốt X, Y, N (N = 2 có thể quá thấp với khách đi hàng nhiều)? Giữ 5 màu hay gộp còn 3 như CEO gợi ý?
1. Kỳ vọng: WR-02 (kế toán làm tay trên đúng bộ file tool đã chạy rồi so) còn cần không? WR-04 → WR-09 đã so W40–W42 với bản human. Không cần thì đóng-huỷ.
2. Kỳ vọng: "Ngày gia hạn chỉ có hiệu lực khi đã tới ngày đó" (thêm ở WR-09, luật R.2) có đúng ý không? Nếu Sales gia hạn sớm thì hạn mới tính từ ngày gia hạn hay từ ngày hết hạn cũ?
3. Đề xuất luật cho RULES (bẫy sống quá một tuần):
   a. Không được ghép phiếu thu với hoá đơn đầu tiên mang số đó; phải chọn hoá đơn cùng số gần nhất trước ngày thu (Misa đánh số lại mỗi năm).
   b. Không được tin tên file Misa (Windows thêm hậu tố (21), (22)); phải đọc ngày bên trong file.
   c. Không được so khớp khách giữa các file Misa theo mã; phải so theo tên đã chuẩn hoá.
   d. Gộp hai câu ở R.2 về thứ tự trừ tiền thu: câu v0.4 "còn lại trừ hoá đơn cũ nhất trước" nay đã bị câu 20/10 (ưu tiên hoá đơn Misa "Đã thanh toán") thay một phần.
4. Luật bị thay khi sinh lại RULES từ v2.0 (xem AUDIT WR-10): mục 1 cũ "Chỉ huy" và giải trừ giới hạn; mục 2 cũ quyền ký nghiệm thu theo vai; mục 3 cũ trần vòng lặp (nay A3); mục 4 cũ "Bốn câu hỏi kiểm tra của user" (nay A10).

## Bẫy đã biết
- Hoá đơn nhập lùi ngày làm dư đầu kỳ Misa khác dư cuối báo cáo tuần trước: xem cột "Backdated vs last report" ở AR Risk và Payable.
- File Tổng hợp và Chi tiết xuất khác giờ thì số dư lệch (W40: Thủy Ngân 308tr); tiền theo Tổng hợp, Review liệt kê.
- File Excel thật trong `handoff/docs/` và `handoff/archive/` bị .gitignore chặn: máy khác không có, phải chép tay.
- Phép thử: 20 phép thử bỏ qua khi thiếu dữ liệu thật trong `tests/fixtures/`; máy ảo Linux cần `pip install --user pytest openpyxl` trước. `tests/conftest.py` hạ ngưỡng số lẻ xuống 0,5đ — phép thử cần ngưỡng thật phải tự đặt lại.

## Báo lỗi giả
Tháng này: 0
