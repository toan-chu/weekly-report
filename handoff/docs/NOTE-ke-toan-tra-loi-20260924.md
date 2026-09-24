# NOTE — Kế toán trả lời đợt 3, 2026-09-24

Về 3 nhóm chênh khi đối chiếu tool với bản tay W33, W34, W35.

## 1. Bốn khách có dư Nợ trong Misa mà bảng tay không có

- **MN Shipping Hải Phòng / Vận Tải Biển Minh Nguyên** (183.514.584): ban đầu
  xuất hoá đơn cho Vận Tải Minh Nguyên, sau sửa lại xuất cho MN Shipping. Cùng
  một khoản, khác pháp nhân trên hoá đơn. Tool lấy theo Misa là đúng.
- **Rustam Yadu** (74.491.924) và **Xiamen Trans-Europeasia** (9.010.411): lúc
  kế toán làm báo cáo, một kế toán viên khác chưa ghi nhận hoá đơn; sau mới ghi nhận lùi.
- **Optimalog** (2.115.321): khách KHÔNG thực nợ. Do mình xuất hoá đơn tiền cọc
  sai nên thừa ra phần thuế phải thu thêm. Kế toán chủ động loại khỏi bảng để
  clear, sẽ thu thêm trong lô sau.
  → Xử lý: khai `--bo-qua`, tool loại khỏi bảng nhưng vẫn liệt kê ở sheet
  Cần xem lại kèm lý do.

## 2. STA Logistic Center (lệch 235.520) và ACS Time Critical (lệch 6.318.221)

Hai nhà này mới làm hoá đơn thay thế cho hoá đơn phát sinh ở ba tuần đó, nên
công nợ hiện tại khác với lúc kế toán làm báo cáo.

## 3. Kết luận quan trọng cho việc nghiệm thu

Cả 5 chênh lệch đều KHÔNG phải lỗi tool và cũng không phải lỗi kế toán. Chúng
đến từ một chuyện duy nhất: **file Misa xuất hôm nay khác với file Misa lúc kế
toán làm báo cáo** — ghi nhận lùi, hoá đơn thay thế, sửa tên pháp nhân trên
hoá đơn.

Hệ quả:
- Không thể đối chiếu một báo cáo cũ với một lần xuất Misa mới. Muốn so thì
  kế toán và tool phải dùng **cùng một cặp file Misa**.
- Báo cáo từ nay ghi lại tên và thời điểm sửa của hai file nguồn ở đầu sheet
  Cần xem lại, để sau này biết bản đó dựa trên lần xuất nào.
