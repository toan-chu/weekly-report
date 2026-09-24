# SPEC — Luật điền sheet Receivable

Chốt ngày 2026-09-23 giữa Chairman, kế toán phụ trách công nợ và Claude.
Cập nhật 2026-09-23 chiều: thêm mục 3b sau khi kế toán trả lời đợt 2.
Rút ra từ 3 tuần dữ liệu thật W36, W37, W38 trong `fixtures/`.

## Đầu vào

| Nguồn | Dùng để làm gì |
|---|---|
| Tổng hợp công nợ phải thu | Danh sách khách và Total. Đây là số tiền đúng. |
| Phân tích công nợ theo tuổi nợ | Chia Total vào các nhóm tuổi. |
| Báo cáo tuần trước | Credit Term, Salesman, số job, Lý do chưa thu hồi. |
| Sổ ghi của tool | Cách chia tuổi tuần trước, ngày mỗi khoản xuất hiện lần đầu. |

Hai file Misa phải cùng kỳ: ngày "đến ngày" của file Tuổi nợ = ngày cuối kỳ
của file Tổng hợp + 1 ngày. Kiểm ở cả 3 tuần đều đúng quy luật này.

## Vì sao không lấy Total từ file Tuổi nợ

File Tuổi nợ chỉ đúng khi hoá đơn có ghi hạn thanh toán và phiếu thu đã được
gắn vào đúng hoá đơn. Thực tế ở Trustana:
- Hoá đơn cũ không ghi hạn thanh toán — W38 có 56,2 tỷ nằm ở ô "Không có hạn nợ".
- Khách trả bằng loại tiền khác loại tiền trên hoá đơn, hoặc trả thừa để bù trừ
  cho lô sau, kế toán không gắn phiếu thu được — hoá đơn vẫn nằm đó như chưa trả.

Hậu quả: W38 file Tuổi nợ cộng ra 73,0 tỷ trong khi thực nợ 10,0 tỷ.

## Luật điền

1. **Danh sách**: khách có số dư cuối kỳ bên Nợ lớn hơn 0 trong file Tổng hợp.
   Khách đã trả hết tự rơi ra. Khách trả trước (dư bên Có) không đưa vào.
2. **Total** = số dư cuối kỳ bên Nợ.
3. **Sáu cột tuổi nợ**: Current / 1-30 / 31-60 / 61-90 / 91-120 / 120+.
   Từ W39 tách riêng 91-120 và 120+ theo yêu cầu của ban lãnh đạo; trước đó
   kế toán gộp hết vào cột 61-90.
   - Current = toàn bộ nợ chưa tới hạn, không quan tâm còn mấy ngày nữa tới hạn.
   - Nhóm A — file Tuổi nợ khớp Total và mọi hoá đơn có hạn: chép thẳng.
     W38 có 31/70 khách thuộc nhóm này.
   - Nhóm còn lại: lấy cách chia của tuần trước trong sổ ghi rồi điều chỉnh
     · nợ tăng thêm → cho vào Current (hoá đơn mới thì chưa tới hạn)
     · nợ giảm đi → trừ vào nhóm tuổi cũ nhất trước
     · đủ ngày tuổi → tự chuyển Current sang 1-30, rồi 31-60, theo số ngày
       thật kể từ ngày khoản đó xuất hiện lần đầu, cộng Credit Term
   - Khách mới chưa từng có trong sổ ghi: toàn bộ vào Current, Credit Term
     mặc định 30 ngày, đánh dấu để soát.
3b. **Kế toán khai đè** hai thứ mà file Misa không có, khai một lần dùng mãi:
   credit term mới (khách vừa ký hợp đồng) và ngoại lệ nghiệp vụ. Xem
   `NOTE-ke-toan-tra-loi-20260923.md`.
4. **Ngoại lệ nghiệp vụ**: có sự kiện chỉ kế toán biết (VD lô tàu chìm, xuất
   lại debit note nên công nợ tính từ ngày mới). Ghi một lần vào sheet Ngoại lệ,
   tool áp dụng cho tới khi khoản đó thu xong. W38 có 3 ca: Việt Nga, BSF (lô tàu chìm, xuất lại debit note) và AHC Logistics
   (đang chờ ký biên bản bù trừ công nợ hai chiều nên để ở Current).
4b. **Khách bị loại khỏi bảng**: có khoản Misa ghi nợ nhưng thực chất khách
   không nợ (VD nợ ảo do lỗi xuất hoá đơn). Kế toán khai `--bo-qua` một lần;
   tool bỏ khỏi bảng nhưng vẫn liệt kê ở sheet Cần xem lại kèm lý do, để không
   ai quên. Hết chuyện thì `--nhan-lai`.
5. **Thứ tự dòng**: nợ lâu nhất lên trên. Cùng nhóm tuổi thì số tiền lớn lên trên.
6. **Tự kiểm tra**: tổng 6 cột của một dòng phải bằng Total dòng đó; tổng cả
   bảng phải bằng tổng file Tổng hợp. Dòng nào không thoả, hoặc tool tự thấy
   không chắc, thì vào sheet "Cần xem lại".
7. **Ngày chốt số**: thứ Bảy. Tuần tính từ thứ Bảy tới thứ Sáu.

## Số trên Misa đổi theo thời điểm xuất

Cùng một tuần, xuất Misa hôm nay và xuất lại tuần sau có thể ra số khác nhau:
ghi nhận lùi, hoá đơn thay thế, sửa tên pháp nhân trên hoá đơn. Vì vậy:
- Báo cáo ghi lại tên và thời điểm sửa của hai file nguồn.
- Chỉ đối chiếu báo cáo máy với bản làm tay khi cả hai dùng CÙNG một cặp file.

## Đứt quãng và chen hàng

Tuổi nợ tính theo ngày thật, không theo số lần chạy. Nghỉ vài tuần rồi chạy
lại vẫn ra tuổi đúng cho khách cũ; nợ mới phát sinh trong lúc nghỉ bị đoán trẻ
hơn thực tế, các dòng này được đánh dấu. Tìm lại được file của tuần bỏ lỡ thì
thả vào, tool xử lý theo thứ tự ngày rồi sinh lại các tuần sau.
