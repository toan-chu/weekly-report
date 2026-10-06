# FIN trả lời về file ghép khách SMS ↔ Misa — 2026-10-06

Nguồn: Chairman chuyển lời FIN (Ranee) trong hội thoại 2026-10-06.
Áp cho: `wr/ghep.py`, phiếu WR-03. Còn đúng tới khi FIN đổi mã khách trên SMS.

1. **Một mã SMS ứng với hai khách Misa thì theo Misa.** SMS chưa đổi mã sang
   "Viện độ bền nhiệt đới" (KH00416) nên vẫn dùng mã cũ "TTND VIETNGA".
   Luật: khi sheet "Ghép khách" và sheet "Misa chưa ghép" mâu thuẫn, sheet
   "Misa chưa ghép" thắng (đó là phía Misa). File kiểm tra vẫn liệt kê để FIN biết.
2. **Mã khách FIN ghi mà Misa không có số dư = khách đã trả, SMS quên gạch.**
   Trên Misa khoản đó tự mất. Tool bỏ số SMS đó, không đưa vào báo cáo, chỉ liệt
   kê ở mục "SMS cần gạch paid".
3. **INTERFREIGHT và INTER-FREIGHT trên SMS là cùng một khách Misa INTER-FREIGHT LTD.**
   FIN ghi "Sai" là nhầm. Một khách Misa có thể có nhiều mã SMS.

Bổ sung từ trước (2026-10-06, cùng phiên):
- SMS đôi khi quên ấn paid. Tiền luôn theo Misa.
- FIN chỉ điền tay nợ quá chầy 2023–2024; nợ cũ khác lấy từ bản final của kế toán trưởng.
- Khách ghi "không có khả năng thu hồi" (Hoàng Kim, Vinh Giang, DLV, DTS, Sao Việt): luôn 120+.
