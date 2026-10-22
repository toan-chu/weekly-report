Mã:         WR-09
Mạch:       WR — báo cáo công nợ phải thu hằng tuần
Nối tiếp:   WR-08
Giao cho:   claude-logisticist   [COWORK-EXEC]
Trạng thái: chờ-nghiệm-thu
Vòng:       1/2

## ĐỀ BÀI

**Mục tiêu nghiệp vụ** (Chairman, 2026-10-20, theo ghi chú FIN sau W42)
1. STA CENTER: dòng "HĐ 1560 ngày 04/09 — 29,4tr" không có trên sổ; 29,4tr là của 2 hoá đơn khác.
2. Tên công ty: TRUSTANA VIETNAM COMPANY LIMITED.
3. Client Master thêm cột "Đã có hợp đồng": có thì tích, chưa có để trống (Sales đề nghị).
4. Chạy W42 từ file FIN gửi, so với bản "(human)" FIN sửa tay, liệt kê thay đổi.

## TODO
- [x] Tiền thu không ghi số HĐ trừ trước vào HĐ Misa ghi "Đã thanh toán" → STA: HĐ 1529 27.957.741 + HĐ 1322 1.463.336 (đúng ý FIN)
- [x] Chênh lệch tỷ giá NVK cùng ngày với phiếu thu ghi số HĐ → gộp vào HĐ đó (OPTIMALOG HĐ 579 = 2.115.321)
- [x] Khó đòi có trong bảng job nợ cũ vẫn hiện dòng HĐ (Minh Khôi HĐ 2519)
- [x] CONG_TY = TRUSTANA VIETNAM COMPANY LIMITED
- [x] Client Master: cột Đã có hợp đồng (✓, thả xuống), ngày gia hạn chỉ tính khi đã tới; Summary đếm khách chưa có HĐ
- [x] tools/nang_cap_client_master.py: nâng cấp file FIN đang dùng (82 khách, 13 ✓ theo Số hợp đồng)
- [x] tests/test_wr9.py

## AUDIT
```
python -m pytest            → 72 passed, 20 skipped
W42 (workspace test, mốc = W41 human) → 68 khách; phải thu 15.651.640.892; quá hạn 5.648.539.192;
                               nợ xấu 1.214.313.194; phải trả 7.574.718.156
So W42 (human)              → 68/68 khách, tổng khớp hết, term và Sales khớp hết; dòng job khớp 64/68
LibreOffice tính lại        → 0 lỗi công thức
```
4 khách còn khác, tool theo sổ + trạng thái Misa (chờ FIN xác nhận): OPTIMALOG (579 vs 1019), ACT 247
(13.898 của HĐ 1790), RUNWAY và AIRSPACE (tỷ giá gộp vào HĐ thay vì dòng riêng).

Ký nghiệm thu (Chairman): ____  [RUNTIME CHƯA KIỂM]

## HISTORY
### 2026-10-20
Failure: tiền thu không ghi số HĐ trừ theo ngày HĐ cũ nhất, trong khi Misa đã đối trừ theo cách khác
(HĐ 1560 hạch toán 22/07 nhưng ngày HĐ 04/09).
Lessons: trạng thái "Đã thanh toán" trên Misa là dấu vết đối trừ của kế toán — dùng nó để phân bổ trước FIFO.
