# NOTE — Kế toán trả lời đợt 2, 2026-09-23

Người trả lời: kế toán phụ trách công nợ. Hỏi qua Chairman.

## 1. Ba khách Misa báo quá hạn nhưng bản tay để Current

- **AHC Logistics (Xiamen)**: khách này mình cũng có công nợ PHẢI TRẢ cho họ, bù
  trừ xong thì mình còn phải trả tiền cho họ. Nhưng chưa ký biên bản bù trừ nên
  kế toán vẫn ghi nhận công nợ phải thu, và để ở **nợ trong hạn**.
  → Xử lý: khai ngoại lệ, ép về Current. Ký biên bản bù trừ xong thì bỏ ngoại lệ.
- **Optec Express (HK)**: khách thanh toán công nợ theo tháng, cuối tháng 9 trả
  nợ của cả tháng 8, nên kế toán chưa chuyển sang quá hạn. Misa thì ghi nhận 30
  ngày kể từ ngày xuất hoá đơn nên mới có phần quá hạn.
  → Kế toán chốt: **làm theo Misa**.
- **Grand Forwarding**: trước không có hợp đồng nên để 30 ngày; mới ký hợp đồng
  hạn 15 ngày, báo cáo tay chưa cập nhật.
  → Kế toán chốt: **làm theo Misa**, và credit term đổi thành 15 ngày.

## 2. LLC UPP lệch 1 tỷ ở W38

Tại thời điểm làm báo cáo, một kế toán viên khác chưa ghi nhận 1 tỷ đó; tới 22/09 mới ghi
nhận lùi lại nên số trên Misa cao hơn bản tay.
→ **Lấy số trên Misa.** Đúng với luật đã chốt: tiền theo file Tổng hợp.

## 3. Cột số job

Giữ nguyên số job của tuần trước, kế toán tự sửa. **Không xoá trắng.**

## Việc phải làm khi cài đặt

```
python runner.py --ngoai-le "AHC LOGISTICS(XIAMEN) CO.,LTD" --nhom "Current" \
                 --vi-sao "cho ky bien ban bu tru cong no hai chieu"
python runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"
```
