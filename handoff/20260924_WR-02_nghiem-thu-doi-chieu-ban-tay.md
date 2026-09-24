Mã:         WR-02
Mạch:       WR — báo cáo công nợ phải thu hằng tuần
Nối tiếp:   WR-01
Giao cho:   chairman
Trạng thái: chờ-đề-bài
Vòng:       0/2

## ĐỀ BÀI

**Mục tiêu nghiệp vụ**
Đo độ chính xác thật của tool bằng một phép so hợp lệ: kế toán làm bảng bằng tay
trên ĐÚNG cặp file Misa mà tool đã chạy, rồi so từng khách từng cột.

**Ngoài phạm vi**
- Không sửa luật tính số trong phiếu này. Có sai thì ghi nhận, mở phiếu riêng.
- Không đụng tới cột Salesman và số job.

**Kế thừa**
- Luật tính số: `handoff/docs/SPEC-logic-bao-cao-tuan.md`.
- Vì sao phải cùng một cặp file: `handoff/docs/NOTE-ke-toan-tra-loi-20260924.md`
  mục 3 — số trên Misa đổi theo thời điểm xuất, so hai lần xuất khác nhau là
  kết luận sai về tool.
- Ba lần so trước đây (W33, W34, W35) đều không thoả điều kiện này, con số
  60/72 và 47/68 chỉ là sàn dưới.

**Kịch bản nghiệm thu**
CHỜ CHAIRMAN VIẾT. Gợi ý các mốc đo được để Chairman chọn và sửa lời:
- Kế toán xuất 2 file Misa, thả vào 01_Input, và giữ lại đúng 2 file đó.
- Kế toán làm bảng bằng tay trên chính 2 file đó, thả vào 03_BanTay.
- Chạy `python tests/so_sanh.py <file máy> <file tay>`.
- Ngưỡng đạt: bao nhiêu phần trăm dòng khớp cả 6 nhóm tuổi thì coi là đạt, và
  mọi dòng lệch phải giải thích được bằng một luật đã ghi trong SPEC.

## TODO

(chưa mở, chờ ĐỀ BÀI đủ)

## AUDIT

(trống)

## HISTORY

### 2026-09-24
Failure    không
           Báo lỗi giả: không
Lessons    Tách phần nghiệm thu nghiệp vụ ra khỏi WR-01 để phiếu đó không phải
           nằm mở chờ lịch làm việc của kế toán.
-- claude-logisticist
