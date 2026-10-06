# MAP — weekly-report

Kiến trúc hiện tại. Mọi thay đổi ở tầng này phải vẽ ở đây TRƯỚC khi đổi mã.

Cập nhật: 2026-10-06 (WR-03). Trạng thái: v0.3 — thêm nguồn SMS, file ghép khách,
bản final của kế toán trưởng làm trí nhớ nợ cũ.

## Hai nơi, đừng nhầm

```
THƯ MỤC MÃ NGUỒN (chỉ có mã)            THƯ MỤC ĐÍCH (người dùng tự trỏ) — Chairman chốt 06/10
weekly-report/                           Cong-No-Workspace/
  1_Khoi_tao_workspace.bat                 *.xlsx       báo cáo tuần tool sinh ra, nằm ngay ở gốc;
  2_Kiem_tra_ghep_khach.bat                             lần đầu FIN bỏ bản final gần nhất vào đây
  3_Chay_ngay.bat                          input/       mỗi tuần: 2 file Misa + 1 file SMS
  runner.py  wr/  tools/  tests/           sample/      file ghép khách SMS ↔ Misa + file KIEM_TRA
  template/   mẫu báo cáo chuẩn            _tool/       (ẩn) sổ ghi, nhật ký — người dùng không đụng
  settings.json  chỉ ghi đường dẫn
  handoff/    phòng bàn giao
```
Đường dẫn chỉ nằm trong settings.json của từng máy. Cấm ghi cứng vào mã hay tài liệu.

## Luồng chạy mỗi tuần

```
 input/: Misa Tổng hợp ─┐                        sample/:  bảng ghép (FIN, 1 lần)
         Misa Tuổi nợ  ─┼─► kiểm cùng kỳ          gốc/:     báo cáo các tuần trước (có thể đã sửa tay)
         SMS AR-AP     ─┘   (TH + 1 ngày = TN)               → đọc lại vào sổ ghi trước mỗi lần chạy
                  │                                       │
                  ▼                                       ▼
   ① GHÉP mỗi khách Misa ← các dòng SMS còn nợ (qua bảng ghép, rồi tên/mã trùng)
   ② TIỀN  = Misa Tổng hợp, luôn luôn
   ③ TUỔI
        khó đòi (FIN khai)       → cả số tiền vào 120+
        SMS > Misa               → cắt dòng SMS cũ nhất cho bằng Misa, liệt kê để FIN gạch paid
        SMS ≤ Misa               → dòng SMS: hạn = ngày ETD/ETA + credit term
                                   phần dư = nợ cũ: lấy lô CŨ NHẤT trong sổ ghi trước đó
                                   thiếu sổ ghi → hạn = ngày sớm nhất trong file SMS, đánh dấu
        không có file SMS        → luật cũ v0.2 (file Tuổi nợ + sổ ghi)
   ④ ghi <gốc>/<tên báo cáo tuần>.xlsx + sheet "Review" + trang sổ ghi mới
      file cùng tên đã có (có thể đã sửa tay) → KHÔNG ghi đè, đặt tên "(chay lai dd-mm HHhMM)"
```

## Sổ ghi

Một trang cho mỗi ngày chốt. Mỗi khách: số dư, 6 nhóm tuổi, các lô nợ kèm hạn,
credit term, salesman, job, ghi chú, lý do. Mọi báo cáo ở thư mục gốc (bản tool sinh
rồi kế toán sửa, hoặc bản final kế toán trưởng) được đọc lại thành trang của đúng ngày
chốt ghi trong file: người thắng máy. Cùng ngày chốt nhiều bản thì bản sửa sau cùng thắng.
Vì vậy chỉ LẦN ĐẦU cần bỏ bản final tuần trước; các tuần sau tool tự đọc báo cáo của nó.

## Nguồn dữ liệu

| Cần gì | Lấy ở đâu |
|---|---|
| Danh sách khách, Total | Misa Tổng hợp |
| Ngày chốt, kiểm cùng kỳ | Misa Tuổi nợ |
| Ngày từng job, salesman, số job | SMS AR-AP (cột Partner Code, Job, ETD, ETA, A/R remain) |
| SMS khách nào = Misa khách nào | sample/ (FIN điền 1 lần, bổ sung khi có khách mới) |
| Tuổi của nợ cũ không có trong SMS | Báo cáo tuần trước ở thư mục gốc (lần đầu: bản final kế toán trưởng) |
| Khách khó đòi | sample/, ghi chú "không có khả năng thu hồi" |
| Credit term, lý do, ghi chú | Khai báo của kế toán > sổ ghi/final tuần trước |

## Ba nút cho người dùng

1. `1_Khoi_tao_workspace.bat` — cài thư viện, chọn thư mục đích, tạo ngăn, đăng ký lịch 15 phút.
2. `2_Kiem_tra_ghep_khach.bat` — đọc lại báo cáo ở gốc, đọc file ghép khách,
   đối chiếu với Misa/SMS mới nhất trong input/, mở file KIEM_TRA trong sample/.
3. `3_Chay_ngay.bat` — chạy ngay thay vì chờ lịch.

## Ngoài phạm vi hiện tại

Sheet Payable, Dashboard. Phase Power Automate đọc báo cáo mới nhất ở thư mục gốc để gửi email.
