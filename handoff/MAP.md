# MAP — weekly-report

Kiến trúc hiện tại. Mọi thay đổi ở tầng này phải vẽ ở đây TRƯỚC khi đổi mã.

Cập nhật: 2026-09-23. Trạng thái: mới dựng phòng bàn giao, **chưa có mã chạy**.

## Hai nơi, đừng nhầm

```
THƯ MỤC MÃ NGUỒN (chỉ có mã)            THƯ MỤC ĐÍCH (người dùng tự trỏ)
weekly-report/                           <thư mục đích>/
  runner.py                                01_Input/   kế toán thả 2 file Misa
  setup.bat                                02_Output/  báo cáo tuần máy sinh ra
  wr/ tools/ tests/                        03_BanTay/  bản kế toán tự làm, nếu có
  template/         mẫu báo cáo chuẩn      so-ghi/     trí nhớ của tool
  settings.json     chỉ ghi đường dẫn      luu-tru/    bản cũ, không xoá
  handoff/          phòng bàn giao         log/        nhật ký mỗi lần chạy
    docs/fixtures/  dữ liệu W36-W38
```
Thư mục đích do người dùng trỏ lúc cài đặt — OneDrive, SharePoint đã đồng bộ,
ổ D hay ổ mạng đều được. Wizard dò sẵn các thư mục đồng bộ để chọn nhanh, và
luôn cho phép gõ đường dẫn bất kỳ. Đổi chỗ sau này:
`python runner.py --dat-thu-muc "D:\..."`, xem lại: `--xem-cau-hinh`.
Đường dẫn chỉ nằm trong settings.json của từng máy. Cấm ghi cứng vào mã hay
tài liệu. Thư mục mã nguồn không chứa dữ liệu chạy: sổ ghi, nhật ký, báo cáo
đều nằm trong thư mục đích.

## Luồng chạy

```
[kế toán]  thả 2 file Misa vào 01_Input        ← thao tác duy nhất của con người
    │
[lịch nền] đến giờ thì gọi runner
    │
    ├─ đọc ngày chốt trong từng file; hai file lệch kỳ → dừng, báo lỗi
    ├─ lập danh sách khách và Total          ← file Tổng hợp
    ├─ chia Total vào 6 nhóm tuổi            ← file Tuổi nợ + sổ ghi
    ├─ mang cột người dùng từ báo cáo tuần trước sang
    ├─ tự kiểm tra: tổng dòng = Total, tổng bảng = tổng Misa
    ├─ ghi 02_Output/<tên báo cáo tuần>.xlsx theo mẫu trong template/
    └─ ghi thêm một trang vào so-ghi/
```

## Sổ ghi

Một trang cho mỗi ngày chốt số. Nội dung mỗi trang: từng khách, số dư, cách
chia 6 nhóm tuổi, ngày mỗi khoản xuất hiện lần đầu, và nguồn của dòng đó
(máy tự tính hay chép từ bản tay). Đây là thứ thay cho trí nhớ của kế toán.

## Nguồn dữ liệu

| Cần gì | Lấy ở đâu |
|---|---|
| Danh sách khách, Total | File Tổng hợp công nợ phải thu |
| Tuổi nợ | File Phân tích công nợ theo tuổi nợ |
| Credit Term, Salesman, số job, Lý do chưa thu hồi | Báo cáo tuần trước |
| Tuổi của nợ chưa có hạn | Sổ ghi (ngày xuất hiện lần đầu) + Credit Term |
| Ngoại lệ nghiệp vụ (VD lô tàu chìm) | Sheet Ngoại lệ trong sổ ghi |

## Ngoài phạm vi hiện tại

Sheet Payable, General, Dashboard, Hard-to-collect. Cột Salesman và số job
lấy tự động từ SMS. Phase Power Automate đọc 02_Output để gửi Teams và email.
