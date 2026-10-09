# MAP — weekly-report

Kiến trúc hiện tại. Mọi thay đổi ở tầng này phải vẽ ở đây TRƯỚC khi đổi mã.

Cập nhật: 2026-10-08 (WR-04). Trạng thái: v0.4 — full credit report: phải thu tính tuổi
từ ngày hoá đơn thật, thêm phải trả, dòng tiền, nhóm rủi ro, lịch sử tuần, cách tính.

## Hai nơi, đừng nhầm

```
THƯ MỤC MÃ NGUỒN (chỉ có mã)            THƯ MỤC ĐÍCH (người dùng tự trỏ)
weekly-report/                           Cong-No-Workspace/
  1_Khoi_tao_workspace.bat                 *_Credit_Report_*.xlsx   báo cáo tuần, ngay ở gốc;
  2_Kiem_tra_ghep_khach.bat                                         chị KTT sửa thẳng vào đây
  runner.py  wr/  tools/  tests/           input/   mỗi tuần 4 file (xem dưới)
  template/   mẫu báo cáo chuẩn            sample/  file khai 1 lần: ghép khách, Master List
  settings.json  chỉ ghi đường dẫn         _tool/   (ẩn) sổ ghi, nhật ký, lưu trữ
  handoff/    phòng bàn giao
```
Không còn nút "Chạy ngay": tool tự chạy khi đủ file; ép chạy sớm thì bấm Run trong
Task Scheduler. Đường dẫn chỉ nằm trong settings.json của từng máy.

## Sáu file mỗi tuần (input/), xuất CÙNG NGÀY

| File | Cho biết gì | Bắt buộc? |
|---|---|---|
| Misa Tổng hợp công nợ phải thu (TK131) | số dư từng khách = TIỀN đúng, kỳ báo cáo | có |
| Misa Chi tiết công nợ phải thu (TK131), **từ 01/01/2022 đến nay** | từng hoá đơn có ngày, từng phiếu thu có TK đối ứng | có |
| Misa Tổng hợp công nợ phải trả (TK331) | nợ vendor đầu kỳ, phát sinh, cuối kỳ | chờ 2 giờ rồi chạy thiếu |
| SMS AR-AP | Salesman của từng job (job + tiền đã lấy từ Misa khi có 2 file dưới) | chờ 2 giờ rồi chạy thiếu |
| Misa Bán hàng, **từ đầu năm** (WR-06) | trạng thái thanh toán từng hoá đơn | chờ 2 giờ; thiếu thì job lấy từ SMS (cách cũ) |
| Misa Sổ chi tiết tài khoản 131, **từ đầu năm** (WR-06) | job của từng hoá đơn ("Mã đối tượng THCP"), phiếu thu sau kỳ | chờ 2 giờ rồi chạy thiếu |

Mã: `wr/banhang.py` đọc 2 file WR-06; `wr/tindung.py` dựng dòng con + `doi_chieu()`.

Không có file Chi tiết mà có file Tuổi nợ cũ → chạy luật v0.3 (giữ nguyên để không gãy).

## Luồng chạy mỗi tuần

```
 input/ 4 file ──► nhận dạng theo NỘI DUNG (không tin tên file), ghép theo kỳ
 sample/ ghép khách, Master List ──► nhận dạng theo nội dung
 gốc/ báo cáo các tuần trước ──► đọc lại vào sổ ghi (bản người thắng bản máy)
                 │
   ① TIỀN mỗi khách   = Tổng hợp 131 (dư Nợ cuối kỳ). Luôn luôn.
   ② TUỔI             = Chi tiết 131: hoá đơn theo ngày hoá đơn; phiếu thu ghi số HĐ thì
                         trừ đúng HĐ đó, còn lại trừ HĐ CŨ NHẤT trước; hạn = ngày HĐ + term
                         ("14 ngày của tháng tiếp theo" = ngày 14 tháng sau);
                         quá hạn từ HÔM SAU ngày đến hạn (FIN 08/10)
                         khó đòi (bảng ghép / lý do có "khó đòi") → cả số vào 120+
                         Chi tiết ≠ Tổng hợp → tiền theo Tổng hợp, ghi Review
   ③ JOB, SALESMAN    = SMS (qua bảng ghép) > tuần trước
   ④ RỦI RO           = luật của chị KTT (S1..S6, N0..N4), so với tuần trước trong sổ ghi
   ⑤ PHẢI TRẢ         = Tổng hợp 331: mọi vendor còn dư Có + vendor trả xong trong tuần
   ⑥ DÒNG TIỀN        = phiếu thu trong kỳ của Chi tiết 131 tách theo TK đối ứng
                         (112 ngân hàng, 111 tiền mặt, 331 cấn trừ, 515/635 tỷ giá, khác)
                         + phát sinh Nợ 331 (đã trả vendor)
   ⑦ ghi <gốc>/<năm>_W<tuần>_Credit_Report_<từ>-<đến>.xlsx, file cùng tên đã có thì
      KHÔNG ghi đè, đặt tên "(chay lai dd-mm HHhMM)"; cập nhật sổ ghi + lịch sử tuần
```

## File báo cáo tuần (đầu vào của dashboard WR-05)

| Sheet | Một dòng là | Nguồn |
|---|---|---|
| Summary | 1 tuần (tuần này + mọi tuần trước) | tổng hợp các sheet dưới + sổ ghi |
| Receivables | 1 khách còn nợ: tuổi nợ, Salesman, Job Number, lý do; dòng con (gập) = job SMS + Sales + số tiền, dòng cân đối | ①②③ |
| Invoices | 1 hoá đơn còn nợ: ngày, hạn, số ngày quá hạn, nhóm; cộng theo khách = Total | ② |
| AR Risk | 1 khách: đầu kỳ, tăng, giảm, cuối kỳ, so tuần trước, S1..S6, nhóm rủi ro, xu hướng | ④ |
| Payable | 1 vendor: đầu kỳ, nợ mới, đã trả, cuối kỳ, trạng thái | ⑤ |
| Cash Flow | 1 khách đã trả / 1 vendor đã được trả trong tuần | ⑥ |
| Methodology | 1 con số: tính thế nào, lấy từ file nào cột nào | cố định trong mã |
| Review | đối chiếu và những chỗ tool không chắc | mọi bước |

Excel ghi SỐ, không ghi công thức (dashboard đọc file bằng trình duyệt, không tính lại được);
cách tính viết bằng lời ở sheet Methodology.

## Dashboard (WR-05)

`dashboard/src` (HTML, CSS, JS, thư viện đọc Excel, font) → `dashboard/build.py` → một file
`Credit_Dashboard.html` chạy offline. Mỗi lần quét, tool chép file này vào thư mục gốc workspace.
Người xem kéo file báo cáo tuần vào trang; dashboard chỉ đọc số, riêng nhóm rủi ro tính lại khi
lý do trễ hạn trong Receivables khác bản tool viết. Nút In / Lưu PDF dùng hộp in của trình duyệt.
WR-07: mỗi tab là một trang (Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả · Dòng tiền). Bấm
một khách / một Sales → trang chi tiết; bộ lọc ghi trên thanh địa chỉ (`#sales?s=…&c=…`) nên đi
theo qua các tab và nút Back chạy. In / Lưu PDF in đủ mọi trang.

## Sổ ghi (_tool/so-ghi.json)

Một trang cho mỗi ngày chốt: từng khách (số dư, 6 nhóm, lô nợ, term, salesman, job, lý do)
+ `lich_su_tuan` (các con số của sheet Summary từng tuần). Báo cáo ở thư mục gốc được đọc
lại thành trang của đúng ngày chốt ghi trong file. Cùng ngày chốt: bản không có tab Review
(bản người làm) thắng bản tool; cùng loại thì bản sửa sau cùng thắng.

## Hai nút cho người dùng

1. `1_Khoi_tao_workspace.bat` — cài thư viện, chọn thư mục đích, tạo ngăn, đăng ký lịch 15 phút.
2. `2_Kiem_tra_ghep_khach.bat` — đối chiếu file ghép khách với Misa/SMS mới nhất, mở file KIEM_TRA.

## Ngoài phạm vi hiện tại

Nợ vendor theo job/loại phí từ SMS (cần ghép vendor SMS ↔ Misa).
Customer Master List từ SharePoint (sẽ nhận dạng khi có file mẫu). Power Automate gửi email.
