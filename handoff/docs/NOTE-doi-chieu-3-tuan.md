# NOTE — Đối chiếu luật với 3 tuần làm tay

Ngày chạy: 2026-09-22. Dữ liệu: `fixtures/` (W36, W37, W38, mỗi
tuần 2 file Misa + 1 báo cáo kế toán làm tay).
Cách chạy: áp luật trong `SPEC-logic-bao-cao-tuan.md`, so từng khách từng cột
với bản kế toán làm tay. Khớp nghĩa là lệch dưới 2 đồng.

## Kết quả

| | W37 | W38 |
|---|---|---|
| Nhóm A — chép thẳng từ file Tuổi nợ | 25/29 | 26/31 |
| Nhóm B — chưa có hạn nợ, dùng sổ ghi tuần trước | 29/32 | 22/30 |
| Nhóm C — hai file lệch nhau, dùng sổ ghi tuần trước | 10/13 | 8/9 |
| **Tổng dòng khớp** | **64/74** | **56/70** |
| Giá trị khớp | 7,79/9,40 tỷ (83%) | 7,90/10,00 tỷ (79%) |

Total lấy từ file Tổng hợp: khớp 68/73 ở W36, 70/74 ở W37, 69/70 ở W38.

## 14 dòng W38 chưa khớp, chia 4 dạng

1. Ngoại lệ nghiệp vụ (2): Việt Nga, BSF — lô tàu chìm, xuất lại debit note.
   Xử lý bằng sheet Ngoại lệ.
2. Nợ chưa có hạn đã tới lúc chuyển nhóm (7): PROLOGISTOV, QINGDAO, RELOPET,
   Đối Tác Hàng Hoá, EASYFLYERS, ALLIANCE 21, SPA. Luật tính tuổi theo ngày
   xuất hiện lần đầu trong sổ ghi giải quyết được nhóm này, nhưng chỉ sau khi
   sổ ghi đã chạy đủ vài tuần.
3. Misa báo quá hạn nhưng bản tay để Current (3): AHC, GRAND, OPTEC. Nhiều khả
   năng kế toán mang số tuần trước sang mà quên đọc lại file Misa. Nếu đúng thì
   ở 3 dòng này máy đúng hơn bản tay. ĐANG CHỜ KẾ TOÁN XÁC NHẬN.
4. Kế toán tự phán đoán (2): Tập đoàn IDC, Hội Tụ Thông Minh.

Cả 4 dạng đều không sai Total, chỉ có thể xếp nhầm nhóm tuổi.

## Điểm còn treo

LLC UPP ở W38: bản tay 1,11 tỷ, file Tổng hợp 2,11 tỷ, lệch đúng 1 tỷ.
Chưa rõ bên nào đúng. ĐANG CHỜ KẾ TOÁN.
