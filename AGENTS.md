# weekly-report — bộ định tuyến

Công cụ sinh báo cáo công nợ phải thu hằng tuần của Trustana Việt Nam.
Kế toán thả 2 file Excel xuất từ Misa vào thư mục OneDrive; một tiến trình
Python chạy nền theo lịch đọc 2 file đó cộng với sổ ghi của chính nó, rồi
điền ra file báo cáo tuần theo mẫu chuẩn. Không ai phải điền tay.

Đây là tool đầu tiên trong chuỗi credit control. Các tool sau và phase
Power Automate dùng chung thư mục OneDrive này.

## Trước khi làm bất cứ việc gì

1. Đọc `handoff/RULES.md` — luật đứng của dự án.
2. Đọc `handoff/STATE.md` — hiện trạng, việc đang mở.
3. Mở ĐÚNG MỘT phiếu việc được giao trong `handoff/`.

Đụng tới kiến trúc thì đọc thêm `handoff/MAP.md`, và vẽ ở đó TRƯỚC khi đổi mã.

Cấm quét cả thư mục `handoff/`. Cấm đọc phiếu đã đóng khi không có trường
`Nối tiếp` trỏ tới. Cấm đọc `handoff/docs/` khi không có dòng nào trỏ tới.

Chuẩn phối hợp: `../handoff-standard-v1.1.md`.
