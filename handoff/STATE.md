Cập nhật:  2026-09-24 21:00 UTC+7 — bởi claude-logisticist
Chỉ huy:   claude   [COWORK-EXEC]

## Việc đang mở
   WR-02   giao cho chairman   chờ-đề-bài   vòng 0/2
           Nghiệm thu bằng phép so hợp lệ: kế toán làm tay trên đúng cặp file
           Misa mà tool đã chạy. Chờ Chairman viết kịch bản nghiệm thu.

   WR-01   giao cho claude-logisticist   chờ-nghiệm-thu   vòng 1/2
           Logic ĐẠT (29 phép thử). Chạy trên Windows của Chairman ĐẠT một phần
           (bằng chứng: runner.log 23/09 16:04, ra file W39). Còn 3 dòng chờ
           Chairman hoặc Codex ký: setup.bat, Task Scheduler, hiển thị Excel.

## Hôm nay
   Đổi nhịp chạy nền sang 15 phút. Dọn rác trong thư mục mã nguồn và chặn không
   cho sinh lại: lệnh chỉ đọc không tạo thư mục, cấu hình của máy khác thì báo
   chứ không đẻ thư mục tên "C:\Users\...", tắt sinh .pyc khi chạy nền, tắt
   cache của pytest. Viết lại README cho FIN đọc. Khởi tạo git, commit đầu tiên.
   29 phép thử.

## Hôm 24/09 (chiều)
   Rà tính di động: cấu hình theo tên máy, chạy nền không chết vì tiếng Việt,
   Task Scheduler gọi qua chay_nen.bat, file kết quả đang mở thì hoãn. 25 phép thử.
   Đối chiếu 3 bản tay W33-W35: nối từ bản tay tuần trước thì khớp 60/72 (W34)
   và 47/68 (W35). Ba nhóm chênh đã ghi trong phiếu, cần kế toán trả lời.

## Hôm 24/09 (sáng)
   Chairman thả 6 file W33-W35 vào 01_Input, bấm Run không ra gì. Do tôi xoá
   nhầm settings.json hôm qua, cộng với việc tool chỉ ghép được 1 cặp file.
   Đã sửa cả hai: ghép cặp theo kỳ và chạy nhiều tuần lần lượt, mọi nhánh dừng
   sớm đều ghi nhật ký. Chạy lại ra đủ 3 báo cáo W33, W34, W35. 21 phép thử.
   Tách "thiếu Credit Term / Salesman" khỏi "lỗi số liệu" để bảng bớt tô màu.

## Hôm 23/09
   Chairman viết kịch bản nghiệm thu, mở việc WR-01.
   Dựng xong tool: đọc 2 file Misa, sổ ghi theo lô nợ có hạn, luật chia 6 nhóm
   tuổi, ghi file theo mẫu chuẩn kẻ màu thương hiệu, wizard cài đặt 1 lần.
   Chạy kịch bản W37: ra đúng file, 74 khách, tổng 9.547.752.597 khớp Misa,
   45 dòng vào sheet Cần xem lại. Thả nhầm file lệch kỳ thì dừng, không ra file.
   Đối chiếu bản kế toán làm tay: khớp 62/74 dòng, 12 dòng lệch giải thích được.
   Dọn mẫu chuẩn còn một sheet Receivable trắng; xoá 2 file fixture trùng.

   Chiều: kế toán trả lời đợt 2. Thêm cơ chế khai đè credit term và ngoại lệ.
   Chạy tiếp W38 nối từ W37: 70 khách, khớp bản tay 56/70; dòng lệch to nhất là
   LLC UPP mà kế toán đã xác nhận máy đúng. 15 phép thử đạt.

   Cuối chiều: cho người dùng tự trỏ thư mục đích (--dat-thu-muc / --xem-cau-hinh),
   tool tạo đủ 6 ngăn bên trong. Thư mục mã nguồn không còn chứa dữ liệu chạy.
   Bắt được lỗi lệnh chạy tay bỏ qua cấu hình máy, đã sửa. 17 phép thử đạt.

   Tối: kế toán thử bộ dữ liệu giả W39, hỏi vì sao tổng báo cáo khác tổng file
   Tuổi nợ. Số đúng cả, nhưng báo cáo chưa giải thích. Thêm khối đối chiếu ở
   đầu sheet Cần xem lại. 19 phép thử đạt.

## Hôm qua
   Đọc 3 tuần dữ liệu thật W36, W37, W38 (mỗi tuần: 2 file Misa + 1 báo cáo
   kế toán làm tay). Rút ra luật điền báo cáo, chạy thử luật đó trên dữ liệu
   thật: khớp 64/74 dòng ở W37 và 56/70 dòng ở W38.
   Chốt với kế toán: danh sách và Total lấy từ file Tổng hợp, tuổi nợ lấy từ
   file Tuổi nợ, hai bên mâu thuẫn thì tiền theo Tổng hợp.
   Phát hiện kế toán đang mang cách chia của tuần trước sang tuần sau — đây là
   thứ tool phải thay thế bằng sổ ghi.
   Chairman chốt điều kiện: kế toán không còn thao tác gì ngoài thả 2 file.
   Dựng phòng bàn giao theo Chuẩn Bàn Giao v1.1.

## Vừa đóng
   chưa có

## Đang treo, cần Chairman
   -1. Ký 3 dòng còn lại trong ô ký nghiệm thu của WR-01, rồi phiếu chuyển
      đóng-đạt. Viết kịch bản nghiệm thu cho WR-02.
   0. Chốt cách nghiệm thu mới: kế toán làm tay trên ĐÚNG cặp file Misa đã thả
      vào 01_Input, rồi mới so với bản máy. So với bản làm từ lần xuất khác là
      kết luận sai về tool.
   1. Ký nghiệm thu phần chạy trên Windows: chạy setup.bat, xem Task Scheduler
      có tạo được tác vụ không, mở file kết quả bằng Excel xem màu và tiếng Việt.
   2. Máy nào chạy lịch nền, mấy giờ thứ mấy. Wizard đang đặt 30 phút quét một lần.

## Bẫy đã biết
   - File Misa tải về bị Windows thêm hậu tố (21), (22) vào tên. Con số đó
     không có nghĩa. Luôn đọc ngày chốt ghi bên trong file, không tin tên file.
   - File Tuổi nợ phồng số: W38 cộng ra 73 tỷ trong khi thực nợ 10,0 tỷ, vì
     chứa hoá đơn đã trả nhưng chưa gắn phiếu thu.
   - 56,2/73 tỷ trong file Tuổi nợ nằm ở ô "Không có hạn nợ" — hoá đơn cũ
     không ghi hạn thanh toán.
   - Báo cáo mẫu cũ có công thức trỏ tới ô cố định Receivable!E78; mẫu mới
     tự tính dòng TỔNG theo số dòng thật nên hết bẫy này.
   - Lệnh chạy tay mà không đọc settings.json thì sổ ghi rơi vào thư mục mã
     nguồn, tool tưởng khởi động lạnh mà không báo gì rõ ràng.
   - Mã khách của cùng một khách khác nhau giữa hai file Misa, nên tool so khớp
     theo TÊN đã chuẩn hoá. Đổi tên khách trong Misa là mất liên kết sổ ghi.

## Báo lỗi giả
   Tháng này: 0
