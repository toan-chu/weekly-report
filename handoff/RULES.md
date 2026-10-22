# RULES — weekly-report

Sinh từ `Chuẩn Bàn Giao v1.1` (chốt 2026-08-25), bản gốc ở
`../../handoff-standard-v1.1.md`. Lệch phiên bản thì đề xuất cập nhật,
không tự sửa.

Chỉ chứa luật đứng — thứ còn đúng vào tháng sau. Trần 200 dòng.

---

## 1. Chỉ huy hiện tại và cách đổi

Chỉ huy: `claude` (nhân dạng `claude-logisticist`), thẻ `[COWORK-EXEC]`.

Chairman giao trực tiếp việc này cho Claude, tức trường hợp giải trừ giới hạn:
Claude có toàn quyền kỹ thuật và toàn quyền thực thi trên tool này. Giải trừ
chấm dứt khi Chairman thu hồi, hoặc khi Codex tiếp nhận — lúc đó đổi dòng
"Chỉ huy" trong `STATE.md` và ghi một dòng vào khoang HISTORY của phiếu đang mở.

Hai thứ không bao giờ giải trừ được:
- Phạm vi nghiệp vụ vẫn chờ Chairman.
- Phán quyết runtime trên máy Windows: Claude không tự nhận. Gắn
  `[RUNTIME CHƯA KIỂM]` và chuyển Codex hoặc Chairman xác nhận.

## 2. Quyền ký nghiệm thu

Chia theo đường lấy bằng chứng, không theo vai.

- Claude ký được: logic tính toán, so khớp số giữa file máy sinh và file kế
  toán làm tay, đối chiếu công thức. Bắt buộc dán lệnh và kết quả vào khoang AUDIT.
- Claude KHÔNG ký được: phép thử đạt/hỏng chạy trên máy Windows của kế toán,
  `1_Khoi_tao_workspace.bat` có cài được lịch chạy không, Task Scheduler có gọi đúng không,
  toàn vẹn file, mã hoá ký tự tiếng Việt khi mở bằng Excel thật. Đó là độc
  quyền Codex hoặc Chairman, chứng minh bằng lệnh và kết quả dán vào.

## 3. Trần vòng lặp và đường lên Chairman

- Nghiệm thu báo HỎNG tối đa 2 vòng. Hết trần thì chuyển `chờ-chairman`
  kèm 2 phương án nén. Cấm vòng thứ ba.
- `KHÔNG ĐO ĐƯỢC` không tính là một vòng: sửa phép thử trước, không sửa sản phẩm.
- Trạng thái chỉ đẩy tiến. Đóng rồi không mở lại; làm tiếp thì mở phiếu mới,
  ghi `Nối tiếp` trỏ về phiếu cũ.

## 4. Bốn câu của Chairman

1. "Kịch bản nghiệm thu là gì? Đọc cho tôi bằng tiếng Việt: ai bấm gì, thấy gì,
   trong bao lâu."
2. "Đã chạy chưa? Dán lệnh và kết quả vào."
3. "Làm hỏng thử một chỗ đi — phép thử có đỏ lên không?"
4. "Có chỗ nào KHÔNG ĐO ĐƯỢC không? Vì sao?"

---

## 5. Luật đứng của riêng tool này

### 5.1 Ranh giới với người dùng

- Thao tác hằng tuần duy nhất của kế toán là thả 6 file, xuất cùng lúc, vào `input/`:
  Tổng hợp 131, Chi tiết 131 từ 01/01/2022, Tổng hợp 331, SMS AR-AP (Chairman chốt
  2026-10-08, WR-04) + Bán hàng và Sổ chi tiết tài khoản 131, hai file xuất TỪ ĐẦU NĂM (FIN 09/10, WR-06). Không có nút "chạy ngay": ép chạy bằng Run trong Task Scheduler. Thao tác một lần: điền file ghép
  khách vào `sample/`, bỏ bản final tuần gần nhất vào thư mục gốc. Bố cục thư mục đích:
  gốc (báo cáo) + `input/` + `sample/` + `_tool/` ẩn — Chairman chốt 06/10.
  Mọi thiết kế ép kế toán mở file ra sửa, điền, hay chạy lệnh đều là vi phạm.
- Tool không bao giờ chặn không ra báo cáo vì thiếu thông tin. Không chắc chỗ
  nào thì vẫn ra file, đánh dấu chỗ đó vào sheet "Cần xem lại".
- Tool không xoá và không sửa đè file nào của kế toán. File cũ chuyển sang
  thư mục lưu trữ, không biến mất.

### 5.2 Số liệu

- Tuổi nợ (v0.4) tính từ từng hoá đơn trong file Chi tiết 131: hạn = ngày hoá đơn + credit
  term; quá hạn từ HÔM SAU ngày đến hạn (FIN 08/10). Phiếu thu ghi số hoá đơn thì trừ đúng
  hoá đơn đó, còn lại trừ hoá đơn cũ nhất trước. Cấm suy tuổi từ ngày ETD/ETA khi đã có file Chi tiết.
- Payable lấy toàn bộ số dư phải trả trên Misa, không lọc (FIN 08/10). Sheet Payable y chang mẫu FIN:
  No · Code · Name · Current · Total, chỉ số tổng, không phân tích, không tô màu (FIN 09/10).
- Job chưa thanh toán (v0.5) lấy từ Misa: Bán hàng (trạng thái) + Sổ chi tiết 131 (job). Chỉ "Đã
  thanh toán" mới tính là đã trả; "Chưa" và "Một phần" liệt kê đủ tiền hoá đơn (FIN 09/10). Trạng
  thái Misa là tại lúc xuất file: phiếu thu sau ngày cuối kỳ thì hoá đơn vẫn là nợ của kỳ.
- (v0.6, WR-08) Tiền dòng con = số CÒN NỢ của từng hoá đơn theo sổ (Chi tiết 131 đã trừ phiếu
  thu), job lấy từ Sổ chi tiết 131, hoá đơn trước năm nay lấy job từ bảng job nợ cũ. Dòng con
  luôn cộng bằng Total. Chứng từ không phải số hoá đơn (NVK) = "Điều chỉnh / chênh lệch tỷ giá".
  Cột Đối chiếu chỉ còn để FIN dọn trạng thái trên Misa.
- (WR-08) Sheet AR Risk ghi CÔNG THỨC của kế toán trưởng (S1–S6, Signals, Tier, Trend, Reason,
  Required action); dòng TOTAL ở Receivables = SUMIFS chỉ cộng dòng khách. Các sheet khác ghi số.
  Dashboard tự tính lại phần công thức. Đổi luật tính thì phải sửa Methodology cùng lần sửa.
- (WR-08) Credit term ưu tiên: file mẫu credit term trong `sample/` (nhận theo tiêu đề cột, dùng
  được nguyên file của kế toán trưởng) > khai `--credit-term` > tuần trước. Khách chưa có term
  liệt kê ở sheet Contracts. Không còn phiếu khách mới: Sales điền thẳng Client Master (13/10).
- (13/10) Client Master là file chung trên OneDrive. Trong `sample/` thì không cần cài gì; nằm chỗ khác
  thì mỗi máy khai đường dẫn (`--client-master`, lưu `client_master_theo_may` trong settings.json, không
  đi theo git). Có file chung → chỉ đọc file đó. Không thấy file → dùng `sample/` + cảnh báo ở Review.
- (20/10, FIN W42) Tiền thu không ghi số HĐ: trừ trước vào HĐ Misa (Bán hàng) ghi "Đã thanh toán" tại ngày chốt,
  rồi mới FIFO. Chênh lệch tỷ giá (NVK) cùng ngày với phiếu thu ghi số HĐ → cộng vào HĐ đó.
- (20/10) Client Master có cột "Đã có hợp đồng" (✓). Ngày gia hạn chỉ có hiệu lực khi đã tới ngày đó.
  Đổi mẫu thì nâng cấp file đang dùng bằng `tools/nang_cap_client_master.py`, không dựng lại từ đầu.
- (13/10, Chairman chọn A) S7 lỡ hẹn: lý do có ngày/tháng hẹn trả ('15/9', 'tháng 9' = cuối tháng,
  'đầu tháng' = ngày 10, 'giữa tháng' = ngày 20) đã qua ngày chốt mà quá hạn không giảm so với tuần
  trước → tính như S1–S3 (2 dấu hiệu = N3). Cột Promise date tool điền từ lý do, FIN sửa tay được.
  Luật đọc ngày phải giống nhau ở `tindung.ngay_hen` và `ngayHen` của dashboard.
- (13/10, KTT qua Chairman) N3 chỉ hiện câu "việc cần làm" trong báo cáo; FIN tự kiểm và quyết có
  ngưng dịch vụ hay không. Tool không khoá, không gửi gì.
- (13/10) Client Master (file chung OneDrive): FIN là người sửa lần cuối. Tool không có cột "xác nhận";
  thay vào đó sheet Contracts báo "Credit term vừa đổi" mỗi khi term khác tuần trước để FIN soát.
- (13/10) Hợp đồng hết hạn chưa gia hạn: tuổi nợ vẫn tính theo term cũ. Tool chỉ cảnh báo (sheet
  Contracts + Summary + dashboard). Hạn đang áp dụng = ngày gia hạn + 12 tháng, không có thì ngày hết hạn.
- (WR-08) Số dư dưới 1.000đ (tiền lẻ) không đưa vào báo cáo, liệt kê ở Review.
- (WR-08) Khách nhiều Sales: cột "Jobs by Salesman" tách job + tiền theo từng Sales trên SMS;
  job không thấy trên SMS ghi "Không thấy job trên SMS".

- Có file SMS: SMS chỉ quyết NGÀY từng job, không bao giờ quyết tiền. SMS ghi nhiều
  hơn Misa thì bỏ job cũ nhất và liệt kê cho FIN gạch paid
  (`handoff/docs/NOTE-ke-toan-tra-loi-20261006.md`).
- Total của một khách luôn lấy từ file **Tổng hợp công nợ phải thu**, cột số dư
  cuối kỳ bên Nợ. Khi file Tổng hợp và file Tuổi nợ mâu thuẫn, tiền theo Tổng hợp.
- Danh sách khách = khách có số dư cuối kỳ bên Nợ lớn hơn 0. Khách trả trước
  (số dư bên Có) không đưa vào báo cáo.
- Tổng các cột tuổi nợ của một dòng phải bằng Total của dòng đó. Không thoả
  thì dòng đó bị đánh dấu, không được im lặng cho qua.
- Tuổi nợ tính theo **số ngày thật đã trôi qua**, không tính theo số lần tool
  chạy. Nghỉ mấy tuần rồi chạy lại vẫn phải ra tuổi đúng.
- Các file đầu vào phải cùng kỳ: Chi tiết 131 và Tổng hợp 131 cùng ngày cuối kỳ (v0.4);
  Tuổi nợ = Tổng hợp + 1 ngày (v0.3). Lệch thì dừng và báo lỗi, cấm ra file.

### 5.2b Những thứ không suy ra được từ file Misa

- Hợp đồng mới ký, biên bản bù trừ, khách trả theo tháng — file Misa không có.
  Kế toán khai một lần bằng lệnh `--credit-term` hoặc `--ngoai-le`, tool nhớ mãi.
  Cấm bắt kế toán khai lại mỗi tuần.
- Số job và Salesman chưa lấy được tự động thì mang nguyên của tuần trước sang.
  Cấm xoá trắng: kế toán sửa nhanh hơn là điền lại từ đầu.

### 5.2c Đối chiếu

- Chỉ so báo cáo của tool với bản kế toán làm tay khi hai bên dùng CÙNG một cặp
  file Misa. Số trên Misa đổi theo thời điểm xuất, nên so hai lần xuất khác
  nhau là kết luận sai về tool.
- Mỗi báo cáo phải ghi lại tên và thời điểm sửa của hai file nguồn.

### 5.3 Sổ ghi

- Mỗi lần chạy ghi một trang theo ngày chốt số: từng khách, số dư, cách chia
  nhóm tuổi, và ngày mà mỗi khoản xuất hiện lần đầu.
- Trang của một ngày chốt đã có thì ghi đè đúng trang đó, không tạo thêm bản.
- Khi kế toán có file làm tay cho một tuần, bản của con người thắng bản của máy. Bản người
  làm = file không có tab Review, HOẶC tên có chữ "human", HOẶC file tool ghi mà giờ sửa đã đổi
  (FIN mở ra sửa thẳng rồi lưu — sổ ghi nhớ giờ ghi trong `file_tool`). Cùng ngày chốt thì bản
  người thắng mọi bản tool, kể cả bản sửa sau.
- Mất sổ ghi thì chạy ở chế độ khởi động lạnh và đánh dấu, không được dừng.

### 5.3a Không bao giờ để tool im lặng

- Mọi nhánh dừng sớm đều phải ghi một dòng vào nhật ký nói rõ vì sao. Cấm
  `SystemExit` hay `return` trơn: chạy nền thì không ai thấy màn hình.
- Cấm xoá `settings.json` của máy người dùng khi dọn dẹp. Mất file đó là tool
  đứng im mà không ai biết.

### 5.3b Hai thư mục, không lẫn

- Thư mục mã nguồn chỉ chứa mã, mẫu chuẩn và phòng bàn giao. Cấm để sổ ghi,
  nhật ký, file đầu vào hay báo cáo ở đó, kể cả khi chạy thử.
- Thư mục đích do người dùng trỏ, tool tự tạo đủ các ngăn bên trong. Không ép
  người dùng phải đặt ở OneDrive.

### 5.4 Hai máy, hai môi trường

- Cấm ghi đường dẫn tuyệt đối vào bất kỳ file nào đi theo dự án. Chỗ nào cần
  chỉ đường thì mô tả dấu hiệu nhận dạng, để lúc cài đặt tự dò.
- Cấu hình theo từng máy không đi theo git.

### 5.5 Dashboard

- Nhãn trên biểu đồ dùng tên ngắn (bỏ "CÔNG TY TNHH / CỔ PHẦN / THƯƠNG MẠI / XUẤT NHẬP KHẨU…" ở đầu,
  "VIỆT NAM / LLC / LTD…" ở cuối); tooltip và bảng giữ tên đầy đủ (Chairman 13/10).
- Sửa bất cứ file nào trong `dashboard/src` thì phải chạy `python dashboard/build.py` rồi commit
  cả `dashboard/Credit_Dashboard.html`. Máy FIN không build; nó chỉ nhận bản đã đóng gói qua git.
- Không bắt người dùng chép dashboard bằng tay: tool tự đặt bản mới nhất vào thư mục gốc workspace
  mỗi lần quét.
- Lịch chạy nền gọi thẳng `pythonw.exe runner.py`, không qua .bat (Chairman 09/10: không bật cửa sổ
  CMD). Máy cài trước đó tự được runner đổi lịch một lần (`sua_lich_chay_an`).
- Luật màu (Chairman 08/10): đỏ và các màu lân cận (cam, vàng cam) chỉ dùng cho điều TIÊU CỰC;
  xanh lá cho điều TÍCH CỰC; xanh dương, tím cho TRUNG TÍNH. Cấm dùng đỏ/cam để phân biệt Sales,
  khách hay vendor.
