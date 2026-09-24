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
  `setup.bat` có cài được lịch chạy không, Task Scheduler có gọi đúng không,
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

- Thao tác duy nhất của kế toán là thả 2 file Misa vào thư mục `01_Input`.
  Mọi thiết kế ép kế toán mở file ra sửa, điền, hay chạy lệnh đều là vi phạm.
- Tool không bao giờ chặn không ra báo cáo vì thiếu thông tin. Không chắc chỗ
  nào thì vẫn ra file, đánh dấu chỗ đó vào sheet "Cần xem lại".
- Tool không xoá và không sửa đè file nào của kế toán. File cũ chuyển sang
  thư mục lưu trữ, không biến mất.

### 5.2 Số liệu

- Total của một khách luôn lấy từ file **Tổng hợp công nợ phải thu**, cột số dư
  cuối kỳ bên Nợ. Khi file Tổng hợp và file Tuổi nợ mâu thuẫn, tiền theo Tổng hợp.
- Danh sách khách = khách có số dư cuối kỳ bên Nợ lớn hơn 0. Khách trả trước
  (số dư bên Có) không đưa vào báo cáo.
- Tổng các cột tuổi nợ của một dòng phải bằng Total của dòng đó. Không thoả
  thì dòng đó bị đánh dấu, không được im lặng cho qua.
- Tuổi nợ tính theo **số ngày thật đã trôi qua**, không tính theo số lần tool
  chạy. Nghỉ mấy tuần rồi chạy lại vẫn phải ra tuổi đúng.
- Hai file đầu vào phải cùng kỳ: ngày "đến ngày" của file Tuổi nợ bằng ngày
  cuối kỳ của file Tổng hợp cộng một ngày. Lệch thì dừng và báo lỗi, cấm ra file.

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
- Khi kế toán có file làm tay cho một tuần, bản của con người thắng bản của máy.
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
