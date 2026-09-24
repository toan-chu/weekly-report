<div align="center">

# 📊 Báo cáo công nợ tuần

**Tự động dựng bảng công nợ phải thu hằng tuần cho TRUSTANA Việt Nam**

[![Trustana](https://img.shields.io/badge/TRUSTANA-4d148c?style=for-the-badge&logoColor=white)](#)
[![Phiên bản](https://img.shields.io/badge/phiên%20bản-0.2-ff6200?style=for-the-badge)](#)
[![Phép thử](https://img.shields.io/badge/phép%20thử-29%20đạt-2ea043?style=for-the-badge)](#-dành-cho-người-bảo-trì)
[![Python](https://img.shields.io/badge/Python-3.9%2B-4d148c?style=for-the-badge&logo=python&logoColor=white)](#)
[![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#)

### Kế toán thả 2 file Misa vào một thư mục. Báo cáo tự xuất hiện. Hết.

</div>

---

## 🎯 Ai làm gì

| Vai trò | Việc phải làm | Mất bao lâu | Cần biết code? |
|---|---|---|---|
| 👩‍💼 **Kế toán** | Xuất 2 file từ Misa, thả vào thư mục `01_Input` | ~5 phút/tuần | ❌ Không |
| 🧑‍💼 **Sales** | Điền cột *Lý do chưa thu hồi* vào file kết quả | vài phút | ❌ Không |
| 🧑‍🔧 **Người cài đặt** | Nháy đúp `setup.bat`, một lần duy nhất | ~5 phút | ❌ Không |
| 🤖 **Tool** | Đọc file, tính tuổi nợ, điền mẫu chuẩn, nhớ cho tuần sau | 2 giây | — |

> 💡 Trước đây bước "tính tuổi nợ và điền bảng" mất vài tiếng mỗi tuần và phải nhớ
> số liệu của tuần trước. Giờ máy làm, kế toán chỉ còn việc xuất file.

---

## 🔄 Luồng chạy

```
        👩‍💼 KẾ TOÁN                          🤖 TOOL                        📤 KẾT QUẢ
   ─────────────────────            ──────────────────────          ────────────────────

   Xuất 2 file từ Misa
   ┌──────────────────┐
   │ Tổng hợp công nợ │  ──┐
   │ phải thu         │    │
   └──────────────────┘    │      ╔══════════════════════╗
   ┌──────────────────┐    ├────► ║  15 phút thức dậy    ║
   │ Phân tích công nợ│  ──┘      ║  một lần, thấy file  ║
   │ theo tuổi nợ     │           ║  mới thì bắt tay     ║
   └──────────────────┘           ╚══════════╤═══════════╝
            │                                │
            ▼                                ▼
    📁 01_Input/                   1️⃣  Đọc kỳ trong file
    (thả file vào đây)                 hai file lệch kỳ → dừng, báo lỗi
                                   2️⃣  Danh sách khách + cột Total
                                       ← file Tổng hợp (số tiền đúng)
                                   3️⃣  Chia 6 nhóm tuổi nợ
                                       ← file Tuổi nợ + sổ ghi tuần trước
                                   4️⃣  Chép sang Credit Term, Salesman,
                                       số job, Lý do của tuần trước
                                   5️⃣  Tự kiểm: tổng dòng = Total
                                       tổng bảng = tổng Misa
                                                │
                                                ▼
                                   📁 02_Output/  🎉 file báo cáo tuần
                                   📁 so-ghi/     🧠 nhớ cho tuần sau
                                   📄 01_Input/   ✅ file cũ đổi tên [DONE]
```

---

## 📁 Thư mục làm việc

Lúc cài đặt, bạn chọn **thư mục đích** — đặt ở OneDrive, ổ D, ổ mạng đều được.
Tool tự tạo đủ các ngăn bên trong:

```
📂 Cong-No-Workspace/
   ├── 📥 01_Input/    ← 👩‍💼 kế toán thả 2 file Misa vào đây
   ├── 📤 02_Output/   ← 🎉 báo cáo tuần xuất hiện ở đây
   ├── ✍️  03_BanTay/   ← bản kế toán tự làm, khi cần đối chiếu
   ├── 🧠 so-ghi/      ← trí nhớ của tool, đừng xoá
   ├── 🗃️  luu-tru/     ← bản cũ
   └── 📜 log/         ← nhật ký mỗi lần chạy
```

> ⚠️ **Đừng xoá thư mục `so-ghi/`.** Đó là chỗ tool nhớ tuần trước khách nợ bao
> nhiêu và mỗi khoản nợ từ ngày nào. Mất nó thì tuần sau tool phải đoán lại từ đầu.

---

## 📋 Dùng thế nào

### 👩‍💼 Kế toán — mỗi tuần

1. Xuất từ Misa 2 file, **cùng một kỳ**:
   - *Tổng hợp công nợ phải thu khách hàng* — kỳ từ thứ Bảy đến thứ Sáu
   - *Phân tích công nợ phải thu theo tuổi nợ* — đến ngày thứ Bảy
2. Thả cả 2 file vào `📥 01_Input`. Tên file là gì cũng được, kể cả Windows tự
   thêm `(1)`, `(2)` — tool đọc ngày bên trong file, không nhìn tên.
3. Chờ tối đa 15 phút. Báo cáo xuất hiện trong `📤 02_Output`, hai file đầu vào
   được đổi tên thành `[DONE] ...` cho biết đã xử lý xong.

> 🗂️ Thả nhiều tuần cùng lúc cũng được — tool ghép từng cặp cùng kỳ rồi chạy
> lần lượt từ tuần cũ nhất, mỗi tuần một file kết quả.

### 📖 Đọc file kết quả

**Sheet `Receivable`** — bảng công nợ theo mẫu chuẩn, 14 cột.

| Màu dòng | Nghĩa là |
|---|---|
| ⬜ Trắng | Máy lấy thẳng số từ Misa, yên tâm |
| 🟨 Vàng | Máy suy ra từ sổ ghi, nên ngó qua |
| 🟥 Đỏ | Khách mới, hoặc ngoại lệ do kế toán khai — nên xem kỹ |

**Sheet `Cần xem lại`** — mở phần đầu trước, có sẵn bảng đối chiếu:

```
Tổng bảng báo cáo này                    1.635.150.000   ✅ phải bằng dòng dưới
Tổng file Tổng hợp công nợ (dư Nợ)       1.635.150.000   ✅ KHỚP
Tổng file Phân tích tuổi nợ              1.905.150.000   ℹ️ KHÔNG dùng để đối chiếu
Chênh lệch giữa hai file Misa              270.000.000   do khách liệt kê bên dưới
Nợ trong hạn trên báo cáo                  426.400.000
  · lấy thẳng từ file Tuổi nợ              228.600.000
  · chưa xác định được tuổi                197.800.000
```

> ❓ **Vì sao tổng báo cáo không bằng tổng file Tuổi nợ?** Vì tiền luôn lấy theo
> file **Tổng hợp**. File Tuổi nợ thường cao hơn do còn hoá đơn thực tế đã thu
> nhưng chưa gắn được chứng từ thanh toán. Hai con số này không bao giờ bằng nhau.

---

## 🚀 Cài đặt — một lần duy nhất

```
1️⃣  Cài Python từ python.org      ⚠️ nhớ tick "Add Python to PATH"
2️⃣  Nháy đúp setup.bat            wizard hỏi thư mục đích, bạn gõ số để chọn
3️⃣  Nạp trạng thái ban đầu        từ báo cáo tuần gần nhất làm tay
```

```bat
python runner.py --nap-bao-cao "2026_W38_Bang_cong_no_tuan_....xlsx" --ngay-chot 2026-09-19
```

Wizard tự dò các thư mục OneDrive trên máy, tự tạo các ngăn, tự cài thư viện,
tự đăng ký lịch chạy nền **15 phút một lần** vào Task Scheduler.

### 🏷️ Khai những thứ file Misa không có — khai một lần, tool nhớ mãi

```bat
:: khách vừa ký hợp đồng, đổi hạn thanh toán
python runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"

:: ngoại lệ nghiệp vụ: lô tàu chìm, xuất lại debit note
python runner.py --ngoai-le "BSF LLC" --nhom "1 - 30" --vi-sao "lo tau chim, xuat lai DN"

:: chờ ký biên bản bù trừ công nợ hai chiều nên để ở nợ trong hạn
python runner.py --ngoai-le "AHC LOGISTICS(XIAMEN) CO.,LTD" --nhom "Current" --vi-sao "cho ky bien ban bu tru"

:: nợ ảo do lỗi xuất hoá đơn, loại khỏi bảng nhưng vẫn ghi ở sheet Cần xem lại
python runner.py --bo-qua "OPTIMALOG LLC" --vi-sao "no ao do loi xuat hoa don tien coc"
python runner.py --nhan-lai "OPTIMALOG LLC"      :: khi hết chuyện
```

### ⚙️ Đổi thư mục đích, đổi nhịp chạy

```bat
python runner.py --xem-cau-hinh
python runner.py --dat-thu-muc "D:\Cong ty\Cong no"

schtasks /change /tn "Bao cao cong no tuan" /ri 15      :: 15 phút một lần
schtasks /query  /tn "Bao cao cong no tuan"
```

---

## ❓ Hỏng thì làm gì

> 🛟 **Nguyên tắc: tool không bao giờ chặn không ra báo cáo.** Chỗ nào không chắc
> thì vẫn ra file, tô màu và ghi vào sheet *Cần xem lại*.

| Hiện tượng | Nghĩa là | Làm gì |
|---|---|---|
| 🟥 File đầu vào bị đổi thành `[LOI] ...` | Dữ liệu không hợp lệ | Mở file `[LOI] ... doc-vi-sao-hong.txt` cùng thư mục, sửa rồi bỏ chữ `[LOI]` khỏi tên file |
| ⏱️ Nhật ký ghi *"Hai file lệch kỳ"* | Hai file xuất ở hai thời điểm khác nhau | Xuất lại cho cùng kỳ rồi thả vào |
| 🕳️ Nhật ký ghi *"chưa có file Tổng hợp cùng kỳ đi kèm"* | Thiếu một trong hai file | Xuất bổ sung file còn thiếu |
| 😴 Không thấy file kết quả | Máy tắt, hoặc lịch chạy chưa gọi | Task Scheduler → *Bao cao cong no tuan* → Run |
| 🤐 Bấm Run mà không có gì xảy ra | Mất `settings.json` nên tool không biết thư mục đích | `python runner.py --xem-cau-hinh`, rồi `--dat-thu-muc` |
| 🔒 Nhật ký ghi *"Tạm hoãn kỳ này"* | File kết quả đang mở trong Excel | Đóng file, lần chạy sau tool làm tiếp |
| 🆘 Tool hỏng cả tuần | | Kế toán làm tay như cũ, thả bản làm tay vào `03_BanTay`; tool coi bản của người là đúng và chạy tiếp từ đó |

📜 Nhật ký mỗi lần chạy: `log/runner.log` **trong thư mục đích**.

---

## 🧭 Những điều nên biết

- 📅 **Tuổi nợ tính theo ngày thật**, không theo số lần tool chạy. Nghỉ ba tuần
  rồi chạy lại vẫn ra tuổi đúng cho khách cũ.
- 💰 **Tiền luôn lấy theo file Tổng hợp.** Hai file Misa mâu thuẫn thì tiền thắng,
  và khách gây lệch được liệt kê trong sheet *Cần xem lại*.
- 🕰️ **Số trên Misa đổi theo thời điểm xuất** — ghi nhận lùi, hoá đơn thay thế.
  Vì vậy chỉ so báo cáo của tool với bản làm tay khi cả hai dùng **cùng một cặp
  file**. Báo cáo có ghi tên và giờ sửa của hai file nguồn để đối chiếu sau này.
- ✍️ **Cột Salesman và số job** hiện được mang nguyên từ tuần trước sang, kế toán
  tự sửa. Lấy tự động từ SMS là việc của phase sau.

---

## 🛠️ Dành cho người bảo trì

```
weekly-report/            chỉ chứa mã nguồn, không có dữ liệu chạy
├── runner.py             điểm chạy, cũng là chỗ nhận mọi lệnh
├── wr/                   misa · soghi · tinhtoan · baocao · caidat · thuonghieu
├── tools/                make_template · setup_wizard
├── tests/                29 phép thử + so_sanh.py để đối chiếu bản tay
├── template/             mẫu chuẩn, sheet Receivable trắng
└── handoff/              RULES · STATE · MAP · phiếu việc · tài liệu
```

> 🔐 Thư mục `handoff/docs/fixtures/` chứa dữ liệu công nợ thật để chạy phép thử
> đối chiếu. Nó **không đi theo repo**. Không có nó thì 20 phép thử tự bỏ qua,
> 9 phép thử còn lại vẫn chạy. Chép dữ liệu vào đó là 29 phép thử chạy lại đủ.

```bat
python -m pytest                       :: 29 phép thử (20 cần dữ liệu thật)
python tools/make_template.py          :: dựng lại mẫu chuẩn
python tests/tao_du_lieu_gia.py "<01_Input>"   :: sinh dữ liệu giả để thử
```

📖 Luật tính số: `handoff/docs/SPEC-logic-bao-cao-tuan.md`
📐 Luật phối hợp: `handoff/RULES.md` — đọc trước khi sửa bất cứ thứ gì

<div align="center">

**TRUSTANA VIỆT NAM** · tool đầu tiên trong chuỗi credit control

</div>
