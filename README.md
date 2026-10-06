<div align="center">

# 📊 Báo cáo công nợ tuần

**Tự động dựng bảng công nợ phải thu hằng tuần cho TRUSTANA Việt Nam**

[![Trustana](https://img.shields.io/badge/TRUSTANA-4d148c?style=for-the-badge&logoColor=white)](#)
[![Phiên bản](https://img.shields.io/badge/phiên%20bản-0.3-ff6200?style=for-the-badge)](#)
[![Phép thử](https://img.shields.io/badge/phép%20thử-47-2ea043?style=for-the-badge)](#-dành-cho-người-bảo-trì)
[![Python](https://img.shields.io/badge/Python-3.9%2B-4d148c?style=for-the-badge&logo=python&logoColor=white)](#)
[![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#)

### Kế toán thả 3 file (2 Misa + 1 SMS) vào một thư mục. Báo cáo tự xuất hiện. Hết.

</div>

---

## 🎯 Ai làm gì

| Vai trò | Việc phải làm | Mất bao lâu | Cần biết code? |
|---|---|---|---|
| 👩‍💼 **Kế toán** | Xuất 2 file từ Misa + 1 file AR-AP từ SMS, thả vào `input` | ~5 phút/tuần | ❌ Không |
| 🧑‍💼 **Sales** | Điền cột *Lý do chưa thu hồi* vào file kết quả | vài phút | ❌ Không |
| 🧑‍🔧 **Người cài đặt** | Nháy đúp `1_Khoi_tao_workspace.bat`, một lần duy nhất | ~5 phút | ❌ Không |
| 👩‍💼 **FIN** | Điền file ghép khách SMS ↔ Misa một lần, bấm `2_Kiem_tra_ghep_khach.bat` | ~30 phút, 1 lần | ❌ Không |
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
    📁 input/                      1️⃣  Đọc kỳ trong file
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
                                   📂 thư mục gốc 🎉 file báo cáo tuần
                                                  🧠 tuần sau tool đọc lại
                                   📄 input/      ✅ file cũ đổi tên [DONE]
```

---

## 📁 Thư mục làm việc

Lúc cài đặt, bạn chọn **thư mục đích** — đặt ở OneDrive, ổ D, ổ mạng đều được.
Bên trong chỉ có 2 ngăn cho người dùng:

```
📂 Cong-No-Workspace/
   ├── 📊 2026_W40_Bang_cong_no_tuan_....xlsx   ← 🎉 báo cáo tuần nằm ngay ở đây
   ├── 📊 w39 final.xlsx                         ← lần đầu: bản final gần nhất của KTT
   ├── 📥 input/    ← 👩‍💼 mỗi tuần thả 2 file Misa + 1 file SMS vào đây
   └── 🔗 sample/   ← file ghép khách SMS ↔ Misa (FIN điền 1 lần) + file KIEM_TRA
```

> 🔁 **Tuần sau tool tự đọc báo cáo tuần trước nằm ở đây** để biết nợ cũ bao nhiêu tuổi.
> Kế toán sửa tay thẳng vào file báo cáo (lý do, credit term, nhóm tuổi) thì tuần sau
> tool tính tiếp theo bản đã sửa. Đừng xoá báo cáo tuần gần nhất.
>
> Tool còn một ngăn ẩn `_tool/` (trí nhớ, nhật ký). Không cần mở, đừng xoá.

---

## 📋 Dùng thế nào

### 👩‍💼 Kế toán — mỗi tuần

1. Xuất **3 file, cùng một kỳ**:
   - Misa: *Tổng hợp công nợ phải thu khách hàng* — kỳ từ thứ Bảy đến thứ Sáu
   - Misa: *Phân tích công nợ phải thu theo tuổi nợ* — đến ngày thứ Bảy
   - SMS: báo cáo *AR-AP* (có cột Partner Code, File/Job No., ETD, ETA, A/R remain)
2. Thả cả 3 file vào `📥 input`. Tên file là gì cũng được — tool đọc bên trong file.
3. Chờ tối đa 15 phút, hoặc nháy đúp `3_Chay_ngay.bat`. Báo cáo xuất hiện trong
   thư mục gốc, ba file đầu vào được đổi tên thành `[DONE] ...`.
4. Kế toán trưởng sửa thẳng vào file báo cáo ở thư mục gốc — tuần sau tool tính
   tiếp theo bản đã sửa (người thắng máy).

> ⏳ Thả 2 file Misa mà chưa có file SMS: tool chờ 2 giờ rồi mới chạy theo luật cũ.

### 🧮 Tool ghép 3 nguồn thế nào

```
Tiền mỗi khách   ← Misa Tổng hợp (luôn luôn)
Tuổi từng job    ← SMS: ngày ETD/ETA + credit term
Nợ cũ ngoài SMS  ← báo cáo tuần trước ở thư mục gốc, cộng thêm ngày đã trôi
Khó đòi          ← file ghép khách, luôn ở nhóm 120+
SMS > Misa       ← khách đã trả mà SMS quên gạch: bỏ job cũ nhất, liệt kê để FIN ấn paid
```

> 🗂️ Thả nhiều tuần cùng lúc cũng được — tool ghép từng cặp cùng kỳ rồi chạy
> lần lượt từ tuần cũ nhất, mỗi tuần một file kết quả.

### 📖 Đọc file kết quả

**Sheet `Receivable`** — bảng công nợ theo mẫu chuẩn, 14 cột.

| Màu dòng | Nghĩa là |
|---|---|
| ⬜ Trắng | Máy lấy thẳng số từ Misa, yên tâm |
| 🟨 Vàng | Máy suy ra từ sổ ghi, nên ngó qua |
| 🟥 Đỏ | Khách mới, hoặc ngoại lệ do kế toán khai — nên xem kỹ |

**Sheet `Review`** — mở phần đầu trước, có sẵn bảng đối chiếu:

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
1️⃣  Cài Python từ python.org            ⚠️ nhớ tick "Add Python to PATH"
2️⃣  Nháy đúp 1_Khoi_tao_workspace.bat   chọn thư mục đích, tool tạo đủ các ngăn
3️⃣  Bỏ file ghép khách (FIN đã điền)    vào sample
    Bỏ bản final gần nhất (KTT duyệt)   vào thư mục gốc — chỉ lần đầu
4️⃣  Nháy đúp 2_Kiem_tra_ghep_khach.bat  Excel mở file KIEM_TRA: dòng cần sửa, khách mới chưa ghép,
                                         và chạy thử từng khách để FIN soát
```

Nút 1 tự cài thư viện, tự đăng ký lịch chạy nền **15 phút một lần** vào Task Scheduler.
Khi có khách mới, nút 2 liệt kê sẵn đúng khuôn cột của file ghép khách — FIN chép dòng
sang, điền ô vàng, lưu lại là xong.

### 🏷️ Khai những thứ file Misa không có — khai một lần, tool nhớ mãi

```bat
:: khách vừa ký hợp đồng, đổi hạn thanh toán
python runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"

:: ngoại lệ nghiệp vụ: lô tàu chìm, xuất lại debit note
python runner.py --ngoai-le "BSF LLC" --nhom "1 - 30" --vi-sao "lo tau chim, xuat lai DN"

:: chờ ký biên bản bù trừ công nợ hai chiều nên để ở nợ trong hạn
python runner.py --ngoai-le "AHC LOGISTICS(XIAMEN) CO.,LTD" --nhom "Current" --vi-sao "cho ky bien ban bu tru"

:: nợ ảo do lỗi xuất hoá đơn, loại khỏi bảng nhưng vẫn ghi ở sheet Review
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
> thì vẫn ra file, tô màu và ghi vào sheet *Review*.

| Hiện tượng | Nghĩa là | Làm gì |
|---|---|---|
| 🟥 File đầu vào bị đổi thành `[LOI] ...` | Dữ liệu không hợp lệ | Mở file `[LOI] ... doc-vi-sao-hong.txt` cùng thư mục, sửa rồi bỏ chữ `[LOI]` khỏi tên file |
| ⏱️ Nhật ký ghi *"Hai file lệch kỳ"* | Hai file xuất ở hai thời điểm khác nhau | Xuất lại cho cùng kỳ rồi thả vào |
| 🕳️ Nhật ký ghi *"chưa có file Tổng hợp cùng kỳ đi kèm"* | Thiếu một trong hai file | Xuất bổ sung file còn thiếu |
| 😴 Không thấy file kết quả | Máy tắt, hoặc lịch chạy chưa gọi | Task Scheduler → *Bao cao cong no tuan* → Run |
| 🤐 Bấm Run mà không có gì xảy ra | Mất `settings.json` nên tool không biết thư mục đích | `python runner.py --xem-cau-hinh`, rồi `--dat-thu-muc` |
| 🔒 Nhật ký ghi *"Tạm hoãn kỳ này"* | File kết quả đang mở trong Excel | Đóng file, lần chạy sau tool làm tiếp |
| 🆘 Tool hỏng cả tuần | | Kế toán làm tay như cũ, thả bản làm tay vào thư mục gốc; tool coi bản của người là đúng và chạy tiếp từ đó |

📜 Nhật ký mỗi lần chạy: `_tool/runner.log` **trong thư mục đích** (ngăn ẩn).

---

## 🧭 Những điều nên biết

- 📅 **Tuổi nợ tính theo ngày thật**, không theo số lần tool chạy. Nghỉ ba tuần
  rồi chạy lại vẫn ra tuổi đúng cho khách cũ.
- 💰 **Tiền luôn lấy theo file Tổng hợp.** Hai file Misa mâu thuẫn thì tiền thắng,
  và khách gây lệch được liệt kê trong sheet *Review*.
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
├── wr/                   misa · ghep · soghi · tinhtoan · baocao · kiemtra · caidat · thuonghieu
├── tools/                make_template · setup_wizard
├── tests/                47 phép thử + so_sanh.py để đối chiếu bản tay
├── template/             mẫu chuẩn, sheet Receivable trắng
└── handoff/              RULES · STATE · MAP · phiếu việc · tài liệu
```

> 🔐 Thư mục `handoff/docs/fixtures/` chứa dữ liệu công nợ thật để chạy phép thử
> đối chiếu. Nó **không đi theo repo**. Không có nó thì 20 phép thử tự bỏ qua,
> 27 phép thử còn lại (dữ liệu giả) vẫn chạy. Chép dữ liệu vào đó là 47 phép thử chạy lại đủ.

```bat
python -m pytest                       :: 47 phép thử (20 cần dữ liệu thật)
python tools/make_template.py          :: dựng lại mẫu chuẩn
python tests/tao_du_lieu_gia.py "<input>"   :: sinh dữ liệu giả để thử
```

📖 Luật tính số: `handoff/docs/SPEC-logic-bao-cao-tuan.md`
📐 Luật phối hợp: `handoff/RULES.md` — đọc trước khi sửa bất cứ thứ gì

<div align="center">

**TRUSTANA VIỆT NAM** · tool đầu tiên trong chuỗi credit control

</div>
