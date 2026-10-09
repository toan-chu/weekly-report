<div align="center">

# 📊 Báo cáo công nợ tuần — Full Credit Report

**Tự dựng báo cáo phải thu, phải trả và dòng tiền hằng tuần cho TRUSTANA Việt Nam**

[![Trustana](https://img.shields.io/badge/TRUSTANA-4d148c?style=for-the-badge&logoColor=white)](#)
[![Phiên bản](https://img.shields.io/badge/phiên%20bản-0.5-ff6200?style=for-the-badge)](#)
[![Phép thử](https://img.shields.io/badge/phép%20thử-78-2ea043?style=for-the-badge)](#-dành-cho-người-bảo-trì)
[![Python](https://img.shields.io/badge/Python-3.9%2B-4d148c?style=for-the-badge&logo=python&logoColor=white)](#)
[![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#)

### Kế toán thả 6 file (5 Misa + 1 SMS) vào một thư mục. Báo cáo tự xuất hiện. Hết.

</div>

> **Một file Excel mỗi tuần cho sếp thấy: tiền khách còn nợ và đã trả, tiền mình còn nợ
> vendor và đã trả, khách nào đáng lo.**
> Kế toán thả file → tool chạy ngầm → chị kế toán trưởng mở file ở thư mục gốc (sau này kéo
> vào dashboard) để trình sếp.

---

## 🎯 Ai làm gì

| Vai trò | Việc phải làm | Mất bao lâu | Cần biết code? |
|---|---|---|---|
| 👩‍💼 **FIN** | Xuất 5 file Misa + 1 file AR-AP từ SMS, **cùng lúc**, thả vào `input` | ~10 phút/tuần | ❌ Không |
| 👩‍💼 **FIN** | Xem cột **Đối chiếu** + tab Review, sửa chỗ ĐỎ / CAM trên Misa | tuỳ tuần | ❌ Không |
| 👩‍💼 **FIN** | Điền file ghép khách + bảng job nợ cũ một lần, bấm `2_Kiem_tra_ghep_khach.bat` | ~30 phút, 1 lần | ❌ Không |
| 👩‍💼 **Kế toán trưởng** | Đọc báo cáo, sửa lý do trễ hạn ngay trên file | tuỳ | ❌ Không |
| 🧑‍🔧 **Người cài đặt** | Nháy đúp `1_Khoi_tao_workspace.bat`, một lần duy nhất | ~5 phút | ❌ Không |
| 🤖 **Tool** | Đọc file, tính tuổi nợ từ hoá đơn, phân nhóm rủi ro, ghi báo cáo | ~10 giây | — |

---

## 🔄 Luồng chạy

```
  👩‍💼 FIN (mỗi tuần, xuất cùng lúc)              🤖 TOOL (15 phút thức dậy 1 lần, chạy ẩn)
  ─────────────────────────────────            ──────────────────────────────────
  Misa  Tổng hợp công nợ phải thu    ─┐        1️⃣ nhận dạng file theo nội dung; thiếu file → chờ 2 giờ
  Misa  Chi tiết công nợ phải thu     │        2️⃣ tiền mỗi khách ← Tổng hợp phải thu
        (từ 01/01/2022)               │        3️⃣ tuổi nợ ← từng hoá đơn trong file Chi tiết
  Misa  Tổng hợp công nợ phải trả     ├──► 📥  4️⃣ job chưa thanh toán ← Bán hàng + Sổ chi tiết TK131
  Misa  Bán hàng (từ đầu năm)         │  input    Salesman của job ← SMS
  Misa  Sổ chi tiết TK131 (đầu năm)   │        5️⃣ đối chiếu: tổng job vs dư nợ → tô màu cho FIN
  SMS   AR-AP                        ─┘        6️⃣ nhóm rủi ro ← luật của kế toán trưởng
                                               7️⃣ phải trả (số tổng), dòng tiền, so với tuần trước
  🔗 sample/: ghép khách + job nợ cũ (1 lần)
  📊 gốc/: báo cáo các tuần trước ───────────►    (đọc lại bản kế toán đã sửa)
                                                          │
                                                          ▼
                                   📊 gốc/2026_W41_Credit_Report_....xlsx   ✅ input đổi tên [DONE]
```

---

## 📁 Thư mục làm việc

```
📂 Cong-No-Workspace/
   ├── 📊 2026_W40_Credit_Report_....xlsx   ← 🎉 báo cáo tuần nằm ngay ở đây; sửa thẳng vào đây
   ├── 🌐 Credit_Dashboard.html             ← mở bằng Chrome/Edge, kéo báo cáo tuần vào
   ├── 📥 input/    ← mỗi tuần thả 6 file vào đây; xong tool đổi tên thành [DONE] ...
   ├── 🔗 sample/   ← những file FIN điền MỘT LẦN, tool đọc lại mỗi tuần (bảng dưới)
   └── 🙈 _tool/    ← ẩn: trí nhớ, nhật ký. Không cần mở, đừng xoá
```

**Các file trong `sample/`**

| File | Ai tạo | Ý nghĩa | FIN làm gì |
|---|---|---|---|
| `SAMPLE_ghep-khach-SMS-Misa_FIN-da-dien.xlsx` | FIN | Mã khách bên SMS ứng với mã khách nào bên Misa (tên hai bên khác nhau) | Có khách mới thì thêm dòng (file KIEM_TRA chỉ ra dòng nào) |
| `Tham_chieu_job_no_cu.xlsx` (tên nào cũng được, tool nhận theo nội dung) | Tool, lần đầu bấm nút 2 — gói W40 có sẵn | Hoá đơn **trước năm nay** còn nợ mà Misa không cho biết job | Điền cột Job, Sales (ô vàng) một lần |
| `KIEM_TRA_ghep_khach_<ngày giờ>.xlsx` | Tool, mỗi lần bấm nút 2 | Kết quả kiểm tra: khách mới chưa ghép, dòng ghép sai, nợ cũ chưa có job | Chỉ đọc; sửa ở hai file trên |

> 🔁 Tuần sau tool đọc lại báo cáo tuần trước ở thư mục gốc để lấy lý do trễ hạn, credit term
> và làm mốc so sánh. Bản kế toán trưởng tự làm (không có tab Review) luôn thắng bản tool
> cùng tuần. Đừng xoá báo cáo tuần gần nhất.

---

## 📋 Dùng thế nào

### 👩‍💼 FIN — mỗi tuần

**Xuất 6 file trong cùng một buổi** (số Misa đổi theo lúc xuất: tool so được hai file lệch nhau là
do xuất khác lúc, nhưng xuất cùng lúc thì không phải dò). Kỳ báo cáo: **thứ Bảy → thứ Sáu**.
Tên file là gì cũng được — tool đọc tiêu đề bên trong file.

| # | Lấy ở | Báo cáo (tiêu đề trong file) | Chọn khi xuất | Tool dùng để |
|---|---|---|---|---|
| 1 | Misa | **Tổng hợp công nợ phải thu khách hàng** | TK 131 · loại tiền Tổng hợp · thứ Bảy → thứ Sáu | Tiền còn nợ của từng khách (số chuẩn) |
| 2 | Misa | **Chi tiết công nợ phải thu khách hàng** | TK 131 · **từ 01/01/2022** → thứ Sáu | Tuổi nợ từng hoá đơn, tiền thu trong tuần |
| 3 | Misa | **Tổng hợp công nợ phải trả nhà cung cấp** | TK 331 · thứ Bảy → thứ Sáu | Tab Payable (số tổng), tiền đã trả vendor |
| 4 | Misa | **Bán hàng** (danh sách chứng từ bán hàng) | **Từ 01/01 năm nay** → hôm nay · có cột *TT thanh toán* | Hoá đơn nào chưa thanh toán |
| 5 | Misa | **Sổ chi tiết tài khoản** | TK 131 · **từ 01/01 năm nay** → hôm nay · có cột *Mã đối tượng THCP* | Hoá đơn thuộc job nào; phiếu thu sau kỳ |
| 6 | SMS | **AR-AP** | Như mọi tuần | Tên Sales của từng job |

Rồi:

1. Thả cả 6 file vào `📥 input`.
2. Chờ tối đa 15 phút (tool chạy ẩn, không hiện cửa sổ). Muốn có ngay: mở **Task Scheduler** →
   *Bao cao cong no tuan* → **Run**.
3. Mở báo cáo ở thư mục gốc → tab **Review** → xử lý theo màu:

| Màu ở cột *Đối chiếu job* | Nghĩa là | FIN làm gì |
|---|---|---|
| 🟩 Khớp | Tổng job chưa thanh toán = dư nợ | Không cần làm gì |
| 🟥 ĐỎ | Job chưa thanh toán **nhiều hơn** dư nợ | Đã thu mà chưa đối trừ vào hoá đơn, hoặc thanh toán một phần Misa chưa trừ → đối trừ trên Misa |
| 🟨 VÀNG | Lệch nhỏ (≤ 2%) | Nghi quên đánh giá chênh lệch tỷ giá |
| 🟧 CAM | Dư nợ **nhiều hơn** job chưa thanh toán | Nợ trước năm nay → điền bảng job nợ cũ; hoặc hai file xuất khác lúc (cột ghi rõ) → xuất lại cùng lúc |
| ⬜ Khó đòi | Nợ khó đòi đã khai | Không cần làm gì |

Review còn liệt kê: khách Misa đã hết nợ mà hoá đơn vẫn ghi "chưa thanh toán" (ĐỎ — đối trừ trên
Misa), và tên khách trên file Bán hàng không tìm thấy trên Misa (TÍM — kiểm tra tên).

### 📖 Đọc file kết quả

| Sheet | Để làm gì |
|---|---|
| **Summary** | Trang bìa: các ô số chính, tuần này so với tuần trước, lịch sử các tuần |
| **Receivables** | Bảng công nợ theo mẫu chuẩn. Bấm **+** (hoặc nút **2** góc trái) để mở các job chưa thanh toán của khách (lấy từ Misa): hoá đơn nào, Sales nào, còn bao nhiêu. Cột cuối **Đối chiếu job**: Khớp / ĐỎ / VÀNG / CAM — chỗ FIN cần xem lại |
| **AR Risk** | Đầu kỳ, tăng, giảm, cuối kỳ; dấu hiệu S1–S6; nhóm rủi ro N0–N4; việc cần làm |
| **Payable** | Theo mẫu của FIN: vendor còn nợ, chỉ số tổng (Current = Total), không phân tích |
| **Cash Flow** | Tiền khách trả trong tuần (ngân hàng, tiền mặt, cấn trừ, tỷ giá) và tiền trả vendor |
| **Invoices** | Từng hoá đơn còn nợ — lọc một khách, cộng lại đúng bằng số ở Receivables |
| **Methodology** | Từng con số tính thế nào, lấy từ file nào — để trả lời khi sếp hỏi |
| **Review** | Đối chiếu tổng, cảnh báo, những chỗ tool không chắc |

```
Tuổi nợ   = hạn của từng hoá đơn (ngày hoá đơn + credit term), quá hạn tính từ HÔM SAU ngày đến hạn
Tiền thu  = phiếu thu ghi số hoá đơn thì trừ đúng hoá đơn đó; còn lại trừ hoá đơn cũ nhất trước
Khó đòi   = luôn ở nhóm 120+
Tiền      = luôn theo file Tổng hợp; file Chi tiết lệch thì ghi ở Review
Job       = mỗi hoá đơn KHÔNG ghi "Đã thanh toán" trên Bán hàng; tiền = tiền hoá đơn, hoá đơn nhiều job
            tách theo dòng hàng của từng job trên Sổ chi tiết. "Đã thanh toán" mà phiếu thu nằm sau
            ngày cuối kỳ thì vẫn tính là nợ của kỳ.
```

### 📊 Dashboard

Thư mục gốc workspace luôn có `Credit_Dashboard.html`. Mở bằng Chrome hoặc Edge, kéo file báo cáo
tuần vào trang. Mỗi tab là một trang riêng: Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả ·
Dòng tiền. Bấm một khách hoặc một Sales để vào trang chi tiết; các tab khác lọc theo luôn, bỏ lọc
bằng dấu ✕ trên thanh lọc, nút Back của trình duyệt quay lại được. **In / Lưu PDF** in đủ mọi trang
để gửi sếp (chọn "Save as PDF").
Không cần mạng, dữ liệu không rời khỏi máy.

---

## 🚀 Cài đặt — một lần duy nhất

```
1️⃣  Cài Python từ python.org            ⚠️ nhớ tick "Add Python to PATH"
2️⃣  Nháy đúp 1_Khoi_tao_workspace.bat   chọn thư mục đích, tool tạo các ngăn + lịch 15 phút chạy ẩn
3️⃣  Bỏ file ghép khách (FIN đã điền)    vào sample
    Bỏ bản final gần nhất (KTT duyệt)   vào thư mục gốc — chỉ lần đầu, để có mốc so sánh
4️⃣  Nháy đúp 2_Kiem_tra_ghep_khach.bat  Excel mở file KIEM_TRA; lần đầu tool tạo luôn
                                        sample/Tham_chieu_job_no_cu.xlsx để FIN điền job nợ cũ
```

### 📦 Bắt đầu trên máy FIN bằng gói tuần 40

Gói `sample_w40` (zip) có đủ mọi file để chạy lại tuần 40, làm mốc cho tuần 41:

| File trong gói | Bỏ vào | Là gì |
|---|---|---|
| `W39_bao-cao-final_nguoi-lam.xlsx` | thư mục gốc | Bản final W39 kế toán trưởng duyệt — mốc "tuần trước" |
| `SAMPLE_ghep-khach-SMS-Misa_FIN-da-dien.xlsx` | `sample/` | File ghép khách FIN đã điền (đã sửa TTND VIETNGA → KH00416) |
| `SAMPLE_Tham_chieu_job_no_cu.xlsx` | `sample/` | Bảng job nợ cũ tool đã dựng sẵn từ số W40 (1 hoá đơn: Minh Khôi 32tr) — FIN điền Job, Sales ở ô vàng |
| `W40_misa-tong-hop-phai-thu-131.xlsx` | `input/` | File 1 tuần 40 |
| `W40_misa-chi-tiet-phai-thu-131_tu-2022.xlsx` | `input/` | File 2 tuần 40 |
| `W40_misa-tong-hop-phai-tra-331.xlsx` | `input/` | File 3 tuần 40 |
| `W40_misa-ban-hang_tu-dau-nam.xlsx` | `input/` | File 4 tuần 40 |
| `W40_misa-so-chi-tiet-131_tu-dau-nam.xlsx` | `input/` | File 5 tuần 40 |
| `W40_sms-AR-AP.xlsx` | `input/` | File 6 tuần 40 |

Thứ tự: cài đặt (bước 1–2 ở trên) → giải nén, bỏ file đúng chỗ → điền ô vàng trong
`SAMPLE_Tham_chieu_job_no_cu.xlsx` (làm sau cũng được) → Task Scheduler → **Run** → có báo cáo W40 ở
thư mục gốc. Bảng job nợ cũ không nằm trong git vì chứa số liệu khách; máy nào chưa có thì bấm
`2_Kiem_tra_ghep_khach.bat` là tool tự tạo.
Từ tuần 41: chỉ thả 6 file mới vào `input/`.

### 🏷️ Khai những thứ file Misa không có — khai một lần, tool nhớ mãi

```bat
python runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"
python runner.py --ngoai-le "BSF LLC" --nhom "1 - 30" --vi-sao "lo tau chim, xuat lai DN"
python runner.py --bo-qua "OPTIMALOG LLC" --vi-sao "no ao do loi xuat hoa don tien coc"
python runner.py --nhan-lai "OPTIMALOG LLC"
```

---

## ❓ Hỏng thì làm gì

> 🛟 Tool không bao giờ chặn không ra báo cáo. Chỗ nào không chắc thì vẫn ra file và ghi ở *Review*.

| Hiện tượng | Nghĩa là | Làm gì |
|---|---|---|
| 🟥 File đầu vào bị đổi thành `[LOI] ...` | Dữ liệu không hợp lệ | Đọc file `[LOI] ... doc-vi-sao-hong.txt` cùng thư mục, sửa rồi bỏ chữ `[LOI]` |
| ⏳ Đã thả file mà chưa có báo cáo | Đang chờ file còn thiếu trong 6 file (tối đa 2 giờ) | Thả nốt file còn thiếu; nhật ký ghi tên file đang chờ |
| 📅 File Chi tiết 131 xuất tới cuối tháng, Tổng hợp tới thứ Sáu | Bình thường (FIN W41) | Không cần làm gì: tool tự bỏ phần sau ngày cuối kỳ, nhật ký ghi số bút toán đã bỏ |
| 🟧 Đối chiếu ghi *"hai file xuất khác lúc"* | Misa sửa hoá đơn giữa hai lần xuất | Xuất lại cả 6 file cùng lúc |
| 🪟 Lịch chạy vẫn bật cửa sổ đen | Máy cài trước 09/10 | Lần chạy kế tiếp tool tự sửa; nếu vẫn bật, chạy lại `1_Khoi_tao_workspace.bat` |
| ⚠️ Review ghi *"Chi tiết 131 và Tổng hợp 131 lệch"* | Hai file xuất khác thời điểm | Xuất lại cả hai cùng lúc |
| 😴 Không thấy file kết quả | Máy tắt, hoặc lịch chưa gọi | Task Scheduler → *Bao cao cong no tuan* → Run |
| 🔒 Nhật ký ghi *"Tạm hoãn kỳ này"* | File kết quả đang mở trong Excel | Đóng file, lần chạy sau tool làm tiếp |

📜 Nhật ký: `_tool/runner.log` trong thư mục đích.

---

## 🛠️ Dành cho người bảo trì

```
weekly-report/            chỉ chứa mã nguồn, không có dữ liệu chạy
├── runner.py             điểm chạy, cũng là chỗ nhận mọi lệnh
├── wr/                   misa · chitiet · banhang · tindung · ghifull · ghep · soghi · tinhtoan · baocao · kiemtra
├── tools/                make_template · setup_wizard
├── tests/                78 phép thử (20 cần dữ liệu thật trong handoff/docs/fixtures/)
├── dashboard/            src + build.py → Credit_Dashboard.html (tool chép vào workspace)
├── template/             mẫu chuẩn sheet Receivables
└── handoff/              RULES · STATE · MAP · phiếu việc · tài liệu
```

```bat
python -m pytest                       :: 78 phép thử
python runner.py --tong-hop A.xlsx --chi-tiet B.xlsx --phai-tra C.xlsx --sms D.xlsx ^
                 --ban-hang E.xlsx --so-chi-tiet F.xlsx --ra <thư mục>
python dashboard/build.py              :: sửa dashboard/src thì dựng lại rồi commit cả file html
```

📐 Kiến trúc: `handoff/MAP.md` · Luật phối hợp: `handoff/RULES.md` — đọc trước khi sửa

### 🆕 Có gì mới ở v0.5 (09/10/2026)

- **Job lấy từ Misa** (Bán hàng + Sổ chi tiết TK131) thay cho SMS: tiền theo tỷ giá hoá đơn nên khớp sổ; SMS chỉ còn cho tên Sales (WR-06).
- **Cột Đối chiếu job** tô màu + gợi ý nguyên nhân; Review gom theo màu cho FIN (WR-06).
- **Payable theo mẫu FIN**: chỉ số tổng, không phân tích, không tô màu.
- **Dashboard nhiều trang**: mỗi tab một trang, bấm khách / Sales vào trang chi tiết, bộ lọc đi theo, nút Back chạy (WR-07).
- **Chạy ẩn**: Task Scheduler không bật cửa sổ CMD nữa (WR-07).

<div align="center">

**TRUSTANA VIỆT NAM** · tool đầu tiên trong chuỗi credit control

</div>
