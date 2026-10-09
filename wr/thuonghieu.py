# -*- coding: utf-8 -*-
"""Màu và phông của thương hiệu Trustana, dùng chung cho mọi file Excel tool sinh ra."""

TIM = "4D148C"      # cột số liệu công nợ
CAM = "FF6200"      # cột định danh khách hàng
TRANG = "FFFFFF"
XAM_NHAT = "F2F2F2"
VANG_CANH_BAO = "FFF2CC"
DO_CANH_BAO = "FCE4E4"

# Chữ trên báo cáo — báo cáo có lúc gửi đối tác nước ngoài nên để tiếng Anh.
# Sửa chữ ở đây là đủ, tool ghi lại các dòng này mỗi lần chạy.
CONG_TY = "TRUSTANA VIETNAM"
TEN_BAO_CAO = "Overdue Debtors Reports"
TEN_BANG = "Accounts receivables"
SHEET_BANG = "Receivables"
SHEET_REVIEW = "Review"
# v0.4 — các sheet còn lại của full credit report (dashboard WR-05 đọc theo đúng tên này)
SHEET_SUMMARY = "Summary"
SHEET_RISK = "AR Risk"
SHEET_PAYABLE = "Payable"
SHEET_CASH = "Cash Flow"
SHEET_METHOD = "Methodology"
SHEET_INVOICES = "Invoices"
TEN_BAO_CAO_TRA = "Creditors Report"
TEN_BANG_TRA = "Accounts payable"

PHONG_TIEU_DE = "Roboto"   # heading
PHONG_NOI_DUNG = "Calibri" # body

# Cột định danh (nền cam) và cột tiền (nền tím) của sheet Receivable
COT = [
    ("No.", "ten", 6),
    ("Code", "ten", 16),
    ("Customer's Name", "ten", 46),
    ("Credit Term", "ten", 14),
    ("Total", "tien", 16),
    ("Current", "tien", 15),
    ("1 - 30", "tien", 15),
    ("31 - 60", "tien", 15),
    ("61 - 90", "tien", 15),
    ("91 - 120", "tien", 15),
    ("120+", "tien", 15),
    ("Salesman", "ten", 24),
    ("Job Number", "ten", 30),
    ("Reason for late payment", "ten", 60),
]

# v0.5 (WR-06): cột thêm khi có file Bán hàng — dư nợ Misa so với tổng job chưa thanh toán
COT_DOI_CHIEU = ("Đối chiếu job (FIN)", 46)

# Dòng 1 công ty · 2 tên báo cáo · 3 ngày chốt · 4 tên bảng · 5 tên cột
DONG_TIEU_DE = 5      # dòng chứa tên cột
DONG_DAU_DU_LIEU = 6  # dòng dữ liệu đầu tiên


# --- Tên Salesman: CEO chốt 08/10 dùng TÊN ĐẦY ĐỦ như trên SMS -------------
# Báo cáo cũ của kế toán trưởng dùng tên ngắn; tool đổi sang tên đầy đủ theo bảng này.
# Đối chiếu từ W40: cùng khách, bản KTT ghi tên ngắn, SMS ghi tên đầy đủ.
# Có nhân viên mới thì thêm một dòng ở đây.
TEN_SALES_DAY_DU = {
    "hiếu": "Trần Văn Hiếu",
    "hiếu 1": "Phạm Trần Hiếu",
    "hoài": "Phạm Thị Thương Hoài",
    "trân": "Võ Lê Huyền Trân",
    "hằng": "Lê Thị Hằng",
    "hanna": "Ngô Thúy Hằng",
    "thư": "Kim Thanh Thư",
    "tân": "Nguyễn Duy Tân",
}

# Lý do trễ hạn do tool tự viết khi Sales chưa ghi — đánh dấu để tuần sau không
# nhầm là lời giải trình của Sales.
NHAN_TU_DONG = "[TỰ ĐỘNG]"


def ten_sales(ten: str) -> str:
    """'Admin,Admin' → 'Admin'; 'Hiếu' → 'Trần Văn Hiếu'; tên đầy đủ giữ nguyên."""
    import unicodedata
    phan = []
    for x in str(ten or "").split(","):
        x = unicodedata.normalize("NFC", x).strip()
        if not x or x in ("+",):
            continue
        x = TEN_SALES_DAY_DU.get(x.lower(), x)
        if x not in phan:
            phan.append(x)
    return ", ".join(phan)
