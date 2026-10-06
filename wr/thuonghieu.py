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
TEN_BAO_CAO = "Debtors and Creditors Ageing Reports"
TEN_BANG = "Accounts receivable"
SHEET_BANG = "Receivable"
SHEET_REVIEW = "Review"

PHONG_TIEU_DE = "Roboto"   # heading
PHONG_NOI_DUNG = "Calibri" # body

# Cột định danh (nền cam) và cột tiền (nền tím) của sheet Receivable
COT = [
    ("No.", "ten", 6),
    ("Code", "ten", 16),
    ("Customers' Name", "ten", 46),
    ("Credit Term", "ten", 14),
    ("Total", "tien", 16),
    ("Current", "tien", 15),
    ("1 - 30", "tien", 15),
    ("31 - 60", "tien", 15),
    ("61 - 90", "tien", 15),
    ("91 - 120", "tien", 15),
    ("120+", "tien", 15),
    ("Salesman", "ten", 12),
    ("Notes", "ten", 34),
    ("Reason for late payment", "ten", 34),
]

# Dòng 1 công ty · 2 tên báo cáo · 3 ngày chốt · 4 tên bảng · 5 tên cột
DONG_TIEU_DE = 5      # dòng chứa tên cột
DONG_DAU_DU_LIEU = 6  # dòng dữ liệu đầu tiên
