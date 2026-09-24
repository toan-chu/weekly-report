# -*- coding: utf-8 -*-
"""Màu và phông của thương hiệu Trustana, dùng chung cho mọi file Excel tool sinh ra."""

TIM = "4D148C"      # cột số liệu công nợ
CAM = "FF6200"      # cột định danh khách hàng
TRANG = "FFFFFF"
XAM_NHAT = "F2F2F2"
VANG_CANH_BAO = "FFF2CC"
DO_CANH_BAO = "FCE4E4"

PHONG_TIEU_DE = "Roboto"   # heading
PHONG_NOI_DUNG = "Calibri" # body

# Cột định danh (nền cam) và cột tiền (nền tím) của sheet Receivable
COT = [
    ("STT", "ten", 6),
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
    ("Ghi chú", "ten", 34),
    ("Lý do chưa thu hồi được công nợ", "ten", 34),
]

DONG_TIEU_DE = 7      # dòng chứa tên cột
DONG_DAU_DU_LIEU = 8  # dòng dữ liệu đầu tiên
