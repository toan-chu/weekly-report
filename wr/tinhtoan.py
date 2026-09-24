# -*- coding: utf-8 -*-
"""Luật điền sheet Receivable. Bản đặc tả: handoff/docs/SPEC-logic-bao-cao-tuan.md

Tóm tắt:
  - Danh sách khách và cột Total  <- file Tổng hợp (đây là số tiền đúng)
  - Sáu nhóm tuổi                 <- file Tuổi nợ nếu tin được, nếu không thì sổ ghi
  - Các cột do người điền         <- mang từ trang trước của sổ ghi sang
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import soghi as sg
from .misa import DongTuoiNo, TongHop, TuoiNo

LECH_CHO_PHEP = 2.0  # đồng


@dataclass
class DongBaoCao:
    ma: str
    ten: str
    ten_chuan: str
    credit_term: str
    total: float
    nhom: List[float]
    salesman: str = ""
    ghi_chu: str = ""
    ly_do: str = ""
    nguon: str = ""                 # misa | so-ghi | khoi-dong-lanh | ngoai-le
    cho_xem_lai: List[str] = field(default_factory=list)      # vấn đề về SỐ
    thieu_thong_tin: List[str] = field(default_factory=list)  # thiếu cột người điền
    lo: List[sg.Lo] = field(default_factory=list)

    @property
    def nhom_gia_nhat(self) -> int:
        for i in range(5, -1, -1):
            if self.nhom[i] > 0.5:
                return i
        return 0


def _tin_duoc_file_tuoi_no(a: Optional[DongTuoiNo], total: float) -> bool:
    """File Tuổi nợ chỉ dùng được khi khớp số tiền với file Tổng hợp VÀ mọi hoá
    đơn đều có ghi hạn thanh toán."""
    return bool(a) and abs(a.tong - total) < LECH_CHO_PHEP and a.khong_han < 1


def tinh_mot_khach(
    ten_chuan: str,
    ma: str,
    ten: str,
    total: float,
    a: Optional[DongTuoiNo],
    truoc: Optional[sg.HoSoKhach],
    ngay_chot: dt.date,
    ngoai_le: Dict[str, dict],
    dieu_chinh: Optional[Dict[str, dict]] = None,
) -> DongBaoCao:
    dieu_chinh = dieu_chinh or {}
    term = (dieu_chinh.get(ten_chuan, {}).get("credit_term")
            or (truoc.credit_term if truoc else "") or "")
    d = DongBaoCao(
        ma=ma, ten=ten, ten_chuan=ten_chuan, credit_term=term, total=round(total, 2),
        nhom=[0.0] * 6,
        salesman=(truoc.salesman if truoc else ""),
        ghi_chu=(truoc.ghi_chu if truoc else ""),
        ly_do=(truoc.ly_do if truoc else ""),
    )
    term_ngay = sg.so_ngay_term(term)

    ngoai = ngoai_le.get(ten_chuan)
    if ngoai and "nhom" in ngoai:
        i = sg.TEN_NHOM.index(ngoai["nhom"])
        d.nhom = [0.0] * 6
        d.nhom[i] = round(total, 2)
        d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
        d.nguon = "ngoai-le"
        d.cho_xem_lai.append(f"Ngoại lệ do kế toán khai: xếp cả {total:,.0f} vào nhóm {ngoai['nhom']}"
                             + (f" — {ngoai.get('ly_do')}" if ngoai.get("ly_do") else ""))
        return d

    if _tin_duoc_file_tuoi_no(a, total):
        d.nhom = [round(a.truoc_han, 2)] + [round(x, 2) for x in a.qua_han]
        d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
        d.nguon = "misa"
    else:
        if truoc and truoc.lo:
            lo = sg.dieu_chinh_lo(truoc.lo, total, ngay_chot, term_ngay)
            d.nguon = "so-ghi"
        else:
            lo = [sg.Lo(round(total, 2), (ngay_chot + dt.timedelta(days=term_ngay)).isoformat(), True)]
            d.nguon = "khoi-dong-lanh"
            d.cho_xem_lai.append(
                "Khách chưa có trong sổ ghi: tạm xếp toàn bộ vào Current, tuổi nợ là ước tính"
            )
        d.lo = lo
        d.nhom = sg.nhom_tu_lo(lo, ngay_chot)
        if a is None:
            d.cho_xem_lai.append("Không có trong file Tuổi nợ, chia theo sổ ghi")
        elif abs(a.tong - total) >= LECH_CHO_PHEP:
            d.cho_xem_lai.append(
                f"Hai file Misa lệch nhau: Tổng hợp {total:,.0f} — Tuổi nợ {a.tong:,.0f}. "
                "Chia theo sổ ghi, tiền lấy theo Tổng hợp"
            )
        elif a.khong_han >= 1:
            d.cho_xem_lai.append(
                f"Hoá đơn không ghi hạn thanh toán: {a.khong_han:,.0f}. Tuổi nợ tính theo sổ ghi"
            )

    if not term:
        d.thieu_thong_tin.append("Chưa có Credit Term, tạm tính 30 ngày")
    if not d.salesman:
        d.thieu_thong_tin.append("Chưa có Salesman")
    lech = round(d.total - sum(d.nhom), 2)
    if abs(lech) >= 1:
        d.nhom[0] = round(d.nhom[0] + lech, 2)   # tổng các nhóm luôn phải bằng Total
        d.cho_xem_lai.append(f"Tổng các nhóm lệch Total {lech:,.0f}, đã dồn phần lệch vào Current")
    return d


def tinh_bang(
    th: TongHop,
    tn: TuoiNo,
    so: sg.SoGhi,
    ngay_chot: Optional[dt.date] = None,
) -> List[DongBaoCao]:
    ngay_chot = ngay_chot or tn.den_ngay
    truoc = so.trang_truoc(ngay_chot) or {}
    ra: List[DongBaoCao] = []
    da_bo: List[tuple] = []
    for k in th.khach_con_no():
        if k.ten_chuan in so.bo_qua:
            da_bo.append((k.ten, k.du_no, so.bo_qua[k.ten_chuan].get("ly_do", "")))
            continue
        ra.append(
            tinh_mot_khach(
                ten_chuan=k.ten_chuan, ma=k.ma, ten=k.ten, total=k.du_no,
                a=tn.dong.get(k.ten_chuan), truoc=truoc.get(k.ten_chuan),
                ngay_chot=ngay_chot, ngoai_le=so.ngoai_le, dieu_chinh=so.dieu_chinh,
            )
        )
    # nợ lâu nhất lên trên; cùng nhóm thì số tiền lớn lên trên
    ra.sort(key=lambda d: (-d.nhom_gia_nhat, -d.total))
    tinh_bang.da_bo = da_bo          # để báo cáo liệt kê ở sheet Cần xem lại
    return ra


def thanh_ho_so(d: DongBaoCao) -> sg.HoSoKhach:
    return sg.HoSoKhach(
        ma=d.ma, ten=d.ten, credit_term=d.credit_term, salesman=d.salesman,
        ghi_chu=d.ghi_chu, ly_do=d.ly_do, tong=d.total, nhom=d.nhom, lo=d.lo, nguon=d.nguon,
    )
