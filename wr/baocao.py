# -*- coding: utf-8 -*-
"""Ghi file báo cáo tuần theo mẫu chuẩn, và đọc lại báo cáo cũ để nạp sổ ghi."""
from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import soghi as sg
from . import thuonghieu as th
from .misa import chuan_hoa_ten
from .tinhtoan import DongBaoCao

DINH_DANG_TIEN = "#,##0;[Red]-#,##0"
THANG_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def ten_file_bao_cao(tu_ngay: dt.date, den_ngay: dt.date) -> str:
    tuan = den_ngay.isocalendar()[1]
    return (f"{den_ngay.year}_W{tuan:02d}_Bang_cong_no_tuan_"
            f"{tu_ngay:%d.%m}-{den_ngay:%d.%m.%Y}.xlsx")


def _ngay_en(d: dt.date) -> str:
    return f"At {d.day:02d} {THANG_EN[d.month - 1]} {d.year}"


def ghi_bao_cao(
    dong: List[DongBaoCao],
    ngay_chot: dt.date,
    mau_chuan: Path,
    dich: Path,
    tong_misa: Optional[float] = None,
    doi_chieu: Optional[dict] = None,
) -> Path:
    dich.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(mau_chuan, dich)
    wb = openpyxl.load_workbook(dich)
    ws = wb["Receivable"]
    ws["A3"] = _ngay_en(ngay_chot)

    vien = Side(style="thin", color="D9D9D9")
    kieu_vien = Border(left=vien, right=vien, top=vien, bottom=vien)
    hang = th.DONG_DAU_DU_LIEU
    for stt, d in enumerate(dong, start=1):
        gia_tri = [stt, d.ma, d.ten, d.credit_term, d.total] + \
                  [v if abs(v) > 0.4 else None for v in d.nhom] + \
                  [d.salesman, d.ghi_chu, d.ly_do]
        for cot, v in enumerate(gia_tri, start=1):
            o = ws.cell(row=hang, column=cot, value=v)
            o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o.border = kieu_vien
            if 5 <= cot <= 11:
                o.number_format = DINH_DANG_TIEN
                o.alignment = Alignment(horizontal="right")
            elif cot == 1:
                o.alignment = Alignment(horizontal="center")
            else:
                o.alignment = Alignment(vertical="top", wrap_text=cot >= 13)
        if d.cho_xem_lai:
            mau = th.DO_CANH_BAO if d.nguon in {"khoi-dong-lanh", "ngoai-le"} else th.VANG_CANH_BAO
            for cot in range(1, len(th.COT) + 1):
                ws.cell(row=hang, column=cot).fill = PatternFill("solid", fgColor=mau)
        hang += 1

    hang_tong = hang
    o = ws.cell(row=hang_tong, column=2, value="TỔNG")
    o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    for cot in range(1, len(th.COT) + 1):
        c = ws.cell(row=hang_tong, column=cot)
        c.fill = PatternFill("solid", fgColor=th.TIM)
        c.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        if 5 <= cot <= 11:
            chu = get_column_letter(cot)
            c.value = f"=SUM({chu}{th.DONG_DAU_DU_LIEU}:{chu}{hang_tong - 1})"
            c.number_format = DINH_DANG_TIEN
            c.alignment = Alignment(horizontal="right")

    ws.auto_filter.ref = f"A{th.DONG_TIEU_DE}:{get_column_letter(len(th.COT))}{hang_tong - 1}"
    _ghi_sheet_xem_lai(wb, dong, ngay_chot, tong_misa, doi_chieu or {})
    wb.save(dich)
    return dich


def _ghi_sheet_xem_lai(wb, dong: List[DongBaoCao], ngay_chot: dt.date, tong_misa,
                      doi_chieu: dict) -> None:
    ws = wb.create_sheet("Cần xem lại")
    ws["A1"] = f"Đối chiếu và những dòng tool thấy không chắc — chốt số {ngay_chot:%d/%m/%Y}"
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)

    tong_bang = sum(d.total for d in dong)
    tong_tuoi_no = doi_chieu.get("tong_tuoi_no")
    lech_khach = doi_chieu.get("khach_lech") or []

    hang = 3
    def _dong_doi_chieu(nhan, gia_tri, ghi_chu="", mau="000000"):
        nonlocal hang
        o = ws.cell(row=hang, column=1, value=nhan)
        o.font = Font(name=th.PHONG_NOI_DUNG, size=10, bold=True, color=mau)
        c = ws.cell(row=hang, column=2, value=gia_tri)
        c.number_format = DINH_DANG_TIEN
        c.font = Font(name=th.PHONG_NOI_DUNG, size=10, bold=True, color=mau)
        o2 = ws.cell(row=hang, column=4, value=ghi_chu)
        o2.font = Font(name=th.PHONG_NOI_DUNG, size=10, color=mau)
        o2.alignment = Alignment(wrap_text=True, vertical="top")
        hang += 1

    nguon = doi_chieu.get("nguon") or []
    for nhan, mo_ta in nguon:
        o = ws.cell(row=hang, column=1, value=nhan)
        o.font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True)
        o2 = ws.cell(row=hang, column=4, value=mo_ta)
        o2.font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True)
        hang += 1
    if nguon:
        hang += 1

    _dong_doi_chieu("Tổng bảng báo cáo này", tong_bang,
                    "Phải bằng dòng ngay dưới. Lệch là tool có lỗi.")
    if tong_misa is not None:
        khop = abs(tong_bang - tong_misa) < 2
        _dong_doi_chieu("Tổng file Tổng hợp công nợ (dư Nợ)", tong_misa,
                        "KHỚP" if khop else "KHÔNG KHỚP — báo lại người làm tool",
                        "000000" if khop else "C00000")
    if tong_tuoi_no is not None:
        chenh = tong_tuoi_no - (tong_misa if tong_misa is not None else tong_bang)
        _dong_doi_chieu("Tổng file Phân tích tuổi nợ", tong_tuoi_no,
                        "Con số này KHÔNG dùng để đối chiếu tổng báo cáo. "
                        "Nó thường cao hơn vì còn hoá đơn đã thu nhưng chưa gắn "
                        "chứng từ thanh toán.")
        if abs(chenh) >= 2:
            _dong_doi_chieu("Chênh lệch giữa hai file Misa", chenh,
                            "Do các khách liệt kê ngay dưới. Tiền luôn lấy theo "
                            "file Tổng hợp.", "B26B00")
            hang += 1
            o = ws.cell(row=hang, column=1, value="Khách làm lệch hai file")
            o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.CAM)
            hang += 1
            for ten, v_th, v_tn in lech_khach:
                ws.cell(row=hang, column=1, value=ten).font = Font(name=th.PHONG_NOI_DUNG, size=10)
                for cot, v in ((2, v_th), (3, v_tn)):
                    c = ws.cell(row=hang, column=cot, value=v)
                    c.number_format = DINH_DANG_TIEN
                    c.font = Font(name=th.PHONG_NOI_DUNG, size=10)
                ws.cell(row=hang, column=4,
                        value="Tổng hợp | Tuổi nợ").font = Font(name=th.PHONG_NOI_DUNG, size=9,
                                                                italic=True)
                hang += 1
    # Nợ trong hạn: vì sao báo cáo khác cột "Nợ trước hạn" của file Tuổi nợ
    cur_tu_misa = sum(d.nhom[0] for d in dong if d.nguon == "misa")
    cur_khac = sum(d.nhom[0] for d in dong if d.nguon != "misa")
    hang += 1
    _dong_doi_chieu("Nợ trong hạn trên báo cáo", cur_tu_misa + cur_khac,
                    "Bằng hai dòng dưới cộng lại.")
    _dong_doi_chieu("  · lấy thẳng từ file Tuổi nợ", cur_tu_misa,
                    "Khớp với cột Nợ trước hạn của file Tuổi nợ, phần các khách "
                    "mà hai file Misa khớp nhau.")
    if cur_khac >= 1:
        ten_khac = ", ".join(d.ten for d in dong if d.nguon != "misa" and d.nhom[0] >= 1)
        _dong_doi_chieu("  · chưa xác định được tuổi, tạm để ở đây", cur_khac,
                        f"Hoá đơn không ghi hạn thanh toán, hoặc hai file Misa lệch "
                        f"nhau, nên chưa tính được tuổi. Gồm: {ten_khac}. "
                        "Các dòng này được tô màu ở sheet Receivable.", "B26B00")
    hang += 1

    tieu_de = ["Khách hàng", "Total", "Nguồn số", "Vì sao cần xem lại"]
    for i, t in enumerate(tieu_de, start=1):
        o = ws.cell(row=hang, column=i, value=t)
        o.fill = PatternFill("solid", fgColor=th.CAM if i != 2 else th.TIM)
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    hang += 1
    thieu_term = sum(1 for d in dong if any("Credit Term" in x for x in d.thieu_thong_tin))
    thieu_sale = sum(1 for d in dong if any("Salesman" in x for x in d.thieu_thong_tin))
    # vấn đề về số liệu xếp trước, rồi mới tới thiếu thông tin do người điền
    for d in sorted(dong, key=lambda x: (not x.cho_xem_lai, -x.total)):
        ly_do = d.cho_xem_lai + d.thieu_thong_tin
        if not ly_do:
            continue
        ws.cell(row=hang, column=1, value=d.ten).font = Font(name=th.PHONG_NOI_DUNG, size=10)
        c = ws.cell(row=hang, column=2, value=d.total)
        c.number_format = DINH_DANG_TIEN
        c.font = Font(name=th.PHONG_NOI_DUNG, size=10)
        ws.cell(row=hang, column=3, value=d.nguon).font = Font(name=th.PHONG_NOI_DUNG, size=10)
        o = ws.cell(row=hang, column=4, value=" | ".join(ly_do))
        o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
        o.alignment = Alignment(wrap_text=True, vertical="top")
        hang += 1
    da_bo = doi_chieu.get("da_bo") or []
    if da_bo:
        hang += 1
        o = ws.cell(row=hang, column=1, value="Khách bị loại khỏi bảng theo khai báo của kế toán")
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.CAM)
        hang += 1
        for ten, tien, ly_do in da_bo:
            ws.cell(row=hang, column=1, value=ten).font = Font(name=th.PHONG_NOI_DUNG, size=10)
            c = ws.cell(row=hang, column=2, value=tien)
            c.number_format = DINH_DANG_TIEN
            c.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o = ws.cell(row=hang, column=4, value=ly_do)
            o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o.alignment = Alignment(wrap_text=True, vertical="top")
            hang += 1

    hang += 1
    o = ws.cell(row=hang, column=1,
                value=f"Thiếu thông tin do người điền: {thieu_term} khách chưa có Credit Term, "
                      f"{thieu_sale} khách chưa có Salesman. Đây không phải lỗi số liệu.")
    o.font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True)
    for cot, rong in zip("ABCD", (46, 16, 16, 80)):
        ws.column_dimensions[cot].width = rong


# --- nạp trạng thái ban đầu từ một báo cáo đã có -------------------------

ANH_XA_NHOM = {
    "current": 0, "1 - 30": 1, "1-30": 1, "31 - 60": 2, "31-60": 2,
    "61 - 90": 3, "61-90": 3, "60+": 3, "91 - 120": 4, "91 -120": 4, "91-120": 4,
    "120+": 5, "trên 120": 5,
}


def doc_bao_cao(duong_dan: Path) -> Dict[str, sg.HoSoKhach]:
    """Đọc sheet Receivable của một báo cáo tuần (kể cả bản kế toán làm tay,
    kể cả bản cũ chỉ có 4 cột tuổi nợ) để nạp vào sổ ghi."""
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    ws = wb["Receivable"]
    dong_tieu_de = None
    for r in range(1, 15):
        gia_tri = [str(c.value or "").strip().lower() for c in ws[r]]
        if "code" in gia_tri and "total" in gia_tri:
            dong_tieu_de = r
            break
    if dong_tieu_de is None:
        raise ValueError(f"{Path(duong_dan).name}: không tìm thấy dòng tiêu đề của sheet Receivable")

    cot: Dict[str, int] = {}
    nhom_cot: Dict[int, int] = {}
    for c in ws[dong_tieu_de]:
        ten = str(c.value or "").strip().lower()
        if not ten:
            continue
        if ten in ANH_XA_NHOM:
            nhom_cot[c.column] = ANH_XA_NHOM[ten]
        elif "customer" in ten:
            cot["ten"] = c.column
        elif ten == "code":
            cot["ma"] = c.column
        elif "credit term" in ten:
            cot["term"] = c.column
        elif ten == "total":
            cot["total"] = c.column
        elif "salesman" in ten:
            cot["salesman"] = c.column
        elif "ghi chú" in ten:
            cot["ghi_chu"] = c.column
        elif "lý do" in ten:
            cot["ly_do"] = c.column

    ra: Dict[str, sg.HoSoKhach] = {}
    for r in range(dong_tieu_de + 1, ws.max_row + 1):
        ten = ws.cell(row=r, column=cot["ten"]).value if "ten" in cot else None
        ma = ws.cell(row=r, column=cot.get("ma", 2)).value
        if not ten or str(ma or "").strip().upper().startswith(("TỔNG", "TONG")):
            continue
        total = ws.cell(row=r, column=cot["total"]).value
        if not isinstance(total, (int, float)):
            continue
        nhom = [0.0] * 6
        for c, i in nhom_cot.items():
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                nhom[i] += float(v)
        lay = lambda k: str(ws.cell(row=r, column=cot[k]).value or "").strip() if k in cot else ""
        tc = chuan_hoa_ten(ten)
        ra[tc] = sg.HoSoKhach(
            ma=str(ma or "").strip(), ten=str(ten).strip(), credit_term=lay("term"),
            salesman=lay("salesman"), ghi_chu=lay("ghi_chu"), ly_do=lay("ly_do"),
            tong=round(float(total), 2), nhom=[round(v, 2) for v in nhom], nguon="ban-tay",
        )
    return ra


def nap_vao_so_ghi(so: sg.SoGhi, duong_dan: Path, ngay_chot: dt.date) -> int:
    khach = doc_bao_cao(duong_dan)
    for hs in khach.values():
        hs.lo = sg.lo_tu_nhom(hs.nhom, ngay_chot, uoc_tinh=True)
    so.ghi_trang(ngay_chot, khach)
    return len(khach)
