# -*- coding: utf-8 -*-
"""Ghi file báo cáo tuần theo mẫu chuẩn, và đọc lại báo cáo cũ để nạp sổ ghi."""
from __future__ import annotations

import datetime as dt
import re
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


def ghep_ghi_chu(job: str, ghi_chu: str) -> str:
    """Cột Ghi chú của mẫu chuẩn chứa cả số job: 'job: A, B | ghi chú'."""
    phan = []
    if job:
        phan.append(f"job: {job}")
    if ghi_chu:
        phan.append(ghi_chu)
    return " | ".join(phan)


def tach_ghi_chu(o: str):
    """Ngược với ghep_ghi_chu. Trả (job, ghi chú)."""
    o = str(o or "").strip()
    if not o.lower().startswith("job:"):
        return "", o
    dau, _, sau = o.partition(" | ")
    return dau[4:].strip(), sau.strip()


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
    ws = wb[th.SHEET_BANG] if th.SHEET_BANG in wb.sheetnames else wb.worksheets[0]
    ws.title = th.SHEET_BANG
    # chữ tiêu đề lấy từ thuonghieu.py mỗi lần chạy, sửa ở đó là đủ
    ws["A1"], ws["A2"], ws["A4"] = th.CONG_TY, th.TEN_BAO_CAO, th.TEN_BANG
    for i, (ten_cot, _, _) in enumerate(th.COT, start=1):
        ws.cell(row=th.DONG_TIEU_DE, column=i).value = ten_cot
    ws["A3"] = _ngay_en(ngay_chot)

    vien = Side(style="thin", color="D9D9D9")
    kieu_vien = Border(left=vien, right=vien, top=vien, bottom=vien)
    hang = th.DONG_DAU_DU_LIEU
    for stt, d in enumerate(dong, start=1):
        gia_tri = [stt, d.ma, d.ten, d.credit_term, d.total] + \
                  [v if abs(v) > 0.4 else None for v in d.nhom] + \
                  [d.salesman, ghep_ghi_chu(d.job, d.ghi_chu), d.ly_do]
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
            mau = (th.DO_CANH_BAO if d.nguon in {"khoi-dong-lanh", "ngoai-le", "uoc-tinh"}
                   else th.VANG_CANH_BAO)
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
    ws = wb.create_sheet(th.SHEET_REVIEW)
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
    if doi_chieu.get("che_do") == "sms":
        hang += 1
        tong_nguon: Dict[str, float] = {}
        for d in dong:
            for k, v in d.phan_nguon.items():
                tong_nguon[k] = tong_nguon.get(k, 0) + v
        _dong_doi_chieu("Tuổi nợ tính từ đâu", tong_bang, "Bằng các dòng dưới cộng lại.")
        for k, nhan, giai_thich, mau in (
            ("sms", "  · ngày job trên SMS", "ETD/ETA của từng job + credit term.", "000000"),
            ("so-ghi", "  · bản final/sổ ghi tuần trước", "Nợ cũ không có trong SMS: giữ nhóm tuổi "
             "kế toán trưởng đã xếp, cộng thêm số ngày đã trôi qua.", "000000"),
            ("kho-doi", "  · khó đòi (FIN khai)", "Luôn ở nhóm 120+.", "000000"),
            ("ngoai-le", "  · ngoại lệ kế toán khai", "", "000000"),
            ("uoc-tinh", "  · chưa rõ tuổi, đang ước tính", "Không có trong SMS lẫn bản final. "
             "Dòng tô đỏ ở sheet Receivable — kế toán xếp lại nhóm trên bản final.", "C00000"),
        ):
            if tong_nguon.get(k, 0) >= 1:
                _dong_doi_chieu(nhan, tong_nguon[k], giai_thich, mau)
    else:
        hang += 1
        _ghi_no_trong_han(ws, dong, _dong_doi_chieu)
        hang = ws.max_row + 1
    hang += 1
    _bang_khach_xem_lai(ws, dong, doi_chieu, hang)


def _ghi_no_trong_han(ws, dong, _dong_doi_chieu) -> None:
    # Nợ trong hạn: vì sao báo cáo khác cột "Nợ trước hạn" của file Tuổi nợ
    cur_tu_misa = sum(d.nhom[0] for d in dong if d.nguon == "misa")
    cur_khac = sum(d.nhom[0] for d in dong if d.nguon != "misa")
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


def _bang_khach_xem_lai(ws, dong, doi_chieu, hang) -> None:
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

    for khoa, tieu, cot_gt, ghi in (
        ("sms_cat", "SMS cần gạch paid — SMS ghi nợ nhiều hơn Misa, tool đã bỏ phần này",
         ("Khách hàng", "Số tiền bỏ", "Job", "Ngày ETD/ETA"),
         "FIN kiểm tra và ấn paid trên SMS cho các job này."),
        ("sms_khong_misa", "SMS còn nợ nhưng Misa không còn — không đưa vào báo cáo",
         ("Khách trên SMS", "Số tiền SMS", "Job", "Vì sao"),
         "Nhiều khả năng đã thu tiền mà SMS chưa gạch paid."),
        ("sms_chua_ghep", "Khách SMS chưa có trong bảng ghép — số này CHƯA vào báo cáo",
         ("Mã | tên trên SMS", "Số tiền SMS", "Số dòng", ""),
         "Bổ sung vào file ghép khách (nút 2_Kiem_tra_ghep_khach sẽ liệt kê sẵn)."),
    ):
        ds = doi_chieu.get(khoa) or []
        if not ds:
            continue
        hang += 1
        o = ws.cell(row=hang, column=1, value=f"{tieu} — {ghi}")
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.CAM)
        hang += 1
        for i, t in enumerate(cot_gt, start=1):
            ws.cell(row=hang, column=i, value=t).font = Font(name=th.PHONG_NOI_DUNG, size=9, bold=True)
        hang += 1
        for dong_ds in ds:
            if khoa == "sms_cat":
                vals = (dong_ds[0], dong_ds[2], dong_ds[1],
                        dong_ds[3].strftime("%d/%m/%Y") if dong_ds[3] else "")
            elif khoa == "sms_khong_misa":
                vals = (f"{dong_ds[0]} | {dong_ds[1]}", dong_ds[3], dong_ds[2], dong_ds[4])
            else:
                vals = (f"{dong_ds[0]} | {dong_ds[1]}", dong_ds[2], dong_ds[3], "")
            for i, v in enumerate(vals, start=1):
                c = ws.cell(row=hang, column=i, value=v)
                c.font = Font(name=th.PHONG_NOI_DUNG, size=10)
                if i == 2:
                    c.number_format = DINH_DANG_TIEN
            hang += 1

    hang += 1
    thieu_term = sum(1 for d in dong if any("Credit Term" in x for x in d.thieu_thong_tin))
    thieu_sale = sum(1 for d in dong if any("Salesman" in x for x in d.thieu_thong_tin))
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
    "not-yet-due": 0, "not yet due": 0, "over 120": 5,
}


def _tim_bang(wb):
    """Tìm sheet và dòng tiêu đề của bảng công nợ: sheet Receivable của mẫu chuẩn,
    hoặc sheet bất kỳ (VD 'data' trong bản final kế toán trưởng)."""
    thu_tu = (["Receivable"] if "Receivable" in wb.sheetnames else []) + \
             [n for n in wb.sheetnames if n != "Receivable"]
    for ten in thu_tu:
        ws = wb[ten]
        for r, hang in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            gia_tri = [str(v or "").strip().lower() for v in hang]
            if "code" in gia_tri and any(v in ("total", "total receivables") for v in gia_tri):
                return ws, r
    return None, None


def ngay_chot_trong_file(duong_dan: Path) -> Optional[dt.date]:
    """Ngày chốt của một báo cáo tuần: 'Reporting period dd/mm/yyyy - dd/mm/yyyy'
    (lấy ngày sau) trong bản final, hoặc 'At 26 Sep 2026' ở ô A3 của mẫu chuẩn."""
    wb = openpyxl.load_workbook(duong_dan, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for hang in ws.iter_rows(min_row=1, max_row=12, values_only=True):
                chu = " ".join(str(v) for v in hang if v is not None)
                if "period" in chu.lower() or "kỳ" in chu.lower():
                    ngay = re.findall(r"(\d{1,2})/(\d{1,2})/(\d{4})", chu)
                    if ngay:
                        d, m, y = ngay[-1]
                        return dt.date(int(y), int(m), int(d))
                m = re.search(r"\bAt (\d{1,2}) (\w{3}) (\d{4})", chu)
                if m and m.group(2) in THANG_EN:
                    return dt.date(int(m.group(3)), THANG_EN.index(m.group(2)) + 1, int(m.group(1)))
    finally:
        wb.close()
    return None


def doc_bao_cao(duong_dan: Path) -> Dict[str, sg.HoSoKhach]:
    """Đọc sheet Receivable của một báo cáo tuần (kể cả bản kế toán làm tay,
    kể cả bản cũ chỉ có 4 cột tuổi nợ) để nạp vào sổ ghi."""
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    ws, dong_tieu_de = _tim_bang(wb)
    if ws is None:
        raise ValueError(f"{Path(duong_dan).name}: không tìm thấy bảng công nợ (dòng tiêu đề có Code và Total)")

    cot: Dict[str, int] = {}
    nhom_cot: Dict[int, int] = {}
    for c in ws[dong_tieu_de]:
        ten = str(c.value or "").strip().lower()
        if not ten:
            continue
        ten_nhom = re.sub(r"\s+", " ", ten).replace(" days", "").strip()
        if ten_nhom in ANH_XA_NHOM:
            nhom_cot[c.column] = ANH_XA_NHOM[ten_nhom]
        elif "customer" in ten and "ten" not in cot:
            cot["ten"] = c.column
        elif ten == "code":
            cot["ma"] = c.column
        elif "credit term" in ten:
            cot["term"] = c.column
        elif ten in ("total", "total receivables"):
            cot["total"] = c.column
        elif "salesman" in ten and "salesman" not in cot:
            cot["salesman"] = c.column
        elif ten in ("job", "job no", "số job"):
            cot["job"] = c.column
        elif "ghi chú" in ten or ten in ("comment", "notes", "note"):
            cot["ghi_chu"] = c.column
        elif "lý do" in ten or "reason" in ten:
            cot["ly_do"] = c.column

    ra: Dict[str, sg.HoSoKhach] = {}
    for r in range(dong_tieu_de + 1, ws.max_row + 1):
        ten = ws.cell(row=r, column=cot["ten"]).value if "ten" in cot else None
        ma = ws.cell(row=r, column=cot.get("ma", 2)).value
        if not ten or str(ma or "").strip().upper().startswith(("TỔNG", "TONG")):
            continue
        if chuan_hoa_ten(ten) in ("TOTAL", "TỔNG", "TỔNG CỘNG", "GRAND TOTAL"):
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
        job_gc, ghi_chu = tach_ghi_chu(lay("ghi_chu"))
        job = re.sub(r"(?i)^\s*job\s*:\s*", "", lay("job")) or job_gc
        ra[tc] = sg.HoSoKhach(
            ma=str(ma or "").strip(), ten=str(ten).strip(), credit_term=lay("term"),
            salesman=lay("salesman"), ghi_chu=ghi_chu, ly_do=lay("ly_do"),
            tong=round(float(total), 2), nhom=[round(v, 2) for v in nhom], nguon="ban-tay",
            job=job,
        )
    return ra


def nap_vao_so_ghi(so: sg.SoGhi, duong_dan: Path, ngay_chot: dt.date) -> int:
    khach = doc_bao_cao(duong_dan)
    for hs in khach.values():
        hs.lo = sg.lo_tu_nhom(hs.nhom, ngay_chot, uoc_tinh=True)
    so.ghi_trang(ngay_chot, khach)
    return len(khach)
