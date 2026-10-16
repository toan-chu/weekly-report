# -*- coding: utf-8 -*-
"""Client master list (credit term + hợp đồng) — WR-08.

File sống trên OneDrive, Sales và FIN cùng điền. Tool chỉ ĐỌC cột "Mã kh" và "Thời hạn"
(ghep.doc_credit_term tìm theo tên tiêu đề trong 6 dòng đầu), nên tiêu đề dòng 6 phải giữ đúng chữ.
Module này dựng file theo màu/phông của báo cáo: dùng cho file mẫu và cho lần tạo đầu tiên
(kiemtra.tao_file_credit_term)."""
from pathlib import Path
from typing import Dict, List, Sequence

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import thuonghieu as th
from . import ghep

SHEET_MASTER = "Client Master"
SHEET_HUONG_DAN = "Hướng dẫn"
DONG_NHOM = 5          # dải tên nhóm cột
DONG_TIEU_DE = 6       # tiêu đề cột (tool đọc dòng này)
DONG_DAU = 7
DONG_CUOI = 400        # công thức + danh sách thả xuống kéo sẵn tới đây cho dòng Sales thêm mới
NGAY = "dd/mm/yyyy"

# (khoá, tiêu đề, nhóm, độ rộng, loại) — loại: ten | ngay | so | cong_thuc
COT = [
    ("stt",      "STT",                   "Khách hàng", 6,  "cong_thuc"),
    ("ma",       "Mã kh",                 "Khách hàng", 16, "ten"),
    ("ten",      "Khách hàng",            "Khách hàng", 44, "ten"),
    ("sales",    "Sales phụ trách",       "Khách hàng", 24, "ten"),
    ("pic",      "Người liên hệ (PIC)",   "Khách hàng", 20, "ten"),
    ("mst",      "Mã số thuế",            "Khách hàng", 14, "ten"),
    ("chu_tk",   "Chủ tài khoản",         "Tài khoản ngân hàng", 22, "ten"),
    ("so_tk",    "Số tài khoản",          "Tài khoản ngân hàng", 18, "ten"),
    ("ngan_hang","Ngân hàng",             "Tài khoản ngân hàng", 18, "ten"),
    ("chi_nhanh","Chi nhánh",             "Tài khoản ngân hàng", 16, "ten"),
    ("so_hd",    "Số hợp đồng",           "Hợp đồng", 16, "ten"),
    ("loai_hd",  "Loại hợp đồng",         "Hợp đồng", 18, "ten"),
    ("term",     "Thời hạn",              "Hợp đồng", 30, "ten"),
    ("hieu_luc", "Ngày hiệu lực",         "Hợp đồng", 13, "ngay"),
    ("het_han",  "Ngày hết hạn",          "Hợp đồng", 13, "ngay"),
    ("gia_han",  "Ngày gia hạn",          "Gia hạn", 13, "ngay"),
    ("het_moi",  "Hết hạn sau gia hạn",   "Gia hạn", 13, "cong_thuc"),
    ("han_dung", "Hạn đang áp dụng",      "Tình trạng hợp đồng", 13, "cong_thuc"),
    ("con_lai",  "Còn lại (ngày)",        "Tình trạng hợp đồng", 10, "cong_thuc"),
    ("trang_thai","Trạng thái",           "Tình trạng hợp đồng", 22, "cong_thuc"),
    ("job_dau",  "Ngày job đầu tiên",     "Ghi chú", 13, "ngay"),
    ("ghi_chu",  "Ghi chú",               "Ghi chú", 44, "ten"),
]
VI_TRI = {k: get_column_letter(i) for i, (k, *_r) in enumerate(COT, start=1)}
LOAI_HD = ["Hợp đồng nguyên tắc", "Hợp đồng từng lô", "Báo giá / email", "Chưa có hợp đồng"]
TRANG_THAI = ["Đã hết hạn", "Sắp hết hạn ≤30 ngày", "Sắp hết hạn ≤60 ngày", "Sắp hết hạn ≤90 ngày",
              "Còn hiệu lực", "Chưa có hạn HĐ"]

_VIEN = Side(style="thin", color="D9D9D9")
VIEN = Border(left=_VIEN, right=_VIEN, top=_VIEN, bottom=_VIEN)
_TIM_NHAT = "E9E1F3"   # nền dải nhóm của cột hợp đồng
_CAM_NHAT = "FFE8D9"   # nền dải nhóm của cột định danh


def _mau_cot(nhom: str) -> str:
    return th.CAM if nhom in ("Khách hàng", "Tài khoản ngân hàng") else th.TIM


def _cong_thuc(k: str, r: int) -> str:
    c = VI_TRI
    if k == "stt":
        return f'=IF({c["ten"]}{r}="","",ROW()-{DONG_TIEU_DE})'
    if k == "het_moi":
        return f'=IF({c["gia_han"]}{r}="","",EDATE({c["gia_han"]}{r},12))'
    if k == "han_dung":
        return f'=IF({c["het_moi"]}{r}<>"",{c["het_moi"]}{r},IF({c["het_han"]}{r}="","",{c["het_han"]}{r}))'
    if k == "con_lai":
        return f'=IF({c["han_dung"]}{r}="","",{c["han_dung"]}{r}-TODAY())'
    if k == "trang_thai":
        h, s = f'{c["han_dung"]}{r}', f'{c["con_lai"]}{r}'
        return (f'=IF({c["ten"]}{r}="","",IF({h}="","{TRANG_THAI[5]}",IF({s}<0,"{TRANG_THAI[0]}",'
                f'IF({s}<=30,"{TRANG_THAI[1]}",IF({s}<=60,"{TRANG_THAI[2]}",IF({s}<=90,"{TRANG_THAI[3]}",'
                f'"{TRANG_THAI[4]}"))))))')
    raise KeyError(k)


def _dau_trang(ws, ghi_chu: str) -> None:
    ws["A1"], ws["A2"], ws["A3"], ws["A4"] = (th.CONG_TY, "Client Master List", ghi_chu,
                                               "Credit term & hợp đồng khách hàng")
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)
    ws["A3"].font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True)
    ws["A4"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)


def _tieu_de(ws) -> None:
    dau = 1
    for i in range(1, len(COT) + 2):
        if i > len(COT) or COT[i - 1][2] != COT[dau - 1][2]:
            nhom = COT[dau - 1][2]
            o = ws.cell(row=DONG_NHOM, column=dau, value=nhom)
            if i - 1 > dau:
                ws.merge_cells(start_row=DONG_NHOM, start_column=dau, end_row=DONG_NHOM, end_column=i - 1)
            for j in range(dau, i):
                ws.cell(row=DONG_NHOM, column=j).fill = PatternFill(
                    "solid", fgColor=_CAM_NHAT if _mau_cot(nhom) == th.CAM else _TIM_NHAT)
            o.font = Font(name=th.PHONG_TIEU_DE, size=9, bold=True, color=_mau_cot(nhom))
            o.alignment = Alignment(horizontal="center", vertical="center")
            dau = i
    for i, (k, ten, nhom, rong, loai) in enumerate(COT, start=1):
        o = ws.cell(row=DONG_TIEU_DE, column=i, value=ten)
        o.fill = PatternFill("solid", fgColor=_mau_cot(nhom))
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        o.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="left" if loai == "ten" else "center")
        o.border = VIEN
        ws.column_dimensions[get_column_letter(i)].width = rong
    ws.row_dimensions[DONG_TIEU_DE].height = 32
    ghi = {"ma": "FIN điền: mã khách trên Misa. Tool ghép báo cáo theo cột này.",
           "term": "Tool đọc cột này mỗi tuần. Viết: 'thanh toán ngay', '30 days', "
                   "'14 ngày của tháng tiếp theo'… — xem sheet Hướng dẫn.",
           "sales": "Tên đầy đủ như trên SMS. Khách nhiều Sales: ghi cách nhau dấu phẩy.",
           "het_moi": "Tự tính: ngày gia hạn + 12 tháng.",
           "trang_thai": "Tự tính từ Hạn đang áp dụng (≤30 / ≤60 / ≤90 ngày)."}
    for k, t in ghi.items():
        cm = Comment(t, "Trustana")
        cm.width, cm.height = 260, 70
        ws[f"{VI_TRI[k]}{DONG_TIEU_DE}"].comment = cm


def _o_du_lieu(o, loai: str, co_gia_tri: bool) -> None:
    o.font = Font(name=th.PHONG_NOI_DUNG, size=10, color="595959" if loai == "cong_thuc" else None)
    o.border = VIEN
    if loai == "cong_thuc":
        o.fill = PatternFill("solid", fgColor=th.XAM_NHAT)
    if loai in ("ngay",) or o.column_letter in (VI_TRI["het_moi"], VI_TRI["han_dung"]):
        o.number_format = NGAY
        o.alignment = Alignment(horizontal="center", vertical="top")
    elif o.column_letter in (VI_TRI["stt"], VI_TRI["con_lai"]):
        o.number_format = "0"
        o.alignment = Alignment(horizontal="center", vertical="top")
    else:
        o.alignment = Alignment(vertical="top", wrap_text=False)


def ghi_master(dich: Path, dong: Sequence[Dict], sales: Sequence[str], ghi_chu: str = "") -> Path:
    """dong: [{khoá cột: giá trị}] — khoá như COT; cột công thức tự điền. sales: danh sách thả xuống."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_MASTER
    ws.sheet_view.showGridLines = False
    _dau_trang(ws, ghi_chu or "File chung trên OneDrive · Sales và FIN điền trực tiếp · "
                              "tool báo cáo tuần đọc cột Mã kh và Thời hạn")
    _tieu_de(ws)
    for r in range(DONG_DAU, DONG_CUOI + 1):
        d = dong[r - DONG_DAU] if r - DONG_DAU < len(dong) else {}
        for i, (k, _t, _n, _w, loai) in enumerate(COT, start=1):
            o = ws.cell(row=r, column=i)
            if loai == "cong_thuc":
                o.value = _cong_thuc(k, r)
            else:
                v = d.get(k)
                o.value = v if v not in ("", None) else None
            _o_du_lieu(o, loai, o.value is not None)

    c = VI_TRI
    # Thả xuống: gợi ý, không chặn (Sales vẫn gõ được kiểu khác)
    hd = SHEET_HUONG_DAN
    for cot, nguon in (("term", f"'{hd}'!$B$22:$B$29"), ("sales", f"'{hd}'!$B$33:$B$52"),
                       ("loai_hd", f"'{hd}'!$B$56:$B$59")):
        dv = DataValidation(type="list", formula1=nguon, allow_blank=True, showErrorMessage=False)
        dv.add(f"{c[cot]}{DONG_DAU}:{c[cot]}{DONG_CUOI}")
        ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2015,1,1)", allow_blank=True,
                        errorTitle="Ngày", error="Nhập ngày dạng dd/mm/yyyy")
    for k in ("hieu_luc", "het_han", "gia_han", "job_dau"):
        dv.add(f"{c[k]}{DONG_DAU}:{c[k]}{DONG_CUOI}")
    ws.add_data_validation(dv)

    do = PatternFill("solid", fgColor=th.DO_CANH_BAO)
    vang = PatternFill("solid", fgColor=th.VANG_CANH_BAO)
    tt = f"{c['trang_thai']}{DONG_DAU}"
    ws.conditional_formatting.add(f"{c['trang_thai']}{DONG_DAU}:{c['trang_thai']}{DONG_CUOI}",
        FormulaRule(formula=[f'{tt}="{TRANG_THAI[0]}"'], fill=do, font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"{c['trang_thai']}{DONG_DAU}:{c['trang_thai']}{DONG_CUOI}",
                                  FormulaRule(formula=[f'{tt}="{TRANG_THAI[1]}"'], fill=vang))
    # Khách có tên mà chưa có term / chưa có mã → ô vàng cho người điền
    for k in ("term", "ma"):
        ws.conditional_formatting.add(
            f"{c[k]}{DONG_DAU}:{c[k]}{DONG_CUOI}",
            FormulaRule(formula=[f'AND(${c["ten"]}{DONG_DAU}<>"",{c[k]}{DONG_DAU}="")'], fill=vang))
    # Trùng mã kh → đỏ
    ws.conditional_formatting.add(
        f"{c['ma']}{DONG_DAU}:{c['ma']}{DONG_CUOI}",
        FormulaRule(formula=[f'AND({c["ma"]}{DONG_DAU}<>"",COUNTIF(${c["ma"]}${DONG_DAU}:${c["ma"]}${DONG_CUOI},'
                             f'{c["ma"]}{DONG_DAU})>1)'], fill=do))

    ws.freeze_panes = f"{c['sales']}{DONG_DAU}"
    ws.auto_filter.ref = f"A{DONG_TIEU_DE}:{get_column_letter(len(COT))}{DONG_CUOI}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{DONG_NHOM}:{DONG_TIEU_DE}"
    ws.print_area = f"A1:{get_column_letter(len(COT))}{DONG_DAU + max(len(dong), 1) + 4}"

    _huong_dan(wb.create_sheet(SHEET_HUONG_DAN), sales)
    wb.save(dich)
    return Path(dich)


def _huong_dan(ws, sales: Sequence[str]) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 70
    ws["A1"] = th.CONG_TY
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"] = "Hướng dẫn điền Client Master List"
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)

    def muc(r, ten, mau=th.CAM):
        o = ws.cell(row=r, column=1, value=ten)
        o.font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=mau)

    def dong(r, a, b="", fill=None):
        for col, v in ((2, a), (3, b)):
            o = ws.cell(row=r, column=col, value=v)
            o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o.alignment = Alignment(vertical="top", wrap_text=True)
            o.border = VIEN
            if fill and col == 2:
                o.fill = PatternFill("solid", fgColor=fill)

    muc(4, "1. Ai điền gì")
    dong(5, "Sales", "Thêm dòng khách mới ở cuối bảng; điền tên, PIC, MST, hợp đồng, thời hạn, ngày hết hạn.")
    dong(6, "FIN", "Điền cột Mã kh (mã khách trên Misa) và xác nhận Thời hạn đúng hợp đồng.")
    dong(7, "Tool báo cáo tuần", "Chỉ ĐỌC cột Mã kh và Thời hạn. Không ghi gì vào file này.")
    dong(8, "Cột nền xám", "Tự tính (STT, hết hạn sau gia hạn, còn lại, trạng thái) — đừng gõ đè.")
    dong(9, "Tiêu đề dòng 6", "Giữ nguyên chữ. Chèn cột thêm được; đổi tên cột Mã kh / Thời hạn thì tool không đọc được.")

    muc(11, "2. Màu trong bảng")
    dong(12, "Ô vàng ở Mã kh / Thời hạn", "Khách đã có tên nhưng còn thiếu — cần điền.", th.VANG_CANH_BAO)
    dong(13, "Mã kh nền đỏ", "Mã bị trùng ở hai dòng — giữ một dòng.", th.DO_CANH_BAO)
    dong(14, "Trạng thái nền đỏ", "Hợp đồng đã hết hạn — cần gia hạn.", th.DO_CANH_BAO)
    dong(15, "Trạng thái nền vàng", "Hết hạn trong 30 ngày tới.", th.VANG_CANH_BAO)
    dong(16, "Chưa có hạn HĐ", "Chưa nhập Ngày hết hạn (hoặc ngày gia hạn).")

    muc(18, "3. Cách ghi Thời hạn", th.TIM)
    ws.cell(row=19, column=2, value="Chọn trong danh sách hoặc gõ đúng một trong các kiểu dưới. "
                                    "Cột phải = tool báo cáo hiểu thế nào.").font = Font(
        name=th.PHONG_NOI_DUNG, size=10, italic=True)
    for col, t in ((2, "Thời hạn"), (3, "Tool hiểu là")):
        o = ws.cell(row=21, column=col, value=t)
        o.fill = PatternFill("solid", fgColor=th.TIM)
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    goi_y = list(ghep.GOI_Y_TERM)
    for i in range(8):
        t = goi_y[i] if i < len(goi_y) else ""
        dong(22 + i, t, ghep.hieu_term(t) if t else "")

    muc(31, "4. Danh sách Sales (nguồn thả xuống — có người mới thì thêm vào đây)", th.CAM)
    o = ws.cell(row=32, column=2, value="Sales")
    o.fill = PatternFill("solid", fgColor=th.CAM)
    o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    for i in range(20):
        dong(33 + i, sales[i] if i < len(sales) else "")

    muc(54, "5. Loại hợp đồng (nguồn thả xuống)", th.TIM)
    o = ws.cell(row=55, column=2, value="Loại hợp đồng")
    o.fill = PatternFill("solid", fgColor=th.TIM)
    o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    for i, t in enumerate(LOAI_HD):
        dong(56 + i, t)


# --- Đọc lại Client Master + cảnh báo hợp đồng cho báo cáo tuần (WR-08, Chairman 13/10) -------------
# Luật đã chốt: FIN là người sửa lần cuối file này; hợp đồng hết hạn chưa gia hạn vẫn tính tuổi nợ theo
# term cũ (tool không đổi term, chỉ cảnh báo); N3 chỉ hiện câu việc cần làm, FIN tự quyết ngưng dịch vụ.
NGUONG_SAP_HET = (30, 60, 90)


def _cong_thang(d, n: int):
    import calendar
    import datetime as dt
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def doc_master(thu_muc) -> Dict[str, Dict]:
    """{MÃ KH (hoa): {ten, term, so_hd, het_han, gia_han, han_dung, sales}} từ mọi bảng term trong sample/.
    Nhận cả file Client Master lẫn file nháp của KTT (tìm cột theo tên). File sửa sau thắng."""
    ra: Dict[str, Dict] = {}
    for ws, (r0, i_ma, i_term) in ghep._cac_bang_term(thu_muc):
        hang = [list(r) for r in ws.iter_rows(min_row=r0, values_only=True)]
        if not hang:
            continue
        tieu = [ghep._o(v).lower() for v in hang[0]]

        def cot(*ten, thu=0):
            vt = [i for i, x in enumerate(tieu) if x in ten]
            return vt[thu] if len(vt) > thu else None
        i_ten = cot("khách hàng", "customer's name", "tên khách")
        i_hd = cot("số hợp đồng", "contract no.")
        i_het = cot("ngày hết hạn", "expiry date")
        i_gh = cot("ngày gia hạn", "renewal date")
        i_sales = cot("sales phụ trách", "salesman")
        g = lambda r, i: r[i] if i is not None and i < len(r) else None
        for r in hang[1:]:
            ma = ghep._o(g(r, i_ma)).upper()
            if not ma or ma in ("MÃ KH", "CODE"):
                continue
            het, gh = ghep._ngay(g(r, i_het)), ghep._ngay(g(r, i_gh))
            ra[ma] = dict(ten=ghep._o(g(r, i_ten)), term=ghep._o(g(r, i_term)), so_hd=ghep._o(g(r, i_hd)),
                          het_han=het, gia_han=gh, han_dung=_cong_thang(gh, 12) if gh else het,
                          sales=ghep._o(g(r, i_sales)))
    return ra


def _chuan_term(t) -> str:
    return " ".join(str(t or "").lower().split())


def canh_bao_hop_dong(master: Dict[str, Dict], rui_ro, ngay_chot, term_truoc: Dict[str, str]) -> List[Dict]:
    """Danh sách cảnh báo cho sheet Contracts. rui_ro: DongRuiRo của tuần (khách còn nợ).
    muc: 'do' | 'vang' | '' — dùng tô màu."""
    ra: List[Dict] = []
    con_no = {r.ma.upper(): r for r in rui_ro}

    def them(loai, muc, ma, ten, sales, du_no, m, ghi_chu):
        han = (m or {}).get("han_dung")
        ra.append(dict(loai=loai, muc=muc, ma=ma, ten=ten, sales=sales, du_no=du_no,
                       so_hd=(m or {}).get("so_hd", ""), term=(m or {}).get("term", ""), han_dung=han,
                       con_lai=(han - ngay_chot).days if han else None, ghi_chu=ghi_chu))

    # 1–2. Hạn hợp đồng: đã hết hạn mà còn nợ / còn phát sinh; sắp hết hạn (mọi khách trong danh sách)
    for ma, m in master.items():
        han = m["han_dung"]
        if not han:
            continue
        r = con_no.get(ma)
        ten = (r.ten if r else "") or m["ten"]
        sales = (r.salesman if r else "") or m["sales"]
        du_no = r.cuoi_ky if r else 0.0
        con = (han - ngay_chot).days
        if con < 0 and r:
            ghi = f"Hết hạn {han:%d/%m/%Y}, vẫn còn nợ" + (f"; tuần này phát sinh {r.tang:,.0f}".replace(",", ".") if r.tang > 0.5 else "")
            them("Hợp đồng đã hết hạn", "do", ma, ten, sales, du_no, m, ghi + ". Tuổi nợ vẫn tính theo term cũ.")
        elif 0 <= con <= NGUONG_SAP_HET[-1]:
            moc = next(n for n in NGUONG_SAP_HET if con <= n)
            them(f"Sắp hết hạn ≤{moc} ngày", "vang" if moc == 30 else "", ma, ten, sales, du_no, m,
                 f"Hết hạn {han:%d/%m/%Y} — Sales làm gia hạn, FIN cập nhật Ngày gia hạn.")
    # 3–4. Khách đang nợ: chưa có term trong danh sách · term vừa bị đổi (thiếu số HĐ / ngày hết hạn: chỉ đếm)
    for ma, r in con_no.items():
        m = master.get(ma)
        if not m or not m["term"]:
            them("Chưa có credit term", "vang", r.ma, r.ten, r.salesman, r.cuoi_ky, m,
                 f"Chưa có trong Client Master — tool đang tính theo: {r.credit_term or '30 ngày (mặc định)'}.")
            continue
        cu = term_truoc.get(ma)
        if cu and _chuan_term(cu) != _chuan_term(r.credit_term):
            them("Credit term vừa đổi", "vang", r.ma, r.ten, r.salesman, r.cuoi_ky, m,
                 f"Tuần trước: {cu} → tuần này: {r.credit_term}. FIN xác nhận đúng hợp đồng.")
    thu_tu = {"Hợp đồng đã hết hạn": 0, "Credit term vừa đổi": 1, "Chưa có credit term": 2}
    ra.sort(key=lambda x: (thu_tu.get(x["loai"], 3 if x["loai"].startswith("Sắp") else 4),
                           x["con_lai"] if x["con_lai"] is not None else 0, -x["du_no"]))
    return ra
