# -*- coding: utf-8 -*-
"""v0.4 — full credit report (WR-04).

Một lần chạy tính đủ các phần mà file Excel tuần cần:
  ① Receivables  tiền theo Tổng hợp 131, tuổi theo hoá đơn trong Chi tiết 131
  ② AR Risk      luật rủi ro của chị kế toán trưởng (sheet "data W39 so W40" + "Tham chieu")
  ③ Payable      Tổng hợp 331
  ④ Cash Flow    phiếu thu trong kỳ tách theo TK đối ứng + tiền đã trả vendor
  ⑤ Summary      các con số tuần này, lưu vào sổ ghi để làm lịch sử nhiều tuần
Luật chốt: handoff/20261008_WR-04_full-credit-report.md, handoff/MAP.md.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import chitiet as ctm
from . import ghep
from . import soghi as sg
from . import thuonghieu as th_
from .misa import PhaiTra, TongHop, chuan_hoa_ten
from .tinhtoan import DongBaoCao, _cac_job, _chenh_ty_gia

# --- tham số của chị KTT (sheet "Tham chieu") ------------------------------
SO_DAU_HIEU_NO_XAU = 2
NGUONG_CEO = 100_000_000
TU_KHOA_LOAI_TRU = ("cấn trừ", "bù trừ", "thanh toán hộ", "supplier của mình")
TU_KHOA_PHAP_LY = ("đang kiện", "mất liên lạc", "khó đòi", "công nợ xấu", "phá sản")

N0, NX, N4, N3, N2, N1 = ("0·Chưa đến hạn", "X·Không phải rủi ro tín dụng",
                          "N4·Nợ xấu–sự kiện pháp lý", "N3·Nợ xấu", "N2·Cần theo dõi",
                          "N1·Quá hạn bình thường")


@dataclass
class DongRuiRo:
    ma: str
    ten: str
    credit_term: str
    dau_ky: float          # Misa: dư đầu kỳ
    tang: float            # Misa: phát sinh Nợ (nợ mới)
    giam: float            # Misa: phát sinh Có (đã thu, cấn trừ, tỷ giá)
    cuoi_ky: float         # = Total của Receivables
    nhom: List[float]
    salesman: str
    job: str
    ly_do: str
    tt_total: float        # tuần trước (U)
    tt_qua_han: float      # (V)
    tt_61: float           # (W)
    co_tuan_truoc: bool
    nhap_bo_sung: Optional[float]   # dư đầu kỳ Misa − dư cuối báo cáo tuần trước
    s: List[int] = field(default_factory=lambda: [0] * 6)
    so_dau_hieu: int = 0
    ma_dau_hieu: str = ""
    nhom_rui_ro: str = ""
    xu_huong: str = ""
    viec_can_lam: str = ""
    ceo: bool = False

    @property
    def qua_han(self) -> float:
        return round(sum(self.nhom[1:]), 2)


@dataclass
class DongTra:
    ma: str
    ten: str
    dau_ky: float
    no_moi: float
    da_tra: float
    cuoi_ky: float
    trang_thai: str
    nhap_bo_sung: Optional[float] = None


@dataclass
class DongThu:
    ma: str
    ten: str
    theo_nhom: Dict[str, float]

    @property
    def tong(self) -> float:
        return round(sum(self.theo_nhom.values()), 2)


@dataclass
class KetQuaFull:
    dong: List[DongBaoCao]
    rui_ro: List[DongRuiRo] = field(default_factory=list)
    phai_tra: Optional[List[DongTra]] = None
    thu: List[DongThu] = field(default_factory=list)
    tra_vendor: List[DongTra] = field(default_factory=list)
    chi_so: Dict[str, float] = field(default_factory=dict)
    chi_so_truoc: Dict[str, float] = field(default_factory=dict)
    lich_su: List[dict] = field(default_factory=list)
    da_bo: List[tuple] = field(default_factory=list)
    lech_chi_tiet: List[tuple] = field(default_factory=list)    # tên, Tổng hợp, Chi tiết
    khach_tra_xong: List[tuple] = field(default_factory=list)   # tên, dư tuần trước
    vendor_tra_truoc: List[tuple] = field(default_factory=list) # mã, tên, dư Nợ
    sms_cat: List[tuple] = field(default_factory=list)
    sms_khong_misa: List[tuple] = field(default_factory=list)
    sms_chua_ghep: List[tuple] = field(default_factory=list)
    # v0.5 (WR-06): đối chiếu job chưa thanh toán trên Misa
    job_misa: bool = False                                      # có file Bán hàng hay không
    het_no_con_hd: List[tuple] = field(default_factory=list)    # tên, số HĐ, tiền: Misa hết nợ, HĐ vẫn "chưa TT"
    ten_khong_ghep: List[tuple] = field(default_factory=list)   # tên, số HĐ, tiền: không tìm ra khách Misa


def _co(chu: str, tu_khoa) -> bool:
    c = (chu or "").lower()
    return any(t in c for t in tu_khoa)


def _ghep_sms(th: TongHop, dong_sms, bang):
    """Gán từng dòng SMS cho một mã khách Misa (bảng ghép > trùng mã > trùng tên)."""
    con_no = {k.ma for k in th.khach_con_no()}
    theo_ma = {k.ma.upper(): k.ma for k in th.dong.values()}
    theo_ten = {k.ten_chuan: k.ma for k in th.dong.values()}
    theo_khach: Dict[str, list] = {}
    khong_misa, chua_ghep = [], {}
    for x in dong_sms or []:
        ma, cach = ghep.ma_misa_cua(x, bang, theo_ma, theo_ten)
        if ma is not None:
            ma = theo_ma.get(ma.upper(), ma)
        if ma is None and cach == "chưa ghép":
            chua_ghep.setdefault(x.ma_sms, []).append(x)
        elif ma is None or ma not in con_no:
            ly_do = ("FIN khai không có trên Misa" if ma is None else
                     f"Misa không còn nợ ({ma}) — nhiều khả năng đã thu, SMS chưa gạch paid")
            khong_misa.append((x.ma_sms, x.ten_sms, x.job, x.so_tien, ly_do))
        else:
            theo_khach.setdefault(ma, []).append(x)
    gop: Dict[tuple, list] = {}
    for ma, ten, job, tien, ly_do in khong_misa:
        g = gop.setdefault((ma, job), [ma, ten, job, 0.0, ly_do])
        g[3] += tien
    chua = [(m, ds[0].ten_sms, sum(x.so_tien for x in ds), len(ds)) for m, ds in chua_ghep.items()]
    return theo_khach, [tuple(g) for g in gop.values()], chua


def _xu_huong(r: DongRuiRo) -> str:
    """Câu tự sinh, chép đúng công thức cột 'Trend (Auto)' của chị KTT."""
    s1, s2, s3, s4, s5, s6 = r.s
    q, v, p = r.qua_han, r.tt_qua_han, r.nhom[5]
    c = []
    if q < 0.5:
        c.append("Chưa đến hạn thanh toán.")
    if not r.co_tuan_truoc or r.tt_total < 0.5:
        c.append("Khách mới xuất hiện trong bảng tuần này.")
    if s1:
        c.append("Khách phát sinh lô mới nhưng nợ cũ không giảm.")
    if s2:
        c.append("Quá hạn đứng yên so với tuần trước, chưa thu được đồng nào.")
    if r.tt_total > 0.5 and 0.5 < q < v - 0.5:
        c.append(f"Đã thu bớt {v - q:,.0f} so với tuần trước.")
    if s3:
        c.append("Nợ đang già đi, tiền chuyển sang nhóm tuổi cao hơn.")
    if s6:
        c.append("Toàn bộ dư nợ đã quá hạn.")
    if p > 0.5:
        c.append(f"Có {p:,.0f} nằm ở nhóm trên 120 ngày.")
    if s5:
        c.append("Chưa có điều khoản thanh toán nên chưa xác định chắc chắn được quá hạn.")
    if s4:
        c.append("CHƯA CÓ GIẢI TRÌNH TỪ SALES.")
    return " ".join(c)


def _chia_tuoi_dong_con(d: DongBaoCao) -> None:
    """Chia sáu nhóm tuổi của khách cho các dòng con (Chairman 08/10: số tiền job phải nằm đúng
    cột 1-30, 31-60...). Không có ngày hoá đơn cho từng job, nên phân bổ theo thứ tự: nợ cũ không
    có job nhận nhóm già nhất trước, rồi tới job cũ nhất. Cộng các dòng con theo từng cột luôn
    bằng đúng dòng khách."""
    con_lai = list(d.nhom)
    for g in d.dong_con:                  # dòng FIN khai theo hoá đơn: nhóm tuổi đã biết chắc
        if g.get("nhom_co_dinh"):
            g["nhom"] = g["nhom_co_dinh"]
            con_lai = [round(a - b, 2) for a, b in zip(con_lai, g["nhom"])]
    con = [g for g in d.dong_con if not g.get("nhom_co_dinh")]
    if d.can_doi > 0:
        con = [{"_can_doi": True, "tien": d.can_doi}] + sorted(con, key=lambda g: (g.get("ngay") or dt.date.min))
    else:
        con = sorted(con, key=lambda g: (g.get("ngay") or dt.date.min))
    for g in con:
        tien, nhom = g["tien"], [0.0] * 6
        for i in range(5, -1, -1):
            lay = min(max(con_lai[i], 0), tien)
            if lay > 0:
                nhom[i] = round(lay, 2)
                con_lai[i] = round(con_lai[i] - lay, 2)
                tien = round(tien - lay, 2)
        if tien > 0.5:                     # job SMS nhiều hơn phần còn lại (tỷ giá): để ở Current
            nhom[0] = round(nhom[0] + tien, 2)
        if g.get("_can_doi"):
            d.can_doi_nhom = nhom
        else:
            g["nhom"] = nhom


# --- v0.5 (WR-06): đối chiếu dư nợ Misa với tổng job chưa thanh toán ---------------------------
DUNG_SAI_KHOP = 1_000          # VND: làm tròn
TY_LE_LECH_NHO = 0.02          # lệch ≤ 2% dư nợ: nghi tỷ giá
MAU_DOI_CHIEU = {"khop": "Khớp", "do": "Đỏ", "vang": "Vàng", "cam": "Cam", "tim": "Tím"}


def _dong_con_misa(d: DongBaoCao, hds, term: str, tai_ngay: dt.date, sales_job: Dict[str, str],
                   sales_cu: str) -> None:
    """Mỗi job của hoá đơn chưa thanh toán (luật FIN: chỉ 'Đã thanh toán' mới bỏ) thành một dòng
    con. Biết ngày hoá đơn nên biết đúng nhóm tuổi của từng job."""
    for h in sorted(hds, key=lambda h: h.ngay):
        han = sg.han_tu_term(h.ngay, term)
        i = sg.nhom_tu_so_ngay((tai_ngay - han).days)
        for job, tien in h.jobs.items():
            nhom = [0.0] * 6
            nhom[i] = tien
            d.dong_con.append({"job": job or f"HĐ {h.so_hd} (chưa có job)",
                               "sales": sales_job.get(job.upper(), "") if job else "",
                               "ngay": h.ngay, "nguon_ngay": "HĐ", "tien": tien, "so_hd": h.so_hd,
                               "trang_thai": (h.trang_thai if h.chua_tt else
                                              f"{h.trang_thai} ngày {h.thu_ngay:%d/%m} (sau kỳ)"),
                               "nhom_co_dinh": nhom, "misa": True})
    if d.dong_con:
        d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con if not g["job"].startswith("HĐ "))))
        tien_sale: Dict[str, float] = {}
        for g in d.dong_con:
            if g["sales"]:
                tien_sale[g["sales"]] = tien_sale.get(g["sales"], 0) + g["tien"]
        if tien_sale:
            d.salesman = ", ".join(sorted(tien_sale, key=lambda t: -tien_sale[t]))
        elif sales_cu:
            d.salesman = sales_cu


def doi_chieu(T: float, dong_con: List[dict]) -> dict:
    """FIN 09/10: cộng các job chưa thanh toán, so với dư nợ của khách; lệch thì báo màu để FIN
    xem lại (quên đánh giá chênh lệch tỷ giá, thanh toán một phần mà Misa không tự trừ...)."""
    S = round(sum(g["tien"] for g in dong_con), 2)
    L = round(T - S, 2)
    mot_phan = sum(1 for g in dong_con if "một phần" in (g.get("trang_thai") or "").lower())
    if abs(L) <= DUNG_SAI_KHOP:
        mau, goi_y = "khop", ""
    elif not dong_con:
        mau, goi_y = "cam", ("Không có hoá đơn chưa thanh toán nào trên file Bán hàng: nợ trước năm nay "
                             "(điền bảng job nợ cũ) hoặc tên khách hai file khác nhau.")
    elif abs(L) <= TY_LE_LECH_NHO * max(T, S):
        mau, goi_y = "vang", "Lệch nhỏ: nghi quên đánh giá chênh lệch tỷ giá, hoặc làm tròn."
    elif L < 0:
        mau = "do"
        goi_y = ("Tổng job chưa thanh toán LỚN hơn dư nợ: đã thu nhưng chưa đối trừ vào hoá đơn, "
                 "hoặc thanh toán một phần mà Misa chưa tự trừ.")
        if mot_phan:
            goi_y += f" Có {mot_phan} dòng 'Thanh toán một phần'."
    else:
        mau, goi_y = "cam", ("Dư nợ LỚN hơn tổng job chưa thanh toán: có nợ chưa gắn job (nợ trước năm "
                             "nay → điền bảng job nợ cũ) hoặc hoá đơn chưa lên file Bán hàng.")
    return {"mau": mau, "lech": L, "tong_job": S, "goi_y": goi_y}


def danh_gia_rui_ro(r: DongRuiRo) -> DongRuiRo:
    J, Q = r.cuoi_ky, r.qua_han
    U, V, W = r.tt_total, r.tt_qua_han, r.tt_61
    nop = sum(r.nhom[3:])
    r.s = [
        int(U > 0.5 and J > U + 0.5 and Q >= V - 0.5),            # S1 nợ tăng, quá hạn không giảm
        int(Q > 0.5 and abs(Q - V) < 1),                          # S2 quá hạn đứng yên
        int(nop > W + 0.5),                                       # S3 nợ trên 60 ngày tăng
        int(not r.ly_do.strip()),                                 # S4 chưa có giải trình
        int(not r.credit_term.strip()),                           # S5 chưa có term
        int(J > 0.5 and abs(Q - J) < 1),                          # S6 toàn bộ đã quá hạn
    ]
    r.so_dau_hieu = sum(r.s[:3])
    r.ma_dau_hieu = " ".join(f"S{i + 1}" for i, v in enumerate(r.s) if v)
    if Q < 0.5:
        r.nhom_rui_ro = N0
    elif _co(r.ly_do, TU_KHOA_LOAI_TRU):
        r.nhom_rui_ro = NX
    elif _co(r.ly_do, TU_KHOA_PHAP_LY):
        r.nhom_rui_ro = N4
    elif r.nhom[5] > 0.5 or r.so_dau_hieu >= SO_DAU_HIEU_NO_XAU:
        r.nhom_rui_ro = N3
    elif r.so_dau_hieu >= 1:
        r.nhom_rui_ro = N2
    else:
        r.nhom_rui_ro = N1
    if r.nhom_rui_ro == N4:
        r.viec_can_lam = ("Chuyển hồ sơ thu hồi. Rà soát chứng từ gốc để xét trích dự phòng "
                          "theo TT48/2019.")
    elif r.nhom_rui_ro == N3:
        r.viec_can_lam = ("Dừng cấp dịch vụ mới. Chốt công văn đòi nợ có xác nhận gửi. Kế toán "
                          "trưởng ký xác nhận phân nhóm." if r.nhom[5] > 0.5 else
                          "Dừng cấp dịch vụ mới. Yêu cầu cam kết NGÀY cụ thể bằng văn bản trong 7 ngày.")
    r.ceo = Q >= NGUONG_CEO and not r.ly_do.strip()
    r.xu_huong = _xu_huong(r)
    return r


def tinh_full(th: TongHop, ct: ctm.ChiTiet, so: sg.SoGhi, dong_sms=None, bang=None,
              pt: Optional[PhaiTra] = None, ngay_chot: Optional[dt.date] = None,
              job_no_cu: Optional[Dict[str, list]] = None, job_misa=None) -> KetQuaFull:
    # Ngày chốt (tên trang sổ ghi, "At dd Mon yyyy") = hôm sau ngày cuối kỳ, như file Tuổi nợ cũ.
    # Tuổi nợ tính tại ngày cuối kỳ: hoá đơn đến hạn đúng ngày đó vẫn là "Current".
    ngay_chot = ngay_chot or th.den_ngay + dt.timedelta(days=1)
    tai_ngay = th.den_ngay
    truoc = so.trang_truoc(ngay_chot) or {}
    truoc_theo_ma = {v.ma.upper(): v for v in truoc.values() if v.ma}
    theo_khach, khong_misa, chua_ghep = _ghep_sms(th, dong_sms, bang)
    kq = KetQuaFull(dong=[], sms_khong_misa=khong_misa, sms_chua_ghep=chua_ghep)
    sales_job = {}
    for x in dong_sms or []:
        for j in _cac_job(x.job or ""):
            if x.salesman and th_.ten_sales(x.salesman):
                sales_job.setdefault(j.upper(), th_.ten_sales(x.salesman))
    if job_misa is not None:
        kq.job_misa = True
        for ten, n, tien in job_misa.khong_ghep_tai(tai_ngay):
            kct_ = ct.khach.get(chuan_hoa_ten(ten))
            if kct_ is not None and abs(kct_.so_du(tai_ngay)) < 1:
                kq.het_no_con_hd.append((ten, n, tien))
            else:
                kq.ten_khong_ghep.append((ten, n, tien))
    kho_doi = bang.kho_doi if bang else {}

    for k in th.khach_con_no():
        if k.ten_chuan in so.bo_qua:
            kq.da_bo.append((k.ten, k.du_no, so.bo_qua[k.ten_chuan].get("ly_do", "")))
            continue
        T = round(k.du_no, 2)
        tr = truoc.get(k.ten_chuan) or truoc_theo_ma.get(k.ma.upper())
        term = (so.dieu_chinh.get(k.ten_chuan, {}).get("credit_term")
                or (tr.credit_term if tr else "") or "")
        term_ngay = sg.so_ngay_term(term)
        d = DongBaoCao(ma=k.ma, ten=k.ten, ten_chuan=k.ten_chuan, credit_term=term, total=T,
                       nhom=[0.0] * 6, salesman=tr.salesman if tr else "",
                       ghi_chu=tr.ghi_chu if tr else "", ly_do=tr.ly_do if tr else "")
        ngoai = so.ngoai_le.get(k.ten_chuan)
        kct = ct.khach.get(k.ten_chuan)
        if ngoai and "nhom" in ngoai:
            d.nhom[sg.TEN_NHOM.index(ngoai["nhom"])] = T
            d.lo = sg.lo_tu_nhom(d.nhom, tai_ngay)
            d.nguon, d.phan_nguon = "ngoai-le", {"ngoai-le": T}
            d.cho_xem_lai.append(f"Ngoại lệ do kế toán khai: xếp cả {T:,.0f} vào nhóm {ngoai['nhom']}")
        elif k.ma in kho_doi or "khó đòi" in d.ly_do.lower():
            d.nhom[5] = T
            d.lo = sg.lo_tu_nhom(d.nhom, tai_ngay)
            d.nguon, d.phan_nguon = "kho-doi", {"kho-doi": T}
            d.ly_do = d.ly_do or kho_doi.get(k.ma, "Khó đòi")
        else:
            con, _ = kct.hoa_don_con_no(tai_ngay) if kct else ([], 0.0)
            # mỗi hoá đơn: [ngày HĐ, số HĐ, hạn, còn nợ] — giữ lại cho sheet Invoices
            hd = [[n, so_hd, sg.han_tu_term(n, term), t] for n, so_hd, t in con]
            S = round(sum(x[3] for x in hd), 2)
            phan = {"hoa-don": min(S, T)} if S else {}
            if kct is None:
                d.cho_xem_lai.append("Không thấy khách này trong file Chi tiết 131 (tên khác nhau?) — "
                                     "tạm tính là nợ mới")
            if S > T + 1:
                du = round(S - T, 2)        # Chi tiết ghi nợ nhiều hơn Tổng hợp: trừ tiếp vào HĐ cũ
                for x in hd:
                    tru = min(x[3], du)
                    x[3], du = round(x[3] - tru, 2), round(du - tru, 2)
                    if du <= 0:
                        break
                hd = [x for x in hd if x[3] > 0.5]
            if abs(S - T) >= 1:
                kq.lech_chi_tiet.append((k.ten, T, S))
                d.cho_xem_lai.append(f"Chi tiết 131 ghi {S:,.0f}, Tổng hợp 131 ghi {T:,.0f}: tiền theo "
                                     "Tổng hợp. Thường do hai file xuất khác thời điểm")
            lo = [sg.Lo(x[3], x[2].isoformat()) for x in hd]
            if S < T - 1:
                thieu = round(T - S, 2)
                han = tai_ngay + dt.timedelta(days=term_ngay)
                lo.append(sg.Lo(thieu, han.isoformat(), True))
                hd.append([None, "(Tổng hợp > Chi tiết)", han, thieu])
                phan["uoc-tinh"] = thieu
            d.hoa_don = [tuple(x) for x in hd]
            d.lo = lo
            d.nhom = sg.nhom_tu_lo(lo, tai_ngay)
            d.phan_nguon = phan
            d.nguon = "uoc-tinh" if "uoc-tinh" in phan else "hoa-don"

        # job và salesman từ SMS — mỗi job còn nợ thành một dòng con dưới khách (CEO 08/10:
        # sếp cần biết ai đang làm job của khách nợ để hỏi thẳng)
        ds = theo_khach.get(k.ma, [])
        if job_misa is not None:
            # v0.5 (WR-06): job + tiền lấy từ Misa (Bán hàng + Sổ chi tiết), SMS chỉ cho tên sales
            _dong_con_misa(d, job_misa.chua_tt_cua(k.ma, tai_ngay), term, tai_ngay, sales_job,
                           tr.salesman if tr else "")
            for h in job_misa.theo_ma.get(k.ma, []):
                n = ctm._so_hd(h.so_hd)
                if n is not None:
                    d.hd_misa[(n, h.ngay.year)] = (h.trang_thai, ", ".join(j for j in h.jobs if j))
        elif ds:
            jobs = [[x, x.so_tien] for x in sorted(ds, key=lambda x: (x.ngay or dt.date.min))]
            S = sum(x.so_tien for x in ds)
            if S > T + _chenh_ty_gia(T) + 1 and d.nguon not in {"kho-doi", "ngoai-le"}:
                can_bo = S - T              # SMS ghi nhiều hơn Misa: job cũ nhất đã thu, chưa gạch paid
                for j in jobs:
                    if can_bo <= 0.5:
                        break
                    bo = min(j[1], can_bo)
                    can_bo -= bo
                    j[1] -= bo
                    kq.sms_cat.append((k.ten, j[0].job, bo, j[0].ngay))
            gop: Dict[str, dict] = {}
            for x, con in jobs:
                if con <= 0.5:
                    continue
                g = gop.setdefault(x.job, {"job": x.job, "sales": th_.ten_sales(x.salesman),
                                           "ngay": x.ngay, "nguon_ngay": x.nguon_ngay, "tien": 0.0})
                g["tien"] = round(g["tien"] + con, 2)
            d.dong_con = sorted(gop.values(), key=lambda g: (g["ngay"] or dt.date.min))
            d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con)))
            tien_sale: Dict[str, float] = {}
            for g in d.dong_con:
                if g["sales"]:
                    tien_sale[g["sales"]] = tien_sale.get(g["sales"], 0) + g["tien"]
            if tien_sale:
                d.salesman = ", ".join(sorted(tien_sale, key=lambda t: -tien_sale[t]))
        elif tr:
            d.job = ", ".join(_cac_job(tr.job))
        # nợ cũ không có job trên SMS: FIN khai một lần ở sample/ (hoá đơn nào thuộc job nào, Sales nào)
        con_lai = round(T - sum(g["tien"] for g in d.dong_con), 2)
        if job_misa is not None:
            con_lai = float("inf")            # nợ cũ trước năm nay: FIN khai theo hoá đơn, không giới hạn
        if job_no_cu and con_lai >= 1 and d.hoa_don:
            mo = {}
            for ngay, so_hd, han, tien in d.hoa_don:
                if job_misa is not None and (ngay is None or ngay >= job_misa.tu_ngay):
                    continue                  # hoá đơn năm nay đã có job trên Misa
                n = ctm._so_hd(str(so_hd))
                if n is not None:
                    mo.setdefault(n, []).append((ngay, han, tien))
            for so_hd, job, sales in job_no_cu.get(k.ma.upper(), []):
                if con_lai < 1:
                    break
                for ngay, han, tien in mo.pop(so_hd, []):
                    lay = round(min(tien, con_lai), 2)
                    nhom = [0.0] * 6      # biết đúng hoá đơn nên biết đúng nhóm tuổi
                    nhom[sg.nhom_tu_so_ngay((tai_ngay - han).days)] = lay
                    d.dong_con.append({"job": job or f"HĐ {so_hd}", "sales": th_.ten_sales(sales), "ngay": ngay,
                                       "nguon_ngay": "HĐ", "tien": lay, "tham_chieu": True, "nhom_co_dinh": nhom})
                    con_lai = round(con_lai - lay, 2)
            if d.dong_con:
                d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con)))
                if not d.salesman:
                    d.salesman = ", ".join(dict.fromkeys(g["sales"] for g in d.dong_con if g["sales"]))
        if job_misa is not None:
            d.doi_chieu = doi_chieu(T, d.dong_con)
            du_ct = kct.so_du(tai_ngay) if kct else None
            if (d.doi_chieu["mau"] != "khop" and du_ct is not None and abs(du_ct - T) > DUNG_SAI_KHOP
                    and abs(du_ct - d.doi_chieu["tong_job"]) <= DUNG_SAI_KHOP):
                # FIN 09/10 (GOLDEN GLOBE): job khớp Chi tiết 131, chỉ Tổng hợp 131 lệch → file xuất khác lúc
                d.doi_chieu["goi_y"] = (f"Tổng job khớp Chi tiết 131 ({du_ct:,.0f}); chỉ file Tổng hợp 131 ghi "
                                        f"{T:,.0f}: hai file xuất khác lúc, Misa đã sửa hoá đơn ở giữa. "
                                        "Xuất lại cả 6 file cùng lúc là khớp.")
            if d.nguon == "kho-doi" and d.doi_chieu["mau"] == "cam" and not d.dong_con:
                d.doi_chieu.update(mau="kho_doi", goi_y="Nợ khó đòi FIN đã khai ở sample/: không cần job.")
        elif d.dong_con and abs(con_lai) >= 1:
            d.can_doi = con_lai
        _chia_tuoi_dong_con(d)
        if not term:
            d.thieu_thong_tin.append("Chưa có Credit Term, tạm tính 30 ngày")
        d.salesman = th_.ten_sales(d.salesman)
        if not d.salesman:
            d.thieu_thong_tin.append("Chưa có Salesman")
        lech = round(d.total - sum(d.nhom), 2)
        if abs(lech) >= 1:
            d.nhom[0] = round(d.nhom[0] + lech, 2)
            d.cho_xem_lai.append(f"Tổng các nhóm lệch Total {lech:,.0f}, đã dồn vào Current")
        kq.dong.append(d)

        r = DongRuiRo(
            ma=k.ma, ten=k.ten, credit_term=term, dau_ky=round(k.dau_ky, 2), tang=round(k.ps_no, 2),
            giam=round(k.ps_co, 2), cuoi_ky=T, nhom=list(d.nhom), salesman=d.salesman, job=d.job,
            ly_do=d.ly_do, tt_total=round(tr.tong, 2) if tr else 0.0,
            tt_qua_han=round(sum(tr.nhom[1:]), 2) if tr else 0.0,
            tt_61=round(sum(tr.nhom[3:]), 2) if tr else 0.0, co_tuan_truoc=tr is not None,
            nhap_bo_sung=round(k.dau_ky - tr.tong, 2) if tr else None,
        )
        kq.rui_ro.append(danh_gia_rui_ro(r))
        if not d.ly_do.strip():
            d.ly_do_tu_dong = f"{th_.NHAN_TU_DONG} {r.xu_huong}".strip()

    if job_misa is not None:
        # khách có trên Tổng hợp nhưng không còn dư nợ, mà hoá đơn vẫn ghi chưa thanh toán
        da_co = {d.ma for d in kq.dong} | {k.ma for k in th.khach_con_no()}
        ten_ma = {k.ma: k.ten for k in th.dong.values()}
        for ma in job_misa.theo_ma:
            if ma in da_co:
                continue
            hds = job_misa.chua_tt_cua(ma, tai_ngay)
            if hds:
                kq.het_no_con_hd.append((ten_ma.get(ma, ma), len(hds), round(sum(h.tien for h in hds), 2)))
        kq.het_no_con_hd.sort(key=lambda x: -x[2])

    # gộp SMS cần gạch paid theo job
    gop: Dict[tuple, list] = {}
    for khach_, job, tien, ngay in kq.sms_cat:
        g = gop.setdefault((khach_, job), [khach_, job, 0.0, ngay])
        g[2] += tien
    kq.sms_cat = [tuple(g) for g in gop.values()]
    kq.dong.sort(key=lambda d: (-d.nhom_gia_nhat, -d.total))
    thu_tu = {N4: 0, N3: 1, N2: 2, N1: 3, NX: 4, N0: 5}
    kq.rui_ro.sort(key=lambda r: (thu_tu.get(r.nhom_rui_ro, 9), -r.qua_han, -r.cuoi_ky))

    # khách tuần trước còn nợ, tuần này hết
    con = {k.ten_chuan for k in th.khach_con_no()}
    for tc, hs in truoc.items():
        if tc not in con and hs.tong > 0.5:
            kq.khach_tra_xong.append((hs.ten, hs.tong))

    # --- tiền thu trong kỳ, tách theo TK đối ứng
    for tc, kh in ct.khach.items():
        theo = kh.thu_trong_ky(th.tu_ngay, th.den_ngay)
        if theo:
            ma = th.dong[tc].ma if tc in th.dong else ""
            kq.thu.append(DongThu(ma=ma, ten=kh.ten, theo_nhom=theo))
    kq.thu.sort(key=lambda x: -x.tong)

    # --- phải trả
    if pt is not None:
        truoc_tra = so.phai_tra_truoc(ngay_chot)
        kq.phai_tra = []
        for v in pt.dong:
            dau = round(v.dk_co - v.dk_no, 2)
            cuoi = round(v.du_co, 2)
            if v.du_no > 0.5:
                kq.vendor_tra_truoc.append((v.ma, v.ten, v.du_no))
            if cuoi <= 0.5 and not (dau > 0.5 and v.ps_no > 0.5):
                if v.ps_no > 0.5:
                    kq.tra_vendor.append(DongTra(v.ma, v.ten, dau, v.ps_co, v.ps_no, cuoi, ""))
                continue
            if cuoi <= 0.5:
                tt = "Paid off this week"
            elif dau <= 0.5:
                tt = "New this week"
            elif v.ps_no > 0.5:
                tt = "Partly paid"
            else:
                tt = "Outstanding"
            dong_tra = DongTra(v.ma, v.ten, dau, round(v.ps_co, 2), round(v.ps_no, 2), cuoi, tt,
                               round(dau - truoc_tra.get(v.ma, 0.0), 2) if truoc_tra is not None else None)
            kq.phai_tra.append(dong_tra)
            if v.ps_no > 0.5:
                kq.tra_vendor.append(dong_tra)
        kq.phai_tra.sort(key=lambda x: (x.trang_thai == "Paid off this week", -x.cuoi_ky))
        kq.tra_vendor.sort(key=lambda x: -x.da_tra)

    kq.chi_so = chi_so_tuan(kq, th, pt, ngay_chot)
    kq.chi_so_truoc = _chi_so_truoc(so, ngay_chot, truoc)
    return kq


def chi_so_tuan(kq: KetQuaFull, th: TongHop, pt: Optional[PhaiTra], ngay_chot: dt.date) -> dict:
    nhom = [round(sum(d.nhom[i] for d in kq.dong), 2) for i in range(6)]
    thu: Dict[str, float] = {}
    for x in kq.thu:
        for k, v in x.theo_nhom.items():
            thu[k] = round(thu.get(k, 0) + v, 2)
    xau = [r for r in kq.rui_ro if r.nhom_rui_ro in (N3, N4)]
    cs = {
        "ngay_chot": ngay_chot.isoformat(),
        "tuan": f"W{th.den_ngay.isocalendar()[1]:02d}",
        "tu_ngay": th.tu_ngay.isoformat(), "den_ngay": th.den_ngay.isoformat(),
        "tong_phai_thu": round(sum(d.total for d in kq.dong), 2),
        "chua_den_han": nhom[0], "qua_han": round(sum(nhom[1:]), 2),
        "qua_han_1_30": nhom[1], "qua_han_31_60": nhom[2], "qua_han_61_90": nhom[3],
        "qua_han_91_120": nhom[4], "qua_han_120": nhom[5],
        "so_khach": len(kq.dong),
        "no_moi": round(sum(k.ps_no for k in th.dong.values()), 2),
        "da_thu": round(sum(k.ps_co for k in th.dong.values()), 2),
        "da_thu_tien": round(thu.get("ngan-hang", 0) + thu.get("tien-mat", 0), 2),
        "thu_ngan_hang": thu.get("ngan-hang", 0.0), "thu_tien_mat": thu.get("tien-mat", 0.0),
        "thu_can_tru": thu.get("can-tru", 0.0), "thu_ty_gia_phi": thu.get("ty-gia-phi", 0.0),
        "thu_khac": thu.get("khac", 0.0),
        "no_xau": round(sum(r.qua_han for r in xau), 2), "so_khach_no_xau": len(xau),
        "no_xau_120": round(sum(r.nhom[5] for r in xau), 2),
        "no_xau_phap_ly": round(sum(r.qua_han for r in xau if r.nhom_rui_ro == N4), 2),
        "ceo_can_xem": sum(1 for r in kq.rui_ro if r.ceo),
    }
    if pt is not None:
        cs.update({
            "tong_phai_tra": round(pt.tong_con_no, 2), "da_tra_vendor": round(pt.tong_da_tra, 2),
            "no_vendor_moi": round(pt.tong_no_moi, 2),
            "so_vendor": sum(1 for v in pt.dong if v.du_co > 0.5),
            "dong_tien_rong": round(cs["da_thu_tien"] - pt.tong_da_tra, 2),
        })
    return cs


def _chi_so_truoc(so: sg.SoGhi, ngay_chot: dt.date, truoc: dict) -> dict:
    cu = [n for n in so.lich_su_tuan if dt.date.fromisoformat(n) < ngay_chot]
    if cu:
        return dict(so.lich_su_tuan[max(cu)])
    if not truoc:
        return {}
    # chưa có lịch sử (tuần đầu dùng v0.4): dựng lại phần phải thu từ trang sổ ghi tuần trước
    nhom = [round(sum(h.nhom[i] for h in truoc.values()), 2) for i in range(6)]
    return {"tong_phai_thu": round(sum(h.tong for h in truoc.values()), 2), "chua_den_han": nhom[0],
            "qua_han": round(sum(nhom[1:]), 2), "qua_han_1_30": nhom[1], "qua_han_31_60": nhom[2],
            "qua_han_61_90": nhom[3], "qua_han_91_120": nhom[4], "qua_han_120": nhom[5],
            "so_khach": len(truoc)}
