# -*- coding: utf-8 -*-
"""W42 (FIN 17–19/10): tiền thu không ghi số HĐ trừ trước vào HĐ Misa đã ghi "Đã thanh toán"; chênh lệch tỷ giá
khi đối trừ đi theo hoá đơn của phiếu thu cùng ngày; cột "Đã có hợp đồng" trong Client Master."""
import datetime as dt

from wr import chitiet as ctm
from wr import ghep

D = dt.date


def _bt(ngay, so_hd, no=0.0, co=0.0, dg="", ngay_hd=None):
    return ctm.ButToan(ngay_ht=ngay, ngay_hd=ngay_hd or (ngay if no else None), so_hd=so_hd, so_ct="X",
                       dien_giai=dg, tk_doi_ung="1122", no=no, co=co)


def test_tien_khong_ghi_so_hd_tru_truoc_vao_hd_misa_da_thanh_toan():
    # STA CENTER W42: HĐ 1560 (hạch toán 22/07, ngày HĐ 04/09) Misa ghi đã TT; HĐ 1529 Misa ghi chưa TT
    k = ctm.KhachChiTiet("STA", "STA", [
        _bt(D(2026, 7, 22), "00001560", no=100, ngay_hd=D(2026, 9, 4)),
        _bt(D(2026, 8, 31), "00001529", no=30),
        _bt(D(2026, 9, 1), "", co=100, dg="Thu tiền thanh toán HD"),     # không ghi số
    ])
    con, _ = k.hoa_don_con_no(D(2026, 10, 16))
    assert [(s, t) for _, s, t in con] == [("00001560", 30.0)], "không có trạng thái Misa: trừ HĐ cũ nhất trước"
    con, _ = k.hoa_don_con_no(D(2026, 10, 16), {(1560, 2026)})
    assert [(s, t) for _, s, t in con] == [("00001529", 30.0)], "HĐ Misa đã TT được trừ trước"


def test_chenh_lech_ty_gia_di_theo_hoa_don_cua_phieu_thu_cung_ngay():
    # OPTIMALOG W42: phiếu thu 04/06 "theo HD 579" + NVK chênh lệch tỷ giá cùng ngày → còn nợ dồn về HĐ 579
    k = ctm.KhachChiTiet("OPT", "OPT", [
        _bt(D(2026, 4, 30), "00000579", no=1000),
        _bt(D(2026, 6, 4), "", co=990, dg="Thu tiền của OPTIMALOG LLC theo HD 579"),
        _bt(D(2026, 6, 4), "", no=20, dg="Xử lý chênh lệch tỷ giá khi thực hiện đối trừ chứng từ"),
    ])
    con, _ = k.hoa_don_con_no(D(2026, 10, 16))
    assert [(s, t) for _, s, t in con] == [("00000579", 30.0)]


def test_client_master_cot_da_co_hop_dong(tmp_path):
    from wr import khachhang as kh
    kh.ghi_master(tmp_path / "SAMPLE_credit_term_khach.xlsx", [
        dict(ma="A", ten="A", term="30 days", co_hd=kh.DAU_TICH, het_han=D(2026, 12, 31), gia_han=D(2027, 1, 1)),
        dict(ma="B", ten="B", term="15 days")], [])
    m = kh.doc_master(tmp_path, D(2026, 10, 20))
    assert m["A"]["co_hd"] is True and m["B"]["co_hd"] is False
    assert m["A"]["han_dung"] == D(2026, 12, 31), "ngày gia hạn chưa tới thì vẫn là ngày hết hạn"
    assert kh.doc_master(tmp_path, D(2027, 1, 5))["A"]["han_dung"] == D(2028, 1, 1)
    assert ghep.doc_credit_term(tmp_path) == {"A": "30 days", "B": "15 days"}
