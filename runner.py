# -*- coding: utf-8 -*-
"""Tool báo cáo công nợ tuần — điểm chạy.

Lịch chạy nền gọi:
    python runner.py                 quét thư mục 01_Input, thấy đủ 2 file thì chạy

Chạy tay khi thử nghiệm:
    python runner.py --tong-hop A.xlsx --tuoi-no B.xlsx --ra <thư mục>
    python runner.py --nap-bao-cao "2026_W36_....xlsx" --ngay-chot 2026-09-05
    python runner.py --ngoai-le "TEN KHACH" --nhom "1 - 30" --vi-sao "lô tàu chìm"
    python runner.py --credit-term "TEN KHACH" --gia-tri "15 days"
    python runner.py --bo-qua "TEN KHACH" --vi-sao "no ao do loi xuat hoa don"
    python runner.py --nhan-lai "TEN KHACH"

Đặt hoặc đổi thư mục đích (nơi kế toán thả file và nhận kết quả):
    python runner.py --dat-thu-muc "<đường dẫn thư mục đích>"
    python runner.py --xem-cau-hinh
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
import traceback
from pathlib import Path
from typing import List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wr import baocao, caidat, misa
from wr import soghi as sg
from wr import tinhtoan
from wr.misa import LoiDocFile, chuan_hoa_ten

# Mặc định ghi cạnh mã nguồn; có thư mục đích rồi thì chuyển vào đó.
NHAT_KY = Path(__file__).resolve().parent / "log" / "runner.log"


def dat_nhat_ky(duong_dan: Path) -> None:
    global NHAT_KY
    NHAT_KY = Path(duong_dan)


def _in(dong: str) -> None:
    """Chạy nền bằng pythonw thì không có màn hình, và console Windows có thể
    không hiện được tiếng Việt. Cả hai đều không được phép làm chết tiến trình."""
    try:
        print(dong)
    except (AttributeError, ValueError, UnicodeEncodeError, OSError):
        try:
            sys.stdout.buffer.write((dong + "\n").encode("utf-8", "replace"))
        except Exception:
            pass


def ghi_log(dong: str) -> None:
    try:
        NHAT_KY.parent.mkdir(parents=True, exist_ok=True)
        with NHAT_KY.open("a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {dong}\n")
    except OSError:
        pass          # không ghi được nhật ký cũng không được chặn việc chính
    _in(dong)


def _ghep_cap(files: List[Path]):
    """Ghép từng cặp file CÙNG KỲ, trả về danh sách cặp xếp theo thứ tự thời gian.

    Thả nhiều tuần vào một lúc cũng chạy được: mỗi tuần một cặp, chạy lần lượt
    từ tuần cũ nhất, để sổ ghi nối đúng thứ tự.
    """
    ds_th, ds_tn, la = [], [], []
    for f in files:
        loai = misa.nhan_dang(f)
        try:
            if loai == "tong-hop":
                ds_th.append((misa.doc_tong_hop(f).den_ngay, f))
            elif loai == "tuoi-no":
                ds_tn.append((misa.doc_tuoi_no(f).den_ngay, f))
            else:
                la.append((f, "không nhận ra là file Misa nào"))
        except Exception as loi:
            la.append((f, str(loi)))

    cap, con_tn = [], dict(ds_tn)
    for ngay_th, f_th in sorted(ds_th):
        mong_doi = ngay_th + dt.timedelta(days=1)
        f_tn = next((f for n, f in sorted(ds_tn) if n == mong_doi and f in con_tn.values()), None)
        if f_tn is None:
            la.append((f_th, f"chưa có file Tuổi nợ chốt ngày {mong_doi:%d/%m/%Y} đi kèm"))
            continue
        con_tn = {n: f for n, f in con_tn.items() if f != f_tn}
        cap.append((ngay_th, f_th, f_tn))
    for _, f in sorted(con_tn.items()):
        la.append((f, "chưa có file Tổng hợp cùng kỳ đi kèm"))
    return cap, la


def chay_mot_lan(
    duong_tong_hop: Path,
    duong_tuoi_no: Path,
    thu_muc_ra: Path,
    duong_so_ghi: Path,
    mau_chuan: Path,
) -> Path:
    th = misa.doc_tong_hop(duong_tong_hop)
    tn = misa.doc_tuoi_no(duong_tuoi_no)
    misa.kiem_cung_ky(th, tn)

    so = sg.SoGhi(duong_so_ghi)
    ngay_chot = tn.den_ngay
    truoc = so.ngay_trang_truoc(ngay_chot)
    if truoc is None:
        ghi_log("  ! Sổ ghi chưa có trang nào trước ngày chốt — chạy chế độ khởi động lạnh")
    elif (ngay_chot - truoc).days > 10:
        ghi_log(f"  ! Trang gần nhất của sổ ghi là {truoc:%d/%m/%Y}, cách {(ngay_chot - truoc).days} ngày. "
                "Nợ mới phát sinh trong lúc nghỉ sẽ bị đoán trẻ hơn thực tế")

    dong = tinhtoan.tinh_bang(th, tn, so, ngay_chot)
    dich = thu_muc_ra / baocao.ten_file_bao_cao(th.tu_ngay, th.den_ngay)
    khach_lech = []
    for k, v in tn.dong.items():
        t = th.dong.get(k)
        so_tong_hop = t.du_no if t else 0.0
        if abs(v.tong - so_tong_hop) >= 2:
            khach_lech.append((v.ten, so_tong_hop, v.tong))
    khach_lech.sort(key=lambda x: -abs(x[2] - x[1]))
    def _mo_ta(f: Path) -> str:
        t = dt.datetime.fromtimestamp(f.stat().st_mtime)
        return f"{f.name} (file sửa lần cuối {t:%d/%m/%Y %H:%M})"

    baocao.ghi_bao_cao(
        dong, ngay_chot, mau_chuan, dich, tong_misa=th.tong_du_no,
        doi_chieu={
            "tong_tuoi_no": sum(v.tong for v in tn.dong.values()),
            "khach_lech": khach_lech[:30],
            "da_bo": getattr(tinhtoan.tinh_bang, "da_bo", []),
            "nguon": [("Nguồn dữ liệu", _mo_ta(Path(duong_tong_hop))),
                      ("", _mo_ta(Path(duong_tuoi_no))),
                      ("", "Số trong hai file này thay đổi theo thời điểm xuất. "
                           "Chỉ so báo cáo với bản làm tay khi cả hai dùng cùng một lần xuất Misa.")],
        },
    )

    so.ghi_trang(ngay_chot, {d.ten_chuan: tinhtoan.thanh_ho_so(d) for d in dong})
    so.ghi()

    can_xem = sum(1 for d in dong if d.cho_xem_lai)
    ghi_log(f"  -> {dich.name}: {len(dong)} khách, tổng {th.tong_du_no:,.0f}, "
            f"{can_xem} dòng cần xem lại")
    return dich


def quet_thu_muc(cfg: caidat.CauHinh) -> int:
    cfg.tao_cac_thu_muc()
    files = [f for f in sorted(cfg.vao.glob("*.xlsx"))
             if not f.name.startswith(("[DONE]", "[LOI]", "~$"))]
    if not files:
        return 0

    cap, la = _ghep_cap(files)
    for f, vi_sao in la:
        ghi_log(f"  Bỏ qua {f.name}: {vi_sao}")
    if not cap:
        ghi_log(f"Chưa có cặp file nào đủ để chạy. Đang có: {[f.name for f in files]}")
        return 0
    if len(cap) > 1:
        ghi_log(f"Có {len(cap)} kỳ trong thư mục, chạy lần lượt từ kỳ cũ nhất")

    xong = 0
    for ngay, duong_th, duong_tn in cap:
        ghi_log(f"Bắt đầu kỳ đến {ngay:%d/%m/%Y}: {duong_th.name} + {duong_tn.name}")
        try:
            chay_mot_lan(duong_th, duong_tn, cfg.ra, cfg.so_ghi, cfg.mau_chuan)
        except PermissionError as loi:
            ghi_log(f"  Tạm hoãn kỳ này: {loi}. Nhiều khả năng file kết quả đang "
                    "mở trong Excel. Đóng file đi, lần chạy sau tool làm tiếp.")
            continue
        except Exception as loi:  # không bao giờ để tiến trình chết im lặng
            ghi_log(f"  LỖI: {loi}")
            for f in (duong_th, duong_tn):
                f.rename(f.with_name("[LOI] " + f.name))
            (cfg.vao / f"[LOI] {dt.date.today():%Y%m%d} - doc-vi-sao-hong.txt").write_text(
                f"{dt.datetime.now():%d/%m/%Y %H:%M}\n\n{loi}\n\n"
                "Sửa xong thì bỏ chữ [LOI] khỏi tên file, tool sẽ chạy lại ở lần kế tiếp.\n\n"
                + traceback.format_exc(),
                encoding="utf-8",
            )
            continue
        for f in (duong_th, duong_tn):
            f.rename(f.with_name("[DONE] " + f.name))
        xong += 1
    return xong


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Sinh báo cáo công nợ tuần từ 2 file Misa")
    p.add_argument("--tong-hop"); p.add_argument("--tuoi-no"); p.add_argument("--ra")
    p.add_argument("--so-ghi"); p.add_argument("--mau-chuan")
    p.add_argument("--nap-bao-cao"); p.add_argument("--ngay-chot")
    p.add_argument("--ngoai-le"); p.add_argument("--nhom"); p.add_argument("--vi-sao", default="")
    p.add_argument("--credit-term"); p.add_argument("--gia-tri")
    p.add_argument("--bo-qua"); p.add_argument("--nhan-lai")
    p.add_argument("--dat-thu-muc", help="thư mục đích trên máy này")
    p.add_argument("--khong-tao-thu-muc-con", action="store_true",
                   help="dùng thẳng thư mục đã trỏ, không tạo Cong-No-Workspace bên trong")
    p.add_argument("--xem-cau-hinh", action="store_true")
    a = p.parse_args(argv)

    if a.dat_thu_muc:
        cfg = caidat.dat_workspace(a.dat_thu_muc, not a.khong_tao_thu_muc_con)
        dat_nhat_ky(cfg.log)
        ghi_log(f"Thư mục đích: {cfg.workspace}")
        for d in cfg.cac_thu_muc():
            _in(f"    {d}")
        return 0

    if a.xem_cau_hinh:
        try:
            cfg = caidat.doc()
        except caidat.ChuaDatThuMuc as loi:
            _in(str(loi))
            return 2
        _in(f"Máy: {caidat.ten_may()}")
        _in(f"Thư mục mã nguồn: {caidat.GOC}")
        _in(f"Thư mục đích: {cfg.workspace}")
        for d in cfg.cac_thu_muc():
            _in(f"    {d}")
        _in(f"Mẫu chuẩn: {cfg.mau_chuan}")
        return 0

    # Luôn dùng cấu hình của máy nếu đã đặt thư mục đích, kể cả khi chạy tay,
    # để sổ ghi và nhật ký không bao giờ rơi vào thư mục mã nguồn.
    cfg = None
    try:
        if caidat.TEP_CAU_HINH.exists():
            cfg = caidat.doc()
    except caidat.ChuaDatThuMuc:
        cfg = None
    if cfg:
        dat_nhat_ky(cfg.log)
    duong_so_ghi = Path(a.so_ghi) if a.so_ghi else (cfg.so_ghi if cfg else Path("so-ghi.json"))
    mau_chuan = Path(a.mau_chuan) if a.mau_chuan else (
        cfg.mau_chuan if cfg else caidat.GOC / "template" / "Bang_cong_no_MAU_CHUAN.xlsx")

    if a.nap_bao_cao:
        if not a.ngay_chot:
            p.error("--nap-bao-cao cần kèm --ngay-chot YYYY-MM-DD")
        so = sg.SoGhi(duong_so_ghi)
        n = baocao.nap_vao_so_ghi(so, Path(a.nap_bao_cao), dt.date.fromisoformat(a.ngay_chot))
        so.ghi()
        ghi_log(f"Đã nạp {n} khách từ {Path(a.nap_bao_cao).name} vào sổ ghi, chốt {a.ngay_chot}")
        return 0

    if a.ngoai_le:
        if a.nhom not in sg.TEN_NHOM:
            p.error(f"--nhom phải là một trong {sg.TEN_NHOM}")
        so = sg.SoGhi(duong_so_ghi)
        so.ngoai_le[chuan_hoa_ten(a.ngoai_le)] = {"nhom": a.nhom, "ly_do": a.vi_sao}
        so.ghi()
        ghi_log(f"Đã khai ngoại lệ cho {a.ngoai_le}: xếp vào nhóm {a.nhom}")
        return 0

    if a.credit_term:
        if not a.gia_tri:
            p.error("--credit-term cần kèm --gia-tri, ví dụ --gia-tri \"15 days\"")
        so = sg.SoGhi(duong_so_ghi)
        so.dieu_chinh.setdefault(chuan_hoa_ten(a.credit_term), {})["credit_term"] = a.gia_tri
        so.ghi()
        ghi_log(f"Đã khai credit term cho {a.credit_term}: {a.gia_tri}")
        return 0

    if cfg is None and not any([a.tong_hop, a.nap_bao_cao, a.ngoai_le, a.credit_term,
                                a.bo_qua, a.nhan_lai]):
        ghi_log("Chưa đặt thư mục đích. Chạy setup.bat, hoặc: "
                'python runner.py --dat-thu-muc "<đường dẫn>"')
        return 2

    if a.bo_qua or a.nhan_lai:
        so = sg.SoGhi(duong_so_ghi)
        if a.bo_qua:
            so.bo_qua[chuan_hoa_ten(a.bo_qua)] = {"ly_do": a.vi_sao or "kế toán khai loại khỏi bảng"}
            ghi_log(f"Đã loại {a.bo_qua} khỏi báo cáo: {a.vi_sao or '(không ghi lý do)'}")
        else:
            so.bo_qua.pop(chuan_hoa_ten(a.nhan_lai), None)
            ghi_log(f"Đã nhận lại {a.nhan_lai} vào báo cáo")
        so.ghi()
        return 0

    if a.tong_hop or a.tuoi_no:
        if not (a.tong_hop and a.tuoi_no and a.ra):
            p.error("chạy tay cần đủ --tong-hop, --tuoi-no và --ra")
        try:
            chay_mot_lan(Path(a.tong_hop), Path(a.tuoi_no), Path(a.ra), duong_so_ghi, mau_chuan)
        except LoiDocFile as loi:
            ghi_log(f"LỖI: {loi}")
            return 2
        return 0

    return 0 if quet_thu_muc(cfg) >= 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
