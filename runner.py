# -*- coding: utf-8 -*-
"""Tool báo cáo công nợ tuần — điểm chạy.

Lịch chạy nền gọi:
    python runner.py                 quét thư mục input, thấy đủ 3 file (2 Misa + 1 SMS) thì chạy

Kiểm tra file ghép khách (nút 2_Kiem_tra_ghep_khach.bat):
    python runner.py --kiem-tra-ghep-khach

Chạy tay khi thử nghiệm:
    python runner.py --tong-hop A.xlsx --tuoi-no B.xlsx [--sms C.xlsx --ghep D.xlsx] --ra <thư mục>
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
import os
import sys
import time
import traceback
from pathlib import Path
from typing import List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wr import baocao, caidat, ghep, misa
from wr import soghi as sg
from wr import tinhtoan
from wr.misa import LoiDocFile, chuan_hoa_ten

# Mặc định ghi cạnh mã nguồn; có thư mục đích rồi thì chuyển vào đó.
NHAT_KY = Path(__file__).resolve().parent / "log" / "runner.log"
# Kế toán thả 2 file Misa trước, file SMS sau: chờ tối đa chừng này rồi mới chạy thiếu SMS
CHO_SMS_GIAY = 2 * 3600


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


def _mo_ta(f: Path) -> str:
    t = dt.datetime.fromtimestamp(f.stat().st_mtime)
    return f"{f.name} (file sửa lần cuối {t:%d/%m/%Y %H:%M})"


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


def _doi_ten(f: Path, tien_to: str) -> None:
    """Đổi tên file đầu vào. File đang bị khoá (OneDrive, Excel) thì ghi nhật ký,
    không để tiến trình chạy nền chết im lặng."""
    try:
        f.rename(f.with_name(f"{tien_to} {f.name}"))
    except OSError as loi:
        ghi_log(f"  ! Không đổi tên được {f.name} thành {tien_to}: {loi}")


def nap_ban_tay(thu_muc: Path, so: sg.SoGhi) -> int:
    """Nạp các báo cáo tuần nằm ở thư mục gốc (bản tool sinh rồi kế toán sửa, hoặc
    bản final kế toán trưởng bỏ vào) vào sổ ghi, đúng ngày chốt ghi trong file.
    Cùng một ngày chốt có nhiều bản thì bản sửa sau cùng thắng. Bản của người
    thắng bản của máy. File đã nạp (cùng giờ sửa) thì bỏ qua."""
    n = 0
    if not thu_muc or not Path(thu_muc).is_dir():
        return 0
    for f in sorted(Path(thu_muc).glob("*.xlsx"), key=lambda f: f.stat().st_mtime):
        if f.name.startswith("~$"):
            continue
        dau = f"{f.stat().st_mtime:.0f}"
        if so.da_nap.get(f.name) == dau:
            continue
        try:
            ngay = baocao.ngay_chot_trong_file(f)
            if ngay is None:
                ghi_log(f"  ! {f.name}: không thấy ngày chốt (dòng 'Reporting period dd/mm/yyyy - "
                        "dd/mm/yyyy' hoặc 'At dd Mon yyyy'), chưa nạp")
                continue
            k = baocao.nap_vao_so_ghi(so, f, ngay)
        except Exception as loi:
            ghi_log(f"  ! Không đọc được bản final {f.name}: {loi}")
            continue
        so.da_nap[f.name] = dau
        n += 1
        ghi_log(f"Đã đọc lại báo cáo {f.name}: {k} khách, chốt {ngay:%d/%m/%Y} — tuần sau tính tiếp từ bản này")
    return n


def chay_mot_lan(
    duong_tong_hop: Path,
    duong_tuoi_no: Path,
    thu_muc_ra: Path,
    duong_so_ghi: Path,
    mau_chuan: Path,
    duong_sms: Optional[Path] = None,
    thu_muc_ghep: Optional[Path] = None,
    duong_ghep: Optional[Path] = None,
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

    them: dict = {}
    nguon_them: list = []
    if duong_sms:
        dong_sms = ghep.doc_sms(duong_sms)
        f_ghep = Path(duong_ghep) if duong_ghep else ghep.tim_file_ghep(thu_muc_ghep)
        bang = ghep.doc_bang_ghep(f_ghep) if f_ghep else None
        if bang is None:
            ghi_log("  ! Chưa có file ghép khách trong thư mục sample — chỉ ghép được khách trùng tên/mã")
        else:
            for ma_sms, cu, moi in bang.xung_dot:
                ghi_log(f"  ! Mã SMS {ma_sms} được ghép với cả {cu} và {moi}: theo phía Misa ({moi})")
        kq = tinhtoan.tinh_bang_sms(th, tn, so, dong_sms, bang, ngay_chot)
        dong, da_bo = kq.dong, kq.da_bo
        them = {"che_do": "sms", "sms_cat": kq.sms_cat, "sms_khong_misa": kq.sms_khong_misa,
                "sms_chua_ghep": kq.sms_chua_ghep}
        nguon_them = [("", _mo_ta(Path(duong_sms))),
                      ("", f"Bảng ghép khách: {f_ghep.name}" if f_ghep else "Bảng ghép khách: (chưa có)")]
    else:
        dong = tinhtoan.tinh_bang(th, tn, so, ngay_chot)
        da_bo = getattr(tinhtoan.tinh_bang, "da_bo", [])
    dich = thu_muc_ra / baocao.ten_file_bao_cao(th.tu_ngay, th.den_ngay)
    if dich.exists():
        # Báo cáo cũ của tuần này có thể đã được kế toán sửa tay: không ghi đè (RULES 5.1)
        dich = dich.with_name(f"{dich.stem} (chay lai {dt.datetime.now():%d-%m %Hh%M}){dich.suffix}")
    khach_lech = []
    for k, v in tn.dong.items():
        t = th.dong.get(k)
        so_tong_hop = t.du_no if t else 0.0
        if abs(v.tong - so_tong_hop) >= 2:
            khach_lech.append((v.ten, so_tong_hop, v.tong))
    khach_lech.sort(key=lambda x: -abs(x[2] - x[1]))

    baocao.ghi_bao_cao(
        dong, ngay_chot, mau_chuan, dich, tong_misa=th.tong_du_no,
        doi_chieu={
            # Chế độ SMS không dùng file Tuổi nợ để chia tuổi, nên không liệt kê độ vênh của nó
            "tong_tuoi_no": None if duong_sms else sum(v.tong for v in tn.dong.values()),
            "khach_lech": [] if duong_sms else khach_lech[:30],
            "da_bo": da_bo,
            **them,
            "nguon": [("Nguồn dữ liệu", _mo_ta(Path(duong_tong_hop))),
                      ("", _mo_ta(Path(duong_tuoi_no))), *nguon_them,
                      ("", "Số trong hai file này thay đổi theo thời điểm xuất. "
                           "Chỉ so báo cáo với bản làm tay khi cả hai dùng cùng một lần xuất Misa.")],
        },
    )

    trang_cu = so.trang.get(ngay_chot.isoformat()) or {}
    if any(v.nguon == "ban-tay" for v in trang_cu.values()):
        ghi_log(f"  Sổ ghi ngày {ngay_chot:%d/%m/%Y} đã lấy từ báo cáo trong thư mục gốc, giữ nguyên")
    else:
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

    so = sg.SoGhi(cfg.so_ghi)
    if nap_ban_tay(cfg.ban_tay, so):
        so.ghi()

    ds_sms = [f for f in files if ghep.la_file_sms(f)]
    cap, la = _ghep_cap([f for f in files if f not in ds_sms])
    for f, vi_sao in la:
        ghi_log(f"  Bỏ qua {f.name}: {vi_sao}")
    if not cap:
        ghi_log(f"Chưa có cặp file Misa nào đủ để chạy. Đang có: {[f.name for f in files]}")
        return 0
    sms = max(ds_sms, key=lambda f: f.stat().st_mtime) if ds_sms else None
    if sms is None:
        moi_nhat = max(f.stat().st_mtime for _, a, b in cap for f in (a, b))
        if time.time() - moi_nhat < CHO_SMS_GIAY:
            ghi_log("Đã có 2 file Misa, đang chờ file SMS (AR-AP). Tuần này không có SMS thì sau "
                    f"{CHO_SMS_GIAY // 3600} giờ tool tự chạy theo luật cũ.")
            return 0
        ghi_log("  ! Không có file SMS — chạy theo luật cũ (file Tuổi nợ + sổ ghi)")
    if len(cap) > 1:
        ghi_log(f"Có {len(cap)} kỳ trong thư mục, chạy lần lượt từ kỳ cũ nhất"
                + (". File SMS chỉ dùng cho kỳ mới nhất" if sms else ""))

    xong = 0
    for i, (ngay, duong_th, duong_tn) in enumerate(cap):
        sms_ky = sms if i == len(cap) - 1 else None
        ghi_log(f"Bắt đầu kỳ đến {ngay:%d/%m/%Y}: {duong_th.name} + {duong_tn.name}"
                + (f" + {sms_ky.name}" if sms_ky else ""))
        try:
            chay_mot_lan(duong_th, duong_tn, cfg.ra, cfg.so_ghi, cfg.mau_chuan,
                         duong_sms=sms_ky, thu_muc_ghep=cfg.ghep)
        except PermissionError as loi:
            ghi_log(f"  Tạm hoãn kỳ này: {loi}. Nhiều khả năng file kết quả đang "
                    "mở trong Excel. Đóng file đi, lần chạy sau tool làm tiếp.")
            continue
        except Exception as loi:  # không bao giờ để tiến trình chết im lặng
            ghi_log(f"  LỖI: {loi}")
            for f in (duong_th, duong_tn):
                _doi_ten(f, "[LOI]")
            (cfg.vao / f"[LOI] {dt.date.today():%Y%m%d} - doc-vi-sao-hong.txt").write_text(
                f"{dt.datetime.now():%d/%m/%Y %H:%M}\n\n{loi}\n\n"
                "Sửa xong thì bỏ chữ [LOI] khỏi tên file, tool sẽ chạy lại ở lần kế tiếp.\n\n"
                + traceback.format_exc(),
                encoding="utf-8",
            )
            continue
        for f in (duong_th, duong_tn) + ((sms_ky,) if sms_ky else ()):
            _doi_ten(f, "[DONE]")
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
    p.add_argument("--sms", help="file SMS AR-AP (chạy tay)")
    p.add_argument("--ghep", help="file ghép khách SMS ↔ Misa (chạy tay)")
    p.add_argument("--kiem-tra-ghep-khach", action="store_true",
                   help="đối chiếu file ghép khách với Misa/SMS mới nhất, mở file kiểm tra")
    p.add_argument("--khong-mo-file", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--mo-ket-qua", action="store_true",
                   help="chạy xong thì mở thư mục gốc chứa báo cáo (nút 3_Chay_ngay.bat)")
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

    if a.kiem_tra_ghep_khach:
        if cfg is None:
            ghi_log("Chưa đặt thư mục đích. Chạy 1_Khoi_tao_workspace.bat trước.")
            return 2
        from wr import kiemtra
        cfg.tao_cac_thu_muc()
        so = sg.SoGhi(cfg.so_ghi)
        if nap_ban_tay(cfg.ban_tay, so):
            so.ghi()
        try:
            f = kiemtra.kiem_tra(cfg, so, ghi_log)
        except PermissionError as loi:
            ghi_log(f"Không ghi được file kiểm tra: {loi}. Đóng file KIEM_TRA đang mở rồi bấm lại.")
            return 1
        if f is None:
            return 2
        if os.name == "nt" and not a.khong_mo_file:
            try:
                os.startfile(str(f))       # mở bằng Excel cho FIN xem ngay
            except OSError:
                pass
        return 0

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
        ghi_log("Chưa đặt thư mục đích. Chạy 1_Khoi_tao_workspace.bat, hoặc: "
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
            chay_mot_lan(Path(a.tong_hop), Path(a.tuoi_no), Path(a.ra), duong_so_ghi, mau_chuan,
                         duong_sms=Path(a.sms) if a.sms else None,
                         duong_ghep=Path(a.ghep) if a.ghep else None,
                         thu_muc_ghep=cfg.ghep if cfg else None)
        except LoiDocFile as loi:
            ghi_log(f"LỖI: {loi}")
            return 2
        return 0

    try:
        xong = quet_thu_muc(cfg)
    except Exception as loi:      # chạy nền: không bao giờ chết im lặng
        ghi_log(f"LỖI khi quét thư mục: {loi}\n{traceback.format_exc()}")
        return 1
    if a.mo_ket_qua:
        if not xong:
            _in("Không có báo cáo mới. Kiểm tra thư mục input đã đủ 3 file (2 Misa + 1 SMS) chưa.")
        if os.name == "nt":
            try:
                os.startfile(str(cfg.ra))
            except OSError:
                pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
