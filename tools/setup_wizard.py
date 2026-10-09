# -*- coding: utf-8 -*-
"""Wizard cài đặt — chạy qua 1_Khoi_tao_workspace.bat, người dùng chỉ gõ số.

Gọi từ 1_Khoi_tao_workspace.bat. Làm 4 việc:
  1. Dò các thư mục OneDrive / SharePoint đã đồng bộ trên máy
  2. Tạo các ngăn thư mục làm việc
  3. Ghi settings.json cho đúng máy này
  4. Đăng ký lịch chạy nền vào Task Scheduler của Windows
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

GOC = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(GOC))
from wr import caidat  # noqa: E402

TEN_TAC_VU = "Bao cao cong no tuan"


def do_thu_muc_dong_bo() -> list[Path]:
    nha = Path(os.path.expanduser("~"))
    ra = []
    for p in sorted(nha.iterdir()):
        if p.is_dir() and ("onedrive" in p.name.lower() or "sharepoint" in p.name.lower()):
            ra.append(p)
            for con in sorted(p.iterdir())[:20]:
                if con.is_dir() and not con.name.startswith("."):
                    ra.append(con)
    return ra


def chon_thu_muc_dich() -> Path:
    """Người dùng trỏ vào đâu cũng được — OneDrive, SharePoint, ổ D, ổ mạng."""
    ung_vien = do_thu_muc_dong_bo()
    print("\n=== Chọn THƯ MỤC ĐÍCH — nơi kế toán thả file và nhận báo cáo ===")
    print("(Thư mục chứa mã nguồn giữ nguyên, không có file dữ liệu nào ở đó)")
    for i, p in enumerate(ung_vien, start=1):
        print(f"  [{i}] {p}")
    print("  [0] Tự gõ đường dẫn bất kỳ")
    tra_loi = input(f"Chọn [0-{len(ung_vien)}]: ").strip()
    if tra_loi == "0" or not tra_loi.isdigit() or not (1 <= int(tra_loi) <= len(ung_vien)):
        goc = Path(input("Dán đường dẫn thư mục: ").strip().strip('"'))
    else:
        goc = ung_vien[int(tra_loi) - 1]
    tao_con = input(
        f'Tạo thư mục con "Cong-No-Workspace" bên trong {goc}? [C/k]: '
    ).strip().lower() not in {"k", "n", "khong", "không"}
    return goc, tao_con


def tao_file_khoi_dong() -> Path:
    """Sinh file .bat nằm cạnh mã nguồn để Task Scheduler gọi.

    Gọi thẳng python kèm đường dẫn dài có dấu cách vào Task Scheduler rất dễ
    gãy, nên gói lại thành một file .bat rồi chỉ đăng ký đúng file đó.
    """
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    trinh_chay = pythonw if pythonw.exists() else Path(sys.executable)
    f = GOC / "chay_nen.bat"
    f.write_text(
        "@echo off\r\n"
        "set PYTHONDONTWRITEBYTECODE=1\r\n"
        f'cd /d "{GOC}"\r\n'
        f'"{trinh_chay}" "{GOC / "runner.py"}"\r\n',
        encoding="utf-8",
    )
    return f


def lenh_chay_an() -> str:
    """Lệnh Task Scheduler gọi thẳng pythonw.exe (không qua .bat) để không bật cửa sổ CMD
    (Chairman 09/10: mỗi lần chạy hiện CMD nhìn mệt)."""
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    trinh_chay = pythonw if pythonw.exists() else Path(sys.executable)
    return f'"{trinh_chay}" "{GOC / "runner.py"}"'


def main() -> int:
    print("=" * 62)
    print(" CÀI ĐẶT TOOL BÁO CÁO CÔNG NỢ TUẦN")
    print("=" * 62)

    goc, tao_con = chon_thu_muc_dich()
    cfg = caidat.dat_workspace(goc, tao_con)
    print(f"\nThư mục đích: {cfg.workspace}")
    for d in cfg.cac_thu_muc():
        print(f"   {d.name}/")
    print(f"Đã ghi cấu hình riêng của máy này: {caidat.TEP_CAU_HINH.name}")
    ws = cfg.workspace

    if os.name == "nt":
        khoi_dong = tao_file_khoi_dong()      # vẫn giữ để chạy tay bằng nháy đúp
        ket_qua = subprocess.run(
            ["schtasks", "/create", "/tn", TEN_TAC_VU, "/tr", lenh_chay_an(),
             "/sc", "minute", "/mo", "15", "/f"],
            check=False, capture_output=True, text=True,
        )
        if ket_qua.returncode == 0:
            print(f'\nĐã đăng ký lịch chạy nền "{TEN_TAC_VU}", 15 phút một lần.')
            print("Xem hoặc tạm dừng: bấm phím Windows, gõ Task Scheduler.")
        else:
            print("\nKhông đăng ký được lịch chạy nền. Lý do Windows báo:")
            print((ket_qua.stderr or ket_qua.stdout).strip()[:400])
            print(f"Vẫn chạy tay được bằng cách nháy đúp: {khoi_dong}")
    else:
        print("\nKhông phải Windows nên bỏ qua bước đăng ký lịch chạy.")

    print("\nCòn 2 việc, làm một lần duy nhất:")
    print(f"  1. Bỏ file ghép khách SMS ↔ Misa (FIN đã điền) vào   {cfg.ghep}")
    print(f"  2. Bỏ bản final tuần gần nhất của kế toán trưởng vào  {cfg.ra}")
    print("     (chỉ cần lần đầu — từ tuần sau tool tự đọc báo cáo tuần trước ở đây)")
    print("  rồi nháy đúp 2_Kiem_tra_ghep_khach.bat để xem tool đã hiểu đúng chưa.")
    print(f"\nXong. Mỗi tuần chỉ cần thả 6 file vào {cfg.vao}, xuất cùng lúc:")
    print("  Misa: Tổng hợp công nợ phải thu · Chi tiết công nợ phải thu (từ 01/01/2022) ·")
    print("        Tổng hợp công nợ phải trả · Bán hàng (từ đầu năm) · Sổ chi tiết TK131 (từ đầu năm)")
    print("  SMS:  AR-AP")
    print("Muốn có ngay, không chờ lịch 15 phút: mở Task Scheduler, chọn tác vụ, bấm Run.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
