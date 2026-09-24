# -*- coding: utf-8 -*-
"""Wizard cài đặt — chạy qua setup.bat, người dùng chỉ gõ số.

Làm 4 việc:
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
        khoi_dong = tao_file_khoi_dong()
        ket_qua = subprocess.run(
            ["schtasks", "/create", "/tn", TEN_TAC_VU, "/tr", f'"{khoi_dong}"',
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

    print("\nCòn một việc nữa, làm một lần duy nhất: nạp trạng thái ban đầu từ")
    print("báo cáo tuần gần nhất kế toán đã làm tay, ví dụ")
    print('   python runner.py --nap-bao-cao "...W38...xlsx" --ngay-chot 2026-09-19')
    print("\nXong. Từ giờ kế toán chỉ cần thả 2 file Misa vào thư mục 01_Input.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
