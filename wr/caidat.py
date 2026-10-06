# -*- coding: utf-8 -*-
"""Cấu hình theo từng máy. Không đi theo git, vì đường dẫn mỗi máy mỗi khác."""
from __future__ import annotations

import json
import os
import platform
import re
import socket
from pathlib import Path
from typing import Optional

GOC = Path(__file__).resolve().parents[1]
TEP_CAU_HINH = GOC / "settings.json"

class ChuaDatThuMuc(Exception):
    """settings.json chưa có thư mục đích — tool không biết đọc và ghi ở đâu."""


MAC_DINH = {
    "workspace": "",                 # thư mục làm việc trên OneDrive
    # Bố cục Chairman chốt 2026-10-06: gốc chứa báo cáo tuần, chỉ 2 ngăn con cho người
    #   <gốc>/          báo cáo tuần tool sinh ra; tuần sau tool đọc lại bản mới nhất ở đây
    #   <gốc>/input/    3 file mỗi tuần (2 Misa + 1 SMS)
    #   <gốc>/sample/   file ghép khách SMS ↔ Misa và file KIEM_TRA
    #   <gốc>/_tool/    trí nhớ và nhật ký của tool (ẩn trên Windows), người dùng không đụng
    "thu_muc_vao": "input",
    "thu_muc_ra": "",
    "thu_muc_ban_tay": "",
    "thu_muc_ghep": "sample",
    "thu_muc_so_ghi": "_tool",
    "thu_muc_luu_tru": "_tool/luu-tru",
    "thu_muc_log": "_tool",
    "credit_term_mac_dinh": 30,
}


DUONG_DAN_WINDOWS = re.compile(r"^[A-Za-z]:[\\/]")


class CauHinh:
    def __init__(self, raw: dict):
        self.raw = {**MAC_DINH, **raw}
        ws = _workspace_cua_may(self.raw)
        if ws and os.name != "nt" and DUONG_DAN_WINDOWS.match(ws):
            raise ChuaDatThuMuc(
                f"Cấu hình đang trỏ tới một đường dẫn Windows ({ws}) nhưng máy này "
                "không phải Windows. Đây là cấu hình của máy khác — đặt lại bằng "
                '--dat-thu-muc "<đường dẫn trên máy này>".'
            )
        if not ws:
            raise ChuaDatThuMuc(
                f"Máy {ten_may()} chưa đặt thư mục đích. Chạy 1_Khoi_tao_workspace.bat một lần, hoặc: "
                'python runner.py --dat-thu-muc "<đường dẫn>"'
            )
        self.workspace = Path(ws)

    def _d(self, khoa: str) -> Path:
        """Chỉ trả về đường dẫn. Tạo thư mục là việc của tao_cac_thu_muc(),
        để lệnh chỉ đọc (VD --xem-cau-hinh) không đẻ ra thư mục lạ."""
        return self.workspace / self.raw[khoa]

    @property
    def vao(self) -> Path: return self._d("thu_muc_vao")
    @property
    def ra(self) -> Path: return self._d("thu_muc_ra")
    @property
    def ban_tay(self) -> Path: return self._d("thu_muc_ban_tay")
    @property
    def ghep(self) -> Path: return self._d("thu_muc_ghep")
    @property
    def so_ghi(self) -> Path: return self._d("thu_muc_so_ghi") / "so-ghi.json"
    @property
    def luu_tru(self) -> Path: return self._d("thu_muc_luu_tru")
    @property
    def log(self) -> Path: return self._d("thu_muc_log") / "runner.log"
    @property
    def mau_chuan(self) -> Path: return GOC / "template" / "Bang_cong_no_MAU_CHUAN.xlsx"

    def cac_thu_muc(self):
        ra = []
        for d in [self.ra, self.vao, self.ghep, self.ban_tay, self.so_ghi.parent, self.log.parent,
                  self.luu_tru]:
            if d not in ra:
                ra.append(d)
        return ra

    def tao_cac_thu_muc(self):
        for d in self.cac_thu_muc():
            d.mkdir(parents=True, exist_ok=True)
        an = self.workspace / "_tool"
        if os.name == "nt" and an.is_dir():
            try:   # ẩn ngăn của tool để FIN chỉ thấy input/ và sample/
                import subprocess
                subprocess.run(["attrib", "+h", str(an)], check=False, capture_output=True)
            except OSError:
                pass
        return self.cac_thu_muc()


def ten_may() -> str:
    """Thư mục mã nguồn có thể nằm trên OneDrive và đồng bộ sang máy khác, nên
    mỗi máy giữ đường dẫn riêng của mình trong cùng một file cấu hình."""
    return (platform.node() or socket.gethostname() or "may-khong-ten").strip()


def _doc_raw() -> dict:
    if not TEP_CAU_HINH.exists():
        return {}
    try:
        return json.loads(TEP_CAU_HINH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _workspace_cua_may(raw: dict) -> str:
    theo_may = raw.get("workspace_theo_may") or {}
    return theo_may.get(ten_may()) or raw.get("workspace") or ""


def dat_workspace(duong_dan, tao_thu_muc_con: bool = True) -> "CauHinh":
    """Đặt thư mục đích và tạo sẵn các ngăn bên trong.

    Người dùng trỏ vào đâu cũng được: OneDrive, SharePoint đã đồng bộ, ổ D,
    ổ mạng. Thư mục chứa mã nguồn không đụng tới.
    """
    goc = Path(str(duong_dan).strip().strip('"')).expanduser()
    if tao_thu_muc_con:
        goc = goc / "Cong-No-Workspace"
    goc.mkdir(parents=True, exist_ok=True)
    raw = _doc_raw()
    raw.setdefault("workspace_theo_may", {})[ten_may()] = str(goc)
    raw.pop("workspace", None)          # khoá cũ, không dùng nữa
    TEP_CAU_HINH.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    cfg = CauHinh(raw)
    cfg.tao_cac_thu_muc()
    return cfg


def doc(duong_dan: Optional[Path] = None) -> CauHinh:
    if duong_dan is None:
        return CauHinh(_doc_raw())
    p = Path(duong_dan)
    raw = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    return CauHinh(raw)
