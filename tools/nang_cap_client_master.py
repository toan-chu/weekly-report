"""Nâng cấp file Client Master đang dùng lên mẫu mới nhất (VD thêm cột "Đã có hợp đồng"). Giữ nguyên dữ liệu,
đổi ngày gõ dạng chữ thành ngày thật. Dùng: python tools/nang_cap_client_master.py "<file cũ>" ["<file mới>"]
Không ghi file mới thì ghi cạnh file cũ, thêm đuôi _moi."""
import sys, datetime as dt, openpyxl
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from wr import khachhang as kh, ghep
src = sys.argv[1]
dst = sys.argv[2] if len(sys.argv) > 2 else str(Path(src).with_name(Path(src).stem + '_moi.xlsx'))
wb = openpyxl.load_workbook(src, data_only=False)
ws = wb[kh.SHEET_MASTER]
tieu = [str(ws.cell(kh.DONG_TIEU_DE, c).value or "").strip() for c in range(1, ws.max_column + 1)]
khoa = {t: k for k, t, *_ in kh.COT}
def gia(v, loai):
    if isinstance(v, str):
        v = v.replace("\xa0", " ").strip()
        if not v: return None
    if loai == "ngay" and v is not None and not isinstance(v, (dt.date, dt.datetime)):
        d = ghep._ngay(v); return d or v
    if isinstance(v, dt.datetime): return v.date() if loai == "ngay" else v
    return v
loai = {k: l for k, _, _, _, l in kh.COT}
dong = []
for r in range(kh.DONG_DAU, ws.max_row + 1):
    d = {}
    for c, t in enumerate(tieu, start=1):
        k = khoa.get(t)
        if not k or loai[k] == "cong_thuc": continue
        d[k] = gia(ws.cell(r, c).value, loai[k])
    if not d.get("ten") and not d.get("ma"): continue
    if "co_hd" not in d or d["co_hd"] is None:
        d["co_hd"] = kh.DAU_TICH if d.get("so_hd") else None
    dong.append(d)
hd = wb["Hướng dẫn"]
sales = [str(hd.cell(r, 2).value).strip() for r in range(33, 53) if hd.cell(r, 2).value and str(hd.cell(r, 2).value).strip()]
kh.ghi_master(dst, dong, sales)
print("Đã ghi", dst); print(len(dong), "khách;", sum(1 for d in dong if d.get("co_hd")), "đã ✓;", sum(1 for d in dong if d.get("term")), "có term;",
      sum(1 for d in dong if isinstance(d.get("het_han"), str) or isinstance(d.get("hieu_luc"), str)), "ngày còn dạng chữ")
