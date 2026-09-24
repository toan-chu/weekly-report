# -*- coding: utf-8 -*-
"""So file máy sinh với bản kế toán làm tay, in bảng lệch."""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from wr.baocao import doc_bao_cao

may = doc_bao_cao(sys.argv[1]); tay = doc_bao_cao(sys.argv[2])
print(f"máy {len(may)} dòng | tay {len(tay)} dòng")
print("chỉ có ở máy:", [may[k].ten[:30] for k in may.keys()-tay.keys()])
print("chỉ có ở tay:", [tay[k].ten[:30] for k in tay.keys()-may.keys()])
gop = lambda n: n[:3]+[n[3]+n[4]+n[5]]
khop=khop_gop=tong_khop=0; lech=[]
for k in may.keys() & tay.keys():
    a,b=may[k],tay[k]
    if abs(a.tong-b.tong)>=2: lech.append(("TOTAL",b.ten,a.tong,b.tong,None,None)); continue
    sau_gop = all(abs(x-y)<2 for x,y in zip(gop(a.nhom),gop(b.nhom)))
    if all(abs(x-y)<2 for x,y in zip(a.nhom,b.nhom)): khop+=1; khop_gop+=1; tong_khop+=b.tong
    elif sau_gop: khop_gop+=1; tong_khop+=b.tong; lech.append(("TACH",b.ten,a.tong,b.tong,a.nhom,b.nhom))
    else: lech.append(("NHOM",b.ten,a.tong,b.tong,a.nhom,b.nhom))
tong=sum(v.tong for v in tay.values())
print(f"khớp cả 6 nhóm: {khop}/{len(tay)} dòng")
print(f"khớp khi gộp 61+ (so công bằng với bản cũ 4 cột): {khop_gop}/{len(tay)} dòng | giá trị {tong_khop/1e9:.2f}/{tong/1e9:.2f} tỷ = {tong_khop/tong:.0%}")
for t,ten,ta,tb,na,nb in sorted(lech,key=lambda x:-x[3]):
    if t=="TACH": print(f"  TÁCH   {ten[:34]:34} {tb:>14,.0f}  máy {[round(x/1e6,1) for x in na]}  tay {[round(x/1e6,1) for x in nb]}")
    elif t=="TOTAL": print(f"  TOTAL  {ten[:34]:34} máy {ta:>15,.0f} | tay {tb:>15,.0f}")
    else: print(f"  NHÓM   {ten[:34]:34} {tb:>14,.0f}  máy {[round(x/1e6,1) for x in na]}  tay {[round(x/1e6,1) for x in nb]}")
