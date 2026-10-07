"""Hoàn tác: ảnh chụp chênh lệch (sua/dv/rong) và đảo lại đúng."""
import sys, os, re, types
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pim_core as C
src = open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py"), encoding="utf-8").read()
i, j = src.index("def _trang_thai_sua"), src.index("def khu_hoan_tac")
class SS(dict):
    __getattr__ = dict.get
    def __setattr__(s, k, v): s[k] = v
ss = SS(sua={}, dv={}, rong={})
saved = []
ns = {"ss": ss, "C": C, "bump": lambda: None, "luu": lambda p, m: saved.append(m) or ghi_nhan_hoan_tac(m)}
exec(src[i:j], ns)
ghi_nhan_hoan_tac = ns["ghi_nhan_hoan_tac"]
ss._tt = ns["_trang_thai_sua"]()
ss.sua[("5005", "S1", "a")] = "x"; ss.dv[("5005", "a")] = "cm"
ns["luu"](["settings"], "Tách")
assert len(ss.hoan_tac) == 1 and ss.hoan_tac[0]["so"] == 2
ss.sua[("5005", "S1", "a")] = "y"
ns["luu"](["settings"], "Sửa")
ns["hoan_tac_lan_cuoi"]()
assert ss.sua[("5005", "S1", "a")] == "x" and len(ss.hoan_tac) == 1
ns["hoan_tac_lan_cuoi"]()
assert not ss.sua and not ss.dv and not ss.hoan_tac
ss.sua[("5005", "S1", "a")] = "x"; ss.sua[("5005", "S1", "b")] = "z"; ss.dv[("5005", "a")] = "cm"
ns["hoan_tac_theo_cot"]([("5005", "a")])
assert list(ss.sua) == [("5005", "S1", "b")] and not ss.dv
print("OK — hoàn tác: 4/4")
