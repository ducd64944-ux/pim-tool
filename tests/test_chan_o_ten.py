"""Hàng rào: ô điền theo TÊN bị chặn khỏi file import (bo_o)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pim_core as C
import pandas as pd

b = {"title": "T", "ten_nh": "t", "attr": ["model_code", "sku", "variant_code", "a", "b"], "ten": {},
     "rows": [{"sku": "S1", "model": "M1", "variant": "", "cate_pim": "5005", "vals": {"a": "1", "b": "X"}}]}
imp = pd.DataFrame([{"model_code": "M1", "sku": "S1", "variant_code": "", "category_code": "5005"}])
x = C.xuat_file_import({"5005": b}, imp, {}, {}, {}, bo_o={"S1\tb"})
assert x["bo_o_ten"] == 1, x.get("bo_o_ten")
x2 = C.xuat_file_import({"5005": b}, imp, {}, {}, {})
assert x2.get("bo_o_ten", 0) == 0
print("OK — chặn ô theo tên: 1/1")
