"""Nhiều máy cùng lúc: phát hiện xung đột + gộp 3 chiều (chạy: python tests/test_dong_bo.py)."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import dong_bo as DB
from gh_store import KHONG_CO, Store, git_sha, json_to_bytes


def test_store_xung_dot():
    with tempfile.TemporaryDirectory() as t:
        S = Store("local", local_dir=t)
        assert S.doc2("shared/a.json") == (None, KHONG_CO)
        ok, _, _, xd = S.luu({"shared/a.json": b"1"}, "m", {"shared/a.json": KHONG_CO})
        assert ok and not xd
        d, pb = S.doc2("shared/a.json")
        assert d == b"1" and pb == git_sha(b"1")
        # máy B lưu trước
        assert S.luu({"shared/a.json": b"2"}, "m", {"shared/a.json": pb})[0]
        # máy A (vẫn cầm phiên bản cũ) lưu -> bị chặn, báo xung đột
        ok, _, _, xd = S.luu({"shared/a.json": b"3"}, "m", {"shared/a.json": pb})
        assert not ok and xd == {"shared/a.json": git_sha(b"2")}
        assert S.doc("shared/a.json") == b"2"  # không bị ghi đè
        # không đổi gì so với bản đã đọc -> không ghi
        ok, msg, ghi, _ = S.luu({"shared/a.json": b"2"}, "m", {"shared/a.json": git_sha(b"2")})
        assert ok and ghi == []


def test_gop():
    goc = {"a": 1, "b": 2}
    out, vc = DB.tron_dict(goc, {"a": 1, "b": 3, "c": 4}, {"a": 9, "b": 2, "d": 5})
    assert out == {"a": 9, "b": 3, "c": 4, "d": 5} and vc == []
    out, vc = DB.tron_dict(goc, {"b": 2}, {"a": 1, "b": 2, "e": 1})       # mình xoá a, họ thêm e
    assert out == {"b": 2, "e": 1}
    out, vc = DB.tron_dict(goc, {"a": 5, "b": 2}, {"a": 6, "b": 2})       # cả 2 cùng sửa a
    assert out["a"] == 5 and vc == ["a"]
    s, _ = DB.tron_settings({"sua": {"x": "1"}, "map_ten": "tat"}, {"sua": {"x": "1", "y": "2"}, "map_ten": "tat"},
                            {"sua": {"x": "1", "z": "3"}, "map_ten": "cau_hinh"})
    assert s["sua"] == {"x": "1", "y": "2", "z": "3"} and s["map_ten"] == "cau_hinh"
    l, _ = DB.tron_ds_id([{"id": 1}], [{"id": 1}, {"id": 2}], [{"id": 1, "x": 1}, {"id": 3}])
    assert [x["id"] for x in l] == [1, 3, 2] and l[0] == {"id": 1, "x": 1}
    c = ["cate", "prop_id", "ma"]
    g = pd.DataFrame([["1", "10", "a"], ["1", "11", "b"]], columns=c)
    m = pd.DataFrame([["1", "10", "a"], ["1", "11", "B2"], ["1", "12", "c"]], columns=c)   # sửa 11, thêm 12
    h = pd.DataFrame([["1", "10", "a"], ["1", "11", "b"], ["2", "20", "z"]], columns=c)    # họ thêm ngành 2
    out, vc = DB.tron_bang(g, m, h, ["cate", "prop_id"])
    assert sorted(map(tuple, out.values.tolist())) == [("1", "10", "a"), ("1", "11", "B2"), ("1", "12", "c"),
                                                       ("2", "20", "z")]


def test_gop_khong_doi_mapping_cu():
    c = ["cate", "prop_id", "ma"]
    g = pd.DataFrame([["2162", "100", "deep_A"], ["2162", "100", "deep_B"]], columns=c)  # trùng có sẵn: A hiệu lực
    m = pd.concat([g, pd.DataFrame([["3", "1", "x"]], columns=c)], ignore_index=True)
    h = pd.concat([g, pd.DataFrame([["4", "1", "y"]], columns=c)], ignore_index=True)
    out, vc = DB.tron_bang(g, m, h, ["cate", "prop_id"])
    assert out.values.tolist()[0] == ["2162", "100", "deep_A"] and len(out) == 4 and vc == []
    s, _ = DB.tron_settings({"a": 1, "b": 2}, {"b": 2}, {"a": 1, "b": 2})
    assert "a" not in s                                                               # mình xoá khoá đơn -> xoá


if __name__ == "__main__":
    test_store_xung_dot()
    test_gop()
    test_gop_khong_doi_mapping_cu()
    print("OK — đồng bộ nhiều máy: 3/3 test pass")
