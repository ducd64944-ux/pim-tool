"""ĐỐI CHIẾU VỚI BẢN DESKTOP (66.py) — chạy cùng 1 bộ dữ liệu qua 2 engine, so từng ô file import.
Bản web (chế độ map_ten="tat", không quy tắc sửa) PHẢI ra file giống hệt desktop, trừ 1 lỗi desktop đã sửa có chủ đích:
  * DATA PIM có OptionValue nhưng OptionCode trống -> desktop ghi "215361, " (dư dấu phẩy); web bỏ mã rỗng + ghi log.
Chạy: python tests/test_parity_desktop.py
"""
import contextlib, io, os, sys, tempfile, types
from pathlib import Path

GOC = Path(__file__).resolve().parent
sys.path[:0] = [str(GOC.parent), str(GOC / "desktop_goc")]


def _stub_tk():
    class _M(types.ModuleType):
        def __getattr__(self, n):
            return type(n, (object,), {"__init__": lambda s, *a, **k: None})
    for n in ["tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox", "tkinter.simpledialog",
              "tkinter.font", "tkinter.scrolledtext"]:
        sys.modules.setdefault(n, _M(n))
    for n in ["ttk", "filedialog", "messagebox", "simpledialog", "font", "scrolledtext"]:
        try:
            setattr(sys.modules["tkinter"], n, sys.modules["tkinter." + n])
        except Exception:  # noqa: BLE001
            pass


IMPORT = [["model_code", "sku", "variant_code", "category_code"], ["Mã model", "Mã sản phẩm ERP", "Mã biến thể", "Mã danh mục PIM"],
          ["M1", "S1", "", "C100"], ["M1", "S2", "V2", "C100"], ["M1", "S5", "", "C100"], ["M3", "S3", "", "C100"],
          ["M4", "S4", "", "C200"], ["M6", "S6", "", "C300"], ["M7", "S7", "", "C100"]]
SP = [["PRODUCTID", "PRODUCTCODE", "PRODUCTNAME", "PROPERTYID", "PROPERTYNAME", "PROPVALUE", "CATEGORYID"],
      ["1", "S1", "a", "10", "Khối lượng", "9", "100"], ["1", "S1", "a", "11", "Màu filter", "Đen", "100"],
      ["1", "S1", "a", "11", "Màu filter", "đen ", "100"], ["1", "S1", "a", "11", "Màu filter", "Trắng.", "100"],
      ["1", "S1", "a", "11", "Màu filter", "Xanh", "100"],
      ["1", "S1", "a", "12", "Chất liệu", "Nhựa", "100"], ["1", "S1", "a", "12", "Chất liệu", "Nhựa", "100"],
      ["1", "S1", "a", "12", "Chất liệu", "Gỗ", "100"], ["1", "S1", "a", "999", "Màu sắc", "Đỏ", "100"],
      ["1", "S1", "a", "13", "Ghi chú", "Không", "100"], ["2", "S2", "b", "10", "Khối lượng", "9.5", "100"],
      ["5", "S5", "b", "10", "Khối lượng", "9|10", "100"], ["4", "S4", "d", "10", "Khối lượng", "3", "200"],
      ["6", "S6", "d", "30", "X", "abc", "300"],
      ["7", "S7", "e", "10", "Khối lượng", "12.5", ""], ["7", "S7", "e", "11", "Màu filter", "Trắng", ""]]
PIM = [["Code", "Name", "Type", "Group", "Active", "OptionCode", "OptionValue"],
       ["Mã thuộc tính", "Tên", "", "", "", "Mã option", "Giá trị"],
       ["speaker_filter", "Màu", "select", "", "1", "215361", "Đen"],
       ["speaker_filter", "Màu", "select", "", "1", "215362", "Trắng"],
       ["speaker_filter", "Màu", "select", "", "1", "215363", "Xanh"]]
MT = [["Mã ngành hàng CMS", "Tên ngành hàng CMS", "Mã thuộc tính", "Tên thuộc tính TSKT", "Mã TSKT (MASTER)", "Tên TSKT (MASTER)"],
      ["100", "Loa", "10", "Khối lượng", "product_mass", "Khối lượng"], ["100", "Loa", "12", "Chất liệu", "material_text", "Chất liệu"],
      ["100", "Loa", "13", "Ghi chú", "note_text", "Ghi chú"], ["300", "Quạt", "30", "X", "zz_text", "ZZ"]]
MF = [["MÃ NGÀNH HÀNG CMS", "TÊN NGÀNH HÀNG CMS", "MÃ THUỘC TÍNH FILTER", "TÊN THUỘC TÍNH CMS", "Mã thuộc tính mới"],
      ["100", "Loa", "11", "Màu", "speaker_filter"]]
CH = [["MÃ", "TÊN", "CỘT"],
      ["100", "Loa kéo karaoke di động chính hãng", "product_mass", "speaker_filter", "material_text", "color_text",
       "note_text", "family_code"],
      ["", "", "Khối lượng", "Màu", "Chất liệu", "Màu sắc", "Ghi chú", "Họ"],
      ["200", "Tủ lạnh", "fridge_mass"], ["", "", "Khối lượng"]]
DV = [["MÃ NGÀNH HÀNG", "TÊN NGÀNH HÀNG", "MÃ TSKT", "TÊN TSKT", "ĐƠN VỊ"], ["100", "Loa", "product_mass", "Khối lượng", "kg"]]


def _desktop(tmp: Path):
    _stub_tk()
    import pim_app_66 as D
    from openpyxl import Workbook, load_workbook
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in [("IMPORT", IMPORT), ("DATA SP", SP), ("DATA PIM", PIM), ("MAPPING TSKT MOI", MT),
                       ("MAPPING FILTER MOI", MF), ("CẤU HÌNH CATEGORY", CH), ("ĐƠN VỊ", DV)]:
        sh = wb.create_sheet(name)
        for r in rows:
            sh.append(r)
    p = tmp / "ws.xlsx"
    wb.save(p)
    with contextlib.redirect_stdout(io.StringIO()):
        D.chay_tat_ca(p, tmp / "out", xuat=False)
        x = D.xuat_file_import_nhanh(p, tmp / "out")
    out = {}
    for f in Path(x["thu_muc"]).glob("*.xlsx"):
        ws = load_workbook(f).active
        out[f.name.rsplit("_", 2)[0]] = [[c.value for c in r] for r in ws.iter_rows()]
    return out


def _web():
    import pim_core as W
    from openpyxl import load_workbook
    sp = W.pd.DataFrame([[W.chuan_hoa_id(r[0]), r[1], r[2], r[3], r[4], W.chuan_hoa_key(r[5]), r[6]] for r in SP[1:]],
                        columns=W.COT_DATA_SP)
    imp = W.pd.DataFrame([r[:4] for r in IMPORT[2:]], columns=W.COT_IMPORT)
    mt, _ = W.doc_mapping_tskt(MT)
    mf, _ = W.doc_mapping_filter(MF)
    dp, _ = W.doc_data_pim(PIM)
    r = W.chay_map(sp, imp, W.doc_cau_hinh_ngang(CH), mt, mf, W.option_maps(dp))  # mặc định = desktop
    x = W.xuat_file_import(r["bang"], imp, {}, {("100", "product_mass"): "kg"}, {})
    out = {}
    for n, d, _ in x["files"]:
        ws = load_workbook(io.BytesIO(d)).active
        out[n.rsplit("_", 2)[0]] = [[c.value for c in rr] for rr in ws.iter_rows()]
    return out, r


def test_giong_desktop():
    with tempfile.TemporaryDirectory() as t:
        d = _desktop(Path(t))
    w, r = _web()
    assert set(d) == set(w), (sorted(d), sorted(w))
    import re
    for k in d:
        # KHÁC DUY NHẤT cố ý so với desktop: ô FILTER nhiều mã viết sát dấu phẩy (25,30) vì PIM không cắt khoảng trắng
        hdr = d[k][0]
        dk = [d[k][0], d[k][1]] + [[re.sub(r"\s*,\s*", ",", c) if isinstance(c, str) and "filter" in (hdr[j] or "")
                                    and "tskt" not in (hdr[j] or "") else c for j, c in enumerate(row)] for row in d[k][2:]]
        assert dk == w[k], f"{k}\nDESKTOP {dk}\nWEB     {w[k]}"
    assert any("215361,215363" in str(row) for rows in w.values() for row in rows)
    assert not any(", " in str(c) and "filter" in (rows[0][j] or "") and "tskt" not in (rows[0][j] or "")
                   for rows in w.values() for row in rows[2:] for j, c in enumerate(row) if c)
    lg = {x[3].split(" — ")[0] for x in r["log"]}
    assert "CATEGORYID không tồn tại trong mapping" in lg and "Không có tab TSKT lẫn cấu hình" in lg


if __name__ == "__main__":
    test_giong_desktop()
    print("OK — web ra file import GIỐNG HỆT desktop 66.py")
