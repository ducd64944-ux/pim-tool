"""Kiểm thử nhanh engine bằng dữ liệu giả (chạy: python -m pytest tests -q  hoặc  python tests/test_core.py)."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import pim_core as C


def _du_lieu():
    sp = pd.DataFrame([
        ["1", "SKU1", "Loa A", "100", "Sâu", "6.95", "2162"],
        ["1", "SKU1", "Loa A", "101", "Tiện ích", "Chống nước", "2162"],
        ["1", "SKU1", "Loa A", "101", "Tiện ích", "Có đèn LED", "2162"],
        ["1", "SKU1", "Loa A", "102", "Sản xuất tại", "Đang cập nhật", "2162"],
        ["1", "SKU1", "Loa A", "999", "Tổng công suất", "35 W", "2162"],
    ], columns=C.COT_DATA_SP)
    imp = pd.DataFrame([["M1", "SKU1", "", "LOA"]], columns=C.COT_IMPORT)
    cfg = {"2162": {"ten": "Loa", "cot": ["deep_tskt_master", "utilities_filter_master", "made_in_tskt_master",
                                          "total_power_tskt_master"],
                    "ten_cot": {"deep_tskt_master": "Sâu", "total_power_tskt_master": "Tổng công suất"}}}
    mt = pd.DataFrame([["2162", "Loa", "100", "Sâu", "deep_tskt_master", "Sâu"],
                       ["2162", "Loa", "102", "Sản xuất tại", "made_in_tskt_master", ""]], columns=C.COT_MAP_TSKT)
    mf = pd.DataFrame([["2162", "Loa", "101", "Tiện ích", "utilities_filter_master"]], columns=C.COT_MAP_FILTER)
    dp = pd.DataFrame([["utilities_filter_master", "Tiện ích", "", "", "TRUE", "13", "Chống nước"],
                       ["utilities_filter_master", "Tiện ích", "", "", "TRUE", "23", "Có đèn LED"]],
                      columns=C.COT_DATA_PIM)
    return sp, imp, cfg, mt, mf, dp


def test_map_kiem_tra_xuat():
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    opt = C.option_maps(dp)
    r = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh")
    v = r["bang"]["2162"]["rows"][0]["vals"]
    assert v["deep_tskt_master"] == "6.95"
    assert v["utilities_filter_master"] == "13, 23"          # FILTER -> mã option
    assert v["total_power_tskt_master"] == "35 W"            # map dự phòng theo tên cột cấu hình
    spec = pd.DataFrame([["SKU1", "M1", "LOA", "", "utilities_filter_master", "", '["13", " 23"]'],
                         ["SKU1", "M1", "LOA", "", "deep_tskt_master", "", "6.95"]], columns=C.COT_SPEC)
    dv = {("2162", "deep_tskt_master"): "cm"}
    rong = {"đang cập nhật": [C.HD_TRONG, ""]}
    k = C.tinh_kiem_tra(r["bang"], imp, spec, opt, {}, dv, rong)
    tt = {d["ma"]: d["trang_thai"] for d in k["chi_tiet"]["SKU1"]["dong"]}
    assert tt["utilities_filter_master"] == C.TRANG_THAI_GIONG   # ["13", " 23"] đã làm sạch
    assert tt["deep_tskt_master"] == C.TRANG_THAI_DON_VI         # 6.95 -> 6.95 cm
    assert tt["made_in_tskt_master"] == C.TRANG_THAI_BO_QUA      # Đang cập nhật -> để trống
    x = C.xuat_file_import(r["bang"], imp, {("2162", "SKU1", "made_in_tskt_master"): "Việt Nam"}, dv, rong)
    assert len(x["files"]) == 1 and x["so_o_sua"] == 1 and x["so_o_dv"] == 1


def test_tach_va_lam_sach():
    assert C.lam_sach_gia_tri_pim('["905", " 903"]', "utilities_filter_master") == "905, 903"
    assert C.them_don_vi("13 kg", "kg") == "13 kg" and C.them_don_vi("9", "kg") == "9 kg"
    assert C.tach_kich_thuoc_ghep("Ngang 9.8 cm - Sâu 8.5 cm - Cao 16 cm", ["Ngang", "Cao", "Sâu"]) == \
        {"Ngang": "9.8 cm", "Sâu": "8.5 cm", "Cao": "16 cm"}
    assert C.khoa_gia_tri_rong("KHÔNG.") == C.khoa_gia_tri_rong("không")


def test_kiem_tra_thong_minh():
    rows = [{"sku": f"S{i}", "model": "", "variant": "", "cate_pim": "", "so_dong": i,
             "vals": {"mass_tskt_master": v, "deep_tskt_master": d, "made_in_tskt_master": m}}
            for i, (v, d, m) in enumerate([("2.5", "20 cm", "Trung Quốc")] * 10 +
                                          [("91.5", "472 mm", "Trung Quốc."), ("3.1", "18. 4 cm", "trung quốc")])]
    bang = {"2162": {"title": "TSKT 2162 Loa", "attr": ["mass_tskt_master", "deep_tskt_master", "made_in_tskt_master"],
                     "ten": {}, "rows": rows}}
    g = C.kiem_tra_thong_minh(bang, {}, {}, {})
    gy = {(r.sku, r.ma): r.goi_y for r in g.itertuples()}
    assert gy[("S10", "mass_tskt_master")] == "9.15"         # 91.5 lệch ~37x -> sai dấu thập phân
    assert gy[("S10", "deep_tskt_master")] == "47.2 cm"      # mm lẫn trong cột cm
    assert gy[("S11", "deep_tskt_master")] == "18.4 cm"      # dấu thập phân bị tách
    assert gy[("S10", "made_in_tskt_master")] == "Trung Quốc"
    assert gy[("S11", "made_in_tskt_master")] == "Trung Quốc"


def test_doc_json_ai():
    import ai_helper as A
    assert A.doc_json('```json\n{"goi_y": [{"ma": "x"}]}\n```') == {"goi_y": [{"ma": "x"}]}
    assert A.doc_json("không có json") is None


def test_doi_soat():
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    sp = pd.concat([sp, pd.DataFrame([
        ["1", "SKU1", "Loa A", "101", "Tiện ích", "Chống nước IPX7", "2162"],     # FILTER không khớp option
        ["1", "SKU1", "Loa A", "103", "Màu", "Đen", "2162"],                     # mapping trỏ cột không có trong cấu hình
        ["1", "SKU1", "Loa A", "104", "Bluetooth", "5.3", "2162"],               # chỉ ngành khác có mapping
    ], columns=C.COT_DATA_SP)], ignore_index=True)
    mt = pd.concat([mt, pd.DataFrame([["2162", "Loa", "103", "Màu", "color_tskt_master", ""],
                                      ["1942", "Tivi", "104", "Bluetooth", "made_in_tskt_master", ""]],
                                     columns=C.COT_MAP_TSKT)], ignore_index=True)
    opt = C.option_maps(dp)
    r = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh")
    d = C.doi_soat_map(sp, imp, r["bang"], cfg, mt, mf, opt)
    L = d["loi"].set_index("_loai")
    assert "Chống nước IPX7" in L.loc[["filter"], "Giá trị CMS"].tolist()
    assert "13=Chống nước" in L.loc["filter", "Gợi ý"]                       # gợi ý option gần giống
    assert L.loc["cot", "Mã cột"] == "color_tskt_master"
    assert L.loc["nganh_khac", "Mã cột"] == "made_in_tskt_master"
    # lưu quy đổi -> map lại -> hết lỗi FILTER, ô có mã 13
    qd = {f"utilities_filter_master\t{C.khoa_quy_doi('Chống nước IPX7')}": "13"}
    r2 = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh", qd)
    assert r2["bang"]["2162"]["rows"][0]["vals"]["utilities_filter_master"] == "13, 23"
    d2 = C.doi_soat_map(sp, imp, r2["bang"], cfg, mt, mf, opt, qd)
    assert "filter" not in set(d2["loi"]._loai)




def test_de_xuat_sua_cms():
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    opt = C.option_maps(dp)
    # thành viên đề xuất: Sâu 6.95 sai -> 7.95 (chỉ SKU1); FILTER Tiện ích "13, 23" -> "13" (mọi SKU cùng giá trị)
    d1 = C.tao_de_xuat("tv1", "TV 1", C.PV_SKU, "2162", "deep_tskt_master", "6.95", "7.95", sku="SKU1")
    d2 = C.tao_de_xuat("tv1", "TV 1", C.PV_GIA_TRI, "2162", "utilities_filter_master", "13, 23", "13")
    sk, gt, qd, n = C.quy_tac_hieu_luc(({}, {}, {}), [d1, d2], {})
    assert n == 2
    r = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh", qd, gt, sk)
    v = r["bang"]["2162"]["rows"][0]["vals"]
    assert v["deep_tskt_master"] == "7.95" and v["utilities_filter_master"] == "13"
    assert r["goc_cms"]["SKU1\tdeep_tskt_master"] == "6.95" and r["tom_tat"]["o_sua_theo_quy_tac"] == 2
    # admin từ chối d1 -> không còn áp; duyệt d2 -> vào bộ chung
    duyet = {d1["id"]: {"trang_thai": C.DX_TU_CHOI}, d2["id"]: {"trang_thai": C.DX_DUYET}}
    chung_sk, chung_gt, chung_qd = {}, {}, {}
    C.ap_de_xuat_vao_quy_tac(d2, chung_sk, chung_gt, chung_qd)
    sk, gt, qd, n = C.quy_tac_hieu_luc((chung_sk, chung_gt, chung_qd), [d1, d2], duyet)
    assert n == 0
    v = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh", qd, gt, sk)["bang"]["2162"]["rows"][0]["vals"]
    assert v["deep_tskt_master"] == "6.95" and v["utilities_filter_master"] == "13"
    # quy tắc theo SKU tự bỏ khi CMS đã đổi giá trị
    sp2 = sp.copy()
    sp2.loc[sp2.PROPERTYID == "100", "PROPVALUE"] = "8"
    sk = {"SKU1\tdeep_tskt_master": {"cu": C.khoa_quy_doi("6.95"), "moi": "7.95"}}
    v = C.chay_map(sp2, imp, cfg, mt, mf, opt, "cau_hinh", {}, {}, sk)["bang"]["2162"]["rows"][0]["vals"]
    assert v["deep_tskt_master"] == "8"
    # quy đổi FILTER từ đề xuất
    d3 = C.tao_de_xuat("tv1", "", C.PV_QUY_DOI, "2162", "utilities_filter_master", "Chống nước", "23")
    _, _, qd, _ = C.quy_tac_hieu_luc(({}, {}, {}), [d3], {})
    v = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh", qd)["bang"]["2162"]["rows"][0]["vals"]
    assert v["utilities_filter_master"] == "23"


def test_hoan_thien_quy_tac_mau():
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    opt = C.option_maps(dp)
    r = C.chay_map(sp, imp, cfg, mt, mf, opt, "cau_hinh")
    qt = {"2162\tdeep_tskt_master": {"loai": C.QT_SO, "tham_so": "cm", "muc": C.MUC_LOI},
          "*\tmade_in_tskt_master": {"loai": C.QT_BAT_BUOC, "muc": C.MUC_LOI}}
    d = C.do_hoan_thien(r["bang"], cfg, qt, {}, {}, {"đang cập nhật": [C.HD_TRONG, ""]})
    vp = set(d["vi_pham"]["Mã cột"])
    assert vp == {"deep_tskt_master", "made_in_tskt_master"}          # 6.95 chưa có cm; made_in bị để trống
    assert not d["sku"].iloc[0]["Đủ bắt buộc"] and d["nganh"].iloc[0]["Số cột bắt buộc"] == 1
    d2 = C.do_hoan_thien(r["bang"], cfg, qt, {}, {("2162", "deep_tskt_master"): "cm"}, {})
    assert set(d2["vi_pham"]["Mã cột"]) == set()                      # thêm đơn vị + giữ "Đang cập nhật" -> đạt
    assert C.la_cot_filter("utilities_filter_master") and not C.la_cot_filter("filter_tskt_master")
    assert r["chon"][0]["MÃ NH"] == "2162" and r["chon"][0]["NGUỒN CATE"] == "CATEGORYID"
    b = C.xuat_workspace_mau(imp, sp, dp, cfg, mt, mf, r["bang"], r["chon"], r["log"])
    w = C.doc_workspace_cu(b, "ws.xlsx")                              # vòng lại web -> web
    r2 = C.chay_map(w["data_sp"], w["import"], w["cau_hinh"], w["map_tskt"], w["map_filter"],
                    C.option_maps(w["data_pim"]), "cau_hinh")
    assert r2["bang"]["2162"]["rows"][0]["vals"] == r["bang"]["2162"]["rows"][0]["vals"]
    assert w["chon_nganh"] == {"2162": True}


def test_tach_sheet_lon_va_chon_so():
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    goc = C.EXCEL_MAX_DONG
    C.EXCEL_MAX_DONG = 3                                   # giả lập giới hạn dòng Excel
    try:
        b = C.xuat_workspace_mau(imp, sp, dp, cfg, mt, mf, {}, [{"CHỌN": 1, "MÃ NH": "2162"}])
    finally:
        C.EXCEL_MAX_DONG = goc
    w = C.doc_workspace_cu(b, "ws.xlsx")
    assert len(w["data_sp"]) == len(sp) and "DATA SP (2)" in w["sheets"]
    assert w["chon_nganh"] == {"2162": True}               # CHỌN = 1 (số) vẫn là chọn


def test_ten_cot_khac_nhau_va_nap_1_cuc():
    import io
    from openpyxl import Workbook
    def xlsx(rows):
        wb = Workbook(); ws = wb.active
        for r in rows:
            ws.append(r)
        b = io.BytesIO(); wb.save(b); return b.getvalue()
    # CMS export tiêu đề tiếng Việt không dấu + thứ tự khác
    cms = xlsx([["Gia tri", "Ma ERP", "Ma thuoc tinh", "Ten thuoc tinh", "Ma nganh hang"],
                ["6.95", "SKU1", "100", "Sâu", "2162"]])
    df, loi = C.doc_cms_export(cms, "a.xlsx")
    assert not loi and df.iloc[0].tolist() == ["", "SKU1", "", "100", "Sâu", "6.95", "2162"]
    # CMS export KHÔNG tiêu đề -> theo vị trí như desktop
    cms2 = xlsx([["1", "SKU1", "Loa", "100", "Sâu", "6.95", "2162"], ["1", "SKU1", "Loa", "101", "X", "Y", "2162"],
                 ["2", "SKU2", "Loa", "100", "Sâu", "7", "2162"]])
    r = C.nhan_dien_file(cms2, "b.xlsx")
    assert r["loai"] == "cms" and len(r["data_sp"]) == 3 and r["data_sp"].PROPVALUE.tolist()[0] == "6.95"
    # danh sách SKU tiêu đề tiếng Việt + cột TSKT -> IMPORT + spec (tự lọc TSKT)
    ds = xlsx([["Mã model", "Mã biến thể", "Mã sản phẩm ERP", "deep_tskt_master"], ["M1", "", "SKU1", "6.95 cm"],
               ["M1", "V2", "SKU2", ""]])
    r = C.nhan_dien_file(ds, "c.xlsx")
    assert r["loai"] == "sku" and r["import"].sku.tolist() == ["SKU1", "SKU2"] and len(r["spec"]) == 1
    # mapping tiêu đề lạ -> theo vị trí (A ngành, C mã thuộc tính, E mã PIM)
    mt, e = C.doc_mapping_tskt([["x", "y", "z", "t", "u", "v"], ["2162", "Loa", "100", "Sâu", "deep_tskt_master", "Sâu"]])
    assert not e and mt.iloc[0][["cate", "prop_id", "ma"]].tolist() == ["2162", "100", "deep_tskt_master"]
    # file xin data + tách SKU trống khỏi file import
    sp, imp, cfg, mt, mf, dp = _du_lieu()
    imp2 = pd.concat([imp, pd.DataFrame([["M9", "SKU9", "", ""]], columns=C.COT_IMPORT)], ignore_index=True)
    rr = C.chay_map(sp, imp2, cfg, mt, mf, C.option_maps(dp))
    x = C.ds_xin_data(imp2, sp, rr["bang"], rr["log"])
    assert x["sku"].sku.tolist() == ["SKU9"] and x["model"].iloc[0]["Tình trạng"].startswith("MODEL")
    sp_loc, tk = C.loc_data_sp(pd.concat([sp, sp.assign(PRODUCTCODE="KHAC")]), imp)
    assert len(sp_loc) == len(sp) and tk["sku_bo"] == 1
    # biến đổi hàng loạt: đổi đơn vị + thêm chữ, không đụng FILTER
    dv = {("2162", "deep_tskt_master", "bd"): [{"kieu": "them_sau", "a": "cm"}],
          ("*", "utilities_filter_master", "bd"): [{"kieu": "them_sau", "a": "x", "pham_vi": "tat_ca"}]}
    xx = C.xuat_file_import(rr["bang"], imp, {}, dv, {})
    assert xx["so_o_bd"] == 1


def test_gop_kich_thuoc_va_ca_hai():
    assert C.gop_kich_thuoc(["30 cm", "20 cm", "10 cm"], "{1} x {2} x {3} {dv}") == "30 x 20 x 10 cm"
    assert C.gop_kich_thuoc(["30", "20", "10"], "Dài {1} - Rộng {2} - Cao {3}", dv="cm") == \
        "Dài 30 cm - Rộng 20 cm - Cao 10 cm"
    assert C.gop_kich_thuoc(["30 cm", "", "10 cm"], "{1} x {2} x {3} {dv}") == ""          # thiếu -> bỏ qua
    assert C.gop_kich_thuoc(["20", "20", "10"], "Dài {1} - Rộng {2} - Cao {3}", dv="cm") == \
        "Dài 20 cm - Rộng 20 cm - Cao 10 cm"
    assert C.ap_buoc("12", {"kieu": "ca_hai", "a": "Khoảng", "b": "cm"}) == "Khoảng 12 cm"
    assert C.ap_buoc("Có", {"kieu": "ca_hai", "a": "Khoảng", "b": "cm"}) == "Có"             # rule cũ: chỉ số trơn


if __name__ == "__main__":
    test_doi_soat()
    test_map_kiem_tra_xuat()
    test_tach_va_lam_sach()
    test_kiem_tra_thong_minh()
    test_doc_json_ai()
    test_de_xuat_sua_cms()
    test_hoan_thien_quy_tac_mau()
    test_tach_sheet_lon_va_chon_so()
    test_ten_cot_khac_nhau_va_nap_1_cuc()
    test_gop_kich_thuoc_va_ca_hai()
    print("OK — 10/10 test pass")
