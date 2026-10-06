#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
TGDĐ CMS -> PIM — APP LOCAL (1 FILE DUY NHẤT: engine + GUI + workspace)
============================================================================
Gộp toàn bộ vào 1 file để tránh lỗi "quên giữ đủ 2 file cùng thư mục"
(No module named 'pim_tool') từng gặp khi tách riêng pim_tool.py/pim_gui.py.
Chạy: double-click file này (double-click là chạy GUI ngay, KHÔNG cần đối
số dòng lệnh) — Windows sẽ tự mở bằng Python nếu đã cài đặt đúng cách.
Lần chạy ĐẦU TIÊN, app TỰ TẠO thư mục "PIM_Data" (cạnh file .py này) kèm
sẵn 1 file Excel mẫu ("du_lieu_pim.xlsx") có đủ các sheet cần thiết —
KHÔNG cần tự chuẩn bị file Excel nhiều sheet đúng cấu trúc trước. Vào
thẳng tab "📋 Quản lý dữ liệu" gõ cấu hình (CẤU HÌNH CATEGORY, MAPPING
TSKT MOI, MAPPING FILTER MOI, IMPORT), dùng nút "📥 Nhập dữ liệu thô từ
file khác..." để nạp file CMS export vào, rồi bấm "🚀 CHẠY TẤT CẢ".
NGUYÊN TẮC NGHIỆP VỤ (giữ nguyên từ Google Apps Script, đã qua nhiều lượt
đối chiếu + sửa lỗi thật trên Sheet):
  1) TSKT (đọc từ sheet "MAPPING TSKT MOI") -> LUÔN LUÔN là TEXT, giữ
     nguyên văn PROPVALUE từ DATA SP, KHÔNG bao giờ tra số qua DATA PIM.
     Nhiều giá trị nối bằng "|" (KHÔNG cách 2 bên).
  2) FILTER (đọc từ sheet "MAPPING FILTER MOI") -> LUÔN LUÔN là SỐ, tra
     đúng cặp (mã filter, giá trị text) qua DATA PIM. Không khớp được ->
     để TRỐNG (không suy đoán/bịa số) + ghi LOG. Nhiều giá trị nối bằng ", ".
  3) DATA SP: dòng 1 để trống, dòng 2 = header, dữ liệu bắt đầu DÒNG 3 —
     header LUÔN được ghi/refresh lại mỗi lần trích (tự phục hồi).
  4) Trích DATA SP từ sheet dữ liệu thô: tự nhận diện cột theo TÊN (không
     phụ thuộc thứ tự cột nguồn).
  5) Xuất file: mỗi category ra 2 file .xlsx riêng biệt — MODEL (variant_code
     trống) và BIẾN THỂ (variant_code có giá trị) — mọi ô ép định dạng Text.
BỔ SUNG (bản này):
  6) "📏 Đơn vị kích thước / khối lượng": tự tìm cột Sâu/Cao/Ngang/Khối
     lượng/Kích thước..., chọn cm/mm/kg/g/inch -> thêm vào ô SỐ TRƠN; lưu ở
     sheet "ĐƠN VỊ", mỗi lần CHẠY TẤT CẢ tự áp lại. Không đụng FILTER.
  7) "📦 Nạp file export PIM": model_code/sku/variant_code/category_code ->
     IMPORT; spec -> sheet "SPEC PIM TẠM" để đối chiếu (sheet "ĐỐI CHIẾU").
  8) Khi xuất: cảnh báo rõ (sheet "CẢNH BÁO XUẤT" + file .txt trong .zip +
     hộp thoại đỏ) — thiếu model_code/category_code, khác spec PIM, ô kích
     thước/khối lượng còn số trơn, SKU không có trong file PIM đã nạp.
  9) Tab "🔍 Kiểm tra & Đối chiếu": ① map dữ liệu (chưa xuất) -> ② kiểm
     tra + SỬA TRỰC TIẾP (sửa ô, lấy PIM cũ, tự lấp ô trống, đơn vị hàng
     loạt — giữ trong bộ nhớ, tính lại tức thì) -> ③ kiểm tra lần cuối +
     xác nhận -> xuất ĐÚNG 1 LẦN (thư mục .xlsx + .zip) và ghi chỉnh sửa vào
     workspace. Giá trị PIM dạng ["905", " 903"] tự làm sạch thành 905, 903.
CẢI TIẾN so với bản Apps Script gốc: sheet "THUỘC TÍNH CMS/PIM" (mapping
CŨ) KHÔNG còn bắt buộc phải tồn tại nữa (trước đây là quirk kế thừa gây
lỗi ngay từ đầu với workspace mới tạo còn trống).
Dùng dòng lệnh (nâng cao, chạy không giao diện) vẫn được:
    python pim_app.py <file_excel.xlsx> --out <thư_mục_xuất>
============================================================================
"""
from __future__ import annotations
import argparse
import io
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
import unicodedata
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
# ============================================================================
# §1 CONFIG — y hệt CFG trong Code.gs
# ============================================================================
class CFG:
    sheet_nhap = "IMPORT"
    tskth_prefix = "TSKT"
    sheet_chon = "CHỌN NGÀNH HÀNG"
    sheet_cau_hinh = "CẤU HÌNH CATEGORY"
    sheet_log = "LOG"
    sheet_data_sp = ["DATA SP"]
    sheet_data_pim = ["DATA PIM"]
    sheet_mapping = ["THUỘC TÍNH CMS/PIM", "THUỘC TÍNH CMSPIM"]
    sheet_detail = "XỬ LÝ HÀM"
    bo_qua_sheet = ["IMPORT FILTER"]
    header_row = 1
    name_row = 2
    data_start_row = 3
    cot_co_dinh = ["model_code", "sku", "category_code", "variant_code"]
    bot_cot_khi_xuat = ["sku"]
    # DATA SP: dòng 1 để trống, dòng 2 = header, dữ liệu từ DÒNG 3.
    # DATA SP thuc te (doi chieu truc tiep file that nguoi dung xuat): header
    # o DONG 1, du lieu tu DONG 2 - KHONG co dong 1 de trong nhu gia dinh
    # truoc do (tung dua theo anh chup 1 workspace da qua tool nay chinh
    # sua). File goc that su cua nguoi dung dung dung quy uoc chuan: dong 1
    # = header, dong 2 tro di = du lieu.
    sp_start_row = 2
    sp_sku = 2
    sp_name = 3
    sp_prop_id = 4
    sp_prop_name = 5
    sp_prop_value = 6
    sp_cate_id = 7
    pim_start_row = 3
    pim_code = 1
    pim_opt_code = 6
    pim_opt_value = 7
    map_start_row = 2
    map_cate = 1
    map_cate_name = 2
    map_prop_id = 3
    map_pim_code = 5
    map_pim_name = 7
    # TSKT master (text) -> "|" không cách 2 bên. FILTER (số) -> ", ".
    value_separator = "|"
    value_separator_filter = ", "
    file_prefix = "PIM_"
CFG_TF_sheet_tskt_mapping = "MAPPING TSKT MOI"
CFG_TF_sheet_filter_mapping = "MAPPING FILTER MOI"
# BỔ SUNG (đơn vị + đối chiếu spec PIM + cảnh báo khi xuất):
CFG_sheet_don_vi = "ĐƠN VỊ"            # lưu đơn vị đã chọn theo (ngành hàng, mã TSKT)
CFG_sheet_spec_pim = "SPEC PIM TẠM"    # spec lấy từ file export PIM, lưu tạm để đối chiếu
CFG_sheet_doi_chieu = "ĐỐI CHIẾU"      # ô tool điền KHÁC spec PIM đang có (tạo lại mỗi lần chạy)
CFG_sheet_canh_bao = "CẢNH BÁO XUẤT"   # tổng hợp cảnh báo trước khi import (tạo lại mỗi lần chạy)
CFG_sheet_chinh_sua = "CHỈNH SỬA TAY"  # nhật ký các ô sửa trực tiếp ở tab Kiểm tra
CFG_sheet_tien_ich = [CFG_sheet_don_vi, CFG_sheet_spec_pim, CFG_sheet_doi_chieu, CFG_sheet_canh_bao, CFG_sheet_chinh_sua]
# ============================================================================
# §2 CHUẨN HOÁ — y hệt chuanHoaCode_/chuanHoaKey_/chuanHoaId_ trong Code.gs
# ============================================================================
_WS_RE = re.compile(r"\s+")
_NBSP = " "
def chuan_hoa_key(v) -> str:
    """String(v).replace(nbsp,' ').trim() — GIỮ khoảng trắng giữa từ."""
    if v is None:
        return ""
    s = str(v).replace(_NBSP, " ")
    return s.strip()
def chuan_hoa_code(v) -> str:
    """Như chuan_hoa_key nhưng CẮT SẠCH mọi khoảng trắng (dùng cho mã/SKU)."""
    s = chuan_hoa_key(v)
    return _WS_RE.sub("", s)
def chuan_hoa_id(v) -> str:
    """Cắt đuôi '.0' nếu Excel/pandas đọc ID số thành số thập phân."""
    s = chuan_hoa_code(v)
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s
def ep_van_ban_an_toan(v):
    """Ép giá trị số (int/float) đọc được từ file Excel về ĐÚNG dạng
    VĂN BẢN THẬT SỰ (không chỉ đổi number_format hiển thị — cách đó
    KHÔNG đổi kiểu dữ liệu lưu bên trong, Excel vẫn coi là số).
    QUAN TRỌNG: các mã (MÃ NGÀNH HÀNG, MÃ THUỘC TÍNH, MÃ TSKT...) tuy
    trông giống số nhưng LUÔN phải là văn bản — nếu để lẫn lộn (chỗ thì
    số 42358, chỗ thì chữ "42358", chỗ thì "42358.0" do Excel tự đọc số
    thập phân) thì các bước so khớp/lọc theo mã ở phía sau sẽ SO SÁNH
    KHÔNG KHỚP, khiến dữ liệu vừa nhập trông như "biến mất"/"không lấy
    được" dù thật ra vẫn nằm trong sheet, chỉ là không được các bước sau
    nhận diện đúng."""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return str(v)
    if isinstance(v, int):
        return str(v)
    return v
# ============================================================================
# §3 TIỆN ÍCH ĐỌC SHEET — y hệt timSheet_/docHeaderImport_/locCotThuocTinh_/
#    dongCuoi_/dongCuoiMoiCot_ trong Code.gs
# ============================================================================
def tim_sheet(wb: Workbook, ten_list: List[str]) -> Optional[Worksheet]:
    for ten in ten_list:
        if ten in wb.sheetnames:
            return wb[ten]
    return None
def sheet_display_values(sh: Worksheet, min_row: int, max_row: int, max_col: int) -> List[List[str]]:
    """Đọc 1 vùng ra dạng text hiển thị (giống getDisplayValues() của Apps
    Script) — mọi giá trị (kể cả số/ngày) đều ép về string trước khi trả
    về, để logic phía sau xử lý nhất quán."""
    out = []
    for row in sh.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col, values_only=True):
        out.append(["" if v is None else str(v) for v in row])
    return out
def doc_header_import(sh: Worksheet) -> List[Tuple[str, int]]:
    """Đọc dòng CFG.header_row -> [(code, col_1based), ...], bỏ cột trùng
    tên/rỗng — y hệt docHeaderImport_()."""
    if not sh.max_column or sh.max_column < 1:  # None (sheet đọc read_only
        # bị thiếu thẻ <dimension>) hoặc 0 -> coi như không có header
        return []
    raw = sheet_display_values(sh, CFG.header_row, CFG.header_row, sh.max_column)[0]
    cols = []
    seen = set()
    for i, h in enumerate(raw):
        code = chuan_hoa_key(h)
        if not code or code in seen:
            continue
        seen.add(code)
        cols.append((code, i + 1))
    return cols
def loc_cot_thuoc_tinh(cols: List[Tuple[str, int]]) -> List[Tuple[str, int]]:
    co_dinh = {x.lower() for x in CFG.cot_co_dinh}
    return [(c, i) for c, i in cols if c.lower() not in co_dinh]
def dong_cuoi(sh: Worksheet, col: int, start_row: int) -> int:
    """Dòng cuối cùng CÓ dữ liệu trong `col`, tính từ start_row — y hệt
    dongCuoi_()."""
    last_row = sh.max_row
    if not last_row or last_row < start_row:
        return start_row - 1
    for r in range(last_row, start_row - 1, -1):
        v = sh.cell(row=r, column=col).value
        if v is not None and str(v).strip() != "":
            return r
    return start_row - 1
def dong_cuoi_moi_cot(sh: Worksheet, cols: List[Tuple[str, int]]) -> int:
    fix_cols = [(c, i) for c, i in cols if c.lower() in {x.lower() for x in CFG.cot_co_dinh}]
    check = fix_cols if fix_cols else cols[:3]
    m = CFG.data_start_row - 1
    for _, col in check:
        d = dong_cuoi(sh, col, CFG.data_start_row)
        if d > m:
            m = d
    return m
def tim_cot_theo_ten(header: List[str], *ten_bien_the: str) -> int:
    """Tìm cột theo TÊN header (không phân biệt hoa/thường, không phân
    biệt khoảng trắng thừa) — y hệt timCotTheoTen_(). Trả về số cột
    1-based, 0 nếu không tìm thấy."""
    chuan = {chuan_hoa_key(t).upper() for t in ten_bien_the}
    for i, h in enumerate(header):
        if chuan_hoa_key(h).upper() in chuan:
            return i + 1
    return 0
def set_text_format(sh: Worksheet, row: int, col: int, value) -> None:
    c = sh.cell(row=row, column=col, value=value)
    c.number_format = "@"
# ============================================================================
# §4 CAU HINH CATEGORY (bố cục ngang) — y hệt docCauHinh_()
# ============================================================================
class CategoryConfig:
    def __init__(self, name: str = "", dong_dau: int = 0):
        self.name = name
        self.tskt: List[str] = []
        self._seen: set = set()
        self.ten_tskt: Dict[str, str] = {}
        self.dong_dau = dong_dau
_MA_TSKT_RE = re.compile(r"^[A-Za-z0-9_]+$")
def la_ma_tskt(s: str) -> bool:
    return bool(_MA_TSKT_RE.fullmatch(s))
def doc_cau_hinh(wb: Workbook) -> Dict[str, CategoryConfig]:
    cfg: Dict[str, CategoryConfig] = {}
    if CFG.sheet_cau_hinh not in wb.sheetnames:
        return cfg
    sh = wb[CFG.sheet_cau_hinh]
    if sh.max_row < 2 or sh.max_column < 3:
        return cfg
    vals = sheet_display_values(sh, 2, sh.max_row, sh.max_column)
    cate: Optional[str] = None
    last_code_cols: Optional[Dict[int, str]] = None
    for idx, r in enumerate(vals):
        rid = chuan_hoa_id(r[0] if len(r) > 0 else "")
        name = chuan_hoa_key(r[1] if len(r) > 1 else "")
        if rid and rid != cate:
            cate = rid
            last_code_cols = None
        if not cate:
            continue
        if cate not in cfg:
            cfg[cate] = CategoryConfig(name=name, dong_dau=idx + 2)
        elif name and not cfg[cate].name:
            cfg[cate].name = name
        o = cfg[cate]
        cells = []
        for c in range(2, len(r)):
            v = chuan_hoa_key(r[c])
            if v:
                cells.append((c, v))
        if not cells:
            continue
        la_dong_ma = all(la_ma_tskt(v) for _, v in cells) and any("_" in v for _, v in cells)
        if la_dong_ma:
            last_code_cols = {}
            for c, v in cells:
                last_code_cols[c] = v
                if v in o._seen:
                    continue
                o._seen.add(v)
                o.tskt.append(v)
        else:
            if last_code_cols is None:
                continue
            for c, v in cells:
                code = last_code_cols.get(c)
                if code and code not in o.ten_tskt:
                    o.ten_tskt[code] = v
    return cfg
# ============================================================================
# §5 CACHE — y hệt docCache_() + docTsktMapping_()/docFilterMapping_()
# ============================================================================
class Cache:
    def __init__(self):
        self.loi: Optional[str] = None
        self.sp_vals: List[List[str]] = []
        self.co_cot_cate_id = False
        self.cate_khai_bao: Dict[str, str] = {}
        self.cate_mau_thuan: Dict[str, set] = {}
        self.map_theo_cate: Dict[str, Dict[str, str]] = defaultdict(dict)
        self.ten_cate: Dict[str, str] = {}
        self.cate_cua_prop_id: Dict[str, set] = defaultdict(set)
        self.pim_set_theo_cate: Dict[str, set] = defaultdict(set)
        self.ten_pim: Dict[str, str] = {}
        self.option_map: Dict[Tuple[str, str], str] = {}
        self.la_select: set = set()
        self.map_tskt_theo_cate: Dict[str, Dict[str, str]] = defaultdict(dict)
        self.map_filter_theo_cate: Dict[str, Dict[str, str]] = defaultdict(dict)
def doc_cache(wb: Workbook) -> Cache:
    cache = Cache()
    sh_sp = tim_sheet(wb, CFG.sheet_data_sp)
    sh_pim = tim_sheet(wb, CFG.sheet_data_pim)
    sh_map = tim_sheet(wb, CFG.sheet_mapping)
    # CAI TIEN so voi ban Apps Script goc: sheet "THUOC TINH CMSPIM" (bang
    # mapping CU) KHONG con bat buoc nua - luong TSKT+FILTER moi (doc qua
    # doc_tskt_mapping/doc_filter_mapping) khong can du lieu cua no de chay
    # dung; ban GAS cu bat buoc sheet nay ton tai la 1 quirk ke thua khong
    # can thiet, gay loi ngay tu dau voi workspace moi tao con trong. Neu
    # sheet CO du lieu thi van doc de bo sung (khong mat gi), CHI khi
    # THIEU/TRONG thi bo qua em, khong dung pipeline lai.
    if not sh_sp or not sh_pim:
        cache.loi = "Thiếu sheet DATA SP hoặc DATA PIM."
        return cache
    sp_last = dong_cuoi(sh_sp, CFG.sp_sku, CFG.sp_start_row)
    if sp_last < CFG.sp_start_row:
        cache.loi = "Sheet DATA SP trống."
        return cache
    sp_width = min(max(CFG.sp_prop_value, CFG.sp_cate_id), sh_sp.max_column)
    co_cot_cate_id = sp_width >= CFG.sp_cate_id
    sp_vals = sheet_display_values(sh_sp, CFG.sp_start_row, sp_last, sp_width)
    cache.sp_vals = sp_vals
    cache.co_cot_cate_id = co_cot_cate_id
    if co_cot_cate_id:
        for r in sp_vals:
            sku = chuan_hoa_code(r[CFG.sp_sku - 1])
            cid = chuan_hoa_id(r[CFG.sp_cate_id - 1])
            if not sku or not cid:
                continue
            if sku in cache.cate_khai_bao and cache.cate_khai_bao[sku] != cid:
                cache.cate_mau_thuan.setdefault(sku, {cache.cate_khai_bao[sku]}).add(cid)
                continue
            cache.cate_khai_bao[sku] = cid
    if sh_map is not None:
        map_last = dong_cuoi(sh_map, CFG.map_cate, CFG.map_start_row)
        if map_last >= CFG.map_start_row:
            map_width = min(max(CFG.map_pim_code, CFG.map_pim_name), sh_map.max_column)
            map_vals = sheet_display_values(sh_map, CFG.map_start_row, map_last, map_width)
            for r in map_vals:
                cate = chuan_hoa_id(r[CFG.map_cate - 1])
                prop_id = chuan_hoa_id(r[CFG.map_prop_id - 1])
                pim_code = chuan_hoa_key(r[CFG.map_pim_code - 1])
                if not cate or not prop_id or not pim_code:
                    continue
                cache.ten_cate.setdefault(cate, chuan_hoa_key(r[CFG.map_cate_name - 1]))
                cache.map_theo_cate[cate].setdefault(prop_id, pim_code)
                cache.cate_cua_prop_id[prop_id].add(cate)
                cache.pim_set_theo_cate[cate].add(pim_code)
                if map_width >= CFG.map_pim_name:
                    tn = chuan_hoa_key(r[CFG.map_pim_name - 1])
                    if tn:
                        cache.ten_pim.setdefault(pim_code, tn)
    pim_last = dong_cuoi(sh_pim, CFG.pim_code, CFG.pim_start_row)
    if pim_last < CFG.pim_start_row:
        cache.loi = "Sheet DATA PIM trống."
        return cache
    pim_vals = sheet_display_values(sh_pim, CFG.pim_start_row, pim_last, CFG.pim_opt_value)
    for r in pim_vals:
        code = chuan_hoa_key(r[CFG.pim_code - 1])
        if not code:
            continue
        cache.la_select.add(code)
        opt_value = chuan_hoa_key(r[CFG.pim_opt_value - 1])
        if not opt_value:
            continue
        key = (code, opt_value.lower())
        if key not in cache.option_map:
            # BUG THAT tu doi chieu file that: OptionCode trong DATA PIM
            # thuc te la kieu SO (Excel luu vd 215361.0), chuan_hoa_key()
            # chi cat khoang trang - KHONG cat duoi ".0" - se lam ma FILTER
            # tra ra bi sai thanh "215361.0" thay vi "215361". Dung
            # chuan_hoa_id() (da co san logic cat duoi ".0" cho ma so, van
            # giu nguyen ma dang chu nhu "a1").
            cache.option_map[key] = chuan_hoa_id(r[CFG.pim_opt_code - 1])
    return cache
def doc_tskt_mapping(wb: Workbook, cache: Cache) -> None:
    if CFG_TF_sheet_tskt_mapping not in wb.sheetnames:
        raise RuntimeError(
            f'Không tìm thấy sheet "{CFG_TF_sheet_tskt_mapping}". '
            "Tạo sheet này và dán dữ liệu từ 99.xlsx vào trước."
        )
    sh = wb[CFG_TF_sheet_tskt_mapping]
    if sh.max_row < 2:
        return
    header = sheet_display_values(sh, 1, 1, sh.max_column)[0]
    c_cate = tim_cot_theo_ten(header, "Mã ngành hàng CMS", "Mã ngành hàng")
    c_cate_name = tim_cot_theo_ten(header, "Tên ngành hàng CMS", "Tên ngành hàng")
    # "MÃ THUỘC TÍNH TSKT" / "TÊN MASTER": tiêu đề file mẫu mới (TEST HÀNG LOẠT IMPORT THÔNG SỐ NEW)
    c_prop_id = tim_cot_theo_ten(header, "Mã thuộc tính", "Mã thuộc tính TSKT")
    c_pim_code = tim_cot_theo_ten(header, "Mã TSKT (MASTER)", "Mã MASTER")
    c_pim_name = tim_cot_theo_ten(header, "Tên TSKT (MASTER)", "Tên MASTER")
    if not c_cate or not c_prop_id or not c_pim_code:
        raise RuntimeError(
            f'Sheet "{CFG_TF_sheet_tskt_mapping}" thiếu cột bắt buộc (cần có dòng 1 là header với tên: '
            "Mã ngành hàng CMS / Mã thuộc tính / Mã TSKT (MASTER))."
        )
    vals = sheet_display_values(sh, 2, sh.max_row, sh.max_column)
    for r in vals:
        cate = chuan_hoa_id(r[c_cate - 1])
        prop_id = chuan_hoa_id(r[c_prop_id - 1])
        pim_code = chuan_hoa_key(r[c_pim_code - 1])
        if not cate or not prop_id or not pim_code:
            continue
        cache.map_tskt_theo_cate[cate].setdefault(prop_id, pim_code)
        cate_name = chuan_hoa_key(r[c_cate_name - 1]) if c_cate_name else ""
        if cate_name:
            cache.ten_cate.setdefault(cate, cate_name)
        cache.map_theo_cate[cate].setdefault(prop_id, pim_code)
        cache.cate_cua_prop_id[prop_id].add(cate)
        cache.pim_set_theo_cate[cate].add(pim_code)
        if c_pim_name:
            pim_name = chuan_hoa_key(r[c_pim_name - 1])
            if pim_name:
                cache.ten_pim.setdefault(pim_code, pim_name)
def doc_filter_mapping(wb: Workbook, cache: Cache) -> None:
    if CFG_TF_sheet_filter_mapping not in wb.sheetnames:
        raise RuntimeError(
            f'Không tìm thấy sheet "{CFG_TF_sheet_filter_mapping}". '
            "Tạo sheet này và dán dữ liệu từ 88.xlsx vào trước."
        )
    sh = wb[CFG_TF_sheet_filter_mapping]
    if sh.max_row < 2:
        return
    header = sheet_display_values(sh, 1, 1, sh.max_column)[0]
    c_cate = tim_cot_theo_ten(header, "MÃ NGÀNH HÀNG CMS", "Mã ngành hàng CMS")
    c_cate_name = tim_cot_theo_ten(header, "TÊN NGÀNH HÀNG CMS", "Tên ngành hàng CMS")
    c_prop_id = tim_cot_theo_ten(header, "MÃ THUỘC TÍNH FILTER", "Mã thuộc tính filter")
    c_code_new = tim_cot_theo_ten(header, "Mã thuộc tính mới")
    c_code_old = tim_cot_theo_ten(header, "Mã thuộc tính cũ")
    if not c_cate or not c_prop_id or (not c_code_new and not c_code_old):
        raise RuntimeError(
            f'Sheet "{CFG_TF_sheet_filter_mapping}" thiếu cột bắt buộc (cần có dòng 1 là header với tên: '
            "MÃ NGÀNH HÀNG CMS / MÃ THUỘC TÍNH FILTER / Mã thuộc tính mới hoặc Mã thuộc tính cũ)."
        )
    vals = sheet_display_values(sh, 2, sh.max_row, sh.max_column)
    for r in vals:
        cate = chuan_hoa_id(r[c_cate - 1])
        prop_id = chuan_hoa_id(r[c_prop_id - 1])
        filter_code = chuan_hoa_key(r[c_code_new - 1]) if c_code_new else ""
        if not filter_code and c_code_old:
            filter_code = chuan_hoa_key(r[c_code_old - 1])
        if not cate or not prop_id or not filter_code:
            continue
        cache.map_filter_theo_cate[cate].setdefault(prop_id, filter_code)
        cate_name = chuan_hoa_key(r[c_cate_name - 1]) if c_cate_name else ""
        if cate_name:
            cache.ten_cate.setdefault(cate, cate_name)
        cache.map_theo_cate[cate].setdefault(prop_id, filter_code)
        cache.cate_cua_prop_id[prop_id].add(cate)
        cache.pim_set_theo_cate[cate].add(filter_code)
# ============================================================================
# §6 IMPORT + PHÂN LOẠI — y hệt docNhapLieu_()/phanLoai_()
# ============================================================================
class NhapLieuRec:
    __slots__ = ("sku", "model", "variant", "cate_pim")
    def __init__(self, sku: str, model: str, variant: str, cate_pim: str = ""):
        self.sku = sku
        self.model = model
        self.variant = variant
        self.cate_pim = cate_pim  # category_code (Mã danh mục PIM) — dùng để cảnh báo khi xuất
def doc_nhap_lieu(wb: Workbook) -> Tuple[Optional[str], List[NhapLieuRec], Dict[str, NhapLieuRec]]:
    if CFG.sheet_nhap not in wb.sheetnames:
        return f"Không tìm thấy sheet {CFG.sheet_nhap}.", [], {}
    sh = wb[CFG.sheet_nhap]
    cols = doc_header_import(sh)
    col_cua = {c.lower(): i for c, i in cols}
    c_sku = col_cua.get("sku")
    if not c_sku:
        return f'Sheet {CFG.sheet_nhap}: dòng 1 phải có cột "sku".', [], {}
    c_model = col_cua.get("model_code", 0)
    c_variant = col_cua.get("variant_code", 0)
    c_cate_pim = col_cua.get("category_code", 0)
    last_row = dong_cuoi(sh, c_sku, CFG.data_start_row)
    if last_row < CFG.data_start_row:
        return f"Chưa dán SKU vào sheet {CFG.sheet_nhap} (từ dòng 3).", [], {}
    max_col = max(c_sku, c_model, c_variant, c_cate_pim)
    vals = sheet_display_values(sh, CFG.data_start_row, last_row, max_col)
    rows: List[NhapLieuRec] = []
    m: Dict[str, NhapLieuRec] = {}
    for r in vals:
        sku = chuan_hoa_code(r[c_sku - 1])
        if not sku or sku in m:
            continue
        rec = NhapLieuRec(
            sku=sku,
            model=chuan_hoa_key(r[c_model - 1]) if c_model else "",
            variant=chuan_hoa_key(r[c_variant - 1]) if c_variant else "",
            cate_pim=chuan_hoa_id(r[c_cate_pim - 1]) if c_cate_pim else "",
        )
        rows.append(rec)
        m[sku] = rec
    if not rows:
        return f"Không đọc được SKU hợp lệ nào trong {CFG.sheet_nhap}.", [], {}
    return None, rows, m
class CateInfo:
    def __init__(self):
        self.skus: List[str] = []
        self.tu_cate_id = 0
        self.tu_doan = 0
        self.tie = False
def phan_loai(cache: Cache, nl_rows: List[NhapLieuRec]) -> Tuple[Dict[str, CateInfo], List[str], List[list]]:
    sku_set = {r.sku for r in nl_rows}
    log_rows: List[list] = []
    prop_id_theo_sku: Dict[str, set] = defaultdict(set)
    for r in cache.sp_vals:
        sku = chuan_hoa_code(r[CFG.sp_sku - 1])
        if not sku or sku not in sku_set:
            continue
        prop_id = chuan_hoa_id(r[CFG.sp_prop_id - 1])
        if not prop_id:
            continue
        prop_id_theo_sku[sku].add(prop_id)
    theo_cate: Dict[str, CateInfo] = {}
    khong_co_data: List[str] = []
    def them(cate: str, sku: str, tu_cate_id: bool, tie: bool):
        o = theo_cate.setdefault(cate, CateInfo())
        o.skus.append(sku)
        if tu_cate_id:
            o.tu_cate_id += 1
        else:
            o.tu_doan += 1
        if tie:
            o.tie = True
    for rec in nl_rows:
        sku = rec.sku
        if sku in cache.cate_mau_thuan:
            arr = sorted(cache.cate_mau_thuan[sku])
            log_rows.append([sku, " | ".join(arr), "",
                              f"CATEGORYID mâu thuẫn trong DATA SP ({', '.join(arr)}) — dùng giá trị đầu: "
                              f"{cache.cate_khai_bao.get(sku, '')}"])
        khai_bao = cache.cate_khai_bao.get(sku)
        if khai_bao:
            if khai_bao not in cache.map_theo_cate:
                log_rows.append([sku, khai_bao, "",
                                  "CATEGORYID không tồn tại trong mapping — bỏ qua SKU"])
                khong_co_data.append(sku)
                continue
            them(khai_bao, sku, True, False)
            continue
        prop_ids = prop_id_theo_sku.get(sku)
        if not prop_ids:
            log_rows.append([sku, "", "", "Không có dòng nào trong DATA SP (không PROPERTYID, không CATEGORYID)"])
            khong_co_data.append(sku)
            continue
        diem: Dict[str, int] = defaultdict(int)
        for pid in prop_ids:
            for c in sorted(cache.cate_cua_prop_id.get(pid, ())):
                diem[c] += 1
        if not diem:
            log_rows.append([sku, "", "", "PROPERTYID của SKU không khớp cate nào trong mapping"])
            khong_co_data.append(sku)
            continue
        best_cate, best, nhi = None, -1, -1
        for k in sorted(diem):
            v = diem[k]
            if v > best:
                nhi, best, best_cate = best, v, k
            elif v > nhi:
                nhi = v
        tie = nhi == best
        if tie:
            log_rows.append([sku, best_cate, "",
                              f"Tự nhận diện bị trùng điểm giữa 2+ cate — đã chọn {best_cate}, "
                              "nên bổ sung CATEGORYID vào DATA SP"])
        them(best_cate, sku, False, tie)
    return theo_cate, khong_co_data, log_rows
# ============================================================================
# §7 TÌM/TẠO TAB TSKT TRONG CÙNG WORKBOOK — y hệt timTabTskt_()/taoTabTuCauHinh_()
# ============================================================================
def tim_tab_tskt(wb: Workbook, cache: Cache) -> Dict[str, Worksheet]:
    p = CFG.tskth_prefix.lower()
    bo_qua = {x.lower() for x in CFG.bo_qua_sheet}
    for x in (CFG.sheet_nhap, CFG.sheet_chon, CFG.sheet_detail, CFG.sheet_cau_hinh, CFG.sheet_log):
        bo_qua.add(x.lower())
    for x in CFG.sheet_data_sp + CFG.sheet_data_pim + CFG.sheet_mapping:
        bo_qua.add(x.lower())
    bo_qua.add(CFG_TF_sheet_tskt_mapping.lower())
    bo_qua.add(CFG_TF_sheet_filter_mapping.lower())
    for x in CFG_sheet_tien_ich:
        bo_qua.add(x.lower())
    kq: Dict[str, Worksheet] = {}
    for sh in wb.worksheets:
        n = sh.title.lower()
        if n in bo_qua:
            continue
        if n != p and not n.startswith(p + " "):
            continue
        cols = doc_header_import(sh)
        attr = loc_cot_thuoc_tinh(cols)
        if not attr:
            continue
        cate_tu_ten = None
        if n.startswith(p + " "):
            phan_con_lai = chuan_hoa_key(sh.title[len(CFG.tskth_prefix):]).strip()
            token_dau = phan_con_lai.split()[0] if phan_con_lai else ""
            id_thu = chuan_hoa_id(token_dau)
            if id_thu and id_thu in cache.map_theo_cate:
                cate_tu_ten = id_thu
        if cate_tu_ten:
            kq.setdefault(cate_tu_ten, sh)
            continue
        cate, best = None, 0
        for c, pim_set in cache.pim_set_theo_cate.items():
            d = sum(1 for code, _ in attr if code in pim_set)
            if d > best:
                best, cate = d, c
        if cate:
            kq.setdefault(cate, sh)
    return kq
def _safe_sheet_title(name: str) -> str:
    name = re.sub(r'[\\/*?:\[\]]', "_", name).strip()
    return name[:31] if name else "Sheet"
def tao_tab_tu_cau_hinh(wb: Workbook, cate: str, cfg_cate: CategoryConfig, cache: Cache) -> Worksheet:
    ten = f"{CFG.tskth_prefix} {cate} {cfg_cate.name or cache.ten_cate.get(cate, '')}".strip()
    ten = _safe_sheet_title(ten[:90])
    if ten in wb.sheetnames:
        return wb[ten]
    sh = wb.create_sheet(ten)
    header = ["model_code", "sku", "variant_code"] + cfg_cate.tskt
    name_row = []
    for h in header:
        if h == "model_code":
            name_row.append("Mã model")
        elif h == "sku":
            name_row.append("Mã sản phẩm ERP")
        elif h == "variant_code":
            name_row.append("Mã biến thể")
        else:
            name_row.append(cfg_cate.ten_tskt.get(h) or cache.ten_pim.get(h, ""))
    bold = Font(bold=True)
    for col, (h, n) in enumerate(zip(header, name_row), start=1):
        c1 = sh.cell(row=1, column=col, value=h)
        c1.number_format = "@"
        c1.font = bold
        c2 = sh.cell(row=2, column=col, value=n)
        c2.number_format = "@"
        c2.font = bold
    sh.freeze_panes = "A3"
    return sh
# ============================================================================
# §8 GHI LOG — y hệt ghiLog_()
# ============================================================================
def ghi_log(wb: Workbook, rows: List[list]) -> None:
    if not rows:
        return
    if CFG.sheet_log in wb.sheetnames:
        sh = wb[CFG.sheet_log]
    else:
        sh = wb.create_sheet(CFG.sheet_log)
        bold = Font(bold=True)
        headers = ["THỜI GIAN", "SKU", "CATEGORY", "TSKT", "NGUYÊN NHÂN"]
        for col, h in enumerate(headers, start=1):
            c = sh.cell(row=1, column=col, value=h)
            c.font = bold
            c.number_format = "@"
        sh.freeze_panes = "A2"
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    start_row = sh.max_row + 1
    for i, r in enumerate(rows):
        vals = [stamp] + [r[j] if j < len(r) and r[j] else "" for j in range(4)]
        for col, v in enumerate(vals, start=1):
            c = sh.cell(row=start_row + i, column=col, value=v)
            c.number_format = "@"
# ============================================================================
# §9 TỐI ƯU GHI HÀNG LOẠT — định dạng TEXT ở CẤP CỘT thay vì từng ô
# ============================================================================
def dat_dinh_dang_van_ban_cot(sh: Worksheet, so_cot: int) -> None:
    """Đặt định dạng TEXT ("@") Ở CẤP CỘT (column_dimensions) thay vì gọi
    cell.number_format cho TỪNG Ô một trong vòng lặp lớn. Excel áp dụng
    định dạng cột cho mọi ô trong cột đó KHÔNG có định dạng riêng ở cấp ô
    — vì các hàm ghi hàng loạt dưới đây không set number_format riêng cho
    từng ô nữa, cách này cho hiệu quả hiển thị Y HỆT set_text_format() cũ
    nhưng nhanh hơn nhiều lần với sheet lớn (chục nghìn dòng): tránh việc
    openpyxl phải tạo/tra 1 style record mới cho MỖI ô một, vốn là phần
    tốn thời gian nhất khi ghi dữ liệu số lượng lớn — đây chính là nguyên
    nhân chính khiến các bước "nạp dữ liệu" cảm giác rất lâu với file to.
    Gọi hàm này 1 LẦN trước khi bắt đầu vòng lặp ghi giá trị, rồi ghi giá
    trị bằng sh.cell(row=..., column=..., value=...) bình thường, KHÔNG
    set .number_format cho từng ô nữa."""
    for col in range(1, so_cot + 1):
        sh.column_dimensions[get_column_letter(col)].number_format = "@"
def mo_workbook_doc_an_toan(path, *, data_only: bool = False) -> Workbook:
    """Mở workbook để ĐỌC, ưu tiên read_only=True (nhanh, ít RAM hơn hẳn
    với file lớn — đây là 1 phần lý do nạp dữ liệu nhanh hơn sau nâng cấp).
    BUG THẬT gặp phải sau khi đổi sang read_only=True: một số file .xlsx
    — đặc biệt file KHÔNG được tạo trực tiếp bằng Excel (export từ CMS/ERP,
    qua thư viện khác...) — THIẾU thẻ <dimension> hợp lệ trong XML của
    sheet. Ở chế độ read_only, openpyxl TIN vào thẻ này để trả về
    max_row/max_column — thiếu thẻ đó thì trả về None thay vì số, khiến
    MỌI so sánh kiểu `sh.max_row < x` ném lỗi "'<' not supported between
    instances of 'NoneType' and 'int'" (đúng lỗi anh gặp khi nạp file dữ
    liệu CMS export thật). Ở chế độ THƯỜNG (không read_only), openpyxl
    luôn tự QUÉT ô thật để tính max_row/max_column nên không bao giờ trả
    về None, nhưng chậm hơn với file rất lớn.
    Hàm này lấy cả 2 cái lợi: thử read_only=True trước (nhanh) — nếu phát
    hiện sheet đầu tiên có max_row/max_column là None (dấu hiệu file
    thiếu <dimension>) thì TỰ ĐỘNG đóng lại và mở lại ở chế độ thường
    (chậm hơn nhưng luôn đúng, không bao giờ crash vì None)."""
    wb = load_workbook(path, read_only=True, data_only=data_only)
    sh0 = wb.worksheets[0] if wb.worksheets else None
    # BUG THẬT (file export từ PIM): thẻ <dimension ref="A1"/> GHI SAI -> chế
    # độ read_only tưởng sheet chỉ có 1 ô (max_row=max_column=1) dù thực tế
    # có hàng chục cột -> "thiếu cột sku". Gặp kích thước 1x1 cũng mở lại
    # kiểu thường (sheet 1 ô thật thì mở lại cũng rất nhanh).
    if sh0 is not None and (
        sh0.max_row is None or sh0.max_column is None or (sh0.max_row <= 1 and sh0.max_column <= 1)
    ):
        wb.close()
        wb = load_workbook(path, read_only=False, data_only=data_only)
    return wb
def xoa_du_lieu_sheet(wb_path: Path, ten_sheet: str, data_start_row: int) -> dict:
    """Xoá SẠCH dữ liệu (GIỮ NGUYÊN mọi dòng header phía trên `data_start_row`)
    của 1 sheet — dùng cho 2 nút "🗑️ Xóa dữ liệu IMPORT" / "🗑️ Xóa dữ liệu
    DATA SP": nhiều khi người dùng muốn dọn sạch dữ liệu của lần nạp trước
    (SKU cũ trong IMPORT, dòng CMS export cũ trong DATA SP) TRƯỚC khi nạp
    lô dữ liệu mới, thay vì phải nhớ chọn đúng "Ghi đè" mỗi lần nạp."""
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if ten_sheet not in wb.sheetnames:
        wb.close()
        return {"loi": f'Không tìm thấy sheet "{ten_sheet}" trong workspace.'}
    sh = wb[ten_sheet]
    dong_cuoi_hien_tai = dong_cuoi_toan_dong(sh, data_start_row)
    so_dong_xoa = 0
    if dong_cuoi_hien_tai >= data_start_row:
        so_dong_xoa = dong_cuoi_hien_tai - data_start_row + 1
        sh.delete_rows(data_start_row, so_dong_xoa)
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {"loi": None, "so_dong_xoa": so_dong_xoa}
# ============================================================================
# §10 TRÍCH DATA SP TỪ SHEET DỮ LIỆU THÔ — y hệt timSheetDuLieuTho_()/
#     trichDataSp_Silent_()
# ============================================================================
def _ghi_header_data_sp(sh: Worksheet) -> None:
    """LUÔN ghi lại header ở đúng dòng (CFG.sp_start_row - 1) mỗi lần
    trích — tự phục hồi nếu header từng bị xoá/ghi đè nhầm."""
    header = ["PRODUCTID", "PRODUCTCODE", "PRODUCTNAME", "PROPERTYID", "PROPERTYNAME", "PROPVALUE", "CATEGORYID"]
    bold = Font(bold=True)
    row = CFG.sp_start_row - 1
    for col, h in enumerate(header, start=1):
        c = sh.cell(row=row, column=col, value=h)
        c.font = bold
        c.number_format = "@"
    sh.freeze_panes = f"A{CFG.sp_start_row}"
def tim_sheet_du_lieu_tho(wb: Workbook) -> Optional[Worksheet]:
    bo_qua = {x.upper() for x in (
        CFG.sheet_nhap, CFG.sheet_chon, CFG.sheet_cau_hinh, CFG.sheet_log, CFG.sheet_detail,
        CFG_TF_sheet_tskt_mapping, CFG_TF_sheet_filter_mapping, *CFG_sheet_tien_ich,
    )}
    for x in CFG.sheet_data_sp + CFG.sheet_data_pim + CFG.sheet_mapping:
        bo_qua.add(x.upper())
    ung_vien = []
    for sh in wb.worksheets:
        ten = sh.title.upper()
        if ten in bo_qua:
            continue
        if ten.startswith(CFG.tskth_prefix.upper()):
            continue
        if sh.max_column < 1:
            continue
        header = sheet_display_values(sh, 1, 1, sh.max_column)[0]
        co_prop_id = tim_cot_theo_ten(header, "PROPERTYID") > 0
        co_prop_value = tim_cot_theo_ten(header, "PROPVALUE") > 0
        co_ma_sp = tim_cot_theo_ten(header, "PRODUCTCODE") > 0 or tim_cot_theo_ten(header, "PRODUCTID") > 0
        if co_prop_id and co_prop_value and co_ma_sp:
            ung_vien.append(sh)
    if len(ung_vien) == 1:
        return ung_vien[0]
    return None
def _doc_va_lam_sach_du_lieu_tho(sh_tho: Worksheet) -> dict:
    """Đọc + làm sạch dữ liệu từ 1 sheet thô (theo TÊN cột, không phụ thuộc
    thứ tự) — dùng CHUNG cho cả trich_data_sp_silent (RUN-ALL, tự động) lẫn
    trich_data_sp_tuong_tac (nút GUI, chủ động chạy bất kỳ lúc nào) — để
    không có 2 bản logic trích khác nhau dễ lệch nhau về sau."""
    last_col = sh_tho.max_column
    last_row = sh_tho.max_row
    # BẢO VỆ THÊM (dù nơi gọi hàm này giờ đã tự tránh trường hợp None qua
    # mo_workbook_doc_an_toan()): nếu vẫn lỡ nhận sheet mở kiểu read_only
    # từ 1 file thiếu thẻ <dimension>, max_column/max_row trả về None —
    # so sánh `None < 1` sẽ ném lỗi thay vì báo lỗi dễ hiểu. Coi None như
    # "không đọc được kích thước" và báo rõ nguyên nhân.
    if last_col is None or last_row is None:
        return {
            "loi": f'Sheet "{sh_tho.title}" không đọc được kích thước (file có thể bị lỗi định dạng). '
            "Thử mở file bằng Excel rồi \"Save As\" lại thành .xlsx mới trước khi nạp."
        }
    if last_col < 1 or last_row < 2:
        return {"loi": f'Sheet "{sh_tho.title}" không có dữ liệu (từ dòng 2).'}
    header = sheet_display_values(sh_tho, 1, 1, last_col)[0]
    c_product_id = tim_cot_theo_ten(header, "PRODUCTID")
    c_product_code = tim_cot_theo_ten(header, "PRODUCTCODE")
    c_product_name = tim_cot_theo_ten(header, "PRODUCTNAME")
    c_prop_id = tim_cot_theo_ten(header, "PROPERTYID")
    c_prop_name = tim_cot_theo_ten(header, "PROPERTYNAME")
    c_prop_value = tim_cot_theo_ten(header, "PROPVALUE")
    c_cate_id = tim_cot_theo_ten(header, "CATEGORYID")
    thieu = []
    for nm, c in (("PRODUCTID", c_product_id), ("PRODUCTCODE", c_product_code), ("PRODUCTNAME", c_product_name),
                  ("PROPERTYID", c_prop_id), ("PROPERTYNAME", c_prop_name), ("PROPVALUE", c_prop_value),
                  ("CATEGORYID", c_cate_id)):
        if not c:
            thieu.append(nm)
    if thieu:
        return {"loi": f'Sheet "{sh_tho.title}" thiếu cột: {", ".join(thieu)}'}
    max_can_doc = max(c_product_id, c_product_code, c_product_name, c_prop_id, c_prop_name, c_prop_value, c_cate_id)
    raw = sheet_display_values(sh_tho, 2, last_row, max_can_doc)
    ket_qua = []
    for r in raw:
        product_id = chuan_hoa_id(r[c_product_id - 1])
        product_code = chuan_hoa_code(r[c_product_code - 1])
        product_name = chuan_hoa_key(r[c_product_name - 1])
        prop_id = chuan_hoa_id(r[c_prop_id - 1])
        prop_name = chuan_hoa_key(r[c_prop_name - 1])
        prop_value = chuan_hoa_key(r[c_prop_value - 1])
        cate_id = chuan_hoa_id(r[c_cate_id - 1])
        if not product_code:
            continue
        ket_qua.append([product_id, product_code, product_name, prop_id, prop_name, prop_value, cate_id])
    if not ket_qua:
        return {"loi": "Không trích được dòng nào hợp lệ (thiếu PRODUCTCODE)."}
    return {"loi": None, "ket_qua": ket_qua}
def trich_data_sp_silent(wb: Workbook, sh_tho: Worksheet) -> dict:
    kq_doc = _doc_va_lam_sach_du_lieu_tho(sh_tho)
    if kq_doc.get("loi"):
        return kq_doc
    ket_qua = kq_doc["ket_qua"]
    if CFG.sheet_data_sp[0] in wb.sheetnames:
        sh_dich = wb[CFG.sheet_data_sp[0]]
        dong_cuoi_cu = dong_cuoi(sh_dich, 2, CFG.sp_start_row)
        if dong_cuoi_cu >= CFG.sp_start_row:
            sh_dich.delete_rows(CFG.sp_start_row, dong_cuoi_cu - CFG.sp_start_row + 1)
    else:
        sh_dich = wb.create_sheet(CFG.sheet_data_sp[0])
    _ghi_header_data_sp(sh_dich)
    dat_dinh_dang_van_ban_cot(sh_dich, 7)
    for i, row in enumerate(ket_qua):
        row_idx = CFG.sp_start_row + i
        for col, v in enumerate(row, start=1):
            sh_dich.cell(row=row_idx, column=col, value=v)
    return {"loi": None, "so_dong": len(ket_qua)}
def trich_data_sp_tuong_tac(wb_path: Path, ten_sheet_tho: str, che_do: str) -> dict:
    """che_do: 'append' (NỐI TIẾP — giữ dữ liệu DATA SP cũ, thêm vào cuối)
    hoặc 'overwrite' (GHI ĐÈ — xoá hết dữ liệu DATA SP cũ, ghi mới).
    Dùng cho nút GUI "📥 Trích/Cập nhật DATA SP" — CHỦ ĐỘNG chạy được BẤT
    KỲ LÚC NÀO (khác trich_data_sp_silent chỉ chạy tự động trong RUN-ALL
    và CHỈ khi DATA SP đang trống hoàn toàn). Đúng quy trình thực tế:
    người dùng dán nhiều sheet dữ liệu thô riêng theo từng ngành hàng vào
    workspace theo thời gian, mỗi lần cần TRÍCH THÊM (nối tiếp) vào DATA
    SP đã có sẵn, không phải lúc nào cũng ghi đè từ đầu."""
    wb = load_workbook(wb_path)
    if ten_sheet_tho not in wb.sheetnames:
        return {"loi": f'Không tìm thấy sheet "{ten_sheet_tho}" trong workspace.'}
    sh_tho = wb[ten_sheet_tho]
    kq_doc = _doc_va_lam_sach_du_lieu_tho(sh_tho)
    if kq_doc.get("loi"):
        return kq_doc
    ket_qua = kq_doc["ket_qua"]
    if CFG.sheet_data_sp[0] in wb.sheetnames:
        sh_dich = wb[CFG.sheet_data_sp[0]]
        dong_cuoi_cu = dong_cuoi(sh_dich, 2, CFG.sp_start_row)
        if che_do == "overwrite":
            if dong_cuoi_cu >= CFG.sp_start_row:
                sh_dich.delete_rows(CFG.sp_start_row, dong_cuoi_cu - CFG.sp_start_row + 1)
            dong_bat_dau = CFG.sp_start_row
        else:  # append - noi tiep xuong duoi, giu nguyen du lieu cu
            dong_bat_dau = dong_cuoi_cu + 1 if dong_cuoi_cu >= CFG.sp_start_row else CFG.sp_start_row
    else:
        sh_dich = wb.create_sheet(CFG.sheet_data_sp[0])
        dong_bat_dau = CFG.sp_start_row
    _ghi_header_data_sp(sh_dich)
    dat_dinh_dang_van_ban_cot(sh_dich, 7)
    for i, row in enumerate(ket_qua):
        row_idx = dong_bat_dau + i
        for col, v in enumerate(row, start=1):
            sh_dich.cell(row=row_idx, column=col, value=v)
    wb.save(wb_path)
    wb.close()
    return {"loi": None, "so_dong": len(ket_qua), "dong_bat_dau": dong_bat_dau}
def nap_va_loc_nhieu_file(wb_path: Path, src_paths: List[str], che_do: str) -> dict:
    """NẠP NHIỀU FILE CMS EXPORT CÙNG LÚC (mỗi file thường là 1 ngành
    hàng) — nạp + lọc TUẦN TỰ từng file rồi ghi hết vào DATA SP, nhưng CHỈ
    mở + lưu workspace ĐÚNG 1 LẦN cho CẢ LÔ (không phải 1 lần mở/lưu cho
    mỗi file) — giữ đúng tinh thần tối ưu tốc độ: đọc thẳng từng file
    nguồn (không copy vào workspace), gộp kết quả, ghi 1 lần.
    `che_do` áp dụng cho CẢ LÔ, không phải riêng từng file:
      - 'overwrite': CHỈ xoá dữ liệu DATA SP cũ 1 LẦN DUY NHẤT, trước khi
        xử lý file ĐẦU TIÊN — các file sau trong CÙNG LÔ này đều nối tiếp
        vào sau, KHÔNG xoá lẫn nhau (nếu không, file thứ 2 sẽ xoá mất dữ
        liệu file thứ 1 vừa nạp, vô lý cho việc nạp nhiều file cùng lúc).
      - 'append': mọi file đều nối tiếp, giữ nguyên dữ liệu DATA SP cũ.
    KHÔNG dừng cả lô nếu 1 file bị lỗi (thiếu cột, mở không được, file
    trống...) — ghi nhận lỗi riêng cho file đó trong `chi_tiet`, TIẾP TỤC
    xử lý các file còn lại, rồi trả về báo cáo tổng hợp để người dùng thấy
    rõ file nào thành công / file nào lỗi và vì sao, thay vì cả lô bị
    chặn đứng chỉ vì 1 file có vấn đề."""
    if not src_paths:
        return {"loi": "Chưa chọn file nào."}
    try:
        wb_dich = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if CFG.sheet_data_sp[0] in wb_dich.sheetnames:
        sh_dich = wb_dich[CFG.sheet_data_sp[0]]
        dong_cuoi_cu = dong_cuoi(sh_dich, 2, CFG.sp_start_row)
        if che_do == "overwrite":
            if dong_cuoi_cu >= CFG.sp_start_row:
                sh_dich.delete_rows(CFG.sp_start_row, dong_cuoi_cu - CFG.sp_start_row + 1)
            dong_ghi_tiep = CFG.sp_start_row
        else:  # append
            dong_ghi_tiep = dong_cuoi_cu + 1 if dong_cuoi_cu >= CFG.sp_start_row else CFG.sp_start_row
    else:
        sh_dich = wb_dich.create_sheet(CFG.sheet_data_sp[0])
        dong_ghi_tiep = CFG.sp_start_row
    dong_bat_dau_dau_tien = dong_ghi_tiep
    _ghi_header_data_sp(sh_dich)
    dat_dinh_dang_van_ban_cot(sh_dich, 7)
    chi_tiet: list[tuple[str, Optional[int], Optional[str]]] = []  # (ten_file, so_dong, loi)
    tong_dong = 0
    so_file_thanh_cong = 0
    so_file_loi = 0
    for src_path in src_paths:
        ten_file = Path(src_path).name
        try:
            wb_src = mo_workbook_doc_an_toan(src_path, data_only=True)
        except Exception as e:  # noqa: BLE001
            chi_tiet.append((ten_file, None, f"Không mở được: {e}"))
            so_file_loi += 1
            continue
        sh_src = wb_src.worksheets[0]
        kq_doc = _doc_va_lam_sach_du_lieu_tho(sh_src)
        wb_src.close()
        if kq_doc.get("loi"):
            chi_tiet.append((ten_file, None, kq_doc["loi"]))
            so_file_loi += 1
            continue
        ket_qua = kq_doc["ket_qua"]
        for row in ket_qua:
            for col, v in enumerate(row, start=1):
                sh_dich.cell(row=dong_ghi_tiep, column=col, value=v)
            dong_ghi_tiep += 1
        tong_dong += len(ket_qua)
        so_file_thanh_cong += 1
        chi_tiet.append((ten_file, len(ket_qua), None))
    if so_file_thanh_cong == 0:
        wb_dich.close()
        return {
            "loi": "Không nạp được file nào:\n"
            + "\n".join(f"✖ {ten}: {loi}" for ten, _, loi in chi_tiet)
        }
    try:
        wb_dich.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb_dich.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb_dich.close()
    return {
        "loi": None,
        "so_dong": tong_dong,
        "dong_bat_dau": dong_bat_dau_dau_tien,
        "so_file_thanh_cong": so_file_thanh_cong,
        "so_file_loi": so_file_loi,
        "chi_tiet": chi_tiet,
    }
def nap_va_loc_du_lieu_tho(wb_path: Path, src_path: str, che_do: str) -> dict:
    """Nạp 1 FILE CMS export DUY NHẤT — nay chỉ là 1 lớp mỏng gọi
    nap_va_loc_nhieu_file() với danh sách đúng 1 file, để logic nạp/lọc
    CHỈ nằm ở 1 NƠI DUY NHẤT (tránh 2 bản dễ lệch nhau khi sửa về sau).
    Vẫn giữ được để mã nơi khác (nếu có) gọi trực tiếp 1 file không cần
    đổi cách dùng."""
    kq = nap_va_loc_nhieu_file(wb_path, [src_path], che_do)
    if kq.get("loi"):
        return kq
    if kq["so_file_loi"] > 0:
        # Chỉ có đúng 1 file mà lỗi -> trả lỗi thẳng (giữ nguyên hành vi
        # cũ của hàm 1-file: có lỗi thì "loi" khác None).
        _, _, loi_rieng = kq["chi_tiet"][0]
        return {"loi": loi_rieng}
    return {"loi": None, "so_dong": kq["so_dong"], "dong_bat_dau": kq["dong_bat_dau"]}
def nhap_import_tuong_tac(wb_path: Path, src_path: str, che_do: str) -> dict:
    """Nhập DANH SÁCH SKU (model_code/sku/variant_code/category_code) từ 1
    file Excel BÊN NGOÀI vào sheet IMPORT của workspace — tương tự
    trich_data_sp_tuong_tac() nhưng cho IMPORT thay vì DATA SP. Đọc cột
    theo TÊN header ở dòng 1 (model_code/sku/variant_code/category_code,
    không phân biệt hoa/thường, không phụ thuộc thứ tự cột nguồn — khớp
    đúng cấu trúc IMPORT thật: dòng 1 = mã cột, dòng 2 = tên tiếng Việt,
    dữ liệu từ dòng 3). che_do: 'append' (nối tiếp) | 'overwrite' (ghi đè).
    """
    try:
        wb_src = mo_workbook_doc_an_toan(src_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file nguồn: {e}"}
    sh_src = wb_src.worksheets[0]
    cols_src = doc_header_import(sh_src)
    col_map_src = {c.lower(): i for c, i in cols_src}
    c_sku = col_map_src.get("sku")
    if not c_sku:
        wb_src.close()
        return {"loi": 'File nguồn thiếu cột "sku" ở dòng 1 (header phải có tên cột "sku").'}
    c_model = col_map_src.get("model_code", 0)
    c_variant = col_map_src.get("variant_code", 0)
    c_cate = col_map_src.get("category_code", 0)
    last_row_src = dong_cuoi(sh_src, c_sku, CFG.data_start_row)
    if last_row_src < CFG.data_start_row:
        wb_src.close()
        return {"loi": f"Không có dữ liệu SKU nào từ dòng {CFG.data_start_row} trong file nguồn."}
    max_col_src = max(c_sku, c_model, c_variant, c_cate)
    vals = sheet_display_values(sh_src, CFG.data_start_row, last_row_src, max_col_src)
    wb_src.close()
    ket_qua = []  # [(model, sku, variant, cate), ...]
    for r in vals:
        sku = chuan_hoa_code(r[c_sku - 1])
        if not sku:
            continue
        model = chuan_hoa_key(r[c_model - 1]) if c_model else ""
        variant = chuan_hoa_key(r[c_variant - 1]) if c_variant else ""
        cate = chuan_hoa_id(r[c_cate - 1]) if c_cate else ""
        ket_qua.append((model, sku, variant, cate))
    if not ket_qua:
        return {"loi": "Không có dòng SKU hợp lệ nào để nhập."}
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    kq = _ghi_vao_import(wb, ket_qua, che_do)
    if kq.get("loi"):
        wb.close()
        return kq
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return kq
def _ghi_vao_import(wb: Workbook, ket_qua: List[Tuple[str, str, str, str]], che_do: str) -> dict:
    """Ghi [(model, sku, variant, cate), ...] vào sheet IMPORT (đã mở sẵn,
    KHÔNG lưu) — dùng chung cho nút 'Danh sách SKU (IMPORT)' và nút 'Nạp
    file export PIM'. Khớp cột đích theo TÊN mã ở dòng 1 của IMPORT, nên
    thứ tự cột IMPORT (model_code/sku/variant_code/category_code) và thứ tự
    cột file PIM (model_code/sku/category_code/variant_code) lệch nhau
    cũng không sao. Nếu IMPORT chưa có cột category_code -> tự thêm."""
    if CFG.sheet_nhap not in wb.sheetnames:
        return {"loi": f'Workspace thiếu sheet "{CFG.sheet_nhap}" — thử khởi động lại app để tự tạo lại.'}
    sh_dst = wb[CFG.sheet_nhap]
    cols_dst = doc_header_import(sh_dst)
    col_map_dst = {c.lower(): i for c, i in cols_dst}
    c_dst_sku = col_map_dst.get("sku")
    if not c_dst_sku:
        return {"loi": f'Sheet "{CFG.sheet_nhap}" trong workspace thiếu cột "sku" ở dòng 1.'}
    bold = Font(bold=True)
    for code, ten in (("model_code", "Mã model"), ("variant_code", "Mã biến thể"), ("category_code", "Mã danh mục PIM")):
        if code not in col_map_dst:
            col_moi = max(col_map_dst.values()) + 1
            for rr, vv in ((1, code), (2, ten)):
                c = sh_dst.cell(row=rr, column=col_moi, value=vv)
                c.font = bold
                c.number_format = "@"
            col_map_dst[code] = col_moi
    c_dst_model = col_map_dst.get("model_code", 0)
    c_dst_variant = col_map_dst.get("variant_code", 0)
    c_dst_cate = col_map_dst.get("category_code", 0)
    dong_cuoi_cu = dong_cuoi(sh_dst, c_dst_sku, CFG.data_start_row)
    if che_do == "overwrite":
        if dong_cuoi_cu >= CFG.data_start_row:
            sh_dst.delete_rows(CFG.data_start_row, dong_cuoi_cu - CFG.data_start_row + 1)
        dong_bat_dau = CFG.data_start_row
    else:  # append
        dong_bat_dau = dong_cuoi_cu + 1 if dong_cuoi_cu >= CFG.data_start_row else CFG.data_start_row
    dat_dinh_dang_van_ban_cot(sh_dst, max(c_dst_model, c_dst_sku, c_dst_variant, c_dst_cate))
    for i, (model, sku, variant, cate) in enumerate(ket_qua):
        row_idx = dong_bat_dau + i
        if c_dst_model:
            sh_dst.cell(row=row_idx, column=c_dst_model, value=model or None)
        sh_dst.cell(row=row_idx, column=c_dst_sku, value=sku)
        if c_dst_variant:
            sh_dst.cell(row=row_idx, column=c_dst_variant, value=variant or None)
        if c_dst_cate:
            sh_dst.cell(row=row_idx, column=c_dst_cate, value=cate or None)
    # Ghi đè mà danh sách mới NGẮN hơn cũ: delete_rows ở trên đã xoá hết dòng
    # cũ nên không còn sót. Nối tiếp: giữ nguyên.
    return {"loi": None, "so_dong": len(ket_qua), "dong_bat_dau": dong_bat_dau}
# ============================================================================
# §11 ĐIỀN 1 CATE — TSKT (text) + FILTER (số) — y hệt dienMotCateTsktFilter_()
# ============================================================================
def dien_mot_cate_tskt_filter(
    cache: Cache, sh: Worksheet, cate: str, skus: List[str],
    nl_map: Dict[str, NhapLieuRec], cfg: Optional[CategoryConfig], log_all: List[list],
    don_vi: Optional[Dict[str, str]] = None,
) -> dict:
    """`don_vi`: {mã TSKT: đơn vị} của RIÊNG cate này (đọc từ sheet ĐƠN VỊ)
    — chỉ áp cho giá trị TSKT (text) là SỐ TRƠN, KHÔNG BAO GIỜ áp cho FILTER."""
    don_vi = don_vi or {}
    cols = doc_header_import(sh)
    if not cols:
        return {"loi": "tab chưa có header"}
    attr_cols = loc_cot_thuoc_tinh(cols)
    if not attr_cols:
        return {"loi": "tab chưa có cột thuộc tính"}
    header_la = 0
    if cfg and cfg.tskt:
        cho_phep = set(cfg.tskt)
        loai_bo = [(c, i) for c, i in attr_cols if c not in cho_phep]
        for c, _ in loai_bo:
            log_all.append(["", cate, c, "Cột có trên tab nhưng KHÔNG có trong cấu hình category — bỏ qua, không điền"])
        header_la += len(loai_bo)
        attr_cols = [(c, i) for c, i in attr_cols if c in cho_phep]
        if not attr_cols:
            return {"loi": "không cột nào trên tab khớp cấu hình category"}
    tskt_map = cache.map_tskt_theo_cate.get(cate)
    filter_map = cache.map_filter_theo_cate.get(cate)
    if not tskt_map and not filter_map:
        return {"loi": "cate không có trong cả 2 bảng TSKT lẫn FILTER"}
    col_cua_code = {c.lower(): i for c, i in cols}
    header_set = {c for c, _ in attr_cols}
    sku_set = set(skus)
    bucket_tskt: Dict[Tuple[str, str], List[str]] = defaultdict(list)
    bucket_filter: Dict[Tuple[str, str], List[str]] = defaultdict(list)
    unmatched_filter = 0
    unmatched_seen_filter: set = set()
    for r in cache.sp_vals:
        sku = chuan_hoa_code(r[CFG.sp_sku - 1])
        if not sku or sku not in sku_set:
            continue
        prop_id = chuan_hoa_id(r[CFG.sp_prop_id - 1])
        prop_value = chuan_hoa_key(r[CFG.sp_prop_value - 1])
        if not prop_id or not prop_value:
            continue
        if tskt_map and prop_id in tskt_map:
            tskt_code = tskt_map[prop_id]
            if tskt_code in header_set:
                key = (sku, tskt_code)
                if prop_value not in bucket_tskt[key]:
                    bucket_tskt[key].append(prop_value)
        if filter_map and prop_id in filter_map:
            filter_code = filter_map[prop_id]
            if filter_code in header_set:
                if filter_code in cache.la_select:
                    key_opt = (filter_code, prop_value.lower())
                    if key_opt in cache.option_map:
                        gia_tri = cache.option_map[key_opt]
                        key = (sku, filter_code)
                        if gia_tri not in bucket_filter[key]:
                            bucket_filter[key].append(gia_tri)
                    else:
                        uk = (filter_code, prop_value)
                        if uk not in unmatched_seen_filter:
                            unmatched_seen_filter.add(uk)
                            unmatched_filter += 1
                            log_all.append([sku, cate, filter_code,
                                             f'FILTER: giá trị CMS "{prop_value}" chưa có option trong DATA PIM — ô để trống'])
                else:
                    uk2 = (filter_code, "(không có trong DATA PIM)")
                    if uk2 not in unmatched_seen_filter:
                        unmatched_seen_filter.add(uk2)
                        unmatched_filter += 1
                        log_all.append([sku, cate, filter_code,
                                         "FILTER: mã không có trong DATA PIM (chưa từng có option nào) — ô để trống"])
    last_col = max(i for _, i in cols)
    cu = dong_cuoi_moi_cot(sh, cols)
    if cu >= CFG.data_start_row:
        for row in range(CFG.data_start_row, cu + 1):
            for col in range(1, last_col + 1):
                # SỬA BUG: sh.cell(..., value=None) KHÔNG xoá được ô (openpyxl
                # bỏ qua value=None) -> dòng/ô cũ của lần chạy trước bị sót lại
                # và bị xuất kèm. Phải gán .value = None trực tiếp.
                sh.cell(row=row, column=col).value = None
    dat_dinh_dang_van_ban_cot(sh, last_col)
    for i, sku in enumerate(skus):
        row_idx = CFG.data_start_row + i
        nl = nl_map.get(sku)
        if "sku" in col_cua_code:
            sh.cell(row=row_idx, column=col_cua_code["sku"], value=sku)
        if nl and nl.model and "model_code" in col_cua_code:
            sh.cell(row=row_idx, column=col_cua_code["model_code"], value=nl.model)
        if nl and nl.variant and "variant_code" in col_cua_code:
            sh.cell(row=row_idx, column=col_cua_code["variant_code"], value=nl.variant)
        for code, col in attr_cols:
            arr_t = bucket_tskt.get((sku, code))
            arr_f = bucket_filter.get((sku, code))
            if arr_t:
                gia_tri_t = CFG.value_separator.join(arr_t)
                if code in don_vi:
                    gia_tri_t = them_don_vi(gia_tri_t, don_vi[code])
                sh.cell(row=row_idx, column=col, value=gia_tri_t)
            elif arr_f:
                sh.cell(row=row_idx, column=col, value=CFG.value_separator_filter.join(arr_f))
    so_o = len(bucket_tskt) + len(bucket_filter)
    return {"loi": None, "so_o": so_o, "unmatched_filter": unmatched_filter, "header_la": header_la}
# ============================================================================
# §12 XUẤT FILE — TÁCH MODEL / BIẾN THỂ — y hệt taoBlobXlsx_()/taoZipTachModel_()
# ============================================================================
def _tao_blob_xlsx(sh: Worksheet, cate: str, che_pham_vi: Optional[str]) -> dict:
    """che_pham_vi: None = gộp cả model+variant | 'model' | 'variant'."""
    cols = doc_header_import(sh)
    if not cols:
        return {"loi": "tab chưa có header"}
    last_row = dong_cuoi_moi_cot(sh, cols)
    if last_row < CFG.data_start_row:
        return {"loi": 'tab chưa có dữ liệu — chạy điền TSKT/FILTER trước'}
    bot = {x.lower() for x in CFG.bot_cot_khi_xuat}
    giu = [(c, i) for c, i in cols if c.lower() not in bot]
    if not giu:
        return {"loi": "không còn cột để xuất"}
    max_col = max(i for _, i in cols)
    all_vals = sheet_display_values(sh, 1, last_row, max_col)
    col_model = col_variant = 0
    if che_pham_vi:
        for c, i in cols:
            cc = c.lower()
            if cc == "model_code":
                col_model = i
            if cc == "variant_code":
                col_variant = i
        if not col_variant:
            return {"loi": "tab không có cột variant_code — không tách được Model/Biến thể"}
    out_rows = []
    out_rows.append([all_vals[CFG.header_row - 1][i - 1] for _, i in giu])
    out_rows.append([all_vals[CFG.name_row - 1][i - 1] for _, i in giu])
    co_dong = False
    for r in range(CFG.data_start_row - 1, last_row):
        if che_pham_vi:
            co_variant = chuan_hoa_key(all_vals[r][col_variant - 1]) != ""
            if che_pham_vi == "model" and co_variant:
                continue
            if che_pham_vi == "variant" and not co_variant:
                continue
        out_rows.append([all_vals[r][i - 1] for _, i in giu])
        co_dong = True
    if che_pham_vi and not co_dong:
        return {"loi": None, "wb_bytes": None}
    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "Export Product Template"
    dat_dinh_dang_van_ban_cot(ws, len(giu))
    for r_idx, row in enumerate(out_rows, start=1):
        for c_idx, v in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=v)
    ws.freeze_panes = "A3"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    nhan = sh.title or cate or "EXPORT"
    hau = "_MODEL" if che_pham_vi == "model" else ("_BIENTHE" if che_pham_vi == "variant" else "")
    ten = f"{CFG.file_prefix}{nhan}{hau}_{stamp}.xlsx"
    buf = io.BytesIO()
    wb_out.save(buf)
    return {"loi": None, "wb_bytes": buf.getvalue(), "ten": ten}
def tao_zip_tach_model(
    wb: Workbook, cache: Cache, cates: List[str], out_dir: Path,
    them_file: Optional[List[Tuple[str, bytes]]] = None,
) -> dict:
    """`them_file`: file phụ bỏ kèm vào .zip (vd _CANH_BAO_DOC_TRUOC.txt)."""
    tab_cua_cate = tim_tab_tskt(wb, cache)
    entries: List[Tuple[str, bytes]] = []
    loi: List[str] = []
    used_names: set = set()
    for cate in cates:
        tab = tab_cua_cate.get(cate)
        if not tab:
            loi.append(f"✖ {cate} — chưa có tab TSKT")
            continue
        for che_pham_vi in ("model", "variant"):
            kq = _tao_blob_xlsx(tab, cate, che_pham_vi)
            if kq.get("loi"):
                loi.append(f"✖ {cate} ({'Model' if che_pham_vi == 'model' else 'Biến thể'}) — {kq['loi']}")
                continue
            if kq.get("wb_bytes") is None:
                continue
            ten = kq["ten"]
            n = 1
            base = ten
            while ten in used_names:
                n += 1
                ten = base.replace(".xlsx", f"_{n}.xlsx")
            used_names.add(ten)
            entries.append((ten, kq["wb_bytes"]))
    if not entries:
        return {"loi": "Không tạo được file nào.\n" + "\n".join(loi)}
    if them_file:
        entries = list(them_file) + entries
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    ten_zip = f"{CFG.file_prefix}MODEL_BIENTHE_{stamp}.zip"
    out_dir.mkdir(parents=True, exist_ok=True)
    zip_path = out_dir / ten_zip
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for ten, data in entries:
            zf.writestr(ten, data)
    return {"loi": None, "so_file": len(entries), "zip_path": str(zip_path), "canh_bao": loi}
# ============================================================================
# §12b ĐƠN VỊ KÍCH THƯỚC / KHỐI LƯỢNG — thêm hàng loạt cm / mm / kg / g / inch
# ============================================================================
# Nguyên tắc (đã chốt với người dùng):
#   - CHỈ thêm đơn vị vào ô là SỐ TRƠN ("9" -> "9 kg", "6.95" -> "6.95 cm").
#     Ô đã có chữ ("13 kg", "Không", "Ngang 122 cm - ...") GIỮ NGUYÊN.
#   - KHÔNG BAO GIỜ đụng cột FILTER (FILTER là mã số option của PIM).
#   - Lựa chọn lưu theo (mã ngành hàng, mã TSKT) ở sheet "ĐƠN VỊ" và TỰ ÁP
#     LẠI mỗi lần CHẠY TẤT CẢ (vì mỗi lần chạy tab TSKT được điền lại từ đầu).
DON_VI_LUA_CHON = ["", "cm", "mm", "m", "kg", "g", "inch"]
_SO_TRON_RE = re.compile(r"^\d+(?:[.,]\d+)?$")
_KT_CODE_RE = re.compile(
    r"(?:^|_)(deep|depth|high|height|horizontal|width|wide|length|long|thick|thickness|"
    r"mass|weight|size|dimension|dimensions|diameter|tall)(?:_|$)"
)
_KT_TEN_RE = re.compile(
    r"^(sâu|cao|ngang|dày|dài|rộng|nặng)\b|khối lượng|trọng lượng|kích thước|kích cỡ|đường kính",
    re.IGNORECASE,
)
COT_KHONG_PHAI_SPEC = {
    "model_code", "sku", "category_code", "variant_code", "family_code",
    "family_variant_code", "model_activated", "variant_activated",
}
def la_so_tron(v) -> bool:
    return bool(_SO_TRON_RE.fullmatch(chuan_hoa_key(v)))
def them_don_vi(v, dv: str):
    """'9' + 'kg' -> '9 kg'. Không phải số trơn / không có đơn vị -> trả nguyên."""
    s = chuan_hoa_key(v)
    dv = chuan_hoa_key(dv)
    if not dv or not s or not _SO_TRON_RE.fullmatch(s):
        return v
    return f"{s} {dv}"
_MANG_RE = re.compile(r"^\s*\[(.*)\]\s*$", re.S)
def lam_sach_gia_tri_pim(v, code: str = "") -> str:
    """File export PIM ghi giá trị dạng mảng: ["905", " 903", " 21"] ->
    bỏ ngoặc [] và dấu nháy, cắt khoảng trắng -> '905, 903, 21' (FILTER)
    hoặc 'a|b' (TSKT). Giá trị thường giữ nguyên."""
    s = chuan_hoa_key(v)
    m = _MANG_RE.match(s)
    if not m:
        return s
    ben_trong = m.group(1)
    items = re.findall(r'"((?:[^"\\]|\\.)*)"', ben_trong) or ben_trong.split(",")
    items = [chuan_hoa_key(x).strip("'") for x in items]
    items = [x for x in items if x]
    sep = CFG.value_separator_filter if la_cot_filter(code) else CFG.value_separator
    return sep.join(items)
def la_cot_filter(code: str) -> bool:
    return "filter" in code.lower()
def la_cot_kich_thuoc(code: str, ten: str) -> bool:
    """Tự nhận diện cột kích thước/khối lượng theo MÃ (deep/high/mass/...)
    hoặc TÊN tiếng Việt (Sâu/Cao/Ngang/Khối lượng/Kích thước...)."""
    if not code or la_cot_filter(code) or code.lower() in COT_KHONG_PHAI_SPEC:
        return False
    return bool(_KT_CODE_RE.search(code.lower()) or _KT_TEN_RE.search(chuan_hoa_key(ten)))
def goi_y_don_vi(code: str, ten: str) -> str:
    c = code.lower()
    t = chuan_hoa_key(ten).lower()
    if re.search(r"khối lượng|trọng lượng|^nặng", t):
        return "kg"
    if re.search(r"(?:^|_)(mass|weight)(?:_|$)", c) and not re.search(r"(?:^|_)size(?:_|$)", c):
        return "kg"
    if "screen" in c or "màn hình" in t:
        return "inch"
    if la_cot_kich_thuoc(code, ten):
        return "cm"
    return ""
def cate_tu_ten_tab(title: str) -> str:
    """'TSKT 2162 Loa' -> '2162'."""
    t = chuan_hoa_key(title)
    if not t.upper().startswith(CFG.tskth_prefix.upper() + " "):
        return ""
    phan = t[len(CFG.tskth_prefix):].strip().split()
    return chuan_hoa_id(phan[0]) if phan else ""
def doc_don_vi_day_du(wb: Workbook) -> Dict[Tuple[str, str], List[str]]:
    """{(cate, mã TSKT): [tên NH, tên TSKT, đơn vị]} từ sheet ĐƠN VỊ."""
    kq: Dict[Tuple[str, str], List[str]] = {}
    if CFG_sheet_don_vi not in wb.sheetnames:
        return kq
    for r in wb[CFG_sheet_don_vi].iter_rows(min_row=2, max_col=5, values_only=True):
        if not r:
            continue
        r = list(r) + [None] * (5 - len(r))
        cate = chuan_hoa_id("" if r[0] is None else str(ep_van_ban_an_toan(r[0])))
        code = chuan_hoa_key(r[2])
        dv = chuan_hoa_key(r[4])
        if cate and code and dv:
            kq[(cate, code)] = [chuan_hoa_key(r[1]), chuan_hoa_key(r[3]), dv]
    return kq
def doc_don_vi(wb: Workbook) -> Dict[str, Dict[str, str]]:
    out: Dict[str, Dict[str, str]] = defaultdict(dict)
    for (cate, code), v in doc_don_vi_day_du(wb).items():
        out[cate][code] = v[2]
    return out
def _tao_lai_sheet(wb: Workbook, ten: str, headers: List[str]) -> Worksheet:
    if ten in wb.sheetnames:
        vi_tri = wb.sheetnames.index(ten)
        del wb[ten]
        sh = wb.create_sheet(ten, vi_tri)
    else:
        sh = wb.create_sheet(ten)
    bold = Font(bold=True)
    for col, h in enumerate(headers, start=1):
        c = sh.cell(row=1, column=col, value=h)
        c.font = bold
        c.number_format = "@"
    sh.freeze_panes = "A2"
    dat_dinh_dang_van_ban_cot(sh, len(headers))
    return sh
def _ghi_sheet_don_vi(wb: Workbook, du_lieu: Dict[Tuple[str, str], List[str]]) -> None:
    sh = _tao_lai_sheet(wb, CFG_sheet_don_vi, ["MÃ NGÀNH HÀNG", "TÊN NGÀNH HÀNG", "MÃ TSKT", "TÊN TSKT", "ĐƠN VỊ"])
    for r_idx, ((cate, code), (ten_nh, ten_tskt, dv)) in enumerate(sorted(du_lieu.items()), start=2):
        for c_idx, v in enumerate([cate, ten_nh, code, ten_tskt, dv], start=1):
            sh.cell(row=r_idx, column=c_idx, value=v or None)
    for col, w in zip("ABCDE", (16, 24, 42, 34, 10)):
        sh.column_dimensions[col].width = w
def _quet_cot_kich_thuoc_wb(wb: Workbook) -> dict:
    cau_hinh = doc_cau_hinh(wb)
    da_luu = doc_don_vi_day_du(wb)
    tab_cua_cate: Dict[str, str] = {}
    for ten in wb.sheetnames:
        cid = cate_tu_ten_tab(ten)
        if cid and cid not in tab_cua_cate:
            tab_cua_cate[cid] = ten
    thu_tu = list(cau_hinh.keys()) + [c for c in tab_cua_cate if c not in cau_hinh]
    nganh: Dict[str, dict] = {}
    for cate in thu_tu:
        cfg = cau_hinh.get(cate)
        ten_tab = tab_cua_cate.get(cate)
        codes: List[str] = []
        ten_tu_tab: Dict[str, str] = {}
        rows: List[tuple] = []
        h1_idx: Dict[str, int] = {}
        if ten_tab:
            all_rows = list(wb[ten_tab].iter_rows(values_only=True))
            if all_rows:
                h1 = all_rows[0]
                h2 = all_rows[1] if len(all_rows) > 1 else ()
                for j, h in enumerate(h1):
                    code = chuan_hoa_key(h)
                    if code:
                        codes.append(code)
                        ten_tu_tab[code] = chuan_hoa_key(h2[j]) if j < len(h2) else ""
                rows = all_rows[CFG.data_start_row - 1:]
                h1_idx = {chuan_hoa_key(h): j for j, h in enumerate(h1) if chuan_hoa_key(h)}
        if cfg:
            for code in cfg.tskt:
                if code not in codes:
                    codes.append(code)
        cot = []
        for code in codes:
            if code.lower() in COT_KHONG_PHAI_SPEC or la_cot_filter(code):
                continue
            ten = (cfg.ten_tskt.get(code) if cfg else "") or ten_tu_tab.get(code, "")
            tong = so_tron = 0
            vi_du_so: List[str] = []
            vi_du_chu: List[str] = []
            if rows and code in h1_idx:
                j = h1_idx[code]
                for r in rows:
                    if j >= len(r) or r[j] in (None, ""):
                        continue
                    s = chuan_hoa_key(ep_van_ban_an_toan(r[j]))
                    if not s:
                        continue
                    tong += 1
                    if _SO_TRON_RE.fullmatch(s):
                        so_tron += 1
                        if len(vi_du_so) < 3 and s not in vi_du_so:
                            vi_du_so.append(s)
                    elif len(vi_du_chu) < 2 and s not in vi_du_chu:
                        vi_du_chu.append(s[:30])
            cot.append({
                "code": code, "ten": ten, "tong": tong, "so_tron": so_tron,
                "vi_du": ", ".join(vi_du_so + vi_du_chu),
                "la_kt": la_cot_kich_thuoc(code, ten),
                "goi_y": goi_y_don_vi(code, ten),
                "da_luu": da_luu.get((cate, code), ["", "", ""])[2],
            })
        ten_nh = (cfg.name if cfg else "") or (
            " ".join(chuan_hoa_key(ten_tab).split()[2:]) if ten_tab else "")
        nganh[cate] = {"ten": ten_nh, "tab": ten_tab, "cot": cot}
    return {"loi": None, "nganh": nganh}
def quet_cot_kich_thuoc(wb_path: Path) -> dict:
    """Quét workspace: mỗi ngành hàng (CẤU HÌNH CATEGORY + tab TSKT) -> danh
    sách cột TSKT kèm thống kê số ô SỐ TRƠN, ví dụ, gợi ý đơn vị, đơn vị
    đã lưu. Mở read_only cho nhanh (workspace có thể > 10MB)."""
    try:
        wb = load_workbook(wb_path, read_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    try:
        return _quet_cot_kich_thuoc_wb(wb)
    except TypeError:
        wb.close()
        wb = load_workbook(wb_path)  # file thiếu <dimension> -> mở kiểu thường
        return _quet_cot_kich_thuoc_wb(wb)
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass
def luu_va_ap_don_vi(
    wb_path: Path, chon: Dict[str, Dict[str, str]],
    ten_cate: Dict[str, str], ten_tskt: Dict[Tuple[str, str], str],
) -> dict:
    """Lưu lựa chọn đơn vị vào sheet ĐƠN VỊ (đơn vị rỗng = xoá lựa chọn) rồi
    ÁP NGAY vào tab TSKT đang có (chỉ ô số trơn)."""
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    du_lieu = doc_don_vi_day_du(wb)
    for cate, m in chon.items():
        for code, dv in m.items():
            dv = chuan_hoa_key(dv)
            if dv:
                du_lieu[(cate, code)] = [ten_cate.get(cate, ""), ten_tskt.get((cate, code), ""), dv]
            else:
                du_lieu.pop((cate, code), None)
    _ghi_sheet_don_vi(wb, du_lieu)
    tab_cua_cate = {}
    for ten in wb.sheetnames:
        cid = cate_tu_ten_tab(ten)
        if cid and cid not in tab_cua_cate:
            tab_cua_cate[cid] = wb[ten]
    so_o: Dict[str, int] = {}
    for cate, m in chon.items():
        sh = tab_cua_cate.get(cate)
        if sh is None:
            continue
        col_cua = {c: i for c, i in doc_header_import(sh)}
        dem = 0
        for code, dv in m.items():
            dv = chuan_hoa_key(dv)
            if not dv or code not in col_cua or la_cot_filter(code):
                continue
            col = col_cua[code]
            for r in range(CFG.data_start_row, (sh.max_row or 0) + 1):
                cell = sh.cell(row=r, column=col)
                if cell.value is None:
                    continue
                s = str(ep_van_ban_an_toan(cell.value))
                moi = them_don_vi(s, dv)
                if moi != s:
                    cell.value = moi
                    dem += 1
        so_o[cate] = dem
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {"loi": None, "so_o": so_o, "so_dong_luu": len(du_lieu)}
# ============================================================================
# §12c NẠP FILE EXPORT PIM — model/sku/variant/category -> IMPORT, spec -> SPEC PIM TẠM
# ============================================================================
HEADER_SPEC = ["SKU", "MODEL_CODE", "CATEGORY_CODE", "VARIANT_CODE", "MÃ TSKT", "TÊN TSKT",
               "GIÁ TRỊ PIM", "FILE NGUỒN", "THỜI GIAN NẠP"]
def nap_file_export_pim(wb_path: Path, src_path: str, che_do: str) -> dict:
    """File export PIM (dòng 1 = mã cột, dòng 2 = tên, dữ liệu từ dòng 3):
      - model_code / sku / variant_code / category_code -> sheet IMPORT
        (che_do 'overwrite' | 'append', khớp cột theo TÊN mã).
      - MỌI cột spec còn lại (giá trị khác rỗng) -> sheet 'SPEC PIM TẠM'
        dạng dọc (1 dòng = 1 SKU x 1 mã TSKT) — gộp được nhiều ngành hàng có
        bộ cột khác nhau. Nối tiếp: SKU nạp lại thì spec cũ của SKU đó bị
        thay bằng spec mới (không để lẫn giá trị cũ/mới)."""
    try:
        wb_src = mo_workbook_doc_an_toan(src_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file nguồn: {e}"}
    sh_src = wb_src.worksheets[0]
    cols = doc_header_import(sh_src)
    col_map = {c.lower(): i for c, i in cols}
    c_sku = col_map.get("sku")
    if not c_sku:
        wb_src.close()
        return {"loi": 'File nguồn thiếu cột "sku" ở dòng 1 — cần đúng dạng file export PIM (dòng 1 = mã cột).'}
    c_model = col_map.get("model_code", 0)
    c_variant = col_map.get("variant_code", 0)
    c_cate = col_map.get("category_code", 0)
    max_col = max(i for _, i in cols)
    ten_row = sheet_display_values(sh_src, CFG.name_row, CFG.name_row, max_col)
    ten_cot = ten_row[0] if ten_row else [""] * max_col
    cot_spec = [(c, i, chuan_hoa_key(ten_cot[i - 1]) if i - 1 < len(ten_cot) else "")
                for c, i in cols if c.lower() not in COT_KHONG_PHAI_SPEC]
    ds_import: List[Tuple[str, str, str, str]] = []
    spec_moi: List[list] = []
    da_thay: set = set()
    bo_qua_khong_sku = trung_sku = thieu_model = thieu_cate = so_o_spec = 0
    def _txt(r, i):
        if not i or i - 1 >= len(r) or r[i - 1] is None:
            return ""
        return chuan_hoa_key(str(ep_van_ban_an_toan(r[i - 1])))
    for r in sh_src.iter_rows(min_row=CFG.data_start_row, max_col=max_col, values_only=True):
        if not r or all(v in (None, "") for v in r):
            continue
        sku = chuan_hoa_code(_txt(r, c_sku))
        if not sku:
            bo_qua_khong_sku += 1
            continue
        if sku in da_thay:
            trung_sku += 1
            continue
        da_thay.add(sku)
        model = _txt(r, c_model)
        variant = _txt(r, c_variant)
        cate = chuan_hoa_id(_txt(r, c_cate))
        if not model:
            thieu_model += 1
        if not cate:
            thieu_cate += 1
        ds_import.append((model, sku, variant, cate))
        co_spec = False
        for code, i, ten in cot_spec:
            v = lam_sach_gia_tri_pim(_txt(r, i), code)
            if v:
                spec_moi.append([sku, model, cate, variant, code, ten, v])
                so_o_spec += 1
                co_spec = True
        if not co_spec:
            # vẫn ghi 1 dòng để biết SKU này CÓ trong file PIM (spec rỗng)
            spec_moi.append([sku, model, cate, variant, "", "", ""])
    wb_src.close()
    if not ds_import:
        return {"loi": f"Không có dòng nào có sku từ dòng {CFG.data_start_row} trong file nguồn."}
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    kq_imp = _ghi_vao_import(wb, ds_import, che_do)
    if kq_imp.get("loi"):
        wb.close()
        return kq_imp
    giu_lai: List[list] = []
    if che_do == "append" and CFG_sheet_spec_pim in wb.sheetnames:
        for r in wb[CFG_sheet_spec_pim].iter_rows(min_row=2, max_col=len(HEADER_SPEC), values_only=True):
            if not r or r[0] in (None, ""):
                continue
            if chuan_hoa_code(str(r[0])) in da_thay:
                continue
            giu_lai.append(list(r))
    sh = _tao_lai_sheet(wb, CFG_sheet_spec_pim, HEADER_SPEC)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ten_file = Path(src_path).name
    r_idx = 2
    for row in giu_lai:
        for c_idx, v in enumerate(row, start=1):
            sh.cell(row=r_idx, column=c_idx, value=v)
        r_idx += 1
    for row in spec_moi:
        for c_idx, v in enumerate(row + [ten_file, stamp], start=1):
            sh.cell(row=r_idx, column=c_idx, value=v if v != "" else None)
        r_idx += 1
    for col, w in zip("ABCDEFGHI", (16, 12, 14, 14, 40, 28, 40, 30, 20)):
        sh.column_dimensions[col].width = w
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {
        "loi": None, "so_sku": len(ds_import), "so_o_spec": so_o_spec, "so_cot_spec": len(cot_spec),
        "dong_bat_dau": kq_imp["dong_bat_dau"], "bo_qua_khong_sku": bo_qua_khong_sku,
        "trung_sku": trung_sku, "thieu_model": thieu_model, "thieu_cate": thieu_cate,
        "so_dong_spec_giu_lai": len(giu_lai),
    }
def doc_spec_pim(wb: Workbook) -> Tuple[Optional[Dict[str, Dict[str, str]]], Dict[str, Tuple[str, str]]]:
    """-> (spec {sku: {mã: giá trị}}, info {sku: (model_code, category_code)}).
    spec = None nếu CHƯA nạp file export PIM nào (sheet không có / trống)."""
    if CFG_sheet_spec_pim not in wb.sheetnames:
        return None, {}
    spec: Dict[str, Dict[str, str]] = {}
    info: Dict[str, Tuple[str, str]] = {}
    for r in wb[CFG_sheet_spec_pim].iter_rows(min_row=2, max_col=7, values_only=True):
        if not r or r[0] in (None, ""):
            continue
        r = list(r) + [None] * (7 - len(r))
        sku = chuan_hoa_code(str(r[0]))
        spec.setdefault(sku, {})
        info.setdefault(sku, (chuan_hoa_key(r[1]), chuan_hoa_id("" if r[2] is None else str(r[2]))))
        code = chuan_hoa_key(r[4])
        if code:
            spec[sku][code] = lam_sach_gia_tri_pim(r[6], code)
    if not spec:
        return None, {}
    return spec, info
# ============================================================================
# §12d ĐỐI CHIẾU SPEC PIM + CẢNH BÁO TRƯỚC KHI XUẤT
# ============================================================================
_DV_DINH_RE = re.compile(r"(\d)\s+(cm|mm|m|kg|g|inch)\b")
def _tap_gia_tri(v: str, code: str, cache: Cache) -> frozenset:
    """Chuẩn hoá 1 giá trị để so sánh: không phân biệt hoa/thường, khoảng
    trắng thừa, thứ tự các phần nối '|'. FILTER: so theo TẬP mã option; nếu
    PIM xuất ra dạng chữ thì tự tra sang mã qua DATA PIM rồi mới so."""
    s = _WS_RE.sub(" ", lam_sach_gia_tri_pim(v, code))
    if not s:
        return frozenset()
    if la_cot_filter(code):
        out = set()
        for t in re.split(r"[,|]", s):
            t = t.strip()
            if not t:
                continue
            tid = chuan_hoa_id(t)
            if not re.fullmatch(r"\d+", tid):
                tid = cache.option_map.get((code, t.lower()), tid)
            out.add(tid.lower())
        return frozenset(out)
    return frozenset(_DV_DINH_RE.sub(r"\1\2", p.strip().lower()) for p in s.split("|") if p.strip())
def _doc_du_lieu_tab(sh: Worksheet) -> dict:
    cols = doc_header_import(sh)
    ten = {}
    if cols:
        h2 = sheet_display_values(sh, CFG.name_row, CFG.name_row, max(i for _, i in cols))[0]
        ten = {c: chuan_hoa_key(h2[i - 1]) for c, i in cols}
    last = dong_cuoi_moi_cot(sh, cols) if cols else 0
    rows = []
    if cols and last >= CFG.data_start_row:
        max_col = max(i for _, i in cols)
        for r in sh.iter_rows(min_row=CFG.data_start_row, max_row=last, max_col=max_col, values_only=True):
            vals = {c: ("" if r[i - 1] is None else chuan_hoa_key(str(r[i - 1]))) for c, i in cols}
            sku = chuan_hoa_code(vals.get("sku", ""))
            if not sku:
                continue
            rows.append({"sku": sku, "model": vals.get("model_code", ""),
                         "variant": vals.get("variant_code", ""), "vals": vals})
    attr = [c for c, _ in cols if c.lower() not in COT_KHONG_PHAI_SPEC]
    return {"title": sh.title, "attr": attr, "ten": ten, "rows": rows}
def doi_chieu_spec(du_lieu_tab: Dict[str, dict], spec: Dict[str, Dict[str, str]], cache: Cache) -> Tuple[List[list], dict]:
    rows_out: List[list] = []
    st = {"khac": 0, "tool_trong": 0, "pim_trong": 0, "sku_doi_chieu": 0}
    for cate, d in du_lieu_tab.items():
        for row in d["rows"]:
            sku = row["sku"]
            if sku not in spec:
                continue
            st["sku_doi_chieu"] += 1
            cu = spec[sku]
            for code in d["attr"]:
                moi = row["vals"].get(code, "")
                old = cu.get(code, "")
                if not moi and not old:
                    continue
                if _tap_gia_tri(moi, code, cache) == _tap_gia_tri(old, code, cache):
                    continue
                if moi and old and _chi_khac_don_vi(moi, old):
                    st["chi_don_vi"] = st.get("chi_don_vi", 0) + 1
                    continue  # chỉ thêm đơn vị — đúng, không liệt kê
                if moi and old:
                    loai = "KHÁC GIÁ TRỊ"
                    st["khac"] += 1
                elif old:
                    loai = "TOOL ĐỂ TRỐNG — PIM ĐANG CÓ"
                    st["tool_trong"] += 1
                else:
                    st["pim_trong"] += 1  # PIM trống, tool thêm mới — chỉ đếm, không liệt kê
                    continue
                rows_out.append([sku, row["model"], cate, code, d["ten"].get(code, ""), old, moi, loai])
    return rows_out, st
def tong_hop_canh_bao(
    du_lieu_tab: Dict[str, dict], nl_map: Dict[str, NhapLieuRec],
    spec: Optional[Dict[str, Dict[str, str]]], spec_info: Dict[str, Tuple[str, str]],
    dc_stat: Optional[dict],
) -> Tuple[List[list], List[str]]:
    """-> (dòng chi tiết cho sheet CẢNH BÁO XUẤT, các dòng tóm tắt)."""
    rows: List[list] = []
    tom: List[str] = []
    thieu_model: List[str] = []
    thieu_cate: List[str] = []
    lech_model: List[str] = []
    khong_co_pim: List[str] = []
    for cate, d in du_lieu_tab.items():
        for row in d["rows"]:
            sku = row["sku"]
            nl = nl_map.get(sku)
            model = (nl.model if nl else "") or row["model"]
            cate_pim = nl.cate_pim if nl else ""
            if not model:
                thieu_model.append(sku)
                rows.append(["CAO", "Thiếu model_code", sku, cate, "", "IMPORT chưa có Mã model cho SKU này"])
            if not cate_pim:
                thieu_cate.append(sku)
                rows.append(["CAO", "Thiếu category_code", sku, cate, "", "IMPORT chưa có Mã danh mục PIM cho SKU này"])
            if spec is not None:
                if sku not in spec:
                    khong_co_pim.append(sku)
                    rows.append(["TRUNG BÌNH", "SKU không có trong file PIM đã nạp", sku, cate, "",
                                 "Không đối chiếu được spec cho SKU này"])
                else:
                    m_pim, c_pim = spec_info.get(sku, ("", ""))
                    if m_pim and model and m_pim != model:
                        lech_model.append(sku)
                        rows.append(["CAO", "model_code khác file PIM", sku, cate, "",
                                     f"IMPORT: {model} — file PIM: {m_pim}"])
                    if c_pim and cate_pim and c_pim != cate_pim:
                        rows.append(["CAO", "category_code khác file PIM", sku, cate, "",
                                     f"IMPORT: {cate_pim} — file PIM: {c_pim}"])
    if thieu_model:
        tom.append(f"• {len(thieu_model)} SKU THIẾU model_code (vd: {', '.join(thieu_model[:3])})")
    if thieu_cate:
        tom.append(f"• {len(thieu_cate)} SKU THIẾU category_code (vd: {', '.join(thieu_cate[:3])})")
    if lech_model:
        tom.append(f"• {len(lech_model)} SKU có model_code trong IMPORT KHÁC file PIM")
    # --- kích thước / khối lượng còn số trơn ---
    cot_so_tron: List[Tuple[str, str, int]] = []
    for cate, d in du_lieu_tab.items():
        for code in d["attr"]:
            if not la_cot_kich_thuoc(code, d["ten"].get(code, "")):
                continue
            ds = [r["sku"] for r in d["rows"] if la_so_tron(r["vals"].get(code, ""))]
            if not ds:
                continue
            vd = [r["vals"][code] for r in d["rows"] if la_so_tron(r["vals"].get(code, ""))][:3]
            cot_so_tron.append((cate, code, len(ds)))
            rows.append(["TRUNG BÌNH", "Kích thước/khối lượng chưa có đơn vị", f"{len(ds)} SKU", cate, code,
                         f"Còn số trơn, vd: {', '.join(vd)} — SKU: {', '.join(ds[:10])}"
                         + (" ..." if len(ds) > 10 else "")])
    if cot_so_tron:
        tong = sum(n for _, _, n in cot_so_tron)
        ds_txt = ", ".join(f"{code} ({n})" for _, code, n in sorted(cot_so_tron, key=lambda x: -x[2])[:5])
        tom.append(f"• {len(cot_so_tron)} cột kích thước/khối lượng còn {tong} ô SỐ TRƠN chưa có đơn vị: {ds_txt}"
                   + (" ..." if len(cot_so_tron) > 5 else ""))
    # --- đối chiếu spec ---
    if spec is None:
        rows.append(["THÔNG TIN", "Chưa nạp file export PIM", "", "", "",
                     "Bỏ qua đối chiếu spec + kiểm tra SKU — dùng nút '📦 Nạp file export PIM' nếu cần"])
        tom.append("• (Chưa nạp file export PIM — KHÔNG đối chiếu được spec cũ)")
    else:
        if dc_stat and (dc_stat["khac"] or dc_stat["tool_trong"]):
            rows.append(["CAO", "Khác spec PIM đã nạp", "", "", "",
                         f"{dc_stat['khac']} ô khác giá trị, {dc_stat['tool_trong']} ô tool để trống trong khi "
                         f"PIM đang có — chi tiết từng ô ở sheet \"{CFG_sheet_doi_chieu}\""])
            tom.append(f"• {dc_stat['khac']} ô KHÁC spec PIM, {dc_stat['tool_trong']} ô tool ĐỂ TRỐNG trong khi PIM "
                       f"đang có giá trị (xem sheet {CFG_sheet_doi_chieu})")
        if khong_co_pim:
            tom.append(f"• {len(khong_co_pim)} SKU KHÔNG có trong file PIM đã nạp (vd: {', '.join(khong_co_pim[:3])})")
    return rows, tom
def _ghi_sheet_doi_chieu(wb: Workbook, rows: List[list], co_spec: bool) -> None:
    if not co_spec:
        if CFG_sheet_doi_chieu in wb.sheetnames:
            del wb[CFG_sheet_doi_chieu]
        return
    sh = _tao_lai_sheet(wb, CFG_sheet_doi_chieu, ["SKU", "MODEL", "NGÀNH HÀNG (CMS)", "MÃ TSKT", "TÊN TSKT",
                                                   "GIÁ TRỊ PIM CŨ", "GIÁ TRỊ TOOL MỚI", "LOẠI"])
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, v in enumerate(row, start=1):
            sh.cell(row=r_idx, column=c_idx, value=v or None)
    for col, w in zip("ABCDEFGH", (16, 12, 14, 40, 28, 40, 40, 28)):
        sh.column_dimensions[col].width = w
def _ghi_sheet_canh_bao(wb: Workbook, rows: List[list]) -> None:
    sh = _tao_lai_sheet(wb, CFG_sheet_canh_bao, ["MỨC", "LOẠI CẢNH BÁO", "SKU", "NGÀNH HÀNG (CMS)", "MÃ TSKT", "CHI TIẾT"])
    thu_tu = {"CAO": 0, "TRUNG BÌNH": 1, "THÔNG TIN": 2}
    for r_idx, row in enumerate(sorted(rows, key=lambda x: thu_tu.get(x[0], 9)), start=2):
        for c_idx, v in enumerate(row, start=1):
            sh.cell(row=r_idx, column=c_idx, value=v or None)
    for col, w in zip("ABCDEF", (12, 36, 16, 16, 36, 80)):
        sh.column_dimensions[col].width = w
# ============================================================================
# §12e KIỂM TRA NHANH — dùng cho tab "🔍 Kiểm tra & Đối chiếu" (CHỈ ĐỌC)
# ============================================================================
# Không mở workspace ở chế độ ghi, không lưu gì -> nhanh (vài giây) và an
# toàn để bấm bao nhiêu lần cũng được. Không đọc DATA SP (sheet nặng nhất).
#   Giá trị MỚI : tab TSKT trong workspace  HOẶC  file đã xuất (.zip/.xlsx)
#   Spec PIM CŨ : sheet SPEC PIM TẠM        HOẶC  file export PIM chọn thẳng
#                 HOẶC không đối chiếu
# Khớp dòng: theo sku; file đã xuất không có cột sku -> khớp theo
# (model_code, variant_code).
TRANG_THAI_GIONG = "Giống"
TRANG_THAI_KHAC = "KHÁC GIÁ TRỊ"
TRANG_THAI_TOOL_TRONG = "TOOL ĐỂ TRỐNG — PIM ĐANG CÓ"
TRANG_THAI_PIM_TRONG = "PIM TRỐNG — TOOL THÊM MỚI"
def _mo_ro_nhanh(src) -> Workbook:
    """Mở read_only; file khai báo kích thước sai/thiếu -> mở kiểu thường."""
    wb = load_workbook(src, read_only=True, data_only=True)
    for sh in wb.worksheets[:1]:
        if sh.max_row is None or sh.max_column is None or (sh.max_row <= 1 and sh.max_column <= 1):
            wb.close()
            if hasattr(src, "seek"):
                src.seek(0)
            return load_workbook(src, data_only=True)
    return wb
def _bang_tu_rows(title: str, cate: str, rows: List[tuple]) -> dict:
    """rows: dòng 1 = mã cột, dòng 2 = tên, dữ liệu từ dòng 3."""
    kq = {"title": title, "cate": cate, "attr": [], "ten": {}, "rows": []}
    if not rows:
        return kq
    h1 = rows[0]
    h2 = rows[1] if len(rows) > 1 else ()
    cols: List[Tuple[str, int]] = []
    seen: set = set()
    for j, h in enumerate(h1):
        c = chuan_hoa_key(h)
        if c and c not in seen:
            seen.add(c)
            cols.append((c, j))
    kq["ten"] = {c: (chuan_hoa_key(h2[j]) if j < len(h2) else "") for c, j in cols}
    kq["attr"] = [c for c, _ in cols if c.lower() not in COT_KHONG_PHAI_SPEC]
    for r in rows[CFG.data_start_row - 1:]:
        if not r:
            continue
        vals = {}
        for c, j in cols:
            v = r[j] if j < len(r) else None
            vals[c] = "" if v is None else chuan_hoa_key(str(ep_van_ban_an_toan(v)))
        sku = chuan_hoa_code(vals.get("sku", ""))
        model = vals.get("model_code", "")
        variant = vals.get("variant_code", "")
        if not sku and not model and not variant and not any(vals.values()):
            continue
        kq["rows"].append({"sku": sku, "model": model, "variant": variant, "so_dong": len(kq["rows"]) + CFG.data_start_row,
                           "cate_pim": chuan_hoa_id(vals.get("category_code", "")), "vals": vals})
    return kq
def _ten_file_xuat_sang_title(ten_file: str) -> str:
    """'PIM_TSKT 2162 Loa_MODEL_20261005_0455.xlsx' -> 'TSKT 2162 Loa'."""
    t = Path(ten_file).stem
    if t.startswith(CFG.file_prefix):
        t = t[len(CFG.file_prefix):]
    t = re.sub(r"(_MODEL|_BIENTHE)?_\d{8}_\d{4}(_\d+)?$", "", t)
    return t
def doc_gia_tri_moi_tu_file_xuat(paths: List[str]) -> dict:
    """Đọc các file .xlsx đã xuất (hoặc .zip chứa chúng)."""
    bang: List[dict] = []
    loi: List[str] = []
    def _doc(ten: str, src) -> None:
        try:
            wb = _mo_ro_nhanh(src)
            rows = list(wb.worksheets[0].iter_rows(values_only=True))
            wb.close()
        except Exception as e:  # noqa: BLE001
            loi.append(f"{ten}: {e}")
            return
        title = _ten_file_xuat_sang_title(ten)
        bang.append(_bang_tu_rows(title, cate_tu_ten_tab(title) or title, rows))
    for p in paths:
        if p.lower().endswith(".zip"):
            with zipfile.ZipFile(p) as zf:
                for n in zf.namelist():
                    if n.lower().endswith(".xlsx"):
                        _doc(n, io.BytesIO(zf.read(n)))
        else:
            _doc(Path(p).name, p)
    # Gộp MODEL + BIẾN THỂ cùng 1 ngành hàng thành 1 bảng
    gop: Dict[str, dict] = {}
    for b in bang:
        g = gop.setdefault(b["cate"], {"title": b["title"], "cate": b["cate"], "attr": [], "ten": {}, "rows": []})
        for c in b["attr"]:
            if c not in g["attr"]:
                g["attr"].append(c)
        g["ten"].update({k: v for k, v in b["ten"].items() if v})
        g["rows"].extend(b["rows"])
    return {"loi": "\n".join(loi) if loi and not gop else None, "canh_bao_doc": loi, "bang": gop}
def doc_file_export_pim(src_path: str) -> dict:
    """Đọc file export PIM -> danh sách {sku, model, variant, cate, spec{mã: giá trị}}."""
    try:
        wb = mo_workbook_doc_an_toan(src_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file PIM: {e}"}
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    b = _bang_tu_rows(Path(src_path).name, "", rows)
    if not b["rows"]:
        return {"loi": f"File PIM không có dòng dữ liệu nào từ dòng {CFG.data_start_row}."}
    ds = []
    for r in b["rows"]:
        ds.append({"sku": r["sku"], "model": r["model"], "variant": r["variant"], "cate": r["cate_pim"],
                   "spec": {c: lam_sach_gia_tri_pim(r["vals"].get(c, ""), c) for c in b["attr"]
                            if lam_sach_gia_tri_pim(r["vals"].get(c, ""), c)}})
    return {"loi": None, "ds": ds, "ten": b["ten"]}
def _spec_tu_sheet(sh: Worksheet) -> List[dict]:
    theo_sku: Dict[str, dict] = {}
    for r in sh.iter_rows(min_row=2, max_col=7, values_only=True):
        if not r or r[0] in (None, ""):
            continue
        r = list(r) + [None] * (7 - len(r))
        sku = chuan_hoa_code(str(r[0]))
        e = theo_sku.setdefault(sku, {
            "sku": sku, "model": chuan_hoa_key(r[1]),
            "cate": chuan_hoa_id("" if r[2] is None else str(r[2])),
            "variant": chuan_hoa_key(r[3]), "spec": {},
        })
        code = chuan_hoa_key(r[4])
        if code:
            e["spec"][code] = lam_sach_gia_tri_pim(r[6], code)
    return list(theo_sku.values())
TRANG_THAI_DON_VI = "CHỈ THÊM ĐƠN VỊ"
_BO_DV_RE = re.compile(r"\s*(cm|mm|m|kg|g|inch)$", re.IGNORECASE)
def _chi_khac_don_vi(moi: str, old: str) -> bool:
    """PIM cũ '11.5' — tool mới '11.5 kg' -> chỉ là thêm đơn vị, không phải sai."""
    o, m = chuan_hoa_key(old), chuan_hoa_key(moi)
    if not o or not m or not la_so_tron(o) or not _BO_DV_RE.search(m):
        return False
    return _BO_DV_RE.sub("", m).replace(",", ".") == o.replace(",", ".")
def doc_du_lieu_kiem_tra(wb_path: Path, nguon_moi, nguon_cu) -> dict:
    """ĐỌC 1 LẦN (chỉ đọc): IMPORT, DATA PIM (tra option), tab TSKT/file đã
    xuất, spec PIM cũ, đơn vị đã lưu. Kết quả giữ trong bộ nhớ để tab Kiểm
    tra TÍNH LẠI TỨC THÌ mỗi khi sửa ô / chọn đơn vị (không đọc lại file).
    nguon_moi: 'workspace' | [file đã xuất .zip/.xlsx]
    nguon_cu  : 'workspace' | 'khong' | đường dẫn file export PIM."""
    try:
        wb = _mo_ro_nhanh(str(wb_path))
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    try:
        opt: Dict[Tuple[str, str], str] = {}
        opt_ten: Dict[Tuple[str, str], str] = {}        # (mã filter, mã option) -> tên option
        opt_ds: Dict[str, List[Tuple[str, str]]] = defaultdict(list)  # mã filter -> [(mã option, tên)]
        ten_filter: Dict[str, str] = {}
        if CFG.sheet_data_pim[0] in wb.sheetnames:
            for r in wb[CFG.sheet_data_pim[0]].iter_rows(min_row=CFG.pim_start_row, max_col=CFG.pim_opt_value,
                                                         values_only=True):
                if not r or len(r) < CFG.pim_opt_value:
                    continue
                code = chuan_hoa_key(r[CFG.pim_code - 1])
                val = chuan_hoa_key(r[CFG.pim_opt_value - 1])
                if code and val:
                    oc = chuan_hoa_id("" if r[CFG.pim_opt_code - 1] is None else str(r[CFG.pim_opt_code - 1]))
                    opt.setdefault((code, val.lower()), oc)
                    if oc and (code, oc) not in opt_ten:
                        opt_ten[(code, oc)] = val
                        opt_ds[code].append((oc, val))
                    if code not in ten_filter:
                        ten_filter[code] = chuan_hoa_key(r[1])
        tra = type("TraOpt", (), {})()
        tra.option_map = opt
        nl_sku: Dict[str, NhapLieuRec] = {}
        nl_model: Dict[str, NhapLieuRec] = {}
        if CFG.sheet_nhap in wb.sheetnames:
            rows_imp = list(wb[CFG.sheet_nhap].iter_rows(values_only=True))
            if rows_imp:
                idx = {chuan_hoa_key(h).lower(): j for j, h in enumerate(rows_imp[0]) if chuan_hoa_key(h)}
                def _g(r, k):
                    j = idx.get(k)
                    return "" if j is None or j >= len(r) or r[j] is None else str(ep_van_ban_an_toan(r[j]))
                for r in rows_imp[CFG.data_start_row - 1:]:
                    sku = chuan_hoa_code(_g(r, "sku"))
                    if not sku or sku in nl_sku:
                        continue
                    rec = NhapLieuRec(sku, chuan_hoa_key(_g(r, "model_code")), chuan_hoa_key(_g(r, "variant_code")),
                                      chuan_hoa_id(_g(r, "category_code")))
                    nl_sku[sku] = rec
                    if rec.model:
                        nl_model.setdefault(rec.model, rec)
        don_vi_luu = doc_don_vi_day_du(wb)
        canh_bao_doc: List[str] = []
        if nguon_moi == "workspace":
            bang: Dict[str, dict] = {}
            for ten in wb.sheetnames:
                cid = cate_tu_ten_tab(ten)
                if not cid or cid in bang:
                    continue
                bang[cid] = _bang_tu_rows(ten, cid, list(wb[ten].iter_rows(values_only=True)))
            if not bang:
                return {"loi": "Workspace chưa có tab TSKT nào — bấm ① Điền dữ liệu trước, hoặc chọn file đã xuất."}
            nhan_moi = "Tab TSKT trong workspace"
        else:
            kq_x = doc_gia_tri_moi_tu_file_xuat(list(nguon_moi))
            if kq_x.get("loi") or not kq_x["bang"]:
                return {"loi": kq_x.get("loi") or "Không đọc được file đã xuất nào."}
            bang = kq_x["bang"]
            canh_bao_doc = kq_x["canh_bao_doc"]
            nhan_moi = ", ".join(Path(p).name for p in nguon_moi)
        ds_cu: Optional[List[dict]] = None
        ten_cu: Dict[str, str] = {}
        nhan_cu = "Không đối chiếu"
        if nguon_cu == "workspace":
            if CFG_sheet_spec_pim in wb.sheetnames:
                ds_cu = _spec_tu_sheet(wb[CFG_sheet_spec_pim]) or None
            nhan_cu = f'Sheet "{CFG_sheet_spec_pim}"' if ds_cu else f'Sheet "{CFG_sheet_spec_pim}" (TRỐNG — chưa nạp)'
        elif nguon_cu and nguon_cu != "khong":
            kq_p = doc_file_export_pim(nguon_cu)
            if kq_p.get("loi"):
                return {"loi": kq_p["loi"]}
            ds_cu = kq_p["ds"]
            ten_cu = kq_p["ten"]
            nhan_cu = Path(nguon_cu).name
    finally:
        wb.close()
    cu_sku: Dict[str, dict] = {}
    cu_mv: Dict[Tuple[str, str], dict] = {}
    for e in ds_cu or []:
        if e["sku"]:
            cu_sku.setdefault(e["sku"], e)
        if e["model"]:
            cu_mv.setdefault((e["model"], e["variant"]), e)
    return {"loi": None, "tra": tra, "nl_sku": nl_sku, "nl_model": nl_model, "bang": bang,
            "ds_cu": ds_cu, "cu_sku": cu_sku, "cu_mv": cu_mv, "ten_cu": ten_cu, "don_vi_luu": don_vi_luu,
            "opt_ten": opt_ten, "opt_ds": dict(opt_ds), "ten_filter": ten_filter,
            "nhan_moi": nhan_moi, "nhan_cu": nhan_cu, "canh_bao_doc": canh_bao_doc,
            "la_workspace": nguon_moi == "workspace"}
TRANG_THAI_BO_QUA = "ĐỂ TRỐNG — KHÔNG CẬP NHẬT"
GIA_TRI_RONG_MAC_DINH = ["Không", "Không có", "Đang cập nhật", "Hãng không công bố", "Không công bố",
                         "Chưa có thông tin", "Chưa cập nhật", "Không rõ", "-", "N/A"]
HD_GIU, HD_TRONG, HD_THAY = "giu", "trong", "thay"
def khoa_gia_tri_rong(v) -> str:
    """Chuẩn hoá để gộp biến thể: 'Không', 'không', 'Không.', 'Đang  cập nhật' (kể cả
    chữ Việt tổ hợp/dựng sẵn khác nhau) -> cùng 1 khoá."""
    s = unicodedata.normalize("NFC", _WS_RE.sub(" ", chuan_hoa_key(v)))
    return s.strip(" .;:").lower()
_KHOA_RONG_MAC_DINH = {khoa_gia_tri_rong(x) for x in GIA_TRI_RONG_MAC_DINH}
def ap_quy_tac_rong(v: str, code: str, rong: Optional[Dict[str, list]]) -> Tuple[str, Optional[str]]:
    """Áp quy tắc giá trị rỗng (KHÔNG áp cho FILTER). -> (giá trị mới, hành động đã áp | None)."""
    if not rong or not v or la_cot_filter(code) or code.lower() in COT_KHONG_PHAI_SPEC:
        return v, None
    q = rong.get(khoa_gia_tri_rong(v))
    if not q or q[0] == HD_GIU:
        return v, None
    if q[0] == HD_TRONG:
        return "", HD_TRONG
    return chuan_hoa_key(q[1] if len(q) > 1 else ""), HD_THAY
def duong_dan_cau_hinh_kt(wb_path: Path) -> Path:
    return Path(wb_path).with_name(Path(wb_path).stem + "_kiem_tra.json")
def doc_cau_hinh_kt(wb_path: Path) -> dict:
    """File cấu hình nhỏ cạnh workspace: sửa tay, đơn vị, quy tắc giá trị rỗng
    — ghi/đọc tức thì, tự áp lại sau mỗi lần ① map lại."""
    import json
    f = duong_dan_cau_hinh_kt(wb_path)
    try:
        d = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except Exception:  # noqa: BLE001
        d = {}
    return {
        "sua": {tuple(k.split("\t")): v for k, v in d.get("sua", {}).items() if k.count("\t") == 2},
        "don_vi": {tuple(k.split("\t")): v for k, v in d.get("don_vi", {}).items() if k.count("\t") == 1},
        "rong": {k: list(v) for k, v in d.get("rong", {}).items()},
    }
def ghi_cau_hinh_kt(wb_path: Path, sua: dict, don_vi: dict, rong: dict) -> None:
    import json
    d = {"sua": {"\t".join(k): v for k, v in sua.items()},
         "don_vi": {"\t".join(k): v for k, v in don_vi.items() if v},
         "rong": {k: v for k, v in rong.items() if v and v[0] != HD_GIU}}
    f = duong_dan_cau_hinh_kt(wb_path)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, f)
def tinh_kiem_tra(raw: dict, sua: Optional[Dict[Tuple[str, str, str], str]] = None,
                  don_vi: Optional[Dict[Tuple[str, str], str]] = None,
                  rong: Optional[Dict[str, list]] = None) -> dict:
    """TÍNH (không đọc file, < 1 giây). `sua`: {(cate, khoá dòng, mã): giá trị
    sửa tay}. `don_vi`: {(cate, mã): đơn vị} đang chọn ở tab Kiểm tra — áp
    cho ô SỐ TRƠN chưa sửa tay."""
    sua = sua or {}
    don_vi = don_vi or {}
    rong = rong or {}
    khoa_rong_quan_tam = _KHOA_RONG_MAC_DINH | set(rong)
    tk_rong: Dict[str, dict] = {}
    tra, nl_sku, nl_model = raw["tra"], raw["nl_sku"], raw["nl_model"]
    ds_cu, cu_sku, cu_mv, ten_cu = raw["ds_cu"], raw["cu_sku"], raw["cu_mv"], raw["ten_cu"]
    la_ws = raw["la_workspace"]
    chi_tiet: Dict[str, dict] = {}
    khac: List[list] = []
    canh_bao: List[list] = []
    st = defaultdict(int)
    so_tron_theo_cot: Dict[Tuple[str, str], List[str]] = defaultdict(list)
    tk_cot: Dict[Tuple[str, str], dict] = {}
    for cate, b in raw["bang"].items():
        if la_ws and nl_sku:
            if not any(r["sku"] in nl_sku for r in b["rows"]):
                canh_bao.append(["THÔNG TIN", "Bỏ qua tab không thuộc lô IMPORT hiện tại", b["title"], cate, "",
                                 f"{len(b['rows'])} dòng — không SKU nào có trong IMPORT, sẽ KHÔNG được xuất"])
                continue
        for code in b["attr"]:
            if not la_cot_filter(code):
                ten = b["ten"].get(code) or ten_cu.get(code, "")
                tk_cot[(cate, code)] = {"cate": cate, "code": code, "ten": ten, "tong": 0, "so_tron": 0,
                                        "vi_du": [], "la_kt": la_cot_kich_thuoc(code, ten),
                                        "goi_y": goi_y_don_vi(code, ten),
                                        "da_luu": raw["don_vi_luu"].get((cate, code), ["", "", ""])[2]}
        for row in b["rows"]:
            sku, model, variant = row["sku"], row["model"], row["variant"]
            nl = nl_sku.get(sku) if sku else nl_model.get(model)
            model_eff = model or (nl.model if nl else "")
            cate_pim = row["cate_pim"] or (nl.cate_pim if nl else "")
            khoa = sku or (f"{model}|{variant}" if (model or variant) else f"{b['title']}#{row['so_dong']}")
            nhan = sku or (f"model {model}" + (f" / biến thể {variant}" if variant else "") if (model or variant)
                           else f"{b['title']} dòng {row['so_dong']} (TRỐNG model)")
            if la_ws and nl_sku and sku and sku not in nl_sku:
                st["dong_sot"] += 1
                canh_bao.append(["THÔNG TIN", "Dòng trên tab không có trong IMPORT (sẽ tự bỏ khi xuất)", nhan, cate,
                                 "", "Dòng cũ sót lại — nút ③ tự bỏ, không xuất"])
                continue
            st["so_dong"] += 1
            if not model_eff:
                st["thieu_model"] += 1
                canh_bao.append(["CAO", "Thiếu model_code", nhan, cate, "", "Chưa có Mã model"])
            if not cate_pim:
                st["thieu_cate"] += 1
                canh_bao.append(["CAO", "Thiếu category_code", nhan, cate, "", "IMPORT chưa có Mã danh mục PIM"])
            e = None
            if ds_cu is not None:
                e = cu_sku.get(sku) if sku else None
                if e is None:
                    e = cu_mv.get((model_eff, variant))
                if e is None:
                    st["khong_co_pim"] += 1
                    canh_bao.append(["TRUNG BÌNH", "Không có trong file PIM", nhan, cate, "",
                                     "Không đối chiếu được spec"])
                else:
                    if e["model"] and model_eff and e["model"] != model_eff:
                        st["lech_model"] += 1
                        canh_bao.append(["CAO", "model_code khác file PIM", nhan, cate, "",
                                         f"Mới: {model_eff} — PIM: {e['model']}"])
                    if e["cate"] and cate_pim and e["cate"] != cate_pim:
                        st["lech_cate"] += 1
                        canh_bao.append(["CAO", "category_code khác file PIM", nhan, cate, "",
                                         f"Mới: {cate_pim} — PIM: {e['cate']}"])
            cu = e["spec"] if e else {}
            dong_ct = []
            for code in b["attr"]:
                goc = row["vals"].get(code, "")
                key = (cate, khoa, code)
                da_sua = key in sua
                hd = None
                if goc and not la_cot_filter(code):
                    kr = khoa_gia_tri_rong(goc)
                    if kr in khoa_rong_quan_tam:
                        x = tk_rong.setdefault(kr, {"khoa": kr, "so_o": 0, "dang": defaultdict(int), "cot": set()})
                        x["so_o"] += 1
                        x["dang"][goc] += 1
                        x["cot"].add(b["ten"].get(code) or code)
                if da_sua:
                    moi = chuan_hoa_key(sua[key])
                    st["da_sua"] += 1
                else:
                    moi, hd = ap_quy_tac_rong(goc, code, rong)
                    if hd:
                        st["rong_" + hd] += 1
                    dv = don_vi.get((cate, code))
                    if dv and not hd and not la_cot_filter(code):
                        moi2 = them_don_vi(moi, dv)
                        if moi2 != moi:
                            moi = moi2
                            st["them_dv"] += 1
                old = cu.get(code, "")
                ten = b["ten"].get(code) or ten_cu.get(code, "")
                t = tk_cot.get((cate, code))
                if t is not None and moi:
                    t["tong"] += 1
                    if la_so_tron(moi):
                        t["so_tron"] += 1
                        if len(t["vi_du"]) < 3 and moi not in t["vi_du"]:
                            t["vi_du"].append(moi)
                    elif len(t["vi_du"]) < 5 and moi[:25] not in t["vi_du"]:
                        t["vi_du"].append(moi[:25])
                if moi and la_cot_filter(code) and any(
                        not re.fullmatch(r"\d+", chuan_hoa_id(x.strip())) for x in re.split(r"[,|]", moi) if x.strip()):
                    st["filter_chu"] += 1
                    canh_bao.append(["CAO", "FILTER có giá trị không phải mã option", nhan, cate, code,
                                     f'"{moi}" — PIM chỉ nhận mã số; sửa thành mã đúng (xem DATA PIM) hoặc để trống'])
                if moi and t is not None and t["la_kt"] and la_so_tron(moi):
                    st["chua_don_vi"] += 1
                    so_tron_theo_cot[(cate, code)].append(nhan)
                if hd == HD_TRONG:
                    tt = TRANG_THAI_BO_QUA
                elif e is None:
                    tt = "" if not moi else "—"
                elif _tap_gia_tri(moi, code, tra) == _tap_gia_tri(old, code, tra):
                    tt = TRANG_THAI_GIONG if (moi or old) else ""
                elif moi and old:
                    tt = TRANG_THAI_DON_VI if _chi_khac_don_vi(moi, old) else TRANG_THAI_KHAC
                elif old:
                    tt = TRANG_THAI_TOOL_TRONG
                else:
                    tt = TRANG_THAI_PIM_TRONG
                if tt in (TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG, TRANG_THAI_PIM_TRONG, TRANG_THAI_DON_VI,
                          TRANG_THAI_BO_QUA):
                    st[tt] += 1
                if tt in (TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG, TRANG_THAI_PIM_TRONG, TRANG_THAI_DON_VI,
                          TRANG_THAI_BO_QUA) or da_sua:
                    khac.append([nhan, model_eff, cate, code, ten, old, moi, tt or TRANG_THAI_GIONG, khoa, goc, da_sua])
                if moi or old or da_sua:
                    dong_ct.append([code, ten, old, moi, tt, goc, da_sua])
            chi_tiet[khoa] = {"nhan": nhan, "sku": sku, "model": model_eff, "variant": variant,
                              "cate": cate, "tab": b["title"], "co_pim": e is not None, "dong": dong_ct}
    for (cate, code), ds in so_tron_theo_cot.items():
        canh_bao.append(["TRUNG BÌNH", "Kích thước/khối lượng chưa có đơn vị", f"{len(ds)} dòng", cate, code,
                         "VD: " + ", ".join(ds[:8]) + (" ..." if len(ds) > 8 else "")
                         + " — chọn đơn vị ở tab '📏 Đơn vị hàng loạt'"])
    for x in raw["canh_bao_doc"]:
        canh_bao.append(["THÔNG TIN", "Không đọc được file", "", "", "", x])
    thu_tu = {"CAO": 0, "TRUNG BÌNH": 1, "THÔNG TIN": 2}
    canh_bao.sort(key=lambda x: thu_tu.get(x[0], 9))
    return {
        "loi": None, "nhan_moi": raw["nhan_moi"], "nhan_cu": raw["nhan_cu"], "co_doi_chieu": ds_cu is not None,
        "stat": dict(st), "khac": khac, "canh_bao": canh_bao, "chi_tiet": chi_tiet,
        "so_cot_chua_dv": len(so_tron_theo_cot),
        "gia_tri_rong": sorted(
            [{"khoa": x["khoa"], "so_o": x["so_o"], "hien": max(x["dang"].items(), key=lambda kv: kv[1])[0],
              "cac_dang": sorted(x["dang"]), "cot": sorted(x["cot"])} for x in tk_rong.values()],
            key=lambda d: -d["so_o"]),
        "don_vi_cot": [dict(v, vi_du=", ".join(v["vi_du"])) for v in tk_cot.values()],
    }
def kiem_tra_nhanh(wb_path: Path, nguon_moi, nguon_cu, sua=None, don_vi=None) -> dict:
    raw = doc_du_lieu_kiem_tra(wb_path, nguon_moi, nguon_cu)
    if raw.get("loi"):
        return raw
    kq = tinh_kiem_tra(raw, sua, don_vi)
    kq["raw"] = raw
    return kq
def luu_chinh_sua_vao_workspace(
    wb_path: Path, sua: Dict[Tuple[str, str, str], str], don_vi: Dict[Tuple[str, str], str],
    ten_cate: Dict[str, str], ten_tskt: Dict[Tuple[str, str], str],
    rong: Optional[Dict[str, list]] = None,
) -> dict:
    """Ghi 1 LẦN vào workspace (lúc xuất): ô sửa tay + đơn vị đã áp -> tab
    TSKT; lựa chọn đơn vị -> sheet ĐƠN VỊ (lần ① Điền dữ liệu sau tự áp);
    nhật ký sửa tay -> sheet CHỈNH SỬA TAY."""
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if don_vi:
        du_lieu = doc_don_vi_day_du(wb)
        for (cate, code), dv in don_vi.items():
            dv = chuan_hoa_key(dv)
            if dv:
                du_lieu[(cate, code)] = [ten_cate.get(cate, ""), ten_tskt.get((cate, code), ""), dv]
            else:
                du_lieu.pop((cate, code), None)
        _ghi_sheet_don_vi(wb, du_lieu)
    nhat_ky: List[list] = []
    so_o = 0
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for ten in wb.sheetnames:
        cate = cate_tu_ten_tab(ten)
        if not cate:
            continue
        sua_cate = {(k, c): v for (ct, k, c), v in sua.items() if ct == cate}
        dv_cate = {c: v for (ct, c), v in don_vi.items() if ct == cate and v and not la_cot_filter(c)}
        if not sua_cate and not dv_cate and not rong:
            continue
        sh = wb[ten]
        col_cua = {c: i for c, i in doc_header_import(sh)}
        c_sku = col_cua.get("sku")
        if not c_sku:
            continue
        for r in range(CFG.data_start_row, (sh.max_row or 0) + 1):
            sku = chuan_hoa_code(sh.cell(row=r, column=c_sku).value)
            if not sku:
                continue
            for code, col in col_cua.items():
                cell = sh.cell(row=r, column=col)
                cu = "" if cell.value is None else str(ep_van_ban_an_toan(cell.value))
                if (sku, code) in sua_cate:
                    moi = chuan_hoa_key(sua_cate[(sku, code)])
                    nhat_ky.append([stamp, cate, sku, code, cu, moi])
                else:
                    moi, hd = ap_quy_tac_rong(cu, code, rong)
                    if not hd and code in dv_cate:
                        moi = them_don_vi(cu, dv_cate[code])
                    elif not hd:
                        continue
                if moi != cu:
                    cell.value = moi or None
                    so_o += 1
    if nhat_ky:
        if CFG_sheet_chinh_sua in wb.sheetnames:
            sh = wb[CFG_sheet_chinh_sua]
        else:
            sh = _tao_lai_sheet(wb, CFG_sheet_chinh_sua, ["THỜI GIAN", "NGÀNH HÀNG", "SKU", "MÃ TSKT",
                                                          "GIÁ TRỊ TRƯỚC", "GIÁ TRỊ SỬA"])
        for row in nhat_ky:
            sh.append(row)
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở trong Excel?): {e}"}
    wb.close()
    return {"loi": None, "so_o": so_o, "so_sua": len(nhat_ky)}
def xuat_ket_qua_kiem_tra(kq: dict, out_path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    bold = Font(bold=True)
    def _sheet(ten, header, rows):
        sh = wb.create_sheet(ten)
        dat_dinh_dang_van_ban_cot(sh, len(header))
        sh.append(header)
        for c in sh[1]:
            c.font = bold
        for r in rows:
            sh.append([v if v != "" else None for v in r])
        sh.freeze_panes = "A2"
    st = kq["stat"]
    _sheet("TÓM TẮT", ["MỤC", "GIÁ TRỊ"], [
        ["Giá trị mới", kq["nhan_moi"]], ["Spec PIM cũ", kq["nhan_cu"]],
        ["Số dòng kiểm tra", st.get("so_dong", 0)],
        ["Ô khác giá trị", st.get(TRANG_THAI_KHAC, 0)], ["Ô tool để trống (PIM đang có)", st.get(TRANG_THAI_TOOL_TRONG, 0)],
        ["Ô tool thêm mới (PIM trống)", st.get(TRANG_THAI_PIM_TRONG, 0)],
        ["Thiếu model_code", st.get("thieu_model", 0)], ["Thiếu category_code", st.get("thieu_cate", 0)],
        ["Ô kích thước/khối lượng chưa có đơn vị", st.get("chua_don_vi", 0)],
        ["Ô chỉ thêm đơn vị (PIM số trơn)", st.get(TRANG_THAI_DON_VI, 0)], ["Ô sửa tay", st.get("da_sua", 0)],
        ["Không có trong file PIM", st.get("khong_co_pim", 0)], ["Dòng sót không có trong IMPORT", st.get("dong_sot", 0)],
        ["Thời gian", datetime.now().strftime("%Y-%m-%d %H:%M")],
    ])
    _sheet("CẢNH BÁO", ["MỨC", "LOẠI", "SKU / MODEL", "NGÀNH HÀNG", "MÃ TSKT", "CHI TIẾT"], kq["canh_bao"])
    _sheet("KHÁC SPEC PIM", ["SKU / MODEL", "MODEL", "NGÀNH HÀNG", "MÃ TSKT", "TÊN TSKT", "PIM CŨ", "TOOL MỚI", "LOẠI"],
           [r[:8] for r in kq["khac"]])
    wb.save(out_path)
def xuat_file_import_nhanh(wb_path: Path, out_dir: Path,
                           sua: Optional[Dict[Tuple[str, str, str], str]] = None,
                           don_vi: Optional[Dict[Tuple[str, str], str]] = None,
                           rong: Optional[Dict[str, list]] = None) -> dict:
    """Bước ③: xuất file import từ tab TSKT HIỆN CÓ (không điền lại, không
    đọc DATA SP -> nhanh). Ra 1 THƯ MỤC .xlsx mở được ngay (khỏi giải nén)
    + 1 bản .zip. CHỈ xuất dòng có SKU trong IMPORT (chặn dòng cũ sót lại)."""
    try:
        wb = _mo_ro_nhanh(str(wb_path))
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    try:
        sku_import: set = set()
        if CFG.sheet_nhap in wb.sheetnames:
            rows_imp = list(wb[CFG.sheet_nhap].iter_rows(values_only=True))
            if rows_imp:
                idx = {chuan_hoa_key(h).lower(): j for j, h in enumerate(rows_imp[0]) if chuan_hoa_key(h)}
                j = idx.get("sku")
                if j is not None:
                    for r in rows_imp[CFG.data_start_row - 1:]:
                        if j < len(r) and r[j] not in (None, ""):
                            sku_import.add(chuan_hoa_code(str(ep_van_ban_an_toan(r[j]))))
        tabs = [(t, list(wb[t].iter_rows(values_only=True))) for t in wb.sheetnames if cate_tu_ten_tab(t)]
    finally:
        wb.close()
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    thu_muc = out_dir / f"{CFG.file_prefix}IMPORT_{stamp}"
    files: List[Tuple[str, bytes, int]] = []
    bo_dong = so_o_sua = so_o_dv = so_o_rong = 0
    sua = sua or {}
    don_vi = {k: v for k, v in (don_vi or {}).items() if v}
    bot = {x.lower() for x in CFG.bot_cot_khi_xuat}
    for title, rows in tabs:
        if len(rows) < CFG.data_start_row:
            continue
        h1 = [chuan_hoa_key(h) for h in rows[0]]
        giu = [j for j, h in enumerate(h1) if h and h.lower() not in bot]
        cate_tab = cate_tu_ten_tab(title)
        j_sku = next((j for j, h in enumerate(h1) if h.lower() == "sku"), None)
        j_var = next((j for j, h in enumerate(h1) if h.lower() == "variant_code"), None)
        model_rows, var_rows = [], []
        for r in rows[CFG.data_start_row - 1:]:
            if not r or all(v in (None, "") for v in r):
                continue
            sku = chuan_hoa_code(str(r[j_sku])) if j_sku is not None and j_sku < len(r) and r[j_sku] else ""
            if not sku:
                continue
            if sku_import and sku not in sku_import:
                bo_dong += 1
                continue
            out = []
            for j in giu:
                v = "" if j >= len(r) or r[j] is None else str(ep_van_ban_an_toan(r[j]))
                key = (cate_tab, sku, h1[j])
                hd = None
                if key not in sua:
                    v, hd = ap_quy_tac_rong(v, h1[j], rong)
                    if hd:
                        so_o_rong += 1
                if key in sua:
                    v = chuan_hoa_key(sua[key])
                    so_o_sua += 1
                elif not hd and (cate_tab, h1[j]) in don_vi and not la_cot_filter(h1[j]):
                    v2 = them_don_vi(v, don_vi[(cate_tab, h1[j])])
                    if v2 != v:
                        v = v2
                        so_o_dv += 1
                out.append(v)
            co_var = j_var is not None and j_var < len(r) and chuan_hoa_key(r[j_var]) != ""
            (var_rows if co_var else model_rows).append(out)
        for hau, data in (("_MODEL", model_rows), ("_BIENTHE", var_rows)):
            if not data:
                continue
            wb_out = Workbook()
            ws = wb_out.active
            ws.title = "Export Product Template"
            dat_dinh_dang_van_ban_cot(ws, len(giu))
            for r_idx, row in enumerate([[rows[0][j] for j in giu], [rows[1][j] if j < len(rows[1]) else None for j in giu]]
                                        + data, start=1):
                for c_idx, v in enumerate(row, start=1):
                    if v not in (None, ""):
                        ws.cell(row=r_idx, column=c_idx, value=v)
            ws.freeze_panes = "A3"
            buf = io.BytesIO()
            wb_out.save(buf)
            files.append((f"{CFG.file_prefix}{title}{hau}_{stamp}.xlsx", buf.getvalue(), len(data)))
    if not files:
        return {"loi": "Không có dòng nào để xuất (tab TSKT trống hoặc SKU không có trong IMPORT)."}
    thu_muc.mkdir(parents=True, exist_ok=True)
    for ten, data, _ in files:
        (thu_muc / ten).write_bytes(data)
    if sua:
        nk = ["NGÀNH HÀNG\tSKU\tMÃ TSKT\tGIÁ TRỊ SỬA TAY"] + [
            f"{c}\t{k}\t{m}\t{v}" for (c, k, m), v in sorted(sua.items())]
        (thu_muc / "_CHINH_SUA_TAY.txt").write_bytes("\r\n".join(nk).encode("utf-8-sig"))
    zip_path = out_dir / f"{CFG.file_prefix}IMPORT_{stamp}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for ten, data, _ in files:
            zf.writestr(ten, data)
    return {"loi": None, "thu_muc": str(thu_muc), "zip": str(zip_path), "bo_dong": bo_dong,
            "so_o_sua": so_o_sua, "so_o_dv": so_o_dv, "so_o_rong": so_o_rong,
            "files": [(ten, n) for ten, _, n in files]}
# ============================================================================
# §13 RUN-ALL — y hệt chayTatCa_1Click()
# ============================================================================
def dem_dong_co_du_lieu_sheet(path: Path, ten_sheet: str, tu_dong: int) -> int:
    """Đếm số dòng THẬT SỰ có dữ liệu (không tính dòng trống) trong 1
    sheet, tính từ `tu_dong` — dùng để hiện số liệu preview trước khi
    nhập (bao nhiêu dòng đang có sẵn / sẽ có thêm bao nhiêu dòng mới),
    giúp người dùng thấy rõ trước khi bấm xác nhận, tránh hiểu lầm mất
    dữ liệu."""
    try:
        wb = mo_workbook_doc_an_toan(path, data_only=True)
        if ten_sheet not in wb.sheetnames:
            wb.close()
            return 0
        sh = wb[ten_sheet]
        if (sh.max_row or 0) < tu_dong:
            wb.close()
            return 0
        n = sum(
            1 for r in sh.iter_rows(min_row=tu_dong, values_only=True)
            if not all(v in (None, "") for v in r)
        )
        wb.close()
        return n
    except Exception:  # noqa: BLE001
        return -1
def dem_dong_co_du_lieu_file(path: str, tu_dong: int) -> int:
    """Giống dem_dong_co_du_lieu_sheet nhưng cho 1 FILE ngoài (sheet đầu
    tiên) — đếm số dòng dữ liệu thật sự trong file nguồn sắp nhập."""
    try:
        wb = mo_workbook_doc_an_toan(path, data_only=True)
        sh = wb.worksheets[0]
        if (sh.max_row or 0) < tu_dong:
            wb.close()
            return 0
        n = sum(
            1 for r in sh.iter_rows(min_row=tu_dong, values_only=True)
            if not all(v in (None, "") for v in r)
        )
        wb.close()
        return n
    except Exception:  # noqa: BLE001
        return -1
def nhap_sheet_theo_vi_tri(
    wb_path: Path,
    src_path: str,
    ten_sheet_dich: str,
    data_start_row: int,
    che_do: str,
    cot_kiem_tra_trung: Optional[int] = None,
    nguon_data_start_row: Optional[int] = None,
    ep_van_ban: bool = True,
) -> dict:
    """Nhập dữ liệu từ 1 file Excel ngoài vào 1 sheet có sẵn trong
    workspace (MAPPING TSKT MOI / MAPPING FILTER MOI / DATA PIM) — copy
    THEO ĐÚNG VỊ TRÍ CỘT (cột 1 nguồn -> cột 1 đích, cột 2 -> cột 2, ...),
    KHÔNG so khớp theo tên header.
    `cot_kiem_tra_trung`: SỐ CỘT (1-based, không phải tên) để báo trùng
    khi che_do='append'.
    `nguon_data_start_row`: dòng dữ liệu thật bắt đầu ở file NGUỒN.
    `ep_van_ban`: True (mặc định) -> ép MỌI mã số về văn bản thật sự.
    Đặt False cho DATA PIM vì có cột IsActivated (boolean) và
    OptionSortOrder (số dùng để sắp xếp) cần giữ nguyên kiểu — khi False,
    cũng KHÔNG ép định dạng hiển thị "@" ở cấp cột (trước đây bản cũ vẫn
    lỡ set number_format="@" cho DATA PIM dù giữ nguyên kiểu giá trị, gây
    lệch giữa kiểu dữ liệu thật và định dạng hiển thị — nay sửa luôn cho
    nhất quán)."""
    if nguon_data_start_row is None:
        nguon_data_start_row = data_start_row
    try:
        wb_src = mo_workbook_doc_an_toan(src_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file nguồn: {e}"}
    sh_src = wb_src.worksheets[0]
    if sh_src.max_row < nguon_data_start_row:
        wb_src.close()
        return {"loi": f"File nguồn không có dòng dữ liệu nào từ dòng {nguon_data_start_row}."}
    du_lieu_nguon = list(sh_src.iter_rows(min_row=nguon_data_start_row, values_only=True))
    wb_src.close()
    while du_lieu_nguon and all(v in (None, "") for v in du_lieu_nguon[-1]):
        du_lieu_nguon.pop()
    if not du_lieu_nguon:
        return {"loi": "File nguồn không có dữ liệu."}
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if ten_sheet_dich not in wb.sheetnames:
        wb.close()
        return {"loi": f'Không tìm thấy sheet "{ten_sheet_dich}" trong workspace.'}
    sh = wb[ten_sheet_dich]
    if che_do == "overwrite":
        dong_cuoi_toan_sheet = sh.max_row
        if dong_cuoi_toan_sheet >= data_start_row:
            sh.delete_rows(data_start_row, dong_cuoi_toan_sheet - data_start_row + 1)
        dong_bat_dau = data_start_row
    else:  # append
        dong_cuoi_cu = dong_cuoi(sh, 1, data_start_row)
        dong_bat_dau = dong_cuoi_cu + 1 if dong_cuoi_cu >= data_start_row else data_start_row
    id_cu: set = set()
    if che_do == "append" and cot_kiem_tra_trung:
        for r in range(data_start_row, dong_bat_dau):
            v = chuan_hoa_id(sh.cell(row=r, column=cot_kiem_tra_trung).value)
            if v:
                id_cu.add(v)
    so_cot_nguon = max((len(row) for row in du_lieu_nguon), default=0)
    if ep_van_ban:
        dat_dinh_dang_van_ban_cot(sh, so_cot_nguon)
    so_dong_ghi = 0
    trung: list[str] = []
    da_bao: set = set()
    for row in du_lieu_nguon:
        if row is None or all(v in (None, "") for v in row):
            continue
        dong_dich = dong_bat_dau + so_dong_ghi
        for c_idx, gia_tri in enumerate(row):
            if ep_van_ban:
                gia_tri = ep_van_ban_an_toan(gia_tri)
            sh.cell(row=dong_dich, column=c_idx + 1, value=gia_tri)
        if cot_kiem_tra_trung and len(row) >= cot_kiem_tra_trung:
            v = chuan_hoa_id(row[cot_kiem_tra_trung - 1])
            if v and v in id_cu and v not in da_bao:
                trung.append(v)
                da_bao.add(v)
        so_dong_ghi += 1
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {"so_dong": so_dong_ghi, "dong_bat_dau": dong_bat_dau, "trung": trung}
def nhap_du_lieu_theo_sheet(
    wb_path: Path,
    src_path: str,
    ten_sheet_dich: str,
    header_row: int,
    data_start_row: int,
    che_do: str,
    cot_kiem_tra_trung: Optional[str] = None,
    nguon_data_start_row: Optional[int] = None,
) -> dict:
    """Nhập dữ liệu từ 1 file Excel BÊN NGOÀI vào 1 sheet CÓ SẴN trong
    workspace — khớp cột theo TÊN header ở `header_row`.
    GHI CHÚ: hàm này hiện KHÔNG được gọi từ bất kỳ nút nào trong GUI (GUI
    dùng nhap_sheet_theo_vi_tri — khớp theo VỊ TRÍ cột — cho MAPPING
    TSKT/FILTER MOI và DATA PIM). Giữ lại vì vẫn đúng và có thể hữu ích
    nếu sau này cần nhập theo tên cột thay vì vị trí, nhưng không phải
    đường chạy chính — không gọi nhầm hàm này tưởng đang dùng."""
    if nguon_data_start_row is None:
        nguon_data_start_row = data_start_row
    try:
        wb_src = load_workbook(src_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file nguồn: {e}"}
    sh_src = wb_src.worksheets[0]
    if sh_src.max_row < header_row:
        wb_src.close()
        return {"loi": "File nguồn không có dữ liệu."}
    header_nguon = [chuan_hoa_key(h) for h in sheet_display_values(sh_src, header_row, header_row, sh_src.max_column)[0]]
    if sh_src.max_row < nguon_data_start_row:
        wb_src.close()
        return {"loi": f"File nguồn không có dòng dữ liệu nào từ dòng {nguon_data_start_row}."}
    du_lieu_nguon = list(sh_src.iter_rows(min_row=nguon_data_start_row, values_only=True))
    wb_src.close()
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if ten_sheet_dich not in wb.sheetnames:
        wb.close()
        return {"loi": f'Không tìm thấy sheet "{ten_sheet_dich}" trong workspace.'}
    sh = wb[ten_sheet_dich]
    header_dich = [chuan_hoa_key(h) for h in sheet_display_values(sh, header_row, header_row, sh.max_column)[0]]
    vi_tri_dich = {h: i + 1 for i, h in enumerate(header_dich) if h}
    thieu_cot = sorted({h for h in header_nguon if h and h not in vi_tri_dich})
    dong_bat_dau: int
    if che_do == "overwrite":
        dong_cuoi_toan_sheet = sh.max_row
        if dong_cuoi_toan_sheet >= data_start_row:
            sh.delete_rows(data_start_row, dong_cuoi_toan_sheet - data_start_row + 1)
        dong_bat_dau = data_start_row
    else:  # append
        dong_cuoi_cu = dong_cuoi(sh, 1, data_start_row)
        dong_bat_dau = dong_cuoi_cu + 1 if dong_cuoi_cu >= data_start_row else data_start_row
    so_dong_ghi = 0
    for row in du_lieu_nguon:
        if row is None or all(v in (None, "") for v in row):
            continue
        dong_dich = dong_bat_dau + so_dong_ghi
        for c_idx, gia_tri in enumerate(row):
            if c_idx >= len(header_nguon):
                break
            ten_cot = header_nguon[c_idx]
            if not ten_cot or ten_cot not in vi_tri_dich:
                continue
            cot_dich = vi_tri_dich[ten_cot]
            cell = sh.cell(row=dong_dich, column=cot_dich, value=gia_tri)
            cell.number_format = "@"
        so_dong_ghi += 1
    trung: list[str] = []
    if che_do == "append" and cot_kiem_tra_trung and cot_kiem_tra_trung in vi_tri_dich and so_dong_ghi > 0:
        cot_check = vi_tri_dich[cot_kiem_tra_trung]
        gia_tri_cu = set()
        for r in range(data_start_row, dong_bat_dau):
            v = chuan_hoa_key(sh.cell(row=r, column=cot_check).value)
            if v:
                gia_tri_cu.add(v)
        for r in range(dong_bat_dau, dong_bat_dau + so_dong_ghi):
            v = chuan_hoa_key(sh.cell(row=r, column=cot_check).value)
            if v and v in gia_tri_cu and v not in trung:
                trung.append(v)
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {"so_dong": so_dong_ghi, "dong_bat_dau": dong_bat_dau, "thieu_cot": thieu_cot, "trung": trung}
def dong_cuoi_toan_dong(sh: Worksheet, start_row: int) -> int:
    """Dòng cuối cùng có BẤT KỲ ô nào khác rỗng (quét toàn bộ các cột),
    tính từ start_row — dùng cho sheet mà 1 "bản ghi" chiếm nhiều dòng và
    dòng phụ (vd nhãn) có thể để trống hẳn cột đầu tiên (CẤU HÌNH
    CATEGORY: dòng mã TSKT + dòng nhãn tiếng Việt ngay dưới, dòng nhãn có
    cột A/B trống)."""
    last_row = sh.max_row
    if last_row < start_row:
        return start_row - 1
    for r in range(last_row, start_row - 1, -1):
        for c in range(1, sh.max_column + 1):
            v = sh.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                return r
    return start_row - 1
def nhap_cau_hinh_category_tuong_tac(wb_path: Path, src_path: str) -> dict:
    """Nhập CẤU HÌNH CATEGORY (ngành hàng mới) từ 1 file Excel ngoài —
    LUÔN nối tiếp, báo TRÙNG nếu CATEGORY ID đã có sẵn (không chặn)."""
    try:
        wb_src = load_workbook(src_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được file nguồn: {e}"}
    sh_src = wb_src.worksheets[0]
    if sh_src.max_row < 1:
        wb_src.close()
        return {"loi": "File nguồn không có dữ liệu."}
    id0 = chuan_hoa_id(sh_src.cell(row=1, column=1).value)
    co_header = not bool(re.fullmatch(r"\d+", id0))
    dong_bat_dau_doc = 2 if co_header else 1
    if sh_src.max_row < dong_bat_dau_doc:
        wb_src.close()
        return {"loi": "File nguồn không có dòng dữ liệu nào."}
    rows_src = list(sh_src.iter_rows(min_row=dong_bat_dau_doc, values_only=True))
    wb_src.close()
    while rows_src and all(v in (None, "") for v in rows_src[-1]):
        rows_src.pop()
    if not rows_src:
        return {"loi": "File nguồn không có dữ liệu."}
    try:
        wb = load_workbook(wb_path)
    except Exception as e:  # noqa: BLE001
        return {"loi": f"Không mở được workspace: {e}"}
    if CFG.sheet_cau_hinh not in wb.sheetnames:
        wb.close()
        return {"loi": f'Không tìm thấy sheet "{CFG.sheet_cau_hinh}" trong workspace.'}
    sh = wb[CFG.sheet_cau_hinh]
    dong_cuoi_cu = dong_cuoi_toan_dong(sh, 2)
    dong_bat_dau = dong_cuoi_cu + 1 if dong_cuoi_cu >= 2 else 2
    id_cu = set()
    for r in range(2, dong_bat_dau):
        v = chuan_hoa_id(sh.cell(row=r, column=1).value)
        if v:
            id_cu.add(v)
    so_cot_nguon = max((len(row) for row in rows_src), default=0)
    dat_dinh_dang_van_ban_cot(sh, so_cot_nguon)
    so_dong_ghi = 0
    trung: list[str] = []
    id_moi_da_thay: set[str] = set()
    for row in rows_src:
        if row is None or all(v in (None, "") for v in row):
            continue
        dong_dich = dong_bat_dau + so_dong_ghi
        for c_idx, gia_tri in enumerate(row):
            sh.cell(row=dong_dich, column=c_idx + 1, value=ep_van_ban_an_toan(gia_tri))
        id_dong = chuan_hoa_id(row[0]) if len(row) > 0 else ""
        if id_dong and id_dong in id_cu and id_dong not in id_moi_da_thay:
            trung.append(id_dong)
            id_moi_da_thay.add(id_dong)
        so_dong_ghi += 1
    try:
        wb.save(wb_path)
    except Exception as e:  # noqa: BLE001
        wb.close()
        return {"loi": f"Không lưu được workspace (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}"}
    wb.close()
    return {"so_dong": so_dong_ghi, "dong_bat_dau": dong_bat_dau, "thieu_cot": [], "trung": trung}
def dam_bao_cot_day_du(sh: Worksheet, cfg_cate: Optional[CategoryConfig], cache: Cache) -> int:
    """Rà lại tab TSKT/FILTER ĐÃ CÓ SẴN so với danh sách mã hiện tại
    trong CẤU HÌNH CATEGORY — mã nào ĐÃ được khai báo nhưng CHƯA có cột
    tương ứng trên tab thì TỰ ĐỘNG THÊM cột mới vào cuối."""
    if not cfg_cate or not cfg_cate.tskt:
        return 0
    cols = doc_header_import(sh)
    da_co = {c for c, _ in cols}
    thieu = [code for code in cfg_cate.tskt if code not in da_co]
    if not thieu:
        return 0
    last_col = max((i for _, i in cols), default=0)
    bold = Font(bold=True)
    for j, code in enumerate(thieu, start=1):
        col = last_col + j
        c1 = sh.cell(row=1, column=col, value=code)
        c1.number_format = "@"
        c1.font = bold
        ten = cfg_cate.ten_tskt.get(code) or cache.ten_pim.get(code, "")
        c2 = sh.cell(row=2, column=col, value=ten)
        c2.number_format = "@"
        c2.font = bold
    return len(thieu)
def chay_tat_ca(workbook_path: Path, out_dir: Path, xuat: bool = True) -> Optional[dict]:
    """Trả về {'co_canh_bao': bool, 'tom_tat': str} khi chạy hết, None nếu dừng sớm."""
    print(f"Đang mở workbook: {workbook_path}")
    wb = load_workbook(workbook_path)
    bao: List[str] = []
    log_all: List[list] = []
    # ---- B1: Trích DATA SP nếu cần ----
    sh_ds = tim_sheet(wb, CFG.sheet_data_sp)
    can_trich = True
    if sh_ds is not None:
        d = dong_cuoi(sh_ds, 2, CFG.sp_start_row)
        can_trich = d < CFG.sp_start_row
    if can_trich:
        sh_tho = tim_sheet_du_lieu_tho(wb)
        if sh_tho:
            kq = trich_data_sp_silent(wb, sh_tho)
            if kq.get("loi"):
                bao.append(f"B1 Trích DATA SP: ⚠ {kq['loi']}")
            else:
                bao.append(f'B1 Trích DATA SP: ✔ {kq["so_dong"]} dòng từ "{sh_tho.title}"')
        else:
            bao.append("B1 Trích DATA SP: bỏ qua (không tìm được đúng 1 sheet dữ liệu thô)")
    else:
        bao.append("B1 Trích DATA SP: bỏ qua (DATA SP đã có sẵn)")
    # ---- B2: Đọc IMPORT + cache + mapping TSKT/FILTER ----
    loi_nl, nl_rows, nl_map = doc_nhap_lieu(wb)
    if loi_nl:
        print(f"LỖI B2: {loi_nl}")
        return
    cache = doc_cache(wb)
    if cache.loi:
        print(f"LỖI B2: {cache.loi}")
        return
    try:
        doc_tskt_mapping(wb, cache)
        doc_filter_mapping(wb, cache)
    except RuntimeError as e:
        bao.append(f"B2 Đọc mapping TSKT/FILTER: ⚠ {e} (vẫn tiếp tục)")
    # ---- B3: Phân loại (tương đương tự tick TẤT CẢ) ----
    theo_cate, khong_co_data, pl_log = phan_loai(cache, nl_rows)
    log_all.extend(pl_log)
    cau_hinh = doc_cau_hinh(wb)
    tab_cua_cate = tim_tab_tskt(wb, cache)
    _ghi_sheet_chon(wb, theo_cate, cache, cau_hinh, tab_cua_cate)
    cates = [c for c, info in theo_cate.items() if info.skus]
    if not cates:
        print("B3 — Không có SKU nào phân loại được vào cate.")
        return
    bao.append(f"B3 Phân loại: ✔ {len(nl_rows)} SKU → {len(cates)} cate")
    # ---- B4: Điền TSKT + FILTER cho MỌI cate ----
    tong_sku, tong_o = 0, 0
    tong_cot_bo_sung = 0
    don_vi_theo_cate = doc_don_vi(wb)
    tab_da_dien: Dict[str, Worksheet] = {}
    for cate in cates:
        info = theo_cate[cate]
        tab = tab_cua_cate.get(cate)
        cfg = cau_hinh.get(cate)
        if not tab and cfg and cfg.tskt:
            tab = tao_tab_tu_cau_hinh(wb, cate, cfg, cache)
        if not tab:
            log_all.append(["", cate, "", "Không có tab TSKT lẫn cấu hình — bỏ qua"])
            continue
        so_cot_them = dam_bao_cot_day_du(tab, cfg, cache)
        if so_cot_them:
            tong_cot_bo_sung += so_cot_them
            log_all.append(["", cate, "", f"Tự bổ sung {so_cot_them} cột còn thiếu so với cấu hình (vd FILTER mới thêm sau) vào tab có sẵn"])
        kq = dien_mot_cate_tskt_filter(cache, tab, cate, info.skus, nl_map, cfg, log_all,
                                       don_vi=don_vi_theo_cate.get(cate))
        if kq.get("loi"):
            log_all.append(["", cate, "", f"Điền lỗi: {kq['loi']}"])
            continue
        tab_da_dien[cate] = tab
        tong_sku += len(info.skus)
        tong_o += kq.get("so_o", 0)
    bao.append(f"B4 Điền TSKT/FILTER: ✔ {tong_sku} SKU, {tong_o} ô" + (f" (đã tự bổ sung {tong_cot_bo_sung} cột thiếu)" if tong_cot_bo_sung else ""))
    so_cot_dv = sum(len(v) for c, v in don_vi_theo_cate.items() if c in tab_da_dien)
    if so_cot_dv:
        bao.append(f"    (đã tự thêm đơn vị cho {so_cot_dv} cột kích thước/khối lượng theo sheet {CFG_sheet_don_vi})")
    if log_all:
        ghi_log(wb, log_all)
    # ---- B4b: Đối chiếu với spec PIM đã nạp (nếu có) ----
    du_lieu_tab = {cate: _doc_du_lieu_tab(tab) for cate, tab in tab_da_dien.items()}
    spec, spec_info = doc_spec_pim(wb)
    dc_stat = None
    if spec is not None:
        dc_rows, dc_stat = doi_chieu_spec(du_lieu_tab, spec, cache)
        _ghi_sheet_doi_chieu(wb, dc_rows, True)
        bao.append(f"B4b Đối chiếu spec PIM: {dc_stat['sku_doi_chieu']} SKU — {dc_stat['khac']} ô khác, "
                   f"{dc_stat['tool_trong']} ô tool để trống (PIM đang có), {dc_stat['pim_trong']} ô tool thêm mới")
    else:
        _ghi_sheet_doi_chieu(wb, [], False)
        bao.append("B4b Đối chiếu spec PIM: bỏ qua (chưa nạp file export PIM)")
    # ---- B4c: Tổng hợp cảnh báo trước khi xuất ----
    cb_rows, cb_tom = tong_hop_canh_bao(du_lieu_tab, nl_map, spec, spec_info, dc_stat)
    _ghi_sheet_canh_bao(wb, cb_rows)
    co_canh_bao = any(r[0] in ("CAO", "TRUNG BÌNH") for r in cb_rows)
    them_file: List[Tuple[str, bytes]] = []
    if co_canh_bao:
        txt = ["CẢNH BÁO TRƯỚC KHI IMPORT PIM — " + datetime.now().strftime("%Y-%m-%d %H:%M"), ""]
        txt += cb_tom + ["", "CHI TIẾT (MỨC | LOẠI | SKU | NGÀNH HÀNG | MÃ TSKT | CHI TIẾT):"]
        txt += ["\t".join(str(x) for x in r) for r in cb_rows]
        them_file.append(("_CANH_BAO_DOC_TRUOC_KHI_IMPORT.txt", "\r\n".join(txt).encode("utf-8-sig")))
    # ---- B5: Xuất .zip tách Model/Biến thể ----
    if not xuat:
        kq_zip = {"loi": None, "chua_xuat": True}
        bao.append("B5 Xuất file: CHƯA XUẤT (chế độ điền trước — kiểm tra ở tab 🔍 rồi bấm ③ Xuất file import)")
    else:
        kq_zip = tao_zip_tach_model(wb, cache, cates, out_dir, them_file=them_file)
    if kq_zip.get("chua_xuat"):
        pass
    elif kq_zip.get("loi"):
        bao.append(f"B5 Xuất .zip: ⚠ {kq_zip['loi']}")
    else:
        bao.append(f'B5 Xuất .zip: ✔ {kq_zip["so_file"]} file .xlsx ({kq_zip["zip_path"]})')
        for w in kq_zip.get("canh_bao", []):
            bao.append(f"    {w}")
    try:
        wb.save(workbook_path)
    except Exception as e:  # noqa: BLE001
        print(f"\nLỖI khi lưu workbook (file đang mở ở nơi khác? / không đủ quyền ghi?): {e}")
        print("\n".join(bao))
        raise
    print("\n".join(bao))
    tom_tat = ""
    if co_canh_bao:
        tom_tat = (
            "⚠️ CẢNH BÁO — ĐỌC TRƯỚC KHI IMPORT LÊN PIM\n"
            f"(chi tiết: sheet \"{CFG_sheet_canh_bao}\" trong workspace + file _CANH_BAO_DOC_TRUOC_KHI_IMPORT.txt trong .zip)\n\n"
            + "\n".join(l for l in cb_tom if not l.startswith("• (Chưa nạp"))
        )
        print("\n" + "!" * 70 + "\n" + tom_tat + "\n" + "!" * 70)
    elif cb_tom:
        print("\n".join(cb_tom))
    print(f"\nĐã lưu lại workbook: {workbook_path}")
    return {"co_canh_bao": co_canh_bao, "tom_tat": tom_tat}
def _ghi_sheet_chon(
    wb: Workbook, theo_cate: Dict[str, CateInfo], cache: Cache,
    cau_hinh: Dict[str, CategoryConfig], tab_cua_cate: Dict[str, Worksheet],
) -> None:
    if CFG.sheet_chon in wb.sheetnames:
        sh = wb[CFG.sheet_chon]
        wb.remove(sh)
    sh = wb.create_sheet(CFG.sheet_chon)
    bold = Font(bold=True)
    headers = ["CHỌN", "MÃ NH", "TÊN NGÀNH HÀNG", "SỐ SKU", "NGUỒN CATE", "TAB TSKT / CẤU HÌNH", "GHI CHÚ"]
    for col, h in enumerate(headers, start=1):
        c = sh.cell(row=1, column=col, value=h)
        c.font = bold
    sh.freeze_panes = "A2"
    items = sorted(theo_cate.items(), key=lambda kv: -len(kv[1].skus))
    for r_idx, (cate, info) in enumerate(items, start=2):
        tab = tab_cua_cate.get(cate)
        cfg = cau_hinh.get(cate)
        if info.tu_cate_id and info.tu_doan:
            nguon = "CATEGORYID + tự nhận diện"
        elif info.tu_cate_id:
            nguon = "CATEGORYID"
        else:
            nguon = "tự nhận diện"
        tskth = tab.title if tab else (f"(sẽ tự tạo từ cấu hình — {len(cfg.tskt)} TSKT)" if cfg else "")
        ghi_chu = []
        if not tab and not cfg:
            ghi_chu.append("Chưa có tab TSKT lẫn cấu hình — thêm 1 trong 2")
        if cfg and not cfg.tskt:
            ghi_chu.append("Cấu hình rỗng")
        if info.tie:
            ghi_chu.append("Có SKU tự nhận diện trùng điểm (xem LOG)")
        row = [True, cate, cache.ten_cate.get(cate) or (cfg.name if cfg else ""), len(info.skus),
               nguon, tskth, " | ".join(ghi_chu)]
        for c_idx, v in enumerate(row, start=1):
            cell = sh.cell(row=r_idx, column=c_idx, value=v)
            if c_idx >= 2:
                cell.number_format = "@"
# ============================================================================
# WORKSPACE TỰ QUẢN LÝ — thư mục + file Excel mẫu TỰ TẠO ngay lần chạy đầu
# ============================================================================
WORKSPACE_DIR_NAME = "PIM_Data"
WORKSPACE_FILE_NAME = "du_lieu_pim.xlsx"
def _thu_muc_goc_app() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
def dam_bao_workspace() -> Path:
    ws_dir = _thu_muc_goc_app() / WORKSPACE_DIR_NAME
    ws_dir.mkdir(parents=True, exist_ok=True)
    (ws_dir / "output").mkdir(exist_ok=True)
    ws_file = ws_dir / WORKSPACE_FILE_NAME
    if ws_file.exists():
        return ws_file
    wb = Workbook()
    wb.remove(wb.active)
    bold = Font(bold=True)
    sh = wb.create_sheet(CFG.sheet_nhap)
    for col, (code, ten) in enumerate(
        [("model_code", "Mã model"), ("sku", "Mã sản phẩm ERP"),
         ("variant_code", "Mã biến thể"), ("category_code", "Mã danh mục PIM")],
        start=1,
    ):
        c1 = sh.cell(row=1, column=col, value=code)
        c1.font = bold
        c1.number_format = "@"
        c2 = sh.cell(row=2, column=col, value=ten)
        c2.font = bold
        c2.number_format = "@"
    sh.freeze_panes = "A3"
    sh = wb.create_sheet(CFG.sheet_data_sp[0])
    _ghi_header_data_sp(sh)
    sh = wb.create_sheet(CFG.sheet_data_pim[0])
    for col, (code, ten) in enumerate(
        [("Code", "Mã thuộc tính"), ("Name", "Tên"), ("Type", "Loại"), ("Group", "Nhóm"),
         ("Active", "Hoạt động"), ("OptionCode", "Mã giá trị"), ("OptionValue", "Giá trị")],
        start=1,
    ):
        c1 = sh.cell(row=1, column=col, value=code)
        c1.font = bold
        c2 = sh.cell(row=2, column=col, value=ten)
        c2.font = bold
    sh.freeze_panes = "A3"
    sh = wb.create_sheet(CFG.sheet_cau_hinh)
    for col, h in enumerate(["CATEGORY ID", "CATEGORY NAME", "DANH SÁCH TSKT (ngang, mỗi ô 1 mã)"], start=1):
        c = sh.cell(row=1, column=col, value=h)
        c.font = bold
    sh.freeze_panes = "A2"
    sh = wb.create_sheet(CFG_TF_sheet_tskt_mapping)
    for col, h in enumerate(
        ["Mã ngành hàng CMS", "Tên ngành hàng CMS", "Mã thuộc tính", "Tên thuộc tính TSKT",
         "Mã TSKT (MASTER)", "Tên TSKT (MASTER)"],
        start=1,
    ):
        c = sh.cell(row=1, column=col, value=h)
        c.font = bold
    sh.freeze_panes = "A2"
    sh = wb.create_sheet(CFG_TF_sheet_filter_mapping)
    for col, h in enumerate(
        ["MÃ NGÀNH HÀNG CMS", "TÊN NGÀNH HÀNG CMS", "MÃ THUỘC TÍNH FILTER", "TÊN THUỘC TÍNH CMS",
         "Mã thuộc tính cũ", "Mã thuộc tính mới"],
        start=1,
    ):
        c = sh.cell(row=1, column=col, value=h)
        c.font = bold
    sh.freeze_panes = "A2"
    wb.save(ws_file)
    return ws_file
def _ten_sheet_khong_trung(workbook_path: Path, ten_goc: str) -> str:
    ten_sach = re.sub(r'[\\/*?:\[\]]', "_", ten_goc).strip()[:25] or "DuLieuTho"
    if ten_sach.upper().startswith(CFG.tskth_prefix.upper()):
        ten_sach = ("RAW_" + ten_sach)[:25]
    try:
        wb = load_workbook(workbook_path, read_only=True)
        existing = set(wb.sheetnames)
        wb.close()
    except Exception:  # noqa: BLE001
        existing = set()
    ten = ten_sach
    n = 1
    while ten in existing:
        n += 1
        ten = f"{ten_sach}_{n}"[:31]
    return ten
class _StdoutRedirect:
    def __init__(self, q: "queue.Queue[str]"):
        self.q = q
    def write(self, s: str) -> int:
        if s:
            self.q.put(s)
        return len(s)
    def flush(self) -> None:
        pass
def _can_giua_popup(parent: tk.Misc, top: tk.Toplevel) -> None:
    top.update_idletasks()
    w, h = top.winfo_reqwidth(), top.winfo_reqheight()
    try:
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - w) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - h) // 2)
        top.geometry(f"+{x}+{y}")
    except tk.TclError:
        pass
    top.resizable(False, False)
class SheetEditorFrame(ttk.Frame):
    MAX_COLS = 60
    MAX_ROWS_CANH_BAO = 5000
    def __init__(self, parent: tk.Widget, get_file_path):
        super().__init__(parent)
        self.get_file_path = get_file_path
        self._so_cot_dang_hien = 10
        self._sel_anchor: tuple[int, int] | None = None
        self._sel_end: tuple[int, int] | None = None
        # BỔ SUNG: nhớ xem sheet đang tải có bị CẮT BỚT không (nhiều hơn
        # MAX_ROWS_CANH_BAO dòng hoặc nhiều hơn MAX_COLS cột trong file
        # thật) — dùng để CHẶN CỨNG nút Lưu, tránh bug mất dữ liệu đã gặp:
        # trước đây chỉ có dòng chữ cảnh báo, không có gì ngăn người dùng
        # lỡ bấm Lưu rồi bị xoá sạch phần dữ liệu chưa tải lên lưới.
        self._sheet_dang_tai_bi_cat = False
        self._so_dong_that_cua_sheet = 0
        self._so_cot_that_cua_sheet = 0
        self._build_ui()
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)
        ttk.Button(frm_top, text="🔄 Nạp danh sách sheet từ file", command=self._nap_danh_sach_sheet).pack(
            side="left"
        )
        ttk.Label(frm_top, text="  Sheet:").pack(side="left", padx=(14, 0))
        self.var_sheet = tk.StringVar()
        self.cbo_sheet = ttk.Combobox(frm_top, textvariable=self.var_sheet, state="readonly", width=32)
        self.cbo_sheet.pack(side="left", padx=(4, 0))
        ttk.Button(frm_top, text="📥 Tải dữ liệu sheet này", command=self._tai_du_lieu_sheet).pack(
            side="left", padx=(10, 0)
        )
        ttk.Button(frm_top, text="➕ Tạo sheet mới...", command=self._tao_sheet_moi).pack(side="left", padx=(10, 0))
        ttk.Button(frm_top, text="✏️ Đổi tên sheet...", command=self._doi_ten_sheet).pack(side="left", padx=(10, 0))
        frm_top2 = ttk.Frame(self)
        frm_top2.pack(fill="x", padx=10)
        ttk.Button(frm_top2, text="➕ Thêm dòng", command=self._them_dong).pack(side="left")
        ttk.Button(frm_top2, text="🗑️ Xóa dòng đã chọn", command=self._xoa_dong).pack(side="left", padx=(6, 0))
        self.btn_luu = ttk.Button(frm_top2, text="💾 Lưu vào file", style="Accent.TButton", command=self._luu_vao_file)
        self.btn_luu.pack(side="left", padx=(16, 0))
        ttk.Label(
            frm_top2,
            text="(double-click 1 ô để sửa; kéo chuột/Shift+Click chọn vùng; Ctrl+C/Ctrl+V copy-paste; Delete xoá nội dung)",
            foreground="#777",
        ).pack(side="left", padx=(14, 0))
        self.var_status = tk.StringVar(value="Chưa nạp sheet nào — bấm 'Nạp danh sách sheet từ file' trước.")
        ttk.Label(self, textvariable=self.var_status, foreground="#555").pack(anchor="w", padx=10, pady=(2, 0))
        frm_grid = ttk.Frame(self)
        frm_grid.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.tree = ttk.Treeview(frm_grid, show="headings", selectmode="extended")
        self.tree.tag_configure("odd", background="#f3f5fb")
        self.tree.tag_configure("even", background="#ffffff")
        vs = ttk.Scrollbar(frm_grid, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(frm_grid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frm_grid.rowconfigure(0, weight=1)
        frm_grid.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self._bat_dau_sua_o)
        self.tree.bind("<ButtonPress-1>", self._bat_dau_chon_vung, add="+")
        self.tree.bind("<B1-Motion>", self._keo_chon_vung)
        self.tree.bind("<ButtonRelease-1>", self._ket_thuc_chon_vung)
        self.tree.bind("<Shift-Button-1>", self._shift_click_chon_vung)
        self.tree.bind("<Control-c>", self._sao_chep)
        self.tree.bind("<Control-C>", self._sao_chep)
        self.tree.bind("<Control-v>", self._dan)
        self.tree.bind("<Control-V>", self._dan)
        self.tree.bind("<Delete>", self._xoa_noi_dung_vung)
        self.tree.bind("<BackSpace>", self._xoa_noi_dung_vung)
        self._menu_chuot_phai = tk.Menu(self.tree, tearoff=0)
        self._menu_chuot_phai.add_command(label="📋 Copy (Ctrl+C)", command=self._sao_chep)
        self._menu_chuot_phai.add_command(label="📥 Dán (Ctrl+V)", command=self._dan)
        self._menu_chuot_phai.add_command(label="🧹 Xoá nội dung vùng chọn (Delete)", command=self._xoa_noi_dung_vung)
        self._menu_chuot_phai.add_separator()
        self._menu_chuot_phai.add_command(label="➕ Thêm dòng", command=self._them_dong)
        self._menu_chuot_phai.add_command(label="🗑️ Xoá dòng đã chọn", command=self._xoa_dong)
        self.tree.bind("<Button-3>", self._mo_menu_chuot_phai)
        self._editor: tk.Entry | None = None
    def _nap_danh_sach_sheet(self) -> None:
        path = self.get_file_path()
        if not path:
            messagebox.showwarning("Thiếu file", "Chọn file Excel nguồn ở tab '🚀 Chạy pipeline' trước.")
            return
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy file:\n{p}")
            return
        try:
            wb = load_workbook(p, read_only=True)
            ten_sheets = list(wb.sheetnames)
            wb.close()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi mở file", str(e))
            return
        self.cbo_sheet["values"] = ten_sheets
        if ten_sheets:
            self.var_sheet.set(ten_sheets[0])
        self.var_status.set(f"Đã nạp danh sách {len(ten_sheets)} sheet. Chọn 1 sheet rồi bấm 'Tải dữ liệu sheet này'.")
    def _tai_du_lieu_sheet(self) -> None:
        path = self.get_file_path()
        sheet_name = self.var_sheet.get()
        if not path or not sheet_name:
            messagebox.showwarning("Thiếu thông tin", "Nạp danh sách sheet và chọn 1 sheet trước.")
            return
        try:
            wb = load_workbook(Path(path))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi mở file", str(e))
            return
        if sheet_name not in wb.sheetnames:
            messagebox.showerror("Không tìm thấy sheet", f'Sheet "{sheet_name}" không còn tồn tại trong file.')
            return
        sh = wb[sheet_name]
        so_cot_that = sh.max_column
        so_dong_that = sh.max_row
        so_cot = max(min(so_cot_that, self.MAX_COLS), 10)
        canh_bao = ""
        # BỔ SUNG: nhớ lại sheet này có bị cắt bớt hay không (dòng HOẶC
        # cột) để _luu_vao_file() chặn lưu nếu có — xem giải thích ở đó.
        self._sheet_dang_tai_bi_cat = so_dong_that > self.MAX_ROWS_CANH_BAO or so_cot_that > self.MAX_COLS
        self._so_dong_that_cua_sheet = so_dong_that
        self._so_cot_that_cua_sheet = so_cot_that
        if so_dong_that > self.MAX_ROWS_CANH_BAO:
            canh_bao += (
                f" ⚠️ Sheet có {so_dong_that} dòng — chỉ TẢI XEM {self.MAX_ROWS_CANH_BAO} dòng đầu để tránh "
                "treo giao diện."
            )
        if so_cot_that > self.MAX_COLS:
            canh_bao += f" ⚠️ Sheet có {so_cot_that} cột — chỉ TẢI XEM {self.MAX_COLS} cột đầu."
        if self._sheet_dang_tai_bi_cat:
            canh_bao += (
                " KHU VỰC NÀY CHỈ DÙNG ĐỂ XEM sheet lớn — nút '💾 Lưu vào file' sẽ BỊ KHOÁ cho sheet này "
                "(để không lỡ ghi đè mất phần dữ liệu chưa tải lên). Sửa sheet lớn (vd DATA SP) trực tiếp "
                "bằng Excel, hoặc dùng các nút nhập/xoá dữ liệu chuyên dụng ở tab 'Chạy pipeline'."
            )
        so_dong_tai = min(so_dong_that, self.MAX_ROWS_CANH_BAO)
        self._so_cot_dang_hien = so_cot
        self._sel_anchor = None
        self._sel_end = None
        cols = [f"C{i}" for i in range(1, so_cot + 1)]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for i, cid in enumerate(cols, start=1):
            self.tree.heading(cid, text=get_column_letter(i))
            self.tree.column(cid, width=110, anchor="w", stretch=False)
        if so_dong_tai > 0:
            for idx, row in enumerate(sh.iter_rows(min_row=1, max_row=so_dong_tai, max_col=so_cot, values_only=True)):
                vals = ["" if v is None else str(v) for v in row]
                self.tree.insert("", "end", values=vals, tags=("even" if idx % 2 == 0 else "odd",))
        wb.close()
        # Khoá/mở nút Lưu tuỳ theo sheet có bị cắt bớt hay không.
        self.btn_luu.configure(state="disabled" if self._sheet_dang_tai_bi_cat else "normal")
        self.var_status.set(
            f'Sheet "{sheet_name}": {so_dong_tai} dòng × {so_cot} cột đã tải.{canh_bao}'
        )
    def _tao_sheet_moi(self) -> None:
        path = self.get_file_path()
        if not path:
            messagebox.showwarning("Thiếu file", "Chọn file Excel nguồn ở tab '🚀 Chạy pipeline' trước.")
            return
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy file:\n{p}")
            return
        top = tk.Toplevel(self)
        top.title("Tạo sheet mới")
        ttk.Label(top, text="Tên sheet mới:").pack(anchor="w", padx=12, pady=(14, 4))
        var_ten = tk.StringVar()
        ent = ttk.Entry(top, textvariable=var_ten, width=40)
        ent.pack(padx=12)
        ent.focus_set()
        def _tao():
            ten = var_ten.get().strip()
            if not ten:
                messagebox.showwarning("Thiếu tên", "Gõ tên sheet trước.", parent=top)
                return
            try:
                wb = load_workbook(p)
                if ten in wb.sheetnames:
                    messagebox.showerror("Đã tồn tại", f'Sheet "{ten}" đã có sẵn trong file.', parent=top)
                    wb.close()
                    return
                wb.create_sheet(ten)
                wb.save(p)
                wb.close()
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Lỗi", str(e), parent=top)
                return
            top.destroy()
            self._nap_danh_sach_sheet()
            self.var_sheet.set(ten)
            self._tai_du_lieu_sheet()
        ttk.Button(top, text="Tạo", style="Accent.TButton", command=_tao).pack(pady=12)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    def _doi_ten_sheet(self) -> None:
        path = self.get_file_path()
        if not path:
            messagebox.showwarning("Thiếu file", "Chọn file Excel nguồn ở tab '🚀 Chạy pipeline' trước.")
            return
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy file:\n{p}")
            return
        try:
            wb0 = load_workbook(p, read_only=True)
            ten_sheets = list(wb0.sheetnames)
            wb0.close()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi mở file", str(e))
            return
        if not ten_sheets:
            messagebox.showinfo("Chưa có sheet", "File chưa có sheet nào.")
            return
        top = tk.Toplevel(self)
        top.title("Đổi tên sheet")
        ttk.Label(top, text="Sheet cần đổi tên:").pack(anchor="w", padx=12, pady=(14, 4))
        var_cu = tk.StringVar(value=self.var_sheet.get() if self.var_sheet.get() in ten_sheets else ten_sheets[0])
        ttk.Combobox(top, textvariable=var_cu, values=ten_sheets, state="readonly", width=36).pack(padx=12)
        ttk.Label(top, text="Tên mới:").pack(anchor="w", padx=12, pady=(10, 4))
        var_moi = tk.StringVar(value=var_cu.get())
        ent = ttk.Entry(top, textvariable=var_moi, width=40)
        ent.pack(padx=12)
        def _doi():
            ten_cu = var_cu.get()
            ten_moi = var_moi.get().strip()
            if not ten_moi:
                messagebox.showwarning("Thiếu tên", "Gõ tên mới trước.", parent=top)
                return
            try:
                wb = load_workbook(p)
                if ten_cu not in wb.sheetnames:
                    messagebox.showerror("Không tìm thấy sheet", f'Sheet "{ten_cu}" không còn tồn tại.', parent=top)
                    wb.close()
                    return
                if ten_moi != ten_cu and ten_moi in wb.sheetnames:
                    messagebox.showerror("Đã tồn tại", f'Sheet "{ten_moi}" đã có sẵn trong file.', parent=top)
                    wb.close()
                    return
                wb[ten_cu].title = ten_moi
                wb.save(p)
                wb.close()
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Lỗi", str(e), parent=top)
                return
            top.destroy()
            self._nap_danh_sach_sheet()
            self.var_sheet.set(ten_moi)
            self._tai_du_lieu_sheet()
            self.var_status.set(f'✔ Đã đổi tên sheet "{ten_cu}" -> "{ten_moi}".')
        ttk.Button(top, text="Đổi tên", style="Accent.TButton", command=_doi).pack(pady=14)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    def _bat_dau_sua_o(self, event: tk.Event) -> None:
        if self._editor is not None:
            self._huy_sua_o()
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        x, y, w, h = self.tree.bbox(row_id, col_id)
        gia_tri_hien_tai = self.tree.set(row_id, col_id)
        self._editor = tk.Entry(self.tree)
        self._editor.insert(0, gia_tri_hien_tai)
        self._editor.select_range(0, "end")
        self._editor.focus_set()
        self._editor.place(x=x, y=y, width=w, height=h)
        self._editor_row_col = (row_id, col_id)
        self._editor.bind("<Return>", lambda e: self._xac_nhan_sua_o())
        self._editor.bind("<KP_Enter>", lambda e: self._xac_nhan_sua_o())
        self._editor.bind("<Escape>", lambda e: self._huy_sua_o())
        self._editor.bind("<FocusOut>", lambda e: self._xac_nhan_sua_o())
    def _xac_nhan_sua_o(self) -> None:
        if self._editor is None:
            return
        gia_tri_moi = self._editor.get()
        row_id, col_id = self._editor_row_col
        if self.tree.exists(row_id):
            self.tree.set(row_id, col_id, gia_tri_moi)
        self._huy_sua_o()
    def _huy_sua_o(self) -> None:
        if self._editor is not None:
            self._editor.destroy()
            self._editor = None
    def _o_tu_toa_do(self, event: tk.Event) -> tuple[int, int] | None:
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id or col_id == "#0":
            return None
        children = self.tree.get_children()
        try:
            r_idx = children.index(row_id)
        except ValueError:
            return None
        try:
            c_idx = int(col_id.replace("#", ""))
        except ValueError:
            return None
        return (r_idx, c_idx)
    def _bat_dau_chon_vung(self, event: tk.Event) -> None:
        o = self._o_tu_toa_do(event)
        if o is None:
            return
        self.tree.focus_set()
        self._sel_anchor = o
        self._sel_end = o
        self._cap_nhat_highlight()
    def _keo_chon_vung(self, event: tk.Event) -> None:
        if self._sel_anchor is None:
            return
        o = self._o_tu_toa_do(event)
        if o is None:
            return
        if o != self._sel_end:
            self._sel_end = o
            self._cap_nhat_highlight()
    def _ket_thuc_chon_vung(self, event: tk.Event) -> None:
        pass
    def _shift_click_chon_vung(self, event: tk.Event) -> None:
        if self._sel_anchor is None:
            self._sel_anchor = self._o_tu_toa_do(event)
        o = self._o_tu_toa_do(event)
        if o is None:
            return
        self._sel_end = o
        self._cap_nhat_highlight()
    def _pham_vi_chon(self) -> tuple[int, int, int, int] | None:
        if self._sel_anchor is None or self._sel_end is None:
            return None
        children = self.tree.get_children()
        if not children:
            return None
        r1, c1 = self._sel_anchor
        r2, c2 = self._sel_end
        r_min, r_max = sorted((r1, r2))
        c_min, c_max = sorted((c1, c2))
        r_max = min(r_max, len(children) - 1)
        c_max = min(c_max, self._so_cot_dang_hien)
        return (r_min, r_max, c_min, c_max)
    def _cap_nhat_highlight(self) -> None:
        pham_vi = self._pham_vi_chon()
        if pham_vi is None:
            return
        r_min, r_max, c_min, c_max = pham_vi
        children = self.tree.get_children()
        self.tree.selection_set(children[r_min:r_max + 1])
        for i in range(1, self._so_cot_dang_hien + 1):
            chu = get_column_letter(i)
            self.tree.heading(f"C{i}", text=f"[{chu}]" if c_min <= i <= c_max else chu)
        so_dong = r_max - r_min + 1
        so_cot = c_max - c_min + 1
        self.var_status.set(
            f"Đã chọn {so_dong} dòng × {so_cot} cột — Ctrl+C copy, Ctrl+V dán, Delete xoá nội dung "
            "(hoặc chuột phải để xem menu)."
        )
    def _mo_menu_chuot_phai(self, event: tk.Event) -> None:
        o = self._o_tu_toa_do(event)
        if o is not None and (self._sel_anchor is None or not self._trong_vung_dang_chon(o)):
            self._sel_anchor = o
            self._sel_end = o
            self._cap_nhat_highlight()
        try:
            self._menu_chuot_phai.tk_popup(event.x_root, event.y_root)
        finally:
            self._menu_chuot_phai.grab_release()
    def _trong_vung_dang_chon(self, o: tuple[int, int]) -> bool:
        pham_vi = self._pham_vi_chon()
        if pham_vi is None:
            return False
        r_min, r_max, c_min, c_max = pham_vi
        r, c = o
        return r_min <= r <= r_max and c_min <= c <= c_max
    def _sao_chep(self, event: tk.Event | None = None) -> None:
        pham_vi = self._pham_vi_chon()
        if pham_vi is None:
            messagebox.showinfo("Chưa chọn ô", "Click hoặc kéo chuột để chọn ít nhất 1 ô trước khi copy.")
            return
        r_min, r_max, c_min, c_max = pham_vi
        children = self.tree.get_children()
        dong_text = []
        for r in range(r_min, r_max + 1):
            row_id = children[r]
            gia_tri = [str(self.tree.set(row_id, f"C{c}")) for c in range(c_min, c_max + 1)]
            dong_text.append("\t".join(gia_tri))
        self.clipboard_clear()
        self.clipboard_append("\n".join(dong_text))
        self.var_status.set(f"Đã copy {r_max - r_min + 1} dòng × {c_max - c_min + 1} cột vào clipboard.")
    def _dan(self, event: tk.Event | None = None) -> None:
        if not self.tree["columns"]:
            messagebox.showwarning("Chưa tải sheet", "Tải dữ liệu 1 sheet trước khi dán.")
            return
        try:
            noi_dung = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Clipboard trống", "Không có dữ liệu trong clipboard để dán.")
            return
        if noi_dung == "":
            return
        pham_vi = self._pham_vi_chon()
        r_bat_dau = pham_vi[0] if pham_vi else 0
        c_bat_dau = pham_vi[2] if pham_vi else 1
        dong_dan = noi_dung.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if dong_dan and dong_dan[-1] == "":
            dong_dan.pop()
        if not dong_dan:
            return
        so_dong_can = r_bat_dau + len(dong_dan)
        while len(self.tree.get_children()) < so_dong_can:
            self.tree.insert("", "end", values=[""] * self._so_cot_dang_hien)
        children = self.tree.get_children()
        cot_bi_cat_bot = False
        cot_dan_toi_da = c_bat_dau
        for i, dong in enumerate(dong_dan):
            row_id = children[r_bat_dau + i]
            o_trong_dong = dong.split("\t")
            for j, gia_tri in enumerate(o_trong_dong):
                c = c_bat_dau + j
                if c > self._so_cot_dang_hien:
                    cot_bi_cat_bot = True
                    continue
                self.tree.set(row_id, f"C{c}", gia_tri)
                cot_dan_toi_da = max(cot_dan_toi_da, c)
        self._sel_anchor = (r_bat_dau, c_bat_dau)
        self._sel_end = (r_bat_dau + len(dong_dan) - 1, cot_dan_toi_da)
        self._lam_moi_zebra()
        self._cap_nhat_highlight()
        canh_bao = " (⚠️ có cột bị cắt bớt vì vượt quá số cột đang hiển thị)" if cot_bi_cat_bot else ""
        self.var_status.set(
            f"Đã dán {len(dong_dan)} dòng vào bảng.{canh_bao} Nhớ bấm '💾 Lưu vào file' để lưu thật sự."
        )
    def _xoa_noi_dung_vung(self, event: tk.Event | None = None) -> None:
        pham_vi = self._pham_vi_chon()
        if pham_vi is None:
            return
        r_min, r_max, c_min, c_max = pham_vi
        children = self.tree.get_children()
        for r in range(r_min, r_max + 1):
            row_id = children[r]
            for c in range(c_min, c_max + 1):
                self.tree.set(row_id, f"C{c}", "")
        self.var_status.set(f"Đã xoá nội dung {r_max - r_min + 1} dòng × {c_max - c_min + 1} cột đã chọn.")
    def _lam_moi_zebra(self) -> None:
        for idx, item in enumerate(self.tree.get_children()):
            self.tree.item(item, tags=("even" if idx % 2 == 0 else "odd",))
    def _them_dong(self) -> None:
        if not self.tree["columns"]:
            messagebox.showwarning("Chưa tải sheet", "Tải dữ liệu 1 sheet trước khi thêm dòng.")
            return
        self.tree.insert("", "end", values=[""] * self._so_cot_dang_hien)
        self._lam_moi_zebra()
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])
    def _xoa_dong(self) -> None:
        self._huy_sua_o()
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chưa chọn dòng", "Chọn ít nhất 1 dòng trong bảng trước (click vào dòng, giữ Ctrl để chọn nhiều).")
            return
        if not messagebox.askyesno("Xác nhận xoá", f"Xoá {len(sel)} dòng đã chọn?"):
            return
        for item in sel:
            self.tree.delete(item)
        self._lam_moi_zebra()
    def _luu_vao_file(self) -> None:
        self._huy_sua_o()
        # BỔ SUNG — chặn cứng, đây là điểm sửa bug mất dữ liệu: nếu sheet
        # đang mở bị cắt bớt lúc tải (nhiều hơn MAX_ROWS_CANH_BAO dòng
        # hoặc MAX_COLS cột trong file thật), TUYỆT ĐỐI không cho Lưu —
        # trước đây Lưu sẽ XOÁ SẠCH sheet cũ rồi ghi lại CHỈ ĐÚNG những gì
        # đang có trên lưới, làm mất toàn bộ phần chưa tải lên hiển thị.
        if self._sheet_dang_tai_bi_cat:
            messagebox.showerror(
                "Không thể lưu — sheet quá lớn",
                f'Sheet "{self.var_sheet.get()}" có {self._so_dong_that_cua_sheet} dòng × '
                f"{self._so_cot_that_cua_sheet} cột trong file thật, nhưng khu vực này chỉ tải xem tối đa "
                f"{self.MAX_ROWS_CANH_BAO} dòng × {self.MAX_COLS} cột.\n\n"
                "Để tránh mất dữ liệu (phần chưa tải lên sẽ bị xoá nếu lưu), nút Lưu bị khoá cho sheet này.\n"
                "Sửa sheet lớn bằng Excel, hoặc dùng các nút nhập/xoá dữ liệu chuyên dụng ở tab "
                "'🚀 Chạy pipeline' (chúng xử lý toàn bộ dữ liệu, không qua lưới xem này).",
            )
            return
        path = self.get_file_path()
        sheet_name = self.var_sheet.get()
        if not path or not sheet_name:
            messagebox.showwarning("Thiếu thông tin", "Chưa chọn file/sheet để lưu.")
            return
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy file:\n{p}")
            return
        rows = [self.tree.item(i, "values") for i in self.tree.get_children()]
        try:
            wb = load_workbook(p)
            if sheet_name in wb.sheetnames:
                vi_tri_cu = wb.sheetnames.index(sheet_name)
                del wb[sheet_name]
                sh = wb.create_sheet(sheet_name, vi_tri_cu)
            else:
                sh = wb.create_sheet(sheet_name)
            for r_idx, row_vals in enumerate(rows, start=1):
                for c_idx, v in enumerate(row_vals, start=1):
                    cell = sh.cell(row=r_idx, column=c_idx, value=(v if v != "" else None))
                    cell.number_format = "@"
            wb.save(p)
            wb.close()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(
                "Lỗi khi lưu",
                f"{e}\n\nKiểm tra xem file có đang mở trong Excel không — đóng file lại rồi thử lại.",
            )
            return
        self.var_status.set(f'✔ Đã lưu {len(rows)} dòng vào sheet "{sheet_name}" trong file.')
        messagebox.showinfo("Đã lưu", f'Đã lưu {len(rows)} dòng vào sheet "{sheet_name}".')
class KiemTraFrame(ttk.Frame):
    """Tab '🔍 Kiểm tra & Đối chiếu' — quy trình 1 lần:
    ① Điền dữ liệu (map hết vào tab TSKT, CHƯA xuất)
    ② Kiểm tra + CHỈNH SỬA ngay trên tool: sửa tay từng ô, lấy giá trị PIM
       cũ, tự động lấp ô tool để trống, nhập ĐƠN VỊ hàng loạt. Mọi chỉnh
       sửa giữ trong bộ nhớ, tính lại tức thì — chưa ghi file nào.
    ③ Xuất file import: đọc lại + kiểm tra lần cuối -> xác nhận -> xuất
       ĐÚNG 1 LẦN (thư mục .xlsx mở được ngay + .zip), đồng thời ghi chỉnh
       sửa + đơn vị vào workspace."""
    MAX_DONG_HIEN = 3000
    LOC_CAN_XEM = "(cần xem: khác + tool trống)"
    LOC_TAT_CA = "(tất cả)"
    LOC_DA_SUA = "✎ Đã sửa tay"
    MAU = {TRANG_THAI_KHAC: "#fde2df", TRANG_THAI_TOOL_TRONG: "#ffedd5", TRANG_THAI_PIM_TRONG: "#e3ecfd",
           TRANG_THAI_DON_VI: "#e6f4ea", TRANG_THAI_BO_QUA: "#eeeeee", "CAO": "#fde2df", "TRUNG BÌNH": "#fff4d6", "THÔNG TIN": "#eef1f8"}
    def __init__(self, parent: tk.Widget, app: "PimApp"):
        super().__init__(parent)
        self.app = app
        self.raw: Optional[dict] = None
        self.kq: Optional[dict] = None
        self._sua: Dict[Tuple[str, str, str], str] = {}
        self._bien_dv: Dict[Tuple[str, str], tk.StringVar] = {}
        self._dv_ap: Dict[Tuple[str, str], str] = {}
        self._rong: Dict[str, list] = {}
        self._bien_rong: Dict[str, Tuple[tk.StringVar, tk.StringVar]] = {}
        self._cfg_cua: Optional[str] = None
        self._file_moi: List[str] = []
        self._file_pim: str = ""
        self._nhan_sang_khoa: Dict[str, str] = {}
        self._khac_hien: List[list] = []
        self._ct_hien: List[list] = []
        self._ct_khoa: Optional[str] = None
        self._ds_khoa_hien: List[str] = []
        self._dang_chay = False
        self._editor: Optional[tk.Entry] = None
        self._build()
    # ================================================================ UI
    def _build(self) -> None:
        frm_b = ttk.Frame(self)
        frm_b.pack(fill="x", padx=10, pady=(10, 4))
        self.btn_dien = ttk.Button(frm_b, text="① Điền dữ liệu (map hết, chưa xuất)", style="Secondary.TButton",
                                   command=self._buoc_dien)
        self.btn_dien.pack(side="left")
        self.btn_kt = ttk.Button(frm_b, text="② 🔍 Kiểm tra / tải lại", style="Secondary.TButton",
                                 command=self._buoc_kiem_tra)
        self.btn_kt.pack(side="left", padx=(6, 0))
        self.btn_xuat = ttk.Button(frm_b, text="③ 📤 Kiểm tra lần cuối & XUẤT FILE IMPORT...", style="Accent.TButton",
                                   command=self._buoc_xuat)
        self.btn_xuat.pack(side="left", padx=(6, 0))
        self.progress = ttk.Progressbar(frm_b, mode="indeterminate", length=140)
        self.progress.pack(side="left", padx=(14, 0))
        self.btn_excel = ttk.Button(frm_b, text="💾 Lưu kết quả kiểm tra ra Excel...", command=self._luu_excel)
        self.btn_excel.pack(side="right")
        self._nut = [self.btn_dien, self.btn_kt, self.btn_xuat, self.btn_excel]
        frm_n = ttk.Frame(self)
        frm_n.pack(fill="x", padx=10)
        ttk.Label(frm_n, text="Giá trị MỚI:", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.var_moi = tk.StringVar(value="workspace")
        ttk.Radiobutton(frm_n, text="Tab TSKT trong workspace (trước khi xuất — sửa được)", variable=self.var_moi,
                        value="workspace").grid(row=0, column=1, sticky="w", padx=6)
        ttk.Radiobutton(frm_n, text="File đã xuất (.zip/.xlsx) — chỉ xem:", variable=self.var_moi,
                        value="file").grid(row=0, column=2, sticky="w")
        ttk.Button(frm_n, text="Chọn...", command=self._chon_file_moi).grid(row=0, column=3, padx=4)
        self.var_ten_moi = tk.StringVar(value="(chưa chọn)")
        ttk.Label(frm_n, textvariable=self.var_ten_moi, foreground="#777").grid(row=0, column=4, sticky="w")
        ttk.Label(frm_n, text="Spec PIM CŨ:", style="Section.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_cu = tk.StringVar(value="workspace")
        ttk.Radiobutton(frm_n, text=f'Sheet "{CFG_sheet_spec_pim}" (đã nạp)', variable=self.var_cu,
                        value="workspace").grid(row=1, column=1, sticky="w", padx=6, pady=(4, 0))
        ttk.Radiobutton(frm_n, text="File export PIM chọn thẳng:", variable=self.var_cu,
                        value="file").grid(row=1, column=2, sticky="w", pady=(4, 0))
        ttk.Button(frm_n, text="Chọn...", command=self._chon_file_pim).grid(row=1, column=3, padx=4, pady=(4, 0))
        self.var_ten_pim = tk.StringVar(value="(chưa chọn)")
        ttk.Label(frm_n, textvariable=self.var_ten_pim, foreground="#777").grid(row=1, column=4, sticky="w")
        ttk.Radiobutton(frm_n, text="Không đối chiếu", variable=self.var_cu, value="khong").grid(
            row=1, column=5, sticky="w", padx=(10, 0), pady=(4, 0))
        self.var_tt = tk.StringVar(value="Bấm ① để map dữ liệu (chưa xuất), hoặc ② nếu tab TSKT đã có dữ liệu.")
        ttk.Label(self, textvariable=self.var_tt, foreground="#555").pack(anchor="w", padx=10, pady=(6, 2))
        # ---- thẻ số liệu
        self.frm_the = tk.Frame(self, background="#eef1f8")
        self.frm_the.pack(fill="x", padx=10, pady=(2, 4))
        self._the: Dict[str, tk.StringVar] = {}
        for key, nhan, mau, dich in (
            ("so_dong", "Dòng sẽ xuất", "#2f6fed", None),
            (TRANG_THAI_KHAC, "Ô khác spec PIM", "#c0392b", ("khac", TRANG_THAI_KHAC)),
            (TRANG_THAI_TOOL_TRONG, "Tool trống (PIM có)", "#d35400", ("khac", TRANG_THAI_TOOL_TRONG)),
            (TRANG_THAI_DON_VI, "Chỉ thêm đơn vị", "#1f9d55", ("khac", TRANG_THAI_DON_VI)),
            (TRANG_THAI_PIM_TRONG, "Tool thêm mới", "#2557c7", ("khac", TRANG_THAI_PIM_TRONG)),
            ("thieu_model", "Thiếu model_code", "#c0392b", ("cb", "Thiếu model_code")),
            ("thieu_cate", "Thiếu category_code", "#c0392b", ("cb", "Thiếu category_code")),
            ("chua_don_vi", "Ô chưa có đơn vị", "#b7791f", ("dv", None)),
            ("khong_co_pim", "Không có trong PIM", "#b7791f", ("cb", "Không có trong file PIM")),
            ("rong_tong", "Ô Không/Đang cập nhật", "#5b6472", ("rong", None)),
            ("da_sua", "Ô sửa tay", "#6b46c1", ("khac", self.LOC_DA_SUA)),
        ):
            v = tk.StringVar(value="–")
            self._the[key] = v
            f = tk.Frame(self.frm_the, background="white", padx=9, pady=5, cursor="hand2" if dich else "")
            f.pack(side="left", padx=(0, 5))
            l1 = tk.Label(f, textvariable=v, font=("Segoe UI", 14, "bold"), foreground=mau, background="white")
            l1.pack(anchor="w")
            l2 = tk.Label(f, text=nhan, font=("Segoe UI", 9), foreground="#5b6472", background="white")
            l2.pack(anchor="w")
            if dich:
                for w in (f, l1, l2):
                    w.bind("<Button-1>", lambda e, d=dich: self._nhay_toi(d))
        # ---- thanh chỉnh sửa
        frm_s = tk.Frame(self, background="#f3eefe", padx=8, pady=5)
        frm_s.pack(fill="x", padx=10, pady=(0, 4))
        tk.Label(frm_s, text="✎ Chỉnh sửa:", background="#f3eefe", font=("Segoe UI", 10, "bold"),
                 foreground="#6b46c1").pack(side="left")
        for txt, cmd in (
            ("⬅ Lấy PIM cũ (dòng chọn)", lambda: self._lay_pim_cu(False)),
            ("⬅ Lấy PIM cũ (mọi dòng đang lọc)", lambda: self._lay_pim_cu(True)),
            ("✏️ Sửa giá trị dòng chọn...", self._sua_tay_nhieu),
            ("↺ Bỏ sửa dòng chọn", self._bo_sua),
            ("🪄 Tự động: ô tool trống → lấy PIM cũ", self._tu_dong_tool_trong),
            ("🗑 Xoá hết sửa tay", self._xoa_het_sua),
        ):
            ttk.Button(frm_s, text=txt, command=cmd).pack(side="left", padx=(6, 0))
        self.var_cho = tk.StringVar(value="")
        tk.Label(frm_s, textvariable=self.var_cho, background="#f3eefe", foreground="#6b46c1").pack(side="left",
                                                                                                     padx=10)
        self.var_ghi_ws = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_s, text="Khi xuất: ghi thêm vào workspace (chậm ~1 phút, thường không cần)",
                        variable=self.var_ghi_ws).pack(side="right")
        # ---- notebook con
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        # 0 Cảnh báo
        t1 = ttk.Frame(self.nb)
        self.nb.add(t1, text="⚠️ Cảnh báo")
        f1 = ttk.Frame(t1)
        f1.pack(fill="x", pady=4)
        ttk.Label(f1, text="Loại:").pack(side="left")
        self.var_loai_cb = tk.StringVar(value=self.LOC_TAT_CA)
        self.cbo_loai_cb = ttk.Combobox(f1, textvariable=self.var_loai_cb, state="readonly", width=52)
        self.cbo_loai_cb.pack(side="left", padx=4)
        self.cbo_loai_cb.bind("<<ComboboxSelected>>", lambda e: self._ve_cb())
        self.var_dem_cb = tk.StringVar()
        ttk.Label(f1, textvariable=self.var_dem_cb, foreground="#555").pack(side="left", padx=10)
        self.tree_cb = self._tao_tree(t1, [("MỨC", 90), ("LOẠI", 320), ("SKU / MODEL", 170), ("NH", 60),
                                           ("MÃ TSKT", 230), ("CHI TIẾT", 560)])
        self.tree_cb.bind("<Double-1>", self._dbl_cb)
        # 1 Khác spec
        t2 = ttk.Frame(self.nb)
        self.nb.add(t2, text="≠ Khác spec PIM (sửa được)")
        f2 = ttk.Frame(t2)
        f2.pack(fill="x", pady=4)
        ttk.Label(f2, text="Loại:").pack(side="left")
        self.var_loai = tk.StringVar(value=self.LOC_CAN_XEM)
        cbo = ttk.Combobox(f2, textvariable=self.var_loai, state="readonly", width=30, values=[
            self.LOC_CAN_XEM, self.LOC_TAT_CA, TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG, TRANG_THAI_DON_VI,
            TRANG_THAI_PIM_TRONG, TRANG_THAI_BO_QUA, self.LOC_DA_SUA])
        cbo.pack(side="left", padx=4)
        cbo.bind("<<ComboboxSelected>>", lambda e: self._ve_khac())
        ttk.Label(f2, text="  Mã TSKT:").pack(side="left")
        self.var_ma = tk.StringVar(value=self.LOC_TAT_CA)
        self.cbo_ma = ttk.Combobox(f2, textvariable=self.var_ma, state="readonly", width=34)
        self.cbo_ma.pack(side="left", padx=4)
        self.cbo_ma.bind("<<ComboboxSelected>>", lambda e: self._ve_khac())
        ttk.Label(f2, text="  Tìm:").pack(side="left")
        self.var_tim = tk.StringVar()
        ent = ttk.Entry(f2, textvariable=self.var_tim, width=20)
        ent.pack(side="left", padx=4)
        ent.bind("<Return>", lambda e: self._ve_khac())
        ttk.Button(f2, text="Lọc", command=self._ve_khac).pack(side="left")
        self.var_dem_khac = tk.StringVar()
        ttk.Label(f2, textvariable=self.var_dem_khac, foreground="#555").pack(side="left", padx=8)
        ttk.Label(t2, text="Bấm vào ô số của FILTER = mở bảng giải nghĩa mã + chọn lại option · double-click TOOL MỚI = "
                  "sửa tay · double-click PIM CŨ = lấy giá trị PIM cũ · double-click cột khác = xem cả SKU.",
                  foreground="#6b46c1").pack(anchor="w")
        self.tree_khac = self._tao_tree(t2, [("SKU / MODEL", 130), ("NH", 45), ("MÃ TSKT", 200), ("TÊN", 140),
                                             ("PIM CŨ", 200), ("TOOL MỚI", 200), ("GIẢI NGHĨA (FILTER)", 260),
                                             ("LOẠI", 180), ("SỬA", 40)])
        self.tree_khac.bind("<Double-1>", self._dbl_khac)
        self.tree_khac.bind("<ButtonRelease-1>", self._click_khac, add="+")
        # 2 Xem theo SKU
        t3 = ttk.Frame(self.nb)
        self.nb.add(t3, text="🔎 Xem theo SKU")
        trai = ttk.Frame(t3)
        trai.pack(side="left", fill="y", pady=4)
        self.var_tim_sku = tk.StringVar()
        e3 = ttk.Entry(trai, textvariable=self.var_tim_sku, width=30)
        e3.pack(fill="x")
        e3.bind("<KeyRelease>", lambda e: self._ve_ds_sku())
        self.var_chi_khac = tk.BooleanVar(value=True)
        ttk.Checkbutton(trai, text="Chỉ SKU cần xem", variable=self.var_chi_khac,
                        command=self._ve_ds_sku).pack(anchor="w", pady=2)
        self.lst = tk.Listbox(trai, width=34, font=("Consolas", 10), activestyle="none", exportselection=False)
        self.lst.pack(fill="y", expand=True)
        self.lst.bind("<<ListboxSelect>>", lambda e: self._ve_chi_tiet())
        phai = ttk.Frame(t3)
        phai.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=4)
        self.var_tieu_de_sku = tk.StringVar(value="Chọn 1 SKU bên trái.")
        ttk.Label(phai, textvariable=self.var_tieu_de_sku, style="Section.TLabel").pack(anchor="w")
        self.tree_ct = self._tao_tree(phai, [("MÃ TSKT", 210), ("TÊN", 150), ("PIM CŨ", 220), ("TOOL MỚI", 220),
                                             ("GIẢI NGHĨA (FILTER)", 260), ("TRẠNG THÁI", 170), ("SỬA", 40)])
        self.tree_ct.bind("<Double-1>", self._dbl_ct)
        self.tree_ct.bind("<ButtonRelease-1>", self._click_ct, add="+")
        # 3 Đơn vị hàng loạt
        t4 = ttk.Frame(self.nb)
        self.nb.add(t4, text="📏 Đơn vị hàng loạt")
        f4 = ttk.Frame(t4)
        f4.pack(fill="x", pady=4)
        ttk.Button(f4, text="✨ Điền gợi ý (cột còn số trơn, chưa chọn)", command=self._dv_goi_y).pack(side="left")
        ttk.Label(f4, text="   Đặt cùng 1 đơn vị cho mọi cột đang hiện:").pack(side="left")
        self.var_dv_nhanh = tk.StringVar(value="cm")
        ttk.Combobox(f4, textvariable=self.var_dv_nhanh, values=DON_VI_LUA_CHON, width=7).pack(side="left", padx=4)
        ttk.Button(f4, text="Đặt", command=lambda: self._dv_dat_tat_ca(self.var_dv_nhanh.get())).pack(side="left")
        ttk.Button(f4, text="🧹 Bỏ chọn", command=lambda: self._dv_dat_tat_ca("")).pack(side="left", padx=(8, 0))
        self.var_dv_tat_ca = tk.BooleanVar(value=False)
        ttk.Checkbutton(f4, text="Hiện tất cả cột TSKT", variable=self.var_dv_tat_ca,
                        command=self._ve_don_vi).pack(side="left", padx=(12, 0))
        ttk.Button(f4, text="▶ ÁP ĐƠN VỊ & XEM LẠI NGAY", style="Accent.TButton",
                   command=self._dv_ap_ngay).pack(side="right")
        ttk.Label(t4, text="Chỉ thêm vào ô SỐ TRƠN (9 → 9 kg). Ô đã có chữ giữ nguyên, cột FILTER không bị đụng. "
                  "Bấm ÁP để xem kết quả ngay (chưa ghi file); đơn vị được ghi vào file xuất + sheet ĐƠN VỊ ở bước ③. "
                  "Số màu đỏ = cột còn ô số trơn chưa có đơn vị.", foreground="#555", wraplength=1250,
                  justify="left").pack(anchor="w")
        khung = ttk.Frame(t4)
        khung.pack(fill="both", expand=True, pady=(4, 0))
        self.cv_dv = tk.Canvas(khung, background="#ffffff", highlightthickness=0)
        vs = ttk.Scrollbar(khung, orient="vertical", command=self.cv_dv.yview)
        self.cv_dv.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.cv_dv.pack(side="left", fill="both", expand=True)
        self.frm_dv = tk.Frame(self.cv_dv, background="#ffffff")
        self.cv_dv.create_window((0, 0), window=self.frm_dv, anchor="nw")
        self.frm_dv.bind("<Configure>", lambda e: self.cv_dv.configure(scrollregion=self.cv_dv.bbox("all")))
        def _cuon(e):
            self.cv_dv.yview_scroll(int(-1 * (e.delta / 120)) if e.delta else 0, "units")
        self.cv_dv.bind("<Enter>", lambda e: self.cv_dv.bind_all("<MouseWheel>", _cuon))
        self.cv_dv.bind("<Leave>", lambda e: self.cv_dv.unbind_all("<MouseWheel>"))
        # 4 Không / Đang cập nhật / Hãng không công bố
        t5 = ttk.Frame(self.nb)
        self.nb.add(t5, text="🚫 Không / Đang cập nhật")
        f5 = ttk.Frame(t5)
        f5.pack(fill="x", pady=4)
        ttk.Button(f5, text="Tất cả → Để trống (không cập nhật)",
                   command=lambda: self._rong_dat_tat_ca(HD_TRONG)).pack(side="left")
        ttk.Button(f5, text="Tất cả → Giữ để import", command=lambda: self._rong_dat_tat_ca(HD_GIU)).pack(
            side="left", padx=6)
        ttk.Label(f5, text="   Thêm giá trị khác:").pack(side="left")
        self.var_rong_moi = tk.StringVar()
        ttk.Entry(f5, textvariable=self.var_rong_moi, width=24).pack(side="left", padx=4)
        ttk.Button(f5, text="➕ Thêm", command=self._rong_them).pack(side="left")
        ttk.Button(f5, text="▶ ÁP & XEM LẠI NGAY", style="Accent.TButton", command=self._rong_ap).pack(side="right")
        ttk.Label(t5, text="Áp cho mọi cột TSKT (không áp cho FILTER, không đè ô sửa tay). 'Để trống' = ô xuất ra "
                  "để trống, KHÔNG ghi đè giá trị đang có trên PIM (nên thử import 1 SKU để chắc PIM bỏ qua ô trống). "
                  "Lựa chọn tự lưu, lần sau map lại tự áp.", foreground="#555", wraplength=1250,
                  justify="left").pack(anchor="w")
        khung5 = ttk.Frame(t5)
        khung5.pack(fill="both", expand=True, pady=(4, 0))
        self.cv_rong = tk.Canvas(khung5, background="#ffffff", highlightthickness=0)
        vs5 = ttk.Scrollbar(khung5, orient="vertical", command=self.cv_rong.yview)
        self.cv_rong.configure(yscrollcommand=vs5.set)
        vs5.pack(side="right", fill="y")
        self.cv_rong.pack(side="left", fill="both", expand=True)
        self.frm_rong = tk.Frame(self.cv_rong, background="#ffffff")
        self.cv_rong.create_window((0, 0), window=self.frm_rong, anchor="nw")
        self.frm_rong.bind("<Configure>", lambda e: self.cv_rong.configure(scrollregion=self.cv_rong.bbox("all")))
    def _tao_tree(self, parent, cols: List[Tuple[str, int]]) -> ttk.Treeview:
        frm = ttk.Frame(parent)
        frm.pack(fill="both", expand=True)
        tree = ttk.Treeview(frm, show="headings", columns=[f"c{i}" for i in range(len(cols))], selectmode="extended")
        for i, (t, w) in enumerate(cols):
            tree.heading(f"c{i}", text=t)
            tree.column(f"c{i}", width=w, anchor="w", stretch=False)
        for k, m in self.MAU.items():
            tree.tag_configure(k, background=m)
        tree.tag_configure("sua", foreground="#6b46c1")
        vs = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        tree.bind("<Control-c>", lambda e, t=tree: self._copy_tree(t))
        return tree
    def _copy_tree(self, tree: ttk.Treeview) -> None:
        dong = ["\t".join(str(x) for x in tree.item(i, "values")) for i in tree.selection()]
        if dong:
            self.clipboard_clear()
            self.clipboard_append("\n".join(dong))
    # ================================================================ chạy nền
    def _chay(self, ham, xong, ten: str) -> None:
        if self._dang_chay:
            return
        self._dang_chay = True
        self.progress.start(12)
        self.var_tt.set(f"{ten}... (đừng đóng cửa sổ)")
        for b in self._nut:
            b.configure(state="disabled")
        self.app._dat_trang_thai_nut("disabled")
        q: "queue.Queue" = queue.Queue()
        def _nen():
            try:
                q.put(ham())
            except Exception as e:  # noqa: BLE001
                q.put({"loi": f"Lỗi ngoài dự kiến: {e}\n\n{traceback.format_exc(limit=3)}"})
        threading.Thread(target=_nen, daemon=True).start()
        def _cho():
            try:
                kq = q.get_nowait()
            except queue.Empty:
                self.after(150, _cho)
                return
            self._dang_chay = False
            self.progress.stop()
            for b in self._nut:
                b.configure(state="normal")
            self.app._dat_trang_thai_nut("normal")
            xong(kq)
        self.after(150, _cho)
    def _wb(self) -> Optional[Path]:
        p = Path(self.app.var_file.get().strip())
        if not p.exists():
            messagebox.showerror("Không tìm thấy workspace", str(p))
            return None
        return p
    def _out(self) -> Path:
        o = self.app.var_out.get().strip()
        return Path(o) if o else Path(self.app.var_file.get()).parent / "output"
    def _chon_file_moi(self) -> None:
        ds = filedialog.askopenfilenames(title="Chọn file đã xuất (.zip hoặc các file .xlsx)",
                                         filetypes=[("Zip / Excel", "*.zip *.xlsx"), ("Tất cả", "*.*")])
        if ds:
            self._file_moi = list(ds)
            self.var_moi.set("file")
            self.var_ten_moi.set(", ".join(Path(x).name for x in ds)[:90])
    def _chon_file_pim(self) -> None:
        f = filedialog.askopenfilename(title="Chọn file export PIM",
                                       filetypes=[("Excel files", "*.xlsx"), ("Tất cả", "*.*")])
        if f:
            self._file_pim = f
            self.var_cu.set("file")
            self.var_ten_pim.set(Path(f).name)
    # ================================================================ đơn vị: trạng thái
    def _dv_da_luu(self) -> Dict[Tuple[str, str], str]:
        if not self.kq:
            return {}
        return {(c["cate"], c["code"]): c["da_luu"] for c in self.kq.get("don_vi_cot", [])}
    def _dv_doi(self) -> Dict[Tuple[str, str], str]:
        """Đơn vị đã ÁP mà khác với sheet ĐƠN VỊ -> cần ghi khi xuất."""
        luu = self._dv_da_luu()
        return {k: v for k, v in self._dv_ap.items() if v != luu.get(k, "")}
    def _cap_nhat_cho(self) -> None:
        st = self.kq["stat"] if self.kq else {}
        txt = []
        if st.get("da_sua"):
            txt.append(f"{st['da_sua']} ô sửa tay")
        if st.get("them_dv"):
            txt.append(f"{st['them_dv']} ô thêm đơn vị")
        if st.get("rong_trong") or st.get("rong_thay"):
            txt.append(f"{st.get('rong_trong', 0) + st.get('rong_thay', 0)} ô Không/Đang cập nhật xử lý")
        self.var_cho.set(("Sẽ áp khi xuất: " + ", ".join(txt) + " (đã tự lưu)") if txt else "")
    # ================================================================ 3 bước
    def _buoc_dien(self) -> None:
        p = self._wb()
        if not p:
            return
        if (self._sua or self._dv_doi()) and not messagebox.askyesno(
                "① Điền lại dữ liệu",
                "Đang có chỉnh sửa chưa xuất. Điền lại sẽ map lại tab TSKT từ DATA SP — các ô sửa tay + đơn vị "
                "đang chọn VẪN được giữ và áp lại lên dữ liệu mới.\n\nTiếp tục?"):
            return
        out = self._out()
        app = self.app
        def _lam():
            old = sys.stdout
            sys.stdout = _StdoutRedirect(app._log_queue)
            try:
                kq = chay_tat_ca(p, out, xuat=False)
            finally:
                sys.stdout = old
            return {"loi": None if kq else "Pipeline dừng sớm — xem Nhật ký ở tab '🚀 Chạy pipeline'.", "kq": kq}
        def _xong(kq):
            if kq.get("loi"):
                messagebox.showerror("Điền dữ liệu", kq["loi"])
                self.var_tt.set("Điền dữ liệu có lỗi.")
                return
            self.var_moi.set("workspace")
            self._buoc_kiem_tra()
        self._chay(_lam, _xong, "① Đang map dữ liệu vào tab TSKT (chưa xuất)")
    def _buoc_kiem_tra(self, sau=None) -> None:
        p = self._wb()
        if not p:
            return
        if self.var_moi.get() == "file" and not self._file_moi:
            messagebox.showwarning("Thiếu file", "Chọn file đã xuất trước (nút 'Chọn...' dòng Giá trị MỚI).")
            return
        if self.var_cu.get() == "file" and not self._file_pim:
            messagebox.showwarning("Thiếu file", "Chọn file export PIM trước (nút 'Chọn...' dòng Spec PIM CŨ).")
            return
        moi = "workspace" if self.var_moi.get() == "workspace" else list(self._file_moi)
        cu = {"workspace": "workspace", "khong": "khong"}.get(self.var_cu.get(), self._file_pim)
        def _xong(raw):
            if raw.get("loi"):
                messagebox.showerror("Kiểm tra", raw["loi"])
                self.var_tt.set("Kiểm tra có lỗi.")
                return
            self.raw = raw
            if self._cfg_cua != str(p):
                cfg = doc_cau_hinh_kt(p)
                self._sua = cfg["sua"]
                self._dv_ap = cfg["don_vi"]
                self._rong = cfg["rong"]
                self._bien_rong.clear()
                self._cfg_cua = str(p)
            self._tinh_lai(dau=True)
            if sau:
                sau()
        self._chay(lambda: doc_du_lieu_kiem_tra(p, moi, cu), _xong, "② Đang đọc dữ liệu (chỉ đọc)")
    def _tinh_lai(self, dau: bool = False) -> None:
        if not self.raw:
            return
        self.kq = tinh_kiem_tra(self.raw, self._sua, {k: v for k, v in self._dv_ap.items() if v}, self._rong)
        if dau:
            # lần đầu: đơn vị đang chọn = đơn vị đã lưu ở sheet ĐƠN VỊ (nếu file cấu hình chưa có)
            for k, v in self._dv_da_luu().items():
                self._dv_ap.setdefault(k, v)
        self.kq["stat"]["rong_tong"] = sum(x["so_o"] for x in self.kq.get("gia_tri_rong", []))
        if self._cfg_cua:
            try:
                ghi_cau_hinh_kt(Path(self._cfg_cua), self._sua, self._dv_ap, self._rong)
            except Exception:  # noqa: BLE001
                pass
        self._hien_ket_qua(dau=dau)
    def _buoc_xuat(self) -> None:
        self.var_moi.set("workspace")
        self._buoc_kiem_tra(sau=self._xac_nhan_xuat)
    def _xac_nhan_xuat(self) -> None:
        kq = self.kq
        st = kq["stat"]
        dv_doi = self._dv_doi()
        dong = [f"Sẽ xuất {st.get('so_dong', 0)} dòng (đã kiểm tra lại từ workspace)."]
        if st.get("da_sua") or st.get("them_dv") or st.get("rong_trong") or st.get("rong_thay"):
            dong.append(f"Áp vào file xuất: {st.get('da_sua', 0)} ô sửa tay, {st.get('them_dv', 0)} ô thêm đơn vị, "
                        f"{st.get('rong_trong', 0)} ô để trống (không cập nhật), {st.get('rong_thay', 0)} ô thay thế.")
        canh = []
        for k, t in (("thieu_model", "SKU thiếu model_code"), ("thieu_cate", "SKU thiếu category_code"),
                     (TRANG_THAI_KHAC, "ô KHÁC spec PIM (sẽ ghi đè)"),
                     (TRANG_THAI_TOOL_TRONG, "ô tool để trống trong khi PIM đang có"),
                     ("chua_don_vi", "ô kích thước/khối lượng chưa có đơn vị"),
                     ("khong_co_pim", "dòng không có trong file PIM"), ("lech_model", "model_code khác PIM"),
                     ("dong_sot", "dòng không có trong IMPORT (sẽ tự bỏ)"),
                     ("filter_chu", "ô FILTER là chữ, không phải mã option")):
            if st.get(k):
                canh.append(f"  • {st[k]} {t}")
        if canh:
            dong += ["", "⚠️ CÒN CẢNH BÁO:"] + canh
        ghi = self.var_ghi_ws.get() and bool(self._sua or dv_doi or st.get("them_dv") or self._rong)
        if ghi:
            dong += ["", "Sau khi xuất sẽ ghi chỉnh sửa + đơn vị vào workspace (mất thêm ~1 phút)."]
        dong += ["", "XUẤT FILE IMPORT NGAY?"]
        if not messagebox.askyesno("③ Xác nhận xuất" + (" — CÒN CẢNH BÁO" if canh else ""), "\n".join(dong),
                                   icon="warning" if canh else "question"):
            self.var_tt.set("Đã huỷ xuất — tiếp tục chỉnh sửa ở các tab bên dưới.")
            return
        p = self._wb()
        out = self._out()
        sua = dict(self._sua)
        dv_ap = {k: v for k, v in self._dv_ap.items() if v}
        dv_ghi = dict(dv_ap)
        dv_ghi.update(dv_doi)
        ten_cate = {c: " ".join(b["title"].split()[2:]) for c, b in self.raw["bang"].items()}
        ten_tskt = {(c["cate"], c["code"]): c["ten"] for c in self.kq.get("don_vi_cot", [])}
        rong = {k: list(v) for k, v in self._rong.items()}
        def _lam():
            x = xuat_file_import_nhanh(p, out, sua, dv_ap, rong)
            if x.get("loi") or not ghi:
                return x
            y = luu_chinh_sua_vao_workspace(p, sua, dv_ghi, ten_cate, ten_tskt, rong)
            x["ghi_ws"] = y
            return x
        def _xong(x):
            if x.get("loi"):
                messagebox.showerror("Xuất file", x["loi"])
                return
            self.app._out_dir = Path(x["thu_muc"])
            self.app.btn_open_out.configure(state="normal")
            ds = "\n".join(f"  • {t} ({n} dòng)" for t, n in x["files"])
            them = f"\nĐã TỰ BỎ {x['bo_dong']} dòng không có trong IMPORT." if x["bo_dong"] else ""
            ws_txt = ""
            y = x.get("ghi_ws")
            if y:
                if y.get("loi"):
                    ws_txt = f"\n\n⚠️ File import ĐÃ xuất, nhưng KHÔNG ghi được vào workspace: {y['loi']}"
                else:
                    ws_txt = f"\n\n✔ Đã ghi {y['so_o']} ô vào workspace (nhật ký: sheet {CFG_sheet_chinh_sua})."
            self.var_tt.set(f"✔ Đã xuất {len(x['files'])} file vào {x['thu_muc']}")
            mo = messagebox.askyesno(
                "Đã xuất xong",
                f"Thư mục (mở được ngay, không cần giải nén):\n{x['thu_muc']}\n\n{ds}\n"
                f"(đã áp {x.get('so_o_sua', 0)} ô sửa tay, {x.get('so_o_dv', 0)} ô thêm đơn vị, "
                f"{x.get('so_o_rong', 0)} ô Không/Đang cập nhật theo lựa chọn){them}"
                f"\n(Kèm bản .zip: {Path(x['zip']).name}){ws_txt}\n\nMở thư mục ngay?")
            if mo:
                self.app._mo_thu_muc_ket_qua()
            if y and not y.get("loi"):
                self._buoc_kiem_tra()
        self._chay(_lam, _xong, "③ Đang xuất file import" + (" + ghi workspace" if ghi else ""))
    # ================================================================ chỉnh sửa
    def _co_the_sua(self) -> bool:
        if not self.raw or not self.raw.get("la_workspace"):
            messagebox.showinfo("Chưa sửa được",
                                "Chỉ sửa được khi 'Giá trị MỚI' = Tab TSKT trong workspace (bấm ② Kiểm tra trước).")
            return False
        return True
    def _o_chon(self, tat_ca_dang_loc: bool = False) -> List[Tuple[str, str, str, str, str]]:
        """-> [(cate, khoá, mã, PIM cũ, giá trị gốc tool)] của dòng đang chọn ở tab đang mở."""
        tab = self.nb.index("current")
        out = []
        if tab == 1:
            rows = self._khac_hien if tat_ca_dang_loc else [
                self._khac_hien[int(i)] for i in self.tree_khac.selection()]
            for r in rows:
                out.append((r[2], r[8], r[3], r[5], r[9]))
        elif tab == 2 and self._ct_khoa:
            ct = self.kq["chi_tiet"][self._ct_khoa]
            rows = self._ct_hien if tat_ca_dang_loc else [self._ct_hien[int(i)] for i in self.tree_ct.selection()]
            for d in rows:
                out.append((ct["cate"], self._ct_khoa, d[0], d[2], d[5]))
        else:
            messagebox.showinfo("Chọn dòng", "Mở tab '≠ Khác spec PIM' hoặc '🔎 Xem theo SKU' rồi chọn dòng cần sửa.")
        return out
    def _chuan_filter(self, code: str, gia_tri: str) -> str:
        """FILTER phải là MÃ option: chữ -> tra mã qua DATA PIM (không tra được thì giữ, sẽ bị cảnh báo)."""
        opt = self.raw["tra"].option_map if self.raw else {}
        out = []
        for t in re.split(r"[,|]", lam_sach_gia_tri_pim(gia_tri, code)):
            t = t.strip()
            if not t:
                continue
            tid = chuan_hoa_id(t)
            if not re.fullmatch(r"\d+", tid):
                tid = opt.get((code, t.lower()), t)
            if tid not in out:
                out.append(tid)
        return CFG.value_separator_filter.join(out)
    def _dat_sua(self, cate: str, khoa: str, code: str, gia_tri: str, goc: str) -> None:
        if la_cot_filter(code):
            gia_tri = self._chuan_filter(code, gia_tri)
        key = (cate, khoa, code)
        if chuan_hoa_key(gia_tri) == chuan_hoa_key(goc):
            self._sua.pop(key, None)
        else:
            self._sua[key] = chuan_hoa_key(gia_tri)
    def _lay_pim_cu(self, tat_ca: bool) -> None:
        if not self._co_the_sua():
            return
        ds = self._o_chon(tat_ca)
        ds = [x for x in ds if x[3]]
        if not ds:
            messagebox.showinfo("Không có gì để lấy", "Các dòng đang chọn/đang lọc không có giá trị PIM cũ.")
            return
        if tat_ca and not messagebox.askyesno("Lấy PIM cũ", f"Lấy giá trị PIM cũ cho {len(ds)} ô đang lọc?"):
            return
        for cate, khoa, code, old, goc in ds:
            self._dat_sua(cate, khoa, code, old, goc)
        self._tinh_lai()
    def _sua_tay_nhieu(self) -> None:
        if not self._co_the_sua():
            return
        ds = self._o_chon()
        if not ds:
            return
        from tkinter import simpledialog
        dau = self._sua.get(ds[0][:3], ds[0][4])
        v = simpledialog.askstring("Sửa giá trị", f"Giá trị mới cho {len(ds)} ô đang chọn "
                                   "(FILTER: mã số, nhiều mã cách nhau ', '):", initialvalue=dau, parent=self)
        if v is None:
            return
        for cate, khoa, code, _, goc in ds:
            self._dat_sua(cate, khoa, code, v, goc)
        self._tinh_lai()
    def _bo_sua(self) -> None:
        ds = self._o_chon()
        for cate, khoa, code, _, _ in ds:
            self._sua.pop((cate, khoa, code), None)
        if ds:
            self._tinh_lai()
    def _tu_dong_tool_trong(self) -> None:
        if not self._co_the_sua():
            return
        ds = [r for r in self.kq["khac"] if r[7] == TRANG_THAI_TOOL_TRONG and r[5]]
        if not ds:
            messagebox.showinfo("Không có ô nào", "Không có ô nào tool để trống mà PIM đang có giá trị.")
            return
        if not messagebox.askyesno("🪄 Chỉnh tự động",
                                   f"{len(ds)} ô tool để trống trong khi PIM đang có giá trị.\n\n"
                                   "Lấy giá trị PIM cũ điền vào (để import không làm mất dữ liệu PIM)?"):
            return
        for r in ds:
            self._dat_sua(r[2], r[8], r[3], r[5], r[9])
        self._tinh_lai()
    # ---- sửa trực tiếp trên ô (inline)
    def _sua_inline(self, tree: ttk.Treeview, iid: str, col: str, gia_tri: str, luu) -> None:
        self._huy_inline()
        bbox = tree.bbox(iid, col)
        if not bbox:
            return
        x, y, w, h = bbox
        e = tk.Entry(tree, font=("Segoe UI", 10), background="#fffbe6")
        e.insert(0, gia_tri)
        e.select_range(0, "end")
        e.place(x=x, y=y, width=max(w, 220), height=h)
        e.focus_set()
        self._editor = e
        def _ok(_=None):
            if self._editor is None:
                return
            v = e.get()
            self._huy_inline()
            luu(v)
        e.bind("<Return>", _ok)
        e.bind("<KP_Enter>", _ok)
        e.bind("<Escape>", lambda _e: self._huy_inline())
        e.bind("<FocusOut>", _ok)
    def _huy_inline(self) -> None:
        if self._editor is not None:
            w, self._editor = self._editor, None
            w.destroy()
    def _dbl_khac(self, event) -> None:
        iid = self.tree_khac.identify_row(event.y)
        col = self.tree_khac.identify_column(event.x)
        if not iid:
            return
        r = self._khac_hien[int(iid)]
        if col in ("#6", "#7") and la_cot_filter(r[3]):
            self._mo_filter(r[2], r[8], r[3], r[4], r[5], r[6], r[9], r[0], event, chon_list=True)
        elif col == "#6" and self._co_the_sua():
            self._sua_inline(self.tree_khac, iid, col, r[6],
                             lambda v: (self._dat_sua(r[2], r[8], r[3], v, r[9]), self._tinh_lai()))
        elif col == "#5" and r[5] and self._co_the_sua():
            self._dat_sua(r[2], r[8], r[3], r[5], r[9])
            self._tinh_lai()
        else:
            self._mo_sku(r[8])
    def _dbl_ct(self, event) -> None:
        iid = self.tree_ct.identify_row(event.y)
        col = self.tree_ct.identify_column(event.x)
        if not iid or not self._ct_khoa:
            return
        d = self._ct_hien[int(iid)]
        cate = self.kq["chi_tiet"][self._ct_khoa]["cate"]
        khoa = self._ct_khoa
        if col == "#3" and d[2] and self._co_the_sua():
            self._dat_sua(cate, khoa, d[0], d[2], d[5])
            self._tinh_lai()
        elif la_cot_filter(d[0]):
            self._mo_filter(cate, khoa, d[0], d[1], d[2], d[3], d[5], self.kq["chi_tiet"][khoa]["nhan"], event,
                            chon_list=True)
        elif self._co_the_sua():
            self._sua_inline(self.tree_ct, iid, "#4", d[3],
                             lambda v: (self._dat_sua(cate, khoa, d[0], v, d[5]), self._tinh_lai()))
    def _dbl_cb(self, event) -> None:
        iid = self.tree_cb.identify_row(event.y)
        if not iid:
            return
        vals = self.tree_cb.item(iid, "values")
        if vals and "đơn vị" in str(vals[1]):
            self.nb.select(3)
            return
        khoa = self._nhan_sang_khoa.get(str(vals[2])) if vals else None
        if khoa:
            self._mo_sku(khoa)
    # ================================================================ giải nghĩa FILTER
    def _giai_nghia(self, code: str, gia_tri: str) -> str:
        """'6, 27' -> '6=Tivi thông minh · 27=4K' (chỉ cho cột FILTER)."""
        if not self.raw or not la_cot_filter(code) or not gia_tri:
            return ""
        ten = self.raw.get("opt_ten", {})
        out = []
        for t in re.split(r"[,|]", lam_sach_gia_tri_pim(gia_tri, code)):
            t = chuan_hoa_id(t.strip())
            if t:
                out.append(f"{t}={ten.get((code, t), '❓KHÔNG CÓ trong DATA PIM')}")
        return " · ".join(out)
    def _click_khac(self, event) -> None:
        iid = self.tree_khac.identify_row(event.y)
        col = self.tree_khac.identify_column(event.x)
        if not iid or col not in ("#5", "#6", "#7"):
            return
        r = self._khac_hien[int(iid)]
        if la_cot_filter(r[3]):
            self._mo_filter(r[2], r[8], r[3], r[4], r[5], r[6], r[9], r[0], event)
    def _click_ct(self, event) -> None:
        iid = self.tree_ct.identify_row(event.y)
        col = self.tree_ct.identify_column(event.x)
        if not iid or not self._ct_khoa or col not in ("#3", "#4", "#5"):
            return
        d = self._ct_hien[int(iid)]
        if la_cot_filter(d[0]):
            ct = self.kq["chi_tiet"][self._ct_khoa]
            self._mo_filter(ct["cate"], self._ct_khoa, d[0], d[1], d[2], d[3], d[5], ct["nhan"], event)
    def _mo_filter(self, cate, khoa, code, ten, old, moi, goc, nhan, event, chon_list: bool = False) -> None:
        """Cửa sổ nhỏ: giải nghĩa mã FILTER (PIM cũ / tool mới) + danh sách
        TOÀN BỘ option của filter (tìm, tick chọn) để sửa lại cho đúng web."""
        if not self.raw:
            return
        ten_opt = self.raw.get("opt_ten", {})
        ds_opt = self.raw.get("opt_ds", {}).get(code, [])
        def _ma(v):
            return [chuan_hoa_id(t.strip()) for t in re.split(r"[,|]", lam_sach_gia_tri_pim(v, code)) if t.strip()]
        ma_moi, ma_cu = _ma(moi), _ma(old)
        dang_dung = set(ma_moi) | set(ma_cu)
        ds_opt = sorted(ds_opt, key=lambda x: (x[0] not in dang_dung,
                                               int(x[0]) if x[0].isdigit() else 10 ** 12, x[0]))
        top = getattr(self, "_pop", None)
        if top is None or not top.winfo_exists():
            top = tk.Toplevel(self)
            top.transient(self.winfo_toplevel())
            top.bind("<Escape>", lambda e: top.destroy())
            self._pop = top
            x, y = event.x_root + 20, event.y_root + 10
            top.geometry(f"560x520+{max(0, x - 600) if x > self.winfo_screenwidth() - 600 else x}"
                         f"+{max(0, min(y, self.winfo_screenheight() - 560))}")
        for w in top.winfo_children():
            w.destroy()
        ten_f = ten or self.raw.get("ten_filter", {}).get(code, "")
        top.title(f"FILTER: {ten_f} — {code}")
        tk.Label(top, text=f"{ten_f}   ({code})", font=("Segoe UI", 11, "bold"), background="#2f6fed",
                 foreground="white", anchor="w", padx=10, pady=6).pack(fill="x")
        tk.Label(top, text=f"SKU: {nhan}", anchor="w", padx=10, foreground="#5b6472").pack(fill="x", pady=(4, 0))
        khung = tk.Frame(top, padx=10, pady=4)
        khung.pack(fill="x")
        for col_i, (tieu_de, ds, khac_ds) in enumerate((("TOOL MỚI (sẽ import)", ma_moi, ma_cu),
                                                        ("PIM CŨ (đang trên web)", ma_cu, ma_moi))):
            f = tk.Frame(khung)
            f.grid(row=0, column=col_i, sticky="nw", padx=(0, 16))
            tk.Label(f, text=tieu_de, font=("Segoe UI", 10, "bold")).pack(anchor="w")
            if not ds:
                tk.Label(f, text="(trống)", foreground="#777").pack(anchor="w")
            for m in ds:
                tn = ten_opt.get((code, m))
                if tn is None:
                    txt, mau = f"{m}  =  ❓ KHÔNG có trong DATA PIM", "#c0392b"
                else:
                    txt, mau = f"{m}  =  {tn}", ("#1f9d55" if m in khac_ds or not khac_ds else "#d35400")
                tk.Label(f, text=txt, foreground=mau, anchor="w", wraplength=250, justify="left").pack(anchor="w")
        tk.Label(top, text="Xanh = có ở cả 2 bên · Cam = chỉ có 1 bên · Đỏ = mã không tồn tại",
                 foreground="#777", padx=10, anchor="w").pack(fill="x")
        ttk.Separator(top).pack(fill="x", pady=4)
        f2 = tk.Frame(top, padx=10)
        f2.pack(fill="x")
        tk.Label(f2, text=f"Tất cả {len(ds_opt)} option — tick để chọn lại:").pack(side="left")
        var_tim = tk.StringVar()
        e = ttk.Entry(f2, textvariable=var_tim, width=18)
        e.pack(side="right")
        tk.Label(f2, text="Tìm:").pack(side="right")
        f3 = tk.Frame(top, padx=10)
        f3.pack(fill="both", expand=True)
        lst = tk.Listbox(f3, selectmode="multiple", activestyle="none", exportselection=False, font=("Segoe UI", 10))
        sb = ttk.Scrollbar(f3, orient="vertical", command=lst.yview)
        lst.configure(yscrollcommand=sb.set)
        lst.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        chon = set(ma_moi)
        hien: List[str] = []
        def _ve():
            lst.delete(0, "end")
            hien.clear()
            q = var_tim.get().strip().lower()
            for oc, tn in ds_opt:
                if q and q not in tn.lower() and q not in oc:
                    continue
                hien.append(oc)
                lst.insert("end", f"{oc}  —  {tn}" + ("   ← PIM cũ" if oc in ma_cu else "")
                           + ("   ← tool mới" if oc in ma_moi else ""))
                if oc in chon:
                    lst.selection_set("end")
        def _doi(_=None):
            hien_set = set(hien)
            sel = {hien[i] for i in lst.curselection()}
            chon.difference_update(hien_set - sel)
            chon.update(sel)
        lst.bind("<<ListboxSelect>>", _doi)
        var_tim.trace_add("write", lambda *a: _ve())
        _ve()
        f4 = tk.Frame(top, padx=10, pady=8)
        f4.pack(fill="x")
        def _dung():
            if not self._co_the_sua():
                return
            thu_tu = [oc for oc, _ in ds_opt if oc in chon] + [m for m in ma_moi if m in chon and
                                                                 m not in {oc for oc, _ in ds_opt}]
            self._dat_sua(cate, khoa, code, CFG.value_separator_filter.join(thu_tu), goc)
            top.destroy()
            self._tinh_lai()
        def _lay_cu():
            if not self._co_the_sua():
                return
            self._dat_sua(cate, khoa, code, old, goc)
            top.destroy()
            self._tinh_lai()
        ttk.Button(f4, text="✔ Dùng các option đã tick", style="Accent.TButton", command=_dung).pack(side="left")
        if old:
            ttk.Button(f4, text="⬅ Lấy PIM cũ", command=_lay_cu).pack(side="left", padx=6)
        ttk.Button(f4, text="Đóng (Esc)", command=top.destroy).pack(side="right")
        top.lift()
        if chon_list:
            e.focus_set()
    # ================================================================ Không / Đang cập nhật
    NHAN_HD = {HD_GIU: "Giữ — import như cũ", HD_TRONG: "Để trống — không cập nhật", HD_THAY: "Thay bằng…"}
    def _ve_rong(self) -> None:
        for w in self.frm_rong.winfo_children():
            w.destroy()
        if not self.kq:
            return
        ds = list(self.kq.get("gia_tri_rong", []))
        co = {d["khoa"] for d in ds}
        for k in self._rong:
            if k not in co:
                ds.append({"khoa": k, "so_o": 0, "hien": k, "cac_dang": [], "cot": []})
        tieu_de = ["Giá trị", "Số ô", "Các kiểu viết gặp trong data", "Có ở các cột (vd)", "XỬ LÝ", "Thay bằng"]
        for j, t in enumerate(tieu_de):
            tk.Label(self.frm_rong, text=t, font=("Segoe UI", 10, "bold"), background="#2f6fed", foreground="white",
                     padx=6, pady=4, anchor="w").grid(row=0, column=j, sticky="we")
        nhan_sang_hd = {v: k for k, v in self.NHAN_HD.items()}
        for i, d in enumerate(ds, start=1):
            k = d["khoa"]
            q = self._rong.get(k, [HD_GIU, ""])
            if k not in self._bien_rong:
                self._bien_rong[k] = (tk.StringVar(), tk.StringVar())
            v_hd, v_thay = self._bien_rong[k]
            v_hd.set(self.NHAN_HD.get(q[0], self.NHAN_HD[HD_GIU]))
            v_thay.set(q[1] if len(q) > 1 else "")
            nen = "#f3f5fb" if i % 2 else "#ffffff"
            mau = {HD_GIU: "#1f2430", HD_TRONG: "#c0392b", HD_THAY: "#2557c7"}[q[0] if q[0] in self.NHAN_HD else HD_GIU]
            o = [d["hien"], f"{d['so_o']:,}".replace(",", "."), " | ".join(d["cac_dang"][:4]),
                 ", ".join(d["cot"][:4]) + (f" … (+{len(d['cot']) - 4})" if len(d["cot"]) > 4 else "")]
            for j, v in enumerate(o):
                tk.Label(self.frm_rong, text=v, background=nen, anchor="w", padx=6, pady=3,
                         foreground=mau if j == 0 else "#1f2430", wraplength=420, justify="left").grid(
                    row=i, column=j, sticky="we")
            cb = ttk.Combobox(self.frm_rong, textvariable=v_hd, values=list(self.NHAN_HD.values()), state="readonly",
                              width=26)
            cb.grid(row=i, column=4, sticky="w", padx=6, pady=2)
            ttk.Entry(self.frm_rong, textvariable=v_thay, width=22).grid(row=i, column=5, sticky="w", padx=6)
            cb.bind("<<ComboboxSelected>>", lambda e, kk=k: self._rong_doc_1(kk, nhan_sang_hd))
        if not ds:
            tk.Label(self.frm_rong, text="(Không thấy giá trị Không / Đang cập nhật nào trong dữ liệu)",
                     background="#ffffff", foreground="#777", pady=12).grid(row=1, column=0, columnspan=6, sticky="w")
    def _rong_doc_1(self, k: str, nhan_sang_hd: dict) -> None:
        v_hd, v_thay = self._bien_rong[k]
        self._rong[k] = [nhan_sang_hd.get(v_hd.get(), HD_GIU), chuan_hoa_key(v_thay.get())]
    def _rong_doc_het(self) -> None:
        nhan_sang_hd = {v: k for k, v in self.NHAN_HD.items()}
        for k in list(self._bien_rong):
            self._rong_doc_1(k, nhan_sang_hd)
        self._rong = {k: v for k, v in self._rong.items() if v[0] != HD_GIU}
    def _rong_dat_tat_ca(self, hd: str) -> None:
        for k, (v_hd, _) in self._bien_rong.items():
            v_hd.set(self.NHAN_HD[hd])
    def _rong_them(self) -> None:
        v = self.var_rong_moi.get().strip()
        if not v:
            return
        k = khoa_gia_tri_rong(v)
        self._rong_doc_het()
        self._rong.setdefault(k, [HD_TRONG, ""])
        self.var_rong_moi.set("")
        self._tinh_lai()
        self.nb.select(4)
    def _rong_ap(self) -> None:
        if not self._co_the_sua():
            return
        self._rong_doc_het()
        thieu = [k for k, v in self._rong.items() if v[0] == HD_THAY and not v[1]]
        if thieu:
            messagebox.showwarning("Thiếu giá trị thay", "Chưa nhập 'Thay bằng' cho: " + ", ".join(thieu))
            return
        self._tinh_lai()
        st = self.kq["stat"]
        self.var_tt.set(f"✔ Đã áp: {st.get('rong_trong', 0)} ô để trống (không cập nhật), {st.get('rong_thay', 0)} ô "
                        "thay thế. Đã tự lưu — xuất ở bước ③.")
    def _xoa_het_sua(self) -> None:
        if not self._sua:
            messagebox.showinfo("Không có", "Chưa có ô sửa tay nào.")
            return
        if messagebox.askyesno("Xoá hết sửa tay", f"Xoá toàn bộ {len(self._sua)} ô sửa tay đã lưu?"):
            self._sua.clear()
            self._tinh_lai()
    # ================================================================ đơn vị hàng loạt
    def _ve_don_vi(self) -> None:
        for w in self.frm_dv.winfo_children():
            w.destroy()
        if not self.kq:
            return
        ds = self.kq.get("don_vi_cot", [])
        if not self.var_dv_tat_ca.get():
            ds = [c for c in ds if c["la_kt"] or c["da_luu"] or self._dv_ap.get((c["cate"], c["code"]))]
        ds = sorted(ds, key=lambda c: (c["cate"], not c["la_kt"], -c["so_tron"], c["code"]))
        tieu_de = ["NH", "Mã TSKT", "Tên", "Số trơn / có dữ liệu", "Giá trị hiện tại (sau áp)", "Gợi ý", "Đã lưu",
                   "ĐƠN VỊ"]
        for j, t in enumerate(tieu_de):
            tk.Label(self.frm_dv, text=t, font=("Segoe UI", 10, "bold"), background="#2f6fed", foreground="white",
                     padx=6, pady=4, anchor="w").grid(row=0, column=j, sticky="we")
        self._dv_hien: List[Tuple[str, str]] = []
        for i, c in enumerate(ds, start=1):
            key = (c["cate"], c["code"])
            if key not in self._bien_dv:
                self._bien_dv[key] = tk.StringVar()
            self._bien_dv[key].set(self._dv_ap.get(key, c["da_luu"]))
            self._dv_hien.append(key)
            nen = "#f3f5fb" if i % 2 else "#ffffff"
            o = [c["cate"], c["code"], c["ten"], f'{c["so_tron"]} / {c["tong"]}', c["vi_du"], c["goi_y"] or "-",
                 c["da_luu"] or "-"]
            for j, v in enumerate(o):
                tk.Label(self.frm_dv, text=v, background=nen, anchor="w", padx=6, pady=3,
                         foreground="#c0392b" if (j == 3 and c["so_tron"] and c["la_kt"]) else "#1f2430").grid(
                    row=i, column=j, sticky="we")
            ttk.Combobox(self.frm_dv, textvariable=self._bien_dv[key], values=DON_VI_LUA_CHON, width=8).grid(
                row=i, column=7, sticky="w", padx=6, pady=2)
        if not ds:
            tk.Label(self.frm_dv, text="(Không có cột kích thước/khối lượng — tick 'Hiện tất cả cột TSKT')",
                     background="#ffffff", foreground="#777", pady=12).grid(row=1, column=0, columnspan=8, sticky="w")
    def _dv_goi_y(self) -> None:
        if not self.kq:
            return
        theo = {(c["cate"], c["code"]): c for c in self.kq.get("don_vi_cot", [])}
        for key in getattr(self, "_dv_hien", []):
            c = theo.get(key)
            if c and c["so_tron"] and not self._bien_dv[key].get() and c["goi_y"]:
                self._bien_dv[key].set(c["goi_y"])
    def _dv_dat_tat_ca(self, dv: str) -> None:
        for key in getattr(self, "_dv_hien", []):
            self._bien_dv[key].set(dv)
    def _dv_ap_ngay(self) -> None:
        if not self._co_the_sua():
            return
        for key in getattr(self, "_dv_hien", []):
            self._dv_ap[key] = chuan_hoa_key(self._bien_dv[key].get())
        self._tinh_lai()
        st = self.kq["stat"]
        self.var_tt.set(f"✔ Đã áp đơn vị: {st.get('them_dv', 0)} ô được thêm đơn vị — còn {st.get('chua_don_vi', 0)} "
                        "ô số trơn. Chưa ghi file — xuất ở bước ③.")
    # ================================================================ hiển thị
    def _hien_ket_qua(self, dau: bool = False) -> None:
        kq = self.kq
        st = kq["stat"]
        for k, v in self._the.items():
            if not kq["co_doi_chieu"] and k in (TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG, TRANG_THAI_PIM_TRONG,
                                                 TRANG_THAI_DON_VI, "khong_co_pim"):
                v.set("–")
            else:
                v.set(f"{st.get(k, 0):,}".replace(",", "."))
        self._nhan_sang_khoa = {ct["nhan"]: k for k, ct in kq["chi_tiet"].items()}
        self.cbo_loai_cb["values"] = [self.LOC_TAT_CA] + sorted({r[1] for r in kq["canh_bao"]})
        self.cbo_ma["values"] = [self.LOC_TAT_CA] + sorted({r[3] for r in kq["khac"]})
        if dau:
            self.var_loai_cb.set(self.LOC_TAT_CA)
            self.var_ma.set(self.LOC_TAT_CA)
            self.var_loai.set(self.LOC_CAN_XEM)
            self.var_tim.set("")
        self._ve_cb()
        self._ve_khac()
        self._ve_ds_sku()
        if self._ct_khoa and self._ct_khoa in kq["chi_tiet"]:
            self._ve_chi_tiet(self._ct_khoa)
        self._ve_don_vi()
        self._ve_rong()
        self._cap_nhat_cho()
        cao = sum(1 for r in kq["canh_bao"] if r[0] == "CAO")
        self.var_tt.set(
            f"Giá trị mới: {kq['nhan_moi']}  |  Spec PIM cũ: {kq['nhan_cu']}  |  "
            + (f"⚠️ {cao} cảnh báo mức CAO" if cao else "✔ Không có cảnh báo mức CAO")
            + "  —  bấm ô số để lọc."
        )
        if dau:
            self.nb.select(1 if (st.get(TRANG_THAI_KHAC) or st.get(TRANG_THAI_TOOL_TRONG)) else 0)
    def _nap_tree(self, tree: ttk.Treeview, rows: List[list], tag_idx: int, sua_idx: Optional[int] = None) -> int:
        tree.delete(*tree.get_children())
        for i, r in enumerate(rows[: self.MAX_DONG_HIEN]):
            tags = [str(r[tag_idx])]
            if sua_idx is not None and r[sua_idx]:
                tags.append("sua")
            tree.insert("", "end", iid=str(i), values=r, tags=tuple(tags))
        return len(rows)
    def _ve_cb(self) -> None:
        if not self.kq:
            return
        loai = self.var_loai_cb.get()
        rows = [r for r in self.kq["canh_bao"] if loai == self.LOC_TAT_CA or r[1] == loai]
        n = self._nap_tree(self.tree_cb, rows, 0)
        self.var_dem_cb.set(f"{n} dòng" + (f" (hiện {self.MAX_DONG_HIEN} dòng đầu)" if n > self.MAX_DONG_HIEN else ""))
    def _ve_khac(self) -> None:
        if not self.kq:
            return
        loai, ma, tim = self.var_loai.get(), self.var_ma.get(), self.var_tim.get().strip().lower()
        hien = []
        for r in self.kq["khac"]:
            tt, da_sua = r[7], r[10]
            if loai == self.LOC_CAN_XEM and tt not in (TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG) and not da_sua:
                continue
            if loai == self.LOC_DA_SUA and not da_sua:
                continue
            if loai not in (self.LOC_CAN_XEM, self.LOC_TAT_CA, self.LOC_DA_SUA) and tt != loai:
                continue
            if ma != self.LOC_TAT_CA and r[3] != ma:
                continue
            if tim and not any(tim in str(x).lower() for x in (r[0], r[1], r[5], r[6])):
                continue
            hien.append(r)
        self._khac_hien = hien[: self.MAX_DONG_HIEN]
        rows = [[r[0], r[2], r[3], r[4], r[5], r[6], self._giai_nghia(r[3], r[6]), r[7], "✎" if r[10] else ""]
                for r in hien]
        n = self._nap_tree(self.tree_khac, rows, 7, 8)
        self.var_dem_khac.set(f"{n} ô" + (f" (hiện {self.MAX_DONG_HIEN} ô đầu — lọc thêm)" if n > self.MAX_DONG_HIEN
                                          else ""))
    def _ve_ds_sku(self) -> None:
        if not self.kq:
            return
        tim = self.var_tim_sku.get().strip().lower()
        chi_khac = self.var_chi_khac.get()
        self.lst.delete(0, "end")
        self._ds_khoa_hien = []
        for k, ct in self.kq["chi_tiet"].items():
            co_khac = any(d[4] in (TRANG_THAI_KHAC, TRANG_THAI_TOOL_TRONG) for d in ct["dong"])
            co_sua = any(d[6] for d in ct["dong"])
            if chi_khac and not (co_khac or co_sua):
                continue
            if tim and tim not in ct["nhan"].lower() and tim not in ct["model"].lower():
                continue
            self._ds_khoa_hien.append(k)
            self.lst.insert("end", ("● " if co_khac else ("✎ " if co_sua else "   ")) + ct["nhan"])
            if co_khac:
                self.lst.itemconfigure("end", foreground="#c0392b")
            elif co_sua:
                self.lst.itemconfigure("end", foreground="#6b46c1")
    def _ve_chi_tiet(self, khoa: Optional[str] = None) -> None:
        if not self.kq:
            return
        if khoa is None:
            sel = self.lst.curselection()
            if not sel:
                return
            khoa = self._ds_khoa_hien[sel[0]]
        ct = self.kq["chi_tiet"].get(khoa)
        if not ct:
            return
        self._ct_khoa = khoa
        self.var_tieu_de_sku.set(
            f"{ct['nhan']}  |  model: {ct['model'] or '(trống)'}  |  biến thể: {ct['variant'] or '-'}  |  "
            f"tab: {ct['tab']}" + ("" if ct["co_pim"] else "  |  ⚠️ KHÔNG có trong file PIM")
            + "   —  double-click TOOL MỚI để sửa, PIM CŨ để lấy giá trị cũ"
        )
        thu_tu = {TRANG_THAI_KHAC: 0, TRANG_THAI_TOOL_TRONG: 1, TRANG_THAI_DON_VI: 2, TRANG_THAI_PIM_TRONG: 3}
        self._ct_hien = sorted(ct["dong"], key=lambda d: (not d[6], thu_tu.get(d[4], 9), d[0]))
        rows = [[d[0], d[1], d[2], d[3], self._giai_nghia(d[0], d[3]), d[4], "✎" if d[6] else ""]
                for d in self._ct_hien]
        self._nap_tree(self.tree_ct, rows, 5, 6)
    def _mo_sku(self, khoa: str) -> None:
        if not self.kq or khoa not in self.kq["chi_tiet"]:
            return
        nhan = self.kq["chi_tiet"][khoa]["nhan"]
        self.var_chi_khac.set(False)
        self.var_tim_sku.set(nhan)
        self._ve_ds_sku()
        if self._ds_khoa_hien:
            self.lst.selection_clear(0, "end")
            self.lst.selection_set(0)
        self._ve_chi_tiet(khoa)
        self.nb.select(2)
    def _nhay_toi(self, dich) -> None:
        if not self.kq:
            return
        kieu, gt = dich
        if kieu == "khac":
            self.var_loai.set(gt)
            self.var_ma.set(self.LOC_TAT_CA)
            self._ve_khac()
            self.nb.select(1)
        elif kieu == "dv":
            self.nb.select(3)
        elif kieu == "rong":
            self.nb.select(4)
        else:
            self.var_loai_cb.set(gt)
            self._ve_cb()
            self.nb.select(0)
    def _luu_excel(self) -> None:
        if not self.kq:
            messagebox.showinfo("Chưa có kết quả", "Bấm ② Kiểm tra trước.")
            return
        f = filedialog.asksaveasfilename(
            title="Lưu kết quả kiểm tra", defaultextension=".xlsx",
            initialfile=f"KIEM_TRA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not f:
            return
        try:
            xuat_ket_qua_kiem_tra(self.kq, f)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi khi lưu", str(e))
            return
        self.var_tt.set(f"✔ Đã lưu kết quả kiểm tra: {f}")
class PimApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TGDĐ CMS -> PIM — Chuyển đổi TSKT + FILTER")
        self.geometry("1080x720")
        self.minsize(760, 520)
        self._ap_dung_giao_dien()
        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._done_queue: "queue.Queue[tuple]" = queue.Queue()
        self._worker: threading.Thread | None = None
        self._out_dir: Path | None = None
        self._da_xong_thanh_cong = False
        # BỔ SUNG: danh sách MỌI nút thao tác đọc/ghi workspace (Nạp, Xoá,
        # Chạy tất cả...) — khoá HẾT các nút này trong lúc có 1 thao tác
        # đang chạy nền, để không có 2 luồng cùng mở/ghi file .xlsx song
        # song (dễ ghi đè lẫn nhau / hỏng file).
        self._nut_can_khoa_khi_chay: list[tk.Widget] = []
        self._build_ui()
        self.after(100, self._bom_log)
        try:
            ws_file = dam_bao_workspace()
            self.var_file.set(str(ws_file))
            self.var_out.set(str(ws_file.parent / "output"))
            self.var_status.set(f"Đã sẵn sàng workspace: {ws_file}")
        except Exception as e:  # noqa: BLE001
            messagebox.showwarning(
                "Không tự tạo được workspace",
                f"Lỗi: {e}\n\nAnh tự chọn file Excel nguồn bằng nút '📂 Chọn file...' bên dưới.",
            )
    def _ap_dung_giao_dien(self) -> None:
        MAU_NEN = "#eef1f8"
        MAU_KHUNG = "#ffffff"
        MAU_CHINH = "#2f6fed"
        MAU_CHINH_DAM = "#2557c7"
        MAU_THANH_CONG = "#1f9d55"
        MAU_THANH_CONG_DAM = "#178a49"
        MAU_CHU = "#1f2430"
        MAU_CHU_MO = "#5b6472"
        MAU_NGUY_HIEM = "#c0392b"
        MAU_NGUY_HIEM_DAM = "#a5301f"
        self.configure(background=MAU_NEN)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=MAU_NEN, foreground=MAU_CHU, font=("Segoe UI", 10))
        style.configure("TFrame", background=MAU_NEN)
        style.configure("TLabel", background=MAU_NEN, foreground=MAU_CHU)
        style.configure("Section.TLabel", background=MAU_NEN, foreground=MAU_CHU_MO, font=("Segoe UI", 10, "bold"))
        style.configure("TButton", padding=(10, 6), font=("Segoe UI", 10), background="#e3e8f4", foreground=MAU_CHU)
        style.map("TButton", background=[("active", "#d5dcf0")])
        style.configure(
            "Accent.TButton", padding=(14, 8), font=("Segoe UI", 10, "bold"),
            background=MAU_THANH_CONG, foreground="white",
        )
        style.map("Accent.TButton", background=[("active", MAU_THANH_CONG_DAM), ("disabled", "#9fd4b5")])
        style.configure(
            "Secondary.TButton", padding=(10, 6), font=("Segoe UI", 10),
            background=MAU_CHINH, foreground="white",
        )
        style.map("Secondary.TButton", background=[("active", MAU_CHINH_DAM), ("disabled", "#a9c2f5")])
        style.configure(
            "Canh.TButton", padding=(10, 6), font=("Segoe UI", 10, "bold"),
            background=MAU_NGUY_HIEM, foreground="white",
        )
        style.map("Canh.TButton", background=[("active", MAU_NGUY_HIEM_DAM), ("disabled", "#e3a89f")])
        style.configure("TNotebook", background=MAU_NEN, borderwidth=0)
        style.configure(
            "TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 10, "bold"),
            background="#dbe1f0", foreground=MAU_CHU_MO,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", MAU_KHUNG)],
            foreground=[("selected", MAU_CHINH_DAM)],
        )
        style.configure(
            "Treeview", background=MAU_KHUNG, fieldbackground=MAU_KHUNG,
            foreground=MAU_CHU, rowheight=24, font=("Segoe UI", 10), borderwidth=0,
        )
        style.configure(
            "Treeview.Heading", background=MAU_CHINH, foreground="white",
            font=("Segoe UI", 10, "bold"), relief="flat",
        )
        style.map("Treeview.Heading", background=[("active", MAU_CHINH_DAM)])
        style.map("Treeview", background=[("selected", "#bcd0fb")], foreground=[("selected", MAU_CHU)])
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("Horizontal.TProgressbar", background=MAU_CHINH, troughcolor="#dbe1f0")
    # ------------------------------------------------------------------
    # BỔ SUNG: chạy 1 tác vụ đọc/ghi workspace ở THREAD NỀN + khoá nút +
    # thanh tiến trình, dùng chung cho MỌI nút Nạp/Xoá dữ liệu — trước
    # đây các nút này gọi thẳng hàm nặng ngay trên main thread nên khi
    # file lớn, giao diện ĐỨNG HÌNH không phản hồi trong lúc chờ (khiến
    # cảm giác "rất lâu" còn tệ hơn vì tưởng app bị treo, không phải chỉ
    # đang chạy chậm), lại không có nút Hủy/không thấy tiến độ gì.
    # ------------------------------------------------------------------
    def _dat_trang_thai_nut(self, state: str) -> None:
        for w in self._nut_can_khoa_khi_chay:
            try:
                w.configure(state=state)
            except tk.TclError:
                pass
    def _chay_tac_vu_nen(self, ham, callback_xong, *, ten_thao_tac: str = "Đang xử lý") -> None:
        """`ham`: callable không tham số, chạy trong thread nền, trả về 1
        dict kết quả (dạng {"loi": ...} như các hàm engine ở trên).
        `callback_xong(ket_qua)`: được gọi Ở MAIN THREAD khi xong — an
        toàn để show messagebox / cập nhật widget trong đó."""
        self.progress.start(12)
        self.var_status.set(f"{ten_thao_tac}... (đừng đóng cửa sổ)")
        self._dat_trang_thai_nut("disabled")
        q: "queue.Queue" = queue.Queue()
        def _nen():
            try:
                kq = ham()
            except Exception as e:  # noqa: BLE001
                kq = {"loi": f"Lỗi ngoài dự kiến: {e}\n\n{traceback.format_exc(limit=3)}"}
            q.put(kq)
        threading.Thread(target=_nen, daemon=True).start()
        def _cho():
            try:
                kq = q.get_nowait()
            except queue.Empty:
                self.after(120, _cho)
                return
            self.progress.stop()
            self._dat_trang_thai_nut("normal")
            callback_xong(kq)
        self.after(120, _cho)
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        tab_run = ttk.Frame(self.notebook)
        self.notebook.add(tab_run, text="🚀 Chạy pipeline")
        frm_top = ttk.Frame(tab_run)
        frm_top.pack(fill="x", **pad)
        ttk.Label(frm_top, text="File Excel nguồn (.xlsx) — workspace tự quản lý:").grid(
            row=0, column=0, sticky="w"
        )
        self.var_file = tk.StringVar()
        ent_file = ttk.Entry(frm_top, textvariable=self.var_file, width=70)
        ent_file.grid(row=1, column=0, sticky="we", pady=(2, 0))
        ttk.Button(frm_top, text="📂 Chọn file khác...", command=self._chon_file).grid(row=1, column=1, padx=(6, 0))
        ttk.Label(frm_top, text="Thư mục xuất kết quả (.zip):").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.var_out = tk.StringVar()
        ent_out = ttk.Entry(frm_top, textvariable=self.var_out, width=70)
        ent_out.grid(row=3, column=0, sticky="we", pady=(2, 0))
        ttk.Button(frm_top, text="📁 Chọn thư mục...", command=self._chon_out_dir).grid(row=3, column=1, padx=(6, 0))
        frm_top.columnconfigure(0, weight=1)
        frm_btn = ttk.Frame(tab_run)
        frm_btn.pack(fill="x", **pad)
        self.btn_run = ttk.Button(frm_btn, text="🚀 CHẠY TẤT CẢ", style="Accent.TButton", command=self._bam_chay)
        self.btn_run.pack(side="left")
        self._nut_can_khoa_khi_chay.append(self.btn_run)
        self.btn_open_out = ttk.Button(
            frm_btn, text="📂 Mở thư mục kết quả", style="Secondary.TButton",
            command=self._mo_thu_muc_ket_qua, state="disabled",
        )
        self.btn_open_out.pack(side="left", padx=(8, 0))
        self.progress = ttk.Progressbar(frm_btn, mode="indeterminate", length=220)
        self.progress.pack(side="left", padx=(16, 0))
        frm_btn2 = ttk.Frame(tab_run)
        frm_btn2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(frm_btn2, text="Nạp dữ liệu:", style="Section.TLabel").pack(side="left")
        btn = ttk.Button(
            frm_btn2, text="📥 Nạp file dữ liệu (1 hoặc nhiều file, tự lọc luôn)...", style="Accent.TButton",
            command=self._nap_va_loc_mot_buoc,
        )
        btn.pack(side="left", padx=(8, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(frm_btn2, text="📋 Danh sách SKU (IMPORT)...", command=self._mo_dialog_nhap_import)
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn2, text="📦 Nạp file export PIM (IMPORT + spec đối chiếu)...", style="Secondary.TButton",
            command=self._mo_dialog_nap_pim,
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn2, text="⚙ Trích thủ công từ sheet có sẵn trong workspace...",
            command=self._mo_dialog_trich_data_sp,
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        frm_btn2b = ttk.Frame(tab_run)
        frm_btn2b.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(frm_btn2b, text="Xóa dữ liệu cũ (giữ header):", style="Section.TLabel").pack(side="left")
        # BỔ SUNG THEO YÊU CẦU: 2 nút xóa riêng — IMPORT và DATA SP — để
        # dọn sạch dữ liệu của lần nạp trước TRƯỚC KHI nạp lô mới, không
        # cần nhớ chọn đúng "Ghi đè" mỗi lần. Đặt NGAY CẠNH khu vực Nạp dữ
        # liệu (nơi hay hiện lỗi/xác nhận) để dễ thấy khi cần dùng.
        btn = ttk.Button(
            frm_btn2b, text="🗑️ Xóa dữ liệu IMPORT", style="Canh.TButton",
            command=lambda: self._xoa_du_lieu_cu("IMPORT", CFG.sheet_nhap, CFG.data_start_row),
        )
        btn.pack(side="left", padx=(8, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn2b, text="🗑️ Xóa dữ liệu DATA SP", style="Canh.TButton",
            command=lambda: self._xoa_du_lieu_cu("DATA SP", CFG.sheet_data_sp[0], CFG.sp_start_row),
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        frm_btn2c = ttk.Frame(tab_run)
        frm_btn2c.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(frm_btn2c, text="Tiện ích:", style="Section.TLabel").pack(side="left")
        btn = ttk.Button(
            frm_btn2c, text="📏 Đơn vị kích thước / khối lượng (cm, mm, kg, g, inch)...",
            command=self._mo_dialog_don_vi,
        )
        btn.pack(side="left", padx=(8, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        frm_btn3 = ttk.Frame(tab_run)
        frm_btn3.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(frm_btn3, text="Nhập mapping/cấu hình từ file khác:  ", foreground="#555").pack(side="left")
        btn = ttk.Button(
            frm_btn3, text="🧬 Mapping TSKT (đè mới)...",
            command=lambda: self._nhap_sheet_mapping(
                ten_hien_thi="MAPPING TSKT MOI", ung_vien=[CFG_TF_sheet_tskt_mapping],
                data_start_row=CFG.map_start_row,
                che_do_co_dinh="overwrite",
            ),
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn3, text="🧮 Mapping FILTER (đè mới)...",
            command=lambda: self._nhap_sheet_mapping(
                ten_hien_thi="MAPPING FILTER MOI", ung_vien=[CFG_TF_sheet_filter_mapping],
                data_start_row=CFG.map_start_row,
                che_do_co_dinh="overwrite",
            ),
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn3, text="🗄️ DATA PIM (đè mới)...",
            command=lambda: self._nhap_sheet_mapping(
                ten_hien_thi="DATA PIM", ung_vien=CFG.sheet_data_pim,
                data_start_row=CFG.pim_start_row,
                che_do_co_dinh="overwrite", ep_van_ban=False,
            ),
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        btn = ttk.Button(
            frm_btn3, text="🏷️ Cấu hình Category theo NH (nối tiếp)...",
            command=self._nhap_cau_hinh_category,
        )
        btn.pack(side="left", padx=(0, 6))
        self._nut_can_khoa_khi_chay.append(btn)
        self.var_status = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(tab_run, textvariable=self.var_status, foreground="#555").pack(anchor="w", padx=10)
        frm_log = ttk.Frame(tab_run)
        frm_log.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        ttk.Label(frm_log, text="Nhật ký chạy:").pack(anchor="w")
        txt_frame = ttk.Frame(frm_log)
        txt_frame.pack(fill="both", expand=True)
        self.txt_log = tk.Text(
            txt_frame, wrap="word", state="disabled", font=("Consolas", 10),
            background="#1e2430", foreground="#d7deec", insertbackground="#d7deec",
            relief="flat", padx=8, pady=6,
        )
        scroll = ttk.Scrollbar(txt_frame, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tab_kiem_tra = KiemTraFrame(self.notebook, self)
        self.notebook.add(self.tab_kiem_tra, text="🔍 Kiểm tra & Đối chiếu")
        tab_data = SheetEditorFrame(self.notebook, get_file_path=lambda: self.var_file.get())
        self.notebook.add(tab_data, text="📋 Quản lý dữ liệu")
    # ------------------------------------------------------------------
    # BỔ SUNG: 2 nút "Xóa dữ liệu cũ" (IMPORT / DATA SP)
    # ------------------------------------------------------------------
    def _xoa_du_lieu_cu(self, ten_hien_thi: str, ten_sheet: str, data_start_row: int) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        so_dong_hien_tai = dem_dong_co_du_lieu_sheet(p, ten_sheet, data_start_row)
        so_dong_txt = str(so_dong_hien_tai) if so_dong_hien_tai >= 0 else "?"
        if so_dong_hien_tai == 0:
            messagebox.showinfo("Không có gì để xóa", f'Sheet "{ten_sheet}" hiện đang trống.')
            return
        if not messagebox.askyesno(
            f"Xác nhận xóa dữ liệu {ten_hien_thi}",
            f'Sheet "{ten_sheet}" hiện có {so_dong_txt} dòng dữ liệu.\n\n'
            f"⚠️ Sẽ XÓA SẠCH toàn bộ {so_dong_txt} dòng này (giữ nguyên header) — thường dùng để dọn dữ liệu "
            "của lần nạp trước trước khi nạp lô mới.\n\nKHÔNG THỂ HOÀN TÁC. Tiếp tục?",
            icon="warning",
        ):
            return
        def _lam():
            return xoa_du_lieu_sheet(p, ten_sheet, data_start_row)
        def _xong(kq: dict):
            if kq.get("loi"):
                messagebox.showerror(f"Lỗi khi xóa dữ liệu {ten_hien_thi}", kq["loi"])
                self.var_status.set(f"Có lỗi khi xóa dữ liệu {ten_hien_thi}.")
                return
            messagebox.showinfo(
                "Đã xóa xong", f'Đã xóa {kq["so_dong_xoa"]} dòng dữ liệu cũ trong sheet "{ten_sheet}".'
            )
            self.var_status.set(f'✔ Đã xóa {kq["so_dong_xoa"]} dòng dữ liệu cũ trong "{ten_sheet}".')
        self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac=f"Đang xóa dữ liệu {ten_hien_thi}")
    def _nap_va_loc_mot_buoc(self) -> None:
        # NÂNG CẤP: chọn 1 HOẶC NHIỀU file CMS export cùng lúc (mỗi file
        # thường là 1 ngành hàng) — trước đây phải lặp lại "Nạp & Lọc" cho
        # từng file một, giờ chọn hết 1 lần (giữ Ctrl/Shift khi chọn trong
        # hộp thoại), tool tự xử lý tuần tự và gộp kết quả vào DATA SP.
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        src_paths = filedialog.askopenfilenames(
            title="Chọn 1 hoặc nhiều file Excel dữ liệu (vd export từ CMS) — giữ Ctrl/Shift để chọn nhiều",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not src_paths:
            return
        src_paths = list(src_paths)
        top = tk.Toplevel(self)
        top.title("Nạp dữ liệu vào DATA SP")
        if len(src_paths) == 1:
            nguon_txt = f"Nguồn: {Path(src_paths[0]).name}"
        else:
            ten_hien = "\n".join(f"  • {Path(sp).name}" for sp in src_paths[:8])
            them = f"\n  ... và {len(src_paths) - 8} file khác" if len(src_paths) > 8 else ""
            nguon_txt = f"Nguồn: {len(src_paths)} file đã chọn\n{ten_hien}{them}"
        ttk.Label(
            top,
            text=f"{nguon_txt}\n"
            "Tool sẽ tự nhận diện đúng cột (PRODUCTID/PROPERTYID/PROPVALUE/...) và lọc thẳng vào DATA SP. "
            "Nếu chọn nhiều file, tool xử lý lần lượt từng file — 1 file lỗi không làm dừng các file còn lại.",
            wraplength=440, justify="left",
        ).pack(anchor="w", padx=14, pady=(14, 8))
        var_che_do = tk.StringVar(value="append")
        ttk.Label(top, text="Cách ghi vào DATA SP:").pack(anchor="w", padx=14)
        ttk.Radiobutton(
            top, text="Nối tiếp (giữ dữ liệu cũ, thêm dữ liệu mới vào cuối) — khuyến nghị",
            variable=var_che_do, value="append",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        ttk.Radiobutton(
            top,
            text="Ghi đè (xoá hết dữ liệu cũ, rồi ghi dữ liệu mới — mọi file đã chọn nối tiếp nhau, "
            "chỉ xoá dữ liệu CŨ đúng 1 lần)",
            variable=var_che_do, value="overwrite",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        def _chay():
            che_do_chon = var_che_do.get()  # đọc tk.StringVar Ở MAIN THREAD trước khi
            # đưa vào thread nền — KHÔNG được gọi .get() của biến tkinter bên trong
            # _lam() (chạy trong thread nền), Tk không an toàn khi truy cập từ thread
            # khác thread chính, sẽ ném "main thread is not in main loop".
            top.destroy()
            def _lam():
                return nap_va_loc_nhieu_file(p, src_paths, che_do_chon)
            def _xong(kq: dict):
                if kq.get("loi"):
                    messagebox.showerror("Lỗi khi nạp dữ liệu", kq["loi"])
                    self.var_status.set("Có lỗi khi nạp dữ liệu — xem chi tiết ở hộp thoại.")
                    return
                dong_tong = (
                    f'Đã lọc tổng cộng {kq["so_dong"]} dòng vào DATA SP (từ dòng {kq["dong_bat_dau"]}).'
                )
                if len(src_paths) > 1 or kq.get("so_file_loi", 0) > 0:
                    chi_tiet_dong = []
                    for ten, so_dong, loi_rieng in kq.get("chi_tiet", []):
                        if loi_rieng:
                            chi_tiet_dong.append(f"✖ {ten}: {loi_rieng}")
                        else:
                            chi_tiet_dong.append(f"✔ {ten}: {so_dong} dòng")
                    dong_tong += (
                        f"\n\n{kq['so_file_thanh_cong']}/{len(src_paths)} file thành công"
                        + (f", {kq['so_file_loi']} file lỗi" if kq.get("so_file_loi") else "")
                        + ":\n" + "\n".join(chi_tiet_dong)
                    )
                messagebox.showinfo("Đã nạp xong", dong_tong)
                self.var_status.set(
                    f'✔ Đã lọc {kq["so_dong"]} dòng vào DATA SP'
                    + (f' ({kq["so_file_thanh_cong"]}/{len(src_paths)} file).' if len(src_paths) > 1 else ".")
                )
            self._chay_tac_vu_nen(
                _lam, _xong,
                ten_thao_tac=f"Đang nạp và lọc {len(src_paths)} file" if len(src_paths) > 1 else "Đang nạp và lọc dữ liệu",
            )
        ttk.Button(top, text="Nạp & Lọc ngay", style="Accent.TButton", command=_chay).pack(pady=14)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    def _mo_dialog_trich_data_sp(self) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        try:
            # KHÔNG dùng read_only=True ở đây: dialog này chỉ đọc DÒNG 1
            # (header) của MỖI sheet trong workspace để tìm sheet dữ liệu
            # thô — rẻ dù workspace lớn, nên không cần tối ưu tốc độ đọc.
            # Quan trọng hơn: read_only=True có thể trả về max_column=None
            # cho sheet nào bị thiếu thẻ <dimension> hợp lệ (thường gặp ở
            # sheet dữ liệu thô dán/copy từ nguồn ngoài Excel), gây lỗi
            # "'<' not supported between instances of 'NoneType' and
            # 'int'" — mở kiểu thường (quét ô thật) luôn trả về số, không
            # bao giờ None.
            wb = load_workbook(p)
            bo_qua = {x.upper() for x in (
                CFG.sheet_nhap, CFG.sheet_chon, CFG.sheet_cau_hinh, CFG.sheet_log, CFG.sheet_detail,
                CFG_TF_sheet_tskt_mapping, CFG_TF_sheet_filter_mapping, *CFG_sheet_tien_ich,
            )}
            for x in CFG.sheet_data_sp + CFG.sheet_data_pim + CFG.sheet_mapping:
                bo_qua.add(x.upper())
            ung_vien = []
            for sh in wb.worksheets:
                ten = sh.title.upper()
                if ten in bo_qua or ten.startswith(CFG.tskth_prefix.upper()):
                    continue
                if sh.max_column < 1:
                    continue
                header = sheet_display_values(sh, 1, 1, sh.max_column)[0]
                if (tim_cot_theo_ten(header, "PROPERTYID") > 0 and tim_cot_theo_ten(header, "PROPVALUE") > 0
                        and (tim_cot_theo_ten(header, "PRODUCTCODE") > 0 or tim_cot_theo_ten(header, "PRODUCTID") > 0)):
                    ung_vien.append(sh.title)
            wb.close()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi đọc workspace", str(e))
            return
        if not ung_vien:
            messagebox.showinfo(
                "Không có sheet dữ liệu thô",
                "Không tìm thấy sheet nào đủ cột PRODUCTID/PROPERTYID/PROPVALUE trong workspace.\n\n"
                "Dùng nút '📥 Nhập dữ liệu thô từ file khác...' để nạp vào trước.",
            )
            return
        top = tk.Toplevel(self)
        top.title("Trích/Cập nhật DATA SP")
        ttk.Label(top, text="Chọn sheet dữ liệu thô nguồn:").pack(anchor="w", padx=14, pady=(14, 4))
        var_sheet = tk.StringVar(value=ung_vien[0])
        ttk.Combobox(top, textvariable=var_sheet, values=ung_vien, state="readonly", width=40).pack(padx=14)
        var_che_do = tk.StringVar(value="append")
        ttk.Label(top, text="Cách ghi vào DATA SP:").pack(anchor="w", padx=14, pady=(14, 4))
        ttk.Radiobutton(
            top, text="Nối tiếp (giữ dữ liệu cũ, thêm vào cuối) — khuyến nghị",
            variable=var_che_do, value="append",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        ttk.Radiobutton(
            top, text="Ghi đè (xoá hết dữ liệu DATA SP cũ, ghi mới)",
            variable=var_che_do, value="overwrite",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        def _chay():
            top.destroy()
            ten_sheet_chon = var_sheet.get()
            che_do_chon = var_che_do.get()
            def _lam():
                return trich_data_sp_tuong_tac(p, ten_sheet_chon, che_do_chon)
            def _xong(kq: dict):
                if kq.get("loi"):
                    messagebox.showerror("Lỗi khi trích", kq["loi"])
                    return
                messagebox.showinfo(
                    "Đã trích xong",
                    f'Đã trích {kq["so_dong"]} dòng từ "{ten_sheet_chon}" vào DATA SP '
                    f'(từ dòng {kq["dong_bat_dau"]}).',
                )
                self.var_status.set(f'✔ Đã trích {kq["so_dong"]} dòng vào DATA SP.')
            self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac="Đang trích DATA SP")
        ttk.Button(top, text="Trích ngay", style="Accent.TButton", command=_chay).pack(pady=16)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    def _mo_dialog_nhap_import(self) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        src_path = filedialog.askopenfilename(
            title="Chọn file Excel chứa danh sách SKU (cần có cột 'sku' ở dòng 1)",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not src_path:
            return
        top = tk.Toplevel(self)
        top.title("Nhập danh sách SKU vào IMPORT")
        ttk.Label(
            top,
            text=f"Nguồn: {Path(src_path).name}\n"
            "Tự nhận diện cột theo tên ở dòng 1 (model_code/sku/variant_code/category_code).",
            wraplength=400, justify="left",
        ).pack(anchor="w", padx=14, pady=(14, 8))
        var_che_do = tk.StringVar(value="overwrite")
        ttk.Label(top, text="Cách ghi vào IMPORT:").pack(anchor="w", padx=14)
        ttk.Radiobutton(
            top, text="Ghi đè (xoá hết SKU cũ, ghi danh sách mới) — khuyến nghị",
            variable=var_che_do, value="overwrite",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        ttk.Radiobutton(
            top, text="Nối tiếp (giữ SKU cũ, thêm SKU mới vào cuối)",
            variable=var_che_do, value="append",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        def _chay():
            top.destroy()
            che_do_chon = var_che_do.get()
            def _lam():
                return nhap_import_tuong_tac(p, src_path, che_do_chon)
            def _xong(kq: dict):
                if kq.get("loi"):
                    messagebox.showerror("Lỗi khi nhập", kq["loi"])
                    return
                messagebox.showinfo(
                    "Đã nhập xong",
                    f'Đã nhập {kq["so_dong"]} SKU vào sheet IMPORT (từ dòng {kq["dong_bat_dau"]}).',
                )
                self.var_status.set(f'✔ Đã nhập {kq["so_dong"]} SKU vào IMPORT.')
            self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac="Đang nhập danh sách SKU")
        ttk.Button(top, text="Nhập ngay", style="Accent.TButton", command=_chay).pack(pady=14)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    def _nhap_cau_hinh_category(self) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        src_path = filedialog.askopenfilename(
            title="Chọn file Excel chứa (các) ngành hàng mới cần thêm vào CẤU HÌNH CATEGORY",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not src_path:
            return
        so_dong_cu = dem_dong_co_du_lieu_sheet(p, CFG.sheet_cau_hinh, 2)
        so_dong_cu_txt = str(so_dong_cu) if so_dong_cu >= 0 else "?"
        if not messagebox.askyesno(
            "Xác nhận nhập dữ liệu",
            f'Sheet "CẤU HÌNH CATEGORY" hiện có {so_dong_cu_txt} dòng (mã + nhãn).\n\n'
            f"NỐI TIẾP: sẽ GIỮ NGUYÊN dữ liệu cũ, thêm ngành hàng mới từ file:\n"
            f"{Path(src_path).name}\n\nTiếp tục?",
        ):
            return
        def _lam():
            return nhap_cau_hinh_category_tuong_tac(p, src_path)
        def _xong(kq: dict):
            if kq.get("loi"):
                messagebox.showerror("Lỗi khi nhập", kq["loi"])
                return
            canh_bao = ""
            if kq["trung"]:
                canh_bao = (
                    "\n\n⚠️ CATEGORY ID sau đã CÓ SẴN trong CẤU HÌNH CATEGORY (vẫn được thêm mới, "
                    "tự kiểm tra lại xem có bị trùng/nhân đôi ngành hàng không):\n" + ", ".join(kq["trung"])
                )
            messagebox.showinfo(
                "Đã nhập xong",
                f'Đã thêm {kq["so_dong"]} dòng vào "CẤU HÌNH CATEGORY" (nối tiếp, từ dòng {kq["dong_bat_dau"]}).{canh_bao}',
            )
            self.var_status.set(f'✔ Đã thêm {kq["so_dong"]} dòng vào "CẤU HÌNH CATEGORY".')
        self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac="Đang nhập cấu hình Category")
    def _nhap_sheet_mapping(
        self,
        ten_hien_thi: str,
        ung_vien: List[str],
        data_start_row: int,
        che_do_co_dinh: str,
        cot_kiem_tra_trung: int | None = None,
        nguon_data_start_row: int | None = None,
        ep_van_ban: bool = True,
    ) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        try:
            wb = load_workbook(p, read_only=True)
            ten_sheet_dich = next((t for t in ung_vien if t in wb.sheetnames), None)
            wb.close()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Lỗi đọc workspace", str(e))
            return
        if not ten_sheet_dich:
            messagebox.showerror(
                "Không tìm thấy sheet",
                f'Workspace không có sheet nào trong nhóm "{ten_hien_thi}" ({", ".join(ung_vien)}).',
            )
            return
        src_path = filedialog.askopenfilename(
            title=f"Chọn file Excel nguồn để nhập vào {ten_hien_thi}",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not src_path:
            return
        so_dong_cu = dem_dong_co_du_lieu_sheet(p, ten_sheet_dich, data_start_row)
        so_dong_nguon = dem_dong_co_du_lieu_file(src_path, nguon_data_start_row or data_start_row)
        so_dong_cu_txt = str(so_dong_cu) if so_dong_cu >= 0 else "?"
        so_dong_nguon_txt = str(so_dong_nguon) if so_dong_nguon >= 0 else "?"
        if che_do_co_dinh == "overwrite":
            noi_dung = (
                f'Sheet "{ten_sheet_dich}" hiện có {so_dong_cu_txt} dòng dữ liệu.\n\n'
                f"⚠️ GHI ĐÈ: sẽ XOÁ HẾT {so_dong_cu_txt} dòng cũ này, thay bằng "
                f"{so_dong_nguon_txt} dòng mới từ file:\n{Path(src_path).name}\n\n"
                f"Sau khi xong sẽ còn lại: {so_dong_nguon_txt} dòng.\n\nTiếp tục?"
            )
        else:
            tong_sau = (
                str(so_dong_cu + so_dong_nguon) if so_dong_cu >= 0 and so_dong_nguon >= 0 else "?"
            )
            noi_dung = (
                f'Sheet "{ten_sheet_dich}" hiện có {so_dong_cu_txt} dòng dữ liệu.\n\n'
                f"NỐI TIẾP: sẽ GIỮ NGUYÊN {so_dong_cu_txt} dòng cũ, thêm "
                f"{so_dong_nguon_txt} dòng mới từ file:\n{Path(src_path).name}\n\n"
                f"Sau khi xong sẽ có: {tong_sau} dòng (không mất dữ liệu cũ).\n\nTiếp tục?"
            )
        if not messagebox.askyesno("Xác nhận nhập dữ liệu", noi_dung):
            return
        def _lam():
            return nhap_sheet_theo_vi_tri(
                p, src_path, ten_sheet_dich, data_start_row, che_do_co_dinh, cot_kiem_tra_trung,
                nguon_data_start_row, ep_van_ban,
            )
        def _xong(kq: dict):
            if kq.get("loi"):
                messagebox.showerror("Lỗi khi nhập", kq["loi"])
                return
            canh_bao = ""
            if kq["trung"]:
                canh_bao += "\n\n⚠️ Đã có sẵn (trùng) trong sheet đích, vẫn được thêm mới — tự kiểm tra lại:\n" + ", ".join(kq["trung"][:30])
                if len(kq["trung"]) > 30:
                    canh_bao += f" ... (còn {len(kq['trung']) - 30} giá trị trùng khác)"
            kieu = "Ghi đè" if che_do_co_dinh == "overwrite" else "Nối tiếp"
            messagebox.showinfo(
                "Đã nhập xong",
                f'Đã nhập {kq["so_dong"]} dòng vào sheet "{ten_sheet_dich}" ({kieu}, từ dòng {kq["dong_bat_dau"]}).{canh_bao}',
            )
            self.var_status.set(f'✔ Đã nhập {kq["so_dong"]} dòng vào "{ten_sheet_dich}" ({kieu}).')
        self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac=f"Đang nhập {ten_hien_thi}")
    # ------------------------------------------------------------------
    # BỔ SUNG: hộp cảnh báo nổi bật sau khi xuất (đỏ, không lẫn với info)
    # ------------------------------------------------------------------
    def _hien_canh_bao_xuat(self, tom_tat: str) -> None:
        top = tk.Toplevel(self)
        top.title("⚠️ CẢNH BÁO TRƯỚC KHI IMPORT PIM")
        top.configure(background="#fff4f2")
        tk.Label(
            top, text="⚠️  ĐÃ XUẤT FILE — NHƯNG CÓ CẢNH BÁO, KIỂM TRA TRƯỚC KHI IMPORT LÊN PIM",
            background="#c0392b", foreground="white", font=("Segoe UI", 11, "bold"), padx=12, pady=10,
        ).pack(fill="x")
        txt = tk.Text(top, wrap="word", height=16, width=100, font=("Segoe UI", 10),
                      background="#fff4f2", relief="flat", padx=12, pady=10)
        txt.insert("1.0", tom_tat)
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)
        frm = ttk.Frame(top)
        frm.pack(fill="x", padx=12, pady=10)
        ttk.Button(frm, text="📂 Mở thư mục kết quả", style="Secondary.TButton",
                   command=self._mo_thu_muc_ket_qua).pack(side="left")
        def _xem_tab():
            top.destroy()
            self.notebook.select(self.tab_kiem_tra)
            self.tab_kiem_tra.var_moi.set("workspace")
            self.tab_kiem_tra._buoc_kiem_tra()
        ttk.Button(frm, text="🔍 Xem chi tiết trong tab Kiểm tra", style="Secondary.TButton",
                   command=_xem_tab).pack(side="left", padx=(8, 0))
        ttk.Button(frm, text="Đã hiểu", style="Canh.TButton", command=top.destroy).pack(side="right")
        _can_giua_popup(self, top)
        top.resizable(True, True)
        top.transient(self)
        top.grab_set()
    # ------------------------------------------------------------------
    # BỔ SUNG: nạp file export PIM -> IMPORT + SPEC PIM TẠM
    # ------------------------------------------------------------------
    def _mo_dialog_nap_pim(self) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        src_path = filedialog.askopenfilename(
            title="Chọn file export PIM (dòng 1 = mã cột, dòng 2 = tên cột, dữ liệu từ dòng 3)",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not src_path:
            return
        top = tk.Toplevel(self)
        top.title("Nạp file export PIM")
        ttk.Label(
            top,
            text=f"Nguồn: {Path(src_path).name}\n\n"
            "• model_code / sku / variant_code / category_code  →  sheet IMPORT\n"
            f"• Các cột spec còn lại  →  lưu tạm ở sheet \"{CFG_sheet_spec_pim}\".\n"
            f"  Khi bấm CHẠY TẤT CẢ, tool so giá trị mới với spec này (sheet \"{CFG_sheet_doi_chieu}\") "
            "và cảnh báo trước khi xuất.",
            wraplength=520, justify="left",
        ).pack(anchor="w", padx=14, pady=(14, 8))
        var_che_do = tk.StringVar(value="overwrite")
        ttk.Label(top, text="Cách ghi vào IMPORT + SPEC PIM TẠM:").pack(anchor="w", padx=14)
        ttk.Radiobutton(
            top, text="Ghi đè (xoá SKU + spec cũ, thay bằng file này) — khuyến nghị",
            variable=var_che_do, value="overwrite",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        ttk.Radiobutton(
            top, text="Nối tiếp (giữ SKU cũ; SKU nạp lại thì spec cũ của SKU đó được thay mới)",
            variable=var_che_do, value="append",
        ).pack(anchor="w", padx=20, pady=(4, 0))
        def _chay():
            che_do_chon = var_che_do.get()
            top.destroy()
            def _lam():
                return nap_file_export_pim(p, src_path, che_do_chon)
            def _xong(kq: dict):
                if kq.get("loi"):
                    messagebox.showerror("Lỗi khi nạp file PIM", kq["loi"])
                    self.var_status.set("Có lỗi khi nạp file export PIM.")
                    return
                dong = [
                    f'✔ {kq["so_sku"]} SKU → IMPORT (từ dòng {kq["dong_bat_dau"]}).',
                    f'✔ {kq["so_o_spec"]} ô spec ({kq["so_cot_spec"]} cột) → "{CFG_sheet_spec_pim}".',
                ]
                if kq.get("so_dong_spec_giu_lai"):
                    dong.append(f'  (giữ lại {kq["so_dong_spec_giu_lai"]} dòng spec cũ của SKU khác)')
                canh = []
                if kq["thieu_model"]:
                    canh.append(f'⚠️ {kq["thieu_model"]} SKU trống model_code')
                if kq["thieu_cate"]:
                    canh.append(f'⚠️ {kq["thieu_cate"]} SKU trống category_code')
                if kq["bo_qua_khong_sku"]:
                    canh.append(f'⚠️ Bỏ qua {kq["bo_qua_khong_sku"]} dòng không có sku')
                if kq["trung_sku"]:
                    canh.append(f'⚠️ Bỏ qua {kq["trung_sku"]} dòng trùng sku (giữ dòng đầu)')
                noi_dung = "\n".join(dong + ([""] + canh if canh else []))
                (messagebox.showwarning if canh else messagebox.showinfo)("Đã nạp file export PIM", noi_dung)
                self.var_status.set(f'✔ Đã nạp {kq["so_sku"]} SKU + {kq["so_o_spec"]} ô spec từ file PIM.')
            self._chay_tac_vu_nen(_lam, _xong, ten_thao_tac="Đang nạp file export PIM")
        ttk.Button(top, text="Nạp ngay", style="Accent.TButton", command=_chay).pack(pady=14)
        _can_giua_popup(self, top)
        top.transient(self)
        top.grab_set()
    # ------------------------------------------------------------------
    # BỔ SUNG: chọn đơn vị kích thước / khối lượng hàng loạt
    # ------------------------------------------------------------------
    def _mo_dialog_don_vi(self) -> None:
        file_path = self.var_file.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu workspace", "Chưa có file workspace — thử khởi động lại app.")
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy workspace:\n{p}")
            return
        self._chay_tac_vu_nen(
            lambda: quet_cot_kich_thuoc(p), lambda kq: self._hien_dialog_don_vi(p, kq),
            ten_thao_tac="Đang quét các cột kích thước / khối lượng",
        )
    def _hien_dialog_don_vi(self, p: Path, kq: dict) -> None:
        if kq.get("loi"):
            messagebox.showerror("Lỗi khi quét", kq["loi"])
            return
        nganh: Dict[str, dict] = kq["nganh"]
        if not nganh:
            messagebox.showinfo(
                "Chưa có ngành hàng",
                "Chưa có ngành hàng nào trong CẤU HÌNH CATEGORY hoặc tab TSKT — nạp cấu hình trước.",
            )
            return
        self.var_status.set("Sẵn sàng.")
        top = tk.Toplevel(self)
        top.title("📏 Đơn vị kích thước / khối lượng")
        top.geometry("1060x640")
        bien: Dict[Tuple[str, str], tk.StringVar] = {}
        nhan_cate = {f"{c} — {d['ten']}" if d["ten"] else c: c for c, d in nganh.items()}
        frm_tren = ttk.Frame(top)
        frm_tren.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(frm_tren, text="Ngành hàng:").pack(side="left")
        var_cate = tk.StringVar(value=next(iter(nhan_cate)))
        cbo = ttk.Combobox(frm_tren, textvariable=var_cate, values=list(nhan_cate), state="readonly", width=34)
        cbo.pack(side="left", padx=(6, 16))
        var_tat_ca = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_tren, text="Hiện TẤT CẢ cột TSKT (không chỉ cột tự nhận diện là kích thước/khối lượng)",
                        variable=var_tat_ca, command=lambda: _ve()).pack(side="left")
        ttk.Label(
            top,
            text="Đơn vị CHỈ được thêm vào ô là SỐ TRƠN (vd 9 → 9 kg, 6.95 → 6.95 cm). Ô đã có chữ "
            "('13 kg', 'Không', 'Ngang 122 cm - ...') giữ nguyên. Cột FILTER không bao giờ bị đụng tới. "
            f"Lựa chọn lưu ở sheet \"{CFG_sheet_don_vi}\" và TỰ ÁP LẠI mỗi lần CHẠY TẤT CẢ. "
            "Để trống = không thêm đơn vị. Có thể gõ đơn vị khác (vd: lít, W) vào ô.",
            wraplength=1020, justify="left", foreground="#555",
        ).pack(anchor="w", padx=12, pady=(0, 6))
        frm_nhanh = ttk.Frame(top)
        frm_nhanh.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Button(frm_nhanh, text="✨ Điền theo gợi ý (chỉ ô đang trống)",
                   command=lambda: _dien_goi_y()).pack(side="left")
        ttk.Label(frm_nhanh, text="   Đặt cùng 1 đơn vị cho mọi cột đang hiện:").pack(side="left")
        var_nhanh = tk.StringVar(value="cm")
        ttk.Combobox(frm_nhanh, textvariable=var_nhanh, values=DON_VI_LUA_CHON, width=7).pack(side="left", padx=4)
        ttk.Button(frm_nhanh, text="Áp", command=lambda: _dat_tat_ca(var_nhanh.get())).pack(side="left")
        ttk.Button(frm_nhanh, text="🧹 Xoá chọn (cột đang hiện)", command=lambda: _dat_tat_ca("")).pack(
            side="left", padx=(12, 0))
        # vùng cuộn
        frm_khung = ttk.Frame(top)
        frm_khung.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        canvas = tk.Canvas(frm_khung, background="#ffffff", highlightthickness=0)
        vs = ttk.Scrollbar(frm_khung, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        noi_dung = tk.Frame(canvas, background="#ffffff")
        canvas.create_window((0, 0), window=noi_dung, anchor="nw")
        noi_dung.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        def _cuon(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)) if e.delta else 0, "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _cuon))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        var_dem = tk.StringVar()
        dang_hien: List[Tuple[str, str]] = []
        def _cot_hien(cate: str) -> List[dict]:
            ds = nganh[cate]["cot"]
            if not var_tat_ca.get():
                ds = [c for c in ds if c["la_kt"] or c["da_luu"]]
            return sorted(ds, key=lambda c: (not c["la_kt"], -c["so_tron"], c["code"]))
        def _ve() -> None:
            for w in noi_dung.winfo_children():
                w.destroy()
            dang_hien.clear()
            cate = nhan_cate[var_cate.get()]
            info = nganh[cate]
            tieu_de = ["Mã TSKT", "Tên", "Số trơn / có dữ liệu", "Ví dụ giá trị hiện tại", "Gợi ý", "ĐƠN VỊ"]
            for j, t in enumerate(tieu_de):
                tk.Label(noi_dung, text=t, font=("Segoe UI", 10, "bold"), background="#2f6fed",
                         foreground="white", padx=6, pady=4, anchor="w").grid(row=0, column=j, sticky="we")
            ds = _cot_hien(cate)
            for i, c in enumerate(ds, start=1):
                nen = "#f3f5fb" if i % 2 else "#ffffff"
                key = (cate, c["code"])
                if key not in bien:
                    bien[key] = tk.StringVar(value=c["da_luu"])
                dang_hien.append(key)
                mau_so = "#c0392b" if c["so_tron"] and not bien[key].get() else "#1f2430"
                o = [c["code"], c["ten"], f'{c["so_tron"]} / {c["tong"]}', c["vi_du"], c["goi_y"] or "-"]
                for j, v in enumerate(o):
                    tk.Label(noi_dung, text=v, background=nen, anchor="w", padx=6, pady=3,
                             foreground=mau_so if j == 2 else "#1f2430").grid(row=i, column=j, sticky="we")
                ttk.Combobox(noi_dung, textvariable=bien[key], values=DON_VI_LUA_CHON, width=8).grid(
                    row=i, column=5, sticky="w", padx=6, pady=2)
            if not ds:
                tk.Label(noi_dung, text="(Không tự nhận diện được cột kích thước/khối lượng nào — tick "
                         "'Hiện TẤT CẢ cột TSKT' để chọn thủ công)", background="#ffffff",
                         foreground="#777", pady=12).grid(row=1, column=0, columnspan=6, sticky="w")
            tab_txt = f'tab "{info["tab"]}"' if info["tab"] else "CHƯA có tab TSKT (sẽ áp ở lần CHẠY TẤT CẢ)"
            var_dem.set(f"{len(ds)} cột đang hiện — {tab_txt}. Số màu đỏ = cột còn ô số trơn mà chưa chọn đơn vị.")
            canvas.yview_moveto(0)
        def _dien_goi_y() -> None:
            cate = nhan_cate[var_cate.get()]
            goi_y = {c["code"]: c["goi_y"] for c in nganh[cate]["cot"]}
            for key in dang_hien:
                if not bien[key].get() and goi_y.get(key[1]):
                    bien[key].set(goi_y[key[1]])
            _ve()
        def _dat_tat_ca(dv: str) -> None:
            for key in dang_hien:
                bien[key].set(dv)
            _ve()
        cbo.bind("<<ComboboxSelected>>", lambda e: _ve())
        ttk.Label(top, textvariable=var_dem, foreground="#555").pack(anchor="w", padx=12)
        frm_duoi = ttk.Frame(top)
        frm_duoi.pack(fill="x", padx=12, pady=10)
        def _luu() -> None:
            chon: Dict[str, Dict[str, str]] = defaultdict(dict)
            ten_tskt: Dict[Tuple[str, str], str] = {}
            ten_map = {(cate, c["code"]): c for cate, d in nganh.items() for c in d["cot"]}
            for (cate, code), var in bien.items():
                moi = chuan_hoa_key(var.get())
                if moi == ten_map[(cate, code)]["da_luu"]:
                    continue  # không đổi -> không ghi lại, không áp lại
                chon[cate][code] = moi
                ten_tskt[(cate, code)] = ten_map[(cate, code)]["ten"]
            if not chon:
                messagebox.showinfo("Không có thay đổi", "Chưa đổi đơn vị nào so với lần lưu trước.", parent=top)
                return
            so_cot = sum(len(v) for v in chon.values())
            ten_cate = {c: d["ten"] for c, d in nganh.items()}
            top.destroy()
            def _xong(kq2: dict) -> None:
                if kq2.get("loi"):
                    messagebox.showerror("Lỗi khi lưu đơn vị", kq2["loi"])
                    return
                chi_tiet = "\n".join(f"  • NH {c}: thêm đơn vị cho {n} ô" for c, n in kq2["so_o"].items())
                messagebox.showinfo(
                    "Đã lưu đơn vị",
                    f"Đã lưu {so_cot} thay đổi vào sheet \"{CFG_sheet_don_vi}\" "
                    f"(tổng {kq2['so_dong_luu']} cột đang có đơn vị).\n"
                    + (f"Áp ngay vào tab TSKT hiện có:\n{chi_tiet}\n" if chi_tiet else "")
                    + "\nCác lần CHẠY TẤT CẢ sau sẽ tự áp lại.",
                )
                self.var_status.set(f"✔ Đã lưu {so_cot} thay đổi đơn vị kích thước/khối lượng.")
            self._chay_tac_vu_nen(lambda: luu_va_ap_don_vi(p, chon, ten_cate, ten_tskt), _xong,
                                  ten_thao_tac="Đang lưu và áp đơn vị")
        ttk.Button(frm_duoi, text="💾 Lưu & áp vào tab TSKT", style="Accent.TButton", command=_luu).pack(side="left")
        ttk.Button(frm_duoi, text="Đóng", command=top.destroy).pack(side="right")
        _ve()
        top.transient(self)
        top.grab_set()
    def _chon_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file Excel nguồn",
            filetypes=[("Excel files", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if path:
            self.var_file.set(path)
            if not self.var_out.get():
                self.var_out.set(str(Path(path).parent / "output"))
    def _chon_out_dir(self) -> None:
        path = filedialog.askdirectory(title="Chọn thư mục xuất kết quả")
        if path:
            self.var_out.set(path)
    def _ghi_log(self, s: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", s)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
    def _xoa_log(self) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
    def _bom_log(self) -> None:
        try:
            while True:
                s = self._log_queue.get_nowait()
                self._ghi_log(s)
        except queue.Empty:
            pass
        try:
            while True:
                out_dir, loi, kq = self._done_queue.get_nowait()
                self._chay_xong(out_dir, loi, kq)
        except queue.Empty:
            pass
        self.after(100, self._bom_log)
    def _bam_chay(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Đang chạy", "Tool đang xử lý, đợi xong rồi bấm lại.")
            return
        file_path = self.var_file.get().strip()
        out_path = self.var_out.get().strip()
        if not file_path:
            messagebox.showwarning("Thiếu file", "Chọn file Excel nguồn trước đã.")
            return
        wb_path = Path(file_path)
        if not wb_path.exists():
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy file:\n{wb_path}")
            return
        if not out_path:
            out_path = str(wb_path.parent / "output")
            self.var_out.set(out_path)
        out_dir = Path(out_path)
        self._xoa_log()
        self.var_status.set("Đang xử lý — có thể mất chút thời gian nếu file lớn, đừng tắt cửa sổ...")
        self._dat_trang_thai_nut("disabled")
        self.btn_open_out.configure(state="disabled")
        self.progress.start(12)
        self._da_xong_thanh_cong = False
        self._worker = threading.Thread(target=self._chay_nen, args=(wb_path, out_dir), daemon=True)
        self._worker.start()
    def _chay_nen(self, wb_path: Path, out_dir: Path) -> None:
        old_stdout = sys.stdout
        sys.stdout = _StdoutRedirect(self._log_queue)
        loi: str | None = None
        kq = None
        try:
            kq = chay_tat_ca(wb_path, out_dir)
        except Exception:  # noqa: BLE001
            loi = traceback.format_exc()
            self._log_queue.put("\n" + "=" * 70 + "\n")
            self._log_queue.put("LỖI — xem chi tiết bên trên để biết chỗ nào cần sửa:\n")
            self._log_queue.put(loi)
        finally:
            sys.stdout = old_stdout
        self._done_queue.put((out_dir, loi, kq))
    def _chay_xong(self, out_dir: Path, loi: str | None, kq: Optional[dict] = None) -> None:
        self.progress.stop()
        self._dat_trang_thai_nut("normal")
        if loi:
            self.var_status.set("Đã dừng vì có lỗi — xem nhật ký bên dưới.")
            messagebox.showerror(
                "Có lỗi xảy ra",
                "Tool dừng lại vì gặp lỗi. Xem chi tiết trong khung Nhật ký chạy bên dưới.",
            )
            return
        self._da_xong_thanh_cong = True
        self._out_dir = out_dir
        self.var_status.set("✔ Xong! Đã cập nhật file Excel nguồn và tạo file .zip kết quả.")
        self.btn_open_out.configure(state="normal")
        if kq and kq.get("co_canh_bao"):
            self.var_status.set("⚠️ Đã xuất xong nhưng CÓ CẢNH BÁO — đọc sheet CẢNH BÁO XUẤT trước khi import.")
            self._hien_canh_bao_xuat(kq.get("tom_tat", ""))
            return
        messagebox.showinfo("Hoàn tất", "Đã chạy xong toàn bộ pipeline — không có cảnh báo.\n"
                            "Bấm 'Mở thư mục kết quả' để xem file .zip.")
    def _mo_thu_muc_ket_qua(self) -> None:
        if not self._out_dir:
            return
        path = self._out_dir
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as e:  # noqa: BLE001
            messagebox.showwarning("Không mở được thư mục", f"Đường dẫn: {path}\n\nLỗi: {e}")
# ============================================================================
# ĐIỂM VÀO CHƯƠNG TRÌNH
# ============================================================================
def _chay_cli(wb_path: Path, out_dir: Path) -> None:
    if not wb_path.exists():
        print(f"Không tìm thấy file: {wb_path}")
        sys.exit(1)
    try:
        chay_tat_ca(wb_path, out_dir)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        sys.exit(1)
def main() -> None:
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(
            description="TGDĐ CMS -> PIM — chạy không giao diện (RUN-ALL). Không truyền gì = mở GUI."
        )
        parser.add_argument("workbook", type=str, help="Đường dẫn file Excel nguồn (.xlsx)")
        parser.add_argument("--out", type=str, default="./output", help="Thư mục xuất file .zip kết quả")
        args = parser.parse_args()
        _chay_cli(Path(args.workbook), Path(args.out))
        return
    app = PimApp()
    app.mainloop()
if __name__ == "__main__":
    main()