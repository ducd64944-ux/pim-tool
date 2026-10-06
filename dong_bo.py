# -*- coding: utf-8 -*-
"""
dong_bo.py — GỘP 3 CHIỀU khi 2 máy cùng sửa 1 file (giống cách git merge, nhưng theo nghĩa dữ liệu):

    gốc  = bản mình đã đọc lúc đầu
    mình = bản mình định lưu
    họ   = bản đang có trên kho (máy khác vừa lưu)
    -> kết quả = "họ" + các thay đổi CỦA MÌNH so với "gốc"

Nhờ vậy 2 người cùng thêm quy đổi FILTER / cùng sửa tay / cùng gửi đề xuất / admin duyệt trong khi thành viên
gửi… đều không mất của ai. Chỉ khi CẢ HAI cùng sửa ĐÚNG 1 khoá thành 2 giá trị khác nhau thì bản của mình thắng
(và được báo lại).
"""
from __future__ import annotations

import json
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

_XOA = object()


def tron_dict(goc: dict, minh: dict, ho: dict) -> Tuple[dict, List[str]]:
    """Gộp dict phẳng. -> (kết quả, các khoá cả 2 cùng sửa khác nhau)."""
    goc, minh, ho = goc or {}, minh or {}, ho or {}
    out = dict(ho)
    va_cham = []
    for k in set(goc) | set(minh):
        g, m = goc.get(k, _XOA), minh.get(k, _XOA)
        if _bang(g, m):
            continue  # mình không đổi khoá này -> giữ của họ
        h = ho.get(k, _XOA)
        if not _bang(h, g) and not _bang(h, m):
            va_cham.append(str(k))
        if m is _XOA:
            out.pop(k, None)
        else:
            out[k] = m
    return out, va_cham


def _bang(a, b) -> bool:
    if a is _XOA or b is _XOA:
        return a is b
    return json.dumps(a, sort_keys=True, ensure_ascii=False) == json.dumps(b, sort_keys=True, ensure_ascii=False)


def tron_settings(goc: dict, minh: dict, ho: dict) -> Tuple[dict, List[str]]:
    """settings.json của workspace: gộp từng bảng con (sua / don_vi / rong), khoá đơn lẻ: của mình thắng."""
    goc, minh, ho = goc or {}, minh or {}, ho or {}
    out = dict(ho)
    va_cham: List[str] = []
    for k in set(goc) | set(minh) | set(ho):
        g, m, h = goc.get(k), minh.get(k), ho.get(k)
        if isinstance(m, dict) or isinstance(h, dict):
            out[k], vc = tron_dict(g if isinstance(g, dict) else {}, m if isinstance(m, dict) else {},
                                   h if isinstance(h, dict) else {})
            va_cham += [f"{k}:{x}" for x in vc]
        elif k in minh:
            if not _bang(g, m) or k not in ho:
                out[k] = m
        elif k in goc and k in ho and _bang(g, h):  # mình xoá khoá, họ không đổi -> xoá
            out.pop(k, None)
    return out, va_cham


def tron_ds_id(goc: list, minh: list, ho: list, khoa: str = "id") -> Tuple[list, List[str]]:
    """Danh sách bản ghi có id (đề xuất): thêm/sửa/xoá theo id."""
    def theo(ds):
        return {str(x.get(khoa)): x for x in (ds or []) if isinstance(x, dict)}
    out, vc = tron_dict(theo(goc), theo(minh), theo(ho))
    thu_tu = [str(x.get(khoa)) for x in (ho or []) if isinstance(x, dict)]
    them = [k for k in out if k not in set(thu_tu)]
    return [out[k] for k in thu_tu if k in out] + [out[k] for k in them], vc


def tron_ds(goc: list, minh: list, ho: list) -> Tuple[list, List[str]]:
    """Danh sách thường (lịch sử xuất): họ + phần mình thêm, bỏ phần mình xoá."""
    def k(x):
        return json.dumps(x, sort_keys=True, ensure_ascii=False)
    g, m = {k(x) for x in goc or []}, {k(x) for x in minh or []}
    them = [x for x in minh or [] if k(x) not in g]
    xoa = g - m
    return [x for x in ho or [] if k(x) not in xoa] + [x for x in them if k(x) not in {k(y) for y in ho or []}], []


def tron_bang(goc: pd.DataFrame, minh: pd.DataFrame, ho: pd.DataFrame,
              khoa: Optional[List[str]] = None) -> Tuple[pd.DataFrame, List[str]]:
    """Bảng (mapping, DATA PIM…): theo DÒNG. Dòng mình thêm -> thêm; dòng mình xoá -> xoá khỏi bản của họ.
    Có `khoa` (vd cate+prop_id): sau gộp mỗi khoá chỉ giữ 1 dòng, ưu tiên dòng của mình."""
    cols = list(minh.columns)

    def tap(df):
        if df is None or not len(df):
            return []
        d = df.copy()
        for c in cols:
            if c not in d.columns:
                d[c] = ""
        return [tuple(str(v) for v in r) for r in d[cols].itertuples(index=False)]
    g, m, h = tap(goc), tap(minh), tap(ho)
    sg, sm = set(g), set(m)
    them = [r for r in m if r not in sg]
    xoa = sg - sm
    ket = [r for r in h if r not in xoa]
    s_ket = set(ket)
    ket += [r for r in them if r not in s_ket]
    va_cham: List[str] = []
    if khoa and them:
        # chỉ xử lý trùng khoá ở các khoá MÌNH vừa thêm/sửa: dòng của mình thắng, dòng khác cùng khoá bị bỏ.
        # Trùng khoá có sẵn từ trước (không liên quan) giữ nguyên thứ tự -> mapping hiệu lực không đổi.
        ik = [cols.index(c) for c in khoa]
        k_them = {tuple(r[i] for i in ik) for r in them}
        s_them = set(them)
        k_goc = {tuple(r[i] for i in ik): r for r in g}
        for r in h:
            kk = tuple(r[i] for i in ik)
            if kk in k_them and r not in s_them and r not in sg and k_goc.get(kk) != r:
                va_cham.append(" / ".join(kk))  # họ cũng vừa sửa đúng khoá này
        ket = [r for r in ket if tuple(r[i] for i in ik) not in k_them or r in s_them]
    out = pd.DataFrame(ket, columns=cols)
    return out.reset_index(drop=True), va_cham[:50]


def tron_json(loai: str, goc: Any, minh: Any, ho: Any) -> Tuple[Any, List[str]]:
    if loai == "settings":
        return tron_settings(goc, minh, ho)
    if loai == "ds_id":
        return tron_ds_id(goc, minh, ho)
    if loai == "ds":
        return tron_ds(goc, minh, ho)
    if isinstance(minh, dict):
        if loai == "cau_hinh":  # gộp theo ngành; trong 1 ngành bản của mình thắng
            return tron_dict(goc, minh, ho)
        return tron_dict(goc, minh, ho)
    raise ValueError(f"Không gộp được kiểu {loai}")
