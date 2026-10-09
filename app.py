# -*- coding: utf-8 -*-
"""
PIM Tool (web) — TGDĐ CMS -> PIM: map TSKT + FILTER, kiểm tra/đối chiếu, sửa trực tiếp, xuất file import.
Bản Streamlit của tool desktop 66.py. Dữ liệu lưu GitHub (xem gh_store.py), mỗi tài khoản 1 workspace.
"""
from __future__ import annotations

import contextlib
import gc
import hashlib
import hmac
import os
import random
import secrets
import re
import threading
import time
import traceback
from collections import Counter
from datetime import datetime, timezone, timedelta

import pandas as pd
import streamlit as st
import base64 as _b64
from pathlib import Path as _Path

_ASSETS = _Path(__file__).parent / "assets"
_LOGO = _ASSETS / "logo.png"
_FAVICON = _ASSETS / "favicon.png"
_DMX_BANNER = _ASSETS / "dmx_banner.png"  # logo Điện máy XANH (nền xanh)


def _b64_file(ten: str) -> str:
    try:
        return _b64.b64encode((_ASSETS / ten).read_bytes()).decode()
    except Exception:  # noqa: BLE001 - thiếu file logo thì dùng icon chữ
        return ""


LOGO_B64 = _b64_file("logo_nho.png") or _b64_file("logo.png")
DMX_CHU_B64 = _b64_file("dmx_chu.png")       # chữ "Điện máy XANH" nền trong suốt
DMX_BANNER_B64 = _b64_file("dmx_banner.png")
LOI_VUI_B64 = _b64_file("loi_vui.png")  # hình vui khi lỗi hệ thống (non-admin)

# Múi giờ Việt Nam (UTC+7)
_TZ_VN = timezone(timedelta(hours=7))


def logo_img(px: int) -> str:
    return (f'<img src="data:image/png;base64,{LOGO_B64}" width="{px}" height="{px}" '
            f'style="display:block;border-radius:50%" alt="logo">') if LOGO_B64 else "🧩"


def dmx_chu_img(cao: int) -> str:
    return (f'<img src="data:image/png;base64,{DMX_CHU_B64}" height="{cao}" style="display:block;height:{cao}px;'
            f'width:auto" alt="Điện máy XANH">') if DMX_CHU_B64 else ""


def dmx_banner_img(rong: int) -> str:
    return (f'<img src="data:image/png;base64,{DMX_BANNER_B64}" style="display:block;width:{rong}px;max-width:100%;'
            f'height:auto;border-radius:14px;margin:0 auto;box-shadow:0 8px 24px rgba(0,120,200,.25)" '
            f'alt="Điện máy XANH">') if DMX_BANNER_B64 else logo_img(72)


st.set_page_config(page_title="PIM Tool — CMS → PIM", page_icon=str(_FAVICON) if _FAVICON.exists() else "🧩",
                   layout="wide",
                   menu_items={"Get Help": None, "Report a bug": None, "About": None})

import pandas as _pd_de

_goc_data_editor = st.data_editor


def _data_editor_an_toan(data, *a, **kw):
    """Bảng rỗng: pandas để cột kiểu số (float) → cột chữ/chọn báo lỗi StreamlitAPIException. Đổi kiểu cho khớp."""
    try:
        if isinstance(data, _pd_de.DataFrame) and len(data) == 0:
            data = data.copy()
            cc = kw.get("column_config") or {}
            for c in data.columns:
                cf = cc.get(c)
                t = (cf.get("type_config") or {}).get("type") if isinstance(cf, dict) else None
                if t == "checkbox":
                    data[c] = data[c].astype(bool)
                elif t in ("number", "progress"):
                    data[c] = _pd_de.to_numeric(data[c], errors="coerce").astype(float)
                else:
                    data[c] = data[c].astype(object)
    except Exception:  # noqa: BLE001
        pass
    return _goc_data_editor(data, *a, **kw)


st.data_editor = _data_editor_an_toan

_goc_tabs = st.tabs
_RE_SO_TAB = __import__("re").compile(r"\s*\(\s*[\d.,]+\s*\)|\s*✔\s*$")


def _tabs_on_dinh(tabs, *a, **kw):
    """Tên tab có SỐ ĐẾM thay đổi (vd "Nhất quán ngành (3)") làm Streamlit coi là bộ tab mới → nhảy về tab đầu
    mỗi lần bấm. Bỏ phần số đếm khỏi tên để tab đang mở được giữ nguyên."""
    try:
        tabs = [_RE_SO_TAB.sub("", str(t)).strip() or str(t) for t in tabs]
    except Exception:  # noqa: BLE001
        pass
    try:  # nhớ tab đang mở theo khoá cố định -> bấm tab không còn nhảy về tab khác / phải bấm 2 lần
        if "key" not in kw:
            kw = dict(kw, key="tabs_" + format(__import__("zlib").crc32("|".join(tabs).encode()), "x"),
                      on_change="rerun")
        return _goc_tabs(tabs, *a, **kw)
    except Exception:  # noqa: BLE001  (trùng khoá / phiên bản cũ) -> dùng tab thường
        kw.pop("key", None)
        kw.pop("on_change", None)
        return _goc_tabs(tabs, *a, **kw)


st.tabs = _tabs_on_dinh

if _DMX_BANNER.exists():
    st.logo(str(_DMX_BANNER), size="large", icon_image=str(_LOGO) if _LOGO.exists() else None)
elif _LOGO.exists():
    st.logo(str(_LOGO), size="large", icon_image=str(_LOGO))

# ---- Bảo đảm các module phụ (pim_core, gh_store…) là BẢN MỚI sau khi triển khai lại ----
# Máy chủ giữ module đã import trong bộ nhớ (và đã tắt theo dõi file) nên sau khi cập nhật code, app.py mới có thể chạy với
# pim_core cũ -> lỗi "module 'pim_core' has no attribute …". Phát hiện lệch -> nạp lại module, không cần Reboot app.
import importlib as _il  # noqa: E402
import sys as _sys  # noqa: E402

_CORE_CAN = "2026-10-09.4"  # phải khớp pim_core.CORE_VERSION; đổi cả 2 nơi mỗi khi pim_core thêm hàm/hằng mới


def _bao_dam_module_moi() -> None:
    for _n in ("gh_store", "pim_core", "ai_helper", "dong_bo"):
        _m = _sys.modules.get(_n)
        _f = getattr(_m, "__file__", None)
        if _m is None or not _f:
            continue
        try:
            _mt = os.path.getmtime(_f)
            if _n == "pim_core":
                _cu = getattr(_m, "CORE_VERSION", None) != _CORE_CAN
            else:
                _cu = getattr(_m, "_PIM_MT", _mt) != _mt
            if _cu:
                _il.reload(_m)
            _m._PIM_MT = _mt
        except Exception:  # noqa: BLE001  — không nạp lại được thì chạy tiếp với bản đang có
            pass


_bao_dam_module_moi()
import ai_helper as AIH  # noqa: E402
import pim_core as C  # noqa: E402
import dong_bo as DB  # noqa: E402
from gh_store import KHONG_CO, Store, bytes_to_df, bytes_to_json, df_to_bytes, git_sha, json_to_bytes  # noqa: E402

HIEN_AI = False  # ẨN toàn bộ tính năng AI (tab, tự học, cấu hình). Bật lại: đổi True — mã AI vẫn giữ nguyên.
LIEN_HE = {"ten": "Đức Content 234766", "sdt": "0326606655", "line": "1756070012", "email": "nguyenduc6655@gmail.com"}


def the_lien_he() -> None:
    """Tên người làm: bấm vào mở khung liên hệ — khung GIỮ NGUYÊN tới khi bấm ra ngoài, mỗi dòng có nút sao chép."""
    lh = LIEN_HE
    with st.popover(f"📇 {lh['ten']}", width="stretch"):
        st.markdown("**Liên hệ nếu lỗi**")
        for nhan, k in (("📞 SĐT", "sdt"), ("💬 LINE ID", "line"), ("✉️ Email", "email")):
            if lh.get(k):
                a, b = st.columns([1, 2.4], vertical_alignment="center", gap="small")
                a.markdown(f"<span style='font-size:.95rem'>{nhan}</span>", unsafe_allow_html=True)
                with b:
                    st.code(lh[k], language=None)  # có nút sao chép ở góc phải


APP_VERSION = "web-4.19 · 2026-10-09"
ss = st.session_state


# ============================================================================
# HỆ THỐNG GHI LỖI (chỉ admin xem được)
# ============================================================================
def _ghi_loi(err: Exception, noi: str = "") -> None:
    """Ghi lỗi vào ss._nhat_ky_loi cho admin xem. Non-admin không thấy chi tiết."""
    ss.setdefault("_nhat_ky_loi", [])
    ss._nhat_ky_loi.append({
        "luc": datetime.now(_TZ_VN).strftime("%H:%M:%S %d/%m"),
        "noi": noi,
        "loai": type(err).__name__,
        "chi_tiet": traceback.format_exception(err)[-1].strip(),
        "day_du": "".join(traceback.format_exception(err))
    })
    # giữ tối đa 50 lỗi gần nhất
    if len(ss._nhat_ky_loi) > 50:
        ss._nhat_ky_loi = ss._nhat_ky_loi[-50:]


def _hien_loi_vui() -> None:
    """Non-admin: hiện hình vui thay vì lỗi kỹ thuật."""
    if LOI_VUI_B64:
        st.markdown(
            '<div style="text-align:center;padding:40px 20px">'
            f'<img src="data:image/png;base64,{LOI_VUI_B64}" style="max-width:360px;width:100%;border-radius:16px;'
            'box-shadow:0 8px 32px rgba(0,0,0,.15)" alt="Hệ thống đang bảo trì">'
            '<p style="margin-top:16px;font-size:1.1rem;color:#475569;font-weight:600">'
            '🔧 Hệ thống đang xử lý — vui lòng thử lại sau ít phút!</p></div>',
            unsafe_allow_html=True)
    else:
        st.info("🔧 Hệ thống đang xử lý — vui lòng thử lại sau ít phút!")


def _gio_vn() -> str:
    """Trả về giờ Việt Nam dạng '11:01 PM'."""
    now = datetime.now(_TZ_VN)
    return now.strftime("%-I:%M %p").replace("AM", "AM").replace("PM", "PM")

# CSS: KHÔNG được có dòng trống bên trong (Markdown sẽ kết thúc khối HTML ở dòng trống -> CSS bị in ra thành chữ)
_CSS = """<style>
:root {
  --ink: #0f172a; --ink-2: #1e293b; --ink-3: #475569; --line: #e2e8f0; --line-2: #f1f5f9;
  --bg: #ffffff; --bg-2: #f8fafc; --bg-3: #f1f5f9;
  --brand: #1e40af; --brand-2: #3b82f6; --brand-tint: #eff6ff;
  --ok: #15803d; --ok-tint: #dcfce7; --warn: #b45309; --warn-tint: #fef3c7;
  --err: #b91c1c; --err-tint: #fee2e2;
}
html, body, [class*="css"], .stApp {font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;}
#MainMenu, footer, .viewerBadge_container__r5tak, .viewerBadge_link__qRIco,
div[data-testid="manage-app-button"], .stDeployButton,
a[href*="streamlit.io"], a[href*="streamlit.app"], a[href*="streamlit.cloud"],
a[href*="github.com"][target="_blank"][style*="fixed"],
div[class*="viewerBadge"], span[class*="viewerBadge"],
iframe[title="badge"], .st-emotion-cache-h4xjwg,
div[data-testid="manage-app-button"] ~ div,
div.stActionButton, ._profileContainer_gzau3_53,
div[class*="_profileContainer"], a[class*="_profileContainer"],
div[class*="StatusWidget"], div[class*="_hostBadge"],
a[class*="_hostBadge"] {display: none !important; visibility: hidden !important; height: 0 !important; width: 0 !important; overflow: hidden !important; position: absolute !important; left: -9999px !important;}
footer, footer * {visibility: hidden !important; display: none !important; height: 0 !important;}
/* Giữ header (chứa chỉ báo "Đang chạy…" khi bấm) nhưng trong suốt; ẩn nút Deploy/menu */
header[data-testid="stHeader"] {background: transparent; height: 2.4rem;}
div[data-testid="stToolbar"] {visibility: hidden;}
div[data-testid="stStatusWidget"] {background: #fff; border: 1px solid var(--line);
  border-radius: 999px; padding: 2px 10px; box-shadow: 0 2px 8px rgba(15,23,42,.12);}
div[data-testid="stDecoration"] {background: linear-gradient(90deg, #1e40af, #60a5fa); height: 3px;}
.block-container {padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1480px;}
.stApp {background: var(--bg-2);}
h1 {font-weight: 700; letter-spacing:-.3px; font-size: 1.4rem; color: var(--ink); margin: .2rem 0 .4rem;}
h2 {font-weight: 700; font-size: 1.1rem; color: var(--ink); margin: .8rem 0 .4rem; letter-spacing:-.2px;}
h3 {font-weight: 600; font-size: 1rem; color: var(--ink); margin: .6rem 0 .3rem;}
h4, h5, h6 {font-weight: 600; color: var(--ink-2);}
p, li, .stMarkdown {color: var(--ink-2); line-height: 1.55;}
hr {border: 0; border-top: 1px solid var(--line); margin: 1rem 0;}
/* Brand bar */
.brand-bar {background: linear-gradient(90deg, #0f172a 0%, #1e40af 100%);
  color: #fff; padding: 12px 20px; border-radius: 12px; margin-bottom: 14px;
  display: flex; align-items: center; justify-content: space-between; gap: 16px;
  box-shadow: 0 2px 8px rgba(15,23,42,.15);}
.brand-bar .logo {font-weight: 800; font-size: 1.15rem; letter-spacing:-.2px; display:flex; gap:10px; align-items:center;}
.brand-bar .logo .mark {display:inline-flex; width:34px; height:34px; border-radius:50%; background:#fff;
  box-shadow:0 0 0 2px rgba(255,255,255,.35);}
.brand-bar .meta {font-size: .84rem; display:flex; gap:14px; color:#fff;}
.brand-bar .meta .chip {background: rgba(255,255,255,.2); color:#fff; padding: 3px 10px; border-radius: 999px; font-weight:500;}
/* Metric tiles */
div[data-testid="stMetric"] {background: var(--bg); border: 1px solid var(--line); border-radius: 10px;
  padding: 10px 14px; box-shadow: 0 1px 2px rgba(15,23,42,.03);}
div[data-testid="stMetricValue"] {font-size: 1.3rem; font-weight: 700; color: var(--ink);}
div[data-testid="stMetricLabel"] {font-size: .78rem; color: var(--ink-3);}
div[data-testid="stMetricLabel"] p {font-size: .78rem; color: var(--ink-3);}
/* Buttons */
.stButton>button {border-radius: 8px; font-weight: 500; border: 1px solid var(--line); background: var(--bg);
  color: var(--ink-2); padding: 6px 14px; transition: all .15s;}
.stButton>button:hover {border-color: var(--brand-2); color: var(--brand); background: var(--brand-tint);}
.stButton>button[kind="primary"] {background: var(--brand); color: #fff; border-color: var(--brand);
  box-shadow: 0 1px 3px rgba(30,64,175,.3); font-weight: 600;}
.stButton>button[kind="primary"]:hover {background: var(--brand-2); border-color: var(--brand-2); color: #fff;}
/* Containers / cards */
div[data-testid="stContainerBorderless"], div[data-testid="stContainer"] {border-radius: 10px;}
div[data-testid="stContainer"][class*="st-emotion"] {background: var(--bg);}
div[data-testid="stExpander"] {border: 1px solid var(--line); border-radius: 10px; background: var(--bg);
  box-shadow: 0 1px 2px rgba(15,23,42,.03);}
div[data-testid="stExpander"] summary {font-weight: 500; color: var(--ink);}
/* Tabs */
div[data-baseweb="tab-list"] {gap: 2px; border-bottom: 1px solid var(--line);}
button[data-baseweb="tab"] {border-radius: 8px 8px 0 0; font-weight: 500; color: var(--ink-3);
  padding: 8px 14px; background: transparent;}
button[data-baseweb="tab"][aria-selected="true"] {color: var(--brand); background: var(--brand-tint);
  font-weight: 600;}
/* Segmented control (vùng làm việc) */
div[data-testid="stSegmentedControl"] button {background: var(--bg); border: 1px solid var(--line);
  color: var(--ink-2); font-weight: 500; padding: 10px 16px;}
div[data-testid="stSegmentedControl"] button[aria-pressed="true"] {background: var(--brand);
  color: #fff; border-color: var(--brand); box-shadow: 0 1px 3px rgba(30,64,175,.25);}
/* Dataframes */
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {border: 1px solid var(--line);
  border-radius: 10px; overflow: hidden;}
/* Alerts */
div[data-testid="stAlert"] {border: 0; padding: 0; background: transparent;}
div[data-testid="stAlertContainer"] {border-radius: 10px;}
div[data-testid="stAlertContentSuccess"], div[data-testid="stAlertContentSuccess"] p {color: #14532d !important;}
div[data-testid="stAlertContentWarning"], div[data-testid="stAlertContentWarning"] p {color: #713f12 !important;}
div[data-testid="stAlertContentError"], div[data-testid="stAlertContentError"] p {color: #7f1d1d !important;}
div[data-testid="stAlertContentInfo"], div[data-testid="stAlertContentInfo"] p {color: #1e3a8a !important;}
div[data-testid="stAlert"] p {font-weight: 500;}
div[data-testid="stCaptionContainer"], div[data-testid="stCaptionContainer"] p {color: #334155 !important;
  opacity: 1 !important;}
.buoc, .buoc * {color: #1e293b;}
.canh, .canh * {color: #7f1d1d;}
/* Sidebar */
section[data-testid="stSidebar"] {background: var(--bg); border-right: 1px solid var(--line);}
section[data-testid="stSidebar"] .stMarkdown {color: var(--ink-2);}
/* ===== PHẢN HỒI CHUỘT / BÀN PHÍM ===== */
.stButton>button, .stDownloadButton>button, .stFormSubmitButton>button {cursor: pointer;
  transition: background .12s, border-color .12s, color .12s, transform .06s, box-shadow .12s;}
.stButton>button:active, .stDownloadButton>button:active, .stFormSubmitButton>button:active {
  transform: translateY(1px) scale(.985); box-shadow: none;}
.stButton>button:disabled, .stDownloadButton>button:disabled, .stFormSubmitButton>button:disabled {
  opacity: .5; cursor: not-allowed; transform: none;}
.stDownloadButton>button {border-radius: 8px; border: 1px solid var(--line); background: var(--bg); font-weight: 500;}
.stDownloadButton>button:hover {border-color: var(--ok); color: var(--ok); background: var(--ok-tint);}
button:focus-visible, [role="tab"]:focus-visible, [role="radio"]:focus-visible, input:focus-visible,
summary:focus-visible {outline: 2px solid var(--brand-2) !important; outline-offset: 2px;}
button[data-baseweb="tab"] {cursor: pointer; transition: background .12s, color .12s;}
button[data-baseweb="tab"]:hover {color: var(--brand); background: var(--line-2);}
div[data-testid="stSegmentedControl"] button {cursor: pointer; transition: background .12s, color .12s;}
div[data-testid="stSegmentedControl"] button:hover:not([aria-pressed="true"]) {background: var(--brand-tint);
  color: var(--brand); border-color: var(--brand-2);}
div[data-testid="stExpander"] summary {cursor: pointer; border-radius: 10px; transition: background .12s;}
div[data-testid="stExpander"] summary:hover {background: var(--line-2);}
label[data-baseweb="checkbox"], label[data-baseweb="radio"] {cursor: pointer;}
label[data-baseweb="checkbox"]:hover span, label[data-baseweb="radio"]:hover div:first-child {
  box-shadow: 0 0 0 3px rgba(59,130,246,.18);}
section[data-testid="stFileUploaderDropzone"] {border: 1.5px dashed #94a3b8; background: var(--bg);
  border-radius: 10px; transition: border-color .12s, background .12s; cursor: pointer;}
section[data-testid="stFileUploaderDropzone"]:hover {border-color: var(--brand-2); background: var(--brand-tint);}
div[data-testid="stTextInput"] input:focus, div[data-testid="stTextArea"] textarea:focus {
  border-color: var(--brand-2) !important; box-shadow: 0 0 0 3px rgba(59,130,246,.15);}
/* Khi đang xử lý, phần đang chờ mờ nhẹ -> người dùng biết app đang chạy */
.stale-element, [data-stale="true"] {opacity: .55;}
/* Nút chính (kể cả nút trong form): chữ trắng rõ trên nền xanh */
button[kind="primary"], button[kind="primaryFormSubmit"] {background: var(--brand) !important; color: #fff !important;
  border-color: var(--brand) !important; font-weight: 600;}
button[kind="primary"] p, button[kind="primaryFormSubmit"] p {color: #fff !important;}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {background: var(--brand-2) !important;
  border-color: var(--brand-2) !important;}
/* Vùng làm việc đang chọn */
div[data-testid="stButtonGroup"] button[aria-checked="true"] {background: var(--brand) !important;
  color: #fff !important; border-color: var(--brand) !important; font-weight: 600;}
div[data-testid="stButtonGroup"] button[aria-checked="true"] p {color: #fff !important;}
div[data-testid="stButtonGroup"] button[aria-checked="false"]:hover {background: var(--brand-tint);
  color: var(--brand); border-color: var(--brand-2);}
/* Trang đăng nhập */
.login-head {text-align: center; margin: 6vh 0 18px;}
.login-logo {width: 72px; height: 72px; margin: 0 auto 12px; border-radius: 50%; font-size: 32px; line-height: 72px;
  display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 22px rgba(59,130,246,.35);}
.login-title {font-size: 1.5rem; font-weight: 800; color: var(--ink); letter-spacing: -.3px;}
.login-sub {font-size: .88rem; color: var(--ink-3); margin-top: 2px;}
div[data-testid="stForm"] {background: var(--bg); border: 1px solid var(--line); border-radius: 12px;
  padding: 18px 18px 8px; box-shadow: 0 2px 10px rgba(15,23,42,.05);}
/* Legacy classes giữ cho tương thích */
.buoc {background: var(--brand-tint); border-left: 3px solid var(--brand); padding: 10px 14px;
  border-radius: 8px; margin: 6px 0 12px; color: var(--ink-2);}
.canh {background: var(--err-tint); border-left: 3px solid var(--err); padding: 10px 14px;
  border-radius: 8px; margin: 6px 0 12px; color: var(--ink-2);}
.pill {display:inline-block; padding: 3px 10px; border-radius: 999px; font-size: .78rem;
  font-weight: 600; margin-right: 6px;}
.pill-cao{background: var(--err-tint); color: var(--err);}
.pill-tb{background: var(--warn-tint); color: var(--warn);}
.pill-thap{background: var(--ok-tint); color: var(--ok);}
.pill-tt{background: var(--brand-tint); color: var(--brand);}
.card {background: var(--bg); border: 1px solid var(--line); border-radius: 10px;
  padding: 14px 16px; margin: 8px 0 12px; box-shadow: 0 1px 2px rgba(15,23,42,.03);}
.card-h {font-weight: 700; color: var(--ink); font-size: 1.02rem;}
.card-s {color: var(--ink-3); font-size: .88rem;}
/* ================= GIAO DIỆN DỄ ĐỌC (web-2.4) — chỉ CSS, không đổi logic ================= */
/* Chữ & nút to, rõ, dễ bấm */
p, li, .stMarkdown, label, .stCheckbox label p, .stRadio label p {font-size: 1rem;}
h3 {font-size: 1.3rem; font-weight: 700;}
h4 {font-size: 1.12rem;}
.stButton>button, .stDownloadButton>button, .stFormSubmitButton>button {min-height: 44px; font-size: 1rem;
  font-weight: 600; padding: 8px 18px;}
.stButton>button p, .stDownloadButton>button p, .stFormSubmitButton>button p {font-size: 1rem; font-weight: 600;}
div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stTextArea"] textarea,
div[data-baseweb="select"] > div {font-size: 1rem; min-height: 44px;}
label[data-baseweb="checkbox"] > span:first-child {transform: scale(1.15);}
/* 3 vùng làm việc: nút lớn, rõ vùng đang ở */
div[data-testid="stButtonGroup"] button {min-height: 54px; font-size: 1.08rem !important; font-weight: 700;}
div[data-testid="stButtonGroup"] button p {font-size: 1.08rem !important; font-weight: 700;}
/* TAB: xuống dòng thành các nút viên thuốc — thấy hết tab, không phải cuộn ngang, tab đang chọn nổi bật */
div[data-baseweb="tab-list"] {flex-wrap: wrap; gap: 8px; border-bottom: 0; padding: 4px 0 10px; overflow: visible !important;}
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {display: none !important;}
button[data-baseweb="tab"] {border: 1.5px solid var(--line); border-radius: 999px !important; background: var(--bg);
  padding: 8px 16px !important; min-height: 42px; margin: 0 !important;}
button[data-baseweb="tab"] p {font-size: .98rem !important; font-weight: 600; color: var(--ink-2);}
button[data-baseweb="tab"]:hover {border-color: var(--brand-2); background: var(--brand-tint);}
button[data-baseweb="tab"][aria-selected="true"] {background: var(--brand) !important; border-color: var(--brand);
  box-shadow: 0 2px 8px rgba(30,64,175,.28);}
button[data-baseweb="tab"][aria-selected="true"] p {color: #fff !important;}
div[data-testid="stTabs"] [data-baseweb="tab-panel"] {padding-top: 6px;}
div[data-testid="stTabs"] [role="tablist"] {flex-wrap: wrap !important; gap: 8px; padding: 4px 2px 10px;
  border-bottom: 0 !important; overflow: visible !important; width: 100% !important;}
div[data-testid="stTabs"] [role="tablist"]::before, div[data-testid="stTabs"] [role="tablist"]::after {display: none;}
div[data-testid="stTabs"] > div:first-child {overflow: visible !important; border-bottom: 0 !important;}
div[data-testid="stTabs"] .react-aria-SelectionIndicator {display: none !important;}
div[data-testid="stTab"] {border: 1.5px solid var(--line) !important; border-radius: 999px !important;
  background: var(--bg); padding: 8px 16px !important; min-height: 42px; margin: 0 !important; cursor: pointer;
  display: flex; align-items: center; transition: background .12s, border-color .12s, box-shadow .12s;}
div[data-testid="stTab"] p {font-size: .98rem !important; font-weight: 600 !important; color: var(--ink-2) !important;
  white-space: nowrap;}
div[data-testid="stTab"]:hover {border-color: var(--brand-2) !important; background: var(--brand-tint);}
div[data-testid="stTab"][aria-selected="true"] {background: var(--brand) !important; border-color: var(--brand) !important;
  box-shadow: 0 2px 8px rgba(30,64,175,.28);}
div[data-testid="stTab"][aria-selected="true"] p {color: #fff !important;}
div[data-testid="stTab"]:focus-visible {outline: 3px solid #93c5fd !important; outline-offset: 2px;}
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {white-space: normal !important;
  overflow: visible !important; text-overflow: clip !important; line-height: 1.3 !important;}
[data-testid="stMetricLabel"] p {font-size: .9rem !important; color: var(--ink-2) !important;}
div[data-testid="stMetric"] {min-height: 100%;}
.stApp h1 {font-size: 1.7rem !important; font-weight: 800; margin: .2rem 0 .6rem !important;}
.stApp h1 + div, .stApp h1 span {font-size: inherit;}
section[data-testid="stSidebar"] .stButton>button {font-size: .95rem; padding: 6px 10px;}
section[data-testid="stSidebar"] .stButton>button p {white-space: nowrap; font-size: .95rem;}
/* Ô số liệu = nút bấm mở bảng sửa */
[class*="st-key-the_"]:not([class*="the_dong"]) button {min-height: 84px; justify-content: flex-start; text-align: left;
  padding: 10px 14px; white-space: normal;}
[class*="st-key-the_"]:not([class*="the_dong"]) button p {text-align: left; line-height: 1.35; font-size: .95rem;}
[class*="st-key-the_"]:not([class*="the_dong"]) button p strong {font-size: 1.5rem; display: inline-block;}
/* Nút XEM bự */
[class*="st-key-xem_"] button {min-height: 62px; font-size: 1.15rem; font-weight: 700;}
/* Thanh dữ liệu đang nạp: ô nhỏ gọn, bấm được */
[class*="st-key-kho_"]:not([class*="kho_dong"]):not([class*="kho_ok"]):not([class*="kho_xoa"]) button {min-height: 64px;
  padding: 6px 10px; white-space: normal; line-height: 1.3;}
[class*="st-key-kho_"] button p {font-size: .9rem; text-align: center;}
/* web-2.9: KHÔNG còn hiệu ứng rê chuột đổi kích thước/vị trí (gây giật, lag khi cuộn): chữ luôn hiện đủ, đứng yên */
div[data-testid="stCaptionContainer"] p {font-size: .93rem; line-height: 1.55;}
/* VÙNG đang trỏ chuột: viền xanh + bóng → biết đang làm ở khung nào */
div[data-testid="stExpander"] summary p {font-size: 1.02rem; font-weight: 600;}
div[data-testid="stExpander"] details[open] > summary {background: var(--brand-tint); border-bottom: 1px solid var(--line);}
/* Ô bảng số liệu & thanh trạng thái to hơn chút */
div[data-testid="stMetricValue"] {font-size: 1.5rem;}
div[data-testid="stMetricLabel"] p {font-size: .9rem !important;}
/* Thông báo rõ: chữ đậm vừa, cỡ chuẩn */
div[data-testid="stAlert"] p {font-size: 1rem; line-height: 1.55;}
/* ===== Nhận diện Điện máy XANH (web-2.6) ===== */
.brand-bar {background: linear-gradient(90deg, #005a9e 0%, #0091d5 60%, #00aeef 100%) !important;}
.brand-bar .logo .sep {display:inline-block; width:1.5px; height:24px; background: rgba(255,255,255,.55); margin: 0 4px;}
.brand-bar .meta .chip {background: rgba(0,40,80,.35) !important;}
.login-banner {margin: 0 auto 14px; max-width: 320px;}
section[data-testid="stSidebar"] img[data-testid="stLogo"], [data-testid="stSidebarHeader"] img {height: 2.6rem !important;
  border-radius: 8px;}
/* ===== Sửa sau rà soát (web-2.5) ===== */
/* Ô LỖI / CẢNH BÁO / THÔNG BÁO: luôn hiện ĐỦ chữ, không thu gọn (không được bỏ sót nội dung lỗi) */
div[data-testid="stAlertContainer"], div[data-testid="stAlertContainer"]:hover {max-height: none !important;
  overflow: visible !important; -webkit-mask-image: none !important; mask-image: none !important; box-shadow: none !important;}
/* Thanh bên: hiện đủ chữ */
section[data-testid="stSidebar"] div[data-testid="stCaptionContainer"] {max-height: none !important;
  overflow: visible !important; -webkit-mask-image: none !important; mask-image: none !important;}
/* Màn hình cảm ứng (không rê chuột được): hiện đủ chữ luôn */
@media (hover: none), (pointer: coarse) {
  div[data-testid="stCaptionContainer"], .buoc {max-height: none !important; overflow: visible !important;
    -webkit-mask-image: none !important; mask-image: none !important;}
}
div[data-testid="stCaptionContainer"] {margin-bottom: .35rem;}
/* Nút mở lại thanh bên (nằm trong toolbar đang ẩn) — phải luôn thấy, nhất là trên điện thoại */
div[data-testid="stToolbar"] [data-testid="stExpandSidebarButton"],
div[data-testid="stToolbar"] [data-testid="stExpandSidebarButton"] * {visibility: visible !important;}
/* Chữ trên nút không bị cắt "…": cho xuống dòng */
.stButton>button p, .stDownloadButton>button p, .stFormSubmitButton>button p,
section[data-testid="stSidebar"] .stButton>button p {white-space: normal !important; overflow: visible !important;
  text-overflow: clip !important; line-height: 1.25;}
/* Tab: bỏ mũi tên cuộn (đã xuống dòng), tab dài được xuống dòng, viền rõ hơn */
[data-testid="stTabsScrollLeft"], [data-testid="stTabsScrollRight"] {display: none !important;}
div[data-testid="stTab"] {max-width: 100%; border-color: #94a3b8 !important;}
div[data-testid="stTab"][aria-selected="true"] {border-color: var(--brand) !important;}
div[data-testid="stTab"] p {white-space: normal !important;}
div[data-testid="stTab"]:focus-visible {outline-color: #1d4ed8 !important;}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {background: #1d4ed8 !important;
  border-color: #1d4ed8 !important;}
section[data-testid="stFileUploaderDropzone"] {border-color: #64748b;}
@media (max-width: 640px) {
  div[data-testid="stTab"] {min-height: 36px; padding: 5px 12px !important;}
  div[data-testid="stTab"] p {font-size: .88rem !important;}
  div[data-testid="stTabs"] [role="tablist"] {gap: 6px;}
  .brand-bar, .brand-bar .meta {flex-wrap: wrap;}
  .brand-bar .meta .chip {white-space: nowrap;}
}
</style>
<script>
(function(){var u=function(){var c=document.querySelectorAll('.brand-bar .chip');if(!c.length)return;
var t=c[c.length-1];if(!t.textContent.includes('🕐'))return;
var d=new Date();var h=d.getUTCHours()+7;if(h>=24)h-=24;var m=d.getUTCMinutes();
var ap=h>=12?'PM':'AM';var h12=h%12||12;
t.textContent='🕐 '+h12+':'+(m<10?'0':'')+m+' '+ap;};
u();setInterval(u,30000);})();
(function(){function h(){document.querySelectorAll('a').forEach(function(a){var t=(a.textContent||'').toLowerCase();var hr=a.href||'';if((t.includes('hosted with streamlit')||t.includes('created by')||hr.includes('streamlit.io/cloud')||hr.includes('share.streamlit.io'))&&!a.closest('.stApp .block-container')){a.style.display='none';a.style.visibility='hidden';}});};h();setInterval(h,3000);})();
</script>"""
st.markdown(re.sub(r"\n\s*\n", "\n", _CSS), unsafe_allow_html=True)


# ============================================================================
# CẤU HÌNH TỪ SECRETS
# ============================================================================
def sec(key: str, default=None):
    try:
        v = st.secrets.get(key)  # type: ignore[union-attr]
        if v not in (None, ""):
            return v
    except Exception:  # noqa: BLE001 - chưa có secrets.toml
        pass
    return os.environ.get(key, default)


F_TK = "shared/tai_khoan.json"  # tài khoản thành viên tự tạo (lưu trên kho dữ liệu)
_TEN_DN_RE = re.compile(r"^[a-z0-9._-]{3,30}$")


def tk_secrets() -> dict:
    """[users.<tên>] password = "..." (hoặc "sha256:<hex>"), ten = "...", admin = true/false."""
    try:
        u = st.secrets.get("users")  # type: ignore[union-attr]
        if u:
            return {k: dict(v) for k, v in u.items() if hasattr(v, "items")}  # bỏ qua dòng lạc (vd AI_API_KEY)
    except Exception:  # noqa: BLE001
        pass
    if os.environ.get("PIM_DEV") == "1":
        return {"admin": {"password": "admin", "ten": "Admin (chạy thử)", "admin": True}}
    return {}


@st.cache_data(ttl=15, show_spinner=False)
def tk_kho() -> dict:
    try:
        return bytes_to_json(tao_store().doc(F_TK), {}) or {}
    except Exception:  # noqa: BLE001
        return {}


def ds_tai_khoan(ca_cho_duyet: bool = False) -> dict:
    """Secrets (admin cài sẵn) + tài khoản thành viên tự tạo đã được duyệt. Secrets thắng nếu trùng tên."""
    out = dict(tk_kho())
    if not ca_cho_duyet:
        out = {k: v for k, v in out.items() if v.get("trang_thai", "active") == "active"}
    out.update(tk_secrets())
    return out


def tim_tai_khoan(nhap: str, tk: dict) -> str | None:
    """Khớp tên đăng nhập không phân biệt hoa/thường; chấp nhận cả 'Tên hiển thị' nếu không trùng ai."""
    n = nhap.strip().lower()
    if not n:
        return None
    for k in tk:
        if k.lower() == n:
            return k
    cung_ten = [k for k, v in tk.items() if str(v.get("ten", "")).strip().lower() == n]
    return cung_ten[0] if len(cung_ten) == 1 else None


def bam_mat_khau(mk: str) -> str:
    salt = secrets.token_hex(8)
    return f"pbkdf2${salt}${hashlib.pbkdf2_hmac('sha256', mk.encode(), salt.encode(), 120_000).hex()}"


def dung_mat_khau(luu: str, nhap: str) -> bool:
    luu = str(luu or "")
    if luu.startswith("sha256:"):
        return hmac.compare_digest(luu[7:].lower(), hashlib.sha256(nhap.encode()).hexdigest())
    if luu.startswith("pbkdf2$"):
        try:
            _, salt, h = luu.split("$")
            return hmac.compare_digest(h, hashlib.pbkdf2_hmac("sha256", nhap.encode(), salt.encode(), 120_000).hex())
        except Exception:  # noqa: BLE001
            return False
    return bool(luu) and hmac.compare_digest(luu, nhap)


def sua_tk_kho(ham, thong_diep: str) -> tuple[bool, str]:
    """Đọc-sửa-ghi shared/tai_khoan.json có kiểm tra phiên bản (nhiều người đăng ký cùng lúc không mất của nhau)."""
    S = tao_store()
    for _ in range(6):
        data, ver = S.doc2(F_TK)
        d = bytes_to_json(data, {}) or {}
        loi = ham(d)
        if loi:
            return False, loi
        ok, msg, _, xd = S.luu({F_TK: json_to_bytes(d)}, thong_diep, {F_TK: ver})
        if ok:
            tk_kho.clear()
            return True, ""
        if not xd:
            return False, msg
        time.sleep(0.4 + random.random() * 0.6)
    return False, "Kho đang bận, thử lại sau vài giây."


def dang_ky(u: str, ten: str, mk: str, ma: str) -> tuple[bool, str]:
    u = u.strip().lower()
    if not _TEN_DN_RE.match(u):
        return False, "Tên đăng nhập 3–30 ký tự: chữ thường không dấu, số, . _ -  (vd: an.nguyen)."
    if len(mk) < 6:
        return False, "Mật khẩu tối thiểu 6 ký tự."
    if u in tk_secrets():
        return False, "Tên đăng nhập này đã có người dùng."
    ma_moi = str(sec("MA_MOI", "") or "")
    if ma and ma_moi and not hmac.compare_digest(ma.strip(), ma_moi):
        return False, "Mã mời không đúng."
    dung_ma = bool(ma_moi) and bool(ma)
    ban_ghi = {"password": bam_mat_khau(mk), "ten": ten.strip() or u, "admin": False,
               "trang_thai": "active" if dung_ma else "cho_duyet", "tao_luc": C.bay_gio()}

    def them(d):
        if u in d:
            return "Tên đăng nhập này đã có người dùng."
        d[u] = ban_ghi
        return ""
    ok, loi = sua_tk_kho(them, f"[{u}] Đăng ký tài khoản")
    if not ok:
        return False, loi
    return True, ("Tạo xong — đăng nhập được ngay." if dung_ma else
                  "Đã gửi đăng ký — chờ admin duyệt rồi đăng nhập.")


@st.cache_resource
def tao_store() -> Store:
    token, repo = sec("GITHUB_TOKEN", ""), sec("GITHUB_DATA_REPO") or sec("GITHUB_REPO", "")
    if token and repo:
        return Store("github", repo=repo, branch=sec("GITHUB_BRANCH", "main"), token=token,
                     root=sec("DATA_ROOT", "data"))
    return Store("local", root=sec("DATA_ROOT", "data"), local_dir=sec("LOCAL_DATA_DIR", "_du_lieu_cuc_bo"))


# ============================================================================
# ĐĂNG NHẬP
# ============================================================================
def dang_nhap() -> None:
    if ss.get("user"):
        return
    tk = ds_tai_khoan()
    _, giua, _ = st.columns([1, 1.25, 1])
    with giua:
        st.markdown(
            f'<div class="login-head"><div class="login-banner">{dmx_banner_img(320)}</div>'
            '<div class="login-title">PIM Tool</div>'
            '<div class="login-sub">Chuyển thông số CMS → PIM · TGDĐ / ĐMX</div></div>',
            unsafe_allow_html=True)
        if not tk:
            st.error("Chưa cấu hình tài khoản. Vào Streamlit Cloud → app → Settings → Secrets, thêm:")
            st.code('[users.ducd]\npassword = "mat-khau"\nten = "Đức"\nadmin = true', language="toml")
            st.stop()
        t_dn, t_dk = st.tabs(["🔑 Đăng nhập", "🆕 Tạo tài khoản"])
        with t_dn:
            with st.form("dang_nhap"):
                u = st.text_input("Tài khoản (tên đăng nhập)").strip()
                p = st.text_input("Mật khẩu", type="password")
                ok = st.form_submit_button("Đăng nhập", type="primary", width="stretch")
            if ok:
                k = tim_tai_khoan(u, tk)
                if ss.get("sai_mk", 0) >= 8:
                    st.error("Sai quá nhiều lần — tải lại trang sau ít phút.")
                elif k and dung_mat_khau(tk[k].get("password"), p):
                    ss.user, ss.ten = k, tk[k].get("ten", k)
                    ss.admin = bool(tk[k].get("admin", False))
                    ss.ws = k
                    ss.sai_mk = 0
                    st.rerun()
                else:
                    ss.sai_mk = ss.get("sai_mk", 0) + 1
                    cho = tim_tai_khoan(u, ds_tai_khoan(True))
                    if cho and cho not in tk and cho not in tk_secrets():
                        st.warning("Tài khoản này đang chờ admin duyệt.")
                    else:
                        st.error("Sai tài khoản hoặc mật khẩu. Lưu ý: dùng TÊN ĐĂNG NHẬP (vd `ducd`), không phải tên hiển thị.")
        with t_dk:
            st.caption("Thành viên tự tạo tài khoản. Có **mã mời** từ admin thì dùng được ngay; không có thì chờ admin duyệt.")
            with st.form("dang_ky"):
                u2 = st.text_input("Tên đăng nhập (chữ thường, không dấu)")
                t2 = st.text_input("Tên hiển thị")
                m2 = st.text_input("Mật khẩu (≥ 6 ký tự)", type="password")
                m3 = st.text_input("Nhập lại mật khẩu", type="password")
                ma = st.text_input("Mã mời (nếu có)")
                ok2 = st.form_submit_button("Tạo tài khoản", width="stretch")
            if ok2:
                if m2 != m3:
                    st.error("Hai mật khẩu không giống nhau.")
                else:
                    try:
                        done, msg = dang_ky(u2, t2, m2, ma)
                    except Exception as e:  # noqa: BLE001
                        done, msg = False, f"Lỗi lưu: {e}"
                    (st.success if done else st.error)(msg)
        the_lien_he()
    st.stop()


# ============================================================================
# DỮ LIỆU: NẠP / LƯU
# ============================================================================
F_SHARED = {"cau_hinh": "shared/cau_hinh.json", "quy_doi": "shared/quy_doi_filter.json",
            "sua_sku": "shared/sua_sku.json", "sua_gt": "shared/sua_gia_tri.json",
            "dx_duyet": "shared/de_xuat_duyet.json", "quy_tac_kt": "shared/quy_tac_kiem_tra.json",
            "ai_cau_hinh": "shared/ai_cau_hinh.json", "ai_hoc": "shared/ai_hoc.json",
            "map_tskt": "shared/map_tskt.parquet",
            "map_filter": "shared/map_filter.parquet", "data_pim": "shared/data_pim.parquet"}
F_USER = {"import": "import.parquet", "data_sp": "data_sp.parquet", "spec": "spec.parquet",
          "ket_qua": "ket_qua.parquet", "ket_qua_meta": "ket_qua_meta.json", "settings": "settings.json",
          "lich_su": "lich_su.json"}
COT = {"import": C.COT_IMPORT, "data_sp": C.COT_DATA_SP, "spec": C.COT_SPEC, "map_tskt": C.COT_MAP_TSKT,
       "map_filter": C.COT_MAP_FILTER, "data_pim": C.COT_DATA_PIM}


def p_user(ten: str) -> str:
    return f"users/{ss.ws}/{F_USER[ten]}"


def duoc_sua_chung() -> bool:
    return bool(ss.get("admin")) or str(sec("SHARED_EDIT", "admin")).lower() == "all"


# ============================================================================
# BẢO VỆ TẢI CAO — nhiều người dùng chung 1 máy chủ, lô vài triệu dòng
# ============================================================================
def ram_mb() -> tuple:
    """(đang dùng MB, giới hạn MB hoặc None, còn trống MB) của CONTAINER (cgroup v2/v1), không có thì /proc/meminfo."""
    def _doc(p):
        with open(p, encoding="utf-8") as f:
            return f.read().strip()
    gh = dung = None
    try:
        v = _doc("/sys/fs/cgroup/memory.max")
        gh = None if v == "max" else int(v)
        dung = int(_doc("/sys/fs/cgroup/memory.current"))
        for ln in _doc("/sys/fs/cgroup/memory.stat").splitlines():  # bộ nhớ đệm file thu hồi được -> không tính là đang dùng
            if ln.startswith("inactive_file "):
                dung -= int(ln.split()[1])
                break
    except Exception:  # noqa: BLE001
        try:
            v = int(_doc("/sys/fs/cgroup/memory/memory.limit_in_bytes"))
            gh = v if v < (1 << 50) else None
            dung = int(_doc("/sys/fs/cgroup/memory/memory.usage_in_bytes"))
        except Exception:  # noqa: BLE001
            pass
    mi = {}
    try:
        for ln in _doc("/proc/meminfo").splitlines():
            a, _, b = ln.partition(":")
            mi[a] = int(b.split()[0]) * 1024
    except Exception:  # noqa: BLE001
        pass
    if gh is not None and dung is not None:
        return dung // 2**20, gh // 2**20, max(0, gh - dung) // 2**20
    if mi.get("MemTotal"):
        return (mi["MemTotal"] - mi.get("MemAvailable", 0)) // 2**20, mi["MemTotal"] // 2**20, mi.get("MemAvailable", 0) // 2**20
    return 0, None, 10**9  # không đọc được -> không chặn


@st.cache_resource
def _dieu_phoi() -> dict:
    """Dùng CHUNG cho mọi phiên của cả máy chủ: giới hạn số việc nặng chạy cùng lúc (map/xuất file lô lớn)."""
    try:
        toi_da = int(sec("PIM_MAX_JOBS", 0) or 0)
    except (TypeError, ValueError):
        toi_da = 0
    if toi_da <= 0:
        gh = ram_mb()[1]
        toi_da = 2 if (gh and gh >= 6000) else 1
    return {"sem": threading.BoundedSemaphore(toi_da), "toi_da": toi_da, "dang": {}, "lk": threading.Lock()}


@contextlib.contextmanager
def viec_nang(ten: str, uoc_mb: int):
    """Bao quanh việc NẶNG (map, tạo file import lô lớn): (1) đủ RAM mới chạy — thiếu thì báo rõ thay vì làm sập máy chủ;
    (2) xếp hàng nếu đã có đủ việc nặng đang chạy; (3) dọn bộ nhớ ngay khi xong. Dùng: with viec_nang(..) as ok: if ok: ..."""
    d = _dieu_phoi()
    _, gh, con = ram_mb()
    if con < uoc_mb * 1.15:
        st.warning(f"⏳ Máy chủ đang đầy bộ nhớ (còn ~{con:,} MB, việc «{ten}» cần ~{uoc_mb:,} MB"
                   + (f" trên tổng {gh:,} MB" if gh else "") + "). Chờ người khác xong rồi bấm lại, hoặc chia lô nhỏ hơn. "
                   "Dữ liệu của bạn vẫn an toàn, chưa mất gì.")
        yield False
        return
    co_luot = d["sem"].acquire(blocking=False)
    if not co_luot:
        with d["lk"]:
            ai = ", ".join(sorted({v[0] for v in d["dang"].values()})) or "người khác"
        cho = st.empty()
        cho.info(f"⏳ Đang có {len(d['dang'])} việc nặng chạy ({ai}) — bạn đang xếp hàng, tự chạy khi tới lượt…")
        co_luot = d["sem"].acquire(timeout=300)
        cho.empty()
        if not co_luot:
            st.warning("Hàng đợi đang quá dài — bấm lại sau ít phút. Dữ liệu của bạn vẫn an toàn.")
            yield False
            return
    tid = threading.get_ident()
    with d["lk"]:
        d["dang"][tid] = (ss.get("user", "?"), ten, time.time())
    try:
        yield True
    finally:
        with d["lk"]:
            d["dang"].pop(tid, None)
        d["sem"].release()
        tra_ram()


def tra_ram() -> None:
    """Thu gom rác Python + trả bộ nhớ trống về hệ điều hành (malloc_trim — pandas hay giữ lại vùng nhớ đã giải phóng)."""
    try:
        gc.collect()
        gc.collect()
        import ctypes
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except Exception:  # noqa: BLE001
        pass


# Kết quả tính lại được khi cần (lười) -> xoá an toàn, không mất dữ liệu người dùng
_CACHE_DAN_XUAT = ("ds", "ds_ver", "kc_kq", "kc_ver", "dht", "dht_ver", "nq_kq", "nq_ver", "_xd_kq", "_xd_khoa", "ws_mau")


def don_ram(ca_xuat: bool = False) -> dict:
    """Dọn bộ nhớ của PHIÊN này + trả RAM về máy chủ. KHÔNG đụng dữ liệu đang làm (IMPORT, DATA SP, kết quả map, mapping,
    cấu hình, sửa tay, đơn vị). Xoá: các bảng kiểm tra tính lại được, bảng file vừa đọc ở Nạp nhanh, file báo cáo đã dựng;
    ca_xuat=True: cả các file import đã tạo (phải bấm Tạo file import lại nếu muốn tải)."""
    truoc = ram_mb()[0]
    n = 0
    for k in _CACHE_DAN_XUAT:
        if k in ss:
            ss.pop(k, None)
            n += 1
    for k in list(ss.keys()):  # kết quả đọc file của Nạp nhanh (giữ bảng lớn sau khi nạp xong)
        if isinstance(k, str) and k.endswith("_kq") and k != "ai_sku_kq" and f"{k[:-3]}_giay" in ss:
            for h in ("_kq", "_sig", "_giay"):
                ss.pop(k[:-3] + h, None)
            n += 1
    if ca_xuat and "xuat" in ss:
        ss.pop("xuat", None)
        n += 1
    try:
        st.cache_data.clear()
    except Exception:  # noqa: BLE001
        pass
    tra_ram()
    sau = ram_mb()[0]
    return {"truoc": truoc, "sau": sau, "n": n}


def _cb_don_ram() -> None:
    ss["_don_ram_kq"] = don_ram(bool(ss.get("_don_ram_xuat")))


def uoc_mb_map() -> int:
    """Ước RAM đỉnh khi map ~0,7 KB/dòng DATA SP (đo thực: 2 triệu dòng ≈ 1,4 GB) + nền 150 MB."""
    return int(150 + len(ss.get("data_sp", [])) * 0.0007)


def nut_tai(nhan: str, dung_file, **kw) -> None:
    """Nút tải: file xlsx chỉ được DỰNG KHI BẤM (data=hàm) thay vì dựng sau mỗi lần chạy lại trang — bảng vài trăm nghìn
    dòng không còn làm chậm mọi thao tác. Streamlit cũ chưa hỗ trợ -> tự dựng ngay như trước."""
    try:
        st.download_button(nhan, dung_file, **kw)
    except Exception:  # noqa: BLE001
        st.download_button(nhan, dung_file(), **kw)


def duoc_nap_nganh() -> bool:
    """Ai cũng được NẠP/THÊM cấu hình ngành (file mẫu ngành / file SKU). Mọi lần nạp ghi vào lịch sử (commit
    '[người dùng] ...'); admin xem lại + xoá/khôi phục ở Cấu hình → Cấu hình ngành hàng. Muốn siết lại: sửa 1 chỗ này."""
    return True


def bump() -> None:
    ss.ver = ss.get("ver", 0) + 1


F_JSON = ("cau_hinh", "quy_doi", "sua_sku", "sua_gt", "dx_duyet", "quy_tac_kt", "ai_cau_hinh", "ai_hoc")
DX_DIR = "shared/de_xuat"


def p_dx(user: str) -> str:
    return f"{DX_DIR}/{user}.json"


def doc_file(path: str, ghi_nho: bool = True) -> bytes | None:
    """Đọc 1 file từ kho và GHI NHỚ phiên bản đã đọc (để lúc lưu phát hiện máy khác đã sửa -> gộp, không ghi đè)."""
    data, phien_ban = tao_store().doc2(path)
    if ghi_nho:
        ss.setdefault("goc", {})[path] = (phien_ban, data)
    return data


def nap_quy_tac() -> None:
    """Đọc lại (mới nhất) các bảng quy tắc đã duyệt + quyết định duyệt."""
    tao_store().phien_ban_thu_muc("shared")  # làm mới danh sách phiên bản
    for k in ("quy_doi", "sua_sku", "sua_gt", "dx_duyet", "quy_tac_kt"):
        ss[k] = bytes_to_json(doc_file(F_SHARED[k]), {}) or {}


def nap_de_xuat(user: str) -> list:
    return bytes_to_json(doc_file(p_dx(user)), []) or []


def nap_tat_ca_de_xuat() -> None:
    S = tao_store()
    out = []
    for f in S.liet_ke(DX_DIR):
        if f.endswith(".json"):
            out += bytes_to_json(doc_file(f"{DX_DIR}/{f}"), []) or []
    ss.dx_tat_ca = out
    ss.dx_tat_ca_luc = C.bay_gio()


def nap_shared() -> None:
    tao_store().phien_ban_thu_muc("shared")
    ss.cau_hinh = bytes_to_json(doc_file(F_SHARED["cau_hinh"]), {}) or {}
    ss.ai_cau_hinh = bytes_to_json(doc_file(F_SHARED["ai_cau_hinh"]), {}) or {}
    ss.ai_hoc = bytes_to_json(doc_file(F_SHARED["ai_hoc"]), {}) or {}
    nap_quy_tac()
    if ss.get("admin"):
        try:
            nap_tat_ca_de_xuat()
        except Exception:  # noqa: BLE001
            ss.dx_tat_ca = []
    for k in ("map_tskt", "map_filter", "data_pim"):
        ss[k] = bytes_to_df(doc_file(F_SHARED[k]), COT[k])
    ss.opt = C.option_maps(ss.data_pim)
    bump()


def ap_settings(st_: dict) -> None:
    ss.sua = {tuple(k.split("\t")): v for k, v in st_.get("sua", {}).items() if k.count("\t") == 2}
    # (ngành, mã) -> đơn vị (rule cũ) ; (ngành|*, mã, "bd") -> [bước biến đổi hàng loạt]
    ss.dv = {tuple(k.split("\t")): v for k, v in st_.get("don_vi", {}).items() if k.count("\t") in (1, 2)}
    ss.rong = {k: list(v) for k, v in st_.get("rong", {}).items()}
    # web-1.4: mặc định về đúng bản desktop (chỉ map theo PROPERTYID); map theo tên thành tuỳ chọn bật tay
    # web-2.7: LUÔN về Tắt khi mở workspace (map theo tên không phải logic desktop, từng điền nhầm cột)
    ss.map_ten = "tat"
    ss.duyet_ten = list(st_.get("duyet_ten", []))  # các "ngành\tmã" người dùng đã duyệt cho phép vào file import
    ss.luc_luu = st_.get("luc_luu", "")
    ss.hoan_tac = list(st_.get("hoan_tac", []))
    ss._tt = _trang_thai_sua()
    ss.chon_nganh = dict(st_.get("chon_nganh", {}))


def nap_workspace() -> None:
    tao_store().phien_ban_thu_muc(f"users/{ss.ws}")
    for k in ("import", "data_sp", "spec"):
        ss[k] = bytes_to_df(doc_file(p_user(k)), COT[k])
        if k != "import":
            ss[k] = C.nen_df(ss[k])
    meta = bytes_to_json(doc_file(p_user("ket_qua_meta")), {}) or {}
    ss.meta = meta
    df_kq = bytes_to_df(doc_file(p_user("ket_qua")), ["cate", "sku", "model", "variant", "cate_pim", "ma", "gia_tri"])
    ss.bang = C.df_sang_bang(df_kq, meta.get("nganh", {}))
    ap_settings(bytes_to_json(doc_file(p_user("settings")), {}) or {})
    ss.lich_su = bytes_to_json(doc_file(p_user("lich_su")), []) or []
    ss.dx_rieng = nap_de_xuat(ss.ws)
    ss.ws_da_nap = ss.ws
    ss.pop("xuat", None); ss.pop("ws_mau", None)
    ss.pop("xung_dot", None)
    ss.ver_nap = ss.get("ver_nap", 0) + 1  # đổi khoá bảng chọn ngành sau mỗi lần tải lại
    bump()


def _trang_thai_sua() -> dict:
    """Ảnh chụp phần 'chỉnh tay' (sửa tay · đơn vị/biến đổi · Không/Đang cập nhật) để so sánh & hoàn tác."""
    return {"sua": {"\t".join(k): v for k, v in ss.sua.items()},
            "dv": {"\t".join(k): v for k, v in ss.dv.items() if v},
            "rong": {k: v for k, v in ss.rong.items() if v and v[0] != C.HD_GIU}}


def ghi_nhan_hoan_tac(nhan: str) -> None:
    """Gọi trước mỗi lần lưu: nếu phần chỉnh tay đổi so với lần trước -> ghi bản ĐẢO (chỉ phần khác) để hoàn tác."""
    cur, prev = _trang_thai_sua(), ss.get("_tt")
    ss._tt = cur
    if prev is None or cur == prev:
        return
    d = {}
    for ph in ("sua", "dv", "rong"):
        a, b = prev.get(ph, {}), cur.get(ph, {})
        x = {k: a.get(k) for k in set(a) | set(b) if a.get(k) != b.get(k)}
        if x:
            d[ph] = x
    if d:
        ls = list(ss.get("hoan_tac", []))
        ls.append({"luc": C.bay_gio(), "nhan": nhan, "d": d,
                   "so": sum(len(v) for v in d.values())})
        ss.hoan_tac = ls[-15:]
        ss._vua = True


def hoan_tac_lan_cuoi() -> None:
    ls = list(ss.get("hoan_tac", []))
    if not ls:
        return
    e = ls.pop()
    for ph, x in e["d"].items():
        dic = {"sua": ss.sua, "dv": ss.dv, "rong": ss.rong}[ph]
        for k, v in x.items():
            kk = tuple(k.split("\t")) if ph != "rong" else k
            if v is None:
                dic.pop(kk, None)
            else:
                dic[kk] = v
    ss.hoan_tac = ls
    ss._tt = _trang_thai_sua()
    bump()
    luu(["settings"], f"Hoàn tác: {e['nhan']}")


def hoan_tac_theo_cot(chon: list) -> None:
    """chon: [(ngành|*, mã)] -> bỏ sửa tay + đơn vị + biến đổi hàng loạt của các cột đó."""
    for c, m in chon:
        for k in [k for k in ss.sua if k[0] == c and k[2] == m]:
            ss.sua.pop(k, None)
        ss.dv.pop((c, m), None)
        ss.dv.pop((c, m, "bd"), None)
    bump()
    luu(["settings"], f"Hoàn tác {len(chon)} cột")


def _hoan_tac_cot_ui() -> None:
    nhom: dict = {}
    for (c, _s, m), _v in ss.sua.items():
        nhom.setdefault((c, m), [0, "", ""])[0] += 1
    for k, v in ss.dv.items():
        if v and len(k) in (2, 3):
            x = nhom.setdefault((k[0], k[1]), [0, "", ""])
            x[1 if len(k) == 2 else 2] = "có"
    if not nhom:
        st.caption("Không có cột nào đang được chỉnh.")
        return
    st.caption("Tick cột cần trả về đúng giá trị map từ CMS:")
    df = pd.DataFrame([{"Bỏ": False, "Ngành": c, "Cột": m, "Tên cột": ss.bang.get(c, {}).get("ten", {}).get(m, ""),
                        "Sửa tay": a, "Đơn vị": u, "Biến đổi": bd} for (c, m), (a, u, bd) in sorted(nhom.items())])
    ed = st.data_editor(df, hide_index=True, width="stretch", key=f"ed_ht_cot_{len(ss.sua)}_{len(ss.dv)}",
                        disabled=[c for c in df.columns if c != "Bỏ"])
    if st.button("Hoàn tác các cột đã tick", key="btn_ht_cot"):
        chon = [(r["Ngành"], r["Cột"]) for _, r in ed.iterrows() if r["Bỏ"]]
        if chon:
            hoan_tac_theo_cot(chon)
            ss._vua = True
            st.rerun()
        else:
            st.warning("Chưa tick cột nào.")


def thanh_hoan_tac() -> None:
    """Thanh gọn 1 dòng — CHỈ hiện ngay sau khi vừa thao tác (tách/gộp/điền đơn vị/sửa hàng loạt…)."""
    ls = list(ss.get("hoan_tac", []))
    if not ls or not ss.get("_vua"):
        return
    e = ls[-1]
    c = st.columns([5, 1.4, 1.6, 0.5], vertical_alignment="center")
    c[0].markdown(f"↩ Vừa **{e['nhan']}** · {e['so']} mục")
    if c[1].button("Hoàn tác", key="btn_hoan_tac", type="primary", width="stretch"):
        hoan_tac_lan_cuoi()
        ss._vua = False
        st.rerun()
    with c[2].popover("Theo cột", width="stretch"):
        _hoan_tac_cot_ui()
    if c[3].button("✕", key="btn_dong_ht", help="Ẩn thanh này"):
        ss._vua = False
        st.rerun()


def settings_bytes() -> bytes:
    return json_to_bytes({"sua": {"\t".join(k): v for k, v in ss.sua.items()},
                          "don_vi": {"\t".join(k): v for k, v in ss.dv.items() if v},
                          "rong": {k: v for k, v in ss.rong.items() if v and v[0] != C.HD_GIU},
                          "duyet_ten": list(ss.get("duyet_ten", [])), "hoan_tac": list(ss.get("hoan_tac", [])), "cai_dat_ver": 3, "luc_luu": C.bay_gio(),
                          "chon_nganh": ss.get("chon_nganh", {}),
                          "nguoi_luu": ss.user})


# Cách GỘP khi file bị máy khác sửa cùng lúc (xem dong_bo.py). File không có ở đây -> hỏi người dùng.
def kieu_gop(path: str) -> str | None:
    ten = path.rsplit("/", 1)[-1]
    if path.startswith(DX_DIR + "/"):
        return "ds_id"
    if path.startswith("shared/"):
        return {"cau_hinh.json": "cau_hinh", "quy_doi_filter.json": "dict", "sua_sku.json": "dict",
                "sua_gia_tri.json": "dict", "de_xuat_duyet.json": "dict", "quy_tac_kiem_tra.json": "dict", "ai_hoc.json": "dict", "map_tskt.parquet": "bang:cate,prop_id",
                "map_filter.parquet": "bang:cate,prop_id", "data_pim.parquet": "bang:"}.get(ten)
    if path.startswith("users/"):
        return {"settings.json": "settings", "lich_su.json": "ds"}.get(ten)
    return None


def _ap_ket_qua_gop(path: str, obj) -> None:
    """Đưa bản đã gộp vào bộ nhớ phiên để màn hình hiển thị đúng."""
    for k, f in F_SHARED.items():
        if f == path:
            if isinstance(obj, pd.DataFrame):
                ss[k] = obj
                if k == "data_pim":
                    ss.opt = C.option_maps(obj)
            else:
                ss[k] = obj
            return
    if path == p_user("settings"):
        ap_settings(obj)
    elif path == p_user("lich_su"):
        ss.lich_su = obj
    elif path == p_dx(ss.ws):
        ss.dx_rieng = obj


def _gop(path: str, minh: bytes | None) -> tuple:
    """-> (bytes đã gộp | None nếu không gộp được, phiên bản trên kho, mô tả)."""
    kg = kieu_gop(path)
    S = tao_store()
    S.phien_ban_thu_muc(path.rpartition("/")[0])  # làm mới
    ho, pb = S.doc2(path)
    goc = (ss.get("goc", {}).get(path) or (None, None))[1]
    if kg is None or minh is None:
        return None, pb, ""
    if kg.startswith("bang:"):
        cols = None
        for k, f in F_SHARED.items():
            if f == path:
                cols = COT[k]
        khoa = [x for x in kg[5:].split(",") if x] or None
        df, vc = DB.tron_bang(bytes_to_df(goc, cols), bytes_to_df(minh, cols), bytes_to_df(ho, cols), khoa)
        _ap_ket_qua_gop(path, df)
        return df_to_bytes(df), pb, (f"{len(vc)} dòng cả 2 cùng sửa — giữ bản của bạn" if vc else "")
    mac_dinh = [] if kg in ("ds_id", "ds") else {}
    obj, vc = DB.tron_json(kg, bytes_to_json(goc, mac_dinh), bytes_to_json(minh, mac_dinh), bytes_to_json(ho, mac_dinh))
    if kg == "settings":
        obj = {**obj, "luc_luu": C.bay_gio(), "nguoi_luu": ss.user}
    _ap_ket_qua_gop(path, obj)
    return json_to_bytes(obj), pb, (f"{len(vc)} mục cả 2 cùng sửa — giữ bản của bạn" if vc else "")


def luu(phan: list, thong_diep: str, them_file: dict | None = None, ghi_de: dict | None = None) -> bool:
    """phan: tên phần cần lưu ('import','data_sp','spec','ket_qua','settings','lich_su','dx',
    'shared:cau_hinh','shared:map_tskt','shared:map_filter','shared:data_pim',…).
    Nhiều máy cùng lúc: file máy khác vừa sửa được GỘP tự động; không gộp được (dữ liệu lô) thì hỏi."""
    files = {}
    ghi_nhan_hoan_tac(thong_diep)
    for p in phan:
        if p.startswith("shared:"):
            k = p.split(":", 1)[1]
            files[F_SHARED[k]] = json_to_bytes(ss[k]) if k in F_JSON else df_to_bytes(ss[k])
        elif p == "dx":
            files[p_dx(ss.ws)] = json_to_bytes(ss.dx_rieng)
        elif p == "ket_qua":
            files[p_user("ket_qua")] = df_to_bytes(C.bang_sang_df(ss.bang))
            files[p_user("ket_qua_meta")] = json_to_bytes(ss.meta)
        elif p == "settings":
            continue
        elif p == "lich_su":
            files[p_user("lich_su")] = json_to_bytes(ss.lich_su[-300:])
        else:
            files[p_user(p)] = df_to_bytes(ss[p])
    files[p_user("settings")] = settings_bytes()  # luôn kèm (dấu thời gian lưu)
    files.update(them_file or {})
    goc = ss.setdefault("goc", {})
    ky_vong = {p: goc[p][0] for p in files if p in goc}
    # "Ghi đè": chỉ các file đang xung đột, và chỉ khi trên kho VẪN là bản đã thấy lúc báo xung đột (không xoá
    # mất thay đổi mới hơn của máy khác); mọi file khác vẫn kiểm tra phiên bản như thường.
    ky_vong.update(ghi_de or {})
    ghi_chu, ok, msg = [], False, ""
    for _lan in range(3):  # lỗi 5xx tạm thời của GitHub: tự thử lại tối đa 3 lượt trước khi báo lỗi
        ghi_chu = []
        with st.spinner("Đang lưu…"):
            try:
                for _ in range(4):
                    ok, msg, _ghi, xd = tao_store().luu(files, f"[{ss.user}] {thong_diep}", ky_vong)
                    if ok or not xd:
                        break
                    khong_gop = []
                    for path in xd:
                        b, pb, mo_ta = _gop(path, files.get(path))
                        if b is None:
                            khong_gop.append(path)
                            continue
                        files[path], ky_vong[path] = b, pb
                        ghi_chu.append(f"{path.rsplit('/', 1)[-1]}: đã gộp với bản máy khác vừa lưu" +
                                       (f" ({mo_ta})" if mo_ta else ""))
                    if khong_gop:
                        ss.xung_dot = {"files": {p: files[p] for p in files}, "paths": khong_gop,
                                       "xd": {p: xd[p] for p in khong_gop}, "thong_diep": thong_diep}
                        msg = ("Workspace vừa được LƯU TỪ MÁY/TAB KHÁC (" + ", ".join(p.rsplit("/", 1)[-1] for p in khong_gop)
                               + ") — chọn cách xử lý ở khung đỏ trên cùng.")
                        break
            except Exception as e:  # noqa: BLE001
                ok, msg = False, f"Lỗi khi lưu: {e}"
        if ok or ss.get("xung_dot") or "lỗi 5" not in msg:
            break
        time.sleep(3 * (_lan + 1))
    if ok:
        for p, d in files.items():
            goc[p] = (git_sha(d), d)
        ss.luc_luu = C.bay_gio()
        x = ss.get("xung_dot")
        if x and set(x["paths"]) <= set(files):  # chỉ gỡ xung đột khi chính các file đó đã lưu được
            ss.pop("xung_dot", None)
        ss.chua_luu = bool(ss.get("xung_dot"))
        ss.pop("_luu_lai", None)
        st.toast(f"💾 {msg}", icon="✅")
        for g in ghi_chu:
            st.toast("🔀 " + g, icon="ℹ️")
    else:
        ss.chua_luu = True
        if ss.get("xung_dot"):
            st.rerun()  # khung xử lý xung đột vẽ 1 lần ở đầu trang (tránh trùng widget)
        ss._luu_lai = (list(phan), thong_diep)
        st.error(f"⚠️ CHƯA LƯU ĐƯỢC: {msg} — dữ liệu vẫn còn trên màn hình, bấm «🔁 Lưu lại» ở đầu trang.")
    return ok


def thanh_kho(items: list) -> None:
    """Thanh dữ liệu đang nạp: mỗi ô là NÚT — bấm để xem bảng; dữ liệu lô có nút Làm sạch."""
    sel = ss.get("kho_chon")
    cols = st.columns(len(items))
    for c, (k_, nhan, n, dv) in zip(cols, items):
        if c.button(f"{nhan}  \n**{n:,}** {dv}", key=f"kho_{k_}", width="stretch",
                    type="primary" if sel == k_ else "secondary"):
            ss.kho_chon = None if sel == k_ else k_
            st.rerun()
    if not sel:
        return
    nhan = next((x[1] for x in items if x[0] == sel), sel)
    with st.container(border=True):
        top = st.columns([6, 1])
        top[0].markdown(f"**{nhan}** — đang có trong tool")
        if top[1].button("✕ Đóng", key="kho_dong"):
            ss.kho_chon = None
            st.rerun()
        try:
            if sel in ("import", "data_sp", "spec"):
                df = ss.get(sel)
                if df is None or not len(df):
                    st.success("Trống — không có dữ liệu.")
                else:
                    st.caption(f"{len(df):,} dòng" + (" — hiện 300 dòng đầu" if len(df) > 300 else ""))
                    st.dataframe(df.head(300), hide_index=True, width="stretch", height=300)
                    ok = st.checkbox("Tôi chắc chắn muốn xoá dữ liệu này", key=f"kho_ok_{sel}")
                    if st.button(f"🧹 Làm sạch {nhan}", disabled=not ok, key=f"kho_xoa_{sel}"):
                        ss[sel] = df.iloc[0:0]
                        ss.pop("xuat", None)
                        ss.kho_chon = None
                        bump()
                        luu([sel], f"Làm sạch {nhan}")
                        st.rerun()
            elif sel == "bang":
                if not ss.get("bang"):
                    st.success("Chưa có kết quả map.")
                else:
                    st.dataframe(pd.DataFrame([{"Ngành": c_, "Tên": b_["title"], "Số SKU": len(b_["rows"])}
                                               for c_, b_ in ss.bang.items()]), hide_index=True, width="stretch",
                                 height=min(300, 60 + 35 * len(ss.bang)))
                    ok = st.checkbox("Tôi chắc chắn muốn xoá kết quả map", key="kho_ok_bang")
                    if st.button("🧹 Làm sạch kết quả map", disabled=not ok, key="kho_xoa_bang"):
                        ss.bang, ss.meta = {}, {}
                        ss.pop("xuat", None)
                        ss.kho_chon = None
                        bump()
                        luu(["ket_qua"], "Làm sạch kết quả map")
                        st.rerun()
            elif sel == "cau_hinh":
                st.dataframe(pd.DataFrame([{"Ngành": c_, "Tên": v.get("ten", ""), "Số cột": len(v.get("cot", []))}
                                           for c_, v in ss.cau_hinh.items()]), hide_index=True, width="stretch", height=300)
                st.caption("Dữ liệu dùng chung — sửa ở ⚙️ Cấu hình & mapping.")
            else:
                df = ss.get(sel)
                st.caption(f"{len(df):,} dòng — hiện 300 dòng đầu. Dữ liệu dùng chung, sửa ở ⚙️ Cấu hình & mapping.")
                st.dataframe(df.head(300), hide_index=True, width="stretch", height=300)
        except Exception as e:  # noqa: BLE001
            st.info(f"Không hiển thị được bảng này ({type(e).__name__}).")


def thanh_luu_lai() -> None:
    """Lưu thất bại (vd GitHub lỗi 500) → nút lưu lại 1 chạm, không phải làm lại thao tác."""
    x = ss.get("_luu_lai")
    if not x or ss.get("xung_dot"):
        return
    c = st.columns([5, 1.6], vertical_alignment="center")
    if ss.get("admin"):
        c[0].markdown("⚠️ **Chưa lưu được lần gần nhất** — dữ liệu vẫn còn trong phiên làm việc này.")
    else:
        c[0].markdown("⚠️ **Chưa lưu được** — bấm nút bên cạnh để thử lại.")
    if c[1].button("🔁 Lưu lại", type="primary", key="btn_luu_lai"):
        luu(*x)
        st.rerun()


def hien_xung_dot() -> None:
    """Khung xử lý khi cùng 1 workspace bị lưu từ 2 nơi (dữ liệu lô không gộp tự động được)."""
    x = ss.get("xung_dot")
    if not x:
        return
    st.markdown("<div class='canh'><b>⚠️ Workspace này vừa được lưu từ MÁY hoặc TAB KHÁC</b> (cùng tài khoản). "
                "File bị trùng: " + ", ".join(p.rsplit("/", 1)[-1] for p in x["paths"]) +
                ".<br>Chọn 1 trong 2: lấy bản trên kho (bỏ thay đổi chưa lưu ở đây) hoặc ghi đè bằng bản ở đây.</div>",
                unsafe_allow_html=True)
    c = st.columns([1, 1, 3])
    if c[0].button("⬇️ Lấy bản trên kho", type="primary", key="xd_tai"):
        nap_shared()
        nap_workspace()
        st.rerun()
    if c[1].button("⬆️ Ghi đè bằng bản ở đây", key="xd_ghi"):
        ss.pop("xung_dot", None)
        luu([], x["thong_diep"] + " (ghi đè)", them_file={p: x["files"][p] for p in x["paths"]}, ghi_de=x["xd"])
        st.rerun()


# ============================================================================
# QUY TẮC SỬA CMS ĐANG ÁP (đã duyệt chung + đề xuất chờ duyệt của chủ workspace)
# ============================================================================
def qt() -> tuple:
    """-> (sua_sku, sua_gt, quy_doi, số đề xuất riêng đang áp)."""
    return C.quy_tac_hieu_luc((ss.sua_sku, ss.sua_gt, ss.quy_doi), ss.get("dx_rieng", []), ss.dx_duyet)


def dem_dx(ds_dx: list) -> Counter:
    return Counter(C.trang_thai_dx(d, ss.dx_duyet) for d in ds_dx)


# ============================================================================
# KIỂM TRA (tính lại khi có thay đổi)
# ============================================================================
def kq() -> dict:
    if ss.get("kq_ver") != ss.ver or "kq" not in ss:
        dv_luu = dict(ss.dv)
        ss.kq = C.tinh_kiem_tra(ss.bang, ss["import"], ss.spec, ss.opt, ss.sua, ss.dv, ss.rong, dv_luu)
        ss.kq_ver = ss.ver
    return ss.kq


def ttm() -> pd.DataFrame:
    """Kiểm tra thông minh (không AI) — tính lại khi dữ liệu/sửa đổi thay đổi."""
    if ss.get("ttm_ver") != ss.ver or "ttm" not in ss:
        ss.ttm = C.kiem_tra_thong_minh(ss.bang, ss.sua, ss.dv, ss.rong)
        ss.ttm_ver = ss.ver
    return ss.ttm


def ds() -> dict:
    """Đối soát CMS -> kết quả (tính lại khi dữ liệu thay đổi)."""
    # đối soát chỉ phụ thuộc DATA SP + kết quả map + mapping (không phụ thuộc sửa tay/đơn vị) -> không tính lại
    # sau mỗi lần sửa ô (lô lớn đỡ chờ hàng chục giây)
    khoa = (id(ss.data_sp), id(ss.bang), id(ss.map_tskt), id(ss.map_filter), id(ss.cau_hinh), id(ss.spec),
            len(ss.quy_doi), len(ss.sua_sku), len(ss.sua_gt), len(ss.get("dx_rieng", [])), ss.get("ver_map", 0))
    if ss.get("ds_ver") != khoa or "ds" not in ss:
        sk, gt, qd, _ = qt()
        ss.ds = C.doi_soat_map(ss.data_sp, ss["import"], ss.bang, ss.cau_hinh, ss.map_tskt, ss.map_filter, ss.opt,
                               qd, ss.spec, sk, gt)
        ss.ds_ver = khoa
    return ss.ds


def nq(chi_cate: list | None = None) -> dict:
    """QC ngầm NHẤT QUÁN ngành: cấu hình · DATA SP · mapping TSKT/FILTER · DATA PIM (chỉ đọc, có nhớ đệm)."""
    _, _, qd, _ = qt()
    khoa = (ss.get("ver", 0), id(ss.data_sp), len(ss.data_sp), id(ss.map_tskt), len(ss.map_tskt), id(ss.map_filter),
            len(ss.map_filter), id(ss.data_pim), len(ss.cau_hinh),
            sum(len(v.get("cot", [])) for v in ss.cau_hinh.values()), len(qd or {}),
            tuple(chi_cate) if chi_cate else None)
    if ss.get("nq_ver") != khoa or "nq_kq" not in ss:
        ss.nq_kq = C.kiem_tra_nhat_quan(ss.cau_hinh, ss.data_sp, ss.map_tskt, ss.map_filter, ss.opt, qd, chi_cate)
        ss.nq_ver = khoa
    return ss.nq_kq


def hien_nhat_quan(d: dict, key: str, gon: bool = False) -> None:
    """Bảng tóm tắt từng ngành + danh sách vấn đề (dùng cho QC ngầm trước/sau map và khi xem file)."""
    L, N = d["loi"], d["nganh"]
    if not len(N):
        st.caption("Chưa có ngành nào để kiểm (cần DATA SP hoặc cấu hình ngành).")
        return
    n_cao = int((L["Mức"] == "CAO").sum()) if len(L) else 0
    n_tb = int((L["Mức"] == "TB").sum()) if len(L) else 0
    if n_cao:
        st.error(f"⛔ {n_cao} lỗi nhất quán (sót cấu hình / sót mapping / mapping trỏ sai cột / FILTER thiếu option) "
                 "— SKU hoặc giá trị sẽ bị bỏ khi map.")
    elif n_tb:
        st.warning(f"⚠️ {n_tb} điểm cần xem (cột chưa có mapping, thuộc tính CMS chưa map, FILTER không khớp option).")
    else:
        st.success(f"✔ {len(N)} ngành nhất quán: cấu hình · DATA SP · mapping TSKT/FILTER · DATA PIM khớp nhau.")
    st.dataframe(N, hide_index=True, width="stretch", height=min(320, 42 + 35 * len(N)), column_config={
        "% cột có mapping": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
    if len(L):
        hien = L[L["Mức"] != "TT"] if gon else L
        hien = hien.assign(Mức=hien["Mức"].map(MUC_ICON).fillna(hien["Mức"]))
        with st.expander(f"Chi tiết {len(hien)} vấn đề", expanded=not gon and bool(n_cao)):
            st.dataframe(hien, hide_index=True, width="stretch", height=min(420, 42 + 35 * min(len(hien), 11)),
                         column_config={"Chi tiết": st.column_config.TextColumn(width="large"),
                                        "Số": st.column_config.NumberColumn(format="%d")})
            nut_tai("📊 Tải báo cáo nhất quán (.xlsx)",
                               lambda: C.xlsx_nhieu_sheet({"NGÀNH": N, "VẤN ĐỀ": L}),
                               file_name=f"NHAT_QUAN_NGANH_{C.bay_gio()[:10]}.xlsx", key=f"nq_dl_{key}")


def tab_nhat_quan() -> None:
    st.caption("QC ngầm từng ngành của lô: **cấu hình cột** ↔ **DATA SP** ↔ **mapping TSKT / FILTER** ↔ **DATA PIM**. "
               "Phát hiện ngành bị sót cấu hình/mapping, mapping trỏ sai cột, cột luôn trống, thuộc tính CMS chưa map, "
               "cột FILTER không có option. Chỉ kiểm, **không sửa** dữ liệu.")
    tat_ca = st.checkbox("Kiểm cả các ngành không có trong lô này (mọi ngành đã cấu hình)", key="nq_tat_ca")
    cates = sorted(set(ss.cau_hinh) | set(ss.data_sp.CATEGORYID if len(ss.data_sp) else [])) if tat_ca else None
    hien_nhat_quan(nq(cates), "tab")


def kc() -> dict:
    """Kiểm chứng SKU <-> DATA SP (chỉ đọc). Tính lại khi map lại hoặc sửa/đơn vị thay đổi."""
    _, _, qd, _ = qt()
    khoa = (ss.get("ver", 0), ss.get("ver_map", 0), id(ss.data_sp), id(ss.bang), id(ss.map_tskt), id(ss.map_filter),
            len(qd or {}))
    if ss.get("kc_ver") != khoa or "kc_kq" not in ss:
        ss.kc_kq = C.kiem_chung_sku(ss.bang, ss.data_sp, ss["import"], ss.map_tskt, ss.map_filter, ss.opt, qd,
                                    ss.meta.get("goc_cms") or {},
                                    lambda c_, s_, m_, v_: C.bien_doi_o(c_, s_, m_, v_, ss.sua, ss.dv, ss.rong)[0])
        ss.kc_ver = khoa
    return ss.kc_kq


def _doi_chieu_model_sku() -> None:
    """Đối chiếu từng MÃ MODEL: mỗi SKU của model ↔ chính SKU đó trong DATA SP. Ô có thông tin thì hiện, không có thì để trống."""
    if bang_trong() or not len(ss.get("data_sp", [])):
        return
    c = st.columns([2, 2, 3])
    cate = c[0].selectbox("Ngành hàng", list(ss.bang), format_func=lambda x: ss.bang[x]["title"], key="dc_cate")
    b = ss.bang[cate]
    models = sorted({r["model"] for r in b["rows"] if r["model"]})
    mo = c[1].selectbox("Mã model", ["Tất cả"] + models, key=f"dc_model_{cate}")
    che = c[2].radio("Hiện", ["Kết quả map", "DATA SP gốc", "Chỉ ô khác nhau"], horizontal=True, key="dc_che")
    tm: dict = {}
    for cc, pid, ma in ss.map_tskt[["cate", "prop_id", "ma"]].itertuples(index=False):
        if cc == cate:
            tm.setdefault(pid, ma)
    rows = [r for r in b["rows"] if mo == "Tất cả" or r["model"] == mo]
    skus = {r["sku"] for r in rows}
    sp = ss.data_sp[ss.data_sp.PRODUCTCODE.isin(skus)]
    raw: dict = {}
    for sku, pid, val in sp[["PRODUCTCODE", "PROPERTYID", "PROPVALUE"]].itertuples(index=False):
        ma = tm.get(pid)
        if ma and val:
            lst = raw.setdefault((sku, ma), [])
            if val not in lst:
                lst.append(val)
    cot = [m for m in C.cot_tt(b) if not C.la_cot_filter(m)]
    out = []
    for r in rows:
        d = {"Model": r["model"], "SKU": r["sku"]}
        for m in cot:
            vm = r["vals"].get(m, "")
            vr = C.SEP_TSKT.join(raw.get((r["sku"], m), []))
            if che == "Kết quả map":
                d[m] = vm
            elif che == "DATA SP gốc":
                d[m] = vr
            else:
                d[m] = f"{vm}  ≠  {vr}" if (vm or vr) and C.chuan_hoa_key(vm) != C.chuan_hoa_key(vr) else ""
        out.append(d)
    df = pd.DataFrame(out)
    giu = [m for m in cot if m in df.columns and (df[m].astype(str).str.strip() != "").any()]
    df = df[["Model", "SKU"] + giu]
    if che == "Chỉ ô khác nhau":
        df = df[(df[giu].astype(str) != "").any(axis=1)] if giu else df.iloc[0:0]
    if not giu:
        st.success("✔ Không có ô nào để hiện." if che == "Chỉ ô khác nhau" else "Chưa có ô nào có giá trị.")
        return
    ten = b.get("ten", {})
    st.caption(f"{len(df):,} SKU · {len(giu)} cột có thông tin" + (" — hiện 300 SKU đầu, chọn model để xem hết" if len(df) > 300 else ""))
    st.dataframe(df.head(300), hide_index=True, width="stretch", height=min(480, 60 + 35 * min(len(df), 12)),
                 column_config={m: st.column_config.TextColumn(ten.get(m) or m, help=m) for m in giu})


def tab_kiem_chung() -> None:
    d = kc()
    L, t = d["loi"], d["tk"]
    nghiem = int((L["Mức"] == "CAO").sum()) if len(L) else 0
    c = st.columns(4)
    c[0].metric("SKU đã kiểm", f"{t.get('sku', 0):,}")
    c[1].metric("Ô có giá trị", f"{t.get('o', 0):,}")
    c[2].metric("Giá trị truy đúng nguồn", f"{t.get('khop', 0):,}")
    c[3].metric("Lỗi nghiêm trọng", f"{nghiem:,}")
    st.caption("Đi ngược từ **từng ô kết quả** về **DATA SP của chính SKU đó**, theo đúng mapping PROPERTYID như lúc map: "
               "giá trị TSKT phải có trong PROPVALUE của SKU, mã FILTER phải suy ra được từ CMS của SKU. "
               "Chỉ kiểm, **không sửa** dữ liệu.")
    _doi_chieu_model_sku()
    if not len(L):
        st.success(f"✔ 100% khớp — {t.get('o', 0):,} ô của {t.get('sku', 0):,} SKU đều lấy đúng dữ liệu từ đúng SKU "
                   "trong DATA SP. Không có ký tự ẩn.")
        return
    if nghiem:
        st.error(f"⚠️ {nghiem:,} lỗi nghiêm trọng (lệch SKU / FILTER không có nguồn / lệch IMPORT / lệch ngành). "
                 "Không xuất file khi chưa xử lý.")
    loai = st.multiselect("Loại", sorted(L["Loại"].unique()), key="kc_loai",
                          default=[x for x in L["Loại"].unique() if C.KC_MUC.get(x) in ("CAO", "TB")])
    v = L[L["Loại"].isin(loai)] if loai else L
    v = v.assign(Mức=v["Mức"].map(MUC_ICON).fillna(v["Mức"]))
    st.dataframe(v.head(3000), hide_index=True, height=min(480, 60 + 35 * min(len(v), 12)),
                 column_config={"Nguồn / ghi chú": st.column_config.TextColumn(width="large")})
    nut_tai("📊 Tải báo cáo kiểm chứng (.xlsx)", lambda: C.xlsx_nhieu_sheet({"KIỂM CHỨNG": L}),
                       file_name=f"KIEM_CHUNG_SKU_{C.bay_gio()[:10]}.xlsx", key="kc_dl")


def dht() -> dict:
    """Độ hoàn thiện + vi phạm quy tắc kiểm tra (tính lại khi sửa / đổi quy tắc)."""
    khoa = (ss.ver, json_to_bytes(ss.get("quy_tac_kt", {})))
    if ss.get("dht_ver") != khoa or "dht" not in ss:
        ss.dht = C.do_hoan_thien(ss.bang, ss.cau_hinh, ss.get("quy_tac_kt", {}), ss.sua, ss.dv, ss.rong)
        ss.dht_ver = khoa
    return ss.dht


def bang_chon_nganh(key: str) -> list:
    """Bảng CHỌN NGÀNH HÀNG (như sheet của desktop/file mẫu) -> danh sách mã ngành được tick."""
    goc = {x["MÃ NH"]: x for x in ss.meta.get("chon", [])}
    rows = []
    d = dht()["nganh"].set_index("Ngành") if len(dht()["nganh"]) else None
    for c, b in ss.bang.items():
        x = goc.get(c, {})
        rows.append({"CHỌN": bool(ss.get("chon_nganh", {}).get(c, True)), "MÃ NH": c,
                     "TÊN NGÀNH HÀNG": x.get("TÊN NGÀNH HÀNG") or b.get("ten_nh", ""), "SỐ SKU": len(b["rows"]),
                     "NGUỒN CATE": x.get("NGUỒN CATE", ""), "TAB TSKT / CẤU HÌNH": x.get("TAB TSKT / CẤU HÌNH", ""),
                     "% ĐẦY ĐỦ": float(d.loc[c, "% đầy đủ TB"]) if d is not None and c in d.index else 0.0,
                     "GHI CHÚ": x.get("GHI CHÚ", "")})
    if not rows:
        return []
    ed = st.data_editor(pd.DataFrame(rows), hide_index=True, key=key, height=min(400, 40 + 35 * len(rows)),
                        disabled=[k for k in rows[0] if k != "CHỌN"],
                        column_config={"CHỌN": st.column_config.CheckboxColumn(width="small"),
                                       "% ĐẦY ĐỦ": st.column_config.ProgressColumn(min_value=0, max_value=100,
                                                                                   format="%.0f%%")})
    moi = {r["MÃ NH"]: bool(r["CHỌN"]) for r in ed.to_dict("records")}
    if moi != {c: bool(ss.get("chon_nganh", {}).get(c, True)) for c in moi}:
        ss.chon_nganh = {**ss.get("chon_nganh", {}), **moi}
        luu(["settings"], "Chọn ngành hàng")
    return [c for c, v in moi.items() if v]


_AI_PROV = ("groq", "gemini", "openrouter")


def _sach(v) -> str:
    return str(v or "").strip().strip('"').strip("'").strip()


def _doan_prov(key: str) -> str:
    """Đoán nhà cung cấp theo đầu key (tránh khai sai AI_PROVIDER)."""
    if key.startswith("gsk_"):
        return "groq"
    if key.startswith("AIza"):
        return "gemini"
    if key.startswith("sk-or-"):
        return "openrouter"
    return ""


def ai_tu_secrets() -> dict:
    """Đọc cấu hình AI từ Secrets — nhập 1 lần, cả nhóm dùng, đăng xuất không mất. Chấp nhận các kiểu khai:
    AI_PROVIDER/AI_API_KEY/AI_MODEL · GROQ_API_KEY/GEMINI_API_KEY/OPENROUTER_API_KEY · bảng [ai] provider/api_key/model
    (chữ hoa/thường đều được)."""
    bang = {}
    try:
        for ten in ("ai", "AI"):
            if ten in st.secrets and hasattr(st.secrets[ten], "get"):
                bang = {str(k).lower(): v for k, v in dict(st.secrets[ten]).items()}
                break
    except Exception:  # noqa: BLE001 - chưa có secrets.toml
        bang = {}

    # Hay gặp: dán AI_API_KEY bên dưới 1 bảng [users.xxx] → TOML xếp nó vào bảng đó. Quét cả các bảng con.
    long = {}
    try:
        def _quet(d, sau=0):
            for k_, v_ in dict(d).items():
                if hasattr(v_, "items") and sau < 3:
                    _quet(v_, sau + 1)
                elif sau and isinstance(v_, str) and str(k_).upper() in (
                        "AI_PROVIDER", "AI_API_KEY", "AI_MODEL", "AI_BASE_URL", "GROQ_API_KEY", "GEMINI_API_KEY",
                        "OPENROUTER_API_KEY"):
                    long.setdefault(str(k_).upper(), v_)
        _quet(st.secrets)
    except Exception:  # noqa: BLE001
        pass

    def lay(*ten):
        for t in ten:
            v = _sach(sec(t, "") or sec(t.lower(), "") or long.get(t.upper(), ""))
            if v:
                return v
        return ""
    prov = (lay("AI_PROVIDER") or _sach(bang.get("provider"))).lower()
    key = lay("AI_API_KEY") or _sach(bang.get("api_key") or bang.get("key"))
    if not key:
        for p in ([prov] if prov in _AI_PROV else []) + [p for p in _AI_PROV if p != prov]:
            key = lay(f"{p.upper()}_API_KEY") or _sach(bang.get(f"{p}_api_key"))
            if key:
                prov = prov if prov in _AI_PROV and lay(f"{prov.upper()}_API_KEY") == key else p
                break
    prov = _doan_prov(key) or (prov if prov in _AI_PROV else "groq")
    return {"provider": prov, "api_key": key, "model": lay("AI_MODEL") or _sach(bang.get("model")),
            "base_url": lay("AI_BASE_URL") or _sach(bang.get("base_url"))}


def tao_ai() -> AIH.AI:
    # Ưu tiên: key dán tạm trong phiên → Secrets (cố định cho cả nhóm) → bản lưu cũ trên kho (nếu có)
    scf = ai_tu_secrets()
    acf = scf if scf.get("api_key") else (ss.get("ai_cau_hinh") or {})
    if ss.get("ai_key_tam"):
        k_ = ss.ai_key_tam
        return AIH.AI(_doan_prov(k_) or ss.get("ai_prov_tam") or "groq", k_, ss.get("ai_model_tam") or "", "")
    prov = acf.get("provider") or "groq"
    return AIH.AI(prov, acf.get("api_key") or "", acf.get("model") or "", acf.get("base_url") or "")


_MAU_SECRETS_AI = 'AI_PROVIDER = "groq"\nAI_API_KEY = "gsk_...dán key vào đây..."\n# AI_MODEL = "openai/gpt-oss-120b"   # tuỳ chọn'


def huong_dan_secrets_ai() -> None:
    scf = ai_tu_secrets()
    if scf.get("api_key"):
        k = scf["api_key"]
        st.success(f"🔑 Đang dùng key AI trong **Secrets**: {scf['provider']} · `{k[:6]}…{k[-4:]}` ({len(k)} ký tự) · "
                   f"model `{scf.get('model') or '(mặc định)'}` — cả nhóm dùng chung, đăng xuất không mất.")
    else:
        st.warning("Chưa có key AI trong Secrets. Admin vào **Streamlit Cloud → app → ⋮ → Settings → Secrets**, "
                   "dán 2 dòng dưới (ngoài mọi bảng [..], đặt ở đầu file) → **Save changes** → đợi app tự khởi động lại. "
                   "Nhập 1 lần, cả nhóm dùng luôn.")
        st.code(_MAU_SECRETS_AI, language="toml")


def ten_sp(sku: str) -> str:
    d = ss.data_sp
    x = d.PRODUCTNAME[d.PRODUCTCODE == sku] if len(d) else []
    return x.iloc[0] if len(x) else ""


def dat_sua(cate: str, sku: str, ma: str, gia_tri: str, goc: str) -> None:
    if C.la_cot_filter(ma):
        gia_tri = C.ma_filter_tu_gia_tri(ma, gia_tri, ss.opt["option_map"])
    key = (cate, sku, ma)
    if C.chuan_hoa_key(gia_tri) == C.chuan_hoa_key(goc):
        ss.sua.pop(key, None)
    else:
        ss.sua[key] = C.chuan_hoa_key(gia_tri)


def bang_trong() -> bool:
    return not ss.get("bang")


# ============================================================================
# THANH BÊN
# ============================================================================
TRANG = ["🏁 Làm nhanh", "📥 Nạp dữ liệu lô", "🚀 Map & kiểm tra", "📤 Xuất file import", "📮 Đề xuất sửa CMS",
         "⚙️ Cấu hình & mapping", "🧰 Tra cứu", "📘 Hướng dẫn", "👥 Quản trị"]


def so_cho_duyet() -> int:
    return sum(C.trang_thai_dx(d, ss.dx_duyet) == C.DX_CHO for d in ss.get("dx_tat_ca", []))


def thanh_ben() -> None:
    with st.sidebar:
        if ss.admin:
            st.markdown(f"### PIM Tool\n👤 **{ss.ten}** (`{ss.user}`) · 🛡️ admin")
        else:
            st.markdown(f"### PIM Tool\n👤 **{ss.ten}**")
        if st.toggle("🔎 Chữ to hơn (dễ đọc)", key="chu_to", help="Phóng to toàn bộ chữ và nút trong app."):
            st.markdown("<style>html{font-size:18.5px !important}.stApp p,.stApp label,.stApp li{font-size:1.08rem}"
                        "div[data-testid=stTab] p{font-size:1.05rem !important}</style>", unsafe_allow_html=True)
        st.caption("💡 Chữ dài được thu gọn — rê chuột vào để đọc đủ.")
        if ss.admin:
            tk = list(ds_tai_khoan())
            chon = st.selectbox("Workspace đang xem", tk, index=tk.index(ss.ws) if ss.ws in tk else 0,
                                help="Admin xem/sửa được workspace của mọi tài khoản.")
            if chon != ss.ws:
                ss.ws = chon
                nap_workspace()
                st.rerun()
        if ss.ws != ss.user:
            st.warning(f"Đang xem workspace của **{ss.ws}** — mọi thay đổi ghi vào workspace này.")
        if ss.admin and so_cho_duyet():
            st.warning(f"📥 **{so_cho_duyet()} đề xuất sửa CMS chờ duyệt** → 📋 Quản lý dữ liệu → 📮 Đề xuất")
        dm = dem_dx(ss.get("dx_rieng", []))
        if dm:
            st.caption(f"📮 Đề xuất của {'bạn' if ss.ws == ss.user else ss.ws}: ⏳ {dm.get(C.DX_CHO, 0)} chờ · "
                       f"✅ {dm.get(C.DX_DUYET, 0)} duyệt · ❌ {dm.get(C.DX_TU_CHOI, 0)} từ chối")
        st.divider()
        S = tao_store()
        if ss.admin:
            if S.backend == "local":
                st.warning("Chưa cấu hình GITHUB_TOKEN → đang lưu TẠM trên máy chủ (mất khi app khởi động lại).")
            st.caption(f"Lưu trữ: {S.mo_ta}")
            st.caption(f"Lần lưu gần nhất: {ss.get('luc_luu') or '—'}")
        if ss.get("chua_luu"):
            if ss.admin:
                st.error("Có thay đổi CHƯA lưu.")
            else:
                st.warning("⚠️ Nhớ lưu trước khi thoát.")
            if st.button("💾 Lưu lại ngay", width="stretch"):
                luu(["import", "data_sp", "spec", "ket_qua", "settings"], "Lưu thủ công")
                st.rerun()
        c1, c2 = st.columns(2)
        if c1.button("🔄 Tải lại", width="stretch", help="Đọc lại dữ liệu mới nhất từ kho"):
            nap_shared()
            nap_workspace()
            st.rerun()
        if c2.button("🚪 Đăng xuất", width="stretch"):
            for k in list(ss.keys()):
                del ss[k]
            st.rerun()
        if ss.admin:
            the_lien_he()
        # Admin: xem nhật ký lỗi
        if ss.admin and ss.get("_nhat_ky_loi"):
            st.divider()
            with st.expander(f"🔴 Nhật ký lỗi ({len(ss._nhat_ky_loi)})"):
                for e in reversed(ss._nhat_ky_loi[-10:]):
                    st.markdown(f"**{e['luc']}** — `{e['loai']}` tại {e['noi']}")
                    st.caption(e["chi_tiet"])
                    with st.expander("Chi tiết đầy đủ", expanded=False):
                        st.code(e["day_du"], language="python")
                if st.button("🧹 Xóa nhật ký lỗi", key="xoa_loi"):
                    ss._nhat_ky_loi = []
                    st.rerun()


def _chan_nhanh() -> list:
    """Lỗi chặn cho nút Xuất nhanh: chỉ lỗi lệch SKU / không nguồn (kiểm chứng) và thiếu khoá import."""
    try:
        kcl = kc()["loi"]
        n = int((kcl["Mức"] == "CAO").sum()) if len(kcl) else 0
        s_ = kq()["stat"]
    except Exception:  # noqa: BLE001
        return []
    x = [(n, "ô LỆCH SKU / không có nguồn (tab ✅ Kiểm chứng)"), (s_.get("thieu_model", 0), "SKU thiếu model_code"),
         (s_.get("thieu_cate", 0), "SKU thiếu category_code"), (s_.get("filter_chu", 0), "ô FILTER sai mã")]
    out = [f"• {a:,} {b}" for a, b in x if a]
    if out:
        st.markdown("<div class='canh'><b>⛔ LỖI CẦN XỬ LÝ:</b><br>" + "<br>".join(out) + "</div>",
                    unsafe_allow_html=True)
    return out


def khu_kiem_nhanh(key: str, chon: list) -> None:
    """Trước khi tải file: bảng kết quả kiểm tra (bằng code, đọc đúng dữ liệu đã map + DATA SP) + lưới giá trị SẼ XUẤT sửa nhanh."""
    if not chon or bang_trong():
        st.info("Chưa có ngành nào để kiểm tra.")
        return
    k = kq()
    q = qc_tong_hop(k)
    if q:
        st.markdown("**Kết quả kiểm tra**")
        st.dataframe(pd.DataFrame([{"Mức": MUC_ICON[a], "Vùng": b, "Số lượng": n_, "Cách xử lý": g}
                                   for a, b, n_, y, g, act, kind in q]),
                     hide_index=True, width="stretch", height=min(300, 60 + 35 * len(q)))
    else:
        st.success("✔ Không có mục nào cần xử lý — có thể tải file.")
    st.markdown("**Giá trị sẽ xuất** (sửa trực tiếp rồi bấm Lưu — file import dùng đúng các giá trị này)")
    c = st.columns([2, 2])
    cate = c[0].selectbox("Ngành hàng", chon, format_func=lambda x: ss.bang[x]["title"], key=f"{key}_kn_cate")
    b = ss.bang[cate]
    models = sorted({r["model"] for r in b["rows"] if r["model"]})
    mo = c[1].selectbox("Mã model", ["Tất cả"] + models, key=f"{key}_kn_model_{cate}")
    cot = [m for m in C.cot_tt(b) if not C.la_cot_filter(m)]
    rows = [r for r in b["rows"] if mo == "Tất cả" or r["model"] == mo]
    out = []
    for r in rows[:300]:
        d = {"Model": r["model"], "SKU": r["sku"]}
        for m in cot:
            d[m], _ = C.bien_doi_o(cate, r["sku"], m, r["vals"].get(m, ""), ss.sua, ss.dv, ss.rong)
        out.append(d)
    df = pd.DataFrame(out)
    giu = [m for m in cot if m in df.columns and (df[m].astype(str).str.strip() != "").any()]
    if not giu:
        st.info("Ngành này chưa có ô nào có giá trị.")
        return
    df = df[["Model", "SKU"] + giu]
    st.caption(f"{len(df):,} SKU · {len(giu)} cột có thông tin" + (" — hiện 300 SKU đầu, chọn model để xem hết" if len(rows) > 300 else ""))
    ten = b.get("ten", {})
    ed = st.data_editor(df, hide_index=True, width="stretch", height=min(460, 60 + 35 * min(len(df), 11)),
                        key=f"{key}_kn_ed_{cate}_{mo}_{ss.ver}", disabled=["Model", "SKU"],
                        column_config={m: st.column_config.TextColumn(ten.get(m) or m, help=m) for m in giu})
    if st.button("💾 Lưu các ô đã sửa", type="primary", key=f"{key}_kn_luu"):
        goc = {r["sku"]: r for r in b["rows"]}
        n = 0
        for i in range(len(df)):
            for m in giu:
                moi = C.chuan_hoa_key(ed.at[i, m] or "")
                cu_ = C.chuan_hoa_key(df.at[i, m] or "")
                if moi != cu_:
                    dat_sua(cate, df.at[i, "SKU"], m, moi, goc[df.at[i, "SKU"]]["vals"].get(m, ""))
                    n += 1
        if n:
            bump()
            luu(["settings"], f"Kiểm tra nhanh: sửa {n} ô")
            st.rerun()
        st.info("Chưa có ô nào thay đổi.")


def xuat_gon(key: str, chon: list | None = None, canh: list | None = None, can_xn: bool | None = None) -> None:
    """Khối xuất dùng chung (trang Làm nhanh + trang Xuất)."""
    if chon is None:
        chon = [c for c in ss.bang if ss.get("chon_nganh", {}).get(c, True)]
        st.caption("Ngành xuất: " + (", ".join(ss.bang[c]["title"] for c in chon) or "(chưa chọn — xem trang 📤)"))
    c = st.columns([1.2, 1.4, 2])
    giu_sku = c[0].checkbox("Giữ cột sku", value=False, key=f"{key}_sku",
                            help="Mặc định bỏ cột sku như bản desktop (file import theo model_code/variant_code).")
    bo_trong = c[1].checkbox("Tách SKU không có giá trị ra file xin data", value=True, key=f"{key}_trong",
                             help="SKU không có thông số nào sẽ KHÔNG nằm trong file import (bản desktop vẫn xuất dòng "
                                  "trống). Danh sách nằm trong file xin data CMS.")
    with c[2]:
        nut_xin_data(f"{key}_xd")
    canh_bao_o_ten()
    kn = bool(ss.get(f"{key}_kn"))
    if st.button("✕ Đóng kiểm tra nhanh" if kn else "🔎 Kiểm tra nhanh trước khi tải", width="stretch",
                 type="secondary" if kn else "primary", key=f"xem_{key}_kn"):
        ss[f"{key}_kn"] = not kn
        st.rerun()
    if kn:
        with st.container(border=True):
            khu_kiem_nhanh(key, chon)
    xn = True
    # Chỉ bắt tick khi có lỗi CHẶN (can_xn); lưu ý thường không cần tick -> bớt 1 lượt bấm.
    if (can_xn if can_xn is not None else bool(canh)):
        xn = st.checkbox("Tôi đã xem lỗi ở trên và vẫn xuất file", value=False, key=f"{key}_xn")
    if st.button("📤 Tạo file import", type="primary", disabled=not xn or not chon, key=f"{key}_tao"):
        with st.spinner("Đang tạo file…"):
            _uoc_x = int(150 + sum(len(b_.get("rows", [])) * len(b_.get("attr", [])) for b_ in ss.bang.values()) * 0.0004)
            with viec_nang("Tạo file import", _uoc_x) as _ok_x:
                x = (C.xuat_file_import(ss.bang, ss["import"], ss.sua, ss.dv, ss.rong, bo_cot_sku=not giu_sku,
                                        chi_cate=chon, bo_dong_trong=bo_trong, bo_o=o_ten_bi_chan())
                     if _ok_x else None)
        if x is not None:
            ss.xuat = x
            ss.lich_su.append({"Lúc": C.bay_gio(), "Người xuất": ss.user, "Workspace": ss.ws,
                               "File": ", ".join(f[0] for f in x["files"]), "Số dòng": sum(f[2] for f in x["files"]),
                               "Sửa tay": x.get("so_o_sua", 0), "Thêm đơn vị": x.get("so_o_dv", 0),
                               "Biến đổi": x.get("so_o_bd", 0), "Không/Đang cập nhật": x.get("so_o_rong", 0),
                               "Tách xin data": x.get("bo_trong", 0), "Chặn ô theo tên": x.get("bo_o_ten", 0), "Cảnh báo": " | ".join(canh or [])})
            luu(["lich_su"], f"Xuất {len(x['files'])} file import")
            # AI tự học: dùng chính lô vừa xuất (coi như đã được xác nhận) làm mẫu cho lần sau
            if HIEN_AI:
                try:
                    cap_nhat_ai_hoc_tu_bang(chi_cate=list(chon))
                except Exception:  # noqa: BLE001
                    pass
            ss.xuat_ver = ss.get("ver", 0)
    x = ss.get("xuat")
    if x and ss.get("xuat_ver") != ss.get("ver", 0):
        st.warning("Dữ liệu đã thay đổi sau lần tạo file trước — bấm **📤 Tạo file import** lại để có file mới nhất.")
        x = None
    if x:
        st.markdown("#### Tải file (mở được ngay, không cần giải nén)")
        for t, d, n in x["files"]:
            st.download_button(f"⬇️ {t} ({n:,} dòng)", d, file_name=t, key=f"{key}_dl_{t}",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("🗜️ Tải tất cả (.zip)", x["zip"], file_name=f"PIM_IMPORT_{x['stamp']}.zip",
                           mime="application/zip", key=f"{key}_zip")
        st.caption(f"Đã áp {x.get('so_o_sua', 0)} ô sửa tay · {x.get('so_o_dv', 0)} ô thêm đơn vị · "
                   f"{x.get('so_o_bd', 0)} ô biến đổi hàng loạt · {x.get('so_o_rong', 0)} ô Không/Đang cập nhật"
                   + (f" · tách {x['bo_trong']} SKU không có giá trị sang file xin data" if x.get("bo_trong") else "")
                   + (f" · bỏ {x['bo_dong']} dòng không có trong IMPORT" if x.get("bo_dong") else ""))


# ============================================================================
# TRANG: NẠP DỮ LIỆU LÔ
# ============================================================================
VUNG_LO = [  # (khoá, tên vùng, có chế độ ghi đè/nối tiếp)
    ("import", "📋 IMPORT — model · SKU · mã biến thể · category", True),
    ("data_sp", "📦 DATA SP — dữ liệu CMS của sản phẩm", True),
    ("spec", "🧾 SPEC PIM tạm — để đối chiếu", True),
    ("don_vi", "📏 Đơn vị đã đặt theo cột", False),
    ("chon_nganh", "☑️ CHỌN NGÀNH HÀNG", False)]
VUNG_CHUNG = [  # dùng chung, chỉ admin
    ("cau_hinh", "🏷️ Cấu hình ngành hàng (danh sách cột TSKT)"),
    ("map_tskt", "🧬 Mapping TSKT"),
    ("map_filter", "🧮 Mapping FILTER"),
    ("data_pim", "🗂️ DATA PIM (option FILTER)")]


def _so_vung(w: dict, k: str) -> int:
    v = w.get(k)
    return 0 if v is None else len(v)


def nhap_mau(w: dict, chon_lo: dict, chon_chung: list, ten_file: str) -> None:
    """Nạp ĐÚNG các vùng đã chọn từ file mẫu; vùng không chọn giữ nguyên 100%.
    chon_lo: {khoá vùng lô: "ghi_de" | "noi_tiep"} · chon_chung: các khoá dùng chung (admin) gộp theo ngành."""
    phan, bao = [], []
    for k, mode in chon_lo.items():
        n = _so_vung(w, k)
        if not n:
            continue
        if k == "import":
            moi = w[k]
            if mode == "noi_tiep" and len(ss["import"]):
                ss["import"] = pd.concat([ss["import"][~ss["import"].sku.isin(set(moi.sku))], moi], ignore_index=True)
            else:
                ss["import"] = moi
        elif k in ("data_sp", "spec"):
            moi = w[k]
            if mode == "noi_tiep" and len(ss[k]):
                cu = ss[k].astype(object)
                if k == "spec":
                    cu = cu[~cu.sku.isin(set(moi.sku))]
                ss[k] = C.nen_df(pd.concat([cu, moi], ignore_index=True).drop_duplicates(ignore_index=True))
            else:
                ss[k] = C.nen_df(moi)
        elif k == "don_vi":
            ss.dv.update({tuple(kk.split("\t")): v for kk, v in w["don_vi"].items()})
        elif k == "chon_nganh":
            ss.chon_nganh = w["chon_nganh"]
        phan.append(k)
        bao.append(f"{k}: {n:,} ({'nối tiếp' if mode == 'noi_tiep' else 'ghi đè'})" if mode else f"{k}: {n:,}")
    if chon_chung and duoc_sua_chung():
        nap_shared()  # lấy bản mới nhất trước khi gộp (nhiều người cùng dùng)
        cu = {k: ss[k] for k in ("cau_hinh", "map_tskt", "map_filter", "data_pim")}
        out, tk = C.gop_chung_theo_nganh(cu, {k: (w.get(k) if k in chon_chung else None) for k in cu})
        for k, v in out.items():
            ss[k] = v
        ss.opt = C.option_maps(ss.data_pim)
        phan += [f"shared:{k}" for k in tk]
        bao += [f"{k}: {v}" for k, v in tk.items()]
    if not phan:
        st.info("Không có vùng nào được nạp (vùng chọn không có dữ liệu trong file).")
        return
    bump()
    if luu(phan + ["settings"], f"Nạp file mẫu {ten_file}: {', '.join(phan)}"):
        st.success("✔ Đã nạp: " + " · ".join(bao) + " — các vùng còn lại giữ nguyên.")


def nap_lai_toan_bo(w: dict, kem_lo: bool, ten_file: str) -> None:
    """Thay TOÀN BỘ data gốc dùng chung (cấu hình, mapping TSKT, mapping FILTER, DATA PIM) bằng nội dung file."""
    if not duoc_sua_chung():
        st.error("Chỉ admin được nạp lại data gốc.")
        return
    nap_shared()
    phan, bao, giu = [], [], []
    for k, ten in (("cau_hinh", "Cấu hình ngành"), ("map_tskt", "Mapping TSKT"), ("map_filter", "Mapping FILTER"),
                   ("data_pim", "DATA PIM")):
        v = w.get(k)
        if v is None or not len(v):
            giu.append(ten)
            continue
        ss[k] = v
        phan.append(f"shared:{k}")
        bao.append(f"{ten}: {len(v):,}" + (" ngành" if k == "cau_hinh" else " dòng"))
    ss.opt = C.option_maps(ss.data_pim)
    if kem_lo:
        for k in ("import", "data_sp", "spec"):
            if _so_vung(w, k):
                ss[k] = C.nen_df(w[k]) if k != "import" else w[k]
                phan.append(k)
                bao.append(f"{k}: {len(w[k]):,}")
        if w.get("don_vi"):
            ss.dv.update({tuple(kk.split("\t")): v for kk, v in w["don_vi"].items()})
        if w.get("chon_nganh"):
            ss.chon_nganh = w["chon_nganh"]
    if not phan:
        st.warning("File không có phần data gốc nào để nạp.")
        return
    bump()
    if luu(phan + ["settings"], f"Nạp lại toàn bộ data gốc từ {ten_file}"):
        st.success("✔ Đã thay toàn bộ: " + " · ".join(bao) + (f" — giữ nguyên: {', '.join(giu)}" if giu else ""))


def khu_nap_mau() -> None:
    st.markdown("<div class='buoc'><b>📦 Cách nhanh nhất:</b> nạp NGUYÊN file theo mẫu "
                "(<i>du_lieu_pim.xlsx</i> / <i>TEST HÀNG LOẠT IMPORT THÔNG SỐ</i>: IMPORT, DATA SP, DATA PIM, CẤU HÌNH "
                "CATEGORY, MAPPING TSKT MOI, MAPPING FILTER MOI, CHỌN NGÀNH HÀNG, tab TSKT…). Tool tự nhận từng sheet, "
                "kể cả tiêu đề mẫu mới (MÃ THUỘC TÍNH TSKT, MÃ MASTER, MÃ HỌ) và tab TSKT có sẵn.</div>",
                unsafe_allow_html=True)
    f = st.file_uploader("File workspace theo mẫu (.xlsx)", type=["xlsx", "xlsm"], key="up_mau")
    if f is not None and ss.get("mau_ten") != (f.name, f.size):
        t = time.time()
        with st.spinner("Đang đọc file mẫu…"):
            ss.mau = C.doc_workspace_cu(f.getvalue(), f.name)
        ss.mau_ten, ss.mau_giay = (f.name, f.size), time.time() - t
    w = ss.get("mau") if f is not None else None
    if not w:
        return
    dem = {"IMPORT (SKU)": len(w.get("import", [])), "DATA SP (dòng)": len(w.get("data_sp", [])),
           "DATA PIM (option)": len(w.get("data_pim", [])), "Cấu hình (ngành)": len(w.get("cau_hinh", {})),
           "Mapping TSKT": len(w.get("map_tskt", [])), "Mapping FILTER": len(w.get("map_filter", [])),
           "Tab TSKT": len(w.get("tab_tskt", {}))}
    cc = st.columns(len(dem))
    for c, (k, v) in zip(cc, dem.items()):
        c.metric(k, f"{v:,}".replace(",", "."))
    st.caption(f"Đọc trong {ss.get('mau_giay', 0):.1f} giây · sheet: {', '.join(w.get('sheets', []))}")
    for cb in w.get("canh_bao", []):
        st.warning(cb)
    if w.get("tab_tskt"):
        st.caption("Tab TSKT dùng làm danh sách cột (như desktop): " + " · ".join(
            f"{t['tab']} → ngành {c} ({len(t['cot'])} cột)" for c, t in w["tab_tskt"].items()))
    nganh_lo = set(w.get("data_sp", pd.DataFrame(columns=["CATEGORYID"])).CATEGORYID) - {""}
    thieu = [c for c in sorted(nganh_lo) if c not in ss.cau_hinh and c not in w.get("cau_hinh", {})]
    if thieu:
        st.warning(f"Ngành chưa có cấu hình ở kho lẫn trong file: {', '.join(thieu)} — SKU các ngành này sẽ bị bỏ qua.")
    with st.container(border=True):
        st.markdown("##### 🔄 Nạp lại TOÀN BỘ data gốc")
        st.caption("Thay **toàn bộ** Cấu hình ngành hàng · Mapping TSKT (master) · Mapping FILTER (master) · DATA PIM bằng "
                   "nội dung file này (không gộp theo ngành — ngành nào không có trong file sẽ mất). Dùng khi muốn "
                   "làm lại từ đầu theo file gốc. Phần chưa có trong file thì giữ nguyên. Chỉ admin.")
        kem_lo = st.checkbox("Nạp kèm dữ liệu lô trong file (IMPORT · DATA SP · SPEC · đơn vị · chọn ngành) — ghi đè",
                             value=False, key=f"goc_lo_{abs(hash(ss.get('mau_ten')))}")
        ok_goc = st.checkbox("Tôi hiểu: dữ liệu dùng chung hiện tại sẽ bị thay bằng file này", value=False,
                             key=f"goc_ok_{abs(hash(ss.get('mau_ten')))}", disabled=not duoc_sua_chung())
        if st.button("🔄 Nạp lại toàn bộ", type="primary", disabled=not (ok_goc and duoc_sua_chung()),
                     key=f"goc_nut_{abs(hash(ss.get('mau_ten')))}"):
            nap_lai_toan_bo(w, kem_lo, f.name)
            if kem_lo and st.session_state.get("tu_map_sau_nap", True):
                chay_map_ui()
    _oq = st.container(border=True)
    _oq.markdown("**🧭 QC ngầm nội dung file này** (cấu hình · DATA SP · mapping · DATA PIM) — trước khi nạp")
    if _oq.checkbox("Xem QC ngầm file này", value=False, key=f"nq_file_mau_{abs(hash(ss.get('mau_ten')))}"):
      with _oq:
        _ch = w.get("cau_hinh") or ss.cau_hinh
        _mt = w["map_tskt"] if w.get("map_tskt") is not None and len(w["map_tskt"]) else ss.map_tskt
        _mf = w["map_filter"] if w.get("map_filter") is not None and len(w["map_filter"]) else ss.map_filter
        _dp = w["data_pim"] if w.get("data_pim") is not None and len(w["data_pim"]) else ss.data_pim
        _sp = w["data_sp"] if w.get("data_sp") is not None and len(w["data_sp"]) else None
        hien_nhat_quan(C.kiem_tra_nhat_quan(_ch, _sp, _mt, _mf, C.option_maps(_dp)), "file_mau", gon=True)
    st.markdown("##### Chọn vùng cần nạp — vùng không tick giữ nguyên, không bị đụng tới")
    ver = abs(hash(ss.get("mau_ten")))
    ds_lo = [(k, t, m) for k, t, m in VUNG_LO if _so_vung(w, k)]
    ds_chung = [(k, t) for k, t in VUNG_CHUNG if _so_vung(w, k)]

    def dat_hang_loat(lo: bool, chung: bool) -> None:
        for k, _, _ in ds_lo:
            st.session_state[f"v_{k}_{ver}"] = lo
        for k, _ in ds_chung:
            st.session_state[f"v_{k}_{ver}"] = chung and duoc_sua_chung()
    b1, b2, b3, b4 = st.columns(4)
    b1.button("✅ Tất cả vùng", on_click=dat_hang_loat, args=(True, True), key=f"b_all_{ver}")
    b2.button("Chỉ vùng LÔ", on_click=dat_hang_loat, args=(True, False), key=f"b_lo_{ver}")
    b3.button("Chỉ vùng DÙNG CHUNG", on_click=dat_hang_loat, args=(False, True), key=f"b_chung_{ver}")
    b4.button("Bỏ chọn hết", on_click=dat_hang_loat, args=(False, False), key=f"b_none_{ver}")
    chon_lo, chon_chung = {}, []
    st.markdown(f"**Vùng LÔ** → workspace `{ss.ws}`")
    for k, t, co_mode in ds_lo:
        c = st.columns([3.2, 1.2, 2])
        st.session_state.setdefault(f"v_{k}_{ver}", True)
        on = c[0].checkbox(t, key=f"v_{k}_{ver}")
        c[1].caption(f"{_so_vung(w, k):,} dòng" if k != "chon_nganh" else f"{_so_vung(w, k)} ngành")
        mode = None
        if co_mode:
            mode = "noi_tiep" if c[2].radio("Cách ghi", ["Ghi đè", "Nối tiếp"], horizontal=True, key=f"m_{k}_{ver}",
                                            label_visibility="collapsed", disabled=not on,
                                            help="Ghi đè = thay hết vùng này. Nối tiếp = giữ dữ liệu cũ, thêm/cập nhật phần trong file."
                                            ) == "Nối tiếp" else "ghi_de"
        if on:
            chon_lo[k] = mode
    st.markdown("**Vùng DÙNG CHUNG** (cả nhóm dùng — gộp theo ngành, ngành khác giữ nguyên)")
    if not duoc_sua_chung():
        st.caption("Chỉ admin cập nhật dữ liệu dùng chung.")
    for k, t in ds_chung:
        c = st.columns([3.2, 1.2, 2])
        st.session_state.setdefault(f"v_{k}_{ver}", duoc_sua_chung())
        on = c[0].checkbox(t, key=f"v_{k}_{ver}", disabled=not duoc_sua_chung())
        c[1].caption(f"{_so_vung(w, k):,} " + ("ngành" if k == "cau_hinh" else "dòng"))
        if on and duoc_sua_chung():
            chon_chung.append(k)
    n_chon = len(chon_lo) + len(chon_chung)
    if st.button(f"✔ Nạp {n_chon} vùng đã chọn" if n_chon else "✔ Nạp vào tool", type="primary",
                 disabled=not n_chon, key=f"nap_{ver}"):
        nhap_mau(w, chon_lo, chon_chung, f.name)
        if chon_lo and st.session_state.get("tu_map_sau_nap", True):
            chay_map_ui()


# ============================================================================
# NẠP NHANH 1 CỤC (tự nhận loại file, tự lọc) + FILE XIN DATA CMS
# ============================================================================
def khu_nap_nhanh(key: str = "nn") -> None:
    st.markdown("<div class='buoc'><b>⚡ Nạp nhanh 1 cục:</b> kéo thả 1 hoặc nhiều file BẤT KỲ — workspace theo mẫu, "
                "file CMS export, danh sách SKU / file export PIM (model, SKU, biến thể, có hoặc không có cột TSKT) — "
                "hoặc tick ✍️ Tự điền tay để dán cột model / SKU / ID từ Excel. Tool <b>tự nhận loại file, tự nhận cột</b> (tên tiếng Việt/Anh, có/không "
                "dấu, thiếu tiêu đề thì đoán theo nội dung), <b>tự lọc</b> DATA SP theo SKU cần làm.</div>",
                unsafe_allow_html=True)
    for m in ss.pop("flash", []) or []:
        st.success(m)
    for m in ss.pop("flash_err", []) or []:
        st.error(m)
    lan = ss.get(f"{key}_lan", 0)  # đổi khoá sau mỗi lần nạp -> ô chọn file / ô dán tự trống (không nạp lặp)
    fs = st.file_uploader("File (.xlsx / .xlsm / .csv) — chọn được nhiều file", type=["xlsx", "xlsm", "xls", "csv"],
                          accept_multiple_files=True, key=f"{key}_f{lan}")
    # ---- TỰ ĐIỀN TAY model / SKU / mã biến thể (không cần file) ----
    tay_txt = ""
    if st.checkbox("✍️ Tự điền tay model / SKU / mã biến thể (không cần file)", key=f"{key}_tay_on{lan}",
                   help="Dán hàng loạt từng cột (Model / SKU / Mã biến thể) hoặc gõ vào bảng. Chỉ cần SKU; "
                        "model / biến thể / ngành để trống được."):
        cate_mac = st.text_input("Mã ngành mặc định (tuỳ chọn)", key=f"{key}_tay_cate{lan}", placeholder="vd: 2062",
                                 help="Áp cho các dòng để trống category_code.")
        _che = st.radio("Cách nhập", ["📋 Dán theo cột (hàng loạt)", "⌨️ Bảng gõ tay"], horizontal=True,
                        key=f"{key}_tay_che{lan}", label_visibility="collapsed")
        _dong, _loi_tay = [], ""
        _HD = "model_code\tsku\tvariant_code\tcategory_code\tPRODUCTID"
        if _che.startswith("📋"):
            st.caption("Copy 1 cột từ Excel rồi dán vào ô tương ứng — **mỗi dòng 1 giá trị, thứ tự dòng khớp nhau** "
                       "(dòng thứ n của ô này ↔ dòng thứ n của các ô kia). Model / Biến thể chỉ có "
                       "**1 giá trị** thì tool áp cho tất cả dòng. Mã ngành điền ở ô **Mã ngành mặc định** phía trên (áp cho mọi dòng). Mỗi dòng cần **SKU hoặc ID CMS** (SP chưa có SKU "
                       "thì nhập **ID CMS = PRODUCTID** trong file CMS export; có cả hai cũng được).")
            _c3 = st.columns(4)
            _cot = {}
            for _col, _nhan, _k in zip(_c3, ("Model (model_code)", "SKU", "Mã biến thể (variant_code)",
                                             "ID CMS (PRODUCTID)"),
                                       ("model_code", "sku", "variant_code", "pid")):
                with _col:
                    _cot[_k] = st.text_area(_nhan, height=170, key=f"{key}_tay_{_k}{lan}",
                                            placeholder="mỗi dòng 1 giá trị\n(dán cả cột từ Excel)")
            _rows, _ghi, _loi_tay = C.ghep_cot_dan(_cot, cate_mac)
            if _loi_tay:
                st.error(_loi_tay)
            else:
                for _g_ in _ghi:
                    (st.warning if _g_.startswith("⚠️") else st.caption)(_g_)
                _dong = ["\t".join(_r_) for _r_ in _rows]
        else:
            _cau = {c_: st.column_config.TextColumn(c_) for c_ in ("model_code", "sku", "variant_code", "category_code")}
            _cau["PRODUCTID"] = st.column_config.TextColumn("ID CMS (PRODUCTID)")
            _tay = st.data_editor(pd.DataFrame([["", "", "", "", ""]] * 4,
                                               columns=["model_code", "sku", "variant_code", "category_code",
                                                        "PRODUCTID"]),
                                  num_rows="dynamic", hide_index=True, width="stretch", column_config=_cau,
                                  key=f"{key}_tay{lan}")
            for _r in _tay.fillna("").astype(str).itertuples(index=False):
                if _r.sku.strip() or _r.PRODUCTID.strip():
                    _dong.append("\t".join([_r.model_code.strip(), _r.sku.strip(), _r.variant_code.strip(),
                                            (_r.category_code.strip() or cate_mac.strip()), _r.PRODUCTID.strip()]))
        if _dong:
            tay_txt = _HD + "\n" + "\n".join(_dong)
            st.caption(f"✍️ Đang có {len(_dong):,} dòng điền tay — bấm **Nạp vào tool** ở dưới để đưa vào IMPORT.")
        elif not _loi_tay:
            st.caption("Điền ít nhất **SKU hoặc ID CMS** ở 1 dòng thì mới nạp được.")
    sig = tuple((f.name, f.size) for f in fs or []) + ((hash(tay_txt),) if tay_txt else ())
    if not sig:
        return
    if ss.get(f"{key}_sig") != sig:
        kq, t = [], time.time()
        with st.spinner("Đang đọc và nhận diện…"):
            for f in fs or []:
                try:
                    kq.append((f.name, C.nhan_dien_file(f.getvalue(), f.name)))
                except Exception as e:  # noqa: BLE001
                    kq.append((f.name, {"loai": None, "loi": f"Không đọc được: {e}"}))
            if tay_txt:
                r = C.doc_mot_cuc(C.doc_text_dan(tay_txt))
                _nd = tay_txt.count("\n")
                if len(r["import"]) < _nd:
                    r["ghi_chu"].append(f"Nhận {len(r['import']):,}/{_nd:,} dòng — dòng trùng SKU/ID hoặc có ký tự không "
                                        f"hợp lệ bị bỏ.")
                kq.append(("(điền tay)", {"loai": "sku" if len(r["import"]) else None, **r,
                                          "loi": None if len(r["import"]) else " ".join(r["ghi_chu"])}))
        ss[f"{key}_sig"], ss[f"{key}_kq"], ss[f"{key}_giay"] = sig, kq, time.time() - t
    kq = ss[f"{key}_kq"]
    tom = []
    for ten, r in kq:
        if r.get("loai") == "mau":
            nd = (f"IMPORT {len(r.get('import', [])):,} SKU · DATA SP {len(r.get('data_sp', [])):,} dòng · "
                  f"{len(r.get('cau_hinh', {}))} ngành cấu hình · mapping {len(r.get('map_tskt', [])):,}+"
                  f"{len(r.get('map_filter', [])):,}")
        elif r.get("loai") == "cms":
            d = r.get("data_sp")
            nd = f"{len(d):,} dòng · {d.PRODUCTCODE.nunique():,} SKU" if d is not None and len(d) else ""
            if d is not None and len(d):
                _n_id = d.PRODUCTCODE.astype(str).str.startswith(C.ID_TIEN_TO).sum()
                if _n_id:
                    nd += f" · gồm {d.loc[d.PRODUCTCODE.astype(str).str.startswith(C.ID_TIEN_TO), 'PRODUCTID'].nunique():,} SP chưa có code (khớp theo ID CMS)"
        elif r.get("loai") == "nganh":
            nd = " · ".join(f"{v['ten']} ({c}): {len(v['cot'])} cột" for c, v in r["cau_hinh"].items())
        elif r.get("loai") == "sku":
            nd = (f"{len(r['import']):,} SKU · {(r['import'].model_code != '').sum():,} có model · "
                  f"{(r['import'].variant_code != '').sum():,} có biến thể · {len(r['spec']):,} ô TSKT/FILTER"
                  + (f" · {r['so_chi_id']:,} dòng chỉ có ID CMS" if r.get("so_chi_id") else ""))
        else:
            nd = ""
        tom.append({"File": ten, "Nhận là": C.LOAI_FILE.get(r.get("loai"), "❌ không nhận ra"), "Nội dung": nd,
                    "Ghi chú": r.get("loi") or " ".join(r.get("ghi_chu", []) if isinstance(r.get("ghi_chu"), list)
                                                       else []) or " ".join(r.get("canh_bao", []))})
    st.dataframe(pd.DataFrame(tom), hide_index=True, width="stretch")
    st.caption(f"Đọc trong {ss.get(f'{key}_giay', 0):.1f} giây.")
    ds_sku = [r for _, r in kq if r.get("loai") == "sku"]
    if ds_sku:
        with st.expander("👀 Xem trước SKU nhận được (kiểm tra cột đã nhận đúng chưa)", expanded=len(ds_sku) == 1):
            for r in ds_sku:
                st.caption("Cột nhận: " + " · ".join(f"{k} ← “{v}”" for k, v in r.get("nhan_cot", {}).items()))
                st.dataframe(r["import"].head(8), hide_index=True)
    co_mau = any(r.get("loai") == "mau" for _, r in kq)
    co_sku = bool(ds_sku) or any(len(r.get("import", [])) for _, r in kq if r.get("loai") == "mau")
    co_cms = any(r.get("loai") == "cms" for _, r in kq) or any(len(r.get("data_sp", [])) for _, r in kq
                                                                 if r.get("loai") == "mau")
    c = st.columns(4)
    ghi_imp = c[0].radio("IMPORT", ["Thay mới", "Nối thêm"], key=f"{key}_gi", horizontal=True, disabled=not co_sku)
    ghi_sp = c[1].radio("DATA SP", ["Thay mới", "Nối thêm"], key=f"{key}_gs", horizontal=True, disabled=not co_cms)
    loc = c[2].checkbox("Tự lọc DATA SP: chỉ giữ SKU có trong IMPORT", value=True, key=f"{key}_loc")
    tu_map = c[3].checkbox("Map + QC luôn sau khi nạp", value=True, key=f"{key}_map")
    lay_chung = False
    ds_ng = [(ten, r) for ten, r in kq if r.get("loai") == "nganh"]
    lay_ng = False
    if ds_ng:
        moi_ng = {c: v for _, r in ds_ng for c, v in r["cau_hinh"].items()}
        for c, v in moi_ng.items():
            cu = ss.cau_hinh.get(c)
            if cu:
                them = [x for x in v["cot"] if x not in cu.get("cot", [])]
                bo = [x for x in cu.get("cot", []) if x not in v["cot"]]
                st.markdown(f"🏷️ **{v['ten']} ({c})** · {len(v['cot'])} cột — thêm {len(them)} · bỏ {len(bo)}")
            else:
                st.markdown(f"🏷️ **{v['ten']} ({c})** · ngành mới · {len(v['cot'])} cột")
        lay_ng = st.checkbox("Cập nhật Cấu hình ngành hàng từ file mẫu ngành (dùng chung)", value=duoc_nap_nganh(),
                             disabled=not duoc_nap_nganh(), key=f"{key}_ng")
    # ---- GỢI Ý cấu hình ngành từ file SKU (file export PIM có cột TSKT) ----
    ds_gy_cfg = []  # [(ten_file, cate_id_gy, [(ma,tv),...])]
    for ten_file, r in kq:
        if r.get("loai") != "sku":
            continue
        gy = r.get("goi_y_cfg") or {}
        cot_gy = gy.get("cot") or []
        if not cot_gy:
            continue
        ds_gy_cfg.append((ten_file, gy.get("cate_id", ""), cot_gy))
    cat_them = []  # [(cate_id, cate_ten, [(ma,tv),...])]
    if ds_gy_cfg:
        # KHÔNG dùng st.expander ở đây — khu_nap_nhanh có thể đang nằm trong một expander (vd Trang 🏁 Làm nhanh),
        # Streamlit không cho expander lồng nhau -> crash im lặng -> file có vẻ "không nhận".
        with st.container(border=True):
            st.markdown(f"**➕ Thêm cấu hình ngành từ file SKU** "
                        f"({sum(len(c) for _, _, c in ds_gy_cfg):,} cột gợi ý)")
            st.caption("Xác nhận mã + tên ngành. Ngành đã có chỉ bổ sung cột thiếu. Không muốn thêm thì bỏ tick «Áp dụng».")
            st.caption("Lưu vào Cấu hình ngành dùng chung; admin xem lại lịch sử và xoá được.")
            for i, (ten_file, cate_gy, cot_gy) in enumerate(ds_gy_cfg):
                st.markdown(f"**📄 {ten_file}** — {len(cot_gy)} cột thuộc tính")
                cols = st.columns([1.2, 2.5, 1])
                cid = cols[0].text_input("Mã ngành", value=cate_gy, key=f"{key}_cfg_cid_{i}",
                                         help="Số CATEGORYID của ngành (vd: 1988, 9218)")
                cu_ten = (ss.cau_hinh.get(C.chuan_hoa_id(cid), {}) or {}).get("ten", "") if cid else ""
                if not cu_ten:
                    cu_ten = os.path.splitext(ten_file)[0]  # gợi ý tên ngành từ tên file
                cten = cols[1].text_input("Tên ngành", value=cu_ten, key=f"{key}_cfg_cten_{i}",
                                          placeholder="vd: Xe đạp tập thể dục")
                ok = cols[2].checkbox("Áp dụng", value=bool(cid), key=f"{key}_cfg_ok_{i}_{cid}",
                                      disabled=not cid)
                if not cid:
                    st.caption("⚠️ File chưa có category_code → nhập Mã ngành (số CATEGORYID) để nạp.")
                xem = st.checkbox(f"Xem {len(cot_gy)} cột sẽ thêm", key=f"{key}_cfg_xem_{i}")
                if xem:
                    st.dataframe(pd.DataFrame(cot_gy, columns=["Mã cột", "Tên tiếng Việt"]),
                                 hide_index=True, height=min(300, 40 + 30 * min(len(cot_gy), 10)))
                if ok and cid:
                    cat_them.append((C.chuan_hoa_id(cid), (cten or "").strip(), cot_gy))
    # ---- QC ngầm cho ngành SẮP thêm/cập nhật: kiểm trước khi bấm Nạp (chỉ đọc, chưa lưu gì)
    _moi: dict = {}
    if ds_ng and duoc_nap_nganh():
        for c_, v_ in moi_ng.items():
            _moi[c_] = ("thay", v_)
    for cid_, cten_, cot_gy_ in cat_them:
        _moi[cid_] = ("them", {"ten": cten_, "cot": [m for m, _ in cot_gy_], "ten_cot": {m: t for m, t in cot_gy_ if t}})
    if _moi:
        ch_tam = {c_: dict(v_, cot=list(v_.get("cot", [])), ten_cot=dict(v_.get("ten_cot", {})))
                  for c_, v_ in ss.cau_hinh.items()}
        for c_, (kieu, v_) in _moi.items():
            if kieu == "thay" or c_ not in ch_tam:
                ch_tam[c_] = {"ten": v_.get("ten", ""), "cot": list(v_["cot"]), "ten_cot": dict(v_.get("ten_cot", {}))}
            else:
                o_ = ch_tam[c_]
                o_["cot"] += [m for m in v_["cot"] if m not in o_["cot"]]
                o_["ten"] = v_.get("ten") or o_.get("ten", "")
        sp_ds = [ss.data_sp] + [r_["data_sp"] for _, r_ in kq if r_.get("loai") == "cms" and r_.get("data_sp") is not None]
        sp_ds += [r_["data_sp"] for _, r_ in kq if r_.get("loai") == "mau" and r_.get("data_sp") is not None]
        sp_tam = pd.concat([x.astype(object) for x in sp_ds if x is not None and len(x)], ignore_index=True) \
            if any(x is not None and len(x) for x in sp_ds) else ss.data_sp
        _, _, qd_, _ = qt()
        with st.container(border=True):
            st.markdown(f"**🧭 QC ngầm cho {len(_moi)} ngành sắp thêm/cập nhật** — đối chiếu cấu hình mới với DATA SP · "
                        "mapping TSKT/FILTER · DATA PIM hiện có (chưa lưu gì)")
            hien_nhat_quan(C.kiem_tra_nhat_quan(ch_tam, sp_tam, ss.map_tskt, ss.map_filter, ss.opt, qd_, list(_moi)),
                           f"{key}_ngmoi", gon=True)
    if co_mau:
        lay_chung = st.checkbox("File mẫu: cập nhật mapping / cấu hình / DATA PIM dùng chung (gộp theo ngành)",
                                value=False, disabled=not duoc_sua_chung(), key=f"{key}_chung",
                                help="Chỉ admin. Ngành có trong file thay mapping + cấu hình của ngành đó; ngành khác giữ.")
    if not st.button("✔ Nạp vào tool", type="primary", key=f"{key}_ok"):
        return
    imp, spec, sp = [], [], []
    da_ng = []
    if ds_ng and lay_ng and duoc_nap_nganh():
        nap_shared()
        ch = dict(ss.cau_hinh)
        ch.update(moi_ng)
        ss.cau_hinh = ch
        luu(["shared:cau_hinh"], "Nạp cấu hình ngành từ file mẫu: " + ", ".join(
            f"{v['ten']} ({c}) {len(v['cot'])} cột" for c, v in moi_ng.items()))
        da_ng = [f"{v['ten']} ({c}) {len(v['cot'])} cột" for c, v in moi_ng.items()]
    # Áp các gợi ý "thêm cấu hình từ file SKU" (chỉ thêm cột còn thiếu, giữ cột cũ)
    if cat_them and duoc_nap_nganh():
        nap_shared()
        log_nganh = []
        for cid, cten, cot_gy in cat_them:
            moi_tao = cid not in ss.cau_hinh
            o = ss.cau_hinh.setdefault(cid, {"ten": "", "cot": [], "ten_cot": {}})
            # Người thường: ngành ĐÃ có tên thì giữ tên cũ (tránh ghi đè nhầm). Admin: được đổi tên (như trước).
            if cten and (not o.get("ten") or duoc_sua_chung()):
                o["ten"] = cten
            them_n = 0
            for ma, tv in cot_gy:
                if ma not in o["cot"]:
                    o["cot"].append(ma)
                    them_n += 1
                if tv and ma not in o.get("ten_cot", {}):
                    o.setdefault("ten_cot", {})[ma] = tv
            da_ng.append(f"{(o.get('ten') or cten or cid)} ({cid}) +{them_n} cột")
            log_nganh.append(f"{(o.get('ten') or cten or cid)} ({cid}) {'MỚI' if moi_tao else 'bổ sung'} +{them_n} cột")
        luu(["shared:cau_hinh"], "Thêm cấu hình ngành từ file SKU: " + "; ".join(log_nganh))
    for ten, r in kq:
        if r.get("loai") == "mau":
            if lay_chung and duoc_sua_chung():
                nap_shared()
                cu = {k: ss[k] for k in ("cau_hinh", "map_tskt", "map_filter", "data_pim")}
                out, _ = C.gop_chung_theo_nganh(cu, {k: r.get(k) for k in cu})
                for k, v in out.items():
                    ss[k] = v
                ss.opt = C.option_maps(ss.data_pim)
                luu([f"shared:{k}" for k in out], f"Nạp dùng chung từ {ten}")
            if len(r.get("import", [])):
                imp.append(r["import"])
            if len(r.get("data_sp", [])):
                sp.append(r["data_sp"])
            if len(r.get("spec", [])):
                spec.append(r["spec"])
            if r.get("don_vi"):
                ss.dv.update({tuple(kk.split("\t")): v for kk, v in r["don_vi"].items()})
            if r.get("chon_nganh"):
                ss.chon_nganh = r["chon_nganh"]
        elif r.get("loai") == "sku":
            imp.append(r["import"])
            if len(r["spec"]):
                spec.append(r["spec"])
        elif r.get("loai") == "cms" and r.get("data_sp") is not None:
            sp.append(r["data_sp"])
    phan, bao = [], []
    if imp:
        moi = pd.concat(imp, ignore_index=True).drop_duplicates("sku", keep="last")
        if ghi_imp == "Nối thêm":
            moi = pd.concat([ss["import"][~ss["import"].sku.isin(set(moi.sku))], moi], ignore_index=True)
        ss["import"] = moi.reset_index(drop=True)
        phan.append("import")
        bao.append(f"IMPORT {len(ss['import']):,} SKU")
    if spec:
        moi = pd.concat(spec, ignore_index=True)
        cu = ss.spec[~ss.spec.sku.isin(set(moi.sku))].astype(object) if len(ss.spec) else ss.spec
        ss.spec = C.nen_df(pd.concat([cu, moi], ignore_index=True))
        phan.append("spec")
        bao.append(f"spec PIM cũ {moi.sku.nunique():,} SKU")
    if sp:
        moi = pd.concat(sp, ignore_index=True)
        cu = ss.data_sp.astype(object) if ghi_sp == "Nối thêm" and len(ss.data_sp) else ss.data_sp.iloc[0:0]
        ss.data_sp = pd.concat([cu, moi], ignore_index=True).drop_duplicates(ignore_index=True)
        phan.append("data_sp")
    # Dòng nhập theo ID CMS: khai cả SKU lẫn ID -> đổi mã tạm ID_<PRODUCTID> trong DATA SP sang SKU thật
    _ghep = {}
    for _t, r in kq:
        if r.get("loai") == "sku" and r.get("ghep_id"):
            _ghep.update(r["ghep_id"])
    if _ghep:
        ss["ghep_id"] = {**ss.get("ghep_id", {}), **_ghep}
    if len(ss.data_sp) and (ss.get("ghep_id") or (len(ss["import"]) and ss["import"].sku.astype(str)
                                                  .str.startswith(C.ID_TIEN_TO).any())):
        _sp_g, _n_g = C.gan_id_data_sp(ss.data_sp, ss.get("ghep_id", {}), set(ss["import"].sku))
        if _n_g:
            ss.data_sp = _sp_g
            bao.append(f"đã gắn {_n_g:,} dòng DATA SP theo ID CMS vào SKU")
            if "data_sp" not in phan:
                phan.append("data_sp")
    if loc and len(ss["import"]) and len(ss.data_sp):
        _data_sp_truoc = ss.data_sp.copy()
        _loc_kq, tk = C.loc_data_sp(ss.data_sp.astype(object), ss["import"])
        if tk["giu"] == 0 and tk["bo"] > 0:
            # Catastrophic: 0 SKU khớp giữa DATA SP và IMPORT -> KHÔNG LỌC, giữ nguyên DATA SP và báo RÕ
            ss.data_sp = _data_sp_truoc
            _sku_dsp = set(ss.data_sp.PRODUCTCODE.unique()) if len(ss.data_sp) else set()
            _sku_imp = set(ss["import"].sku.unique()) if len(ss["import"]) else set()
            _vd_dsp = list(_sku_dsp)[:3]
            _vd_imp = list(_sku_imp)[:3]
            _msg = (f"⚠️ 2 FILE KHÔNG KHỚP NHAU — 0 SKU trùng. "
                    f"IMPORT ({len(_sku_imp):,} SKU, vd: {', '.join(_vd_imp)}) vs "
                    f"DATA SP ({len(_sku_dsp):,} SKU, vd: {', '.join(_vd_dsp)}). "
                    f"Hai file xuất từ 2 lô / 2 nhóm sản phẩm KHÁC nhau. "
                    f"Cần: file CMS export của ĐÚNG các SKU trong IMPORT. "
                    f"Đã GIỮ NGUYÊN DATA SP ({len(ss.data_sp):,} dòng) để bạn upload file IMPORT đúng.")
            ss.setdefault("flash_err", []).append(_msg)
            bao.append(f"⚠️ 0 SKU IMPORT khớp DATA SP — xem khung đỏ phía trên")
        else:
            ss.data_sp = _loc_kq
            if tk["bo"]:
                bao.append(f"đã lọc bỏ {tk['bo']:,} dòng DATA SP của {tk['sku_bo']:,} SKU không có trong IMPORT")
        if "data_sp" not in phan:
            phan.append("data_sp")
    if "data_sp" in phan:
        ss.data_sp = C.nen_df(ss.data_sp)
        bao.append(f"DATA SP {len(ss.data_sp):,} dòng")
    if len(ss["import"]) and len(ss.data_sp):  # dòng nhập theo ID mà DATA SP không có PRODUCTID đó -> báo rõ
        _id_imp = ss["import"].sku[ss["import"].sku.astype(str).str.startswith(C.ID_TIEN_TO)]
        if len(_id_imp):
            _co = set(ss.data_sp.PRODUCTCODE.astype(object).unique())
            _thieu = [x[len(C.ID_TIEN_TO):] for x in _id_imp if x not in _co]
            if _thieu:
                bao.append(f"⚠️ {len(_thieu):,}/{len(_id_imp):,} dòng nhập theo ID CMS KHÔNG thấy trong DATA SP "
                           f"(ID: {', '.join(_thieu[:5])}{'…' if len(_thieu) > 5 else ''}) — kiểm tra ID hoặc nạp file CMS export đúng")
    if da_ng:
        bao.insert(0, "Cấu hình ngành " + "; ".join(da_ng))
        if not phan:
            bump()
            ss.flash = ["✔ Đã nạp: " + " · ".join(bao)]
            ss.pop(f"{key}_sig", None)
            ss.pop(f"{key}_kq", None)  # bảng file vừa đọc không còn cần -> trả RAM
            ss[f"{key}_lan"] = lan + 1
            if len(ss["import"]) and len(ss.data_sp):
                chay_map_ui()
            st.rerun()
    if not phan and not da_ng:
        # Chẩn đoán cụ thể: tool không nhận ra file nào có dữ liệu để nạp
        chi_tiet = []
        for ten, r in kq:
            lo = r.get("loai") or "❌ không nhận ra"
            if r.get("loi"):
                chi_tiet.append(f"**{ten}**: {lo} — {r['loi']}")
            else:
                chi_tiet.append(f"**{ten}**: {lo}")
        st.error("⚠️ KHÔNG CÓ DỮ LIỆU NÀO ĐỂ NẠP. Có thể các file đã upload không được nhận diện, hoặc file rỗng. "
                 "Kiểm tra từng file bên dưới (xem cột 'Nhận là' trong bảng tóm tắt phía trên); nếu file đúng mà "
                 "'không nhận ra' → gửi file cho mình kiểm tra lại.")
        for d in chi_tiet:
            st.caption("• " + d)
        return
    bump()
    if phan and luu(phan + ["settings"], "Nạp nhanh: " + ", ".join(t for t, _ in kq)):
        ss.flash = ["✔ Đã nạp: " + " · ".join(bao)]
        ss.pop(f"{key}_sig", None)
        ss.pop(f"{key}_kq", None)  # bảng file vừa đọc không còn cần -> trả RAM
        ss[f"{key}_lan"] = lan + 1
        if tu_map and len(ss["import"]) and len(ss.data_sp):
            chay_map_ui()
            if ss.get("map_msg"):
                ss.flash.append(ss.map_msg)
        st.rerun()  # vẽ lại cả trang theo dữ liệu mới (các bước sau tự mở)


def file_xin_data() -> dict:
    x = C.ds_xin_data(ss["import"], ss.data_sp, ss.bang, ss.meta.get("log", []))
    return x


def nut_xin_data(key: str) -> None:
    # Bảo vệ tốc độ: chỉ dựng lại file xin data khi dữ liệu THỰC SỰ đổi (trước đây dựng lại sau mỗi lần bấm bất kỳ)
    khoa = (ss.get("ver", 0), ss.get("ver_map", 0), id(ss["import"]), len(ss["import"]), id(ss.data_sp),
            len(ss.data_sp), id(ss.bang), id(ss.meta))
    if ss.get("_xd_khoa") != khoa or "_xd_kq" not in ss:
        x = file_xin_data()
        n_ = len(x["sku"])
        nm_ = int((x["model"]["Tình trạng"].str.startswith("MODEL")).sum()) if len(x["model"]) else 0
        ss._xd_kq = (n_, nm_, C.xlsx_nhieu_sheet({"XIN DATA CMS": x["sku"], "THEO MODEL": x["model"]}) if n_ else b"")
        ss._xd_khoa = khoa
    n, nm, xlsx_b = ss._xd_kq
    if not n:
        st.caption("✔ Mọi SKU đều có giá trị — không cần xin thêm data CMS.")
        return
    st.download_button(f"📨 Tải file xin data CMS ({n:,} SKU · {nm:,} model không có giá trị)",
                       xlsx_b,
                       file_name=f"XIN_DATA_CMS_{ss.ws}_{C.bay_gio()[:10]}.xlsx", key=key,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _loc_bang(df: pd.DataFrame, q: str, cot: list | None = None) -> pd.DataFrame:
    q = (q or "").strip()
    if not q or df is None or not len(df):
        return df
    cot = [c for c in (cot or list(df.columns)[:8]) if c in df.columns]
    m = pd.Series(False, index=df.index)
    for c in cot:
        m |= df[c].astype(str).str.contains(q, case=False, regex=False, na=False)
    return df[m]


def xem_du_lieu() -> None:
    """Xem NGAY dữ liệu đã nạp (không cần map): lô · data gốc · đơn vị đã đặt."""
    st.caption("Xem những gì đang có trong tool. Gõ vào ô tìm để lọc (mã SKU, model, tên, mã cột…).")
    tabs = st.tabs([f"📋 IMPORT ({len(ss['import']):,})", f"📦 DATA SP ({len(ss.data_sp):,})",
                    f"🧾 SPEC PIM ({len(ss.spec):,})", f"🏷️ Cấu hình ({len(ss.cau_hinh)})",
                    f"🧬 Mapping TSKT ({len(ss.map_tskt):,})", f"🧮 Mapping FILTER ({len(ss.map_filter):,})",
                    f"🗂️ DATA PIM ({len(ss.data_pim):,})", "📏 Đơn vị & biến đổi đã đặt"])

    def bang(i: int, df: pd.DataFrame, key: str, cot: list | None = None) -> None:
        with tabs[i]:
            if df is None or not len(df):
                st.info("Chưa có dữ liệu.")
                return
            q = st.text_input("🔎 Tìm", key=f"xem_q_{key}", placeholder="gõ để lọc")
            d = _loc_bang(df, q, cot)
            st.caption(f"{len(d):,} / {len(df):,} dòng" + (" — hiện 1.000 dòng đầu" if len(d) > 1000 else ""))
            st.dataframe(d.head(1000).astype(str), hide_index=True, height=360)
    bang(0, ss["import"], "imp")
    bang(1, ss.data_sp, "sp", ["PRODUCTCODE", "PRODUCTNAME", "PROPERTYNAME", "PROPVALUE", "CATEGORYID"])
    bang(2, ss.spec, "spec", ["sku", "model_code", "ma", "ten", "gia_tri"])
    with tabs[3]:
        if not ss.cau_hinh:
            st.info("Chưa có cấu hình ngành hàng.")
        else:
            rows = [{"Mã NH": c, "Tên ngành": v.get("ten", ""), "Số cột": len(v.get("cot", [])),
                     "Danh sách cột": ", ".join(v.get("cot", []))} for c, v in sorted(ss.cau_hinh.items())]
            st.dataframe(pd.DataFrame(rows), hide_index=True, height=360)
    bang(4, ss.map_tskt, "mt")
    bang(5, ss.map_filter, "mf")
    bang(6, ss.data_pim, "dp")
    with tabs[7]:
        rows = []
        for k, v in ss.dv.items():
            if not v:
                continue
            if len(k) == 3:
                rows.append({"Loại": "Biến đổi hàng loạt", "Ngành": k[0], "Mã cột": k[1],
                             "Nội dung": " → ".join(f"{C.BD_KIEU.get(x['kieu'], x['kieu'])} [{x.get('a', '')}"
                                                   f"{(' → ' + x['b']) if x.get('b') else ''}]" for x in v)})
            else:
                ten = ss.cau_hinh.get(k[0], {}).get("ten_cot", {}).get(k[1], "")
                rows.append({"Loại": "Đơn vị theo cột", "Ngành": k[0], "Mã cột": f"{k[1]} {('— ' + ten) if ten else ''}".strip(),
                             "Nội dung": v if isinstance(v, str) else str(v)})
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, height=300)
        else:
            st.info("Chưa đặt đơn vị / biến đổi nào. Sau khi Map, vào bước ③ → tab **📏 Đơn vị & kích thước** "
                    "để thêm đơn vị cho dài / rộng / cao / khối lượng…")
        st.caption(f"Ô sửa tay: {len(ss.sua):,} · quy tắc Không/Đang cập nhật: "
                   f"{len([v for v in ss.rong.values() if v[0] != C.HD_GIU])}")


def khu_nap_rieng() -> None:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. File CMS export → DATA SP")
        fs = st.file_uploader("Chọn 1 hoặc nhiều file CMS export (.xlsx/.csv)", type=["xlsx", "xlsm", "csv"],
                              accept_multiple_files=True, key="up_cms")
        che_do = st.radio("Cách ghi", ["Nối tiếp (giữ dữ liệu cũ)", "Ghi đè (xoá DATA SP cũ)"], key="cd_cms",
                          horizontal=True)
        if st.button("📥 Nạp vào DATA SP", type="primary", disabled=not fs):
            ds, bao = [], []
            with st.spinner("Đang đọc file…"):
                for f in fs:
                    df, loi = C.doc_cms_export(f.getvalue(), f.name)
                    bao.append(f"✖ {f.name}: {loi}" if loi else f"✔ {f.name}: {len(df):,} dòng")
                    if not loi:
                        ds.append(df)
            if ds:
                moi = pd.concat(ds, ignore_index=True)
                cu = ss.data_sp if che_do.startswith("Nối") else ss.data_sp.iloc[0:0]
                ss.data_sp = C.nen_df(pd.concat([cu.astype(object), moi], ignore_index=True)
                                      .drop_duplicates(ignore_index=True))
                bump()
                luu(["data_sp"], f"Nạp {len(fs)} file CMS ({len(moi):,} dòng)")
            st.info("\n".join(bao))
        st.caption(f"DATA SP hiện có **{len(ss.data_sp):,}** dòng · "
                   f"{ss.data_sp.PRODUCTCODE.nunique() if len(ss.data_sp) else 0:,} SKU · "
                   f"{ss.data_sp.CATEGORYID.nunique() if len(ss.data_sp) else 0} ngành hàng")
    with c2:
        st.subheader("2. File export PIM → IMPORT + spec cũ")
        fp = st.file_uploader("File export PIM (dòng 1 mã cột, dòng 2 tên, dữ liệu từ dòng 3)",
                              type=["xlsx", "xlsm", "csv"], key="up_pim")
        che_do2 = st.radio("Cách ghi", ["Ghi đè (thay SKU + spec cũ)", "Nối tiếp (SKU nạp lại thì thay spec)"],
                           key="cd_pim", horizontal=True)
        if st.button("📦 Nạp file PIM", type="primary", disabled=not fp):
            with st.spinner("Đang đọc file PIM…"):
                r = C.doc_pim_export(fp.getvalue(), fp.name)
            if r["loi"]:
                st.error(r["loi"])
            else:
                if che_do2.startswith("Ghi"):
                    ss["import"], ss.spec = r["import"], C.nen_df(r["spec"])
                else:
                    moi_sku = set(r["import"].sku)
                    ss["import"] = pd.concat([ss["import"][~ss["import"].sku.isin(moi_sku)], r["import"]],
                                             ignore_index=True)
                    ss.spec = C.nen_df(pd.concat([ss.spec[~ss.spec.sku.isin(moi_sku)].astype(object), r["spec"]],
                                                 ignore_index=True))
                bump()
                luu(["import", "spec"], f"Nạp file PIM {fp.name}")
                msg = f"✔ {len(r['import']):,} SKU → IMPORT · {len(r['spec']):,} ô spec ({r['so_cot_spec']} cột)"
                canh = [f"{r['thieu_model']} SKU trống model_code" if r["thieu_model"] else "",
                        f"{r['thieu_cate']} SKU trống category_code" if r["thieu_cate"] else "",
                        f"bỏ {r['bo_khong_sku']} dòng không có sku" if r["bo_khong_sku"] else "",
                        f"bỏ {r['trung']} dòng trùng sku" if r["trung"] else ""]
                canh = [x for x in canh if x]
                if r.get("bien_the_mau"):
                    msg += f" · 🎨 {r['bien_the_mau']} SKU biến thể theo màu → xuất file MODEL"
                (st.warning if canh else st.success)(msg + ("  ·  ⚠️ " + " · ".join(canh) if canh else ""))
        st.caption(f"IMPORT hiện có **{len(ss['import']):,}** SKU · spec cũ của "
                   f"{ss.spec.sku.nunique() if len(ss.spec) else 0:,} SKU")
    st.divider()
    with st.expander("📋 Danh sách SKU (IMPORT) — xem / sửa trực tiếp / dán thêm", expanded=False):
        c1, c2 = st.columns([2, 1])
        with c1:
            ed = st.data_editor(ss["import"], num_rows="dynamic", hide_index=True,
                                height=320, key=f"ed_imp_{ss.ver}")
            if st.button("💾 Lưu IMPORT đã sửa"):
                ed = ed.fillna("").astype(str)
                ed["sku"] = ed.sku.map(C.chuan_hoa_code)
                ss["import"] = ed[ed.sku != ""].drop_duplicates("sku", ignore_index=True)
                bump()
                luu(["import"], "Sửa IMPORT trực tiếp")
                st.rerun()
        with c2:
            txt = st.text_area("Dán nhanh (mỗi dòng: sku hoặc model<TAB>sku<TAB>variant<TAB>category — copy từ Excel)",
                               height=200)
            if st.button("➕ Thêm vào IMPORT", disabled=not txt.strip()):
                rows = []
                for line in txt.splitlines():
                    p = [x.strip() for x in line.split("\t")]
                    if len(p) == 1 and p[0]:
                        rows.append(["", C.chuan_hoa_code(p[0]), "", ""])
                    elif len(p) >= 2 and p[1]:
                        p += [""] * 4
                        rows.append([p[0], C.chuan_hoa_code(p[1]), p[2], C.chuan_hoa_id(p[3])])
                moi = pd.DataFrame(rows, columns=C.COT_IMPORT)
                moi = moi[moi.sku.str.fullmatch(r"[0-9A-Za-z_\-]+", na=False)]
                ss["import"] = pd.concat([ss["import"][~ss["import"].sku.isin(set(moi.sku))], moi],
                                         ignore_index=True)
                bump()
                luu(["import"], f"Dán {len(moi)} SKU vào IMPORT")
                st.rerun()
            fi = st.file_uploader("Hoặc file danh sách SKU", type=["xlsx", "csv"], key="up_imp")
            if fi and st.button("📋 Nạp file SKU (ghi đè IMPORT)"):
                df, loi = C.doc_import_don_gian(fi.getvalue(), fi.name)
                if loi:
                    st.error(loi)
                else:
                    ss["import"] = df
                    bump()
                    luu(["import"], f"Nạp file SKU {fi.name}")
                    st.rerun()
    with st.expander("🗑️ Dọn dữ liệu lô cũ", expanded=False):
        xn = st.checkbox("Tôi muốn xoá (không hoàn tác được)")
        c = st.columns(4)
        if c[0].button("Xoá DATA SP", disabled=not xn):
            ss.data_sp = ss.data_sp.iloc[0:0]
            bump()
            luu(["data_sp"], "Xoá DATA SP")
            st.rerun()
        if c[1].button("Xoá IMPORT", disabled=not xn):
            ss["import"] = ss["import"].iloc[0:0]
            bump()
            luu(["import"], "Xoá IMPORT")
            st.rerun()
        if c[2].button("Xoá spec PIM cũ", disabled=not xn):
            ss.spec = ss.spec.iloc[0:0]
            bump()
            luu(["spec"], "Xoá spec PIM")
            st.rerun()
        if c[3].button("Xoá kết quả map + sửa tay", disabled=not xn):
            ss.bang, ss.meta, ss.sua = {}, {}, {}
            bump()
            luu(["ket_qua", "settings"], "Xoá kết quả map")
            st.rerun()


# ============================================================================
# TRANG: MAP & KIỂM TRA
# ============================================================================
NHAN_MAP_TEN = {"tat": "Tắt — chỉ theo mã PROPERTYID (đúng bản desktop, khuyên dùng)",
                "cau_hinh": "Thêm: theo tên cột của chính ngành hàng (khác desktop — soát lại)",
                "tham_chieu": "Thêm bảng tham chiếu tên (rộng hơn — cần soát lại)"}



# ============================================================================
# AI tự học từ các lô ĐÃ XUẤT (được coi là đã xác nhận bởi người dùng)
# KHÔNG đụng vào logic map/QC; chỉ GOM top giá trị theo (cate, cột) để sau này so sánh.
# ============================================================================
def cap_nhat_ai_hoc_tu_bang(chi_cate: list | None = None) -> dict:
    """Trích top-20 giá trị + đếm theo (cate, cột) từ GIÁ TRỊ SẼ XUẤT của bang hiện tại.
    Gộp vào ss.ai_hoc rồi lưu lên kho dùng chung. Trả về thống kê để báo."""
    from collections import Counter
    hoc = dict(ss.get("ai_hoc") or {})
    tk = {"cate": 0, "cot": 0, "o": 0}
    for cate, b in ss.bang.items():
        if chi_cate is not None and cate not in chi_cate:
            continue
        c_hoc = dict(hoc.get(cate, {}))
        tk["cate"] += 1
        for code in C.cot_tt(b):
            if code.lower() in C.COT_KHONG_PHAI_SPEC or C.la_cot_filter(code):
                continue
            dem = Counter()
            for r in b["rows"]:
                v, _ = C.bien_doi_o(cate, r["sku"], code, r["vals"].get(code, ""), ss.sua, ss.dv, ss.rong)
                if v:
                    dem[str(v)[:120]] += 1
            if not dem:
                continue
            tk["cot"] += 1
            tk["o"] += sum(dem.values())
            cu = c_hoc.get(code, {}) or {}
            tv_cu = Counter(cu.get("top_values") or {})
            tv_cu.update(dem)
            top = dict(tv_cu.most_common(20))
            c_hoc[code] = {"top_values": top, "total_samples": cu.get("total_samples", 0) + sum(dem.values()),
                           "last_updated": C.bay_gio()[:10],
                           "contributors": sorted(set((cu.get("contributors") or []) + [ss.user]))[:20]}
        hoc[cate] = c_hoc
    ss.ai_hoc = hoc
    try:
        luu(["shared:ai_hoc"], f"[{ss.user}] AI tự học {tk['cate']} ngành · {tk['cot']} cột · {tk['o']} ô")
    except Exception:  # noqa: BLE001
        pass
    return tk


def qc_nguoc_tu_ai_hoc() -> list:
    """QC ngược: so các ô CÓ giá trị hiện tại với mẫu đã học của chính ngành đó.
    Flag: ô có giá trị CHƯA TỪNG GẶP trong 20 giá trị thường gặp, với ngành đã có ≥ 10 mẫu.
    Trả về list dict {cate, sku, ma, ten, gia_tri_hien, goi_y_top, ly_do}."""
    hoc = ss.get("ai_hoc") or {}
    kq = []
    for cate, b in ss.bang.items():
        c_hoc = hoc.get(cate) or {}
        if not c_hoc:
            continue
        for code in C.cot_tt(b):
            if code.lower() in C.COT_KHONG_PHAI_SPEC or C.la_cot_filter(code):
                continue
            info = c_hoc.get(code)
            if not info or info.get("total_samples", 0) < 10:
                continue
            top = info.get("top_values") or {}
            top_set = set(top)
            for r in b["rows"]:
                v, _ = C.bien_doi_o(cate, r["sku"], code, r["vals"].get(code, ""), ss.sua, ss.dv, ss.rong)
                if not v:
                    continue
                v_short = str(v)[:120]
                if v_short in top_set:
                    continue
                goi_y = ", ".join(list(top.keys())[:3])
                kq.append({
                    "cate": cate, "sku": r["sku"], "ma": code,
                    "ten": b["ten"].get(code, ""),
                    "gia_tri_hien": v_short,
                    "goi_y_top": goi_y,
                    "ly_do": f"Chưa gặp trong {info.get('total_samples', 0):,} mẫu đã duyệt — "
                             f"thường gặp: {goi_y}",
                })
    return kq


def gom_o_theo_ten(o: dict) -> list:
    """Gom các ô điền theo TÊN thành nhóm (ngành, mã cột) để hiển thị cảnh báo + duyệt."""
    g: dict = {}
    for k, v in (o or {}).items():
        sku = k.split("\t")[0]
        x = g.setdefault((v["cate"], v["ma"]), {"cate": v["cate"], "ma": v["ma"], "props": {}, "skus": [],
                                               "trung_ten": ""})
        x["props"].update(v.get("props", {}))
        x["skus"].append(sku)
        if v.get("trung_ten"):
            x["trung_ten"] = v["trung_ten"]
    return list(g.values())


def o_ten_nhom() -> list:
    """Các nhóm ô điền theo tên CHƯA được duyệt."""
    duyet = set(ss.get("duyet_ten", []))
    return [x for x in ss.meta.get("o_theo_ten", []) if f"{x['cate']}\t{x['ma']}" not in duyet]


def o_ten_bi_chan() -> set:
    """Tập khoá "sku\tmã" bị chặn khỏi file import (cột điền theo tên chưa duyệt)."""
    return {f"{s_}\t{x['ma']}" for x in o_ten_nhom() for s_ in x["skus"]}


def canh_bao_o_ten(khung=st) -> None:
    """Cảnh báo ĐỎ: ô điền theo tên -> ghi rõ giá trị đang đi vào cột nào. Chặn khỏi import tới khi duyệt."""
    ds_ = o_ten_nhom()
    if not ds_:
        return
    dong = []
    for x in ds_:
        b = ss.bang.get(x["cate"], {})
        sk = set(x["skus"])
        vd = next((r["vals"].get(x["ma"], "") for r in b.get("rows", []) if r["sku"] in sk and r["vals"].get(x["ma"])), "")
        pr = ", ".join(f"«{n}» (PROPERTYID {p})" for p, n in list(x["props"].items())[:3])
        trung = f" ⚠️ tên này CŨNG trùng cột `{x['trung_ten']}` đã map theo mã." if x.get("trung_ten") else ""
        dong.append(f"- Ngành **{x['cate']}** · thuộc tính CMS {pr} → cột **`{x['ma']}`** · "
                    f"**{len(sk)} SKU**, ví dụ `{str(vd)[:60]}`.{trung}")
    khung.error("⛔ **Có ô được điền THEO TÊN (không phải theo mã PROPERTYID — không có trong bản desktop).** "
                "Các cột sau **KHÔNG được đưa vào file import** cho tới khi bạn duyệt "
                "(vào 🧭 QC tổng hợp → mục «Ô điền theo tên»):\n\n" + "\n".join(dong))


def chay_map_ui() -> None:
    with viec_nang("Map dữ liệu", uoc_mb_map()) as _ok:
        if _ok:
            _chay_map_ui_thuc()


def _chay_map_ui_thuc() -> None:
    ss.pop("xuat", None); ss.pop("ws_mau", None)  # map lại → bỏ file xuất cũ, tránh tải nhầm bản trước
    thieu_import = not len(ss.get("import", []))
    thieu_dsp = not len(ss.get("data_sp", []))
    if thieu_import or thieu_dsp:
        chi_tiet = []
        if thieu_import:
            chi_tiet.append("**IMPORT** (danh sách SKU: model_code / sku / variant_code / category_code — "
                            "từ file export PIM như `export_product...xlsx`)")
        if thieu_dsp:
            chi_tiet.append("**DATA SP** (dữ liệu thô CMS: PRODUCTID / PROPERTYID / PROPVALUE / CATEGORYID — "
                            "từ file CMS export, mỗi SKU nhiều dòng)")
        st.error("⚠️ Thiếu dữ liệu để map:\n\n" + "\n".join(f"- {x}" for x in chi_tiet) +
                 "\n\nVào **🚀 Chạy pipeline → ① Nạp dữ liệu** (tab đầu) kéo thả file còn thiếu, "
                 "hoặc dùng **Data gốc** nếu là file workspace theo mẫu (nhiều sheet).")
        return
    t = time.time()
    sk, gt, qd, n_rieng = qt()
    with st.spinner("Đang map dữ liệu…"):
        r = C.chay_map(ss.data_sp, ss["import"], ss.cau_hinh, ss.map_tskt, ss.map_filter, ss.opt, ss.map_ten,
                       qd, gt, sk)
    ss.bang = r["bang"]
    ss.ver_map = ss.get("ver_map", 0) + 1
    ss.meta = {"nganh": {c: {k: b[k] for k in ("title", "ten_nh", "attr", "ten")} for c, b in r["bang"].items()},
               "luc": r["luc"], "tom_tat": r["tom_tat"], "log": r["log"][:5000],
               "chua_map": r["chua_map"].to_dict("records"), "goc_cms": r["goc_cms"], "dx_dang_ap": n_rieng,
               "chon": r["chon"], "o_theo_ten": gom_o_theo_ten(r.get("o_theo_ten", {}))}
    bump()
    luu(["ket_qua", "settings"], f"Map {r['tom_tat']['so_sku']} SKU / {r['tom_tat']['so_nganh']} ngành hàng")
    ss.map_msg = (f"✔ Map xong {r['tom_tat']['so_sku']:,} SKU · {r['tom_tat']['so_nganh']} ngành hàng · "
                  f"{r['tom_tat']['so_o']:,} ô · {time.time() - t:.1f} giây"
                  + (f" · ⚠️ {r['tom_tat']['map_theo_ten']:,} ô điền THEO TÊN (đang bị chặn, chưa vào file import)"
                     if r['tom_tat']['map_theo_ten'] else ""))
    try:  # QC ngầm sau map: báo ngay nếu ngành nào sót cấu hình / mapping
        _L = nq()["loi"]
        _c = int((_L["Mức"] == "CAO").sum()) if len(_L) else 0
        _t = int((_L["Mức"] == "TB").sum()) if len(_L) else 0
        ss.map_msg += (f" · QC ngầm: ⛔ {_c} lỗi nhất quán, {_t} cần xem (tab 🧭)" if _c else
                       (f" · QC ngầm: {_t} điểm cần xem (tab 🧭)" if _t else " · QC ngầm: ✅ nhất quán"))
    except Exception:  # noqa: BLE001
        pass
    st.success(ss.map_msg)


def the_so(k: dict) -> None:
    s = k["stat"]
    co = k["co_doi_chieu"]
    items = [("Dòng sẽ xuất", s.get("so_dong", 0), None),
             ("Ô khác spec PIM", s.get(C.TRANG_THAI_KHAC, 0) if co else "–", "🔴"),
             ("Tool trống (PIM có)", s.get(C.TRANG_THAI_TOOL_TRONG, 0) if co else "–", "🟠"),
             ("Chỉ thêm đơn vị", s.get(C.TRANG_THAI_DON_VI, 0) if co else "–", "🟢"),
             ("Thiếu model/category", s.get("thieu_model", 0) + s.get("thieu_cate", 0), "🔴"),
             ("Ô chưa có đơn vị", s.get("chua_don_vi", 0), "🟡"),
             ("Ô Không/Đang cập nhật", s.get("rong_tong", 0), "⚪"),
             ("FILTER sai mã", s.get("filter_chu", 0), "🔴"),
             ("Ô sửa tay", s.get("bd_sua", 0), "✎"),
             ("Nghi sai (kiểm tra thông minh)", int(ttm().muc_do.isin([C.MUC_CAO, C.MUC_TB]).sum()) if ss.bang else 0,
              "🧠"),
             ("Mất dữ liệu khi map (ô)", ds()["mat_o"] if ss.bang else 0, "🧾"),
             ("SKU không có thông số", _sku_rong() if ss.bang else 0, "🧾")]
    # Mỗi ô là NÚT: bấm → mở ngay bảng xem & sửa bên dưới (cùng bảng với QC tổng hợp); bấm lại để đóng.
    kind_cua = {1: "khac_spec", 2: "tool_trong", 4: "thieu_model" if s.get("thieu_model", 0) else "thieu_cate",
                5: "kich_thuoc", 6: "rong", 7: "filter_chu", 8: "sua_tay", 9: "ttm", 11: "xin_data"}
    goi_y = {0: "Số dòng sẽ có trong file import — tạo file ở vùng ④ Xuất.",
             3: "Ô chỉ khác ở đơn vị so với PIM cũ — an toàn, không cần sửa.",
             10: "Giá trị CMS không map được sang cột nào — xem tab 🧾 Đối soát CMS → kết quả."}
    dang = ss.get("the_chon")
    for h, hang in enumerate((items[:6], items[6:])):
        cols = st.columns(6)
        for n, (c, (t, v, ic)) in enumerate(zip(cols, hang)):
            idx = h * 6 + n
            vs = f"{v:,}".replace(",", ".") if isinstance(v, int) else v
            if c.button(f"**{vs}**  \n{ic or ''} {t}".strip(), key=f"the_{idx}", width="stretch",
                        type="primary" if dang == idx else "secondary"):
                ss.the_chon = None if dang == idx else idx
                st.rerun()
    ss._the_kind = kind_cua.get(dang) if dang is not None else None
    if dang is None:
        return
    nhan = items[dang][0]
    with st.container(border=True):
        top = st.columns([6, 1])
        top[0].markdown(f"**{nhan}** — bảng xem & sửa")
        if top[1].button("✕ Đóng", key="the_dong"):
            ss.the_chon = None
            st.rerun()
        if dang in kind_cua:
            try:
                if kind_cua[dang] == "sua_tay":
                    _bang_sua_tay()
                else:
                    khu_sua_loi(kind_cua[dang], k)
            except Exception as e:  # noqa: BLE001  (vd trùng khoá widget với tab bên dưới)
                st.info("Bảng này đang mở ở tab bên dưới — xem & sửa ở đó. " + type(e).__name__)
        else:
            st.info(goi_y.get(dang, "Mục này không có bảng sửa riêng."))


def _bang_sua_tay() -> None:
    if not ss.sua:
        st.success("Chưa có ô nào sửa tay.")
        return
    df = pd.DataFrame([{"Bỏ": False, "Ngành": c, "SKU": sku, "Cột": m, "Giá trị sửa": v}
                       for (c, sku, m), v in sorted(ss.sua.items())][:3000])
    ed = st.data_editor(df, hide_index=True, width="stretch", height=min(420, 60 + 35 * len(df)),
                        key=f"ed_sua_tay_{len(ss.sua)}", disabled=["Ngành", "SKU", "Cột", "Giá trị sửa"])
    if st.button("↩ Bỏ các ô đã tick (về giá trị CMS)", key="bo_sua_tay_tick") and ed["Bỏ"].any():
        for r in ed[ed["Bỏ"]].to_dict("records"):
            ss.sua.pop((r["Ngành"], r["SKU"], r["Cột"]), None)
        bump()
        luu(["settings"], "Bỏ sửa tay đã tick")
        st.rerun()


def _tab_an_toan(ten: str, ham, *a) -> None:
    """Chạy 1 tab; lỗi ở tab này CHỈ hiện trong tab đó (kèm 1 dòng chi tiết để báo), không làm sập cả trang
    và không chặn các tab/khu phía sau. st.rerun/st.stop (BaseException) vẫn chạy bình thường."""
    try:
        ham(*a)
    except Exception as e:  # noqa: BLE001
        _ghi_loi(e, f"Tab {ten}")
        st.error(f"Tab «{ten}» gặp lỗi: {type(e).__name__}: {str(e)[:300]}")
        if ss.get("admin"):
            st.code("".join(traceback.format_exception(e))[-1500:], language="python")


def trang_map() -> None:
    st.title("🔍 Kiểm tra & Đối chiếu")
    if ss.get("map_flash"):
        st.success(ss.pop("map_flash"))
    # --- RÀO BƯỚC 1: phải nạp dữ liệu trước khi map ---
    _co_import = len(ss.get("import", [])) > 0
    _co_data_sp = len(ss.get("data_sp", [])) > 0
    if not _co_import and not _co_data_sp and not ss.get("bang"):
        st.warning("⚠️ **Chưa có dữ liệu.** Vào **🚀 Chạy pipeline → ① Nạp dữ liệu**, nạp file lô trước rồi quay lại đây.")
        return
    # Cảnh báo khi kết quả map CÓ nhưng IMPORT/DATA SP TRỐNG — thường do session trước lưu được ket_qua
    # mà không lưu được import/data_sp (crash giữa chừng). Người dùng cần nạp lại.
    if ss.bang and (not _co_import or not _co_data_sp):
        st.warning("⚠️ Kết quả map có sẵn ({} ngành) nhưng **IMPORT ({} SKU) hoặc DATA SP ({} dòng) đang trống** "
                   "trong kho — có thể lần nạp trước bị lỗi lưu. Bạn cần vào 🚀 Chạy pipeline → ① Nạp dữ liệu "
                   "nạp lại file lô (SKU + DATA SP), sau đó bấm Map lại.".format(
                       len(ss.bang), len(ss.get("import", [])), len(ss.get("data_sp", []))))
    _xong_nap = _co_import and _co_data_sp
    c1, c2, c3 = st.columns([2, 3, 2])
    with c1:
        if st.button("① Map dữ liệu", type="primary", width="stretch", disabled=not _xong_nap):
            chay_map_ui()
    with c2:
        mt = st.selectbox("Thuộc tính chưa có trong MAPPING (theo mã) — map dự phòng theo tên:",
                          list(NHAN_MAP_TEN), format_func=NHAN_MAP_TEN.get,
                          index=list(NHAN_MAP_TEN).index(ss.get("map_ten", "tat")))
        if mt != ss.map_ten:
            ss.map_ten = mt
            luu(["settings"], "Đổi chế độ map theo tên")
    with c3:
        if ss.meta.get("luc"):
            st.caption(f"Map lần cuối: {ss.meta['luc']}")
            st.caption(" · ".join(f"{k}: {v}" for k, v in ss.meta.get("tom_tat", {}).items()))
    if bang_trong():
        st.info("Chưa có kết quả map — bấm **① Map dữ liệu**.")
        return
    canh_bao_o_ten()
    k = kq()
    the_so(k)
    c = st.columns([1.3, 1.3, 3])
    if c[0].button("✨ Ô tool trống → lấy PIM cũ", width="stretch",
                   help="Lấp các ô tool để trống mà PIM đang có giá trị (để import không làm mất dữ liệu web)"):
        n = 0
        for r in k["khac"]:
            if r["trang_thai"] == C.TRANG_THAI_TOOL_TRONG and r["pim_cu"]:
                dat_sua(r["cate"], r["sku"], r["ma"], r["pim_cu"], r["goc"])
                n += 1
        bump()
        luu(["settings"], f"Tự động lấp {n} ô tool trống bằng PIM cũ")
        st.rerun()
    if c[1].button("🗑 Xoá hết sửa tay", width="stretch", disabled=not ss.sua):
        ss.sua = {}
        bump()
        luu(["settings"], "Xoá hết sửa tay")
        st.rerun()
    sk_, gt_, qd_, n_r = qt()
    c[2].caption(f"Quy tắc sửa CMS: {len(sk_) + len(gt_)} (trong đó {n_r} đề xuất của bạn chờ duyệt) · "
                 f"{ss.meta.get('tom_tat', {}).get('o_sua_theo_quy_tac', 0)} ô đã sửa theo quy tắc. "
                 f"Đang áp: {len(ss.sua)} ô sửa tay · {len([1 for kk, v in ss.dv.items() if v and len(kk) == 2])} cột có đơn vị · "
                 f"{len([1 for kk, v in ss.dv.items() if v and len(kk) == 3])} cột biến đổi hàng loạt · "
                 f"{len([v for v in ss.rong.values() if v[0] != C.HD_GIU])} quy tắc Không/Đang cập nhật. "
                 "Mọi chỉnh sửa tự lưu, áp khi xuất file.")
    kcl = kc()["loi"]
    kc_n = int(kcl["Mức"].isin(["CAO", "TB"]).sum()) if len(kcl) else 0
    nql = nq()["loi"]
    nq_n = int(nql["Mức"].isin(["CAO", "TB"]).sum()) if len(nql) else 0
    _nhan_tab = [
        "🛡️ QC tổng hợp", f"✅ Kiểm chứng SKU ↔ DATA SP{' (' + str(kc_n) + ')' if kc_n else ' ✔'}",
        f"🧭 Nhất quán ngành{' (' + str(nq_n) + ')' if nq_n else ' ✔'}", "⚠️ Cảnh báo",
        "📈 Độ hoàn thiện & quy tắc", "🧾 Đối soát CMS → kết quả", "💡 Gợi ý thông minh & AI",
        "≠ Khác spec PIM (sửa)", "🔎 Theo SKU + FILTER", "📏 Đơn vị & biến đổi hàng loạt", "🚫 Không / Đang cập nhật",
        "📐 Gộp / tách kích thước", "🧩 Thuộc tính chưa map", "📜 Log map"]
    if not HIEN_AI:
        _nhan_tab = [x for x in _nhan_tab if "AI" not in x]
    _tabs = list(st.tabs(_nhan_tab))
    if not HIEN_AI:
        _tabs.insert(6, None)
    (t_qc, t_kc, t_nq, t_cb, t_ht, t_ds, t_ai, t_khac, t_sku, t_dv, t_rong, t_kt, t_cm, t_log) = _tabs
    with t_qc:
        _tab_an_toan('QC tổng hợp', tab_qc, k)
    with t_kc:
        _tab_an_toan('Kiểm chứng SKU ↔ DATA SP', tab_kiem_chung)
    with t_nq:
        _tab_an_toan('Nhất quán ngành', tab_nhat_quan)
    with t_cb:
        _tab_an_toan('Cảnh báo', tab_canh_bao, k)
    with t_ht:
        _tab_an_toan('Độ hoàn thiện & quy tắc', tab_hoan_thien)
    with t_ds:
        _tab_an_toan('Đối soát CMS → kết quả', tab_doi_soat)
    if t_ai is not None:
        with t_ai:
            tab_ai(k)
    with t_khac:
        _tab_an_toan('Khác spec PIM (sửa)', tab_khac, k)
    with t_sku:
        _tab_an_toan('Theo SKU + FILTER', tab_sku, k)
    with t_dv:
        _tab_an_toan('Đơn vị & biến đổi hàng loạt', tab_don_vi, k)
    with t_rong:
        _tab_an_toan('Không / Đang cập nhật', tab_rong, k)
    with t_kt:
        _tab_an_toan('Gộp / tách kích thước', tab_tach_kt)
    with t_cm:
        _tab_an_toan('Thuộc tính chưa map', tab_chua_map)
    with t_log:
        lg = pd.DataFrame(ss.meta.get("log", []), columns=["SKU", "Ngành hàng", "Mã", "Nguyên nhân"])
        st.dataframe(lg, hide_index=True, height=420)
    st.divider()
    st.markdown("### ④ Xuất file import — khi đã kiểm tra & đối chiếu xong")
    khu_xuat(k)


# ============================================================================
# QC TỔNG HỢP — gom mọi vùng kiểm tra: mức độ, ý nghĩa, gợi ý, nút sửa nhanh
# ============================================================================
MUC_ICON = {"CAO": "🔴 Cao", "TB": "🟠 Trung bình", "THAP": "🟢 Thấp", "TT": "ℹ️ Thông tin"}


def qc_tong_hop(k: dict) -> list:
    s = k["stat"]
    co = k["co_doi_chieu"]
    d = ds()
    L = d["loi"]
    xd = file_xin_data()
    ttm_ = ttm()
    h = dht()
    nf = int(L.loc[(L["_loai"] == "filter") & (L["Mức"] == C.MUC_CAO), "Số SKU"].sum()) if len(L) else 0
    nc = int(L.loc[L["_loai"] == "cot", "Số SKU"].sum()) if len(L) else 0
    rong_chua = sum(1 for x in k["gia_tri_rong"] if ss.rong.get(x["khoa"], [C.HD_GIU])[0] == C.HD_GIU)
    kcl = kc()["loi"]
    kc_cao = int((kcl["Mức"] == "CAO").sum()) if len(kcl) else 0
    kc_tb = int((kcl["Mức"] == "TB").sum()) if len(kcl) else 0
    q = [
        ("CAO", "Kiểm chứng SKU ↔ DATA SP: giá trị lệch SKU / không có nguồn", kc_cao,
         "Ô kết quả không truy được về DATA SP của CHÍNH SKU đó", "Xem từng ô bên dưới (tab ✅ Kiểm chứng)", None,
         "kiem_chung"),
        ("TB", "Kiểm chứng: không thấy nguồn / ký tự ẩn", kc_tb,
         "Giá trị không có trong DATA SP của SKU, hoặc ô chứa ký tự ẩn (tự bỏ khi xuất)",
         "Xem từng ô bên dưới (tab ✅ Kiểm chứng)", None, "kiem_chung"),
        ("CAO", "Nhất quán ngành: sót cấu hình / sót mapping / mapping sai cột / FILTER thiếu option",
         int((nq()["loi"]["Mức"] == "CAO").sum()) if len(nq()["loi"]) else 0,
         "Ngành trong DATA SP không khớp cấu hình · mapping TSKT/FILTER · DATA PIM", "Xem bên dưới (tab 🧭)", None,
         "nhat_quan"),
        ("CAO", "Ô điền THEO TÊN chưa duyệt", sum(len(x["skus"]) for x in o_ten_nhom()),
         "Không theo mã PROPERTYID (không có ở desktop) — đang bị chặn khỏi file import",
         "Xem giá trị đang đi vào cột nào, tick duyệt nếu đúng", None, "o_ten"),
        ("CAO", "Thiếu model_code", s.get("thieu_model", 0), "SKU không có Mã model → không import được",
         "Sửa trực tiếp cột Mã model bên dưới", "thieu_model", "thieu_model"),
        ("CAO", "Thiếu category_code", s.get("thieu_cate", 0), "IMPORT chưa có Mã danh mục PIM",
         "Sửa trực tiếp cột Mã danh mục bên dưới", "thieu_cate", "thieu_cate"),
        ("CAO", "ERP / model KHÔNG CÓ GIÁ TRỊ", len(xd["sku"]),
         "CMS chưa có data, ngành chưa có mapping, hoặc không thuộc tính nào map được",
         "Xem danh sách & tải file xin data CMS bên dưới", "xin_data", "xin_data"),
        ("CAO", "Mapping trỏ tới cột không có trong cấu hình", nc,
         "Giá trị CMS đã map nhưng ngành không có cột đó → MẤT khi xuất", "Xem cột thiếu & thêm vào cấu hình bên dưới",
         "them_cot" if nc else None, "cot_thieu"),
        ("CAO", "FILTER không khớp option", nf, "Giá trị CMS không trùng tên option nào trong DATA PIM → ô trống",
         "Chọn option đúng cho từng giá trị bên dưới", None, "filter_khong_khop"),
        ("CAO", "FILTER chứa chữ (sai mã)", s.get("filter_chu", 0), "PIM chỉ nhận mã option số",
         "Sửa trực tiếp từng ô FILTER bên dưới", None, "filter_chu"),
        ("CAO", "Vi phạm quy tắc mức Lỗi", int((h["vi_pham"]["Mức"] == C.MUC_LOI).sum()) if len(h["vi_pham"]) else 0,
         "Không đạt quy tắc chất lượng admin đặt", "Xem danh sách ô vi phạm bên dưới", None, "vi_pham"),
        ("TB", "Ô KHÁC spec PIM", s.get(C.TRANG_THAI_KHAC, 0) if co else 0, "Import sẽ ghi đè giá trị đang trên web",
         "Đối chiếu PIM cũ ↔ TOOL MỚI, sửa bên dưới", None, "khac_spec"),
        ("TB", "Tool trống nhưng PIM đang có", s.get(C.TRANG_THAI_TOOL_TRONG, 0) if co else 0,
         "CMS không có giá trị, PIM cũ có", "Lấy PIM cũ cho các ô này", "lay_pim" if co else None, "tool_trong"),
        ("TB", "Kích thước/khối lượng chưa có đơn vị", s.get("chua_don_vi", 0), "Ô số trơn ở cột kích thước",
         "Bảng dài · rộng · cao: thêm đơn vị + sửa số bên dưới", "don_vi", "kich_thuoc"),
        ("TB", "Giá trị Không / Đang cập nhật chưa chọn cách xử lý", rong_chua,
         "Mặc định GIỮ nguyên khi import", "Chọn Giữ / Để trống / Thay bằng cho từng giá trị bên dưới",
         "rong" if rong_chua else None, "rong"),
        ("TB", "Nghi sai (kiểm tra thông minh)", int(ttm_.muc_do.isin([C.MUC_CAO, C.MUC_TB]).sum()) if len(ttm_) else 0,
         "Lỗi gõ, lẫn đơn vị, giá trị bất thường", "Xem & áp từng gợi ý sửa bên dưới",
         "ttm" if len(ttm_) and ((ttm_.muc_do == C.MUC_CAO) & (ttm_.goi_y != "")).any() else None, "ttm"),
        ("TB", "SKU chưa đủ cột bắt buộc", int((~h["sku"]["Đủ bắt buộc"]).sum()) if len(h["sku"]) else 0,
         "Theo quy tắc Bắt buộc", "Xem danh sách SKU còn thiếu bên dưới", None, "sku_bat_buoc"),
        ("TB", "model_code lặp nhiều dòng MODEL", len(C.model_trung(ss.bang)), "PIM sẽ lấy dòng sau cùng",
         "Xem các model_code bị lặp bên dưới", None, "model_lap"),
        ("THAP", "Thuộc tính CMS chưa có mapping", len(ss.meta.get("chua_map", [])), "Không vào file import",
         "Xem thuộc tính chưa map bên dưới", None, "chua_map"),
        ("TT", "SKU không có trong file PIM cũ", s.get("khong_co_pim", 0) if co else 0, "Không đối chiếu được spec",
         "Bình thường với SKU mới", None, "khong_co_pim"),
    ]
    # Mục người dùng đã mở trong phiên được GIỮ trong bảng dù đã sửa xong (số lượng 0) -> mở lại để chỉnh tiếp, không biến mất
    giu = set(ss.get("qc_da_mo", []))
    return [x for x in q if x[2] or x[6] in giu]


def tab_qc(k: dict) -> None:
    q = qc_tong_hop(k)
    h = dht()
    n = len(h["sku"])
    san_sang = int(h["sku"]["Đủ bắt buộc"].sum()) if n else 0
    cao = sum(x[2] for x in q if x[0] == "CAO")
    c = st.columns(4)
    c[0].metric("Mục cần xử lý", len([x for x in q if x[0] in ("CAO", "TB")]))
    c[1].metric("Lỗi mức Cao (ô/SKU)", f"{cao:,}")
    c[2].metric("% đầy đủ trung bình", f"{h['nganh']['% đầy đủ TB'].mean():.0f}%" if len(h["nganh"]) else "–")
    c[3].metric("SKU sẵn sàng (đủ bắt buộc)", f"{san_sang:,}/{n:,}")
    if not q:
        st.success("✔ Không còn mục nào cần xử lý — cuộn xuống ④ Xuất file import.")
        return
    st.caption("👉 **Bấm vào một dòng** để mở bảng xem & sửa riêng cho lỗi đó (sửa, đối chiếu bằng mắt như Excel).")
    dfq = pd.DataFrame([{"Mức": MUC_ICON[a], "Vùng": b, "Số lượng": n_, "Ý nghĩa": y, "Cách xử lý": g}
                        for a, b, n_, y, g, act, kind in q])
    ev = st.dataframe(dfq, hide_index=True, width="stretch", key=f"qc_tbl_{ss.get('ver_map', 0)}",
                      on_select="rerun", selection_mode="single-row",
                      column_config={"Số lượng": st.column_config.NumberColumn(format="%d")})
    rows = []
    try:
        rows = list(ev.selection["rows"])
    except Exception:  # noqa: BLE001
        rows = list(getattr(getattr(ev, "selection", None), "rows", []) or [])
    # Nhớ mục đang mở THEO TÊN (không theo số thứ tự dòng): sửa xong, danh sách đổi/số lượng giảm vẫn không nhảy đi.
    kinds = [x[6] for x in q]
    sig = "|".join(kinds)
    if ss.get("qc_sig") == sig and rows != ss.get("qc_rows"):
        ss.qc_kind = q[rows[0]][6] if rows and rows[0] < len(q) else None
    ss.qc_sig, ss.qc_rows = sig, rows
    if ss.get("qc_kind") not in kinds:
        ss.qc_kind = None
    sel = next((x for x in q if x[6] == ss.get("qc_kind")), None)
    if sel:
        ss.qc_da_mo = sorted(set(ss.get("qc_da_mo", [])) | {sel[6]})
        muc, ten, so, y, g, act, kind = sel
        st.markdown(f"<div class='card'><span class='pill pill-{muc.lower()}'>{MUC_ICON[muc]}</span> "
                    f"<span class='card-h'>{ten}</span> · <b>{so:,}</b><br><span class='card-s'>{y} — {g}</span></div>",
                    unsafe_allow_html=True)
        if ss.get("_the_kind") == kind:
            st.info("👆 Bảng này đang mở ở ô số liệu phía trên — xem & sửa ở đó.")
        else:
            khu_sua_loi(kind, k)
    else:
        st.info("Chưa chọn dòng nào. Bấm 1 dòng ở bảng trên để sửa; hoặc dùng ⚡ Sửa nhanh hàng loạt bên dưới.")
    acts = {x[5] for x in q if x[5]}
    if not acts:
        return
    with st.expander("⚡ Sửa nhanh hàng loạt (áp cho tất cả, không cần chọn dòng)"):
        cc = st.columns(4)
        i = 0

        def nut(nhan, khoa):
            nonlocal i
            r = cc[i % 4].button(nhan, key=f"qc_{khoa}", width="stretch")
            i += 1
            return r
        if "xin_data" in acts:
            with cc[i % 4]:
                nut_xin_data("qc_xd")
            i += 1
        if "lay_pim" in acts and nut("✨ Ô tool trống → lấy PIM cũ", "lay_pim"):
            for r in k["khac"]:
                if r["trang_thai"] == C.TRANG_THAI_TOOL_TRONG and r["pim_cu"]:
                    dat_sua(r["cate"], r["sku"], r["ma"], r["pim_cu"], r["goc"])
            bump()
            luu(["settings"], "QC: lấy PIM cũ cho ô tool trống")
            st.rerun()
        if "don_vi" in acts and nut("📏 Điền đơn vị gợi ý (cột còn số trơn)", "dv"):
            for x in k["don_vi_cot"]:
                if x["so_tron"] and x["goi_y"] and x["la_kt"] and not ss.dv.get((x["cate"], x["code"])):
                    ss.dv[(x["cate"], x["code"])] = x["goi_y"]
            bump()
            luu(["settings"], "QC: điền đơn vị gợi ý")
            st.rerun()
        if "rong" in acts and nut("🚫 Không/Đang cập nhật → Để trống", "rong_trong"):
            for x in k["gia_tri_rong"]:
                ss.rong.setdefault(x["khoa"], [C.HD_TRONG, ""])
            bump()
            luu(["settings"], "QC: giá trị rỗng → để trống")
            st.rerun()
        if "ttm" in acts and nut("🧠 Áp gợi ý sửa mức Cao", "ttm"):
            t = ttm()
            m = 0
            for r in t[(t.muc_do == C.MUC_CAO) & (t.goi_y != "")].itertuples(index=False):
                b = ss.bang.get(r.cate, {})
                row = next((x for x in b.get("rows", []) if x["sku"] == r.sku), None)
                dat_sua(r.cate, r.sku, r.ma, r.goi_y, row["vals"].get(r.ma, "") if row else "")
                m += 1
            bump()
            luu(["settings"], f"QC: áp {m} gợi ý kiểm tra thông minh mức Cao")
            st.rerun()
        if "them_cot" in acts and duoc_sua_chung() and nut("➕ Thêm cột thiếu vào cấu hình & map lại", "them_cot"):
            L = ds()["loi"]
            cot = L[L["_loai"] == "cot"]
            for cate, ma in cot[["Ngành", "Mã cột"]].drop_duplicates().itertuples(index=False):
                o = ss.cau_hinh.setdefault(cate, {"ten": "", "cot": [], "ten_cot": {}})
                if ma not in o["cot"]:
                    o["cot"].append(ma)
            luu(["shared:cau_hinh"], f"QC: thêm {len(cot)} cột thiếu vào cấu hình")
            chay_map_ui()
            st.rerun()


# ----------------------------------------------------------------------------
# VÙNG XEM & SỬA RIÊNG TỪNG LỖI (bấm 1 dòng ở bảng QC)
# ----------------------------------------------------------------------------
def khu_sua_loi(kind: str, k: dict) -> None:
    if kind == "kiem_chung":
        st.info("👉 Xem và xử lý chi tiết ở tab **✅ Kiểm chứng SKU ↔ DATA SP** (ngay bên cạnh tab này).")
        return
    if kind == "nhat_quan":
        hien_nhat_quan(nq(), "qc")
        return
    if kind == "o_ten":
        canh_bao_o_ten()
        ds_ = o_ten_nhom()
        df_ = pd.DataFrame([{"Duyệt": False, "Ngành": x["cate"], "Cột PIM": x["ma"],
                             "Thuộc tính CMS": "; ".join(f"{n} ({p})" for p, n in x["props"].items()),
                             "Số SKU": len(x["skus"]), "Trùng tên cột đã map": x.get("trung_ten", "")} for x in ds_])
        ed = st.data_editor(df_, hide_index=True, key=f"ed_o_ten_{ss.get('ver_map', 0)}", width="stretch",
                            disabled=[c for c in df_.columns if c != "Duyệt"])
        if st.button("✅ Duyệt các cột đã tick (cho vào file import)", key="btn_duyet_ten"):
            moi = [f"{r['Ngành']}\t{r['Cột PIM']}" for _, r in ed.iterrows() if r["Duyệt"]]
            if moi:
                ss.duyet_ten = sorted(set(ss.get("duyet_ten", [])) | set(moi))
                luu(["settings"], f"Duyệt {len(moi)} cột điền theo tên")
                st.rerun()
        return
    if kind in ("thieu_model", "thieu_cate"):
        _sua_import_thieu("model_code" if kind == "thieu_model" else "category_code")
    elif kind == "xin_data":
        x = file_xin_data()
        if len(x["sku"]):
            st.dataframe(x["sku"].head(1000), hide_index=True, height=360)
        nut_xin_data("sl_xd")
    elif kind == "cot_thieu":
        L = ds()["loi"]
        cot = L[L["_loai"] == "cot"] if len(L) else L
        if not len(cot):
            st.success("✔ Không còn cột thiếu.")
            return
        st.dataframe(cot.drop(columns=[c for c in cot.columns if c.startswith("_")]), hide_index=True, height=320)
        n = cot[["Ngành", "Mã cột"]].drop_duplicates().shape[0]
        if st.button(f"➕ Thêm {n} cột thiếu vào CẤU HÌNH & map lại", type="primary", disabled=not duoc_sua_chung()):
            for cate, ma in cot[["Ngành", "Mã cột"]].drop_duplicates().itertuples(index=False):
                o = ss.cau_hinh.setdefault(cate, {"ten": "", "cot": [], "ten_cot": {}})
                if ma not in o["cot"]:
                    o["cot"].append(ma)
            luu(["shared:cau_hinh"], f"Thêm {n} cột thiếu vào cấu hình")
            chay_map_ui()
            st.rerun()
        if not duoc_sua_chung():
            st.caption("Chỉ admin sửa cấu hình dùng chung.")
    elif kind == "filter_khong_khop":
        st.info("👉 Sửa ở tab **🧾 Đối soát CMS → kết quả** → mục **🔁 Quy đổi FILTER** (gõ giá trị PIM tương ứng).")
    elif kind == "filter_chu":
        _sua_filter_chu()
    elif kind == "khac_spec":
        v = khac_df(k)
        v = v[v.trang_thai == C.TRANG_THAI_KHAC] if len(v) else v
        _loc_roi_grid(v, "sl_khac")
    elif kind == "tool_trong":
        v = khac_df(k)
        v = v[v.trang_thai == C.TRANG_THAI_TOOL_TRONG] if len(v) else v
        if len(v) and st.button("✨ Lấy PIM cũ cho tất cả ô này", key="sl_tt_all"):
            for r in v.itertuples(index=False):
                if r.pim_cu:
                    dat_sua(r.cate, r.sku, r.ma, r.pim_cu, r.goc)
            bump()
            luu(["settings"], "Lấy PIM cũ cho ô tool trống")
            st.rerun()
        _loc_roi_grid(v, "sl_tt")
    elif kind == "kich_thuoc":
        _sua_kich_thuoc(k)
    elif kind == "rong":
        st.info("👉 Chọn cách xử lý ở tab **🚫 Không / Đang cập nhật** (để trống / giữ / thay bằng…).")
    elif kind == "ttm":
        _sua_ttm_nhanh()
    elif kind == "sku_bat_buoc":
        h = dht()
        v = h["sku"][~h["sku"]["Đủ bắt buộc"]] if len(h["sku"]) else h["sku"]
        st.dataframe(v, hide_index=True, height=360)
        st.caption("Cột bắt buộc đặt ở ⚙️ Cấu hình & mapping → ✅ Quy tắc kiểm tra.")
    elif kind == "vi_pham":
        h = dht()
        v = h["vi_pham"][h["vi_pham"]["Mức"] == C.MUC_LOI] if len(h["vi_pham"]) else h["vi_pham"]
        st.dataframe(v, hide_index=True, height=360)
    elif kind == "model_lap":
        mt = C.model_trung(ss.bang)
        st.dataframe(pd.DataFrame({"model_code lặp": mt}), hide_index=True, height=320)
        st.caption("Thường do thiếu variant_code ở IMPORT. Kiểm tra lại danh sách SKU (👀 Xem dữ liệu đã nạp).")
    elif kind == "chua_map":
        cm = pd.DataFrame(ss.meta.get("chua_map", []))
        st.dataframe(cm, hide_index=True, height=360)
        st.caption("Thêm mapping ở ⚙️ Cấu hình & mapping → 🧬 Mapping TSKT / 🧮 Mapping FILTER.")
    elif kind == "khong_co_pim":
        st.info("SKU mới chưa có trong file PIM cũ nên không đối chiếu được — bình thường. Nạp file export PIM "
                "cũ (trang ①) nếu muốn đối chiếu.")
    else:
        st.info("Mục này không có bảng sửa riêng.")


def khac_df(k: dict) -> pd.DataFrame:
    return pd.DataFrame(k["khac"]) if k.get("khac") else pd.DataFrame(
        columns=["sku", "model", "cate", "ma", "ten", "pim_cu", "tool_moi", "goc", "trang_thai", "da_sua"])


def _loc_roi_grid(v: pd.DataFrame, key: str) -> None:
    """① Bảng TẤT CẢ mã TSKT (50 hay 300 mã đều hiện đủ, cuộn được) để tick & áp cho cả mã; ② nút XEM bự → bảng chi tiết."""
    if not len(v):
        st.success("✔ Không còn ô nào thuộc nhóm này.")
        return
    gm = (v.groupby("ma").agg(so=("sku", "size"), ten=("ten", "first")).reset_index()
          .sort_values(["so", "ma"], ascending=[False, True]).reset_index(drop=True))
    st.markdown(f"**{len(gm):,} mã TSKT · {len(v):,} ô** — tick mã cần xử lý:")
    tat = st.checkbox("Chọn tất cả mã", key=f"{key}_all")
    dfm = pd.DataFrame({"Chọn": tat, "Mã TSKT": gm.ma, "Tên": gm.ten, "Số ô": gm.so})
    em = st.data_editor(dfm, hide_index=True, width="stretch", height=min(320, 60 + 35 * len(dfm)),
                        key=f"{key}_ma_{ss.ver}_{int(tat)}", disabled=["Mã TSKT", "Tên", "Số ô"],
                        column_config={"Chọn": st.column_config.CheckboxColumn(width="small")})
    ma_chon = list(em.loc[em["Chọn"], "Mã TSKT"])
    n_ap = int(v.ma.isin(ma_chon).sum())
    if st.button(f"✨ Dùng giá trị PIM cũ cho {len(ma_chon)} mã đã tick ({n_ap:,} ô)", type="primary",
                 disabled=not ma_chon, key=f"{key}_ap_ma"):
        for r in v[v.ma.isin(ma_chon)].itertuples(index=False):
            if r.pim_cu:
                dat_sua(r.cate, r.sku, r.ma, r.pim_cu, r.goc)
        bump()
        luu(["settings"], f"Dùng PIM cũ cho {len(ma_chon)} mã TSKT ({n_ap} ô)")
        st.rerun()
    st.caption("Sửa ở đây = ghi vào «Ô sửa tay» và đi thẳng vào file import khi xuất (giống sửa trong bảng chi tiết).")
    xem = bool(ss.get(f"{key}_xem"))
    if st.button("👁 ĐÓNG BẢNG CHI TIẾT" if xem else "👁 XEM CHI TIẾT TỪNG Ô (sửa riêng từng ô)", width="stretch",
                 type="secondary" if xem else "primary", key=f"xem_{key}"):
        ss[f"{key}_xem"] = not xem
        st.rerun()
    if not xem:
        return
    tim = st.text_input("Lọc theo SKU / giá trị (gõ rồi Enter)", key=f"{key}_tim",
                        placeholder="🔎 gõ SKU hoặc giá trị…")
    d = v[v.ma.isin(ma_chon)] if ma_chon else v
    if tim:
        t = tim.lower()
        d = d[d.sku.str.lower().str.contains(t, regex=False) | d.tool_moi.str.lower().str.contains(t, regex=False)
              | d.pim_cu.str.lower().str.contains(t, regex=False)]
    st.caption(f"{len(d):,} ô" + (" — hiện 1.500 ô đầu, lọc thêm để xem hết" if len(d) > 1500 else "") +
               ("" if ma_chon else " (chưa tick mã nào → hiện tất cả mã)"))
    _grid_khac(d.head(1500), key)


def _sua_import_thieu(cot: str) -> None:
    ten_cot = "Mã model" if cot == "model_code" else "Mã danh mục PIM"
    imp = ss["import"]
    thieu = imp[imp[cot].fillna("").astype(str).str.strip() == ""]
    if not len(thieu):
        st.success(f"✔ Mọi SKU đều có {ten_cot}.")
        return
    st.caption(f"Điền trực tiếp **{ten_cot}** cho {len(thieu):,} SKU bên dưới rồi bấm Lưu. (Sửa cả IMPORT ở 👀 Xem dữ liệu.)")
    hien = thieu[["sku", "model_code", "variant_code", "category_code"]].reset_index().rename(
        columns={"index": "_idx", "sku": "SKU", "model_code": "Mã model", "variant_code": "Mã biến thể",
                 "category_code": "Mã danh mục PIM"}).head(2000)
    ed = st.data_editor(hien, hide_index=True, height=min(480, 90 + 35 * len(hien)), key=f"ed_thieu_{cot}_{ss.ver}",
                        disabled=["_idx", "SKU", "Mã biến thể"] + (["Mã danh mục PIM"] if cot == "model_code"
                                                                   else ["Mã model"]),
                        column_config={"_idx": None})
    if st.button("💾 Lưu & map lại", type="primary", key=f"luu_thieu_{cot}"):
        for _, r in ed.iterrows():
            ss["import"].at[r["_idx"], cot] = C.chuan_hoa_key(r[ten_cot] if cot == "model_code"
                                                              else r["Mã danh mục PIM"])
        bump()
        luu(["import"], f"Điền {ten_cot} cho SKU thiếu")
        chay_map_ui()
        st.rerun()


def _sua_filter_chu() -> None:
    rows = []
    for cate, b in ss.bang.items():
        for code in C.cot_tt(b):
            if not C.la_cot_filter(code):
                continue
            for r in b["rows"]:
                v0, _ = C.bien_doi_o(cate, r["sku"], code, r["vals"].get(code, ""), ss.sua, ss.dv, ss.rong)
                if v0 and re.search(r"[A-Za-zÀ-ỹ]", v0):
                    rows.append({"sku": r["sku"], "cate": cate, "ma": code, "ten": b["ten"].get(code, ""),
                                 "pim_cu": "", "tool_moi": v0, "goc": r["vals"].get(code, ""),
                                 "trang_thai": "FILTER có chữ", "da_sua": False})
    v = pd.DataFrame(rows)
    if not len(v):
        st.success("✔ Không còn ô FILTER nào chứa chữ.")
        return
    st.caption("Các ô FILTER còn chứa chữ (PIM chỉ nhận mã số). Sửa cột TOOL MỚI thành mã option, hoặc để trống. "
               "Xem giải nghĩa để chọn đúng.")
    _grid_khac(v.head(1000), "sl_fchu")


def _sua_ttm_nhanh() -> None:
    g = ttm()
    v = g[(g.muc_do.isin([C.MUC_CAO, C.MUC_TB])) & (g.goi_y != "")] if len(g) else g
    if not len(v):
        st.success("✔ Không có điểm nghi sai nào có gợi ý.")
        return
    v = v.head(1000).reset_index(drop=True)
    df = pd.DataFrame({"_cate": v.cate, "Mức": v.muc_do, "Loại": v.loai, "SKU": v.sku, "Mã TSKT": v.ma,
                       "Tên": v.ten, "Giá trị hiện tại": v.gia_tri, "Gợi ý": v.goi_y, "Lý do": v.ly_do})
    _bang_ap_dung(df, "ed_sl_ttm")


def _sua_kich_thuoc(k: dict) -> None:
    kt = [x for x in k["don_vi_cot"] if x["la_kt"]]
    if not kt:
        st.info("Không có cột kích thước / khối lượng trong các ngành đã map.")
        return
    cates = sorted({x["cate"] for x in kt}, key=lambda c: -sum(y["tong"] for y in kt if y["cate"] == c))
    cate = st.selectbox("Ngành hàng", cates, format_func=lambda c: ss.bang[c]["title"] if c in ss.bang else c,
                        key="skt_cate")
    cols = [(x["code"], x["ten"], x["goi_y"]) for x in kt if x["cate"] == cate]
    # --- (A) đơn vị theo cột ---
    st.markdown("**① Đơn vị cho cột** (chỉ thêm vào ô SỐ TRƠN — đúng rule 66.py)")
    du = pd.DataFrame({"Mã cột": [c for c, _, _ in cols], "Tên": [t for _, t, _ in cols],
                       "Gợi ý": [g for _, _, g in cols],
                       "ĐƠN VỊ": [ss.dv.get((cate, c), "") if isinstance(ss.dv.get((cate, c), ""), str) else ""
                                  for c, _, _ in cols]})
    edu = st.data_editor(du, hide_index=True, key=f"skt_dv_{cate}_{ss.ver}", disabled=["Mã cột", "Tên", "Gợi ý"],
                         column_config={"ĐƠN VỊ": st.column_config.TextColumn(width="small",
                                        help="cm, mm, kg, g, inch, lít…")})
    cu = st.columns([1.5, 1.5, 3])
    if cu[0].button("✨ Điền đơn vị gợi ý", key=f"skt_gy_{cate}"):
        for c, _, g in cols:
            if g and not ss.dv.get((cate, c)):
                ss.dv[(cate, c)] = g
        bump()
        luu(["settings"], "Điền đơn vị gợi ý (kích thước)")
        st.rerun()
    if cu[1].button("▶ Áp đơn vị", type="primary", key=f"skt_apdv_{cate}"):
        for i, (c, _, _) in enumerate(cols):
            v = C.chuan_hoa_key(edu.at[i, "ĐƠN VỊ"] or "")
            if v:
                ss.dv[(cate, c)] = v
            else:
                ss.dv.pop((cate, c), None)
        bump()
        luu(["settings"], "Áp đơn vị kích thước")
        st.rerun()
    # --- (B) bảng SKU × dài/rộng/cao để sửa số bằng mắt ---
    st.markdown("**② Bảng SKU × Dài · Rộng · Cao…** (sửa trực tiếp giá trị, như Excel)")
    nhan, seen = [], {}
    for c, t, _ in cols:
        lb = (t or c).strip()
        if lb in seen or not lb:
            lb = f"{lb} · {c}" if lb else c
        seen[lb] = c
        nhan.append((c, lb))
    rows = []
    b = ss.bang[cate]
    for r in b["rows"]:
        d = {"SKU": r["sku"], "Model": r["model"]}
        co = False
        for c, lb in nhan:
            v0, _ = C.bien_doi_o(cate, r["sku"], c, r["vals"].get(c, ""), ss.sua, ss.dv, ss.rong)
            d[lb] = v0
            if v0:
                co = True
        if co:
            d["_sku"] = r["sku"]
            rows.append(d)
    if not rows:
        st.info("Chưa có SKU nào có giá trị kích thước.")
        return
    tim = st.text_input("🔎 Lọc SKU", key=f"skt_tim_{cate}")
    dfp = pd.DataFrame(rows)
    if tim:
        dfp = dfp[dfp["SKU"].str.contains(tim, case=False, regex=False)]
    dfp = dfp.head(2000).reset_index(drop=True)
    goc = {c: {r["sku"]: r["vals"].get(c, "") for r in b["rows"]} for c, _ in nhan}
    cot_sua = [lb for _, lb in nhan]
    ed = st.data_editor(dfp.drop(columns=["_sku"]), hide_index=True, key=f"skt_pivot_{cate}_{ss.ver}",
                        height=min(520, 90 + 35 * len(dfp)), disabled=["SKU", "Model"])
    st.caption(f"{len(dfp):,} SKU. Sửa ô Dài/Rộng/Cao rồi bấm Lưu. Ô để trống = xoá giá trị đó.")
    if st.button("💾 Lưu thay đổi kích thước & xem lại", type="primary", key=f"skt_luu_{cate}"):
        nmap = dict(nhan)  # lb không có; cần map lb->code
        lb2code = {lb: c for c, lb in nhan}
        m = 0
        for i in range(len(dfp)):
            sku = dfp.at[i, "SKU"]
            for lb in cot_sua:
                code = lb2code[lb]
                moi = C.chuan_hoa_key(ed.at[i, lb] or "")
                cu_v = C.chuan_hoa_key(dfp.at[i, lb] or "")
                if moi != cu_v:
                    dat_sua(cate, sku, code, moi, goc[code].get(sku, ""))
                    m += 1
        bump()
        luu(["settings"], f"Sửa {m} ô kích thước")
        st.rerun()


def tab_hoan_thien() -> None:
    d = dht()
    st.caption("Như Akeneo/Salsify: % ô có giá trị trên các cột thuộc tính của ngành (tính trên GIÁ TRỊ SẼ XUẤT), hạng "
               "A ≥ 90% · B ≥ 80% · C ≥ 70% · D ≥ 60% · E < 60%. Cột BẮT BUỘC và quy tắc kiểm tra do admin đặt ở "
               "⚙️ Cấu hình → ✅ Quy tắc kiểm tra.")
    if len(d["nganh"]):
        st.dataframe(d["nganh"], hide_index=True, column_config={
            "% đầy đủ TB": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.1f%%"),
            "% SKU sẵn sàng": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
    c = st.columns([1, 1, 2])
    hang = c[0].multiselect("Hạng", list("ABCDE"), default=["D", "E"], key="ht_hang")
    chi_thieu = c[1].checkbox("Chỉ SKU thiếu cột bắt buộc", key="ht_bb")
    v = d["sku"]
    if hang:
        v = v[v["Hạng"].isin(hang)]
    if chi_thieu:
        v = v[~v["Đủ bắt buộc"]]
    c[2].caption(f"{len(v):,} SKU (xếp từ ít thông số nhất)")
    st.dataframe(v.head(2000), hide_index=True, height=300, column_config={
        "% đầy đủ": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
    vp = d["vi_pham"]
    st.markdown(f"##### Vi phạm quy tắc kiểm tra ({len(vp):,})")
    if not len(vp):
        st.caption("✔ Không có vi phạm" + ("" if ss.get("quy_tac_kt") else " — chưa đặt quy tắc nào."))
        return
    muc = st.multiselect("Mức", [C.MUC_LOI, C.MUC_CANH_BAO], default=[C.MUC_LOI, C.MUC_CANH_BAO], key="vp_muc")
    st.dataframe(vp[vp["Mức"].isin(muc)].head(3000), hide_index=True, height=320)
    st.caption("Sửa ô sai ở tab ≠ Khác spec / 🔎 Theo SKU (sửa tay) hoặc 📏 Đơn vị hàng loạt; dữ liệu CMS sai thì gửi "
               "📮 Đề xuất sửa CMS.")


def tab_canh_bao(k: dict) -> None:
    cb = k["canh_bao"]
    if not len(cb):
        st.success("✔ Không có cảnh báo.")
        return
    loai = st.multiselect("Lọc loại", sorted(cb["Loại"].unique()))
    v = cb[cb["Loại"].isin(loai)] if loai else cb
    st.dataframe(v, hide_index=True, height=460,
                 column_config={"Chi tiết": st.column_config.TextColumn(width="large")})
    st.caption(f"{len(v):,} dòng. Thiếu model/category: sửa ở IMPORT (trang 📥). Đơn vị: tab 📏.")




# ============================================================================
# TAB: ĐỐI SOÁT CMS -> KẾT QUẢ
# ============================================================================
def _sku_rong() -> int:
    L = ds()["loi"]
    if not len(L):
        return 0
    return int(L.loc[(L["_loai"] == "it_thong_so") & (L["Mức"] == C.MUC_CAO), "Số SKU"].sum() +
               L.loc[L["_loai"] == "khong_data", "Số SKU"].sum())


def khu_quy_doi_filter(L: pd.DataFrame) -> None:
    f = L[(L["_loai"] == "filter") & (L["Mức"] != C.MUC_THAP)] if len(L) else L
    st.caption("Giá trị CMS của cột FILTER không trùng tên option nào trong DATA PIM → ô FILTER bị TRỐNG. Chọn option "
               "đúng cho từng giá trị rồi lưu: bảng quy đổi dùng chung, mọi lần map sau tự khớp.")
    if not len(f):
        st.success("✔ Không có giá trị FILTER nào bị rơi (ngoài các giá trị Không/Đang cập nhật).")
    chon = {}
    for i, (_, r) in enumerate(f.head(40).iterrows()):
        ma, val = r["Mã cột"], r["Giá trị CMS"]
        gy = r["_goi_y"] if isinstance(r.get("_goi_y"), (list, tuple)) else []
        ds_opt = ss.opt["opt_ds"].get(ma, [])
        nhan = {oc: f"{oc} — {tn}" for oc, tn in ds_opt}
        thu_tu = [""] + [a for a, _ in gy] + [oc for oc, _ in ds_opt if oc not in {a for a, _ in gy}]
        cc = st.columns([3, 4])
        cc[0].markdown(f"**{ma}** · “{val}” · {r['Số SKU']} SKU")
        chon[(ma, val)] = cc[1].selectbox("Option", thu_tu, key=f"qd_{i}_{ma}_{val}", label_visibility="collapsed",
                                          format_func=lambda o, nh=nhan, g={a for a, _ in gy}:
                                          "(chưa chọn)" if not o else ("⭐ " if o in g else "") + nh.get(o, o))
    nhan_nut = "💾 Lưu quy đổi & map lại" if ss.admin else "📮 Gửi đề xuất quy đổi (dùng ngay cho bạn) & map lại"
    if len(f) and st.button(nhan_nut, type="primary"):
        items = []
        for (ma, val), oc in chon.items():
            if oc:
                r = f[(f["Mã cột"] == ma) & (f["Giá trị CMS"] == val)].iloc[0]
                items.append(C.tao_de_xuat(ss.user, ss.ten, C.PV_QUY_DOI, r["Ngành"], ma, val, oc,
                                           ten_cot=ss.opt["ten_filter"].get(ma, ""), so_sku=int(r["Số SKU"]),
                                           sku_vd=str(r["SKU ví dụ"]).split(",")[0].strip(),
                                           ly_do="Giá trị CMS chưa có option trong DATA PIM"))
        if items:
            gui_de_xuat(items)
            st.rerun()
    if len(f) and not ss.admin:
        st.caption("Đề xuất áp NGAY cho workspace của bạn; admin duyệt thì mọi người cùng dùng (trang 📮 Đề xuất sửa CMS).")
    if ss.quy_doi:
        with st.expander(f"Bảng quy đổi hiện có ({len(ss.quy_doi)})"):
            qd = pd.DataFrame([{"Mã FILTER": k.split("\t")[0], "Giá trị CMS": k.split("\t")[1], "Option": v,
                                "Tên option": ss.opt["opt_ten"].get((k.split("\t")[0], v), "❓"), "Xoá": False}
                               for k, v in ss.quy_doi.items()])
            ed = st.data_editor(qd, hide_index=True, key=f"ed_qd_{ss.ver}", disabled=["Mã FILTER", "Giá trị CMS",
                                                                                    "Option", "Tên option"])
            if st.button("🗑 Xoá dòng đã tick", disabled=not ss.admin) and ed["Xoá"].any():
                for r in ed[ed["Xoá"]].itertuples(index=False):
                    ss.quy_doi.pop(f"{r[0]}\t{r[1]}", None)
                luu(["shared:quy_doi"], "Xoá quy đổi FILTER")
                chay_map_ui()
                st.rerun()
    return



def tab_doi_soat() -> None:
    d = ds()
    L, P = d["loi"], d["phu"]
    if not len(L) and not len(P):
        st.success("✔ Mọi giá trị CMS đều vào được kết quả.")
        return
    c = st.columns(4)
    for col, muc in zip(c, [C.MUC_CAO, C.MUC_TB, C.MUC_THAP]):
        x = L[L["Mức"] == muc] if len(L) else L
        col.metric(f"Mức {muc}", f"{len(x)} nhóm · {int(x['Số SKU'].sum()) if len(x) else 0:,} SKU-ô")
    with c[3]:
        nut_tai("📊 Tải báo cáo đối soát", lambda: C.xlsx_nhieu_sheet(
        {"LỖI": L.drop(columns=[x for x in L.columns if x.startswith("_")]) if len(L) else L, "ĐỘ PHỦ CỘT": P}),
        file_name=f"DOI_SOAT_{C.bay_gio()[:10]}.xlsx")
    sub = st.radio("Xem", ["❗ Danh sách lỗi", "🔁 Quy đổi FILTER (giá trị CMS → option)", "📊 Độ phủ từng cột"],
                   horizontal=True, label_visibility="collapsed", key="ds_sub")
    sua_chung = duoc_sua_chung()
    if sub.startswith("❗"):
        if not len(L):
            st.success("✔ Không có lỗi.")
            return
        cc = st.columns([2, 3])
        muc = cc[0].multiselect("Mức", [C.MUC_CAO, C.MUC_TB, C.MUC_THAP], default=[C.MUC_CAO, C.MUC_TB], key="ds_muc")
        loai = cc[1].multiselect("Loại", sorted(L["Loại"].unique()), key="ds_loai")
        v = L[L["Mức"].isin(muc)] if muc else L
        if loai:
            v = v[v["Loại"].isin(loai)]
        st.dataframe(v.drop(columns=[x for x in v.columns if x.startswith("_")]), hide_index=True, height=380,
                     column_config={"Gợi ý": st.column_config.TextColumn(width="large"),
                                    "Số SKU": st.column_config.NumberColumn(format="%d")})
        st.caption("Số SKU = số SKU bị ảnh hưởng. Mã ngành khác/thiếu cột: dùng nút bên dưới (admin). "
                   "FILTER không khớp: sang mục 🔁 Quy đổi FILTER.")
        cot = L[L["_loai"] == "cot"]
        nk = L[L["_loai"] == "nganh_khac"]
        b = st.columns(2)
        if len(cot) and b[0].button(f"➕ Thêm {cot[['Ngành', 'Mã cột']].drop_duplicates().shape[0]} cột thiếu vào CẤU HÌNH "
                                     "rồi map lại", disabled=not sua_chung):
            for cate, ma in cot[["Ngành", "Mã cột"]].drop_duplicates().itertuples(index=False):
                o = ss.cau_hinh.setdefault(cate, {"ten": "", "cot": [], "ten_cot": {}})
                if ma not in o["cot"]:
                    o["cot"].append(ma)
            luu(["shared:cau_hinh"], f"Thêm {len(cot)} cột thiếu vào cấu hình (từ đối soát)")
            chay_map_ui()
            st.rerun()
        if len(nk) and b[1].button(f"📋 Copy {len(nk)} mapping từ ngành khác rồi map lại", disabled=not sua_chung):
            them_t, them_f = [], []
            for _, r in nk.iterrows():
                lo, ma = r["_copy"] if isinstance(r.get("_copy"), (tuple, list)) else ("tskt", r["Mã cột"])
                row = [r["Ngành"], ss.cau_hinh.get(r["Ngành"], {}).get("ten", ""), r["Mã thuộc tính CMS"],
                       r["Tên thuộc tính CMS"], ma]
                (them_f if lo == "filter" else them_t).append(row)
            if them_t:
                ss.map_tskt = pd.concat([ss.map_tskt, pd.DataFrame([x + [""] for x in them_t], columns=C.COT_MAP_TSKT)],
                                        ignore_index=True).drop_duplicates(["cate", "prop_id"], keep="last")
            if them_f:
                ss.map_filter = pd.concat([ss.map_filter, pd.DataFrame(them_f, columns=C.COT_MAP_FILTER)],
                                          ignore_index=True).drop_duplicates(["cate", "prop_id"], keep="last")
            luu(["shared:map_tskt", "shared:map_filter"], f"Copy {len(nk)} mapping từ ngành khác")
            chay_map_ui()
            st.rerun()
        if not sua_chung and (len(cot) or len(nk)):
            st.caption("Chỉ admin được sửa cấu hình/mapping dùng chung.")
        return
    if sub.startswith("🔁"):
        khu_quy_doi_filter(L)
        return
    if not len(P):
        return
    chi_thieu = st.checkbox("Chỉ cột chưa phủ hết (< 100%)", value=True)
    v = P[P["% tool"] < 100] if chi_thieu else P
    st.dataframe(v, hide_index=True, height=440,
                 column_config={"% tool": st.column_config.ProgressColumn("% SKU có giá trị", min_value=0, max_value=100,
                                                                          format="%.0f%%")})
    st.caption("Cột 0% + 'Không có mapping nào trỏ tới' = thiếu mapping (dữ liệu CMS có nhưng không vào được). "
               "'PIM cũ có' > 0 mà tool 0% = import có thể làm mất dữ liệu trên web.")

# ============================================================================
# TAB: GỢI Ý THÔNG MINH & AI
# ============================================================================
def ngu_canh_lo(k: dict) -> str:
    s = k["stat"]
    g = ttm()
    dong = [f"Workspace: {ss.ws}. Ngành hàng: " + ", ".join(f"{b['title']} ({len(b['rows'])} SKU, {len(b['attr'])} cột)"
                                                       for b in ss.bang.values())]
    dong.append("Số liệu kiểm tra: " + ", ".join(f"{a}={b}" for a, b in s.items()))
    cb = k["canh_bao"]
    if len(cb):
        dong.append("Cảnh báo theo loại: " + ", ".join(f"{a}: {b}" for a, b in cb["Loại"].value_counts().items()))
    if len(g):
        dong.append("Kiểm tra thông minh theo loại/mức: " + ", ".join(
            f"{a}/{m}: {n}" for (a, m), n in g.groupby(["loai", "muc_do"]).size().items()))
        dong.append("Ví dụ ô nghi sai: " + "; ".join(
            f"{r.ma}='{r.gia_tri}'→'{r.goi_y}' ({r.ly_do})" for r in g[g.muc_do != C.MUC_THAP].head(15).itertuples()))
    L = ds()["loi"]
    if len(L):
        dong.append("Đối soát CMS→kết quả: " + ", ".join(
            f"{a}/{m}: {n} nhóm, {sk} SKU" for (a, m), (n, sk) in
            L.groupby(["Loại", "Mức"])["Số SKU"].agg(["count", "sum"]).iterrows()))
    cm = ss.meta.get("chua_map", [])
    if cm:
        dong.append(f"Thuộc tính CMS chưa map: {len(cm)} (vd: " + ", ".join(x['prop_name'] for x in cm[:10]) + ")")
    dong.append(f"Sửa tay: {len(ss.sua)} ô; đơn vị: {len(ss.dv)} cột; quy tắc Không/Đang cập nhật: {len(ss.rong)}.")
    return "\n".join(dong)


def _bang_ap_dung(df: pd.DataFrame, key: str, cot_goi_y: str = "Gợi ý") -> None:
    """Bảng có cột 'Áp dụng' + gợi ý sửa được -> nút áp dụng thành sửa tay."""
    # Lọc theo MÃ (liệt kê đủ 50/300 mã kèm số ô) + gõ tìm; "áp dụng tất cả" = tất cả dòng ĐANG HIỆN
    if "Mã TSKT" in df.columns and len(df):
        dm = df["Mã TSKT"].value_counts()
        opts = [f"Tất cả {len(dm)} mã ({len(df):,} ô)"] + [f"{m} — {n:,} ô" for m, n in dm.items()]
        cf = st.columns([3, 2])
        pick = cf[0].selectbox("Lọc theo mã TSKT", opts, key=f"{key}_ma_{ss.ver}")
        tim = cf[1].text_input("Lọc SKU / giá trị (gõ rồi Enter)", key=f"{key}_tim", placeholder="🔎 gõ để lọc…")
        if pick != opts[0]:
            df = df[df["Mã TSKT"] == pick.rsplit(" — ", 1)[0]]
        if tim:
            t = tim.lower()
            mk = pd.Series(False, index=df.index)
            for cc_ in ("SKU", "Giá trị hiện tại", "Gợi ý"):
                if cc_ in df.columns:
                    mk |= df[cc_].astype(str).str.lower().str.contains(t, regex=False)
            df = df[mk]
        df = df.reset_index(drop=True)
        st.caption(f"Đang hiện {len(df):,} dòng.")
    hien = df.copy()
    hien.insert(0, "Áp dụng", False)
    ed = st.data_editor(hien, hide_index=True, height=400, key=f"{key}_{ss.ver}_{len(df)}",
                        disabled=[c for c in hien.columns if c not in ("Áp dụng", cot_goi_y)],
                        column_config={"Áp dụng": st.column_config.CheckboxColumn(width="small"), "_cate": None,
                                       cot_goi_y: st.column_config.TextColumn(width="medium"),
                                       "AI": st.column_config.TextColumn("🤖 AI nhận xét", width="medium")})
    c = st.columns([1, 1, 3])
    tat_ca = c[1].button(f"✔ Áp dụng TẤT CẢ {len(df):,} dòng đang hiện", key=f"{key}_all")
    if c[0].button("✔ Áp dụng dòng đã tick", type="primary", key=f"{key}_ap") or tat_ca:
        n = 0
        for i in range(len(ed)):
            gy = str(ed.at[i, cot_goi_y] or "").strip()
            if (tat_ca or ed.at[i, "Áp dụng"]) and gy:
                cate, sku, ma = df.at[i, "_cate"], df.at[i, "SKU"], df.at[i, "Mã TSKT"]
                b = ss.bang.get(cate, {})
                r = next((x for x in b.get("rows", []) if x["sku"] == sku), None)
                dat_sua(cate, sku, ma, gy, r["vals"].get(ma, "") if r else "")
                n += 1
        if n:
            bump()
            luu(["settings"], f"Áp dụng {n} gợi ý")
            st.rerun()
        else:
            st.warning("Chưa có dòng nào được tick / có gợi ý.")


def tab_ai_ra_soat(k: dict) -> None:
    """AI đọc tổng thể lô: cảnh báo, nghi sai, đối soát, mapping thiếu — tự suy luận ra điểm phi logic.
    Chạy 1 lần (hoặc khi bấm chạy lại). Kết quả lưu theo ver_map; thao tác khác không gọi lại AI."""
    ai = tao_ai()
    st.caption("AI đọc **toàn bộ** số liệu lô (cảnh báo, nghi sai, đối soát, mapping, ví dụ ô bất thường) rồi đưa ra "
               "nhận xét logic + đề xuất xử lý theo thứ tự ưu tiên. Chạy 1 phát; không gọi lại trừ khi bạn bấm làm lại.")
    if not ai.co_san:
        st.warning("Chưa cấu hình AI. Vào 👥 Quản trị → 🔑 Lưu API key AI (hoặc đặt trong Secrets).")
        return
    key_cache = f"ra_soat_{ss.get('ver_map', 0)}"
    cached = ss.get(key_cache)
    c = st.columns([1.4, 1, 3])
    chay = c[0].button("🚀 Rà soát ngay" if not cached else "🔄 Rà soát lại", type="primary", key="rs_run")
    if cached:
        c[1].caption(f"Lúc chạy: {cached.get('luc', '')}")
    if chay:
        with st.spinner(f"AI đang rà soát ({ai.mo_ta})… 10–30 giây"):
            try:
                nc = ngu_canh_lo(k)
                mau_o = []
                for cate, b in list(ss.bang.items())[:3]:
                    for r in b["rows"][:40]:
                        for c_ in C.cot_tt(b)[:6]:
                            v = r["vals"].get(c_, "")
                            if v:
                                mau_o.append(f"{cate}·{c_}·{r['sku']}: {str(v)[:60]}")
                            if len(mau_o) >= 80:
                                break
                        if len(mau_o) >= 80:
                            break
                    if len(mau_o) >= 80:
                        break
                prompt = ("Bạn là chuyên gia QC dữ liệu PIM. Hãy đọc kỹ số liệu lô dưới đây và chỉ ra "
                          "**các điểm phi logic / rủi ro / bất thường** mà con người có thể bỏ sót. "
                          "Trả về markdown theo cấu trúc:\n"
                          "## 1. Điểm phi logic (ưu tiên)\n- …\n"
                          "## 2. Nghi vấn dữ liệu (cần kiểm tra tay)\n- …\n"
                          "## 3. Thứ tự xử lý đề xuất\n1. …\n"
                          "## 4. Tối ưu quy trình\n- …\n\n"
                          "Yêu cầu: ngắn gọn, cụ thể, có dẫn chứng mã TSKT/SKU khi có. "
                          "Không bịa số liệu. Nếu dữ liệu ổn, nói rõ.\n\n"
                          "Mẫu ô dữ liệu (cate·ma·sku: giá trị):\n" + "\n".join(mau_o))
                tl = AIH.hoi(ai, prompt, nc, [])
                ss[key_cache] = {"ket_qua": tl, "luc": C.bay_gio(), "model": ai.mo_ta}
                st.rerun()
            except AIH.LoiAI as e:
                st.error(f"Lỗi AI: {e}")
                return
    if cached:
        st.markdown("---")
        st.markdown(cached["ket_qua"])
        st.caption(f"Model: {cached.get('model', '')} · {cached.get('luc', '')}")
    else:
        st.info("Bấm **🚀 Rà soát ngay** để AI đọc toàn bộ lô và đưa ra nhận xét.")


def tab_qc_nguoc(k: dict) -> None:
    """QC ngược: so với THƯ VIỆN AI (gom từ các lô ĐÃ XUẤT thành công) của chính ngành đó."""
    hoc = ss.get("ai_hoc") or {}
    co_hoc = {c: len(v) for c, v in hoc.items() if v}
    st.caption("So giá trị ô hiện tại với **top-20 giá trị thường gặp** của mỗi (ngành · cột) đã học từ các lô "
               "**đã xuất file import** trước đây. Ô có giá trị **chưa từng gặp** sẽ bị flag để bạn kiểm tra. "
               "Thư viện tự cập nhật mỗi lần có ai bấm 📤 Tạo file import.")
    if not hoc:
        st.info("Thư viện AI đang trống — chưa có lô nào xuất file được dùng để học. Khi bạn (hoặc thành viên) "
                "bấm 📤 Tạo file import ở trang Xuất, dữ liệu lô đó sẽ được dùng làm mẫu cho lần sau.")
        return
    st.caption("Thư viện có: " + " · ".join(f"**{c}** ({n} cột)" for c, n in sorted(co_hoc.items())))
    kq = qc_nguoc_tu_ai_hoc()
    if not kq:
        st.success("✔ Không phát hiện ô nào bất thường so với thư viện AI.")
        return
    import pandas as _pd
    df = _pd.DataFrame(kq)
    c = st.columns([2, 3, 2])
    cate_cs = c[0].multiselect("Lọc ngành", sorted(df.cate.unique()), key="qcng_cate")
    ma_cs = c[1].multiselect("Lọc mã TSKT", sorted(df.ma.unique()), key="qcng_ma")
    tim = c[2].text_input("Tìm (SKU / giá trị)", key="qcng_tim")
    v = df
    if cate_cs:
        v = v[v.cate.isin(cate_cs)]
    if ma_cs:
        v = v[v.ma.isin(ma_cs)]
    if tim:
        t = tim.lower()
        v = v[v.sku.str.lower().str.contains(t, regex=False) | v.gia_tri_hien.str.lower().str.contains(t, regex=False)]
    st.caption(f"**{len(v):,}** ô cần xem (tổng {len(df):,}). Giá trị *không bịa* — chỉ so với các giá trị đã có trong thư viện.")
    hien = _pd.DataFrame({"NH": v.cate, "SKU": v.sku, "Mã TSKT": v.ma, "Tên": v.ten,
                          "Giá trị hiện tại": v.gia_tri_hien, "Top thường gặp": v.goi_y_top,
                          "Lý do": v.ly_do})
    st.dataframe(hien.head(1000), hide_index=True, height=min(520, 60 + 35 * min(len(hien), 15)),
                 column_config={"Lý do": st.column_config.TextColumn(width="large")})
    st.caption("Cách xử lý: ô bất thường mà bạn vẫn muốn giữ → kệ nó. Giá trị sai → sửa ở tab ≠ Khác spec PIM "
               "hoặc 🔎 Theo SKU. Bạn xuất file lần này cũng sẽ góp mẫu cho thư viện.")


def tab_ai(k: dict) -> None:
    con = st.radio("Chọn", ["🧠 Kiểm tra thông minh (miễn phí, không cần AI)", "🔍 Rà soát toàn bộ (AI, 1 phát)",
                             "📚 QC ngược (học từ file đã duyệt)", "🤖 AI rà từng SKU", "💬 Hỏi AI"],
                   horizontal=True, label_visibility="collapsed", key="ai_che_do")
    if con.startswith("🔍"):
        tab_ai_ra_soat(k)
        return
    if con.startswith("📚"):
        tab_qc_nguoc(k)
        return
    if con.startswith("🧠"):
        g = ttm()
        if not len(g):
            st.success("✔ Không phát hiện điểm nghi sai nào.")
            return
        c = st.columns([2, 3, 2])
        muc = c[0].multiselect("Mức", [C.MUC_CAO, C.MUC_TB, C.MUC_THAP], default=[C.MUC_CAO, C.MUC_TB])
        loai = c[1].multiselect("Loại", sorted(g.loai.unique()))
        tim = c[2].text_input("Tìm", key="ttm_tim")
        v = g[g.muc_do.isin(muc)] if muc else g
        if loai:
            v = v[v.loai.isin(loai)]
        if tim:
            t = tim.lower()
            v = v[v.sku.str.lower().str.contains(t, regex=False) | v.ma.str.lower().str.contains(t, regex=False)
                  | v.gia_tri.str.lower().str.contains(t, regex=False)]
        v = v.head(1500).reset_index(drop=True)
        st.caption(f"{len(v):,} điểm (tổng {len(g):,}; ẩn mức Thấp theo mặc định). Dò: giá trị bất thường so với cả cột, "
                   "lẫn đơn vị, cùng chữ nhiều kiểu viết, lỗi gõ. Gợi ý có thể sửa trước khi áp dụng.")
        df = pd.DataFrame({"_cate": v.cate, "Mức": v.muc_do, "Loại": v.loai, "SKU": v.sku, "Mã TSKT": v.ma, "Tên": v.ten,
                           "Giá trị hiện tại": v.gia_tri, "Gợi ý": v.goi_y, "Lý do": v.ly_do})
        if "ai_xet" in ss and ss.get("ai_xet_ver") == ss.ver:
            ax = ss.ai_xet
            df["AI"] = [("✔ " if ax[i]["dong_y"] else "✖ ") + ax[i]["ly_do"] if i in ax else "" for i in range(len(df))]
            for i in range(len(df)):
                if i in ax and ax[i]["dong_y"] and ax[i]["de_xuat"]:
                    df.at[i, "Gợi ý"] = ax[i]["de_xuat"]
                elif i in ax and not ax[i]["dong_y"]:
                    df.at[i, "Gợi ý"] = ""  # AI cho là không sai -> không áp dụng hàng loạt
            cot = ["_cate", "Mức", "Loại", "SKU", "Mã TSKT", "Tên", "Giá trị hiện tại", "AI", "Gợi ý", "Lý do"]
            df = df[cot]
        _bang_ap_dung(df, "ed_ttm")
        ai = tao_ai()
        if st.button(f"🤖 Nhờ AI xem lại {min(len(df), 30)} dòng đầu đang lọc", disabled=not (ai.co_san and len(df))):
            mau: dict = {}
            items = []
            for i, r in df.head(30).iterrows():
                key = (r["_cate"], r["Mã TSKT"])
                if key not in mau:
                    vals = Counter(x["vals"].get(r["Mã TSKT"], "") for x in ss.bang[r["_cate"]]["rows"])
                    mau[key] = [a for a, _ in vals.most_common(6) if a][:5]
                items.append({"id": i, "cot": f"{r['Mã TSKT']} ({r['Tên']})", "gia_tri": r["Giá trị hiện tại"],
                              "goi_y_tool": r["Gợi ý"], "ly_do_tool": r["Lý do"], "mau_cot": mau[key]})
            with st.spinner(f"AI đang xem {len(items)} ô ({ai.mo_ta})…"):
                try:
                    kq_ai, _ = AIH.xet_o_danh_dau(ai, ", ".join(b["title"] for b in ss.bang.values()), items)
                    ss.ai_xet, ss.ai_xet_ver = kq_ai, ss.ver
                    st.rerun()
                except AIH.LoiAI as e:
                    st.error(str(e))
        if not ai.co_san:
            st.caption("💡 Bật AI để xem lại các dòng này: xem phần 🤖 AI rà từng SKU → cấu hình key.")
        return
    ai = tao_ai()
    with st.expander(f"⚙️ AI: {ai.mo_ta}" + ("" if ai.co_san else " — CHƯA BẬT"), expanded=not ai.co_san):
        st.caption("AI miễn phí, chỉ gửi thông số sản phẩm (không có dữ liệu cá nhân).")
        huong_dan_secrets_ai()
        st.caption("Hoặc dán key tạm cho riêng phiên này (thử key, đăng xuất là mất):")
        c = st.columns([1, 2, 2])
        prov = c[0].selectbox("Nhà cung cấp", list(AIH.PRESET), format_func=lambda p: AIH.PRESET[p]["ten"],
                              index=list(AIH.PRESET).index(ai.provider))
        key = c[1].text_input("API key (chỉ giữ trong phiên)", type="password", value=ss.get("ai_key_tam", ""))
        model = c[2].text_input("Model (để trống = mặc định)", value=ss.get("ai_model_tam", ""))
        st.caption(f"Lấy key miễn phí: {AIH.PRESET[prov]['lay_key']}")
        cc = st.columns(3)
        if cc[0].button("Dùng cấu hình này"):
            ss.ai_prov_tam, ss.ai_key_tam, ss.ai_model_tam = prov, key.strip(), model.strip()
            st.rerun()
        if cc[1].button("Liệt kê model", disabled=not ai.key):
            try:
                st.write(ai.ds_model())
            except AIH.LoiAI as e:
                st.error(str(e))
    if not ai.co_san:
        return
    if con.startswith("🤖"):
        g = ttm()
        nghi = list(dict.fromkeys(list(g[g.muc_do != C.MUC_THAP].sku) +
                                  [r["sku"] for r in k["khac"] if r["trang_thai"] == C.TRANG_THAI_KHAC]))
        tat_ca = [r["sku"] for b in ss.bang.values() for r in b["rows"]]
        ds = nghi + [x for x in tat_ca if x not in set(nghi)]
        sku = st.selectbox(f"SKU (ưu tiên {len(nghi)} SKU đang nghi sai lên đầu)", ds, key="ai_sku")
        x = k["chi_tiet"].get(sku)
        if not x:
            return
        st.caption(f"{ten_sp(sku) or '(không rõ tên)'} · {x['tab']}")
        if st.button("🤖 AI rà SKU này", type="primary"):
            dong = [{"ma": d["ma"], "ten": d["ten"], "tool_moi": d["tool_moi"], "pim_cu": d["pim_cu"]} for d in x["dong"]]
            with st.spinner(f"AI đang rà ({ai.mo_ta})…"):
                try:
                    gy, raw = AIH.goi_y_sku(ai, ss.bang[x["cate"]]["title"], sku, ten_sp(sku), dong)
                    ss.ai_sku_kq = {"sku": sku, "goi_y": gy, "raw": raw, "ver": ss.ver}
                except AIH.LoiAI as e:
                    st.error(str(e))
        r = ss.get("ai_sku_kq")
        if r and r["sku"] == sku:
            if not r["goi_y"]:
                st.success("✔ AI không thấy điểm sai nào ở SKU này.")
                with st.expander("Câu trả lời gốc của AI"):
                    st.text(r["raw"][:3000])
            else:
                hien_tai = {d["ma"]: d for d in x["dong"]}
                df = pd.DataFrame({"_cate": x["cate"], "Mức": [g_["muc_do"] for g_ in r["goi_y"]], "SKU": sku,
                                   "Mã TSKT": [g_["ma"] for g_ in r["goi_y"]],
                                   "Tên": [hien_tai.get(g_["ma"], {}).get("ten", "") for g_ in r["goi_y"]],
                                   "Giá trị hiện tại": [hien_tai.get(g_["ma"], {}).get("tool_moi", "") for g_ in r["goi_y"]],
                                   "Gợi ý": [g_["de_xuat"] for g_ in r["goi_y"]],
                                   "Lý do": [g_["ly_do"] for g_ in r["goi_y"]]})
                st.caption("AI chỉ đề xuất — kiểm tra lại với web/hãng trước khi áp dụng. Ô 'Gợi ý' trống = AI chỉ nhắc kiểm tra.")
                _bang_ap_dung(df, "ed_ai_sku")
        return
    # ---- Hỏi AI
    ss.setdefault("ai_chat", [])
    c = st.columns(3)
    mau = None
    if c[0].button("📋 Tóm tắt lỗi lô này & nên xử lý gì trước"):
        mau = "Tóm tắt các lỗi/cảnh báo của lô này và cho tôi thứ tự xử lý hợp lý nhất (ngắn gọn, theo bước)."
    if c[1].button("💡 Đề xuất thêm tính năng cho tool"):
        mau = ("Dựa trên số liệu lô này và các tính năng tool đang có, đề xuất 5–8 tính năng/cải tiến giúp map nhanh "
               "và chính xác hơn. Mỗi đề xuất: tên, làm gì, lợi ích, mức ưu tiên.")
    if c[2].button("🗑 Xoá hội thoại"):
        ss.ai_chat = []
        st.rerun()
    for m in ss.ai_chat:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
    hoi = st.chat_input("Hỏi AI về lô dữ liệu / cách xử lý / đề xuất…") or mau
    if hoi:
        with st.chat_message("user"):
            st.markdown(hoi)
        with st.chat_message("assistant"):
            with st.spinner("AI đang trả lời…"):
                try:
                    tl = AIH.hoi(ai, hoi, ngu_canh_lo(k), ss.ai_chat)
                except AIH.LoiAI as e:
                    tl = f"⚠️ {e}"
            st.markdown(tl)
        ss.ai_chat += [{"role": "user", "content": hoi}, {"role": "assistant", "content": tl}]

LOC_CAN_XEM = "Cần xem (khác + tool trống)"


def tab_khac(k: dict) -> None:
    if not k["co_doi_chieu"]:
        st.info("Chưa có spec PIM cũ để đối chiếu — nạp file export PIM ở trang 📥.")
    df = pd.DataFrame(k["khac"])
    if not len(df):
        st.success("✔ Không có ô nào khác spec PIM.")
        return
    c = st.columns([2, 2, 2])
    loai = c[0].selectbox("Loại", [LOC_CAN_XEM, "Tất cả", C.TRANG_THAI_KHAC, C.TRANG_THAI_TOOL_TRONG,
                                   C.TRANG_THAI_DON_VI, C.TRANG_THAI_PIM_TRONG, C.TRANG_THAI_BO_QUA, "✎ Đã sửa tay"])
    ma = c[1].selectbox("Mã TSKT", ["Tất cả"] + sorted(df.ma.unique()))
    tim = c[2].text_input("Tìm (SKU / giá trị)")
    m = pd.Series(True, index=df.index)
    if loai == LOC_CAN_XEM:
        m &= df.trang_thai.isin([C.TRANG_THAI_KHAC, C.TRANG_THAI_TOOL_TRONG]) | df.da_sua
    elif loai == "✎ Đã sửa tay":
        m &= df.da_sua
    elif loai != "Tất cả":
        m &= df.trang_thai == loai
    if ma != "Tất cả":
        m &= df.ma == ma
    if tim:
        t = tim.lower()
        m &= df.sku.str.lower().str.contains(t, regex=False) | df.pim_cu.str.lower().str.contains(t, regex=False) | \
            df.tool_moi.str.lower().str.contains(t, regex=False)
    v = df[m].head(1500).reset_index(drop=True)
    st.caption(f"{int(m.sum()):,} ô" + (" — hiện 1.500 ô đầu, lọc thêm để xem hết" if m.sum() > 1500 else "") +
               ". Sửa cột **TOOL MỚI** hoặc tick **Lấy PIM cũ**, rồi bấm **Áp dụng**.")
    if not len(v):
        return
    _grid_khac(v, "khac")


def _grid_khac(v: pd.DataFrame, key: str) -> None:
    """Bảng sửa & đối chiếu kiểu Excel: cột TOOL MỚI sửa tay / tick Lấy PIM cũ, có giải nghĩa FILTER để soát bằng mắt."""
    v = v.reset_index(drop=True)
    hien = pd.DataFrame({
        "SKU": v.sku, "NH": v.cate, "Mã TSKT": v.ma, "Tên": v.ten, "PIM cũ": v.pim_cu,
        "TOOL MỚI": v.tool_moi, "Giải nghĩa FILTER": [C.giai_nghia_filter(a, b, ss.opt) for a, b in zip(v.ma, v.tool_moi)],
        "Trạng thái": v.trang_thai, "Lấy PIM cũ": False, "✎": v.da_sua.map(lambda x: "✎" if x else "")})
    ed = st.data_editor(hien, hide_index=True, height=min(520, 90 + 35 * len(hien)), key=f"ed_{key}_{ss.ver}",
                        disabled=[c for c in hien.columns if c not in ("TOOL MỚI", "Lấy PIM cũ")],
                        column_config={"Lấy PIM cũ": st.column_config.CheckboxColumn(width="small"),
                                       "TOOL MỚI": st.column_config.TextColumn(width="medium")})
    if st.button("✔ Áp dụng thay đổi", type="primary", key=f"ap_{key}"):
        n = 0
        for i in range(len(v)):
            if ed.at[i, "Lấy PIM cũ"] and v.pim_cu[i]:
                dat_sua(v.cate[i], v.sku[i], v.ma[i], v.pim_cu[i], v.goc[i])
                n += 1
            elif ed.at[i, "TOOL MỚI"] != v.tool_moi[i]:
                dat_sua(v.cate[i], v.sku[i], v.ma[i], ed.at[i, "TOOL MỚI"] or "", v.goc[i])
                n += 1
        bump()
        luu(["settings"], f"Sửa tay {n} ô")
        st.rerun()


def tab_sku(k: dict) -> None:
    ct = k["chi_tiet"]
    if not ct:
        return
    c1, c2 = st.columns([2, 1])
    chi_khac = c2.checkbox("Chỉ SKU cần xem", value=True)
    ds = [s for s, v in ct.items() if not chi_khac or any(
        d["trang_thai"] in (C.TRANG_THAI_KHAC, C.TRANG_THAI_TOOL_TRONG) or d["da_sua"] for d in v["dong"])]
    if not ds:
        st.success("✔ Không còn SKU nào cần xem." if chi_khac else "Không có SKU.")
        return
    sku = c1.selectbox(f"SKU ({len(ds):,})", ds, key="sku_xem")
    x = ct[sku]
    st.caption(f"Model **{x['model'] or '(trống)'}** · biến thể {x['variant'] or '-'} · {x['tab']}"
               + ("" if x["co_pim"] else " · ⚠️ KHÔNG có trong file PIM"))
    tt_thu_tu = {C.TRANG_THAI_KHAC: 0, C.TRANG_THAI_TOOL_TRONG: 1, C.TRANG_THAI_DON_VI: 2, C.TRANG_THAI_PIM_TRONG: 3}
    dong = sorted(x["dong"], key=lambda d: (not d["da_sua"], tt_thu_tu.get(d["trang_thai"], 9), d["ma"]))
    hien = pd.DataFrame({"Mã TSKT": [d["ma"] for d in dong], "Tên": [d["ten"] for d in dong],
                         "PIM cũ": [d["pim_cu"] for d in dong], "TOOL MỚI": [d["tool_moi"] for d in dong],
                         "Giải nghĩa FILTER": [C.giai_nghia_filter(d["ma"], d["tool_moi"], ss.opt) for d in dong],
                         "Trạng thái": [d["trang_thai"] for d in dong], "Lấy PIM cũ": False,
                         "✎": ["✎" if d["da_sua"] else "" for d in dong]})
    ed = st.data_editor(hien, hide_index=True, height=420, key=f"ed_sku_{sku}_{ss.ver}",
                        disabled=[c for c in hien.columns if c not in ("TOOL MỚI", "Lấy PIM cũ")])
    if st.button("✔ Áp dụng cho SKU này", type="primary", key="ap_sku"):
        for i, d in enumerate(dong):
            if ed.at[i, "Lấy PIM cũ"] and d["pim_cu"]:
                dat_sua(x["cate"], sku, d["ma"], d["pim_cu"], d["goc"])
            elif ed.at[i, "TOOL MỚI"] != d["tool_moi"]:
                dat_sua(x["cate"], sku, d["ma"], ed.at[i, "TOOL MỚI"] or "", d["goc"])
        bump()
        luu(["settings"], f"Sửa tay SKU {sku}")
        st.rerun()
    st.caption("Sai do **dữ liệu CMS**? Sau khi sửa, gửi thành đề xuất ở trang 📮 Đề xuất sửa CMS để admin duyệt dùng chung.")
    # ---- giải nghĩa + chọn lại option FILTER
    ma_f = [d["ma"] for d in dong if C.la_cot_filter(d["ma"])]
    if not ma_f:
        return
    st.markdown("##### 🔎 Giải nghĩa & chọn lại FILTER (so với web)")
    mf = st.selectbox("Cột FILTER", ma_f, format_func=lambda m: f"{m} — {ss.opt['ten_filter'].get(m, '')}")
    d = next(dd for dd in dong if dd["ma"] == mf)

    def _ma(v):
        return [C.chuan_hoa_id(t.strip()) for t in re.split(r"[,|]", C.lam_sach_gia_tri_pim(v, mf)) if t.strip()]
    moi, cu = _ma(d["tool_moi"]), _ma(d["pim_cu"])
    c1, c2 = st.columns(2)
    for col, tieu, ds1, ds2 in ((c1, "TOOL MỚI (sẽ import)", moi, cu), (c2, "PIM CŨ (đang trên web)", cu, moi)):
        col.markdown(f"**{tieu}**")
        if not ds1:
            col.caption("(trống)")
        for m in ds1:
            ten = ss.opt["opt_ten"].get((mf, m))
            if ten is None:
                col.markdown(f":red[{m} = ❓ KHÔNG có trong DATA PIM]")
            else:
                col.markdown(f":{'green' if (m in ds2 or not ds2) else 'orange'}[{m} = {ten}]")
    ds_opt = ss.opt["opt_ds"].get(mf, [])
    dang = set(moi) | set(cu)
    ds_opt = sorted(ds_opt, key=lambda o: (o[0] not in dang, int(o[0]) if o[0].isdigit() else 10 ** 12))
    nhan = {oc: f"{oc} — {tn}" + ("  ← PIM cũ" if oc in cu else "") for oc, tn in ds_opt}
    chon = st.multiselect(f"Chọn lại option ({len(ds_opt)} option — gõ để tìm)", list(nhan),
                          default=[m for m in moi if m in nhan], format_func=nhan.get, key=f"ms_{sku}_{mf}")
    if st.button("✔ Dùng các option đã chọn", key=f"dung_{sku}_{mf}"):
        dat_sua(x["cate"], sku, mf, C.SEP_FILTER.join(chon), d["goc"])
        bump()
        luu(["settings"], f"Chọn lại FILTER {mf} cho {sku}")
        st.rerun()


DV_OPTIONS = ["", "cm", "mm", "m", "kg", "g", "inch", "lít", "W", "mAh", "V", "Hz", "dB"]
PV_NHAN = {"so": "Số trơn", "tung_phan": "Từng giá trị (9|10)", "tat_ca": "Có số / chữ+số"}
PV_NGUOC = {v: k_ for k_, v in PV_NHAN.items()}


def _buoc_bang(cate: str, code: str):
    """Bước TRƯỚC/SAU đặt từ bảng ① (đánh dấu nguon='bang') của 1 cột, hoặc None."""
    for x in ss.dv.get((cate, code, "bd")) or []:
        if isinstance(x, dict) and x.get("nguon") == "bang":
            return x
    return None


def _ap_truoc_sau(cate: str, code: str, truoc, sau, pham_vi) -> None:
    """Ghi TRƯỚC / SAU / PHẠM VI của 1 cột từ bảng ①.
    * Chỉ SAU + phạm vi «Số trơn» -> RULE CŨ (ss.dv[(ngành, mã)] = đơn vị), y như trước đây.
    * Có TRƯỚC hoặc phạm vi khác -> 1 bước biến đổi như mục ② (them_truoc / them_sau / ca_hai), đánh dấu nguon='bang'
      để lần sau thay đúng bước này, KHÔNG đụng các bước ② khác của cột."""
    tr, sa = C.chuan_hoa_key(truoc or ""), C.chuan_hoa_key(sau or "")
    pv = PV_NGUOC.get(pham_vi, "so")
    kk = (cate, code, "bd")
    khac = [x for x in (ss.dv.get(kk) or []) if not (isinstance(x, dict) and x.get("nguon") == "bang")]
    if tr or pv != "so":
        ss.dv.pop((cate, code), None)
        if tr or sa:
            kd = "ca_hai" if tr and sa else ("them_truoc" if tr else "them_sau")
            khac.append({"kieu": kd, "pham_vi": pv, "a": tr if kd != "them_sau" else sa,
                         "b": sa if kd == "ca_hai" else "", "he_so": "", "nguon": "bang"})
    elif sa:
        ss.dv[(cate, code)] = sa
    else:
        ss.dv.pop((cate, code), None)
    if khac:
        ss.dv[kk] = khac
    else:
        ss.dv.pop(kk, None)


@st.fragment
def tab_don_vi(k: dict) -> None:
    st.markdown("##### ① Đơn vị theo cột (đúng rule cũ 66.py: chỉ thêm vào ô SỐ TRƠN, không đụng FILTER)")
    ds = pd.DataFrame(k["don_vi_cot"])
    if len(ds):
        tat_ca = st.checkbox("Hiện tất cả cột TSKT (không chỉ cột kích thước/khối lượng)", key="dv_tat_ca")
        if not tat_ca:
            ds = ds[ds.la_kt | (ds.da_luu != "") | ds.apply(lambda r: bool(ss.dv.get((r.cate, r.code))), axis=1)]
        ds = ds.sort_values(["cate", "so_tron"], ascending=[True, False]).reset_index(drop=True)
        _tr_, _sa_, _pv_ = [], [], []
        for a, b in zip(ds.cate, ds.code):
            x_ = _buoc_bang(a, b)
            if x_:  # đã đặt TRƯỚC/SAU từ bảng này (kiểu giống ②) -> nạp lại đúng giá trị
                kd_, aa_, bb_ = x_.get("kieu"), x_.get("a", ""), x_.get("b", "")
                _tr_.append(aa_ if kd_ in ("ca_hai", "them_truoc") else "")
                _sa_.append(bb_ if kd_ == "ca_hai" else (aa_ if kd_ == "them_sau" else ""))
                _pv_.append(PV_NHAN.get(x_.get("pham_vi", "so"), PV_NHAN["so"]))
            else:  # rule cũ: chỉ ĐƠN VỊ (sau) cho ô số trơn
                _tr_.append("")
                _sa_.append(ss.dv.get((a, b), "") if isinstance(ss.dv.get((a, b), ""), str) else "")
                _pv_.append(PV_NHAN["so"])
        hien = pd.DataFrame({"NH": ds.cate, "Mã TSKT": ds.code, "Tên": ds.ten,
                             "Số trơn / có dữ liệu": [f"{a} / {b}" for a, b in zip(ds.so_tron, ds.tong)],
                             "Giá trị hiện tại (sau áp)": ds.vi_du, "Gợi ý": ds.goi_y,
                             "THÊM TRƯỚC": _tr_, "ĐƠN VỊ (SAU)": _sa_, "PHẠM VI": _pv_})
        c = st.columns([1.4, 1, 3])
        if c[0].button("✨ Điền gợi ý (cột còn số trơn, chưa chọn)"):
            for a, b, g, n in zip(ds.cate, ds.code, ds.goi_y, ds.so_tron):
                if n and g and not ss.dv.get((a, b)):
                    ss.dv[(a, b)] = g
            bump()
            luu(["settings"], "Điền đơn vị theo gợi ý")
            st.rerun()
        c[2].caption("Gõ chữ / đơn vị TUỲ Ý vào **THÊM TRƯỚC** (vd: Khoảng, Dài) và/hoặc **ĐƠN VỊ (SAU)** (cm, mm, kg, g, inch, "
                     "W, mAh, lít, giờ…). **PHẠM VI** giống mục ②: Số trơn (mặc định, rule cũ) · Từng giá trị (9|10) · "
                     "Có số/chữ+số (vd «Driver 40mm»). Ô chữ thuần (Không / Đang cập nhật) không bị thêm. "
                     "Bấm Áp để xem kết quả ngay.")
        ed = st.data_editor(hien, hide_index=True, height=min(460, 40 + 35 * len(hien)), key=f"ed_dv_{ss.ver}",
                            disabled=[x for x in hien.columns if x not in ("THÊM TRƯỚC", "ĐƠN VỊ (SAU)", "PHẠM VI")],
                            column_config={
                                "THÊM TRƯỚC": st.column_config.TextColumn(width="small", help="Chữ thêm vào TRƯỚC giá trị"),
                                "ĐƠN VỊ (SAU)": st.column_config.TextColumn(
                                    width="small", help="Thêm vào SAU giá trị, vd: " + ", ".join(DV_OPTIONS[1:])),
                                "PHẠM VI": st.column_config.SelectboxColumn(width="small", options=list(PV_NHAN.values()),
                                                                            required=True)})
        if st.button("▶ Áp đơn vị & xem lại", type="primary"):
            for i in range(len(ds)):
                _ap_truoc_sau(ds.cate[i], ds.code[i], ed.at[i, "THÊM TRƯỚC"], ed.at[i, "ĐƠN VỊ (SAU)"], ed.at[i, "PHẠM VI"])
            bump()
            luu(["settings"], "Áp đơn vị hàng loạt")
            st.rerun()
    st.divider()
    khu_bien_doi(k)


def _mo_ta_buoc(x: dict) -> str:
    a, b = x.get("a", ""), x.get("b", "")
    return f"{C.BD_KIEU.get(x.get('kieu'), x.get('kieu', ''))} [{a}{(' → ' + b) if b else ''}]"


def _bd_mo_bat() -> None:
    """Đã thao tác trong vùng ② -> giữ vùng này MỞ sau mỗi lần chạy lại (không tự gập/nhảy)."""
    ss._bd_mo = True


def _bd_rerun() -> None:
    ss._bd_mo = True
    st.rerun()


def khu_bien_doi(k: dict | None = None) -> None:
    """Vùng nâng cao — mặc định THU GỌN (không mất, bấm mở khi cần); đã thao tác thì giữ mở."""
    n = len([1 for k_, v in ss.dv.items() if len(k_) == 3 and v])
    with st.expander(f"② Biến đổi hàng loạt — nâng cao{f' · đang áp {n} cột' if n else ''} (bấm để mở)",
                     expanded=bool(ss.get("_bd_mo", False))):
        _khu_bien_doi_noi_dung(k)


def _khu_bien_doi_noi_dung(k: dict | None = None) -> None:
    st.caption("Áp sau đơn vị ở ①, trước khi xuất. Không bao giờ áp vào cột FILTER. Ô sửa tay và quy tắc "
               "Không/Đang cập nhật được ưu tiên hơn. **Chọn lại cột đã áp → form tự hiện đúng giá trị đã đặt để sửa.**")
    if bang_trong():
        return
    cot_het = sorted({(c, m) for c, b in ss.bang.items() for m in C.cot_tt(b) if not C.la_cot_filter(m)})
    ten = {(c, m): ss.bang[c]["ten"].get(m, "") for c, m in cot_het}
    # Cột "cần thay thế" (hiện mặc định): cột kích thước/khối lượng, cột còn ô số trơn chưa có đơn vị,
    # và cột ĐÃ có biến đổi. Muốn thấy hết mọi cột -> tick «Hiện tất cả cột».
    can = {(r["cate"], r["code"]) for r in (k or {}).get("don_vi_cot", []) if r.get("la_kt") or r.get("so_tron")}
    can |= {(c, m) for c, m in cot_het if ss.dv.get((c, m, "bd")) or ss.dv.get(("*", m, "bd"))}
    c1, c2 = st.columns([3, 2])
    tat_ca_cot = c2.checkbox("Hiện tất cả cột", key="bd_tat_ca_cot", on_change=_bd_mo_bat,
                             help="Bỏ tick: chỉ hiện cột cần thay thế (kích thước/khối lượng, còn ô số trơn, đã có biến đổi)")
    if tat_ca_cot or not can:
        cot = cot_het
    else:
        cot = [x for x in cot_het if x in can]
    # luôn giữ các cột ĐANG CHỌN trong danh sách (tránh bị mất lựa chọn khi đổi bộ lọc)
    for x in ss.get("bd_cot") or []:
        if x in cot_het and x not in cot:
            cot.append(x)
    cot = sorted(cot)
    if not tat_ca_cot and can:
        c1.caption(f"Đang hiện {len(cot)}/{len(cot_het)} cột cần thay thế — tick «Hiện tất cả cột» để xem hết.")
    chon = c1.multiselect("Cột áp dụng", cot, key="bd_cot", on_change=_bd_mo_bat,
                          format_func=lambda x: f"{x[0]} · {x[1]} — {ten.get(x, '')}")
    # --- cột đã có biến đổi? -> nạp lại đúng giá trị vào form (như lúc chưa áp hàng loạt) ---
    hien_co, buoc_sua = [], None
    if len(chon) == 1:
        cc0, m0 = chon[0]
        kk_ng, kk_all = (cc0, m0, "bd"), ("*", m0, "bd")
        kk_co = kk_ng if ss.dv.get(kk_ng) else (kk_all if ss.dv.get(kk_all) else None)
        hien_co = list(ss.dv.get(kk_co) or []) if kk_co else []
        if hien_co:
            nhan_b = [f"Bước {n}: {_mo_ta_buoc(x)}" for n, x in enumerate(hien_co, 1)] + ["➕ Thêm 1 bước mới"]
            ch = st.selectbox("Biến đổi đang áp cho cột này", nhan_b, key=f"bd_buoc_{cc0}_{m0}_{ss.ver}", on_change=_bd_mo_bat)
            n_ch = nhan_b.index(ch)
            buoc_sua = n_ch if n_ch < len(hien_co) else None
        ctx = (cc0, m0, buoc_sua, len(hien_co))
    else:
        kk_co, ctx = None, None
    if ctx != ss.get("_bd_ctx"):
        ss._bd_ctx = ctx
        if ctx is not None and buoc_sua is not None:
            x = hien_co[buoc_sua]
            ss.bd_kieu, ss.bd_pv = x.get("kieu", "them_sau"), x.get("pham_vi", "so")
            ss.bd_a, ss.bd_b, ss.bd_hs = x.get("a", ""), x.get("b", ""), x.get("he_so", "")
            ss.bd_moi_nganh = kk_co is not None and kk_co[0] == "*"
        elif ctx is not None and not hien_co:
            for k_ in ("bd_a", "bd_b", "bd_hs"):
                ss[k_] = ""
    moi_nganh = c2.checkbox("Áp cho mã cột này ở MỌI ngành", key="bd_moi_nganh", on_change=_bd_mo_bat,
                            help="Lưu theo mã cột (*) thay vì từng ngành — tiện cho cột dùng chung như mass_tskt_master")
    c = st.columns([2, 2, 1.2, 1.2, 1])
    kieu = c[0].selectbox("Kiểu", list(C.BD_KIEU), format_func=C.BD_KIEU.get, key="bd_kieu", on_change=_bd_mo_bat)
    pv = c[1].selectbox("Phạm vi", list(C.BD_PHAM_VI), format_func=C.BD_PHAM_VI.get, key="bd_pv", on_change=_bd_mo_bat,
                        index=2 if kieu in ("thay",) else 0)
    nhan_a = {"them_sau": "Chữ / đơn vị", "them_truoc": "Chữ", "ca_hai": "Chữ phía TRƯỚC", "doi_dv": "Từ đơn vị",
              "thay": "Tìm chữ", "lam_tron": "Số chữ số lẻ"}[kieu]
    a = c[2].text_input(nhan_a, key="bd_a", on_change=_bd_mo_bat, placeholder={"them_sau": "kg", "doi_dv": "mm", "lam_tron": "1",
                                                         "ca_hai": "Khoảng"}.get(kieu, ""))
    b = c[3].text_input({"doi_dv": "Sang đơn vị", "ca_hai": "Chữ phía SAU"}.get(kieu, "Thay bằng"), key="bd_b", on_change=_bd_mo_bat,
                        disabled=kieu not in ("doi_dv", "thay", "ca_hai"),
                        placeholder={"doi_dv": "cm", "ca_hai": "cm"}.get(kieu, ""))
    he_so = c[4].text_input("Hệ số", key="bd_hs", on_change=_bd_mo_bat, disabled=kieu != "doi_dv",
                            help="Để trống nếu là cặp quen thuộc (mm↔cm↔m, g↔kg, inch→cm, ml↔lít, W↔kW, mAh↔Ah, phút↔giờ)")
    buoc = {"kieu": kieu, "pham_vi": pv, "a": a, "b": b, "he_so": he_so}
    if chon and st.checkbox("🔎 Lọc xem giá trị trong cột đã chọn (ô nào có số / đơn vị, ô nào là chữ)", key="bd_loc_xem", on_change=_bd_mo_bat):
        # CHỈ ĐỌC: giá trị hiện tại (sau sửa tay/đơn vị ①, chưa biến đổi hàng loạt) gom theo giá trị
        dv_goc_x = {k_: v_ for k_, v_ in ss.dv.items()
                    if not (len(k_) == 3 and k_[2] == "bd" and any(k_[1] == m_ and k_[0] in (c_, "*") for c_, m_ in chon))}
        dem: dict = {}
        for cc, m in chon:
            for r in ss.bang[cc]["rows"]:
                v0, _ = C.bien_doi_o(cc, r["sku"], m, r["vals"].get(m, ""), ss.sua, dv_goc_x, ss.rong)
                if v0:
                    dem[(m, v0)] = dem.get((m, v0), 0) + 1

        def _tt(v: str) -> str:
            if pv == "so":
                ok_ = C.la_so_tron(v)
            elif pv == "tung_phan":
                ok_ = any(C.la_so_tron(x.strip()) for x in v.split(C.SEP_TSKT))
            else:
                ok_ = bool(re.search(r"\d", v))
            return "✅ có số → được thêm" if ok_ else "⛔ chữ → bỏ qua"
        bang_x = pd.DataFrame([{"Cột": m, "Giá trị": v, "Số ô": n, "Với phạm vi đang chọn": _tt(v)}
                               for (m, v), n in dem.items()])
        if len(bang_x):
            loc_tt = st.radio("Hiện", ["Tất cả", "✅ Được thêm", "⛔ Bỏ qua"], horizontal=True, key="bd_loc_tt", on_change=_bd_mo_bat)
            if loc_tt != "Tất cả":
                bang_x = bang_x[bang_x["Với phạm vi đang chọn"].str.startswith(loc_tt[0])]
            bang_x = bang_x.sort_values(["Cột", "Số ô"], ascending=[True, False]).reset_index(drop=True)
            st.caption(f"{len(bang_x):,} giá trị khác nhau")
            st.dataframe(bang_x, hide_index=True, height=min(320, 40 + 35 * min(len(bang_x), 8)), width="stretch")
        else:
            st.caption("(Cột đã chọn chưa có giá trị.)")
    if chon and (a or b or kieu == "thay"):
        # xem trước tính từ giá trị CHƯA biến đổi hàng loạt của các cột đang chọn (để thấy đúng tác động của bước này)
        dv_goc = {k_: v_ for k_, v_ in ss.dv.items()
                  if not (len(k_) == 3 and k_[2] == "bd" and any(k_[1] == m_ and k_[0] in (c_, "*") for c_, m_ in chon))}
        mau = []
        for cc, m in chon:
            for r in ss.bang[cc]["rows"]:
                v0, _ = C.bien_doi_o(cc, r["sku"], m, r["vals"].get(m, ""), ss.sua, dv_goc, ss.rong)
                v1 = C.ap_buoc(v0, buoc)
                if v0 and v1 != v0:
                    mau.append({"Cột": m, "SKU": r["sku"], "Trước": v0, "Sau": v1})
        st.caption(f"Xem trước: **{len(mau):,} ô** sẽ đổi" + (" — hiện 30 ô đầu" if len(mau) > 30 else ""))
        if mau:
            st.dataframe(pd.DataFrame(mau[:30]), hide_index=True, height=min(300, 40 + 35 * min(30, len(mau))))
        bt = st.columns([1.6, 1.4, 3])
        if buoc_sua is not None:
            if bt[0].button("💾 Lưu thay đổi", type="primary", key="bd_luu_sua"):
                ds_ = list(hien_co)
                ds_[buoc_sua] = buoc
                ss.dv.pop(("*", chon[0][1], "bd"), None)
                ss.dv.pop((chon[0][0], chon[0][1], "bd"), None)
                ss.dv[("*" if moi_nganh else chon[0][0], chon[0][1], "bd")] = ds_
                ss._bd_ctx = None
                bump()
                luu(["settings"], f"Sửa biến đổi hàng loạt cột {chon[0][1]}")
                _bd_rerun()
            if bt[1].button("🗑 Bỏ bước này", key="bd_bo_buoc"):
                ds_ = [x for n, x in enumerate(hien_co) if n != buoc_sua]
                ss.dv.pop(kk_co, None)
                if ds_:
                    ss.dv[kk_co] = ds_
                ss._bd_ctx = None
                bump()
                luu(["settings"], f"Bỏ biến đổi hàng loạt cột {chon[0][1]}")
                _bd_rerun()
        elif bt[0].button(f"▶ Áp cho {len(chon)} cột", type="primary", disabled=not mau):
            for cc, m in chon:
                kk = ("*" if moi_nganh else cc, m, "bd")
                ss.dv[kk] = list(ss.dv.get(kk) or []) + [buoc]
            bump()
            luu(["settings"], f"Biến đổi hàng loạt {len(chon)} cột: {C.BD_KIEU[kieu]}")
            _bd_rerun()
    elif buoc_sua is not None:
        if st.button("🗑 Bỏ bước này", key="bd_bo_buoc2"):
            ds_ = [x for n, x in enumerate(hien_co) if n != buoc_sua]
            ss.dv.pop(kk_co, None)
            if ds_:
                ss.dv[kk_co] = ds_
            ss._bd_ctx = None
            bump()
            luu(["settings"], f"Bỏ biến đổi hàng loạt cột {chon[0][1]}")
            _bd_rerun()
    dang = [(k, v) for k, v in ss.dv.items() if len(k) == 3 and v]
    if dang:
        st.markdown("**Cột đang có biến đổi** (chọn cột ở ô «Cột áp dụng» để xem & sửa):")
        st.dataframe(pd.DataFrame([{"Cột": k[1], "Ngành": "mọi ngành" if k[0] == "*" else k[0],
                                    "Biến đổi": " → ".join(_mo_ta_buoc(x) for x in v)} for k, v in dang]),
                     hide_index=True, height=min(220, 40 + 35 * len(dang)), width="stretch")


def tab_rong(k: dict) -> None:
    ds = list(k["gia_tri_rong"])
    co = {d["khoa"] for d in ds}
    for kk in ss.rong:
        if kk not in co:
            ds.append({"khoa": kk, "so_o": 0, "hien": kk, "cac_dang": [], "cot": []})
    st.caption("Áp cho cột TSKT (không áp FILTER, không đè ô sửa tay). **Để trống** = ô xuất ra để trống, "
               "không ghi đè giá trị trên PIM — nên thử import 1 SKU để chắc PIM bỏ qua ô trống.")
    if not ds:
        st.info("Không thấy giá trị Không / Đang cập nhật nào.")
    nhan_sang_hd = {v: kk for kk, v in C.NHAN_HD.items()}
    hien = pd.DataFrame({"Giá trị": [d["hien"] for d in ds], "Số ô": [d["so_o"] for d in ds],
                         "Kiểu viết gặp": [" | ".join(d["cac_dang"][:4]) for d in ds],
                         "Có ở cột (vd)": [", ".join(d["cot"][:4]) + (f" …(+{len(d['cot']) - 4})" if len(d["cot"]) > 4
                                                                     else "") for d in ds],
                         "XỬ LÝ": [C.NHAN_HD[ss.rong.get(d["khoa"], [C.HD_GIU])[0]] for d in ds],
                         "Thay bằng": [(ss.rong.get(d["khoa"], ["", ""]) + [""])[1] for d in ds]})
    ed = st.data_editor(hien, hide_index=True, key=f"ed_rong_{ss.ver}",
                        disabled=["Giá trị", "Số ô", "Kiểu viết gặp", "Có ở cột (vd)"],
                        column_config={"XỬ LÝ": st.column_config.SelectboxColumn(options=list(C.NHAN_HD.values()),
                                                                                 required=True, width="medium")})
    c = st.columns([1, 1, 2, 1])
    them = c[2].text_input("Thêm giá trị khác", placeholder="vd: Không hỗ trợ")
    if c[0].button("▶ Áp & xem lại", type="primary"):
        moi = dict(ss.rong)
        for i, d in enumerate(ds):
            hd = nhan_sang_hd.get(ed.at[i, "XỬ LÝ"], C.HD_GIU)
            if hd == C.HD_GIU:
                moi.pop(d["khoa"], None)
            else:
                moi[d["khoa"]] = [hd, C.chuan_hoa_key(ed.at[i, "Thay bằng"] or "")]
        if them.strip():
            moi.setdefault(C.khoa_gia_tri_rong(them), [C.HD_TRONG, ""])
        thieu = [kk for kk, v in moi.items() if v[0] == C.HD_THAY and not v[1]]
        if thieu:
            st.error("Chưa nhập 'Thay bằng' cho: " + ", ".join(thieu) + " — chưa áp gì.")
            return
        ss.rong = moi
        bump()
        luu(["settings"], "Áp quy tắc Không/Đang cập nhật")
        st.rerun()
    if c[1].button("Tất cả → Để trống"):
        for d in ds:
            ss.rong[d["khoa"]] = [C.HD_TRONG, ""]
        bump()
        luu(["settings"], "Tất cả giá trị rỗng → để trống")
        st.rerun()


def tab_tach_kt() -> None:
    che_do = st.radio("Chọn", ["🔗 Gộp dài / rộng / cao → 1 cột kích thước", "✂️ Tách cột ghép → các cột con"],
                      horizontal=True, label_visibility="collapsed", key="kt_che_do")
    if che_do.startswith("🔗"):
        tab_gop_kt()
    else:
        tab_tach_kt_con()


def tab_gop_kt() -> None:
    """Gộp 2–4 cột (Dài/Rộng/Cao, Ngang/Cao/Sâu…) thành 1 giá trị trong cột kích thước / kích cỡ của PIM."""
    st.caption("Ví dụ Ngang 30 cm + Cao 20 cm + Sâu 10 cm → cột 'Kích thước' = '30 x 20 x 10 cm' hoặc "
               "'Ngang 30 cm - Cao 20 cm - Sâu 10 cm'. Lấy GIÁ TRỊ SẼ XUẤT (đã có đơn vị/biến đổi). Kết quả ghi dạng "
               "sửa tay (xem lại / bỏ được).")
    cates = list(ss.bang)
    cate = st.selectbox("Ngành hàng", cates, format_func=lambda c: ss.bang[c]["title"], key="gkt_cate")
    b = ss.bang[cate]
    cot = [m for m in C.cot_tt(b) if not C.la_cot_filter(m)]
    ten = b["ten"]

    def nhan(m):
        return f"{m} — {ten.get(m, '')}" if m else "(chọn)"
    la_dich = [m for m in cot if re.search(r"kích thước|kích cỡ|size|dimension", f"{m} {ten.get(m, '')}", re.I)]
    c = st.columns([2, 3])
    ds_dich = [""] + la_dich + [m for m in cot if m not in la_dich]
    dich = c[0].selectbox("Cột ĐÍCH (kích thước / kích cỡ trong PIM)", ds_dich, index=1 if la_dich else 0,
                          format_func=nhan, key=f"gkt_dich_{cate}",
                          help="Tool gợi ý cột có tên chứa 'kích thước / kích cỡ / size'. Không có thì tự chọn.")
    thu_tu = {"dài": 0, "ngang": 0, "rộng": 1, "cao": 2, "sâu": 3, "dày": 3}
    goi_y = sorted([m for m in cot if m != dich and C.la_cot_kich_thuoc(m, ten.get(m, ""))
                    and C.chuan_hoa_ten(ten.get(m, "")).split(" ")[0] in thu_tu and "(" not in ten.get(m, "")],
                   key=lambda m: thu_tu[C.chuan_hoa_ten(ten.get(m, "")).split(" ")[0]])[:3]
    nguon = c[1].multiselect("Cột NGUỒN theo đúng thứ tự {1}, {2}, {3}, {4}", [m for m in cot if m != dich],
                             default=goi_y, max_selections=4, format_func=nhan, key=f"gkt_nguon_{cate}_{dich}")
    c = st.columns([2.2, 1.6, 1, 1.4])
    mau = c[0].selectbox("Mẫu", list(C.MAU_GOP_KT) + ["✍️ Tự nhập mẫu…"], key="gkt_mau",
                         format_func=lambda x: f"{x}  →  {C.MAU_GOP_KT[x]}" if x in C.MAU_GOP_KT else x)
    if mau.startswith("✍️"):
        mau = c[0].text_input("Mẫu tự nhập (dùng {1} {2} {3} {4} và {dv})", value="Rộng {1} - Sâu {2} - Cao {3}",
                              key="gkt_mau_tu")
    bo_dv = c[1].checkbox("Chỉ lấy SỐ, đặt đơn vị 1 lần", value=True, key="gkt_bodv",
                          help="30 cm, 20 cm, 10 cm → 30 x 20 x 10 cm. Bỏ tick = giữ nguyên từng giá trị.")
    dv = c[2].text_input("Đơn vị", key="gkt_dv", placeholder="tự lấy", disabled=not bo_dv)
    thieu = c[3].checkbox("Bỏ qua SKU thiếu 1 giá trị", value=True, key="gkt_thieu")
    chi_trong = st.checkbox("Chỉ điền vào ô đích đang TRỐNG", value=False, key="gkt_trong")
    if not dich:
        st.info("Ngành này chưa có cột tên 'kích thước / kích cỡ' — chọn cột ĐÍCH ở trên.")
        return
    if len(nguon) < 2:
        st.info("Chọn ít nhất 2 cột nguồn.")
        return
    du_kien = []
    for r in b["rows"]:
        gt = [C.bien_doi_o(cate, r["sku"], m, r["vals"].get(m, ""), ss.sua, ss.dv, ss.rong)[0] for m in nguon]
        moi = C.gop_kich_thuoc(gt, mau, bo_dv, dv, thieu)
        hien = C.bien_doi_o(cate, r["sku"], dich, r["vals"].get(dich, ""), ss.sua, ss.dv, ss.rong)[0]
        if moi and moi != hien and (not chi_trong or not hien):
            du_kien.append({"SKU": r["sku"], **{f"{{{i + 1}}}": g for i, g in enumerate(gt)},
                            "Kết quả gộp": moi, "Giá trị hiện tại": hien, "_goc": r["vals"].get(dich, "")})
    st.caption(f"Sẽ điền **{len(du_kien):,}** ô vào `{dich}`.")
    if du_kien:
        st.dataframe(pd.DataFrame(du_kien[:200]).drop(columns=["_goc"]), hide_index=True, height=260)
        if st.button("✔ Gộp & điền", type="primary", key="gkt_ok"):
            for x in du_kien:
                dat_sua(cate, x["SKU"], dich, x["Kết quả gộp"], x["_goc"])
            bump()
            luu(["settings"], f"Gộp kích thước {'+'.join(nguon)} → {dich} ({len(du_kien)} ô)")
            st.rerun()


def tab_tach_kt_con() -> None:
    """Tách cột ghép 'Ngang 122 cm - Cao 71 cm - Dày 7 cm' vào các cột con của ngành hàng."""
    st.caption("Ví dụ cột 'Kích thước loa vệ tinh' = 'Ngang 9.8 cm - Sâu 8.5 cm - Cao 16 cm' → tách vào các cột "
               "Ngang / Sâu / Cao. Kết quả ghi dưới dạng sửa tay (xem lại được, bỏ được).")
    cates = list(ss.bang)
    cate = st.selectbox("Ngành hàng", cates, format_func=lambda c: ss.bang[c]["title"], key="kt_cate")
    b = ss.bang[cate]
    nhan = C.nhan_kich_thuoc_ghep(cate)
    # cột nguồn: cột có giá trị chứa >= 2 nhãn có số
    dem = Counter()
    for r in b["rows"][:400]:
        for m in C.cot_tt(b):
            v = r["vals"].get(m, "")
            if v and len(C.tach_kich_thuoc_ghep(v, nhan)) >= 2:
                dem[m] += 1
    if not dem:
        st.info("Không thấy cột nào chứa giá trị ghép dạng 'Ngang … - Cao …' ở ngành hàng này.")
        return
    nguon = st.selectbox("Cột nguồn (giá trị ghép)", [m for m, _ in dem.most_common()],
                         format_func=lambda m: f"{m} — {b['ten'].get(m, '')} ({dem[m]} dòng mẫu)")
    ten_sang_ma = {C.chuan_hoa_ten(t): m for m, t in b["ten"].items() if t}
    st.markdown("**Nhãn → cột đích** (tool tự gợi ý theo tên cột; để trống = không tách nhãn đó)")
    chon = {}
    cols = st.columns(4)
    for i, n in enumerate(nhan):
        goi_y = ten_sang_ma.get(C.chuan_hoa_ten(n), "")
        tuy_chon = [""] + [m for m in C.cot_tt(b) if not C.la_cot_filter(m) and m != nguon]
        chon[n] = cols[i % 4].selectbox(n, tuy_chon, index=tuy_chon.index(goi_y) if goi_y in tuy_chon else 0,
                                        format_func=lambda m: m and f"{m} — {b['ten'].get(m, '')}" or "(không tách)",
                                        key=f"kt_{cate}_{nguon}_{n}")
    chi_trong = st.checkbox("Chỉ điền vào ô đích đang TRỐNG", value=True)
    du_kien = []
    for r in b["rows"]:
        t = C.tach_kich_thuoc_ghep(r["vals"].get(nguon, ""), nhan)
        for n, v in t.items():
            dich = chon.get(n)
            if dich and (not chi_trong or not r["vals"].get(dich)):
                du_kien.append((r["sku"], dich, v, r["vals"].get(dich, "")))
    st.caption(f"Sẽ điền {len(du_kien):,} ô.")
    if du_kien:
        st.dataframe(pd.DataFrame(du_kien[:200], columns=["SKU", "Cột đích", "Giá trị tách", "Giá trị hiện tại"]),
                     hide_index=True, height=240)
        if st.button("✔ Tách & điền", type="primary"):
            for sku, dich, v, goc in du_kien:
                dat_sua(cate, sku, dich, v, goc)
            bump()
            luu(["settings"], f"Tách kích thước ghép {nguon} ({len(du_kien)} ô)")
            st.rerun()


def tab_chua_map() -> None:
    cm = pd.DataFrame(ss.meta.get("chua_map", []))
    if not len(cm):
        st.success("✔ Mọi thuộc tính CMS của lô này đều đã có mapping.")
        return
    st.caption("Thuộc tính CMS có trong DATA SP nhưng KHÔNG có trong MAPPING TSKT/FILTER (theo mã PROPERTYID). "
               "Dòng 'Đã map theo TÊN' đang được tool tự điền — nên thêm hẳn vào MAPPING TSKT cho chắc.")
    cm = cm.sort_values(["trang_thai", "so_dong"], ascending=[True, False]).reset_index(drop=True)
    cm.insert(0, "Thêm", False)
    ed = st.data_editor(cm, hide_index=True, height=420, key=f"ed_cm_{ss.ver}",
                        disabled=[c for c in cm.columns if c not in ("Thêm", "goi_y_ma")],
                        column_config={"goi_y_ma": st.column_config.TextColumn("Mã MASTER (sửa được)"),
                                       "Thêm": st.column_config.CheckboxColumn(width="small")})
    if not duoc_sua_chung():
        st.caption("Chỉ admin được thêm vào MAPPING dùng chung.")
        return
    if st.button("➕ Thêm dòng đã tick vào MAPPING TSKT (dùng chung)"):
        moi = ed[(ed["Thêm"]) & (ed.goi_y_ma.str.strip() != "")]
        if not len(moi):
            st.warning("Chưa tick dòng nào có Mã MASTER.")
            return
        rows = pd.DataFrame({"cate": moi.cate, "cate_name": [ss.cau_hinh.get(c, {}).get("ten", "") for c in moi.cate],
                             "prop_id": moi.prop_id, "prop_name": moi.prop_name, "ma": moi.goi_y_ma.str.strip(),
                             "ten_ma": ""})
        ss.map_tskt = pd.concat([ss.map_tskt, rows], ignore_index=True).drop_duplicates(["cate", "prop_id"], keep="last")
        luu(["shared:map_tskt"], f"Thêm {len(rows)} dòng MAPPING TSKT từ thuộc tính chưa map")
        st.success(f"Đã thêm {len(rows)} dòng — bấm ① Map dữ liệu lại để áp.")


# ============================================================================
# TRANG: XUẤT FILE
# ============================================================================
def khu_xuat(k: dict) -> None:
    s = k["stat"]
    canh = [(s.get("thieu_model", 0), "SKU thiếu model_code"), (s.get("thieu_cate", 0), "SKU thiếu category_code"),
            (s.get(C.TRANG_THAI_KHAC, 0), "ô KHÁC spec PIM (sẽ ghi đè)"),
            (s.get(C.TRANG_THAI_TOOL_TRONG, 0), "ô tool để trống trong khi PIM đang có"),
            (s.get("chua_don_vi", 0), "ô kích thước/khối lượng chưa có đơn vị"),
            (s.get("khong_co_pim", 0), "dòng không có trong file PIM"), (s.get("filter_chu", 0), "ô FILTER sai mã"),
            (s.get("lech_model", 0), "SKU model_code khác PIM"),
            (ds()["mat_o"], "ô giá trị CMS bị MẤT khi map (xem 🧾 Đối soát)"),
            (_sku_rong(), "SKU KHÔNG có thông số nào (sẽ xuất dòng trống)"),
            (len(C.model_trung(ss.bang)), "model_code lặp ở nhiều dòng file MODEL (PIM sẽ lấy dòng sau cùng — "
                                          "kiểm tra IMPORT có thiếu variant_code không)"),
            (sum((ss.dx_duyet.get(d["id"]) or {}).get("luc", "") > ss.meta.get("luc", "") for d in ss.dx_rieng),
             "đề xuất sửa CMS vừa được admin duyệt/từ chối SAU lần map cuối — nên Map lại"),
            (int((dht()["vi_pham"]["Mức"] == C.MUC_LOI).sum()), "ô vi phạm quy tắc kiểm tra mức LỖI (tab 📈)"),
            (int((~dht()["sku"]["Đủ bắt buộc"]).sum()) if len(dht()["sku"]) else 0,
             "SKU chưa đủ cột BẮT BUỘC (tab 📈 Độ hoàn thiện)")]
    kcl = kc()["loi"]
    kc_cao = int((kcl["Mức"] == "CAO").sum()) if len(kcl) else 0
    # Lỗi CHẶN (cần tick xác nhận mới xuất): dữ liệu sai SKU / thiếu khoá import / FILTER sai mã.
    chan = [(kc_cao, "ô LỆCH SKU / không có nguồn trong DATA SP (tab ✅ Kiểm chứng)"),
            (s.get("thieu_model", 0), "SKU thiếu model_code"), (s.get("thieu_cate", 0), "SKU thiếu category_code"),
            (s.get("filter_chu", 0), "ô FILTER sai mã")]
    chan = [f"• {n:,} {t}" for n, t in chan if n]
    canh = [f"• {n:,} {t}" for n, t in canh if n and t not in ("SKU thiếu model_code", "SKU thiếu category_code",
                                                                  "ô FILTER sai mã")]
    if chan:
        st.markdown("<div class='canh'><b>⛔ LỖI CẦN XỬ LÝ trước khi import:</b><br>" + "<br>".join(chan) +
                    "</div>", unsafe_allow_html=True)
    if canh:
        with st.expander(f"ℹ️ {len(canh)} lưu ý (không chặn xuất)", expanded=not chan):
            st.markdown("<br>".join(canh), unsafe_allow_html=True)
    if not chan and not canh:
        st.success("✔ Không còn cảnh báo — dữ liệu đã kiểm chứng đúng SKU.")
    st.markdown("##### CHỌN NGÀNH HÀNG xuất (giống sheet CHỌN NGÀNH HÀNG của file mẫu)")
    chon = bang_chon_nganh(f"ed_chon_xuat_{ss.ws}_{ss.get('ver_map', 0)}_{ss.get('ver_nap', 0)}_"
                           f"{hash(tuple(ss.bang)) & 0xffffff}")
    xuat_gon("xp", chon, chan + canh, can_xn=bool(chan))
    x = ss.get("xuat")
    if x:
        L = ds()["loi"]
        # Lấy sẵn các bảng (đã có cache) TRƯỚC khi tạo nút: hàm dựng file chạy lúc bấm, NGOÀI luồng của trang nên
        # không được đụng session_state (trước đây gọi dht() trong lambda -> "Failed to generate file for download").
        _dht = dht()
        _sheets = {"CẢNH BÁO": k["canh_bao"], "KHÁC SPEC PIM": pd.DataFrame(k["khac"]),
                   "ĐỘ HOÀN THIỆN": _dht["sku"], "VI PHẠM QUY TẮC": _dht["vi_pham"],
                   "ĐỐI SOÁT CMS": L.drop(columns=[c for c in L.columns if c.startswith("_")]) if len(L) else L}
        nut_tai("📊 Tải báo cáo kiểm tra (.xlsx)", lambda: C.xlsx_nhieu_sheet(_sheets),
                file_name=f"KIEM_TRA_{x['stamp']}.xlsx")
    st.divider()
    st.markdown("##### 📦 Tải workspace theo MẪU (mở bằng Excel / bản desktop 66.py)")
    st.caption("1 file đúng bố cục mẫu: IMPORT, DATA SP, DATA PIM, CẤU HÌNH CATEGORY, MAPPING TSKT MOI, MAPPING FILTER "
               "MOI, CHỌN NGÀNH HÀNG, tab TSKT từng ngành (giá trị SẼ XUẤT, đã áp sửa tay/đơn vị), LOG.")
    if st.button("🧾 Tạo file workspace theo mẫu"):
        with st.spinner("Đang tạo…"):
            ss.ws_mau = C.xuat_workspace_mau(ss["import"], ss.data_sp, ss.data_pim, ss.cau_hinh, ss.map_tskt,
                                             ss.map_filter, {c: b for c, b in ss.bang.items() if c in chon},
                                             [x for x in ss.meta.get("chon", []) if x["MÃ NH"] in chon],
                                             ss.meta.get("log", []), ss.sua, ss.dv, ss.rong)
    if ss.get("ws_mau"):
        st.download_button("⬇️ du_lieu_pim.xlsx", ss.ws_mau, file_name=f"du_lieu_pim_{ss.ws}_{C.bay_gio()[:10]}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def trang_xuat() -> None:
    st.subheader("📤 Xuất file import")
    if bang_trong():
        st.info("Chưa có kết quả map — bấm ① Map dữ liệu ở trên.")
        return
    k = kq()
    khu_xuat(k)


# ============================================================================
# TRANG: CẤU HÌNH & MAPPING (dùng chung)
# ============================================================================
def trang_cau_hinh() -> None:
    st.title("⚙️ Cấu hình & mapping (dùng chung)")
    sua_duoc = duoc_sua_chung()
    if not sua_duoc:
        st.info("Chỉ admin được sửa dữ liệu dùng chung — bạn đang ở chế độ xem.")
    c = st.columns(4)
    c[0].metric("Ngành hàng có cấu hình", len(ss.cau_hinh))
    c[1].metric("Dòng MAPPING TSKT", f"{len(ss.map_tskt):,}")
    c[2].metric("Dòng MAPPING FILTER", f"{len(ss.map_filter):,}")
    c[3].metric("Option DATA PIM", f"{len(ss.data_pim):,}")
    tabs = st.tabs(["📦 Nạp file theo mẫu", "🏷️ Cấu hình ngành hàng", "🧬 Mapping TSKT", "🧮 Mapping FILTER",
                    "🗄️ DATA PIM", "✅ Quy tắc kiểm tra", "🕘 Lịch sử & khôi phục"])
    with tabs[0]:
        khu_nap_mau()
    with tabs[1]:
        tab_cau_hinh(sua_duoc)
    with tabs[2]:
        tab_mapping("map_tskt", C.doc_mapping_tskt, sua_duoc)
    with tabs[3]:
        tab_mapping("map_filter", C.doc_mapping_filter, sua_duoc)
    with tabs[4]:
        f = st.file_uploader("File DATA PIM (cột A Code … F OptionCode, G OptionValue)", type=["xlsx", "csv"],
                             key="up_dpim", disabled=not sua_duoc)
        if f and st.button("🗄️ Thay DATA PIM"):
            df, loi = C.doc_data_pim(C._doc_bang_tho(f.getvalue(), f.name))
            if loi:
                st.error(loi)
            else:
                ss.data_pim = df
                ss.opt = C.option_maps(df)
                bump()
                luu(["shared:data_pim"], f"Thay DATA PIM ({len(df):,} dòng)")
                st.rerun()
        tim = st.text_input("Tìm trong DATA PIM (mã filter / tên option)", key="tim_dpim")
        d = ss.data_pim
        if tim:
            t = tim.lower()
            d = d[d.Code.str.lower().str.contains(t, regex=False) | d.OptionValue.str.lower().str.contains(t, regex=False)
                  | (d.OptionCode == tim)]
        st.dataframe(d.head(2000), hide_index=True, height=380)
    with tabs[5]:
        tab_quy_tac(sua_duoc)
    with tabs[6]:
        tab_lich_su(sua_duoc)


def tab_quy_tac(sua_duoc: bool) -> None:
    st.caption("Quy tắc chất lượng dữ liệu (như Salsify readiness / Pimcore data quality): áp lên GIÁ TRỊ SẼ XUẤT, "
               "báo ở tab 📈 và trang Xuất. Mức **Lỗi** = phải sửa trước khi import; **Cảnh báo** = nên xem. "
               "Cột có quy tắc **Bắt buộc** dùng để tính SKU 'sẵn sàng'.")
    qt = ss.get("quy_tac_kt", {}) or {}
    nganh = ["*"] + sorted(set(ss.cau_hinh) | set(ss.bang))
    cate = st.selectbox("Ngành hàng (* = áp cho mọi ngành có cột đó)", nganh, key="qt_cate",
                        format_func=lambda c: "* (mọi ngành)" if c == "*" else f"{c} — {ss.cau_hinh.get(c, {}).get('ten', '')}")
    rows = [{"Mã cột": k.split("\t", 1)[1], "Loại": v.get("loai"), "Tham số": v.get("tham_so", ""),
             "Mức": v.get("muc", C.MUC_CANH_BAO), "Ghi chú": v.get("ghi_chu", "")}
            for k, v in qt.items() if k.split("\t", 1)[0] == cate and isinstance(v, dict)]
    cot_ds = sorted({m for b in ss.bang.values() for m in C.cot_tt(b)} |
                    {m for c, o in ss.cau_hinh.items() if cate in ("*", c) for m in o.get("cot", [])})
    df = pd.DataFrame(rows, columns=["Mã cột", "Loại", "Tham số", "Mức", "Ghi chú"])
    ed = st.data_editor(df, num_rows="dynamic", hide_index=True, key=f"ed_qt_kt_{cate}_{ss.ver}", disabled=not sua_duoc,
                        column_config={"Mã cột": st.column_config.SelectboxColumn(options=cot_ds, required=True,
                                                                                width="large"),
                                       "Loại": st.column_config.SelectboxColumn(options=list(C.NHAN_QT), required=True),
                                       "Mức": st.column_config.SelectboxColumn(options=[C.MUC_LOI, C.MUC_CANH_BAO],
                                                                               required=True)})
    st.caption(" · ".join(f"**{k}** = {C.NHAN_QT[k]}: {v}" for k, v in C.HUONG_DAN_THAM_SO.items()))
    c = st.columns([1, 1, 3])
    if c[0].button("💾 Lưu quy tắc", type="primary", disabled=not sua_duoc):
        # KHÔNG đọc lại trước khi ghi: luu() gộp 3 chiều với bản máy khác (giữ quy tắc người khác vừa thêm)
        qt = {k: v for k, v in (ss.quy_tac_kt or {}).items() if k.split("\t", 1)[0] != cate}
        for r in ed.fillna("").to_dict("records"):
            if r["Mã cột"] and r["Loại"]:
                qt[f"{cate}\t{r['Mã cột']}"] = {"loai": r["Loại"], "tham_so": str(r["Tham số"]).strip(),
                                                  "muc": r["Mức"] or C.MUC_CANH_BAO, "ghi_chu": str(r["Ghi chú"])}
        ss.quy_tac_kt = qt
        luu(["shared:quy_tac_kt"], f"Lưu quy tắc kiểm tra ngành {cate}")
        st.rerun()
    if cate != "*" and cate in ss.bang and c[1].button("✨ Gợi ý từ dữ liệu lô", disabled=not sua_duoc):
        ss.qt_goi_y = C.goi_y_quy_tac(ss.bang, cate)
    gy = ss.get("qt_goi_y")
    if gy and cate in ss.bang:
        st.markdown("###### Gợi ý (tick dòng muốn thêm, mức Cảnh báo)")
        g = pd.DataFrame(gy)
        g.insert(0, "Thêm", False)
        eg = st.data_editor(g, hide_index=True, key=f"ed_qt_gy_{cate}", disabled=[x for x in g.columns if x != "Thêm"])
        if st.button("➕ Thêm gợi ý đã tick") and eg["Thêm"].any():
            nap_quy_tac()
            qt = dict(ss.quy_tac_kt or {})
            for r in eg[eg["Thêm"]].to_dict("records"):
                qt[f"{cate}\t{r['Mã cột']}"] = {"loai": r["Loại"], "tham_so": r["Tham số"], "muc": C.MUC_CANH_BAO,
                                                  "ghi_chu": "Gợi ý từ dữ liệu"}
            ss.quy_tac_kt, ss.qt_goi_y = qt, None
            luu(["shared:quy_tac_kt"], f"Thêm quy tắc gợi ý ngành {cate}")
            st.rerun()


def tab_lich_su(sua_duoc: bool) -> None:
    S = tao_store()
    if S.backend != "github":
        st.info("Lịch sử phiên bản có khi lưu trên GitHub (mỗi lần lưu = 1 commit).")
        return
    st.caption("Mỗi lần lưu là 1 phiên bản (commit GitHub) — xem ai sửa gì, lúc nào, và KHÔI PHỤC bản cũ nếu lỡ tay "
               "(như versioning của Akeneo). Khôi phục cũng tạo 1 phiên bản mới, không mất lịch sử.")
    ten = {"cau_hinh": "Cấu hình ngành hàng", "map_tskt": "Mapping TSKT", "map_filter": "Mapping FILTER",
           "data_pim": "DATA PIM", "quy_doi": "Quy đổi FILTER", "sua_gt": "Quy tắc sửa CMS theo giá trị",
           "sua_sku": "Quy tắc sửa CMS theo SKU", "quy_tac_kt": "Quy tắc kiểm tra"}
    k = st.selectbox("Dữ liệu", list(ten), format_func=ten.get, key="ls_k")
    try:
        ls = S.lich_su(F_SHARED[k])
    except Exception as e:  # noqa: BLE001
        st.error(str(e))
        return
    if not ls:
        st.caption("(Chưa có phiên bản nào.)")
        return
    df = pd.DataFrame([{"Lúc (UTC)": x["luc"].replace("T", " ").rstrip("Z"), "Nội dung": x["thong_diep"],
                        "Mã": x["sha"][:8]} for x in ls])
    st.dataframe(df, hide_index=True, height=300)
    chon = st.selectbox("Phiên bản muốn xem / khôi phục", range(len(ls)), key="ls_chon",
                        format_func=lambda i: f"{df.at[i, 'Lúc (UTC)']} — {df.at[i, 'Nội dung'][:70]}")
    if st.button("🔍 So với hiện tại"):
        cu = S.doc_ban_cu(F_SHARED[k], ls[chon]["sha"])
        if k in F_JSON:
            a, b = bytes_to_json(cu, {}) or {}, ss[k] or {}
            st.write({"Khoá chỉ có ở bản cũ": len(set(a) - set(b)), "Khoá chỉ có hiện tại": len(set(b) - set(a)),
                      "Khoá khác giá trị": sum(1 for x in set(a) & set(b) if a[x] != b[x])})
        else:
            a, b = bytes_to_df(cu, COT[k]), ss[k]
            ta, tb = set(map(tuple, a.astype(str).values.tolist())), set(map(tuple, b.astype(str).values.tolist()))
            st.write({"Dòng bản cũ": len(a), "Dòng hiện tại": len(b), "Dòng chỉ có ở bản cũ": len(ta - tb),
                      "Dòng chỉ có hiện tại": len(tb - ta)})
    if st.button("↩️ Khôi phục bản này", disabled=not sua_duoc):
        cu = S.doc_ban_cu(F_SHARED[k], ls[chon]["sha"])
        nap_shared()
        ss[k] = (bytes_to_json(cu, {}) or {}) if k in F_JSON else bytes_to_df(cu, COT[k])
        if k == "data_pim":
            ss.opt = C.option_maps(ss.data_pim)
        bump()
        luu([f"shared:{k}"], f"Khôi phục {ten[k]} về {ls[chon]['sha'][:8]}")
        st.rerun()


def khu_nhat_ky_nap_nganh() -> None:
    """ADMIN: xem ai đã nạp/thêm ngành từ file (lấy từ lịch sử commit của file cấu hình) và xoá ngành nạp nhầm.
    Chỉ ĐỌC lịch sử + dùng lại thao tác xoá của tab Cấu hình; lỗi ở đây không ảnh hưởng phần còn lại."""
    with st.expander("📜 Nhật ký nạp ngành từ file (ai nạp gì, lúc nào) — xoá được nếu nạp nhầm"):
        S = tao_store()
        if S.backend != "github":
            st.caption("Lưu cục bộ nên không có nhật ký (chỉ có khi lưu trên GitHub).")
            return
        try:
            ls = [x for x in S.lich_su(F_SHARED["cau_hinh"], 60)
                  if "Thêm cấu hình ngành từ file SKU" in x["thong_diep"] or "Nạp cấu hình ngành từ file mẫu" in x["thong_diep"]]
        except Exception as e:  # noqa: BLE001
            st.caption(f"Không đọc được nhật ký: {e}")
            return
        if not ls:
            st.caption("(Chưa có lần nạp ngành nào.)")
            return
        st.dataframe(pd.DataFrame([{"Lúc (UTC)": x["luc"].replace("T", " ").rstrip("Z"),
                                    "Nội dung": x["thong_diep"]} for x in ls]),
                     hide_index=True, height=min(300, 40 + 35 * len(ls)))
        ids = []
        for x in ls:
            for m in re.findall(r"\((\d+)\)", x["thong_diep"]):
                if m in ss.cau_hinh and m not in ids:
                    ids.append(m)
        if ids:
            xoa = st.selectbox("Ngành đã nạp muốn xoá khỏi cấu hình", ids, key="nk_xoa",
                               format_func=lambda c: f"{c} — {ss.cau_hinh[c].get('ten', '')}")
            if st.button("🗑️ Xoá ngành này khỏi cấu hình", key="nk_xoa_btn"):
                nap_shared()
                ss.cau_hinh.pop(xoa, None)
                bump()
                luu(["shared:cau_hinh"], f"Xoá cấu hình ngành hàng {xoa} (từ nhật ký nạp ngành)")
                st.rerun()
        st.caption("Muốn quay về hẳn 1 thời điểm: tab «Lịch sử & khôi phục».")


def tab_cau_hinh(sua_duoc: bool) -> None:
    if ss.get("admin"):
        khu_nhat_ky_nap_nganh()
    ch = ss.cau_hinh
    tong = pd.DataFrame([{"Mã NH": c, "Tên": v.get("ten", ""), "Số cột": len(v.get("cot", [])),
                          "Số cột FILTER": sum(C.la_cot_filter(x) for x in v.get("cot", []))} for c, v in ch.items()])
    st.dataframe(tong, hide_index=True, height=220)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Xem / sửa cột của 1 ngành hàng**")
        if ch:
            cate = st.selectbox("Ngành hàng", list(ch), format_func=lambda c: f"{c} — {ch[c].get('ten', '')}")
            v = ch[cate]
            df = pd.DataFrame({"Mã cột": v["cot"], "Tên cột": [v.get("ten_cot", {}).get(m, "") for m in v["cot"]]})
            ed = st.data_editor(df, num_rows="dynamic", hide_index=True,
                                key=f"ed_ch_{cate}_{ss.ver}", disabled=not sua_duoc)
            cc = st.columns(2)
            if cc[0].button("💾 Lưu cột ngành hàng", disabled=not sua_duoc):
                ed = ed.fillna("")
                ma = [C.chuan_hoa_key(x) for x in ed["Mã cột"] if C.chuan_hoa_key(x)]
                ch[cate] = {"ten": v.get("ten", ""), "cot": list(dict.fromkeys(ma)),
                            "ten_cot": {C.chuan_hoa_key(a): C.chuan_hoa_key(b) for a, b in
                                        zip(ed["Mã cột"], ed["Tên cột"]) if C.chuan_hoa_key(a)}}
                bump()
                luu(["shared:cau_hinh"], f"Sửa cấu hình ngành hàng {cate}")
                st.rerun()
            if cc[1].button("🗑️ Xoá ngành hàng này", disabled=not sua_duoc):
                ch.pop(cate, None)
                bump()
                luu(["shared:cau_hinh"], f"Xoá cấu hình ngành hàng {cate}")
                st.rerun()
    with c2:
        st.markdown("**Thêm cấu hình**")
        ref = C.cau_hinh_tham_chieu()
        chon = st.multiselect(f"Từ bảng tham chiếu có sẵn ({len(ref)} ngành)", list(ref),
                              format_func=lambda c: f"{c} — {ref[c]['ten']} ({len(ref[c]['cot'])} cột)"
                              + (" ✓ đã có" if c in ch else ""), disabled=not sua_duoc)
        if st.button("➕ Thêm/ghi đè từ tham chiếu", disabled=not (sua_duoc and chon)):
            for c in chon:
                ch[c] = ref[c]
            bump()
            luu(["shared:cau_hinh"], f"Thêm cấu hình {len(chon)} ngành từ tham chiếu")
            st.rerun()
        st.markdown("—")
        f = st.file_uploader("Hoặc từ file template/export PIM của 1 ngành hàng", type=["xlsx", "csv"],
                             key="up_tpl", disabled=not sua_duoc)
        cc = st.columns(2)
        mid = cc[0].text_input("Mã ngành hàng CMS", key="tpl_id")
        mten = cc[1].text_input("Tên ngành hàng", key="tpl_ten")
        if f and mid and st.button("➕ Lấy cột từ template"):
            ch.update(C.cau_hinh_tu_template_pim(f.getvalue(), f.name, mid, mten))
            bump()
            luu(["shared:cau_hinh"], f"Lấy cấu hình {mid} từ template PIM")
            st.rerun()
        f2 = st.file_uploader("Hoặc file CẤU HÌNH CATEGORY (bố cục ngang như bản desktop)", type=["xlsx"],
                              key="up_chn", disabled=not sua_duoc)
        if f2 and st.button("➕ Nạp cấu hình ngang (nối tiếp)"):
            ch.update(C.doc_cau_hinh_ngang(C._doc_bang_tho(f2.getvalue(), f2.name)))
            bump()
            luu(["shared:cau_hinh"], "Nạp cấu hình ngang")
            st.rerun()


def tab_mapping(ten: str, ham_doc, sua_duoc: bool) -> None:
    df = ss[ten]
    c1, c2 = st.columns([3, 2])
    with c2:
        f = st.file_uploader("Nạp file mapping", type=["xlsx", "csv"], key=f"up_{ten}", disabled=not sua_duoc)
        cd = st.radio("Cách ghi", ["Ghi đè toàn bộ", "Nối tiếp (trùng ngành+mã thuộc tính thì thay)"],
                      key=f"cd_{ten}")
        if f and st.button("📥 Nạp mapping", key=f"nap_{ten}"):
            moi, loi = ham_doc(C._doc_bang_tho(f.getvalue(), f.name))
            if loi:
                st.error(loi)
            else:
                ss[ten] = moi if cd.startswith("Ghi") else pd.concat([df, moi], ignore_index=True).drop_duplicates(
                    ["cate", "prop_id"], keep="last")
                bump()
                luu([f"shared:{ten}"], f"Nạp {ten} ({len(moi):,} dòng)")
                st.rerun()
        st.download_button("⬇️ Tải mapping hiện tại (.xlsx)", C.xlsx_nhieu_sheet({ten: df}), file_name=f"{ten}.xlsx",
                           key=f"dl_{ten}")
    with c1:
        cates = sorted(df.cate.unique()) if len(df) else []
        cate = st.selectbox("Ngành hàng", ["(chọn)"] + cates, key=f"cate_{ten}",
                            format_func=lambda c: c if c == "(chọn)" else f"{c} — {ss.cau_hinh.get(c, {}).get('ten', '')}")
        if cate != "(chọn)":
            v = df[df.cate == cate].reset_index(drop=True)
            ed = st.data_editor(v, num_rows="dynamic", hide_index=True, height=360,
                                key=f"ed_{ten}_{cate}_{ss.ver}", disabled=not sua_duoc)
            if st.button("💾 Lưu ngành hàng này", key=f"luu_{ten}", disabled=not sua_duoc):
                ed = ed.fillna("").astype(str)
                ed["cate"] = cate
                ed = ed[(ed.prop_id.str.strip() != "") & (ed.ma.str.strip() != "")]
                ss[ten] = pd.concat([df[df.cate != cate], ed], ignore_index=True)
                bump()
                luu([f"shared:{ten}"], f"Sửa {ten} ngành {cate}")
                st.rerun()


# ============================================================================
# ĐỀ XUẤT SỬA DỮ LIỆU CMS (thành viên đề xuất -> dùng ngay; admin duyệt -> dùng chung lâu dài)
# ============================================================================
def _row_index() -> dict:
    if ss.get("ri_ver") != ss.ver or "ri" not in ss:
        ss.ri = {(c, r["sku"]): r for c, b in ss.bang.items() for r in b["rows"]}
        ss.ri_ver = ss.ver
    return ss.ri


def goc_o(cate: str, sku: str, ma: str) -> str:
    """Giá trị CMS gốc của ô (trước khi áp quy tắc sửa)."""
    g = (ss.meta.get("goc_cms") or {}).get(f"{sku}\t{ma}")
    if g is not None:
        return g
    r = _row_index().get((cate, sku))
    return r["vals"].get(ma, "") if r else ""


def pim_cu_o(sku: str, ma: str) -> str:
    sp = ss.spec
    if not len(sp):
        return ""
    x = sp.gia_tri[(sp.sku == sku) & (sp.ma == ma)]
    return C.lam_sach_gia_tri_pim(x.iloc[0], ma) if len(x) else ""


def kiem_tra_gia_tri(ma: str, v: str) -> str:
    if not C.la_cot_filter(ma):
        return ""
    ma_sai = [t.strip() for t in re.split(r"[,|]", v or "") if t.strip() and (ma, t.strip()) not in ss.opt["opt_ten"]]
    return f"mã option không có trong DATA PIM: {', '.join(ma_sai)}" if ma_sai else ""


def gui_de_xuat(items: list) -> bool:
    """Thành viên: lưu vào danh sách đề xuất của mình (áp ngay cho workspace của mình).
    Admin: duyệt luôn -> ghi vào quy tắc dùng chung."""
    loi = [f"{d['ma']}: {e}" for d in items if (e := kiem_tra_gia_tri(d["ma"], d["gia_tri_moi"]))]
    if loi:
        st.error("Không gửi được — " + "; ".join(loi))
        return False
    if ss.admin:
        nap_quy_tac()
        ds_ad = ss.dx_rieng if ss.ws == ss.user else nap_de_xuat(ss.user)
        for d in items:
            ds_ad.append(d)
            ss.dx_duyet[d["id"]] = {"trang_thai": C.DX_DUYET, "nguoi_duyet": ss.user, "luc": C.bay_gio(),
                                    "ghi_chu": "Admin sửa trực tiếp", "gia_tri": d["gia_tri_moi"]}
            C.ap_de_xuat_vao_quy_tac(d, ss.sua_sku, ss.sua_gt, ss.quy_doi)
        ok = luu(["shared:sua_sku", "shared:sua_gt", "shared:quy_doi", "shared:dx_duyet"],
                 f"Admin áp {len(items)} quy tắc sửa CMS", {p_dx(ss.user): json_to_bytes(ds_ad)})
        if ok:
            ss.dx_tat_ca = ss.get("dx_tat_ca", []) + items
    else:
        moi = {C.khoa_de_xuat(d) for d in items}
        ss.dx_rieng = [d for d in ss.dx_rieng
                       if not (C.trang_thai_dx(d, ss.dx_duyet) == C.DX_CHO and C.khoa_de_xuat(d) in moi)] + items
        ok = luu(["dx"], f"Gửi {len(items)} đề xuất sửa CMS")
    if not ok:
        return False
    for d in items:  # sửa tay trùng với đề xuất -> bỏ (quy tắc đã lo)
        if d["pham_vi"] == C.PV_SKU:
            ss.sua.pop((d["cate"], d["sku"], d["ma"]), None)
        elif d["pham_vi"] == C.PV_GIA_TRI:
            k = C.khoa_quy_doi(d["gia_tri_cu"])
            for key in [kk for kk, v in ss.sua.items() if kk[0] == d["cate"] and kk[2] == d["ma"]
                        and v == d["gia_tri_moi"] and C.khoa_quy_doi(goc_o(*kk)) == k]:
                ss.sua.pop(key, None)
    st.toast(f"📮 {'Đã áp dụng chung' if ss.admin else 'Đã gửi admin & áp dụng ngay cho bạn'}: {len(items)} mục", icon="✅")
    if not bang_trong():
        chay_map_ui()
    return True


def quyet_dinh(ds_qd: list) -> bool:
    """Admin: [(đề xuất, trạng thái, giá trị dùng, ghi chú)] -> cập nhật quy tắc dùng chung (1 commit)."""
    nap_quy_tac()  # đọc bản mới nhất trước khi ghi
    for d, tt, gt, gc in ds_qd:
        cu = (ss.dx_duyet.get(d["id"]) or {}).get("trang_thai", C.DX_CHO)
        ss.dx_duyet[d["id"]] = {"trang_thai": tt, "nguoi_duyet": ss.user, "luc": C.bay_gio(), "ghi_chu": gc,
                                "gia_tri": gt}
        if tt == C.DX_DUYET:
            C.ap_de_xuat_vao_quy_tac(d, ss.sua_sku, ss.sua_gt, ss.quy_doi, gt)
        elif cu == C.DX_DUYET:
            C.bo_quy_tac(d, ss.sua_sku, ss.sua_gt, ss.quy_doi)
    dem = Counter(tt for _, tt, _, _ in ds_qd)
    ok = luu(["shared:sua_sku", "shared:sua_gt", "shared:quy_doi", "shared:dx_duyet"],
             "Duyệt đề xuất: " + ", ".join(f"{v} {k.lower()}" for k, v in dem.items()))
    if ok:
        bump()
    return ok


def _cot_hien(df: pd.DataFrame, bo: tuple = ()) -> pd.DataFrame:
    return df.drop(columns=[c for c in ("id",) + bo if c in df.columns])


def trang_de_xuat() -> None:
    st.title("📮 Đề xuất sửa dữ liệu CMS")
    # --- RÀO: phải có dữ liệu đã map trước mới tạo đề xuất ---
    if not len(ss.get("import", [])) and not len(ss.get("data_sp", [])) and bang_trong():
        st.info("⚠️ Nạp file lô và map trước ở **🚀 Chạy pipeline**, rồi quay lại đây tạo đề xuất.")
        return
    st.markdown("<div class='buoc'>Dữ liệu CMS (TSKT/FILTER) sai 1–2 giá trị → <b>thành viên đề xuất sửa</b>: áp dụng "
                "<b>ngay</b> cho workspace của mình (map lại là có) → <b>admin duyệt</b> → thành quy tắc <b>dùng chung lâu dài"
                "</b> cho mọi tài khoản, mọi lô sau. Bị từ chối thì quy tắc thôi áp (map lại để cập nhật).</div>",
                unsafe_allow_html=True)
    c = st.columns([1, 1, 1, 1, 1.4])
    dm = dem_dx(ss.dx_rieng)
    c[0].metric("⏳ Của tôi chờ duyệt", dm.get(C.DX_CHO, 0))
    c[1].metric("✅ Đã duyệt", dm.get(C.DX_DUYET, 0))
    c[2].metric("❌ Bị từ chối", dm.get(C.DX_TU_CHOI, 0))
    c[3].metric("📚 Quy tắc dùng chung", len(ss.sua_sku) + len(ss.sua_gt) + len(ss.quy_doi))
    if c[4].button("🔄 Cập nhật trạng thái / đề xuất mới", width="stretch"):
        nap_quy_tac()
        ss.dx_rieng = nap_de_xuat(ss.ws)
        if ss.admin:
            nap_tat_ca_de_xuat()
        bump()
        st.rerun()
    if ss.meta.get("luc") and any((ss.dx_duyet.get(d["id"]) or {}).get("luc", "") > ss.meta["luc"] for d in ss.dx_rieng):
        st.warning("Có đề xuất vừa được admin duyệt/từ chối SAU lần map cuối → vào 🚀 Map & kiểm tra bấm **① Map dữ liệu** "
                   "để kết quả dùng đúng quy tắc mới.")
    cho_duyet = [d for d in ss.get("dx_tat_ca", []) if C.trang_thai_dx(d, ss.dx_duyet) == C.DX_CHO] if ss.admin else []
    ten_tab = ["➕ Tạo đề xuất", f"📋 Đề xuất của {'tôi' if ss.ws == ss.user else ss.ws} ({len(ss.dx_rieng)})"]
    if ss.admin:
        ten_tab += [f"📥 Duyệt ({len(cho_duyet)})", "📚 Quy tắc dùng chung"]
    tabs = st.tabs(ten_tab)
    with tabs[0]:
        tab_tao_de_xuat()
    with tabs[1]:
        tab_dx_cua_toi()
    if ss.admin:
        with tabs[2]:
            tab_duyet(cho_duyet)
        with tabs[3]:
            tab_quy_tac_chung()


def tab_tao_de_xuat() -> None:
    if bang_trong():
        st.info("Cần map dữ liệu trước (🚀 Map & kiểm tra) để chọn SKU/cột.")
        return
    if ss.ws != ss.user:
        st.caption(f"Bạn đang xem workspace của {ss.ws}: đề xuất tạo ở đây là của bạn (admin) — áp dụng chung ngay.")
    st.markdown("##### ① Từ các ô đã sửa tay")
    st.caption("Ô bạn đã sửa tay ở tab ≠ Khác spec / 🔎 Theo SKU vì dữ liệu CMS sai → tick để gửi thành đề xuất. "
               "Phạm vi **Mọi SKU cùng giá trị**: mọi SKU trong ngành có đúng giá trị CMS đó đều được sửa.")
    rows = []
    for (cate, sku, ma), v in ss.sua.items():
        goc = goc_o(cate, sku, ma)
        if v == goc:
            continue
        ten = ss.bang.get(cate, {}).get("ten", {}).get(ma, "")
        n, _ = C.dem_sku_cung_gia_tri(ss.bang, cate, ma, goc)
        rows.append({"Gửi": False, "SKU": sku, "NH": cate, "Mã cột": ma, "Tên": ten, "Giá trị CMS (sai)": goc,
                     "Giá trị đúng": v, "Giải nghĩa": C.giai_nghia_filter(ma, v, ss.opt),
                     "Phạm vi": C.NHAN_PHAM_VI[C.PV_SKU], "SKU cùng giá trị": n, "Lý do": ""})
    if not rows:
        st.caption("(Chưa có ô sửa tay nào.)")
    else:
        df = pd.DataFrame(rows)
        ed = st.data_editor(df, hide_index=True, key=f"ed_dx_sua_{ss.ver}", height=min(420, 40 + 35 * len(df)),
                            disabled=[c for c in df.columns if c not in ("Gửi", "Phạm vi", "Lý do", "Giá trị đúng")],
                            column_config={"Gửi": st.column_config.CheckboxColumn(width="small"),
                                           "Phạm vi": st.column_config.SelectboxColumn(
                                               options=[C.NHAN_PHAM_VI[C.PV_SKU], C.NHAN_PHAM_VI[C.PV_GIA_TRI]],
                                               required=True, width="medium"),
                                           "Lý do": st.column_config.TextColumn(width="medium")})
        if st.button("📮 Gửi các dòng đã tick", type="primary", disabled=not ed["Gửi"].any()):
            items, bo = [], 0
            for _, r in ed[ed["Gửi"]].iterrows():
                pv = C.PV_GIA_TRI if r["Phạm vi"] == C.NHAN_PHAM_VI[C.PV_GIA_TRI] else C.PV_SKU
                if pv == C.PV_GIA_TRI and not r["Giá trị CMS (sai)"]:
                    pv, bo = C.PV_SKU, bo + 1
                moi = C.chuan_hoa_key(r["Giá trị đúng"])
                if C.la_cot_filter(r["Mã cột"]):
                    moi = C.ma_filter_tu_gia_tri(r["Mã cột"], moi, ss.opt["option_map"])
                items.append(C.tao_de_xuat(ss.user, ss.ten, pv, r["NH"], r["Mã cột"], r["Giá trị CMS (sai)"], moi,
                                           sku=r["SKU"] if pv == C.PV_SKU else "", ten_cot=r["Tên"],
                                           so_sku=1 if pv == C.PV_SKU else int(r["SKU cùng giá trị"]),
                                           ly_do=r["Lý do"] or "", sku_vd=r["SKU"]))
            if bo:
                st.info(f"{bo} dòng CMS trống → chuyển sang phạm vi 'Chỉ SKU này'.")
            if gui_de_xuat(items):
                st.rerun()
    st.markdown("##### ② Nhập trực tiếp")
    skus = sorted({s for (_, s) in _row_index()})
    c = st.columns([2, 3])
    sku = c[0].selectbox("SKU", skus, key="dx_sku")
    cate = next(cc for (cc, s) in _row_index() if s == sku)
    b = ss.bang[cate]
    ma = c[1].selectbox("Cột", C.cot_tt(b), key="dx_ma",
                        format_func=lambda m: f"{m} — {b['ten'].get(m, '') or ss.opt['ten_filter'].get(m, '')}")
    goc = goc_o(cate, sku, ma)
    hien = _row_index()[(cate, sku)]["vals"].get(ma, "")
    pim = pim_cu_o(sku, ma)
    x = st.columns(3)
    x[0].markdown(f"**CMS:** `{goc or '(trống)'}`" + (f"<br>{C.giai_nghia_filter(ma, goc, ss.opt)}"
                                                       if C.la_cot_filter(ma) else ""), unsafe_allow_html=True)
    x[1].markdown(f"**Kết quả đang dùng:** `{hien or '(trống)'}`" + (" ✔ đã theo quy tắc" if hien != goc else ""))
    x[2].markdown(f"**PIM cũ (web):** `{pim or '(trống)'}`")
    if C.la_cot_filter(ma):
        ds_opt = ss.opt["opt_ds"].get(ma, [])
        nhan = {oc: f"{oc} — {tn}" for oc, tn in ds_opt}
        hien_ma = [t.strip() for t in hien.split(",") if t.strip() in nhan]
        chon = st.multiselect("Giá trị đúng (option)", list(nhan), default=hien_ma, format_func=nhan.get,
                              key=f"dx_ms_{sku}_{ma}")
        moi = C.SEP_FILTER.join(chon)
    else:
        moi = C.chuan_hoa_key(st.text_input("Giá trị đúng", value=hien, key=f"dx_tx_{sku}_{ma}",
                                            help="Nhiều giá trị nối bằng |. Để trống = bỏ giá trị sai."))
    n, _ = C.dem_sku_cung_gia_tri(ss.bang, cate, ma, goc)
    pv_ds = [C.PV_SKU] + ([C.PV_GIA_TRI] if goc else [])
    pv = st.radio("Phạm vi", pv_ds, horizontal=True, key="dx_pv",
                  format_func=lambda p: C.NHAN_PHAM_VI[p] + (f" ({n} SKU)" if p == C.PV_GIA_TRI else ""))
    ly_do = st.text_input("Lý do (giúp admin duyệt nhanh)", key="dx_lydo", placeholder="vd: theo hộp/catalog hãng")
    if st.button("📮 Gửi đề xuất" if not ss.admin else "✅ Áp dụng chung ngay", type="primary",
                 disabled=moi == goc and pv == C.PV_SKU and goc == hien):
        if moi == goc:
            st.warning("Giá trị đúng đang trùng giá trị CMS — không có gì để sửa.")
        elif gui_de_xuat([C.tao_de_xuat(ss.user, ss.ten, pv, cate, ma, goc, moi, sku=sku if pv == C.PV_SKU else "",
                                        ten_cot=b["ten"].get(ma, ""), so_sku=1 if pv == C.PV_SKU else n,
                                        ly_do=ly_do, sku_vd=sku)]):
            st.rerun()


def tab_dx_cua_toi() -> None:
    if not ss.dx_rieng:
        st.caption("(Chưa có đề xuất nào.)")
        return
    df = C.bang_de_xuat(ss.dx_rieng, ss.dx_duyet, ss.opt)
    loc = st.multiselect("Trạng thái", [C.DX_CHO, C.DX_DUYET, C.DX_TU_CHOI, C.DX_THU_HOI], key="dx_loc_toi")
    v = df[df["Trạng thái"].isin(loc)].reset_index(drop=True) if loc else df
    v.insert(0, "Rút", False)
    ed = st.data_editor(_cot_hien(v), hide_index=True, key=f"ed_dx_toi_{ss.ver}", height=min(400, 40 + 35 * len(v)),
                        disabled=[c for c in v.columns if c != "Rút"],
                        column_config={"Rút": st.column_config.CheckboxColumn(width="small",
                                                                              help="Chỉ rút được đề xuất đang chờ")})
    st.caption("⏳ Chờ duyệt: đang áp cho workspace của bạn. ✅ Đã duyệt: áp cho mọi người (giá trị có thể đã được admin "
               "chỉnh). ❌ Từ chối: không còn áp — đọc ghi chú admin, map lại để cập nhật kết quả.")
    if st.button("↩️ Rút đề xuất đã tick (đang chờ)"):
        rut = {v.at[i, "id"] for i in range(len(v)) if ed.at[i, "Rút"] and v.at[i, "Trạng thái"] == C.DX_CHO}
        if rut:
            ss.dx_rieng = [d for d in ss.dx_rieng if d["id"] not in rut]
            if luu(["dx"], f"Rút {len(rut)} đề xuất"):
                chay_map_ui()
                st.rerun()


def tab_duyet(cho: list) -> None:
    if not cho:
        st.success("✔ Không có đề xuất nào chờ duyệt.")
        st.caption(f"Đọc lần cuối: {ss.get('dx_tat_ca_luc', '—')}")
        return
    df = C.bang_de_xuat(cho, ss.dx_duyet, ss.opt)
    by_id = {d["id"]: d for d in cho}
    df["PIM cũ (web)"] = [pim_cu_o(by_id[i]["sku"], by_id[i]["ma"]) if by_id[i].get("sku") else "" for i in df.id]
    df.insert(0, "Quyết định", "⏳ Để sau")
    df["Ghi chú admin"] = ""
    cot = ["Quyết định", "Người đề xuất", "Lúc", "Phạm vi", "Ngành", "SKU", "Mã cột", "Tên cột", "Giá trị CMS (sai)",
           "Giá trị đúng", "Giải nghĩa FILTER", "PIM cũ (web)", "Số SKU", "Lý do", "Ghi chú admin"]
    ed = st.data_editor(df[cot], hide_index=True, key=f"ed_duyet_{ss.ver}", height=min(440, 40 + 35 * len(df)),
                        disabled=[c for c in cot if c not in ("Quyết định", "Giá trị đúng", "Ghi chú admin")],
                        column_config={"Quyết định": st.column_config.SelectboxColumn(
                            options=["⏳ Để sau", "✅ Duyệt", "❌ Từ chối"], required=True, width="medium"),
                            "Giá trị đúng": st.column_config.TextColumn(width="medium",
                                                                       help="Có thể chỉnh trước khi duyệt")})
    st.caption("Duyệt = quy tắc dùng chung cho mọi tài khoản (mỗi lần map sau tự sửa). Từ chối = thôi áp cho người đề xuất. "
               "Phạm vi 'Chỉ SKU này' chỉ áp khi CMS vẫn còn giá trị sai lúc đề xuất (CMS sửa rồi thì tự bỏ qua).")
    c = st.columns(3)
    luu_qd = c[0].button("💾 Lưu quyết định", type="primary")
    duyet_het = c[1].button(f"✅ Duyệt tất cả {len(df)} đề xuất")
    if luu_qd or duyet_het:
        ds_qd, loi = [], []
        for i in range(len(df)):
            q = "✅ Duyệt" if duyet_het else ed.at[i, "Quyết định"]
            if not q or q.startswith("⏳"):
                continue
            d = by_id[df.at[i, "id"]]
            gt = str(ed.at[i, "Giá trị đúng"] or "")
            if C.la_cot_filter(d["ma"]):
                gt = C.ma_filter_tu_gia_tri(d["ma"], gt, ss.opt["option_map"])
            e = kiem_tra_gia_tri(d["ma"], gt)
            if e and q.startswith("✅"):
                loi.append(f"{d['ma']} ({d.get('sku') or d['gia_tri_cu']}): {e}")
                continue
            ds_qd.append((d, C.DX_DUYET if q.startswith("✅") else C.DX_TU_CHOI, gt, str(ed.at[i, "Ghi chú admin"] or "")))
        if loi:
            st.error("Bỏ qua (giá trị không hợp lệ): " + "; ".join(loi))
        if ds_qd and quyet_dinh(ds_qd):
            st.rerun()


def tab_quy_tac_chung() -> None:
    tat_ca = {d["id"]: d for d in ss.get("dx_tat_ca", [])}
    da = [d for d in tat_ca.values() if C.trang_thai_dx(d, ss.dx_duyet) == C.DX_DUYET]
    st.caption(f"Đang dùng chung: {len(ss.sua_sku)} quy tắc theo SKU · {len(ss.sua_gt)} theo giá trị · "
               f"{len(ss.quy_doi)} quy đổi FILTER (gồm cả quy đổi tạo trước khi có tính năng đề xuất).")
    if not da:
        st.caption("(Chưa có đề xuất nào được duyệt.)")
        return
    df = C.bang_de_xuat(da, ss.dx_duyet, ss.opt)
    df.insert(0, "Thu hồi", False)
    ed = st.data_editor(_cot_hien(df, ("Trạng thái",)), hide_index=True, key=f"ed_qt_{ss.ver}", height=min(400, 40 + 35 * len(df)),
                        disabled=[c for c in df.columns if c != "Thu hồi"])
    st.download_button("📊 Tải toàn bộ đề xuất (.xlsx)", C.xlsx_nhieu_sheet(
        {"ĐỀ XUẤT": _cot_hien(C.bang_de_xuat(list(tat_ca.values()), ss.dx_duyet, ss.opt))}),
        file_name=f"DE_XUAT_SUA_CMS_{C.bay_gio()[:10]}.xlsx")
    if st.button("↩️ Thu hồi quy tắc đã tick") and ed["Thu hồi"].any():
        ds_qd = [(tat_ca[df.at[i, "id"]], C.DX_THU_HOI, "", "Thu hồi") for i in range(len(df)) if ed.at[i, "Thu hồi"]]
        if quyet_dinh(ds_qd):
            st.rerun()


# ============================================================================
# TRANG: HƯỚNG DẪN (đọc từ file .md trong repo — sửa file là app cập nhật)
# ============================================================================
def trang_huong_dan() -> None:
    st.title("📘 Hướng dẫn")
    goc = os.path.dirname(os.path.abspath(__file__))
    if ss.get("admin"):
        t = st.tabs(["Dùng hằng ngày", "Cài đặt (admin)", "Quy tắc map (giống desktop)"])
        for tab, f in zip(t, ("HUONG_DAN_SU_DUNG.md", "HUONG_DAN_CAI_DAT.md", "QUY_TAC_MAP.md")):
            with tab:
                p = os.path.join(goc, f)
                st.markdown(open(p, encoding="utf-8").read() if os.path.exists(p) else f"(Thiếu file {f})")
    else:
        p = os.path.join(goc, "HUONG_DAN_SU_DUNG.md")
        st.markdown(open(p, encoding="utf-8").read() if os.path.exists(p) else "(Thiếu file hướng dẫn)")


# ============================================================================
# TRANG: TRA CỨU
# ============================================================================
def trang_tra_cuu() -> None:
    st.title("🧰 Tra cứu")
    t = st.tabs(["🔢 Option FILTER", "📦 SKU trong DATA SP", "📜 Lịch sử xuất"])
    with t[0]:
        c1, c2 = st.columns(2)
        ma = c1.selectbox("Mã FILTER", [""] + sorted(ss.opt["opt_ds"]),
                          format_func=lambda m: m and f"{m} — {ss.opt['ten_filter'].get(m, '')}" or "(chọn)")
        tim = c2.text_input("Tìm theo mã option hoặc tên (vd: 6 hoặc Tivi thông minh)")
        if ma:
            ds = pd.DataFrame(ss.opt["opt_ds"][ma], columns=["Mã option", "Tên"])
            if tim:
                ds = ds[(ds["Mã option"] == tim) | ds["Tên"].str.lower().str.contains(tim.lower(), regex=False)]
            st.dataframe(ds, hide_index=True, height=420)
        elif tim:
            d = ss.data_pim
            d = d[(d.OptionCode == tim) | d.OptionValue.str.lower().str.contains(tim.lower(), regex=False)]
            st.dataframe(d[["Code", "Name", "OptionCode", "OptionValue"]].head(500), hide_index=True)
    with t[1]:
        sku = st.text_input("SKU (PRODUCTCODE)")
        if sku:
            d = ss.data_sp[ss.data_sp.PRODUCTCODE == C.chuan_hoa_code(sku)]
            st.caption(f"{len(d)} dòng thuộc tính CMS")
            st.dataframe(d, hide_index=True, height=420)
    with t[2]:
        st.dataframe(pd.DataFrame(ss.lich_su[::-1]), hide_index=True)


# ============================================================================
# TRANG: QUẢN TRỊ
# ============================================================================
def kiem_tra_he_thong() -> None:
    st.markdown("#### 🩺 Kiểm tra hệ thống")
    st.caption("Chạy sau khi cài đặt / khi app báo lỗi lưu. Mọi dòng ✅ là sẵn sàng.")
    if not st.button("▶ Chạy kiểm tra", key="kt_ht"):
        return
    kq = []
    tk = ds_tai_khoan()
    ad = ", ".join(u for u, v in tk.items() if v.get("admin"))
    kq.append(("Tài khoản trong Secrets", bool(tk) and bool(ad), f"{len(tk)} tài khoản · admin: {ad or 'KHÔNG có admin'}"))
    kq.append(("Mật khẩu đã băm sha256", all(str(v.get("password", "")).startswith("sha256:") for v in tk.values()),
               "nên dùng sha256: cho mọi tài khoản (tạo ở cuối trang này)"))
    S = tao_store()
    kq.append(("Nơi lưu dữ liệu", S.backend == "github", S.mo_ta))
    try:
        kq += S.kiem_tra()
    except Exception as e:  # noqa: BLE001
        kq.append(("Kết nối kho", False, str(e)))
    for mod, ten in (("python_calamine", "Đọc Excel nhanh (python-calamine)"), ("xlsxwriter", "Ghi Excel nhanh (xlsxwriter)"),
                     ("pyarrow", "Lưu bảng nén (pyarrow)")):
        try:
            m = __import__(mod)
            kq.append((ten, True, getattr(m, "__version__", "có")))
        except Exception:  # noqa: BLE001
            kq.append((ten, False, "chưa cài — kiểm tra requirements.txt"))
    import resource
    ram = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
    kq.append(("RAM máy chủ đang dùng (đỉnh)", ram < 2000, f"{ram:,.0f} MB (Streamlit Cloud miễn phí ~2,7 GB cho mọi người)"))
    kq.append(("Dữ liệu dùng chung", len(ss.cau_hinh) > 0 and len(ss.map_tskt) > 0,
               f"{len(ss.cau_hinh)} ngành cấu hình · {len(ss.map_tskt):,} mapping TSKT · {len(ss.map_filter):,} FILTER · "
               f"{len(ss.data_pim):,} option"))
    t = time.time()
    try:
        ok, msg, _, _ = S.luu({"shared/_kiem_tra_ghi.json": json_to_bytes({"luc": C.bay_gio(), "nguoi": ss.user})},
                              f"[{ss.user}] Kiểm tra ghi")
        kq.append(("Ghi thử 1 file", ok, f"{msg} · {time.time() - t:.1f} giây"))
    except Exception as e:  # noqa: BLE001
        kq.append(("Ghi thử 1 file", False, str(e)))
    st.dataframe(pd.DataFrame([{"": "✅" if ok else "❌", "Hạng mục": a, "Chi tiết": b} for a, ok, b in kq]),
                 hide_index=True, width="stretch")
    chan_doan_kho_user()
    nut_sao_luu()
    if HIEN_AI:
        thu_vien_ai_hoc()


def nut_sao_luu() -> None:
    """Sao lưu 1 chạm: gom toàn bộ file dùng chung + workspace đang xem thành 1 file zip tải về máy (dự phòng ngoài GitHub)."""
    import io
    import zipfile
    st.markdown("#### 💾 Sao lưu dự phòng")
    st.caption("Mỗi lần lưu, GitHub đã giữ lịch sử (khôi phục ở ⚙️ Cấu hình → lịch sử). Thêm bản zip này để giữ 1 bản trên máy anh.")
    if st.button("Chuẩn bị bản sao lưu (.zip)", key="btn_sao_luu"):
        with st.spinner("Đang gom file…"):
            buf = io.BytesIO()
            n = 0
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                paths = list(F_SHARED.values()) + [p_user(x) for x in
                                                   ("import", "data_sp", "spec", "ket_qua", "ket_qua_meta", "settings", "lich_su")]
                for p in paths:
                    try:
                        d = doc_file(p)
                    except Exception:  # noqa: BLE001
                        d = None
                    if d:
                        zf.writestr(p, d)
                        n += 1
        ss.sao_luu = (f"SAO_LUU_PIM_{C.bay_gio()[:10]}_{ss.ws}.zip", buf.getvalue(), n)
    x = ss.get("sao_luu")
    if x:
        st.download_button(f"⬇ Tải bản sao lưu ({x[2]} file)", x[1], file_name=x[0], key="dl_sao_luu")


def thu_vien_ai_hoc() -> None:
    """Trang quản lý thư viện AI học (admin): xem số liệu, xoá, xuất Excel."""
    st.markdown("#### 📚 Thư viện AI tự học (từ các lô đã xuất file)")
    hoc = ss.get("ai_hoc") or {}
    if not hoc:
        st.caption("Chưa có dữ liệu. Mỗi lần ai đó bấm 📤 Tạo file import ở trang Xuất, các ô có giá trị sẽ được "
                   "gom vào đây theo (ngành · cột) để làm mẫu cho các lô sau (dùng ở tab 📚 QC ngược).")
        return
    rows = []
    for cate, cols in sorted(hoc.items()):
        for ma, info in sorted((cols or {}).items()):
            rows.append({"Mã NH": cate, "Mã cột": ma,
                         "Số mẫu": info.get("total_samples", 0),
                         "Số giá trị top": len(info.get("top_values") or {}),
                         "Lần cuối": info.get("last_updated", ""),
                         "Người góp": ", ".join(info.get("contributors") or [])[:60]})
    import pandas as _pd
    df = _pd.DataFrame(rows)
    st.caption(f"Tổng: **{len(hoc)}** ngành · **{len(df):,}** cột · "
               f"**{int(df['Số mẫu'].sum()):,}** mẫu")
    st.dataframe(df, hide_index=True, height=min(400, 60 + 35 * min(len(df), 10)))
    cc = st.columns(3)
    if cc[0].button("🗑 Xoá toàn bộ thư viện AI học", disabled=not ss.admin):
        ss.ai_hoc = {}
        luu(["shared:ai_hoc"], f"[{ss.user}] Xoá thư viện AI học")
        st.success("✔ Đã xoá.")
        st.rerun()
    cat_xoa = cc[1].selectbox("Xoá riêng ngành", [""] + sorted(hoc.keys()), key="aihoc_xoa_cate")
    if cc[1].button("🗑 Xoá ngành đã chọn", disabled=not (ss.admin and cat_xoa)):
        ss.ai_hoc = {c: v for c, v in hoc.items() if c != cat_xoa}
        luu(["shared:ai_hoc"], f"[{ss.user}] Xoá thư viện AI học ngành {cat_xoa}")
        st.success(f"✔ Đã xoá ngành {cat_xoa}.")
        st.rerun()


def chan_doan_kho_user() -> None:
    """Hiện danh sách file workspace hiện tại trên kho (đang xem) + số dòng/bytes.
    Dùng khi app báo 'không có dữ liệu' để biết chính xác kho có gì."""
    st.markdown("#### 📦 Chẩn đoán kho workspace đang xem")
    st.caption(f"Workspace: `{ss.ws}` · Giúp xem chính xác file nào có / trống trên GitHub.")
    S = tao_store()
    rows = []
    for ten, f in F_USER.items():
        path = f"users/{ss.ws}/{f}"
        try:
            data = S.doc(path)
        except Exception as e:  # noqa: BLE001
            rows.append({"File": f, "Trạng thái": f"❌ Lỗi: {str(e)[:50]}", "Kích thước": "", "Nội dung": ""})
            continue
        if data is None:
            rows.append({"File": f, "Trạng thái": "⚠️ Không có trên kho", "Kích thước": "", "Nội dung": ""})
            continue
        kich = f"{len(data):,} bytes"
        noi_dung = ""
        try:
            if f.endswith(".parquet"):
                df = bytes_to_df(data, COT.get(ten, []))
                noi_dung = f"{len(df):,} dòng, {len(df.columns)} cột"
            elif f.endswith(".json"):
                j = bytes_to_json(data, None)
                if isinstance(j, dict):
                    noi_dung = f"dict: {len(j)} khoá"
                elif isinstance(j, list):
                    noi_dung = f"list: {len(j)} phần tử"
                else:
                    noi_dung = type(j).__name__
        except Exception as e:  # noqa: BLE001
            noi_dung = f"⚠️ Đọc lỗi: {str(e)[:60]}"
        rows.append({"File": f, "Trạng thái": "✅ Có", "Kích thước": kich, "Nội dung": noi_dung})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    st.caption("Nếu file quan trọng (import, data_sp) hiện **Không có trên kho** mà bang (ket_qua) lại có → "
               "session trước lưu dở, nạp lại ở 🚀 Chạy pipeline → ① Nạp dữ liệu.")


def trang_quan_tri() -> None:
    st.title("👥 Quản trị")
    kiem_tra_he_thong()
    st.divider()
    S = tao_store()
    rows = []
    for u, v in ds_tai_khoan().items():
        try:
            stt = bytes_to_json(S.doc(f"users/{u}/settings.json"), {}) or {}  # chỉ đọc
            ls = bytes_to_json(S.doc(f"users/{u}/lich_su.json"), []) or []
        except Exception as e:  # noqa: BLE001
            stt, ls = {"loi": str(e)}, []
        rows.append({"Tài khoản": u, "Tên": v.get("ten", ""), "Admin": bool(v.get("admin")),
                     "Lưu gần nhất": stt.get("luc_luu", ""), "Ô sửa tay": len(stt.get("sua", {})),
                     "Đề xuất chờ duyệt": sum(d.get("nguoi") == u and C.trang_thai_dx(d, ss.dx_duyet) == C.DX_CHO
                                              for d in ss.get("dx_tat_ca", [])),
                     "Số lần xuất": len(ls), "Xuất gần nhất": ls[-1]["Lúc"] if ls else ""})
    st.dataframe(pd.DataFrame(rows), hide_index=True)
    st.caption("Mở workspace của tài khoản khác: chọn ở ô 'Workspace đang xem' trên thanh bên.")
    if HIEN_AI:
        st.markdown("#### 🤖 AI miễn phí")
        ai = tao_ai()
        st.caption(f"Đang dùng: {ai.mo_ta} · key: {'đã cấu hình' if ai.key else 'CHƯA có'}. Cấu hình cố định trong Secrets: "
                   "AI_PROVIDER = \"groq\" | \"gemini\" | \"openrouter\", AI_API_KEY = \"...\", AI_MODEL = \"...\" (tuỳ chọn).")
        c = st.columns(2)
        if c[0].button("🔌 Kiểm tra kết nối AI", disabled=not ai.co_san):
            try:
                st.success("AI trả lời: " + AIH.AI.chat(ai, [{"role": "user", "content": "Trả lời đúng 1 chữ: OK"}], 20)[:80])
            except AIH.LoiAI as e:
                st.error(str(e))
        if c[1].button("📃 Liệt kê model", disabled=not ai.key):
            try:
                st.write(ai.ds_model())
            except AIH.LoiAI as e:
                st.error(str(e))
        st.markdown("#### 🔑 API key AI (dùng chung cả nhóm, nhập 1 lần trong Secrets)")
        huong_dan_secrets_ai()
        st.caption("Key để trong Secrets của Streamlit (không lưu lên kho dữ liệu vì kho đang Public — tránh lộ key và "
                   "nhà cung cấp tự thu hồi). Đổi key: sửa dòng AI_API_KEY trong Secrets → Save.")
        if (ss.get("ai_cau_hinh") or {}).get("api_key") and st.button("🗑 Xoá key cũ lưu trên kho"):
            ss.ai_cau_hinh = {}
            if luu(["shared:ai_cau_hinh"], f"[{ss.user}] Xoá AI config"):
                st.rerun()

    st.markdown("#### 🆕 Tài khoản thành viên tự đăng ký")
    kho = tk_kho()
    if not kho:
        st.caption("Chưa có ai đăng ký. Thành viên tự vào tab 🆕 Tạo tài khoản ở trang đăng nhập. "
                   "Đặt `MA_MOI = \"...\"` trong Secrets để ai nhập đúng mã được dùng ngay, không cần duyệt.")
    else:
        for u, v in sorted(kho.items()):
            c = st.columns([3, 2, 1, 1, 1])
            c[0].write(f"**{u}** — {v.get('ten', '')}")
            c[1].caption(f"{v.get('trang_thai', 'active')} · {v.get('tao_luc', '')}")
            if v.get("trang_thai") == "cho_duyet" and c[2].button("✅ Duyệt", key=f"dk_ok_{u}"):
                def _d(d, u=u):
                    d[u]["trang_thai"] = "active"
                    return ""
                st.toast(sua_tk_kho(_d, f"[{ss.user}] Duyệt {u}")[1] or "Đã duyệt")
                st.rerun()
            if c[3].button("🔑 Đặt MK", key=f"dk_mk_{u}", help="Đặt lại mật khẩu thành 123456 — nhắn thành viên đổi lại"):
                def _m(d, u=u):
                    d[u]["password"] = bam_mat_khau("123456")
                    return ""
                st.toast(sua_tk_kho(_m, f"[{ss.user}] Đặt lại MK {u}")[1] or "Mật khẩu = 123456")
            if c[4].button("🗑 Xoá", key=f"dk_xoa_{u}"):
                def _x(d, u=u):
                    d.pop(u, None)
                    return ""
                st.toast(sua_tk_kho(_x, f"[{ss.user}] Xoá {u}")[1] or "Đã xoá")
                st.rerun()
    st.markdown("#### Tài khoản admin (Secrets)")
    st.caption("Admin và tài khoản cài sẵn thêm vào Streamlit Secrets (mật khẩu có thể lưu dạng băm sha256 để không lộ chữ thật):")
    mk = st.text_input("Tạo chuỗi băm cho mật khẩu", type="password")
    if mk:
        st.code(f'[users.ten_dang_nhap]\npassword = "sha256:{hashlib.sha256(mk.encode()).hexdigest()}"\n'
                f'ten = "Tên hiển thị"\nadmin = false', language="toml")


def vung_chay() -> None:
    """VÙNG 1 (giống tab 🚀 Chạy pipeline của 66.py): ① Nạp → ② Map → ③ Xuất."""
    for m in ss.pop("flash", []) or []:
        st.success(m)
    for m in ss.pop("flash_err", []) or []:
        st.error(m)
    st.markdown("### ① Nạp dữ liệu")
    t1, t2, t3 = st.tabs(["Dữ liệu lô (SKU · CMS · file mẫu ngành)", "Data gốc (TSKT · FILTER · DATA PIM · cấu hình)",
                          "👀 Xem dữ liệu đã nạp"])
    with t1:
        khu_nap_nhanh("nap")
        with st.expander("Nạp riêng từng loại file (nâng cao)"):
            khu_nap_rieng()
    with t2:
        khu_nap_mau()
    with t3:
        xem_du_lieu()
    st.divider()
    xong1 = len(ss.data_sp) > 0 and len(ss["import"]) > 0
    st.markdown("### ② Map")
    if not xong1:
        st.info("👆 Nạp file ở **bước ①** trước rồi mới bấm Map được.")
    else:
        if len(ss.data_sp):
            _d = nq()
            _L = _d["loi"]
            _c = int((_L["Mức"] == "CAO").sum()) if len(_L) else 0
            _t = int((_L["Mức"] == "TB").sum()) if len(_L) else 0
            _o = st.container(border=True)
            _o.markdown(("⛔" if _c else ("⚠️" if _t else "✅")) + f" **QC ngầm trước khi map** — {len(_d['nganh'])} ngành · "
                        f"{_c} lỗi · {_t} cần xem (cấu hình · mapping TSKT/FILTER · DATA PIM)")
            if _o.checkbox("Xem chi tiết QC ngầm", value=bool(_c), key="nq_truoc_map_xem"):
                with _o:
                    hien_nhat_quan(_d, "truoc_map_v", gon=True)
    cm = st.columns([1.2, 3])
    if cm[0].button("🚀 Map dữ liệu", type="primary", disabled=not xong1, key="nap_nut_map", width="stretch"):
        chay_map_ui()
        if not bang_trong():
            ss.map_flash = ss.get("map_msg", "")
            ss.chuyen_vung = "🔍 Kiểm tra & Đối chiếu"  # như 66.py: map xong tự sang tab kiểm tra
            st.rerun()
    cm[1].caption(f"Map lần cuối: {ss.meta['luc']} — xem kết quả ở 🔍 Kiểm tra." if ss.meta.get("luc") else
                  ("Map xong tự chuyển sang 🔍 Kiểm tra." if xong1 else
                   "Nạp file ở bước ① trước."))
    st.markdown("### ③ Xuất file import")
    if not xong1:
        st.info("👆 Nạp file ở **bước ①** và bấm **Map ở bước ②** trước rồi mới xuất được.")
    elif bang_trong():
        st.caption("👆 Bấm **② Map** ở trên trước — có kết quả rồi mới xuất được.")
    else:
        _cn = _chan_nhanh()
        xuat_gon("ln", None, _cn, can_xn=bool(_cn))
        if ss.lich_su:
            with st.expander("Lần xuất gần đây"):
                st.dataframe(pd.DataFrame(ss.lich_su[-10:][::-1]), hide_index=True)


def vung_du_lieu() -> None:
    """VÙNG 3 (giống tab 📋 Quản lý dữ liệu của 66.py): cấu hình, mapping, đề xuất, tra cứu, hướng dẫn, quản trị."""
    muc = {"⚙️ Cấu hình & mapping": trang_cau_hinh, "📮 Đề xuất sửa CMS": trang_de_xuat, "🧰 Tra cứu": trang_tra_cuu,
           "📘 Hướng dẫn": trang_huong_dan}
    if ss.admin:
        muc["👥 Quản trị"] = trang_quan_tri
    ds_muc = list(muc)
    if ss.get("vung3") not in ds_muc:
        ss.vung3 = ss.get("_vung3_giu") if ss.get("_vung3_giu") in ds_muc else ds_muc[0]
    st.segmented_control("Mục", ds_muc, key="vung3", required=True, label_visibility="collapsed")
    ss._vung3_giu = ss.vung3
    muc[ss.vung3]()


# ============================================================================
# CHẠY
# ============================================================================
dang_nhap()
def _nap_bao_ve(nhan: str, ham) -> None:
    """Tải dữ liệu lúc mở app: lỗi tạm thời (GitHub bận/mạng) -> báo nhẹ + nút Thử lại, KHÔNG sập app, KHÔNG mất đăng nhập."""
    try:
        with st.spinner(nhan):
            ham()
    except Exception as e:  # noqa: BLE001
        _ghi_loi(e, nhan)
        st.warning("⏳ Chưa tải được dữ liệu (kho dữ liệu đang bận hoặc mạng chập chờn). Dữ liệu của bạn vẫn an toàn — "
                   "bấm **Thử lại** sau vài giây.")
        if st.button("🔄 Thử lại", type="primary", key="nap_thu_lai"):
            st.rerun()
        if ss.get("admin"):
            st.caption(f"Chi tiết: {type(e).__name__}: {str(e)[:200]}")
        st.stop()


if "cau_hinh" not in ss:
    _nap_bao_ve("Đang tải dữ liệu dùng chung…", nap_shared)
if ss.get("ws_da_nap") != ss.ws:
    _nap_bao_ve(f"Đang tải workspace {ss.ws}…", nap_workspace)
thanh_ben()
# Non-admin: ẩn chỉ báo hệ thống (Streamlit running/error indicators)
if not ss.get("admin"):
    st.markdown('<style>div[data-testid="stStatusWidget"]{display:none !important;}'
                'div[data-testid="stDecoration"]{display:none !important;}</style>', unsafe_allow_html=True)
hien_xung_dot()
VUNG = ["🚀 Chạy pipeline", "🔍 Kiểm tra & Đối chiếu", "📋 Quản lý dữ liệu"]
if "vung" not in ss and ss.get("_vung_giu") in VUNG:  # giữ vùng đang làm khi ô chọn vùng bị xoá trạng thái
    ss.vung = ss._vung_giu
if ss.get("chuyen_vung") in VUNG:  # nhảy sang vùng khác (vd: map xong -> sang Kiểm tra) TRƯỚC khi vẽ ô chọn vùng
    ss.vung = ss.pop("chuyen_vung")
ss.setdefault("vung", VUNG[0])
if ss.vung not in VUNG:
    ss.vung = VUNG[0]
# Brand bar: admin = đầy đủ thông tin; non-admin = tên + giờ thực
try:
    _logo_html = dmx_chu_img(30) or (f'<span class="mark">{logo_img(34)}</span>')
    if ss.get("admin"):
        _ad = " · 🛡️ Admin"
        _ws = f"Workspace: <b>{ss.ws}</b>"
        _meta = (f'<span class="chip">👤 {ss.ten}{_ad}</span>'
                 f'<span class="chip">{_ws}</span>')
    else:
        _meta = (f'<span class="chip">👤 {ss.ten}</span>'
                 f'<span class="chip">🕐 {_gio_vn()}</span>')
    st.markdown(
        f'<div class="brand-bar">'
        f'<div class="logo">{_logo_html}'
        f'<span class="sep"></span>PIM Tool <span style="opacity:.9;font-weight:500;font-size:.9rem">· CMS → PIM</span></div>'
        f'<div class="meta">{_meta}</div>'
        f'</div>', unsafe_allow_html=True)
except Exception:
    pass

st.segmented_control("Vùng làm việc", VUNG, key="vung", required=True, width="stretch", label_visibility="collapsed")
ss._vung_giu = ss.vung

try:
    # Thanh trạng thái: card ngang với màu tắt khi 0
    try:
        _s_imp = len(ss.get('import', []))
        _s_dsp = len(ss.get('data_sp', []))
        _s_spec = ss.spec.sku.nunique() if ('spec' in ss and len(ss.spec)) else 0
        _s_bang_sku = sum(len(b.get('rows', [])) for b in ss.get('bang', {}).values())
        _s_bang_cate = len(ss.get('bang', {}))
        _s_ch = len(ss.get('cau_hinh', {}))
        _s_mt = len(ss.get('map_tskt', []))
        _s_mf = len(ss.get('map_filter', []))
        thanh_kho([("import", "IMPORT", _s_imp, "SKU"), ("data_sp", "DATA SP", _s_dsp, "dòng"),
                   ("spec", "Spec cũ", _s_spec, "SKU"), ("bang", "Kết quả map", _s_bang_sku, f"SKU · {_s_bang_cate} ngành"),
                   ("cau_hinh", "Cấu hình", _s_ch, "ngành"), ("map_tskt", "Mapping TSKT", _s_mt, "dòng"),
                   ("map_filter", "FILTER", _s_mf, "dòng")])
    except Exception:
        pass
    thanh_luu_lai()
    thanh_hoan_tac()
    if ss.vung == VUNG[0]:
        vung_chay()
    elif ss.vung == VUNG[1]:
        trang_map()
    else:
        vung_du_lieu()
except Exception as _e_main:
    _ghi_loi(_e_main, "Vùng chính")
    if ss.get("admin"):
        st.error(f"**Lỗi hệ thống:** `{type(_e_main).__name__}: {_e_main}`")
        with st.expander("Chi tiết lỗi"):
            st.code("".join(traceback.format_exception(_e_main)), language="python")
    else:
        _hien_loi_vui()

# Footer: mọi người có nút dọn RAM + cảnh báo khi RAM cao; admin thấy thêm version/số liệu
try:
    _u, _g, _c = ram_mb()
    _dp = _dieu_phoi()
    _cao = bool(_g) and _u > 0.75 * _g
    if _cao and time.time() - _dp.get("tra_luc", 0) > 30:  # RAM cao -> tự thu gom (tối đa 30 giây/lần cho cả máy chủ)
        _dp["tra_luc"] = time.time()
        tra_ram()
        _u, _g, _c = ram_mb()
        _cao = bool(_g) and _u > 0.75 * _g
except Exception:  # noqa: BLE001
    _u = _g = _c = 0
    _dp = {"dang": {}, "toi_da": 1}
    _cao = False
_kq_dr = ss.pop("_don_ram_kq", None)
_cf = st.columns([4, 2, 4])
with _cf[1]:
    st.button("🧹 Dọn RAM", key="btn_don_ram", on_click=_cb_don_ram, width="stretch",
              help="Giải phóng bộ nhớ của phiên này và trả RAM về máy chủ khi bị đầy / chậm. KHÔNG mất dữ liệu đang làm "
                   "(IMPORT, DATA SP, kết quả map, mapping, sửa tay…). Chỉ xoá các bảng kiểm tra tính lại được.")
    st.checkbox("Cả file import đã tạo", key="_don_ram_xuat",
                help="Tick thì dọn luôn các file import đã tạo (muốn tải lại phải bấm «Tạo file import» lần nữa).")
if _kq_dr:
    st.success(f"✔ Đã dọn: RAM {_kq_dr['truoc']:,} → {_kq_dr['sau']:,} MB (xoá {_kq_dr['n']} bộ nhớ đệm). "
               "Dữ liệu đang làm vẫn nguyên.")
if _cao:
    st.warning(f"⚠️ RAM máy chủ đang cao ({_u:,}/{_g:,} MB). Bấm «🧹 Dọn RAM»; nếu vẫn cao, chờ người khác xong hoặc chia lô "
               "nhỏ hơn để tránh bị văng.")
if ss.get("admin"):
    _sk = (f" · RAM {_u:,}/{_g:,} MB" if _g else f" · RAM còn {_c:,} MB") + f" · việc nặng {len(_dp['dang'])}/{_dp['toi_da']}"
    st.caption(f"<div style='text-align:center;margin-top:1rem;color:#94a3b8;font-size:.78rem'>"
               f"PIM Tool {APP_VERSION}{_sk}</div>", unsafe_allow_html=True)
