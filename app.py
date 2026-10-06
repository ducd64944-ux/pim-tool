# -*- coding: utf-8 -*-
"""
PIM Tool (web) — TGDĐ CMS -> PIM: map TSKT + FILTER, kiểm tra/đối chiếu, sửa trực tiếp, xuất file import.
Bản Streamlit của tool desktop 66.py. Dữ liệu lưu GitHub (xem gh_store.py), mỗi tài khoản 1 workspace.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import random
import secrets
import re
import time
from collections import Counter

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PIM Tool — CMS → PIM", page_icon="🧩", layout="wide",
                   menu_items={"Get Help": None, "Report a bug": None, "About": None})

import ai_helper as AIH  # noqa: E402
import pim_core as C  # noqa: E402
import dong_bo as DB  # noqa: E402
from gh_store import KHONG_CO, Store, bytes_to_df, bytes_to_json, df_to_bytes, git_sha, json_to_bytes  # noqa: E402

APP_VERSION = "web-1.6 · 2026-10-06 (nạp→map→kiểm tra 1 trang · xem dữ liệu · nạp lại data gốc · nạp theo từng vùng · biến thể màu → MODEL · tự tạo tài khoản · nạp 1 cục · QC tổng hợp · biến đổi hàng loạt · xin data CMS)"
ss = st.session_state

st.markdown("""
<style>
#MainMenu, footer {visibility: hidden;}
div[data-testid="stMetric"] {background: #f6f8fd; border: 1px solid #e3e8f4; border-radius: 10px; padding: 8px 12px;}
div[data-testid="stMetricValue"] {font-size: 1.45rem;}
.buoc {background:#eef4ff;border-left:4px solid #2f6fed;padding:8px 12px;border-radius:6px;margin:4px 0 10px 0;}
.canh {background:#fff4f2;border-left:4px solid #c0392b;padding:8px 12px;border-radius:6px;margin:4px 0 10px 0;}
</style>""", unsafe_allow_html=True)


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
            return {k: dict(v) for k, v in u.items()}
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
    st.title("🧩 PIM Tool — CMS → PIM")
    if not tk:
        st.error("Chưa cấu hình tài khoản. Vào Streamlit Cloud → app → Settings → Secrets, thêm:")
        st.code('[users.ducd]\npassword = "mat-khau"\nten = "Đức"\nadmin = true', language="toml")
        st.stop()
    t_dn, t_dk = st.tabs(["🔑 Đăng nhập", "🆕 Tạo tài khoản"])
    with t_dn:
        with st.form("dang_nhap"):
            u = st.text_input("Tài khoản (tên đăng nhập)").strip()
            p = st.text_input("Mật khẩu", type="password")
            ok = st.form_submit_button("Đăng nhập", type="primary")
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
            ok2 = st.form_submit_button("Tạo tài khoản")
        if ok2:
            if m2 != m3:
                st.error("Hai mật khẩu không giống nhau.")
            else:
                try:
                    done, msg = dang_ky(u2, t2, m2, ma)
                except Exception as e:  # noqa: BLE001
                    done, msg = False, f"Lỗi lưu: {e}"
                (st.success if done else st.error)(msg)
    st.stop()


# ============================================================================
# DỮ LIỆU: NẠP / LƯU
# ============================================================================
F_SHARED = {"cau_hinh": "shared/cau_hinh.json", "quy_doi": "shared/quy_doi_filter.json",
            "sua_sku": "shared/sua_sku.json", "sua_gt": "shared/sua_gia_tri.json",
            "dx_duyet": "shared/de_xuat_duyet.json", "quy_tac_kt": "shared/quy_tac_kiem_tra.json",
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


def bump() -> None:
    ss.ver = ss.get("ver", 0) + 1


F_JSON = ("cau_hinh", "quy_doi", "sua_sku", "sua_gt", "dx_duyet", "quy_tac_kt")
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
    ss.map_ten = st_.get("map_ten", "tat") if st_.get("cai_dat_ver", 1) >= 2 else "tat"
    ss.luc_luu = st_.get("luc_luu", "")
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
    ss.pop("xung_dot", None)
    ss.ver_nap = ss.get("ver_nap", 0) + 1  # đổi khoá bảng chọn ngành sau mỗi lần tải lại
    bump()


def settings_bytes() -> bytes:
    return json_to_bytes({"sua": {"\t".join(k): v for k, v in ss.sua.items()},
                          "don_vi": {"\t".join(k): v for k, v in ss.dv.items() if v},
                          "rong": {k: v for k, v in ss.rong.items() if v and v[0] != C.HD_GIU},
                          "map_ten": ss.get("map_ten", "tat"), "cai_dat_ver": 2, "luc_luu": C.bay_gio(),
                          "chon_nganh": ss.get("chon_nganh", {}),
                          "nguoi_luu": ss.user})


# Cách GỘP khi file bị máy khác sửa cùng lúc (xem dong_bo.py). File không có ở đây -> hỏi người dùng.
def kieu_gop(path: str) -> str | None:
    ten = path.rsplit("/", 1)[-1]
    if path.startswith(DX_DIR + "/"):
        return "ds_id"
    if path.startswith("shared/"):
        return {"cau_hinh.json": "cau_hinh", "quy_doi_filter.json": "dict", "sua_sku.json": "dict",
                "sua_gia_tri.json": "dict", "de_xuat_duyet.json": "dict", "quy_tac_kiem_tra.json": "dict", "map_tskt.parquet": "bang:cate,prop_id",
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
    if ok:
        for p, d in files.items():
            goc[p] = (git_sha(d), d)
        ss.luc_luu = C.bay_gio()
        x = ss.get("xung_dot")
        if x and set(x["paths"]) <= set(files):  # chỉ gỡ xung đột khi chính các file đó đã lưu được
            ss.pop("xung_dot", None)
        ss.chua_luu = bool(ss.get("xung_dot"))
        st.toast(f"💾 {msg}", icon="✅")
        for g in ghi_chu:
            st.toast("🔀 " + g, icon="ℹ️")
    else:
        ss.chua_luu = True
        if ss.get("xung_dot"):
            st.rerun()  # khung xử lý xung đột vẽ 1 lần ở đầu trang (tránh trùng widget)
        st.error(f"⚠️ CHƯA LƯU ĐƯỢC: {msg}")
    return ok


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


def tao_ai() -> AIH.AI:
    prov = ss.get("ai_prov_tam") or sec("AI_PROVIDER", "groq")
    key = ss.get("ai_key_tam") or sec("AI_API_KEY", "") or sec(f"{str(prov).upper()}_API_KEY", "")
    return AIH.AI(prov, key, ss.get("ai_model_tam") or sec("AI_MODEL", ""), sec("AI_BASE_URL", ""))


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


def thanh_ben() -> str:
    with st.sidebar:
        st.markdown(f"### 🧩 PIM Tool\n👤 **{ss.ten}** (`{ss.user}`){' · 🛡️ admin' if ss.admin else ''}")
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
        trang = st.radio("Đi tới", [t for t in TRANG if ss.admin or t != "👥 Quản trị"], key="trang",
                         label_visibility="collapsed")
        if ss.admin and so_cho_duyet():
            st.warning(f"📥 **{so_cho_duyet()} đề xuất sửa CMS chờ duyệt** → trang 📮")
        dm = dem_dx(ss.get("dx_rieng", []))
        if dm:
            st.caption(f"📮 Đề xuất của {'bạn' if ss.ws == ss.user else ss.ws}: ⏳ {dm.get(C.DX_CHO, 0)} chờ · "
                       f"✅ {dm.get(C.DX_DUYET, 0)} duyệt · ❌ {dm.get(C.DX_TU_CHOI, 0)} từ chối")
        st.divider()
        S = tao_store()
        if S.backend == "local":
            st.warning("Chưa cấu hình GITHUB_TOKEN → đang lưu TẠM trên máy chủ (mất khi app khởi động lại).")
        st.caption(f"Lưu trữ: {S.mo_ta}")
        st.caption(f"Lần lưu gần nhất: {ss.get('luc_luu') or '—'}")
        if ss.get("chua_luu"):
            st.error("Có thay đổi CHƯA lưu.")
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
        st.caption(APP_VERSION)
    return trang


# ============================================================================
# TRANG: TỔNG QUAN
# ============================================================================
def trang_tong_quan() -> None:
    st.title("🏁 Làm nhanh — 3 bước")
    c = st.columns(6)
    c[0].metric("SKU trong IMPORT", f"{len(ss['import']):,}")
    c[1].metric("Dòng DATA SP", f"{len(ss.data_sp):,}")
    c[2].metric("SKU có spec PIM cũ", f"{ss.spec.sku.nunique():,}" if len(ss.spec) else "0")
    c[3].metric("Ngành hàng đã map", len(ss.bang))
    c[4].metric("Ô sửa tay", len(ss.sua))
    c[5].metric("Ngành có cấu hình", len(ss.cau_hinh))
    if ss.admin and so_cho_duyet():
        st.warning(f"📥 Có **{so_cho_duyet()} đề xuất sửa dữ liệu CMS** đang chờ bạn duyệt — trang 📮 → tab Duyệt.")
    if not ss.cau_hinh or not len(ss.map_tskt):
        st.info("Lần đầu dùng (admin): ở bước ① nạp file workspace theo mẫu và tick **cập nhật mapping / cấu hình / "
                "DATA PIM dùng chung**.")
    for m in ss.pop("flash", []) or []:
        st.success(m)
    xong1 = len(ss.data_sp) > 0 and len(ss["import"]) > 0
    with st.expander(f"{'✅' if xong1 else '①'} Bước 1 — Nạp dữ liệu (1 cục, tự nhận, tự lọc)", expanded=not xong1):
        khu_nap_nhanh("ln")
    xong2 = not bang_trong()
    with st.expander(f"{'✅' if xong2 else '②'} Bước 2 — Map + QC", expanded=xong1):
        c = st.columns([1, 3])
        if c[0].button("② Map dữ liệu", type="primary", disabled=not xong1, key="ln_nut_map", width="stretch"):
            chay_map_ui()
        if ss.meta.get("luc"):
            c[1].caption(f"Map lần cuối: {ss.meta['luc']} — chi tiết từng vùng ở trang 🚀 Map & kiểm tra.")
        if xong2:
            tab_qc(kq())
    with st.expander("③ Bước 3 — Xuất file import (+ file xin data CMS)", expanded=xong2):
        if xong2:
            xuat_gon("ln")
    if ss.lich_su:
        st.markdown("#### Lần xuất gần đây")
        st.dataframe(pd.DataFrame(ss.lich_su[-10:][::-1]), hide_index=True)


def xuat_gon(key: str, chon: list | None = None, canh: list | None = None) -> None:
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
    xn = True
    if canh:
        xn = st.checkbox("Tôi đã đọc cảnh báo, xuất file", value=False, key=f"{key}_xn")
    if st.button("📤 Tạo file import", type="primary", disabled=not xn or not chon, key=f"{key}_tao"):
        with st.spinner("Đang tạo file…"):
            x = C.xuat_file_import(ss.bang, ss["import"], ss.sua, ss.dv, ss.rong, bo_cot_sku=not giu_sku,
                                   chi_cate=chon, bo_dong_trong=bo_trong)
        ss.xuat = x
        ss.lich_su.append({"Lúc": C.bay_gio(), "Người xuất": ss.user, "Workspace": ss.ws,
                           "File": ", ".join(f[0] for f in x["files"]), "Số dòng": sum(f[2] for f in x["files"]),
                           "Sửa tay": x.get("so_o_sua", 0), "Thêm đơn vị": x.get("so_o_dv", 0),
                           "Biến đổi": x.get("so_o_bd", 0), "Không/Đang cập nhật": x.get("so_o_rong", 0),
                           "Tách xin data": x.get("bo_trong", 0), "Cảnh báo": " | ".join(canh or [])})
        luu(["lich_su"], f"Xuất {len(x['files'])} file import")
    x = ss.get("xuat")
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
                "hoặc dán bảng copy từ Excel. Tool <b>tự nhận loại file, tự nhận cột</b> (tên tiếng Việt/Anh, có/không "
                "dấu, thiếu tiêu đề thì đoán theo nội dung), <b>tự lọc</b> DATA SP theo SKU cần làm.</div>",
                unsafe_allow_html=True)
    for m in ss.pop("flash", []) or []:
        st.success(m)
    lan = ss.get(f"{key}_lan", 0)  # đổi khoá sau mỗi lần nạp -> ô chọn file / ô dán tự trống (không nạp lặp)
    fs = st.file_uploader("File (.xlsx / .xlsm / .csv) — chọn được nhiều file", type=["xlsx", "xlsm", "xls", "csv"],
                          accept_multiple_files=True, key=f"{key}_f{lan}")
    txt = st.text_area("Hoặc dán bảng từ Excel (model / SKU / biến thể…, có hoặc không có dòng tiêu đề)",
                       height=90, key=f"{key}_t{lan}", placeholder="219463\t\t4844439000045\n219464\tV2\t4844439000046")
    sig = tuple((f.name, f.size) for f in fs or []) + ((hash(txt),) if txt.strip() else ())
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
            if txt.strip():
                r = C.doc_mot_cuc(C.doc_text_dan(txt))
                kq.append(("(bảng dán)", {"loai": "sku" if len(r["import"]) else None, **r,
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
        elif r.get("loai") == "nganh":
            nd = " · ".join(f"{v['ten']} ({c}): {len(v['cot'])} cột" for c, v in r["cau_hinh"].items())
        elif r.get("loai") == "sku":
            nd = (f"{len(r['import']):,} SKU · {(r['import'].model_code != '').sum():,} có model · "
                  f"{(r['import'].variant_code != '').sum():,} có biến thể · {len(r['spec']):,} ô TSKT/FILTER")
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
                st.info(f"🏷️ Ngành **{v['ten']} ({c})** đã có cấu hình ({len(cu.get('cot', []))} cột) → file mới "
                        f"{len(v['cot'])} cột: thêm {len(them)} · bỏ {len(bo)}" + (f" ({', '.join(bo[:6])}…)" if bo else ""))
            else:
                st.info(f"🏷️ Ngành **{v['ten']} ({c})** chưa có trong tool → sẽ thêm mới với {len(v['cot'])} cột thuộc tính. "
                        "Sau đó nạp mapping TSKT/FILTER của ngành này (file CMS/mapping) để map được dữ liệu.")
        lay_ng = st.checkbox("Cập nhật Cấu hình ngành hàng từ file mẫu ngành (dùng chung)", value=duoc_sua_chung(),
                             disabled=not duoc_sua_chung(), key=f"{key}_ng")
        if not duoc_sua_chung():
            st.caption("Chỉ admin cập nhật cấu hình dùng chung.")
    if co_mau:
        lay_chung = st.checkbox("File mẫu: cập nhật mapping / cấu hình / DATA PIM dùng chung (gộp theo ngành)",
                                value=False, disabled=not duoc_sua_chung(), key=f"{key}_chung",
                                help="Chỉ admin. Ngành có trong file thay mapping + cấu hình của ngành đó; ngành khác giữ.")
    if not st.button("✔ Nạp vào tool", type="primary", key=f"{key}_ok"):
        return
    imp, spec, sp = [], [], []
    da_ng = []
    if ds_ng and lay_ng and duoc_sua_chung():
        nap_shared()
        ch = dict(ss.cau_hinh)
        ch.update(moi_ng)
        ss.cau_hinh = ch
        luu(["shared:cau_hinh"], "Nạp cấu hình ngành từ file mẫu: " + ", ".join(t for t, _ in ds_ng))
        da_ng = [f"{v['ten']} ({c}) {len(v['cot'])} cột" for c, v in moi_ng.items()]
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
    if loc and len(ss["import"]) and len(ss.data_sp):
        ss.data_sp, tk = C.loc_data_sp(ss.data_sp.astype(object), ss["import"])
        if tk["bo"]:
            bao.append(f"đã lọc bỏ {tk['bo']:,} dòng DATA SP của {tk['sku_bo']:,} SKU không có trong IMPORT")
        if "data_sp" not in phan:
            phan.append("data_sp")
    if "data_sp" in phan:
        ss.data_sp = C.nen_df(ss.data_sp)
        bao.append(f"DATA SP {len(ss.data_sp):,} dòng")
    if da_ng:
        bao.insert(0, "Cấu hình ngành " + "; ".join(da_ng))
        if not phan:
            bump()
            ss.flash = ["✔ Đã nạp: " + " · ".join(bao)]
            ss.pop(f"{key}_sig", None)
            ss[f"{key}_lan"] = lan + 1
            if len(ss["import"]) and len(ss.data_sp):
                chay_map_ui()
            st.rerun()
    bump()
    if phan and luu(phan + ["settings"], "Nạp nhanh: " + ", ".join(t for t, _ in kq)):
        ss.flash = ["✔ Đã nạp: " + " · ".join(bao)]
        ss.pop(f"{key}_sig", None)
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
    x = file_xin_data()
    n, nm = len(x["sku"]), int((x["model"]["Tình trạng"].str.startswith("MODEL")).sum()) if len(x["model"]) else 0
    if not n:
        st.caption("✔ Mọi SKU đều có giá trị — không cần xin thêm data CMS.")
        return
    st.download_button(f"📨 Tải file xin data CMS ({n:,} SKU · {nm:,} model không có giá trị)",
                       C.xlsx_nhieu_sheet({"XIN DATA CMS": x["sku"], "THEO MODEL": x["model"]}),
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


def trang_nap() -> None:
    st.title("📥 Nạp dữ liệu → Map → Kiểm tra đối chiếu")
    ph_so = st.container()  # ô số liệu: vẽ SAU CÙNG để phản ánh dữ liệu vừa nạp trong cùng lượt

    def ve_so() -> None:
        with ph_so:
            c = st.columns(6)
            c[0].metric("SKU trong IMPORT", f"{len(ss['import']):,}")
            c[1].metric("Dòng DATA SP", f"{len(ss.data_sp):,}")
            c[2].metric("Ngành có cấu hình", len(ss.cau_hinh))
            c[3].metric("Mapping TSKT", f"{len(ss.map_tskt):,}")
            c[4].metric("Mapping FILTER", f"{len(ss.map_filter):,}")
            c[5].metric("Option DATA PIM", f"{len(ss.data_pim):,}")
    st.markdown("<div class='buoc'><b>① Nạp</b> (lô hoặc data gốc) → <b>② bấm Map</b> → <b>③ kiểm tra đối chiếu ngay "
                "bên dưới</b>, thêm đơn vị / gộp kích thước, rồi sang trang 📤 Xuất.</div>", unsafe_allow_html=True)
    t1, t2, t3 = st.tabs(["① Nạp dữ liệu lô", "🗄️ Data gốc (TSKT · FILTER · DATA PIM · cấu hình)", "👀 Xem dữ liệu đã nạp"])
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
    cm = st.columns([1.2, 3])
    if cm[0].button("🚀 Map dữ liệu", type="primary", disabled=not xong1, key="nap_nut_map", width="stretch"):
        chay_map_ui()
    cm[1].caption(f"Map lần cuối: {ss.meta['luc']}" if ss.meta.get("luc") else
                  ("Bấm Map để sinh vùng kiểm tra đối chiếu." if xong1 else "Cần có IMPORT và DATA SP trước."))
    if bang_trong():
        ve_so()
        return
    st.markdown("### ③ Kiểm tra & đối chiếu")
    k = kq()
    the_so(k)
    tb = st.tabs(["🛡️ QC tổng hợp", "🧾 Đối soát CMS → kết quả", "📏 Đơn vị & kích thước (dài · rộng · cao)",
                  "📐 Gộp / tách kích thước", "🚫 Không / Đang cập nhật", "≠ Khác spec PIM"])
    with tb[0]:
        tab_qc(k)
    with tb[1]:
        tab_doi_soat()
    with tb[2]:
        tab_don_vi(k)
    with tb[3]:
        tab_tach_kt()
    with tb[4]:
        tab_rong(k)
    with tb[5]:
        tab_khac(k)
    st.caption("Đủ bộ tab (cảnh báo, gợi ý AI, theo SKU…) ở trang 🚀 Map & kiểm tra. Kiểm tra xong → trang 📤 Xuất file import.")
    ve_so()


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


def chay_map_ui() -> None:
    if not len(ss.data_sp) or not len(ss["import"]):
        st.error("Cần nạp DATA SP và IMPORT trước (trang 📥 Nạp dữ liệu lô).")
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
               "chon": r["chon"]}
    bump()
    luu(["ket_qua", "settings"], f"Map {r['tom_tat']['so_sku']} SKU / {r['tom_tat']['so_nganh']} ngành hàng")
    ss.map_msg = (f"✔ Map xong {r['tom_tat']['so_sku']:,} SKU · {r['tom_tat']['so_nganh']} ngành hàng · "
                  f"{r['tom_tat']['so_o']:,} ô · {time.time() - t:.1f} giây"
                  + (f" · {r['tom_tat']['map_theo_ten']:,} ô map theo tên" if r['tom_tat']['map_theo_ten'] else ""))
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
    for hang in (items[:6], items[6:]):
        cols = st.columns(6)
        for c, (t, v, ic) in zip(cols, hang):
            c.metric(f"{ic or ''} {t}".strip(), f"{v:,}".replace(",", ".") if isinstance(v, int) else v)


def trang_map() -> None:
    st.title("🚀 Map & kiểm tra")
    c1, c2, c3 = st.columns([2, 3, 2])
    with c1:
        if st.button("① Map dữ liệu", type="primary", width="stretch"):
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
    k = kq()
    the_so(k)
    c = st.columns([1.3, 1.3, 3])
    if c[0].button("🪄 Ô tool trống → lấy PIM cũ", width="stretch",
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
    tabs = st.tabs(["🛡️ QC tổng hợp", "⚠️ Cảnh báo", "📈 Độ hoàn thiện & quy tắc", "🧾 Đối soát CMS → kết quả",
                    "🤖 Gợi ý thông minh & AI", "≠ Khác spec PIM (sửa)", "🔎 Theo SKU + FILTER",
                    "📏 Đơn vị & biến đổi hàng loạt", "🚫 Không / Đang cập nhật", "📐 Gộp / tách kích thước",
                    "🧩 Thuộc tính chưa map", "📜 Log map"])
    with tabs[0]:
        tab_qc(k)
    tabs = list(tabs[1:])
    tabs = [tabs[0]] + list(tabs[2:]) + [tabs[1]]
    with tabs[-1]:
        tab_hoan_thien()
    with tabs[0]:
        tab_canh_bao(k)
    with tabs[1]:
        tab_doi_soat()
    with tabs[2]:
        tab_ai(k)
    with tabs[3]:
        tab_khac(k)
    with tabs[4]:
        tab_sku(k)
    with tabs[5]:
        tab_don_vi(k)
    with tabs[6]:
        tab_rong(k)
    with tabs[7]:
        tab_tach_kt()
    with tabs[8]:
        tab_chua_map()
    with tabs[9]:
        lg = pd.DataFrame(ss.meta.get("log", []), columns=["SKU", "Ngành hàng", "Mã", "Nguyên nhân"])
        st.dataframe(lg, hide_index=True, height=420)


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
    q = [
        ("CAO", "Thiếu model_code", s.get("thieu_model", 0), "SKU không có Mã model → không import được",
         "Nạp file export PIM / sửa IMPORT (trang 📥)", None),
        ("CAO", "Thiếu category_code", s.get("thieu_cate", 0), "IMPORT chưa có Mã danh mục PIM",
         "Nạp file export PIM (có category_code) hoặc điền ở bảng IMPORT", None),
        ("CAO", "ERP / model KHÔNG CÓ GIÁ TRỊ", len(xd["sku"]),
         "CMS chưa có data, ngành chưa có mapping, hoặc không thuộc tính nào map được",
         "Tải file xin data CMS; trang Xuất có tuỳ chọn tách các SKU này khỏi file import", "xin_data"),
        ("CAO", "Mapping trỏ tới cột không có trong cấu hình", nc,
         "Giá trị CMS đã map nhưng ngành không có cột đó → MẤT khi xuất", "Thêm cột thiếu vào cấu hình (admin)",
         "them_cot" if nc else None),
        ("CAO", "FILTER không khớp option", nf, "Giá trị CMS không trùng tên option nào trong DATA PIM → ô trống",
         "🧾 Đối soát → 🔁 Quy đổi FILTER (có gợi ý option gần giống)", None),
        ("CAO", "FILTER chứa chữ (sai mã)", s.get("filter_chu", 0), "PIM chỉ nhận mã option số",
         "🔎 Theo SKU + FILTER → chọn lại option", None),
        ("CAO", "Vi phạm quy tắc mức Lỗi", int((h["vi_pham"]["Mức"] == C.MUC_LOI).sum()) if len(h["vi_pham"]) else 0,
         "Không đạt quy tắc chất lượng admin đặt", "📈 Độ hoàn thiện & quy tắc", None),
        ("TB", "Ô KHÁC spec PIM", s.get(C.TRANG_THAI_KHAC, 0) if co else 0, "Import sẽ ghi đè giá trị đang trên web",
         "≠ Khác spec PIM: giữ TOOL MỚI hoặc tick Lấy PIM cũ", None),
        ("TB", "Tool trống nhưng PIM đang có", s.get(C.TRANG_THAI_TOOL_TRONG, 0) if co else 0,
         "CMS không có giá trị, PIM cũ có", "Lấy PIM cũ cho các ô này", "lay_pim" if co else None),
        ("TB", "Kích thước/khối lượng chưa có đơn vị", s.get("chua_don_vi", 0), "Ô số trơn ở cột kích thước",
         "Điền đơn vị gợi ý cho mọi cột còn số trơn (rule cũ)", "don_vi"),
        ("TB", "Giá trị Không / Đang cập nhật chưa chọn cách xử lý", rong_chua,
         "Mặc định GIỮ nguyên khi import", "🚫 Không / Đang cập nhật: Giữ / Để trống / Thay bằng", "rong" if rong_chua else None),
        ("TB", "Nghi sai (kiểm tra thông minh)", int(ttm_.muc_do.isin([C.MUC_CAO, C.MUC_TB]).sum()) if len(ttm_) else 0,
         "Lỗi gõ, lẫn đơn vị, giá trị bất thường", "Áp các gợi ý mức Cao (có giá trị sửa)",
         "ttm" if len(ttm_) and ((ttm_.muc_do == C.MUC_CAO) & (ttm_.goi_y != "")).any() else None),
        ("TB", "SKU chưa đủ cột bắt buộc", int((~h["sku"]["Đủ bắt buộc"]).sum()) if len(h["sku"]) else 0,
         "Theo quy tắc Bắt buộc", "📈 Độ hoàn thiện & quy tắc", None),
        ("TB", "model_code lặp nhiều dòng MODEL", len(C.model_trung(ss.bang)), "PIM sẽ lấy dòng sau cùng",
         "Kiểm tra IMPORT có thiếu variant_code", None),
        ("THAP", "Thuộc tính CMS chưa có mapping", len(ss.meta.get("chua_map", [])), "Không vào file import",
         "🧩 Thuộc tính chưa map → thêm vào MAPPING", None),
        ("TT", "SKU không có trong file PIM cũ", s.get("khong_co_pim", 0) if co else 0, "Không đối chiếu được spec",
         "Bình thường với SKU mới", None),
    ]
    return [x for x in q if x[2]]


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
        st.success("✔ Không còn mục nào cần xử lý — sang 📤 Xuất file.")
        return
    st.dataframe(pd.DataFrame([{"Mức": MUC_ICON[a], "Vùng": b, "Số lượng": n_, "Ý nghĩa": y, "Gợi ý xử lý": g,
                                "Sửa nhanh": "⚡" if act else ""} for a, b, n_, y, g, act in q]),
                 hide_index=True, width="stretch",
                 column_config={"Số lượng": st.column_config.NumberColumn(format="%d")})
    acts = {x[5] for x in q if x[5]}
    if not acts:
        return
    st.markdown("###### ⚡ Sửa nhanh (bấm là áp, có thể hoàn tác ở từng tab)")
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
    if "lay_pim" in acts and nut("🪄 Ô tool trống → lấy PIM cũ", "lay_pim"):
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
    if "rong" in acts:
        if nut("🚫 Không/Đang cập nhật → Để trống (không cập nhật)", "rong_trong"):
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
    c[3].download_button("📊 Tải báo cáo đối soát", C.xlsx_nhieu_sheet(
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
    hien = df.copy()
    hien.insert(0, "Áp dụng", False)
    ed = st.data_editor(hien, hide_index=True, height=400, key=f"{key}_{ss.ver}",
                        disabled=[c for c in hien.columns if c not in ("Áp dụng", cot_goi_y)],
                        column_config={"Áp dụng": st.column_config.CheckboxColumn(width="small"), "_cate": None,
                                       cot_goi_y: st.column_config.TextColumn(width="medium"),
                                       "AI": st.column_config.TextColumn("🤖 AI nhận xét", width="medium")})
    c = st.columns([1, 1, 3])
    tat_ca = c[1].button("✔ Áp dụng TẤT CẢ dòng có gợi ý", key=f"{key}_all")
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


def tab_ai(k: dict) -> None:
    con = st.radio("Chọn", ["🧠 Kiểm tra thông minh (miễn phí, không cần AI)", "🤖 AI rà từng SKU", "💬 Hỏi AI"],
                   horizontal=True, label_visibility="collapsed", key="ai_che_do")
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
        st.caption("AI miễn phí, chỉ gửi thông số sản phẩm (không có dữ liệu cá nhân). Admin cấu hình cố định trong "
                   "Secrets (AI_PROVIDER, AI_API_KEY, AI_MODEL); hoặc dán key tạm cho riêng phiên này:")
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
    hien = pd.DataFrame({
        "SKU": v.sku, "NH": v.cate, "Mã TSKT": v.ma, "Tên": v.ten, "PIM cũ": v.pim_cu,
        "TOOL MỚI": v.tool_moi, "Giải nghĩa FILTER": [C.giai_nghia_filter(a, b, ss.opt) for a, b in zip(v.ma, v.tool_moi)],
        "Trạng thái": v.trang_thai, "Lấy PIM cũ": False, "✎": v.da_sua.map(lambda x: "✎" if x else "")})
    ed = st.data_editor(hien, hide_index=True, height=460, key=f"ed_khac_{ss.ver}",
                        disabled=[c for c in hien.columns if c not in ("TOOL MỚI", "Lấy PIM cũ")],
                        column_config={"Lấy PIM cũ": st.column_config.CheckboxColumn(width="small"),
                                       "TOOL MỚI": st.column_config.TextColumn(width="medium")})
    if st.button("✔ Áp dụng thay đổi", type="primary", key="ap_khac"):
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


def tab_don_vi(k: dict) -> None:
    st.markdown("##### ① Đơn vị theo cột (đúng rule cũ 66.py: chỉ thêm vào ô SỐ TRƠN, không đụng FILTER)")
    ds = pd.DataFrame(k["don_vi_cot"])
    if len(ds):
        tat_ca = st.checkbox("Hiện tất cả cột TSKT (không chỉ cột kích thước/khối lượng)", key="dv_tat_ca")
        if not tat_ca:
            ds = ds[ds.la_kt | (ds.da_luu != "") | ds.apply(lambda r: bool(ss.dv.get((r.cate, r.code))), axis=1)]
        ds = ds.sort_values(["cate", "so_tron"], ascending=[True, False]).reset_index(drop=True)
        hien = pd.DataFrame({"NH": ds.cate, "Mã TSKT": ds.code, "Tên": ds.ten,
                             "Số trơn / có dữ liệu": [f"{a} / {b}" for a, b in zip(ds.so_tron, ds.tong)],
                             "Giá trị hiện tại (sau áp)": ds.vi_du, "Gợi ý": ds.goi_y,
                             "ĐƠN VỊ": [ss.dv.get((a, b), "") if isinstance(ss.dv.get((a, b), ""), str) else ""
                                        for a, b in zip(ds.cate, ds.code)]})
        c = st.columns([1.4, 1, 3])
        if c[0].button("✨ Điền gợi ý (cột còn số trơn, chưa chọn)"):
            for a, b, g, n in zip(ds.cate, ds.code, ds.goi_y, ds.so_tron):
                if n and g and not ss.dv.get((a, b)):
                    ss.dv[(a, b)] = g
            bump()
            luu(["settings"], "Điền đơn vị theo gợi ý")
            st.rerun()
        c[2].caption("Gõ đơn vị TUỲ Ý vào cột ĐƠN VỊ (cm, mm, kg, g, inch, W, mAh, lít, giờ…). Ô đã có chữ giữ "
                     "nguyên. Bấm Áp để xem kết quả ngay.")
        ed = st.data_editor(hien, hide_index=True, height=min(460, 40 + 35 * len(hien)), key=f"ed_dv_{ss.ver}",
                            disabled=[x for x in hien.columns if x != "ĐƠN VỊ"],
                            column_config={"ĐƠN VỊ": st.column_config.TextColumn(width="small",
                                                                                help="vd: " + ", ".join(DV_OPTIONS[1:]))})
        if st.button("▶ Áp đơn vị & xem lại", type="primary"):
            for i in range(len(ds)):
                v = C.chuan_hoa_key(ed.at[i, "ĐƠN VỊ"] or "")
                key = (ds.cate[i], ds.code[i])
                if v:
                    ss.dv[key] = v
                else:
                    ss.dv.pop(key, None)
            bump()
            luu(["settings"], "Áp đơn vị hàng loạt")
            st.rerun()
    st.divider()
    khu_bien_doi()


def khu_bien_doi() -> None:
    st.markdown("##### ② Biến đổi hàng loạt (thêm đơn vị / chữ, đổi đơn vị mm→cm, thay chữ, làm tròn)")
    st.caption("Áp sau đơn vị ở ①, trước khi xuất. Không bao giờ áp vào cột FILTER. Ô sửa tay và quy tắc "
               "Không/Đang cập nhật được ưu tiên hơn.")
    if bang_trong():
        return
    cot = sorted({(c, m) for c, b in ss.bang.items() for m in C.cot_tt(b) if not C.la_cot_filter(m)})
    ten = {(c, m): ss.bang[c]["ten"].get(m, "") for c, m in cot}
    c1, c2 = st.columns([3, 2])
    chon = c1.multiselect("Cột áp dụng", cot, key="bd_cot",
                          format_func=lambda x: f"{x[0]} · {x[1]} — {ten.get(x, '')}")
    moi_nganh = c2.checkbox("Áp cho mã cột này ở MỌI ngành", key="bd_moi_nganh",
                            help="Lưu theo mã cột (*) thay vì từng ngành — tiện cho cột dùng chung như mass_tskt_master")
    c = st.columns([2, 2, 1.2, 1.2, 1])
    kieu = c[0].selectbox("Kiểu", list(C.BD_KIEU), format_func=C.BD_KIEU.get, key="bd_kieu")
    pv = c[1].selectbox("Phạm vi", list(C.BD_PHAM_VI), format_func=C.BD_PHAM_VI.get, key="bd_pv",
                        index=2 if kieu in ("thay",) else 0)
    nhan_a = {"them_sau": "Chữ / đơn vị", "them_truoc": "Chữ", "ca_hai": "Chữ phía TRƯỚC", "doi_dv": "Từ đơn vị",
              "thay": "Tìm chữ", "lam_tron": "Số chữ số lẻ"}[kieu]
    a = c[2].text_input(nhan_a, key="bd_a", placeholder={"them_sau": "kg", "doi_dv": "mm", "lam_tron": "1",
                                                         "ca_hai": "Khoảng"}.get(kieu, ""))
    b = c[3].text_input({"doi_dv": "Sang đơn vị", "ca_hai": "Chữ phía SAU"}.get(kieu, "Thay bằng"), key="bd_b",
                        disabled=kieu not in ("doi_dv", "thay", "ca_hai"),
                        placeholder={"doi_dv": "cm", "ca_hai": "cm"}.get(kieu, ""))
    he_so = c[4].text_input("Hệ số", key="bd_hs", disabled=kieu != "doi_dv",
                            help="Để trống nếu là cặp quen thuộc (mm↔cm↔m, g↔kg, inch→cm, ml↔lít, W↔kW, mAh↔Ah, phút↔giờ)")
    buoc = {"kieu": kieu, "pham_vi": pv, "a": a, "b": b, "he_so": he_so}
    if chon and (a or b or kieu == "thay"):
        mau = []
        for cc, m in chon:
            for r in ss.bang[cc]["rows"]:
                v0, _ = C.bien_doi_o(cc, r["sku"], m, r["vals"].get(m, ""), ss.sua, ss.dv, ss.rong)
                v1 = C.ap_buoc(v0, buoc)
                if v0 and v1 != v0:
                    mau.append({"Cột": m, "SKU": r["sku"], "Trước": v0, "Sau": v1})
        st.caption(f"Xem trước: **{len(mau):,} ô** sẽ đổi" + (" — hiện 30 ô đầu" if len(mau) > 30 else ""))
        if mau:
            st.dataframe(pd.DataFrame(mau[:30]), hide_index=True, height=min(300, 40 + 35 * min(30, len(mau))))
        if st.button(f"▶ Áp cho {len(chon)} cột", type="primary", disabled=not mau):
            for cc, m in chon:
                kk = ("*" if moi_nganh else cc, m, "bd")
                ss.dv[kk] = list(ss.dv.get(kk) or []) + [buoc]
            bump()
            luu(["settings"], f"Biến đổi hàng loạt {len(chon)} cột: {C.BD_KIEU[kieu]}")
            st.rerun()
    dang = [(k, v) for k, v in ss.dv.items() if len(k) == 3 and v]
    if dang:
        st.markdown("###### Biến đổi đang áp")
        rows = [{"Ngành": k[0], "Mã cột": k[1], "Bước": " → ".join(
            f"{C.BD_KIEU.get(x['kieu'], x['kieu'])} [{x.get('a', '')}{(' → ' + x['b']) if x.get('b') else ''}] "
            f"({C.BD_PHAM_VI.get(x.get('pham_vi', 'so'), '')[:18]})" for x in v), "Bỏ": False} for k, v in dang]
        ed = st.data_editor(pd.DataFrame(rows), hide_index=True, key=f"ed_bd_{ss.ver}",
                            disabled=["Ngành", "Mã cột", "Bước"])
        if st.button("🗑 Bỏ biến đổi đã tick") and ed["Bỏ"].any():
            for r in ed[ed["Bỏ"]].to_dict("records"):
                ss.dv.pop((r["Ngành"], r["Mã cột"], "bd"), None)
            bump()
            luu(["settings"], "Bỏ biến đổi hàng loạt")
            st.rerun()


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
        for i, d in enumerate(ds):
            hd = nhan_sang_hd.get(ed.at[i, "XỬ LÝ"], C.HD_GIU)
            if hd == C.HD_GIU:
                ss.rong.pop(d["khoa"], None)
            else:
                ss.rong[d["khoa"]] = [hd, C.chuan_hoa_key(ed.at[i, "Thay bằng"])]
        if them.strip():
            ss.rong.setdefault(C.khoa_gia_tri_rong(them), [C.HD_TRONG, ""])
        thieu = [kk for kk, v in ss.rong.items() if v[0] == C.HD_THAY and not v[1]]
        if thieu:
            st.error("Chưa nhập 'Thay bằng' cho: " + ", ".join(thieu))
            return
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
def trang_xuat() -> None:
    st.title("📤 Xuất file import")
    if bang_trong():
        st.info("Chưa có kết quả map — vào 🚀 Map & kiểm tra.")
        return
    k = kq()
    the_so(k)
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
    canh = [f"• {n:,} {t}" for n, t in canh if n]
    if canh:
        st.markdown("<div class='canh'><b>⚠️ CÒN CẢNH BÁO — đọc trước khi import:</b><br>" + "<br>".join(canh) +
                    "</div>", unsafe_allow_html=True)
    else:
        st.success("✔ Không còn cảnh báo.")
    st.markdown("##### CHỌN NGÀNH HÀNG xuất (giống sheet CHỌN NGÀNH HÀNG của file mẫu)")
    chon = bang_chon_nganh(f"ed_chon_xuat_{ss.ws}_{ss.get('ver_map', 0)}_{ss.get('ver_nap', 0)}_"
                           f"{hash(tuple(ss.bang)) & 0xffffff}")
    xuat_gon("xp", chon, canh)
    x = ss.get("xuat")
    if x:
        L = ds()["loi"]
        bc = C.xlsx_nhieu_sheet({"CẢNH BÁO": k["canh_bao"], "KHÁC SPEC PIM": pd.DataFrame(k["khac"]),
                                 "ĐỘ HOÀN THIỆN": dht()["sku"], "VI PHẠM QUY TẮC": dht()["vi_pham"],
                                 "ĐỐI SOÁT CMS": L.drop(columns=[c for c in L.columns if c.startswith("_")])
                                 if len(L) else L})
        st.download_button("📊 Tải báo cáo kiểm tra (.xlsx)", bc, file_name=f"KIEM_TRA_{x['stamp']}.xlsx")
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
    if cate != "*" and cate in ss.bang and c[1].button("🪄 Gợi ý từ dữ liệu lô", disabled=not sua_duoc):
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


def tab_cau_hinh(sua_duoc: bool) -> None:
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
    t = st.tabs(["Dùng hằng ngày", "Cài đặt (admin)", "Quy tắc map (giống desktop)"])
    for tab, f in zip(t, ("HUONG_DAN_SU_DUNG.md", "HUONG_DAN_CAI_DAT.md", "QUY_TAC_MAP.md")):
        with tab:
            p = os.path.join(goc, f)
            st.markdown(open(p, encoding="utf-8").read() if os.path.exists(p) else f"(Thiếu file {f})")


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


# ============================================================================
# CHẠY
# ============================================================================
dang_nhap()
if "cau_hinh" not in ss:
    with st.spinner("Đang tải dữ liệu dùng chung…"):
        nap_shared()
if ss.get("ws_da_nap") != ss.ws:
    with st.spinner(f"Đang tải workspace {ss.ws}…"):
        nap_workspace()
trang = thanh_ben()
hien_xung_dot()
{"🏁 Làm nhanh": trang_tong_quan, "📥 Nạp dữ liệu lô": trang_nap, "🚀 Map & kiểm tra": trang_map,
 "📤 Xuất file import": trang_xuat, "📮 Đề xuất sửa CMS": trang_de_xuat, "📘 Hướng dẫn": trang_huong_dan, "⚙️ Cấu hình & mapping": trang_cau_hinh, "🧰 Tra cứu": trang_tra_cuu,
 "👥 Quản trị": trang_quan_tri}[trang]()
