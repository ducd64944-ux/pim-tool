# -*- coding: utf-8 -*-
"""
gh_store.py — LƯU TRỮ dữ liệu lên GitHub (Streamlit Cloud không có ổ đĩa lâu dài).

  * Mỗi lần lưu = 1 COMMIT duy nhất cho nhiều file (Git Data API: blob -> tree -> commit -> cập nhật ref).
  * NHIỀU MÁY / NHIỀU NGƯỜI CÙNG LÚC — khoá lạc quan (optimistic locking):
      - lúc đọc, ghi nhận "phiên bản" (git blob sha) của từng file;
      - lúc lưu, truyền phiên bản đã đọc (ky_vong). Nếu trên kho file đã bị máy khác đổi -> KHÔNG ghi đè,
        trả về danh sách xung đột để app GỘP (dong_bo.py) hoặc hỏi người dùng;
      - kiểm tra lại ngay trước khi đẩy commit + cập nhật nhánh không-force -> không có khe hở ghi đè.
  * Chỉ đẩy file có thay đổi (so phiên bản đã đọc).
  * Đọc theo blob sha + bộ nhớ đệm theo NỘI DUNG (sha) dùng chung mọi phiên: 10 người mở app chỉ tải DATA PIM 1 lần.
  * Khoá trong tiến trình: các phiên trên cùng máy chủ commit lần lượt (giảm tranh chấp nhánh).
  * Không có token -> lưu thư mục cục bộ (chạy thử trên máy), cùng cơ chế phiên bản/xung đột.
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import random
import threading
import time
from collections import OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests

API = "https://api.github.com"
KHONG_CO = "∅"  # phiên bản của file CHƯA tồn tại


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def git_sha(data: Optional[bytes]) -> str:
    """Đúng sha mà git dùng cho blob -> so trực tiếp với sha trên GitHub."""
    if data is None:
        return KHONG_CO
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


def df_to_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.astype(str).reset_index(drop=True).to_parquet(buf, index=False, compression="zstd")
    return buf.getvalue()


def bytes_to_df(data: Optional[bytes], cols: Optional[List[str]] = None) -> pd.DataFrame:
    if not data:
        return pd.DataFrame(columns=cols or [])
    df = pd.read_parquet(io.BytesIO(data))
    df = df.astype(object).where(df.notna(), "")
    if cols:
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
    return df


def json_to_bytes(obj) -> bytes:
    return json.dumps(obj, ensure_ascii=False, indent=1, sort_keys=True).encode("utf-8")


def bytes_to_json(data: Optional[bytes], default=None):
    if not data:
        return default
    try:
        return json.loads(data.decode("utf-8"))
    except Exception:  # noqa: BLE001
        return default


class _BoNho:
    """Bộ nhớ đệm nội dung theo git sha (giới hạn dung lượng)."""

    def __init__(self, toi_da: int = 400 * 2 ** 20):
        self.toi_da, self.dung, self.d = toi_da, 0, OrderedDict()
        self.khoa = threading.Lock()

    def lay(self, k: str) -> Optional[bytes]:
        with self.khoa:
            v = self.d.get(k)
            if v is not None:
                self.d.move_to_end(k)
            return v

    def dat(self, k: str, v: bytes) -> None:
        if len(v) > self.toi_da // 4:
            return
        with self.khoa:
            if k in self.d:
                return
            self.d[k] = v
            self.dung += len(v)
            while self.dung > self.toi_da and self.d:
                _, x = self.d.popitem(last=False)
                self.dung -= len(x)


class Store:
    """backend 'github' (repo/branch/token) hoặc 'local' (thư mục)."""

    def __init__(self, backend: str, repo: str = "", branch: str = "main", token: str = "",
                 root: str = "data", local_dir: str = "_du_lieu_cuc_bo"):
        self.backend = backend
        self.repo, self.branch, self.token = repo, branch, token
        self.root = root.strip("/")
        self.local = Path(local_dir)
        self._bo_nho = _BoNho()
        self._khoa_ghi = threading.RLock()
        self._ds_tam: Dict[str, Tuple[float, Dict[str, str]]] = {}  # thư mục -> (lúc, {tên: sha})
        self.so_goi_api = 0

    # ---------------------------------------------------------------- tiện ích
    @property
    def mo_ta(self) -> str:
        if self.backend == "github":
            return f"GitHub: {self.repo}@{self.branch}/{self.root}"
        return f"THƯ MỤC CỤC BỘ: {self.local.resolve()} (chỉ để chạy thử — Streamlit Cloud sẽ xoá khi khởi động lại)"

    def _p(self, path: str) -> str:
        return f"{self.root}/{path}".strip("/")

    def _h(self, raw: bool = False) -> dict:
        return {"Authorization": f"Bearer {self.token}", "X-GitHub-Api-Version": "2022-11-28",
                "Accept": "application/vnd.github.raw+json" if raw else "application/vnd.github+json"}

    def _req(self, method: str, url: str, **kw) -> requests.Response:
        r = None
        for lan in range(4):
            try:
                self.so_goi_api += 1
                r = requests.request(method, API + url, timeout=120, **kw)
            except requests.RequestException:
                if lan == 3:
                    raise
                time.sleep(1.5 * (lan + 1))
                continue
            if r.status_code in (502, 503, 504) and lan < 3:
                time.sleep(1.5 * (lan + 1))
                continue
            if r.status_code in (403, 429) and "rate limit" in r.text.lower() and lan < 3:
                cho = int(r.headers.get("retry-after", "0") or 0) or 20
                time.sleep(min(cho, 60))
                continue
            return r
        return r  # pragma: no cover

    # ---------------------------------------------------------------- phiên bản
    def phien_ban_thu_muc(self, thu_muc: str, ref: Optional[str] = None, moi: bool = True) -> Dict[str, str]:
        """{tên file: git sha} của 1 thư mục (1 lần gọi API)."""
        if self.backend == "local":
            d = self.local / self._p(thu_muc)
            if not d.exists():
                return {}
            return {p.name: git_sha(p.read_bytes()) for p in d.iterdir() if p.is_file() and not p.name.endswith(".tmp")}
        key = f"{thu_muc}@{ref or ''}"
        tam = self._ds_tam.get(key)  # 1 lần đọc (phiên khác có thể xoá bộ đệm giữa chừng)
        if not moi and tam and time.time() - tam[0] < 3:
            return tam[1]
        r = self._req("GET", f"/repos/{self.repo}/contents/{self._p(thu_muc)}", headers=self._h(),
                      params={"ref": ref or self.branch})
        if r.status_code == 404:
            out: Dict[str, str] = {}
        elif r.status_code != 200:
            raise RuntimeError(f"Liệt kê {thu_muc} lỗi {r.status_code}: {r.text[:200]}")
        else:
            out = {x["name"]: x["sha"] for x in r.json() if isinstance(x, dict) and x.get("type") == "file"}
        self._ds_tam[key] = (time.time(), out)
        return out

    def phien_ban(self, paths: List[str], ref: Optional[str] = None) -> Dict[str, str]:
        """{path: git sha hoặc KHONG_CO} — gom theo thư mục."""
        theo_dir: Dict[str, List[str]] = {}
        for p in paths:
            d, _, ten = p.rpartition("/")
            theo_dir.setdefault(d, []).append(p)
        out = {}
        for d, ps in theo_dir.items():
            ds = self.phien_ban_thu_muc(d, ref)
            for p in ps:
                out[p] = ds.get(p.rpartition("/")[2], KHONG_CO)
        return out

    # ---------------------------------------------------------------- đọc
    def doc2(self, path: str) -> Tuple[Optional[bytes], str]:
        """-> (nội dung | None, phiên bản git sha | KHONG_CO)."""
        if self.backend == "local":
            f = self.local / self._p(path)
            data = f.read_bytes() if f.exists() else None
            return data, git_sha(data)
        d, _, ten = path.rpartition("/")
        s = self.phien_ban_thu_muc(d, moi=False).get(ten)
        if not s:
            return None, KHONG_CO
        data = self._bo_nho.lay(s)
        if data is None:
            r = self._req("GET", f"/repos/{self.repo}/git/blobs/{s}", headers=self._h(raw=True))
            if r.status_code != 200:
                raise RuntimeError(f"Đọc {path} lỗi {r.status_code}: {r.text[:200]}")
            data = r.content
            if git_sha(data) != s:  # phòng trường hợp API trả JSON thay vì raw
                try:
                    data = base64.b64decode(r.json()["content"])
                except Exception:  # noqa: BLE001
                    pass
            if git_sha(data) != s:  # không bao giờ dùng/nhớ nội dung không khớp phiên bản
                raise RuntimeError(f"Đọc {path}: nội dung không khớp phiên bản {s[:8]} — thử tải lại")
            self._bo_nho.dat(s, data)
        return data, s

    def doc(self, path: str) -> Optional[bytes]:
        return self.doc2(path)[0]

    def liet_ke(self, thu_muc: str) -> List[str]:
        return sorted(self.phien_ban_thu_muc(thu_muc))

    # ---------------------------------------------------------------- ghi
    def luu(self, files: Dict[str, Optional[bytes]], thong_diep: str,
            ky_vong: Optional[Dict[str, str]] = None) -> Tuple[bool, str, List[str], Dict[str, str]]:
        """files: {đường dẫn: bytes | None (xoá)}.
        ky_vong: {đường dẫn: phiên bản đã đọc}. File có ky_vong mà trên kho đã khác -> KHÔNG ghi gì cả,
        trả về xung đột {đường dẫn: phiên bản hiện tại} để gộp rồi lưu lại.
        -> (ok, thông báo, danh sách file đã ghi, xung đột)."""
        ky_vong = ky_vong or {}
        # bỏ file không đổi so với bản đã đọc
        doi = {p: d for p, d in files.items() if not (p in ky_vong and ky_vong[p] == git_sha(d))}
        if not doi:
            return True, "Không có thay đổi cần lưu.", [], {}
        kv = {p: ky_vong[p] for p in doi if p in ky_vong}
        with self._khoa_ghi:
            if self.backend == "local":
                hien = self.phien_ban(list(kv))
                xd = {p: hien[p] for p in kv if hien[p] != kv[p]}
                if xd:
                    return False, f"{len(xd)} file đã bị thay đổi ở nơi khác.", [], xd
                for p, d in doi.items():
                    f = self.local / self._p(p)
                    if d is None:
                        if f.exists():
                            f.unlink()
                    else:
                        f.parent.mkdir(parents=True, exist_ok=True)
                        tmp = f.with_suffix(f.suffix + ".tmp")
                        tmp.write_bytes(d)
                        tmp.replace(f)
                return True, f"Đã lưu {len(doi)} file (cục bộ).", list(doi), {}
            return self._luu_github(doi, thong_diep, kv)

    def _dam_bao_repo_co_commit(self) -> str:
        """Repo trống -> tạo commit đầu. Trả về chuỗi lỗi (rỗng nếu ổn) để báo rõ cho người dùng."""
        r = self._req("GET", f"/repos/{self.repo}/git/ref/heads/{self.branch}", headers=self._h())
        if r.status_code in (404, 409):
            w = self._req("PUT", f"/repos/{self.repo}/contents/{self._p('README.md')}", headers=self._h(),
                          json={"message": "Khởi tạo kho dữ liệu PIM tool", "branch": self.branch,
                                "content": base64.b64encode("Kho dữ liệu của PIM tool.\n".encode()).decode()})
            if w.status_code not in (200, 201, 422):
                return (f"Kho dữ liệu đang trống và không tự khởi tạo được ({w.status_code}). Kiểm tra token có quyền "
                        f"Contents: Read and write cho repo {self.repo}, hoặc tạo repo với 'Add a README file'. "
                        f"Chi tiết: {w.text[:150]}")
        return ""

    def _luu_github(self, doi: Dict[str, Optional[bytes]], thong_diep: str,
                    kv: Dict[str, str]) -> Tuple[bool, str, List[str], Dict[str, str]]:
        loi_kd = self._dam_bao_repo_co_commit()
        if loi_kd:
            return False, loi_kd, [], {}
        loi = ""
        blob_sha: Dict[str, str] = {}
        for lan in range(8):
            r = self._req("GET", f"/repos/{self.repo}/git/ref/heads/{self.branch}", headers=self._h())
            if r.status_code != 200:
                return False, f"Không đọc được nhánh {self.branch}: {r.status_code} {r.text[:200]}", [], {}
            head = r.json()["object"]["sha"]
            if kv:  # kiểm tra phiên bản NGAY TẠI commit đầu nhánh hiện tại
                hien = self.phien_ban(list(kv), ref=head)
                xd = {p: hien[p] for p in kv if hien[p] != kv[p]}
                if xd:
                    self._ds_tam.clear()  # danh sách phiên bản tạm đã cũ
                    return False, f"{len(xd)} file đã bị thay đổi ở nơi khác.", [], xd
            r = self._req("GET", f"/repos/{self.repo}/git/commits/{head}", headers=self._h())
            base_tree = r.json()["tree"]["sha"]
            items = []
            for p, d in doi.items():
                if d is None:
                    items.append({"path": self._p(p), "mode": "100644", "type": "blob", "sha": None})
                    continue
                if p not in blob_sha:
                    rb = self._req("POST", f"/repos/{self.repo}/git/blobs", headers=self._h(),
                                   json={"content": base64.b64encode(d).decode(), "encoding": "base64"})
                    if rb.status_code != 201:
                        return False, f"Tạo blob {p} lỗi {rb.status_code}: {rb.text[:200]}", [], {}
                    blob_sha[p] = rb.json()["sha"]
                    self._bo_nho.dat(blob_sha[p], d)
                items.append({"path": self._p(p), "mode": "100644", "type": "blob", "sha": blob_sha[p]})
            rt = self._req("POST", f"/repos/{self.repo}/git/trees", headers=self._h(),
                           json={"base_tree": base_tree, "tree": items})
            if rt.status_code != 201:
                return False, f"Tạo tree lỗi {rt.status_code}: {rt.text[:200]}", [], {}
            rc = self._req("POST", f"/repos/{self.repo}/git/commits", headers=self._h(),
                           json={"message": thong_diep, "tree": rt.json()["sha"], "parents": [head]})
            if rc.status_code != 201:
                return False, f"Tạo commit lỗi {rc.status_code}: {rc.text[:200]}", [], {}
            ru = self._req("PATCH", f"/repos/{self.repo}/git/refs/heads/{self.branch}", headers=self._h(),
                           json={"sha": rc.json()["sha"], "force": False})
            if ru.status_code == 200:
                self._ds_tam.clear()
                return True, f"Đã lưu {len(doi)} file lên GitHub (1 commit).", list(doi), {}
            loi = f"{ru.status_code}: {ru.text[:200]}"
            time.sleep(0.5 * (lan + 1) + random.random())  # có người khác vừa commit -> lấy đầu nhánh mới, kiểm tra lại, thử lại
        return False, f"Cập nhật nhánh thất bại sau nhiều lần thử ({loi}).", [], {}

    # ---------------------------------------------------------------- lịch sử / khôi phục (GitHub)
    def lich_su(self, path: str, so: int = 30) -> List[dict]:
        """Các commit đã sửa 1 file (mới nhất trước): [{sha, luc, thong_diep}]."""
        if self.backend == "local":
            return []
        r = self._req("GET", f"/repos/{self.repo}/commits", headers=self._h(),
                      params={"path": self._p(path), "sha": self.branch, "per_page": so})
        if r.status_code != 200:
            raise RuntimeError(f"Đọc lịch sử lỗi {r.status_code}: {r.text[:200]}")
        return [{"sha": c["sha"], "luc": c["commit"]["committer"]["date"], "thong_diep": c["commit"]["message"]}
                for c in r.json()]

    def doc_ban_cu(self, path: str, commit: str) -> Optional[bytes]:
        if self.backend == "local":
            return None
        r = self._req("GET", f"/repos/{self.repo}/contents/{self._p(path)}", headers=self._h(raw=True),
                      params={"ref": commit})
        if r.status_code == 404:
            return None
        if r.status_code != 200:
            raise RuntimeError(f"Đọc bản cũ lỗi {r.status_code}: {r.text[:200]}")
        return r.content

    def kiem_tra(self) -> List[Tuple[str, bool, str]]:
        """Tự kiểm tra kết nối/quyền (trang Kiểm tra hệ thống)."""
        out = []
        if self.backend == "local":
            try:
                self.local.mkdir(parents=True, exist_ok=True)
                f = self.local / ".thu_ghi"
                f.write_text("ok")
                f.unlink()
                out.append(("Ghi thư mục cục bộ", True, str(self.local.resolve())))
            except Exception as e:  # noqa: BLE001
                out.append(("Ghi thư mục cục bộ", False, str(e)))
            return out
        r = self._req("GET", f"/repos/{self.repo}", headers=self._h())
        if r.status_code != 200:
            out.append(("Truy cập repo", False, f"{r.status_code} — kiểm tra GITHUB_DATA_REPO và quyền token"))
            return out
        j = r.json()
        out.append(("Truy cập repo", True, f"{j.get('full_name')} ({'private' if j.get('private') else 'PUBLIC — nên để private'})"))
        quyen = (j.get("permissions") or {})
        out.append(("Quyền ghi (Contents: write)", bool(quyen.get("push", True)),
                    "có" if quyen.get("push", True) else "token chỉ có quyền đọc"))
        rl = self._req("GET", "/rate_limit", headers=self._h())
        if rl.status_code == 200:
            c = rl.json().get("resources", {}).get("core", {})
            out.append(("Hạn mức API còn lại", c.get("remaining", 0) > 200, f"{c.get('remaining')}/{c.get('limit')} lượt/giờ"))
        return out
