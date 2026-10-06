# -*- coding: utf-8 -*-
"""
ai_helper.py — AI MIỄN PHÍ gợi ý lỗi sai / đề xuất (tuỳ chọn, cần API key free).

Dùng chuẩn "OpenAI-compatible chat completions" nên đổi nhà cung cấp chỉ cần đổi
cấu hình, không sửa code:
  * groq       — https://console.groq.com/keys  (free tier: openai/gpt-oss-120b, llama-3.3-70b-versatile …)
  * gemini     — https://aistudio.google.com/apikey (free tier Flash; Google dùng dữ liệu free tier để cải thiện sản phẩm)
  * openrouter — https://openrouter.ai/keys (các model có đuôi ":free")
Tên model thay đổi theo thời gian -> có hàm ds_model() để chọn model đang chạy được.

AI CHỈ ĐỀ XUẤT — người dùng tick mới thành sửa tay. Prompt cấm bịa thông số.
"""
from __future__ import annotations

import json
import re
import time
from typing import Dict, List, Tuple

import requests

PRESET: Dict[str, dict] = {
    "groq": {"ten": "Groq (miễn phí)", "base": "https://api.groq.com/openai/v1", "model": "openai/gpt-oss-120b",
             "lay_key": "https://console.groq.com/keys"},
    "gemini": {"ten": "Google Gemini (free tier)", "base": "https://generativelanguage.googleapis.com/v1beta/openai",
               "model": "gemini-flash-latest", "lay_key": "https://aistudio.google.com/apikey"},
    "openrouter": {"ten": "OpenRouter (model :free)", "base": "https://openrouter.ai/api/v1", "model": "",
                   "lay_key": "https://openrouter.ai/keys"},
}


class LoiAI(Exception):
    pass


class AI:
    def __init__(self, provider: str, key: str, model: str = "", base: str = ""):
        p = PRESET.get(provider, PRESET["groq"])
        self.provider = provider if provider in PRESET else "groq"
        self.key = (key or "").strip()
        self.model = (model or p["model"]).strip()
        self.base = (base or p["base"]).rstrip("/")

    @property
    def co_san(self) -> bool:
        return bool(self.key and self.model)

    @property
    def mo_ta(self) -> str:
        return f"{PRESET[self.provider]['ten']} · model {self.model or '(chưa chọn)'}"

    def _h(self) -> dict:
        return {"Authorization": f"Bearer {self.key}", "Content-Type": "application/json"}

    def ds_model(self) -> List[str]:
        r = requests.get(f"{self.base}/models", headers=self._h(), timeout=30)
        if r.status_code != 200:
            raise LoiAI(_loi_http(r))
        ids = [m.get("id", "") for m in r.json().get("data", []) if isinstance(m, dict)]
        ids = [i.split("/", 1)[1] if i.startswith("models/") else i for i in ids]
        if self.provider == "openrouter":
            ids = [i for i in ids if i.endswith(":free")]
        return sorted(i for i in ids if i)

    def chat(self, messages: List[dict], max_tokens: int = 2000, temperature: float = 0.2) -> str:
        if not self.co_san:
            raise LoiAI("Chưa cấu hình API key / model cho AI.")
        body = {"model": self.model, "messages": messages, "temperature": temperature, "max_tokens": max_tokens}
        for lan in range(3):
            try:
                r = requests.post(f"{self.base}/chat/completions", headers=self._h(), json=body, timeout=120)
            except requests.RequestException as e:
                if lan == 2:
                    raise LoiAI(f"Không kết nối được tới AI: {e}")
                time.sleep(2)
                continue
            if r.status_code == 429 and lan < 2:
                cho = min(float(r.headers.get("retry-after", "5") or 5), 20)
                time.sleep(cho)
                continue
            if r.status_code != 200:
                raise LoiAI(_loi_http(r))
            try:
                return r.json()["choices"][0]["message"]["content"] or ""
            except Exception:  # noqa: BLE001
                raise LoiAI(f"AI trả về dữ liệu lạ: {r.text[:300]}")
        raise LoiAI("AI đang quá tải — thử lại sau ít phút.")


def _loi_http(r: requests.Response) -> str:
    if r.status_code in (401, 403):
        return "API key sai hoặc không có quyền (401/403) — kiểm tra lại key."
    if r.status_code == 404:
        return "Không tìm thấy model/endpoint (404) — bấm 'Liệt kê model' để chọn model đang có."
    if r.status_code == 429:
        return "Hết lượt miễn phí tạm thời (429) — đợi ít phút hoặc đổi model/nhà cung cấp."
    return f"Lỗi AI {r.status_code}: {r.text[:300]}"


def doc_json(txt: str):
    """Lấy khối JSON đầu tiên trong câu trả lời (AI hay bọc ```json ... ```)."""
    txt = re.sub(r"```(?:json)?", "", txt or "")
    a, b = txt.find("{"), txt.rfind("}")
    if a < 0 or b <= a:
        return None
    try:
        return json.loads(txt[a:b + 1])
    except Exception:  # noqa: BLE001
        return None


QUY_TAC = """Quy tắc BẮT BUỘC:
- KHÔNG bịa thông số. Chỉ đề xuất khi có căn cứ ngay trong dữ liệu được đưa: mâu thuẫn giữa các ô, sai đơn vị
  hoặc sai độ lớn (vd khối lượng loa 91.5 kg), lỗi chính tả/định dạng, giá trị đặt nhầm cột, khác PIM cũ mà PIM cũ hợp lý hơn.
- Không chắc giá trị đúng -> để "de_xuat" rỗng, ghi lý do cần kiểm tra.
- KHÔNG sửa cột có chữ "filter" trong mã (đó là mã option số).
- Giữ định dạng web: nhiều giá trị nối bằng "|", số và đơn vị cách 1 khoảng trắng (6.95 cm), không thêm chữ thừa.
- Lý do viết tiếng Việt, ngắn gọn (dưới 25 từ)."""


def goi_y_sku(ai: AI, nganh: str, sku: str, ten_sp: str, dong: List[dict]) -> Tuple[List[dict], str]:
    """dong: [{ma, ten, tool_moi, pim_cu}] -> ([{ma, de_xuat, ly_do, muc_do}], câu trả lời thô)."""
    bang = "\n".join(f"{d['ma']} | {d['ten']} | {str(d['tool_moi'])[:160]} | {str(d['pim_cu'])[:160]}"
                     for d in dong[:90] if d.get("tool_moi") or d.get("pim_cu"))
    msgs = [{"role": "system", "content":
             "Bạn là chuyên viên kiểm duyệt thông số kỹ thuật (TSKT) sản phẩm điện máy của Thế Giới Di Động / "
             "Điện Máy Xanh. Rà bảng thông số 1 sản phẩm, chỉ ra ô SAI hoặc ĐÁNG NGỜ và đề xuất giá trị sửa.\n"
             + QUY_TAC + '\nChỉ trả về JSON: {"goi_y":[{"ma":"<mã cột>","de_xuat":"<giá trị hoặc rỗng>",'
             '"ly_do":"...","muc_do":"Cao|Trung bình|Thấp"}]}. Không có gì sai: {"goi_y":[]}.'},
            {"role": "user", "content": f"Ngành hàng: {nganh}\nSKU: {sku}\nTên sản phẩm: {ten_sp}\n"
                                        f"Bảng (mã | tên cột | giá trị TOOL sẽ import | giá trị PIM cũ đang trên web):\n{bang}"}]
    raw = ai.chat(msgs)
    js = doc_json(raw) or {}
    ma_hop_le = {d["ma"] for d in dong}
    out = []
    for g in js.get("goi_y", []) if isinstance(js, dict) else []:
        if isinstance(g, dict) and g.get("ma") in ma_hop_le and "filter" not in g["ma"].lower():
            out.append({"ma": g["ma"], "de_xuat": str(g.get("de_xuat") or "").strip(),
                        "ly_do": str(g.get("ly_do") or "").strip(), "muc_do": str(g.get("muc_do") or "Trung bình")})
    return out, raw


def xet_o_danh_dau(ai: AI, nganh: str, items: List[dict]) -> Tuple[Dict[int, dict], str]:
    """Nhờ AI xem lại các ô kiểm tra thông minh đã đánh dấu.
    items: [{id, cot, gia_tri, goi_y_tool, ly_do_tool, mau_cot}] -> {id: {dong_y, de_xuat, ly_do}}."""
    dong = "\n".join(json.dumps(i, ensure_ascii=False) for i in items[:30])
    msgs = [{"role": "system", "content":
             "Bạn là chuyên viên kiểm duyệt TSKT điện máy. Tool kiểm tra tự động đã đánh dấu các ô nghi sai kèm gợi ý. "
             "Với MỖI ô, cho biết có đồng ý là lỗi không, và giá trị nên sửa.\n" + QUY_TAC +
             '\nChỉ trả về JSON: {"ket_qua":[{"id":<số>,"dong_y":true|false,"de_xuat":"<giá trị hoặc rỗng>",'
             '"ly_do":"..."}]}'},
            {"role": "user", "content": f"Ngành hàng: {nganh}\nCác ô (JSON mỗi dòng; mau_cot = vài giá trị phổ biến "
                                        f"khác trong cùng cột để so sánh):\n{dong}"}]
    raw = ai.chat(msgs, max_tokens=2500)
    js = doc_json(raw) or {}
    out: Dict[int, dict] = {}
    for g in js.get("ket_qua", []) if isinstance(js, dict) else []:
        try:
            out[int(g.get("id"))] = {"dong_y": bool(g.get("dong_y")), "de_xuat": str(g.get("de_xuat") or "").strip(),
                                     "ly_do": str(g.get("ly_do") or "").strip()}
        except Exception:  # noqa: BLE001
            continue
    return out, raw


TINH_NANG_TOOL = """Tool hiện có: nạp file CMS export (DATA SP) + file export PIM (IMPORT + spec cũ); map TSKT (text nối |)
và FILTER (mã option tra DATA PIM, nối ", "); map dự phòng theo tên cột; kiểm tra: khác spec PIM, tool trống, chỉ thêm đơn vị,
thiếu model/category, FILTER sai mã; sửa tay / lấy PIM cũ; đơn vị hàng loạt; quy tắc Không/Đang cập nhật (giữ/để trống/thay);
giải nghĩa + chọn lại option FILTER; tách kích thước ghép; thuộc tính chưa map -> thêm vào MAPPING; kiểm tra thông minh
(giá trị bất thường, lẫn đơn vị, nhiều kiểu viết, lỗi gõ); xuất .xlsx MODEL/BIẾN THỂ; lưu GitHub; nhiều tài khoản;
đề xuất sửa dữ liệu CMS sai (thành viên dùng ngay, admin duyệt thành quy tắc dùng chung)."""


def hoi(ai: AI, cau_hoi: str, ngu_canh: str, lich_su: List[dict]) -> str:
    msgs = [{"role": "system", "content":
             "Bạn là trợ lý của PIM Tool (chuyển thông số sản phẩm từ CMS sang PIM cho Thế Giới Di Động / Điện Máy Xanh). "
             "Trả lời tiếng Việt, ngắn gọn, đi thẳng vào việc, dùng gạch đầu dòng khi liệt kê. Dựa trên số liệu lô hiện tại "
             "được cung cấp; không bịa số liệu. Khi được hỏi đề xuất tính năng: nêu cụ thể tính năng, lợi ích, mức ưu tiên.\n"
             + TINH_NANG_TOOL + "\n\nSỐ LIỆU LÔ HIỆN TẠI:\n" + ngu_canh[:12000]}]
    for m in lich_su[-6:]:
        msgs.append({"role": m["role"], "content": m["content"][:4000]})
    msgs.append({"role": "user", "content": cau_hoi})
    return ai.chat(msgs, max_tokens=1800, temperature=0.4)
