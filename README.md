# 🧩 PIM Tool — TGDĐ CMS → PIM (bản web Streamlit)

Bản web của tool desktop `66.py`: map TSKT + FILTER từ file CMS export sang định dạng import PIM, đối chiếu với spec PIM đang có, sửa trực tiếp, xuất file import. Mỗi tài khoản có 1 workspace riêng; dữ liệu lưu lên GitHub nên mở lại app vẫn còn nguyên.

## Cấu trúc

| File | Vai trò |
|---|---|
| `app.py` | Giao diện Streamlit (đăng nhập, các trang) |
| `pim_core.py` | Engine map / kiểm tra / xuất file (không phụ thuộc Streamlit) |
| `gh_store.py` | Lưu trữ GitHub: 1 lần lưu = 1 commit, khoá lạc quan theo phiên bản (nhiều máy không ghi đè nhau) |
| `dong_bo.py` | Gộp 3 chiều khi 2 máy cùng sửa 1 file |
| `HUONG_DAN_*.md`, `QUY_TAC_MAP.md` | Hướng dẫn cài đặt / sử dụng / quy tắc map (hiện trong trang 📘) |
| `ai_helper.py` | AI miễn phí (Groq / Gemini / OpenRouter, chuẩn OpenAI-compatible) — tuỳ chọn |
| `data/reference/` | Bảng tham chiếu: cột từng ngành hàng, tên thuộc tính → mã master, nhãn kích thước ghép |

## Cài đặt

Xem **[HUONG_DAN_CAI_DAT.md](HUONG_DAN_CAI_DAT.md)** — từng bước trên trình duyệt (2 repo private, token, Streamlit Cloud, Secrets, 🩺 kiểm tra hệ thống). Trong app có trang **📘 Hướng dẫn** hiển thị lại các file này.

Chạy trên máy: `pip install -r requirements.txt` rồi `streamlit run app.py`.

## Kiểm thử

```bash
python tests/test_core.py              # engine
python tests/test_parity_desktop.py    # web ra file import GIỐNG HỆT desktop 66.py (so từng ô)
python tests/test_dong_bo.py           # nhiều máy: phát hiện xung đột + gộp 3 chiều
python tests/test_github_gia_lap.py    # GitHub giả lập: 20 luồng ghi đồng thời không mất dữ liệu
python tests/so_sanh_desktop_file_that.py "file mẫu.xlsx"   # so desktop vs web trên file thật
```

## Dùng hằng ngày

1. **📥 Nạp dữ liệu lô**: file CMS export → DATA SP; file export PIM → IMPORT (model/sku/variant/category) + spec cũ.
2. **🚀 Map & kiểm tra**: bấm **① Map dữ liệu** (vài giây), rồi xử lý theo các ô số:
   - ≠ Khác spec PIM: sửa cột TOOL MỚI hoặc tick Lấy PIM cũ → Áp dụng.
   - 🪄 Ô tool trống → lấy PIM cũ (1 nút).
   - 📏 Đơn vị hàng loạt: Điền gợi ý → Áp.
   - 🚫 Không / Đang cập nhật: Giữ / Để trống / Thay bằng….
   - 🔎 Theo SKU + FILTER: giải nghĩa mã FILTER (vd `6 = Tivi thông minh`), chọn lại option.
   - 📐 Tách kích thước ghép: "Ngang … - Cao … - Sâu …" → các cột con.
   - 🧩 Thuộc tính chưa map: thuộc tính CMS chưa có trong MAPPING → tick để thêm.
   - 🧾 Đối soát CMS → kết quả: giá trị nào CÓ trong CMS mà KHÔNG vào được file import, và vì sao:
     - FILTER không có option, kèm gợi ý option gần giống. Lưu vào **bảng quy đổi** dùng chung thì mọi lần map sau tự khớp.
     - Mapping trỏ tới cột không có trong cấu hình (có nút thêm cột).
     - Mapping chỉ có ở ngành khác (có nút copy).
     - Thuộc tính chưa map, có gợi ý cột từ bảng tham chiếu.
     - SKU không có / thiếu nhiều thông số, cột luôn trống cả lô, lệch FILTER↔TSKT, nhiều thuộc tính gộp 1 cột.
     - **Độ phủ từng cột**: % SKU có giá trị, so với PIM cũ.
   - 🤖 Gợi ý thông minh & AI:
     - **Kiểm tra thông minh** (miễn phí, không cần AI): giá trị bất thường so với cả cột (vd loa 91.5 kg → gợi ý 9.15), lẫn đơn vị (472 mm trong cột cm → 47.2 cm), cùng chữ nhiều kiểu viết (kể cả "Không" nhìn giống hệt nhưng khác mã chữ), lỗi gõ (18. 4 cm, 56.3 Kg, dấu thừa, "A|A").
     - **AI rà từng SKU** / **Nhờ AI xem lại** các ô nghi sai. AI chỉ đề xuất, tick mới áp dụng. AI không được sửa cột FILTER và không được bịa thông số.
     - **Hỏi AI**: tóm tắt lỗi lô và thứ tự xử lý, đề xuất tính năng mới cho tool, hỏi đáp tự do.
3. **📤 Xuất file import**: đọc cảnh báo → xác nhận → tải từng file .xlsx (hoặc .zip).

## 📮 Đề xuất sửa dữ liệu CMS sai (thành viên đề xuất → admin duyệt)

Khi dữ liệu CMS (TSKT hoặc FILTER) sai 1–2 giá trị:

1. **Thành viên** vào trang **📮 Đề xuất sửa CMS → ➕ Tạo đề xuất**:
   - ① tick các ô đã sửa tay (ở tab ≠ Khác spec / 🔎 Theo SKU) để gửi thành đề xuất, hoặc
   - ② nhập trực tiếp: chọn SKU + cột → thấy giá trị CMS / kết quả đang dùng / PIM cũ → nhập giá trị đúng (FILTER: chọn option).
   - Phạm vi: **Chỉ SKU này** hoặc **Mọi SKU cùng giá trị** (trong ngành). FILTER không có option: gửi từ 🧾 Đối soát → 🔁 Quy đổi FILTER.
   - Đề xuất **áp dụng ngay** cho workspace của người đề xuất (tool tự map lại).
2. **Admin** thấy thông báo "N đề xuất chờ duyệt" ở thanh bên / Tổng quan → tab **📥 Duyệt**: xem giá trị CMS, giá trị đề xuất, PIM cũ, lý do → chỉnh giá trị nếu cần → **Duyệt / Từ chối** (hoặc Duyệt tất cả).
   - Duyệt → thành **quy tắc dùng chung**: mọi tài khoản, mọi lô sau map ra là đã sửa.
   - Từ chối → thôi áp cho người đề xuất (họ thấy trạng thái + ghi chú admin, map lại để cập nhật).
   - Tab **📚 Quy tắc dùng chung**: xem, tải Excel, **thu hồi** quy tắc.
3. Quy tắc "Chỉ SKU này" chỉ áp khi CMS **vẫn còn** giá trị sai lúc đề xuất — CMS sửa rồi thì tự bỏ qua (có ghi log).
4. Admin sửa/đề xuất thì được duyệt luôn.

Mọi chỉnh sửa (sửa tay, đơn vị, quy tắc Không/Đang cập nhật) tự lưu và tự áp lại sau mỗi lần map.

## Lưu trữ trong repo dữ liệu

```
data/shared/cau_hinh.json, quy_doi_filter.json, map_tskt.parquet, map_filter.parquet, data_pim.parquet   ← dùng chung (admin sửa)
data/shared/sua_sku.json, sua_gia_tri.json, de_xuat_duyet.json   ← quy tắc sửa CMS đã duyệt + quyết định của admin
data/shared/de_xuat/<tài khoản>.json                            ← đề xuất của từng tài khoản (mỗi người chỉ ghi file của mình → không đè nhau)
data/shared/quy_tac_kiem_tra.json                               ← quy tắc chất lượng dữ liệu (bắt buộc, số+đơn vị, danh sách…)
data/users/<tài khoản>/import.parquet, data_sp.parquet, spec.parquet,
                       ket_qua.parquet, ket_qua_meta.json, settings.json, lich_su.json
```

Parquet nén rất nhỏ: DATA SP 222.943 dòng chỉ khoảng 450 KB.

## Bật AI miễn phí (tuỳ chọn)

1. Lấy key free: Groq (khuyên dùng, free tier có model `openai/gpt-oss-120b`, `llama-3.3-70b-versatile`) tại console.groq.com/keys; hoặc Gemini tại aistudio.google.com/apikey.
2. Thêm vào Secrets: `AI_PROVIDER = "groq"`, `AI_API_KEY = "..."` (xem `secrets.toml.example`).
3. Trang 👥 Quản trị → **Kiểm tra kết nối AI**. Nếu báo 404 model: bấm **Liệt kê model** rồi điền `AI_MODEL`.

Free tier có giới hạn lượt/ngày; hết lượt thì app báo 429, đợi ít phút hoặc đổi model. Dữ liệu gửi đi chỉ là thông số sản phẩm. Lưu ý gói free của Gemini cho phép Google dùng dữ liệu để cải thiện sản phẩm.

## Lưu ý

- "Để trống — không cập nhật" chỉ an toàn nếu PIM bỏ qua ô trống khi import. Hãy thử 1 SKU trước.
- Map dự phòng theo tên mặc định chỉ khớp đúng tên cột của chính ngành hàng. Chế độ "bảng tham chiếu" rộng hơn nhưng cần soát lại: bảng tham chiếu có dòng sai, vd "Tổng công suất" → `total_capacity_tskt_master` ở ngành Loa.
- Cùng 1 tài khoản mở 2 nơi cùng lúc: lần lưu sau ghi đè file trước của workspace đó.
