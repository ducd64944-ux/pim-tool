# Dùng hằng ngày — 3 bước

> **Giao diện 3 vùng ngang** (giống 66.py), chọn ở thanh trên cùng:
> 1. **🚀 Chạy pipeline**: ① Nạp (dữ liệu lô · data gốc · 👀 xem dữ liệu đã nạp) → ② 🚀 Map (xong tự chuyển sang vùng 2) → ③ Xuất file import.
> 2. **🔍 Kiểm tra & Đối chiếu**: tab **🛡️ QC tổng hợp** gom mọi lỗi; **bấm 1 dòng lỗi** để mở bảng xem & sửa riêng cho lỗi đó (sửa, đối chiếu bằng mắt như Excel). Ví dụ: dòng *Kích thước* mở bảng **SKU × Dài · Rộng · Cao · Sâu · Ngang · Khối lượng** để điền đơn vị và sửa số trực tiếp; dòng *Ô KHÁC spec PIM* mở bảng PIM cũ ↔ TOOL MỚI; dòng *FILTER không khớp* mở bảng chọn option đúng. Ngoài ra còn đối soát, gộp kích thước, Không/Đang cập nhật, gợi ý AI…
> 2b. Cuối vùng **🔍 Kiểm tra & Đối chiếu** có **④ Xuất file import**: chọn ngành ở CHỌN NGÀNH HÀNG → bấm 📤 Tạo file import (file MODEL / BIENTHE, bỏ cột sku, mọi ô Text — đúng chuẩn 66.py), tải thêm 📊 báo cáo kiểm tra và 📦 workspace theo mẫu. Vùng 🚀 Chạy pipeline cũng có ③ Xuất nhanh.
> 3. **📋 Quản lý dữ liệu**: ⚙️ Cấu hình & mapping · 📮 Đề xuất sửa CMS · 🧰 Tra cứu · 📘 Hướng dẫn · 👥 Quản trị (admin).
>
> - **👀 Xem dữ liệu đã nạp**: xem IMPORT, DATA SP, SPEC, cấu hình, mapping, DATA PIM, đơn vị/biến đổi đã đặt — có ô tìm, không cần map.
> - **🔄 Nạp lại TOÀN BỘ data gốc** (tab Data gốc, admin): thay toàn bộ cấu hình ngành + mapping TSKT + mapping FILTER + DATA PIM bằng file gốc (tuỳ chọn kèm dữ liệu lô).

## Cách nhanh nhất: trang 🏁 Làm nhanh (3 bước trên 1 trang)

### ① Nạp 1 cục — kéo thả file BẤT KỲ hoặc dán từ Excel

Tool **tự nhận loại file**:

- **Workspace theo mẫu** (nhiều sheet: IMPORT, DATA SP, CẤU HÌNH, MAPPING…): lấy dữ liệu lô. Muốn chọn kỹ từng vùng thì dùng **📦 Nạp file theo mẫu** (bên dưới): tick đúng vùng cần nạp (IMPORT · DATA SP · SPEC PIM tạm · Đơn vị · Chọn ngành · Cấu hình · Mapping TSKT · Mapping FILTER · DATA PIM), mỗi vùng lô chọn **Ghi đè** hoặc **Nối tiếp**. Vùng không tick **giữ nguyên hoàn toàn**. Có nút ✅ Tất cả / Chỉ vùng LÔ / Chỉ vùng DÙNG CHUNG / Bỏ chọn hết. Vùng dùng chung chỉ admin nạp được.
- **File CMS export** → DATA SP:
  - tiêu đề tiếng Anh/Việt, có/không dấu đều nhận (Mã ERP, Mã thuộc tính, Giá trị…);
  - không có tiêu đề thì đọc theo vị trí cột A–G như 66.py.
- **File mẫu ngành hàng** (`Export Product Template`: dòng 1 = tên ngành, mã ngành, các mã cột thuộc tính; dòng 2 = tên cột; không có dữ liệu) → thêm/cập nhật **Cấu hình ngành hàng** (danh sách cột TSKT/FILTER). Tool báo ngành mới hay thêm/bỏ bao nhiêu cột so với cấu hình cũ. Chỉ admin cập nhật. Sau đó nạp mapping của ngành để map được dữ liệu.
- **Danh sách SKU / file export PIM** (model, SKU, biến thể, category, kèm hoặc không kèm cột TSKT/FILTER):
  - tên cột tuỳ ý (Mã model, Mã sản phẩm ERP, Mã biến thể…);
  - **không có tiêu đề thì tool đoán theo nội dung** (cột số dài nhất là SKU…), có bảng xem trước để kiểm tra lại;
  - **Mã họ biến thể có chữ `color`** (vd `lvl_1_color_iden_master`) = biến thể theo màu → SKU đó xuất vào file **MODEL**, không phải BIENTHE;
  - cột TSKT/FILTER được **tách tự động** thành spec PIM cũ để đối chiếu.

Sau khi nạp:

- **Tự lọc:** DATA SP chỉ giữ dòng của SKU có trong IMPORT (bỏ tick nếu muốn giữ hết).
- **Map + QC chạy luôn.** Ô chọn file và ô dán tự trống lại sau khi nạp, nên không bị nạp lặp.

### ② Map + 🛡️ QC tổng hợp

Một bảng gom mọi vùng kiểm tra, mỗi dòng có mức độ, ý nghĩa và gợi ý xử lý: thiếu model/category, ERP/model không có giá trị, mapping trỏ tới cột thiếu, FILTER không khớp, khác spec PIM, đơn vị, Không/Đang cập nhật, nghi sai, quy tắc, model lặp…

Các nút **⚡ sửa nhanh**:

- Lấy PIM cũ cho ô trống
- Điền đơn vị gợi ý
- Không/Đang cập nhật → để trống
- Áp gợi ý sửa mức Cao
- Thêm cột thiếu vào cấu hình (admin)
- Tải file xin data

### ③ Xuất

- Tạo file import: đúng chuẩn desktop, chỉ xuất các ngành đã tick ở bảng CHỌN NGÀNH HÀNG.
- **📨 File xin data CMS**: danh sách ERP (SKU) và MODEL không có giá trị, kèm lý do:
  - CMS chưa có dữ liệu,
  - ngành chưa có mapping/cấu hình,
  - hoặc có dữ liệu nhưng không map được thuộc tính nào.

  Mặc định các SKU này được **tách khỏi file import**. Bỏ tick nếu muốn xuất dòng trống như desktop.

## ② Map & kiểm tra (trang 🚀)

Bấm **① Map dữ liệu**, rồi xem các ô số. Thứ tự nên làm:

| Tab | Làm gì |
|---|---|
| ⚠️ Cảnh báo | Thiếu model/category, FILTER sai mã, … |
| ≠ Khác spec PIM | Sửa cột **TOOL MỚI** hoặc tick **Lấy PIM cũ** → Áp dụng |
| 🪄 (nút) | Ô tool trống mà PIM đang có → lấy PIM cũ |
| 📏 Đơn vị & biến đổi hàng loạt | ① Đơn vị theo cột (gõ đơn vị **tuỳ ý**; rule cũ: chỉ ô **số trơn**, không đụng FILTER) · ② Biến đổi hàng loạt cho cột chọn: thêm chữ/đơn vị phía **sau**, phía **trước**, hoặc **cả hai** (vd `Khoảng 12 cm`), **đổi đơn vị** (mm→cm, g→kg, inch→cm… tự quy đổi số), thay chữ, làm tròn. Phạm vi: chỉ ô số trơn / từng số trong ô `9\|10` / mọi ô. Có **xem trước** trước khi áp, bỏ được bất cứ lúc nào |
| 📐 Gộp / tách kích thước | **Gộp** 2–4 cột (Dài/Rộng/Cao, Ngang/Cao/Sâu…) thành 1 giá trị trong cột kích thước/kích cỡ của PIM theo mẫu: `30 x 20 x 10 cm`, `Dài 30 cm - Rộng 20 cm - Cao 10 cm`, hoặc tự nhập mẫu `{1} {2} {3} {dv}`. **Tách** cột ghép `Ngang … - Cao …` ra các cột con. Có xem trước, ghi dạng sửa tay |
| 🚫 Không / Đang cập nhật | Chọn **Giữ** / **Để trống — không cập nhật** / **Thay bằng…** |
| 🔎 Theo SKU + FILTER | Giải nghĩa mã FILTER (vd `6 = Tivi thông minh`), chọn lại option |
| 📈 Độ hoàn thiện & quy tắc | % đầy đủ từng SKU/ngành (hạng A–E), SKU thiếu cột bắt buộc, ô vi phạm quy tắc |
| 🧾 Đối soát CMS → kết quả | Giá trị CMS **không vào được** file import và lý do; quy đổi FILTER |
| 🤖 Gợi ý thông minh & AI | Giá trị bất thường, lẫn đơn vị, lỗi gõ; AI (nếu bật) |

Mọi chỉnh sửa tự lưu và tự áp lại sau mỗi lần map.

**Dữ liệu CMS sai 1–2 giá trị?** Dùng trang **📮 Đề xuất sửa CMS**:

- Bạn dùng được giá trị đã sửa ngay.
- Admin duyệt xong thì cả nhóm dùng lâu dài.

## ③ Xuất file (trang 📤)

1. Đọc khung cảnh báo.
2. Tick ngành ở bảng **CHỌN NGÀNH HÀNG** (giống sheet trong file mẫu).
3. Bấm **📤 Tạo file import**.
4. Tải về: từng file `PIM_TSKT <mã> <tên>_MODEL_…xlsx` / `_BIENTHE_…xlsx`, hoặc tải 1 file .zip.
   - Mọi ô ở dạng Text.
   - Bỏ cột sku.
   - Đúng bố cục bản desktop.
5. **📦 Tải workspace theo mẫu**: 1 file Excel đúng bố cục mẫu. Mở bằng Excel hoặc bản desktop 66.py để làm tiếp hoặc lưu trữ.

---

## Admin

- **⚙️ Cấu hình & mapping**:
  - Nạp file mẫu (gộp theo ngành).
  - Sửa cấu hình, mapping, DATA PIM.
  - **✅ Quy tắc kiểm tra**: Bắt buộc · Là số + đơn vị · Khoảng giá trị · Danh sách giá trị · Độ dài · Regex. Có nút **🪄 Gợi ý từ dữ liệu lô**.
  - **🕘 Lịch sử & khôi phục**: xem ai sửa gì, khôi phục bản cũ.
- **📮 Đề xuất sửa CMS → 📥 Duyệt**: duyệt hoặc từ chối đề xuất của thành viên.
- **👥 Quản trị**:
  - 🩺 Kiểm tra hệ thống.
  - Thống kê tài khoản.
  - AI.
  - Tạo mật khẩu băm.
- **Xem workspace của người khác**: chọn ở ô "Workspace đang xem" trên thanh bên.
