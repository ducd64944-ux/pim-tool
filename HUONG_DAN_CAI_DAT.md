# Hướng dẫn cài đặt PIM Tool (bản web) — làm 1 lần, khoảng 20 phút

Bạn cần: 1 tài khoản **GitHub** và 1 tài khoản **Streamlit Community Cloud** (đăng nhập bằng chính GitHub). Mọi bước làm trên trình duyệt, không cần cài gì vào máy.

Tóm tắt: tạo **2 repo private** → tạo **1 token** → đưa code lên repo code → tạo app trên Streamlit → dán **Secrets** → đăng nhập admin → **🩺 Kiểm tra hệ thống** → nạp **file mẫu**.

---

## Bước 1 — Tạo 2 repo PRIVATE trên GitHub

1. Vào github.com, góc trên phải bấm **+** rồi chọn **New repository**.
2. Tạo repo **code**:
   - Repository name: `pim-tool`.
   - Chọn **Private**.
   - Bấm **Create repository**.
3. Làm lại để tạo repo **dữ liệu**:
   - Tên: `pim-tool-data`.
   - Chọn **Private**.
   - Tick **Add a README file** để repo không trống.
   - Bấm **Create repository**.

> Vì sao phải 2 repo? Mỗi lần tool lưu dữ liệu là 1 commit. Nếu commit vào repo code, Streamlit sẽ khởi động lại app, và người đang dùng bị văng ra. Tách repo dữ liệu riêng thì không bị như vậy.

## Bước 2 — Đưa code lên repo `pim-tool`

1. Giải nén `pim-tool.zip` ra máy.
2. Trên GitHub, mở repo `pim-tool` → bấm **uploading an existing file** (hoặc **Add file** → **Upload files**).
3. Kéo **toàn bộ nội dung** bên trong thư mục vừa giải nén vào trang (các file `app.py`, `pim_core.py`… phải nằm ở gốc repo, không nằm trong một thư mục con).
4. Thư mục `.streamlit` và `.devcontainer` là thư mục ẩn: trên Windows bật **View → Hidden items** mới thấy, trên macOS nhấn `Cmd + Shift + .` khi chọn file.
5. Kéo xong bấm **Commit changes**.

Kiểm tra: ở gốc repo phải thấy `app.py`, `requirements.txt` và thư mục `.streamlit/` (có `config.toml`).

**KHÔNG** đưa file `secrets.toml` thật lên GitHub. Chỉ có file mẫu `secrets.toml.example` là được đưa lên.

## Bước 3 — Tạo token cho repo dữ liệu

1. Trên GitHub, bấm ảnh đại diện → **Settings** → **Developer settings** (cuối menu trái) → **Personal access tokens** → **Fine-grained tokens** → **Generate new token**.
2. Điền như sau:
   - **Token name:** `pim-tool-data`.
   - **Expiration:** chọn thời hạn dài nhất bạn chấp nhận (ví dụ 1 năm). Ghi lại ngày hết hạn để thay token trước ngày đó.
   - **Repository access:** chọn **Only select repositories** → chọn `pim-tool-data`.
   - **Permissions → Repository permissions → Contents:** chọn **Read and write**. Mục Metadata tự bật Read-only, giữ nguyên.
3. Bấm **Generate token**, rồi **copy ngay** chuỗi `github_pat_...` (chỉ hiện 1 lần).

## Bước 4 — Tạo app trên Streamlit Community Cloud

1. Vào share.streamlit.io và đăng nhập bằng GitHub. Nếu được hỏi, cho phép truy cập **repo private**.
2. Bấm **Create app** → **Deploy a public app from GitHub** (app vẫn có mật khẩu đăng nhập riêng). Điền:
   - **Repository:** `<tài-khoản-github>/pim-tool`.
   - **Branch:** `main`.
   - **Main file path:** `app.py`.
   - **App URL:** đặt tên dễ nhớ.
3. Bấm **Advanced settings**:
   - **Python version:** chọn 3.11 hoặc 3.12.
   - **Secrets:** dán nội dung bên dưới, rồi sửa lại cho đúng.

```toml
GITHUB_TOKEN = "github_pat_...dán token ở bước 3..."
GITHUB_DATA_REPO = "<tài-khoản-github>/pim-tool-data"
GITHUB_BRANCH = "main"

[users.ducd]
password = "mat-khau-tam"
ten = "Đức"
admin = true

[users.an]
password = "mat-khau-tam-2"
ten = "Anh An"
admin = false
```

4. Bấm **Save**, rồi **Deploy**. Lần đầu app cài thư viện mất khoảng 2–4 phút.

> Muốn thêm hoặc bớt người dùng: vào app → **⋮** → **Settings** → **Secrets**, sửa phần `[users.xxx]` rồi **Save**. App tự khởi động lại, dữ liệu không mất.

## Bước 5 — Đăng nhập và tự kiểm tra

1. Mở link app, đăng nhập bằng tài khoản `admin = true`.
2. Vào trang **👥 Quản trị** → bấm **▶ Chạy kiểm tra**. Tất cả các dòng cần ✅:
   - **Truy cập repo / Quyền ghi:** nếu ❌, xem lại `GITHUB_DATA_REPO` và quyền **Contents: Read and write** của token.
   - **Ghi thử 1 file:** nếu ✅, mở repo `pim-tool-data` sẽ thấy commit "Kiểm tra ghi".
   - **Mật khẩu đã băm sha256:** nên đổi sang dạng băm. Cuối trang Quản trị có ô tạo chuỗi `sha256:...`; dán chuỗi đó vào Secrets thay cho mật khẩu chữ.

## Bước 6 — Nạp dữ liệu dùng chung lần đầu (admin)

1. Vào trang **📥 Nạp dữ liệu lô**, khung **📦 Nạp file theo mẫu**.
2. Chọn file workspace (`du_lieu_pim.xlsx` hoặc `TEST HÀNG LOẠT IMPORT THÔNG SỐ NEW.xlsx`).
3. Tool hiện số lượng đọc được từ từng sheet: IMPORT, DATA SP, DATA PIM, cấu hình, mapping, tab TSKT.
4. Tick cả 2 ô: **Dữ liệu LÔ** và **Mapping / cấu hình / DATA PIM → dùng chung**.
5. Bấm **✔ Nạp vào tool**. Tool lưu lên GitHub rồi tự map luôn.

Từ giờ mọi tài khoản dùng chung mapping, cấu hình và DATA PIM này. Mỗi người có workspace (lô hàng) riêng.

---

## Thành viên tự tạo tài khoản

- Ở trang đăng nhập có tab **🆕 Tạo tài khoản**. Thành viên tự điền tên đăng nhập (chữ thường, không dấu), tên hiển thị, mật khẩu.
- **Không có mã mời** → tài khoản ở trạng thái *chờ duyệt*. Admin vào **👥 Quản trị → Tài khoản thành viên tự đăng ký** bấm **✅ Duyệt**.
- **Có mã mời** → dùng được ngay. Đặt mã trong Secrets: `MA_MOI = "ma-cua-ban"`, rồi gửi mã cho thành viên.
- Admin có nút **🔑 Đặt MK** (đặt lại thành 123456) và **🗑 Xoá**.
- Tài khoản tự đăng ký lưu ở `shared/tai_khoan.json` trong repo dữ liệu, mật khẩu băm PBKDF2 (không lưu chữ thật). Tài khoản **admin** vẫn khai trong Secrets (`[users.xxx]`).
- Đăng nhập bằng **tên đăng nhập** (vd `ducd`), không phân biệt hoa/thường; nếu gõ nhầm tên hiển thị (vd `Duccontent`) mà không trùng ai thì vẫn nhận.

## Dùng đồng thời nhiều máy, nhiều người — tool tự lo

- **Mỗi tài khoản 1 workspace riêng**, nên người này không đụng dữ liệu lô của người kia.
- **Lúc lưu, tool kiểm tra phiên bản.** Nếu máy khác vừa lưu cùng file:
  - Dữ liệu **gộp được** thì tool **tự gộp**: mapping, cấu hình, quy đổi FILTER, quy tắc, đề xuất, sửa tay, đơn vị. Ví dụ hai người cùng thêm quy đổi thì giữ đủ cả hai.
  - Cùng **1 tài khoản mở ở 2 máy/tab** và cùng map lại lô: tool **không ghi đè** mà hiện khung đỏ cho chọn **Lấy bản trên kho** hoặc **Ghi đè bằng bản ở đây**.
- **Không cần đóng app** khi người khác đang dùng.
- Nên dùng **mỗi người 1 tài khoản**, không dùng chung tài khoản.

## Giới hạn và dữ liệu lớn

- **File tải lên:** tối đa 500 MB/file (đặt trong `.streamlit/config.toml`).
- **Đọc Excel:** dùng `python-calamine`, nhanh gấp 20–50 lần cách cũ.
  - File mẫu 10 MB (có vùng định dạng thừa 200.000 dòng) đọc xong trong khoảng 2 giây.
  - DATA SP 1 triệu dòng đọc trong khoảng 10–20 giây.
- **Thời gian xử lý đã đo** (15.000 SKU, 2,4 triệu dòng DATA SP): map khoảng 13 giây, kiểm tra khoảng 26 giây, xuất 20 file khoảng 23 giây. DATA SP chỉ chiếm khoảng 36 MB RAM nhờ nén cột.
- **RAM:** Streamlit Cloud miễn phí cho khoảng 2,7 GB RAM, dùng chung cho mọi người đang mở app. Lô lớn hơn 20.000 SKU nên chia nhỏ, hoặc chạy trên máy công ty (xem phần dưới).
- **GitHub:** mỗi file lưu tối đa 100 MB. DATA SP 2,4 triệu dòng chỉ khoảng 6 MB sau khi nén parquet, nên còn rất dư.

## Chạy trên máy công ty / máy cá nhân (không cần Streamlit Cloud)

```bash
pip install -r requirements.txt
streamlit run app.py
```

- Muốn lưu lên GitHub như bản Cloud: tạo file `.streamlit/secrets.toml`, nội dung giống bước 4.
- Không có token thì dữ liệu lưu ở thư mục `_du_lieu_cuc_bo/`, phù hợp để chạy thử.
- Máy công ty nhiều RAM thì chạy lô lớn thoải mái. Người khác trong mạng LAN mở `http://<ip-máy>:8501`.

## Lỗi thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| "Chưa cấu hình tài khoản" | Secrets thiếu phần `[users.xxx]` |
| Lưu báo 401/403 | Token sai hoặc hết hạn, hoặc thiếu quyền Contents: Read and write |
| Lưu báo 404 | `GITHUB_DATA_REPO` sai tên (phải là `tài-khoản/tên-repo`) |
| App chậm hoặc văng khi lô rất lớn | Chia lô nhỏ hơn hoặc chạy trên máy công ty |
| Khung đỏ "Workspace vừa được lưu từ máy khác" | Chọn **Lấy bản trên kho** (an toàn) hoặc **Ghi đè** nếu chắc bản ở máy này đúng |
| Token sắp hết hạn | Tạo token mới (bước 3), dán vào Secrets, Save |

## Công suất & chạy nhiều người (web-4.6)

Số đo thực (1 phiên, DATA SP 2 triệu dòng, 80.000 SKU): map ≈ 10 giây, RAM đỉnh ≈ 1,4 GB, sau đó giữ ≈ 0,75 GB.
Ước tính RAM cần = (số người đang mở lô lớn × ~0,8 GB) + 1 lần map đỉnh (~1,5 GB).

| Số người cùng lúc (mỗi người ~2 triệu dòng) | RAM máy chủ nên có |
|---|---|
| 1–2 | 4 GB |
| 5–6 | 8 GB (khuyến nghị 12–16 GB) |

Streamlit Community Cloud (miễn phí) chỉ có ~1–2,7 GB nên KHÔNG đủ cho lô 2 triệu dòng × nhiều người. Cần máy chủ riêng
(VPS / Cloud Run / Hugging Face Spaces trả phí…) chạy `streamlit run app.py` với cùng Secrets.

Bảo vệ đã có trong app:
- Việc nặng (Map, Tạo file import) chạy **xếp hàng**: mặc định 1 việc/lần (2 nếu máy ≥ 6 GB RAM). Đổi bằng Secret `PIM_MAX_JOBS`.
- Trước khi chạy việc nặng, app kiểm tra RAM còn trống; thiếu thì báo rõ, KHÔNG làm sập máy chủ.
- File xlsx báo cáo chỉ dựng khi bấm tải; tab Đơn vị & biến đổi chạy riêng (fragment).
- Admin thấy RAM + số việc nặng đang chạy ở chân trang.
