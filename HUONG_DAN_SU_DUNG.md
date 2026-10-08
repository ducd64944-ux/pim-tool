# Hướng dẫn sử dụng PIM Tool — 3 bước

## Bước ① Nạp dữ liệu

1. Vào **🚀 Chạy pipeline** → phần **① Nạp dữ liệu**
2. Kéo thả file vào hoặc bấm **Browse files**
3. Tool tự nhận loại file:
   - **File CMS export** → DATA SP
   - **File mẫu ngành** → Cấu hình ngành
   - **Danh sách SKU / file PIM** → IMPORT
   - **Workspace theo mẫu** → lấy nhiều loại cùng lúc
4. Không có file? Dán bảng từ Excel, hoặc tick **✍️ Tự điền tay model / SKU / mã biến thể** rồi **dán theo cột**: copy cột Model / SKU / Mã biến thể từ Excel dán vào 3 ô (mỗi dòng 1 giá trị, thứ tự dòng khớp nhau; Model chỉ 1 giá trị thì áp cho tất cả SKU). Chỉ cần SKU; có thể đặt *Mã ngành mặc định*. Muốn gõ tay từng dòng thì chọn **Bảng gõ tay** → bấm **Nạp vào tool**

> **Lưu ý:** Phải có cả IMPORT và DATA SP mới chuyển sang bước ② được.

---

## Bước ② Map

1. Sau khi nạp xong → bấm **🚀 Map dữ liệu**
2. Tool tự chuyển sang trang **🔍 Kiểm tra & Đối chiếu**
3. Xem bảng QC, sửa lỗi nếu có:
   - **Ô trống** → bấm ✨ Lấy PIM cũ
   - **Đơn vị** → bấm 📏 Đơn vị & biến đổi
   - **FILTER sai** → bấm vào dòng lỗi để chọn lại

> **Mẹo:** Sửa xong bấm Map lại để cập nhật kết quả.

---

## Bước ③ Xuất file import

1. Kiểm tra xong → xuống phần **③ Xuất file import** (hoặc vào trang **📤 Xuất**)
2. Tick chọn ngành hàng cần xuất
3. Bấm **📤 Tạo file import**
4. Tải file .xlsx về máy (từng file hoặc .zip)

> File xuất đúng chuẩn PIM: dạng Text, bỏ cột sku, đúng bố cục import.

---

## Thêm

- **📮 Đề xuất sửa CMS**: dữ liệu CMS sai → đề xuất sửa → admin duyệt → dùng chung cho mọi người.
- **⚙️ Cấu hình**: xem/sửa mapping, cấu hình ngành (admin).
- **💾 Lưu**: thay đổi tự lưu lên kho. Nút 💾 trên thanh bên để lưu thủ công.
- **Lỗi?** Thử tải lại trang. Nếu vẫn lỗi → liên hệ admin.
