# Hướng dẫn sử dụng PIM Tool — 3 bước

## Bước ① Nạp dữ liệu

1. Vào **🚀 Chạy pipeline** → phần **① Nạp dữ liệu**
2. Kéo thả file vào hoặc bấm **Browse files**
3. Tool tự nhận loại file:
   - **File CMS export** → DATA SP
   - **File mẫu ngành** → Cấu hình ngành
   - **Danh sách SKU / file PIM** → IMPORT
   - **Workspace theo mẫu** → lấy nhiều loại cùng lúc
4. Không có file? Tick **✍️ Tự điền tay model / SKU / mã biến thể** rồi **dán theo cột**: copy cột Model / SKU / Mã biến thể / ID CMS từ Excel dán vào 4 ô (mỗi dòng 1 giá trị, khớp theo thứ tự dòng; ô ngắn hơn thì phần cuối để trống — vd SKU 4 dòng, ID 9 dòng thì 5 dòng cuối chỉ có ID; Model / Biến thể chỉ 1 giá trị thì áp cho tất cả dòng; tự bỏ dòng tiêu đề, báo dòng trùng / ID không phải số / dòng thiếu cả SKU lẫn ID). Mỗi dòng cần **SKU hoặc ID CMS** (SP chưa có SKU/code thì nhập ID CMS = PRODUCTID trong file CMS export; file CMS phải nạp cùng hoặc đã nạp trước); đặt *Mã ngành mặc định* ở ô phía trên (áp cho mọi dòng). Muốn gõ tay từng dòng thì chọn **Bảng gõ tay** → bấm **Nạp vào tool**

> **File import theo ID:** file có cột `PRODUCTID` (hoặc ID CMS / ID model / ID biến thể) thay cho cột SKU cũng nạp được. SP trong CMS export chưa có PRODUCTCODE vẫn được giữ, khớp theo PRODUCTID.

> **Lưu ý:** Phải có cả IMPORT và DATA SP mới chuyển sang bước ② được.

---

## Bước ② Map

1. Sau khi nạp xong → bấm **🚀 Map dữ liệu**
2. Tool tự chuyển sang trang **🔍 Kiểm tra & Đối chiếu**
3. Xem bảng QC, sửa lỗi nếu có:
   - **Ô trống** → bấm ✨ Lấy PIM cũ
   - **Đơn vị** → bấm 📏 Đơn vị & biến đổi — bảng ① có cột **THÊM TRƯỚC**, **ĐƠN VỊ (SAU)** và **PHẠM VI** (Số trơn / Từng giá trị / Có số-chữ+số) giống mục ②; chỉ điền SAU + Số trơn là rule cũ
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
- **🧹 Làm mới lô (xoá nhanh)** (ngay dưới thanh IMPORT / DATA SP / Spec cũ / Kết quả map): bấm → xem số liệu sẽ xoá → **Xoá ngay**. Xoá IMPORT, DATA SP, kết quả map và sửa tay của lô đang làm để nạp lô mới. **Giữ nguyên** Cấu hình, Mapping TSKT/FILTER, DATA PIM (dùng chung), đơn vị/biến đổi hàng loạt, quy tắc. Spec cũ chỉ xoá khi tick riêng. Không hoàn tác được.
- **🧹 Dọn RAM** (cuối trang): bấm khi trang chậm / RAM đầy. Giải phóng bộ nhớ, không mất dữ liệu đang làm. Khi RAM > 75% tool cũng tự dọn và hiện cảnh báo.
- **Lỗi?** Thử tải lại trang. Nếu vẫn lỗi → liên hệ admin.
