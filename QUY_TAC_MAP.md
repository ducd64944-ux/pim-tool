# Quy tắc map — giữ đúng như bản desktop 66.py

File `tests/test_parity_desktop.py` tự kiểm tra các quy tắc dưới đây. Nó chạy **cùng một bộ dữ liệu** qua engine desktop và engine web, rồi so **từng ô** của file import.

`tests/so_sanh_desktop_file_that.py <file.xlsx>` làm việc tương tự trên file thật. Với file mẫu *TEST HÀNG LOẠT IMPORT THÔNG SỐ NEW* (46 SKU, 1.481 ô có giá trị), hai bản cho ra **0 ô khác nhau**.

| # | Quy tắc | Chi tiết |
|---|---|---|
| 1 | Nhận ngành của SKU | Lấy theo `CATEGORYID` trong DATA SP. Không có CATEGORYID thì **tự nhận diện**: ngành có nhiều PROPERTYID khớp mapping nhất (bằng điểm thì lấy mã ngành nhỏ nhất). CATEGORYID mâu thuẫn thì dùng giá trị đầu. |
| 2 | Ngành không có trong mapping | **Bỏ qua SKU**, ghi log "CATEGORYID không tồn tại trong mapping". |
| 3 | Danh sách cột | Lấy từ **tab "TSKT <mã> <tên>"** nếu file có tab. Cột cấu hình còn thiếu thì thêm vào cuối. Không có tab thì lấy từ **CẤU HÌNH CATEGORY**. Không có cả hai thì **bỏ qua ngành**. |
| 4 | Cột không được điền | `model_code, sku, category_code, variant_code` luôn bỏ qua khi điền. Cột chỉ có trên tab mà không có trong cấu hình thì vẫn xuất nhưng để trống. Cột như `family_code` có trong cấu hình thì vẫn xuất (trống nếu không có mapping). |
| 5 | TSKT | Văn bản nguyên văn PROPVALUE. Nhiều giá trị thì nối bằng `\|`, bỏ trùng, giữ thứ tự xuất hiện trong DATA SP. |
| 6 | FILTER | Tra **mã option** trong DATA PIM theo (mã cột, giá trị viết thường, bỏ khoảng trắng đầu/cuối). Nhiều mã thì nối bằng `, `. Không tra được thì **để trống + ghi log** (không bịa). Web thêm chuẩn hoá Unicode NFC: chỉ giúp khớp thêm các chữ trông giống hệt nhau, không đổi những mã đã khớp. |
| 7 | Map theo tên | **Tắt mặc định** (đúng desktop). Chỉ khi bật tay ở trang 🚀 mới dùng. |
| 8 | Đơn vị | Chỉ thêm vào ô **số trơn** (`9` → `9 kg`). Ô `9\|10`, ô có chữ và cột FILTER thì không đụng tới. |
| 9 | Thứ tự áp khi xuất | Sửa tay → quy tắc Không/Đang cập nhật → đơn vị. |
| 10 | So spec PIM | Làm sạch `["905", " 903"]` → `905, 903`. So khớp SKU, không có thì so theo (model, biến thể). |
| 11 | File xuất | Sheet `Export Product Template`. Dòng 1 là mã cột, dòng 2 là tên, mọi ô dạng Text `@`, cố định A3, **bỏ cột sku**. Có `variant_code` thì vào file **BIENTHE**, không có thì vào file **MODEL**. Tên file `PIM_<tên tab 31 ký tự>_MODEL/BIENTHE_<thời gian>.xlsx`. |
| 13 | Tên cột khác nhau | Tên cột so khớp **không phân biệt hoa/thường, dấu, khoảng trắng, gạch dưới** (danh sách tên chấp nhận: `TEN_COT` trong pim_core.py). Mapping/CMS export không nhận ra tiêu đề → đọc **theo vị trí cột** như desktop. |
| 14 | Biến đổi hàng loạt | Áp **sau** đơn vị của rule cũ; thứ tự: sửa tay → Không/Đang cập nhật → đơn vị → biến đổi. Không bao giờ áp vào cột FILTER. |
| 15 | Cột FILTER | Mã có chữ `filter` **và không có** `tskt` (vd `filter_tskt_master` = "Bộ lọc" là cột TSKT). |
| 12 | Tiêu đề file mẫu mới | MAPPING TSKT MOI nhận cả `MÃ THUỘC TÍNH TSKT` / `MÃ MASTER` / `TÊN MASTER` / `MÃ HỌ`. File 66.py đi kèm cũng đã được sửa để đọc được tiêu đề mới. |

## Chỗ web khác desktop (có chủ đích)

- **Option có tên nhưng mã trống trong DATA PIM:** desktop ghi `215361, ` (thừa dấu phẩy), web bỏ mã trống và ghi log.
- **Tính năng chỉ có trên web** (chỉ có tác dụng khi được dùng):
  - quy đổi FILTER
  - đề xuất sửa CMS
  - quy tắc kiểm tra
  - map theo tên (bật tay)
- **Cảnh báo thêm:** `model_code` lặp ở nhiều dòng file MODEL. Desktop cũng xuất y như vậy, web chỉ cảnh báo thêm.
- **Tách SKU không có giá trị** khỏi file import (bật mặc định ở trang Xuất, tắt được). Desktop vẫn xuất dòng trống; web đưa các SKU này vào file xin data CMS.

## Rule bổ sung (web-1.6)

| # | Quy tắc | Chi tiết |
|---|---|---|
| 16 | Biến thể theo màu | Khi nạp danh sách SKU / file export PIM: nếu cột **Mã họ biến thể** (`family_variant_code`) có chữ **`color`** (vd `lvl_1_color_iden_master`) thì SKU đó coi là biến thể theo màu: bỏ `variant_code` khi nạp, nên xuất vào file import **MODEL**, không phải **BIENTHE**. Họ biến thể khác (không có chữ color) giữ nguyên rule cũ. |
| 17 | Ký tự ẩn (web-2.0) | Khi đọc file và khi ghi file import, tool **bỏ** ký tự ẩn: zero-width (U+200B–U+200D, U+2060), BOM (U+FEFF), soft hyphen (U+00AD), ký tự điều khiển, ký tự định hướng (U+202A–U+202E), và `\r` (xlsx ghi thành `_x000D_` trên PIM); khoảng trắng lạ (NBSP, en/em-space…) đổi thành khoảng trắng thường. Giữ `\n`/`\t`. Ô không có ký tự ẩn ghi **y nguyên** → file import vẫn giống hệt desktop. |
| 18 | Kiểm chứng SKU ↔ DATA SP (web-2.0) | Sau khi map, mọi ô kết quả được truy ngược về DATA SP **của chính SKU đó** theo đúng mapping PROPERTYID: TSKT phải có trong PROPVALUE, mã FILTER phải suy ra được từ CMS. Báo LỆCH SKU / không có nguồn / lệch IMPORT / lệch ngành. Chỉ kiểm, không sửa. Lỗi mức CAO chặn xuất (cần tick xác nhận). |
