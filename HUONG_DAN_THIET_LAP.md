# Hướng dẫn thiết lập cấu trúc mới (1 file Google Sheet / năm học)

## 1. Tạo file Google Sheet cho từng năm học

Tạo 2 file Google Sheet mới (không cần để trong folder đặc biệt nào, để đâu
cũng được vì hệ thống tra theo ID chứ không dò theo tên/folder):

| Tên file gợi ý     | Mã năm học dùng trong hệ thống |
|---------------------|-------------------------------|
| CSHS_2026-2027      | `2627`                        |
| CSHS_2027-2028      | `2728`                        |

Tên file bạn đặt gì cũng được (chỉ để bạn dễ nhận diện) — hệ thống không đọc
tên file. Cái hệ thống cần là **ID file**, lấy trong URL:

```
https://docs.google.com/spreadsheets/d/AbC123XyzID_NAM_O_DAY/edit
                                        └────────┬────────┘
                                              đây là ID
```

## 2. Trong MỖI file, tạo đúng 4 tab với tên chính xác (phân biệt hoa/thường)

### Tab `BaoCao`
Dòng 1 là tiêu đề, dữ liệu bắt đầu từ dòng 2, script tự ghi:

| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| Dấu thời gian | Ngày | Lớp | Quản lý giờ ăn | Quản lý giờ ngủ | Báo cơm & trang trí lớp | Công tác phối hợp | Công tác vệ sinh | Phát triển bản thân | Văn hóa tổ chức |

→ Bạn chỉ cần tạo tab trống, gõ đúng 10 tiêu đề cột ở dòng 1. Script sẽ tự
động ghi dữ liệu từ dòng 2 trở đi.

### Tab `TrucCong`
| A | B | C | D | E | F |
|---|---|---|---|---|---|
| Thứ | Ngày | Cổng 1 - Người 1 | Cổng 1 - Người 2 | Cổng 2 - Người 1 | Cổng 2 - Người 2 |

→ Nếu bạn không tự tạo, script sẽ **tự tạo tab này** (kèm tiêu đề) trong lần
đầu tiên lưu lịch trực — không bắt buộc tạo tay.

### Tab `DSGV` (danh sách giáo viên/nhân sự)
Giữ nguyên các cột bạn đang dùng ở sheet `DSGV_2627`/`DSGV_2526` cũ (copy y
nguyên header + dữ liệu qua). Lưu ý bắt buộc:
- Phải có **1 cột tên đúng là "Lớp"** (hoặc "lop") — dùng để hệ thống tự suy
  ra danh sách lớp cho tab Báo cáo, Trực cổng & Trông trưa.
- Nên có **1 cột tên chứa "Bộ phận"** (VD: GVCN, HC, NK, TPT...) — dùng để
  module Trông trưa tự động điền Bộ phận khi chọn Họ và tên. Nếu không có
  cột này, người dùng vẫn nhập tay Bộ phận bình thường, không bắt buộc.
- Cột nào tên chứa "xe" hoặc "đi xe" sẽ tự hiển thị dạng chọn (có đi xe / không).
- Cột nào tên chứa "điện thoại"/"sđt"/"phone" sẽ tự thêm số 0 phía trước nếu thiếu.
- Ngoài ra muốn thêm bớt cột gì cũng được, form nhập sẽ tự sinh theo đúng tiêu đề cột.

### Tab `TrongTrua` (bảng chấm công trông trưa/trông ngủ)
Không bắt buộc tạo tay — script sẽ **tự tạo tab này** (kèm 5 cột tiêu đề:
Dấu thời gian, Ngày, Họ và tên, Bộ phận, Lớp) trong lần đầu tiên lưu dữ liệu
Trông trưa. Mỗi dòng ứng với 1 người trông 1 lớp vào 1 ngày cụ thể. Khi bấm
"Xuất Excel" theo tháng ở màn hình Trông trưa, hệ thống tự tổng hợp dữ liệu
tab này thành bảng chấm công dạng ma trận (người theo hàng, ngày làm việc
trong tháng theo cột) giống mẫu bảng chấm công cũ của trường.

### Tab `DanhMucLoi` (danh mục lỗi vi phạm để hiện dạng thẻ chọn nhanh)
Dòng 1 là tiêu đề (nội dung tùy bạn đặt cho dễ nhìn), dữ liệu từ dòng 2.
**Thứ tự 6 cột A→F cố định theo đúng nhóm sau** (đây là quy ước script đọc
theo vị trí cột, không theo tên tiêu đề):

| A (Giờ ăn) | B (Giờ ngủ) | C (Báo cơm & trang trí) | D (Vệ sinh) | E (Bản thân) | F (Văn hóa) |
|---|---|---|---|---|---|
| Đi trễ giờ ăn | Ngủ trễ | Quên báo cơm | Không dọn bàn | ... | ... |
| Bỏ bữa | Nói chuyện giờ ngủ | ... | ... | ... | ... |

→ Copy nguyên dữ liệu từ sheet `DanhMucLoi` cũ sang, giữ nguyên thứ tự cột.

## 3. Khai báo ID file vào Script Properties

Vào Apps Script Editor của project → biểu tượng ⚙️ **Project Settings** (góc
trái) → kéo xuống mục **Script Properties** → **Add script property**, thêm
lần lượt:

| Property   | Value                          |
|------------|---------------------------------|
| `CSHS_2627`| ID file Google Sheet năm 2026-2027 |
| `CSHS_2728`| ID file Google Sheet năm 2027-2028 |

Không cần sửa gì trong code khi thêm năm học mới — chỉ cần thêm 1 dòng
Script Property mới với key `CSHS_<mã năm học 4 số>`, và thêm 1 `<option>`
tương ứng trong dropdown ở `Index.html`.

## 4. Quyền truy cập

Vì `SpreadsheetApp.openById()` chỉ cần file đó cùng thuộc quyền truy cập của
tài khoản chạy script (thường là chính bạn, nếu bạn là người deploy web app
"Execute as: Me"), bạn **không cần chia sẻ quyền gì thêm** miễn là các file
năm học cũng thuộc cùng tài khoản Google với người triển khai web app.

## 5. Dữ liệu cũ

Bạn đã xác nhận sẽ tự copy tay dữ liệu cũ (từ các sheet `BaoCao`, `DSGV_2627`,
`DSGV_2526`, `TrucCong_2627`... trong file container hiện tại) sang đúng tab
tương ứng ở 2 file mới. Vài lưu ý khi copy tay:
- Cột "Ngày" nên copy dạng text `yyyy-MM-dd` hoặc để Google Sheets tự nhận
  diện là Date — cả 2 cách hệ thống đều đọc được (dùng `new Date(...)`).
- Với `BaoCao`: hiện dữ liệu cũ đang gộp chung mọi năm học vào 1 sheet, bạn
  cần tự lọc theo cột "Ngày" để tách đúng report nào thuộc năm học nào trước
  khi dán vào từng file.

## 6. Sau khi hoàn tất

Deploy lại Web App (Deploy → Manage deployments → Edit → New version) để áp
dụng code mới.
