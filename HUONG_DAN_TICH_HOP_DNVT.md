# Hướng dẫn thêm module "Đề nghị vật tư" vào hệ thống

## 1. Thêm 3 file HTML mới vào Apps Script
Tạo 3 file mới trong Apps Script Editor (Tệp > Tạo tệp mới > HTML), đặt tên **chính xác**:
- `dnvt` → dán nội dung file `dnvt.html`
- `dnvt_CSS` → dán nội dung file `dnvt_CSS.html`
- `dnvt_JS` → dán nội dung file `dnvt_JS.html`

## 2. Bổ sung Code.gs
Mở file `Code.gs`, kéo xuống cuối file, dán toàn bộ nội dung trong file
`Code_gs_them_moi_DNVT.txt` vào (thứ tự hàm trong Apps Script không quan trọng).

⚠️ Kiểm tra lại 3 hằng số đầu module, sửa nếu cần:
```javascript
var DNVT_NGUOI_DE_NGHI = "Hoàng Thị Hà";
var DNVT_TRUONG_PHONG  = "Nguyễn Thị Xuân Đào";
var DNVT_PHONG_BO_PHAN = "Chăm sóc học sinh";
```

## 3. Sửa `Index.html` — thêm nút menu + khung trang mới

**3a.** Thêm 1 nút vào `.menu-grid` (sau nút "Trông trưa"):
```html
      <!-- Nút 5: Đề nghị vật tư -->
      <button class="menu-card" onclick="moChucNang('dnvtPage')">
        <div class="card-icon">📦</div>
        <p class="card-title">Đề nghị vật tư</p>
      </button>
```

**3b.** Thêm 1 khung màn hình mới, đặt sau khối `<!-- MÀN HÌNH 5: TRÔNG TRƯA -->` (trước dòng `<?!= include('Index_JS'); ?>`):
```html
  <!-- MÀN HÌNH 6: ĐỀ NGHỊ VẬT TƯ -->
  <div id="dnvtPage" class="screen-page" style="display: none;">
    <button style="margin: 10px; padding: 8px 12px; border-radius: 20px;" onclick="veTrangChu()">⬅ Trang chủ</button>
    <?!= include('dnvt'); ?>
  </div>
```

## 4. Sửa `Index_JS.html` — thêm `dnvtPage` vào logic chuyển trang

Trong hàm `moChucNang(pageId)`, thêm dòng ẩn `dnvtPage` vào cùng nhóm với các trang khác, và thêm nhánh gọi hàm tải dữ liệu:

```javascript
function moChucNang(pageId) {
  document.getElementById("mainPage").style.display = "none";
  document.getElementById("cshsPage").style.display = "none";
  document.getElementById("trucCongPage").style.display = "none";
  document.getElementById("nhanSuPage").style.display = "none";
  document.getElementById("trongTruaPage").style.display = "none";
  document.getElementById("dnvtPage").style.display = "none";        // THÊM DÒNG NÀY

  document.getElementById(pageId).style.display = "block";
  hienHeaderToTruong(true);

  if (pageId === 'cshsPage' && typeof taiDanhSachLopBaoCao === 'function') {
    taiDanhSachLopBaoCao();
  } else if (pageId === 'trucCongPage' && typeof taiDuLieuTuSheet === 'function') {
    taiDuLieuTuSheet();
  } else if (pageId === 'nhanSuPage' && typeof taiDanhSachGV === 'function') {
    taiDanhSachGV();
  } else if (pageId === 'trongTruaPage' && typeof taiDanhSachGVTrongTrua === 'function') {
    taiDanhSachGVTrongTrua();
  } else if (pageId === 'dnvtPage' && typeof taiDanhSachDNVT === 'function') {   // THÊM NHÁNH NÀY
    taiDanhSachDNVT();
  }
}
```

Và thêm dòng ẩn `dnvtPage` vào hàm `veTrangChu()`:
```javascript
function veTrangChu() {
  document.getElementById("cshsPage").style.display = "none";
  document.getElementById("trucCongPage").style.display = "none";
  document.getElementById("nhanSuPage").style.display = "none";
  document.getElementById("trongTruaPage").style.display = "none";
  document.getElementById("dnvtPage").style.display = "none";        // THÊM DÒNG NÀY
  document.getElementById("mainPage").style.display = "block";
  hienHeaderToTruong(false);
}
```

## 5. Yêu cầu bắt buộc đối với sheet `Template_DNVT`

Vì hệ thống **tìm vị trí ô theo nhãn chữ** (không hard-code số dòng/cột — để cô
thoải mái chỉnh sửa mẫu về sau mà không vỡ code), sheet `Template_DNVT` trong
**mỗi file năm học** (CSHS_2627, CSHS_2728...) bắt buộc phải có sẵn, còn nguyên
các nhãn chữ sau (không phân biệt hoa/thường/dấu, nhưng không được đổi tên nhãn):

| Nhãn cần có trong ô | Ý nghĩa |
|---|---|
| `Ngày:` | Ngày lập phiếu (tự điền ngày hiện tại) |
| `Người đề xuất` | Tự điền `DNVT_NGUOI_DE_NGHI` |
| `Trưởng phòng/bộ phận` | Tự điền `DNVT_TRUONG_PHONG` |
| `Phòng/Bộ phận` (không có chữ "Trưởng" phía trước) | Tự điền `DNVT_PHONG_BO_PHAN` |
| `Ngày sử dụng` | Tự điền "Tháng MM/YYYY" |
| `TÊN MẶT HÀNG` | Dòng tiêu đề bảng — **bắt buộc đúng 7 cột liên tiếp theo thứ tự**: TT, Tên mặt hàng, ĐVT, Số lượng, Đơn giá, Thành tiền, Ghi chú |
| `TỔNG` | Dòng tổng cộng, nằm ngay dưới bảng hàng hóa (ô Thành tiền của dòng này sẽ tự động điền công thức SUM) |

Giá trị ghi vào phiếu (Người đề xuất, Ngày, Ngày sử dụng...) sẽ được ghi vào
**ô ngay bên phải nhãn** (tự động bỏ qua vùng ô đã gộp của nhãn nếu có).

Ô "Phòng/Bộ phận" và ô "Trưởng phòng/bộ phận" **phải khác dòng nhau** như
trong ảnh mẫu — hệ thống dựa vào việc nhãn nào bắt đầu bằng chữ gì để phân
biệt 2 ô này, nên không được gộp chung 1 dòng.

Số dòng trống có sẵn giữa dòng tiêu đề bảng và dòng TỔNG **không quan trọng**
(có thể chỉ để 0-3 dòng mẫu) — hệ thống sẽ tự động chèn thêm dòng (kèm sao
chép định dạng/viền của dòng liền trên) khi nhập nhiều mặt hàng hơn số dòng
trống hiện có.

## 6. Cách hoạt động (tóm tắt)
- Cô chọn **Tháng sử dụng** → hệ thống tải danh sách mặt hàng đã lưu của
  đúng tháng đó (nếu sheet "Tháng X.YY" chưa tồn tại thì coi như trống).
- Thêm mặt hàng → vào **danh sách chờ** (chưa ghi Sheet).
- Bấm **💾 Lưu danh sách chờ** → hệ thống tự tạo sheet tháng đó (copy từ
  `Template_DNVT`, điền sẵn Người đề xuất/Trưởng phòng/Phòng-Bộ phận/Ngày sử
  dụng) nếu chưa có, rồi gộp thêm các mặt hàng mới vào, tự đánh lại STT và
  tính lại TỔNG.
- Mỗi mặt hàng **đã lưu** có nút Xóa riêng, xóa ngay trên Sheet (có xác nhận).
- Nút **Excel / PDF** xuất trực tiếp đúng sheet của tháng đang xem (không tạo
  sheet tạm, vì bản thân sheet tháng đã là phiếu hoàn chỉnh).
- Tab "Tra cứu" và "Thống kê" hiện để dạng "đang phát triển" giống mẫu chung
  của các module khác — nếu cô cần (VD: xem tổng chi theo quý, so sánh giữa
  các tháng...), tôi làm thêm được, cứ nói cụ thể là cô muốn xem gì.
