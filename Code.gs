function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hệ thống Quản lý CSHS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ==========================================================================
   CẤU HÌNH FILE THEO NĂM HỌC
   --------------------------------------------------------------------------
   Mỗi năm học ứng với DUY NHẤT 1 file Google Sheet riêng (không còn dùng
   folder Drive để dò tìm). ID của từng file được lưu trong Script Properties
   (Project Settings > Script Properties trong Apps Script Editor), với key
   dạng "CSHS_<mã năm học>", ví dụ:

     CSHS_2627   ->   ID file Google Sheet năm học 2026-2027
     CSHS_2728   ->   ID file Google Sheet năm học 2027-2028
     
   Mỗi file năm học BẮT BUỘC có đủ 4 tab, đặt tên chính xác (phân biệt hoa
   thường) như sau:
     - BaoCao
     - TrucCong
     - DSGV
     - DanhMucLoi

   Xem chi tiết cấu trúc cột từng tab trong file HUONG_DAN_THIET_LAP.md.
   ========================================================================== */

var MA_NAM_HOC_MAC_DINH = "2627";

// Mở đúng file Google Sheet của năm học tương ứng (tra theo Script Properties)
function moFileNamHoc(maNamHoc) {
  var ma = (maNamHoc || MA_NAM_HOC_MAC_DINH).toString().trim();
  var props = PropertiesService.getScriptProperties();
  var fileId = props.getProperty("CSHS_" + ma);

  if (!fileId) {
    throw new Error(
      "Chưa cấu hình file cho năm học mã '" + ma + "'. " +
      "Vào Apps Script > Project Settings > Script Properties, thêm key 'CSHS_" + ma + "' " +
      "với giá trị là ID file Google Sheet của năm học đó."
    );
  }
  return SpreadsheetApp.openById(fileId);
}

// Lấy đúng 1 tab (sheet) trong file của năm học, báo lỗi rõ ràng nếu thiếu tab
function laySheetNamHoc(maNamHoc, tenTab) {
  var ss = moFileNamHoc(maNamHoc);
  var sheet = ss.getSheetByName(tenTab);
  if (!sheet) {
    throw new Error(
      "Không tìm thấy tab '" + tenTab + "' trong file năm học mã '" +
      (maNamHoc || MA_NAM_HOC_MAC_DINH) + "'. Kiểm tra lại tên tab (phân biệt hoa/thường)."
    );
  }
  return sheet;
}

// 1. LƯU BÁO CÁO MỚI (Gộp Thẻ vi phạm & Ghi chú vào chung 1 Cột)
function luuBaoCao(data) {
  try {
    var sheet = laySheetNamHoc(data.maNamHoc, "BaoCao");

    // Hàm gộp các Thẻ vi phạm và Ghi chú lại với nhau
    function gopNoiDung(arrChips, ghiChu) {
      var parts = [];
      if (Array.isArray(arrChips) && arrChips.length > 0) {
        parts.push(arrChips.join(", "));
      }
      if (ghiChu && ghiChu.trim() !== "") {
        parts.push(ghiChu.trim());
      }
      return parts.join(" - "); // Ví dụ: "Đi sai đường - Gia Hoàng"
    }

    // Xử lý ngày chuẩn múi giờ
    var ngayLuu = "";
    if (data.ngay) {
      var parts = data.ngay.split("-");
      if (parts.length === 3) {
        var dateObj = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
        ngayLuu = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        ngayLuu = data.ngay;
      }
    }

    // Mảng 10 Cột khớp 100% tiêu đề Sheet
    var rowData = [
      new Date(),                                              // Cột A (1): Dấu thời gian
      ngayLuu,                                                 // Cột B (2): Ngày
      data.lop || "",                                          // Cột C (3): Lớp
      gopNoiDung(data.viPhamGioAn, data.ghichuGioAn),          // Cột D (4): Quản lý giờ ăn
      gopNoiDung(data.viPhamGioNgu, data.ghichuGioNgu),        // Cột E (5): Quản lý giờ ngủ
      gopNoiDung(data.viPhamBaoCom, data.ghichuBaoCom),        // Cột F (6): Báo cơm & trang trí
      data.ghichuPhoiHop || "",                                // Cột G (7): Công tác phối hợp
      gopNoiDung(data.viPhamVeSinh, data.ghichuVeSinh),        // Cột H (8): Công tác vệ sinh
      gopNoiDung(data.viPhamBanThan, data.ghichuBanThan),      // Cột I (9): Phát triển bản thân
      gopNoiDung(data.viPhamVanHoa, data.ghichuVanHoa)         // Cột J (10): Văn hóa tổ chức
    ];

    // Lưu dòng dữ liệu vào Sheet
    sheet.appendRow(rowData);

    // --- TỰ ĐỘNG KẺ BORDER CHO DÒNG VỪA THÊM ---
    var lastRow = sheet.getLastRow();
    var lastColumn = rowData.length; // 10 cột từ A đến J
    sheet.getRange(lastRow, 1, 1, lastColumn).setBorder(
      true, true, true, true, true, true,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
    // ------------------------------------------

    return "Gửi báo cáo thành công!";
  } catch (error) {
    return "Lỗi: " + error.toString();
  }
}

// 2. TRA CỨU BÁO CÁO (Đọc theo 10 Cột)
function layDanhSachBaoCao(filter) {
  try {
    var sheet = laySheetNamHoc(filter.maNamHoc, "BaoCao");
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var result = [];
    var tuNgay = null, denNgay = null;

    if (filter.tuNgay) {
      var p1 = filter.tuNgay.split("-");
      tuNgay = new Date(p1[0], p1[1] - 1, p1[2], 0, 0, 0);
    }
    if (filter.denNgay) {
      var p2 = filter.denNgay.split("-");
      denNgay = new Date(p2[0], p2[1] - 1, p2[2], 23, 59, 59);
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;

      var rawDate = new Date(row[1]);
      var ngayBC = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate(), 12, 0, 0);

      var matchTuNgay = !tuNgay || ngayBC >= tuNgay;
      var matchDenNgay = !denNgay || ngayBC <= denNgay;

      var lopInSheet = row[2] ? row[2].toString().trim() : "";
      var lopFilter = filter.lop ? filter.lop.toString().trim() : "";
      var matchLop = !filter.lop || filter.lop === "Tất cả lớp" || lopInSheet === lopFilter;

      if (matchTuNgay && matchDenNgay && matchLop) {
        var chiTietHTML = [];
        var chiTietText = [];

        var addSec = function(title, val) {
          if (val && val.toString().trim() !== "") {
            chiTietHTML.push("<b>" + title + ":</b> " + val);
            chiTietText.push("• " + title + ": " + val);
          }
        };

        addSec("Quản lý giờ ăn", row[3]);
        addSec("Quản lý giờ ngủ", row[4]);
        addSec("Báo cơm & trang trí", row[5]);
        addSec("Công tác phối hợp", row[6]);
        addSec("Công tác vệ sinh", row[7]);
        addSec("Phát triển bản thân", row[8]);
        addSec("Văn hóa tổ chức", row[9]);

        var ngayFormatted = Utilities.formatDate(ngayBC, Session.getScriptTimeZone(), "dd/MM/yyyy");

        result.push({
          rowIndex: i + 1,
          ngay: ngayFormatted,
          lop: row[2],
          noiDung: chiTietHTML.join("<br>") || "Không có vi phạm/ghi chú",
          rawTextZalo: "📌 BÁO CÁO CSHS - LỚP " + row[2] + " (" + ngayFormatted + ")\n" + (chiTietText.join("\n") || "Không có vi phạm.")
        });
      }
    }
    return result.reverse();
  } catch (error) {
    return [];
  }
}

// 3. THỐNG KÊ (Đã bổ sung tính cả các lớp 0 lần vi phạm)
function layThongKe(filter) {
  try {
    var sheet = laySheetNamHoc(filter && filter.maNamHoc, "BaoCao");
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { topLoi: [], topLop: [], topLopItNhat: [] };

    var loiCount = {};
    var lopCount = {};

    // 1. Khởi tạo danh sách tất cả các lớp với số lần vi phạm ban đầu = 0
    var dslop = layDanhSachLop(filter && filter.maNamHoc); // Lấy danh sách lớp đúng theo năm học đang xem thống kê
    dslop.forEach(function(l) {
      if (l && l !== "Tất cả lớp") {
        lopCount[l] = 0;
      }
    });

    var tuNgay = null, denNgay = null;
    if (filter && filter.tuNgay) {
      var p1 = filter.tuNgay.split("-");
      tuNgay = new Date(p1[0], p1[1] - 1, p1[2], 0, 0, 0);
    }
    if (filter && filter.denNgay) {
      var p2 = filter.denNgay.split("-");
      denNgay = new Date(p2[0], p2[1] - 1, p2[2], 23, 59, 59);
    }

    // 2. Đếm số lần vi phạm thực tế theo khoảng thời gian
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;

      var rawDate = new Date(row[1]);
      var ngayBC = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate(), 12, 0, 0);

      var matchTuNgay = !tuNgay || ngayBC >= tuNgay;
      var matchDenNgay = !denNgay || ngayBC <= denNgay;

      if (matchTuNgay && matchDenNgay) {
        var lop = row[2] ? row[2].toString().trim() : "";
        if (lop) {
          lopCount[lop] = (lopCount[lop] || 0) + 1;
        }

        // Đếm các cột vi phạm
        [3, 4, 5, 7, 8, 9].forEach(function(colIdx) {
          var str = row[colIdx];
          if (str) {
            var cleanStr = str.toString().split("-")[0];
            var arr = cleanStr.split(",");
            arr.forEach(function(item) {
              var key = item.trim();
              if (key) loiCount[key] = (loiCount[key] || 0) + 1;
            });
          }
        });
      }
    }

    // Top 5 nhiều nhất (Chỉ lấy các lớp có vi phạm > 0)
    var sortObj = function(obj) {
      return Object.keys(obj)
        .map(function(k) { return { name: k, count: obj[k] }; })
        .filter(function(item) { return item.count > 0; })
        .sort(function(a, b) { return b.count - a.count; })
        .slice(0, 5);
    };

    // Top 5 ít nhất (Sắp xếp tăng dần, bao gồm cả lớp 0 lần)
    var sortObjAsc = function(obj) {
      return Object.keys(obj)
        .map(function(k) { return { name: k, count: obj[k] }; })
        .sort(function(a, b) { return a.count - b.count; })
        .slice(0, 5);
    };

    return {
      topLoi: sortObj(loiCount),
      topLop: sortObj(lopCount),
      topLopItNhat: sortObjAsc(lopCount)
    };
  } catch (e) {
    return { topLoi: [], topLop: [], topLopItNhat: [] };
  }
}

// 4. XÓA BÁO CÁO (theo đúng file năm học tương ứng)
function xoaMotBaoCao(rowIndex, maNamHoc) {
  try {
    var sheet = laySheetNamHoc(maNamHoc, "BaoCao");
    sheet.deleteRow(rowIndex);
    return "Đã xóa báo cáo thành công!";
  } catch (error) {
    return "Lỗi khi xóa: " + error.toString();
  }
}

// Lấy danh sách Lớp động (từ tab DSGV) và TỰ ĐỘNG TÁCH các lớp nối bằng dấu "+"
function layDanhSachLop(maNamHoc) {
  try {
    var sheet = laySheetNamHoc(maNamHoc, "DSGV");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];

    // 1. Tìm cột chứa tiêu đề "Lớp"
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var lopColIndex = -1;

    for (var i = 0; i < headers.length; i++) {
      var headerText = headers[i].toString().trim().toLowerCase();
      if (headerText === "lớp" || headerText === "lop") {
        lopColIndex = i + 1;
        break;
      }
    }

    if (lopColIndex === -1) return [];

    // 2. Lấy dữ liệu toàn bộ cột Lớp
    var data = sheet.getRange(2, lopColIndex, lastRow - 1, 1).getValues();
    var rawLopList = [];

    // 3. Xử lý tách các chuỗi dạng "4A0+4A2" hoặc "4A1 + 4B0"
    data.forEach(function(r) {
      var val = r[0] ? r[0].toString().trim() : "";
      if (val) {
        var subLops = val.split('+');
        subLops.forEach(function(sub) {
          var cleanSub = sub.trim();
          if (cleanSub) {
            rawLopList.push(cleanSub);
          }
        });
      }
    });

    // 4. Lọc bỏ trùng lặp
    var uniqueLop = Array.from(new Set(rawLopList));

    return uniqueLop;
  } catch (error) {
    Logger.log("Lỗi đọc danh sách lớp: " + error.toString());
    return [];
  }
}

// 5. Lấy Danh mục lỗi từ tab DanhMucLoi (theo đúng năm học)
function layDanhMucLoi(maNamHoc) {
  try {
    var sheet = laySheetNamHoc(maNamHoc, "DanhMucLoi");
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    var result = { viPhamGioAn: [], viPhamGioNgu: [], viPhamBaoCom: [], viPhamVeSinh: [], viPhamBanThan: [], viPhamVanHoa: [] };
    var keyMapping = ["viPhamGioAn", "viPhamGioNgu", "viPhamBaoCom", "viPhamVeSinh", "viPhamBanThan", "viPhamVanHoa"];

    for (var col = 0; col < keyMapping.length; col++) {
      var key = keyMapping[col];
      for (var row = 1; row < data.length; row++) {
        var val = data[row][col];
        if (val && val.toString().trim() !== "") result[key].push(val.toString().trim());
      }
    }
    return result;
  } catch (e) { return {}; }
}

// 6. Xóa báo cáo theo bộ lọc (dùng khi bấm nút "Xóa" hàng loạt ở Tab tra cứu)
function xoaBaoCaoFilter(filter) {
  try {
    var sheet = laySheetNamHoc(filter.maNamHoc, "BaoCao");
    var list = layDanhSachBaoCao(filter);
    if (list.length === 0) return "Không có báo cáo nào để xóa!";
    list.sort((a, b) => b.rowIndex - a.rowIndex);
    list.forEach(item => sheet.deleteRow(item.rowIndex));
    return "Đã xóa " + list.length + " báo cáo!";
  } catch (error) { return "Lỗi: " + error.toString(); }
}

function xuatFileBaoCaoChuan(filter, dinhDang) {
  var ss = moFileNamHoc(filter.maNamHoc);
  var sheet = ss.getSheetByName("BaoCao");
  var dataFull = sheet.getDataRange().getValues();
  var dsBaoCao = layDanhSachBaoCao(filter);

  // 1. Dọn dẹp sạch sẽ tất cả các Sheet tạm cũ còn sót lại trước đó (trong đúng file năm học này)
  xoaTatCaSheetTam(filter.maNamHoc);

  // 2. Tạo tên file & tên Sheet tạm trùng khớp nhau
  var lopTen = (filter && filter.lop) ? filter.lop : "Tất cả lớp";
  var tuNgayStr = filter.tuNgay ? Utilities.formatDate(new Date(filter.tuNgay), Session.getScriptTimeZone(), "d.M.yyyy") : "";
  var denNgayStr = filter.denNgay ? Utilities.formatDate(new Date(filter.denNgay), Session.getScriptTimeZone(), "d.M.yyyy") : "";

  var rangeDateStr = "";
  if (tuNgayStr && denNgayStr) {
    rangeDateStr = "_" + tuNgayStr + "-" + denNgayStr;
  } else if (tuNgayStr) {
    rangeDateStr = "_" + tuNgayStr;
  }

  var baseFileName = "BaoCao_" + lopTen + rangeDateStr;

  // 3. Đặt tên Sheet tạm TRÙNG VỚI tên file muốn xuất
  var tempSheetName = baseFileName;
  var tempSheet = ss.insertSheet(tempSheetName);

  // 4. DÒNG 1: Tiêu đề lớn chính giữa
  tempSheet.getRange("A1:I1").merge()
    .setValue("BÁO CÁO CHĂM SÓC HỌC SINH")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  var tuNgayTxt = filter.tuNgay ? Utilities.formatDate(new Date(filter.tuNgay), Session.getScriptTimeZone(), "dd/MM/yyyy") : "...";
  var denNgayTxt = filter.denNgay ? Utilities.formatDate(new Date(filter.denNgay), Session.getScriptTimeZone(), "dd/MM/yyyy") : "...";

  // 5. DÒNG 2: Thông tin đúng vị trí các cột
  tempSheet.getRange("C2").setValue("Tổ trưởng: Nguyễn Thị Xuân Đào").setFontWeight("bold");
  tempSheet.getRange("D2").setValue("Từ ngày:").setHorizontalAlignment("right");
  tempSheet.getRange("E2").setValue(tuNgayTxt).setHorizontalAlignment("center");
  tempSheet.getRange("F2").setValue("Đến ngày:").setHorizontalAlignment("right");
  tempSheet.getRange("G2").setValue(denNgayTxt).setHorizontalAlignment("center");

  // 6. DÒNG 3: Tiêu đề cột
  var headers = [
    "Ngày", "Lớp", "Quản lý giờ ăn",
    "Quản lý giờ ngủ", "Báo cơm & trang trí", "Công tác phối hợp",
    "Công tác vệ sinh", "Phát triển bản thân", "Văn hóa tổ chức"
  ];
  tempSheet.getRange(3, 1, 1, 9).setValues([headers]);

  // 7. Đưa dữ liệu từ Dòng 4
  var listReversed = dsBaoCao.slice().reverse();
  for (var k = 0; k < listReversed.length; k++) {
    var rIndex = listReversed[k].rowIndex - 1;
    if (dataFull[rIndex]) {
      var rowContent = dataFull[rIndex].slice(1, 10);
      if (rowContent[0] instanceof Date) {
        rowContent[0] = Utilities.formatDate(rowContent[0], Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      tempSheet.appendRow(rowContent);
    }
  }

  // 8. Định dạng khung bảng, độ rộng cột & chiều cao dòng
  var lastRow = tempSheet.getLastRow();
  if (lastRow >= 3) {
    var rangeTable = tempSheet.getRange(3, 1, lastRow - 2, 9);
    var headerRange = tempSheet.getRange(3, 1, 1, 9);

    rangeTable.setBorder(true, true, true, true, true, true);
    rangeTable.setWrap(true);
    rangeTable.setVerticalAlignment("middle");

    headerRange.setFontWeight("bold").setBackground("#e8f5e9").setHorizontalAlignment("center");

    var columnWidths = [100, 60, 240, 240, 240, 240, 240, 240, 240];
    for (var col = 1; col <= 9; col++) {
      tempSheet.setColumnWidth(col, columnWidths[col - 1]);
    }

    if (lastRow > 3) {
      tempSheet.getRange(4, 1, lastRow - 3, 2).setHorizontalAlignment("center");
    }

    tempSheet.setRowHeight(1, 40);
    tempSheet.setRowHeight(2, 25);
    tempSheet.setRowHeight(3, 28);
  }

  SpreadsheetApp.flush();

  // 9. Tạo URL Export
  var ssId = ss.getId();
  var sheetId = tempSheet.getSheetId();
  var exportUrl = "";

  if (dinhDang === 'pdf') {
    exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&gid=" + sheetId + "&size=A4&portrait=false&fitw=true&gridlines=true";
  } else {
    exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx&gid=" + sheetId;
  }

  return {
    success: true,
    downloadUrl: exportUrl,
    tempSheetName: tempSheetName,
    fileName: baseFileName + (dinhDang === 'pdf' ? '.pdf' : '.xlsx'),
    maNamHoc: filter.maNamHoc || MA_NAM_HOC_MAC_DINH // trả về để client dùng khi gọi xoaSheetTam
  };
}

// Hàm quét xóa tự động tất cả các sheet tạm cũ (trong đúng file năm học)
function xoaTatCaSheetTam(maNamHoc) {
  var ss = moFileNamHoc(maNamHoc);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    if (sName.indexOf("TEMP_") === 0 || sName.indexOf("TempExport") !== -1 || (sName.indexOf("BaoCao_CSHS_") === 0 && sheets.length > 1)) {
      try { ss.deleteSheet(sheets[i]); } catch(e) {}
    }
  }
}

// Hàm dọn dẹp sheet tạm sau khi tải xong (trong đúng file năm học)
function xoaSheetTam(sheetName, maNamHoc) {
  if (!sheetName) return;
  var ss = moFileNamHoc(maNamHoc);
  var targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(targetSheet); } catch(e) {}
  }
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// Tìm chỉ số cột "Họ tên" theo tiêu đề (không hardcode cột 1), để không bị lệch
// khi tab DSGV có thêm/bớt cột (ví dụ thêm cột STT ở đầu). Nếu không tìm thấy,
// trả về 0 (cột đầu tiên) như phương án dự phòng.
function timCotHoTen(headers) {
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || "").toString().trim().toLowerCase();
    if (h.indexOf("họ tên") !== -1 || h.indexOf("họ và tên") !== -1 || h.indexOf("tên") !== -1) {
      return i;
    }
  }
  return 0;
}

// 1. Tải danh sách Giáo viên & Lịch trực theo tuần (theo đúng file năm học)
function layDuLieuTrucCong(tuNgayStr, maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);

    // 1. Lấy danh sách GV (tự tìm đúng cột "Họ tên" theo tiêu đề, không hardcode cột 1)
    var sheetGV = ss.getSheetByName("DSGV");
    var dsgv = [];
    if (sheetGV && sheetGV.getLastRow() >= 2) {
      var lastColGV = sheetGV.getLastColumn();
      var allGV = sheetGV.getRange(1, 1, sheetGV.getLastRow(), lastColGV).getDisplayValues();
      var headersGV = allGV[0];
      var colHoTen = timCotHoTen(headersGV);
      dsgv = allGV.slice(1).map(function(row) { return row[colHoTen]; }).filter(String);
    }

    // 2. Lấy lịch trực
    var sheetTruc = ss.getSheetByName("TrucCong");
    var lichTruc = [];
    if (sheetTruc && sheetTruc.getLastRow() >= 2) {
      var data = sheetTruc.getRange(2, 1, sheetTruc.getLastRow() - 1, 6).getValues();
      lichTruc = data.map(function(row) {
        var d = new Date(row[1]);
        var ngayFormatted = isNaN(d.getTime()) ? "" : Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        return {
          thu: row[0],
          ngay: ngayFormatted,
          c1_p1: row[2] || "",
          c1_p2: row[3] || "",
          c2_p1: row[4] || "",
          c2_p2: row[5] || ""
        };
      });
    }

    return { dsgv: dsgv, lichTruc: lichTruc };
  } catch (err) {
    return { dsgv: [], lichTruc: [], loi: err.toString() };
  }
}

// 2. Lưu lịch trực vào tab TrucCong (đúng file năm học)
function luuLichTrucCong(data, maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrucCong");

    if (!sheet) {
      sheet = ss.insertSheet("TrucCong");
      sheet.appendRow(["Thứ", "Ngày", "Cổng 1 - Người 1", "Cổng 1 - Người 2", "Cổng 2 - Người 1", "Cổng 2 - Người 2"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e8eaf6");
    }

    var existingData = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues() : [];

    data.forEach(function(row) {
      var foundIndex = -1;

      for (var i = 0; i < existingData.length; i++) {
        var sheetNgay = Utilities.formatDate(new Date(existingData[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (sheetNgay === row.ngay) {
          foundIndex = i + 2;
          break;
        }
      }

      if (foundIndex !== -1) {
        sheet.getRange(foundIndex, 1, 1, 6).setValues([[row.thu, row.ngay, row.c1_p1, row.c1_p2, row.c2_p1, row.c2_p2]]);
      } else {
        sheet.appendRow([row.thu, row.ngay, row.c1_p1, row.c1_p2, row.c2_p1, row.c2_p2]);
      }
    });

    return "Lưu lịch trực thành công!";
  } catch (err) {
    return "Lỗi: " + err.toString();
  }
}

// 1. Lấy dữ liệu động bao gồm Tiêu đề (Hàng 1) và Danh sách (tab DSGV, đúng file năm học)
function layDanhSachGVTheoSheet(maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("DSGV");
    if (!sheet) {
      return { headers: [], data: [], loi: "Không tìm thấy tab 'DSGV' trong file năm học này." };
    }
    if (sheet.getLastRow() < 1) return { headers: [], data: [] };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    var headers = allValues[0];
    var data = [];

    for (var i = 1; i < allValues.length; i++) {
      if (allValues[i].join("").trim() !== "") {
        data.push({
          rowNum: i + 1,
          values: allValues[i]
        });
      }
    }
    return { headers: headers, data: data };
  } catch (err) {
    return { headers: [], data: [], loi: err.toString() };
  }
}

// Hàm phụ trợ: bỏ dấu tiếng Việt (dùng để so khớp tên tiêu đề không phân biệt dấu)
function boDauTiengVietGS(str) {
  if (!str) return "";
  return str.toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d").replace(/Đ/g, "D")
    .toLowerCase();
}

// Xác định các cột KHÓA HOÀN TOÀN theo tên tiêu đề (STT, Mã nhân viên, Thâm niên):
// đây thường là các cột có công thức tự tính trên Sheet (STT tăng dần, Mã NV tự sinh,
// Thâm niên tính theo ngày vào làm...), nên KHÔNG BAO GIỜ được phép ghi đè từ dữ liệu
// GUI gửi lên, dù là thêm mới hay cập nhật.
function layDanhSachCotKhoaDSGV(headers) {
  return headers.map(function(h) {
    var t = boDauTiengVietGS(h);
    return t.includes("stt") || t.includes("ma nhan vien") || t.includes("tham nien");
  });
}

// 2. Lưu/Cập nhật dòng dữ liệu nhân sự (tab DSGV, đúng file năm học)
function luuNhanSuToSheet(maNamHoc, rowValues, rowNum) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("DSGV");
    if (!sheet) return { success: false, message: "Không tìm thấy tab DSGV!" };

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var isCotKhoa = layDanhSachCotKhoaDSGV(headers);

    if (rowNum && rowNum > 1) {
      // CẬP NHẬT: chỉ ghi các cột KHÔNG bị khóa. Các cột STT/Mã nhân viên/Thâm niên
      // giữ nguyên giá trị (hoặc công thức) sẵn có trên Sheet, không lấy từ GUI.
      for (var i = 0; i < rowValues.length; i++) {
        if (!isCotKhoa[i]) {
          sheet.getRange(rowNum, i + 1).setValue(rowValues[i]);
        }
      }
      return { success: true, message: "Cập nhật thành công!" };
    } else {
      sheet.appendRow(rowValues);

      var newRow = sheet.getLastRow();
      if (newRow > 2) {
        var numCols = rowValues.length;
        var sourceRange = sheet.getRange(newRow - 1, 1, 1, numCols);
        var targetRange = sheet.getRange(newRow, 1, 1, numCols);

        // 1. Lấy danh sách công thức của dòng ngay phía trên
        var formulas = sourceRange.getFormulas()[0];

        // 2. Sao chép toàn bộ (gồm Định dạng + Công thức) từ dòng trên xuống dòng mới
        //    -> các cột STT/Mã NV/Thâm niên (nếu là công thức) sẽ tự tính lại đúng cho dòng mới
        sourceRange.copyTo(targetRange);

        // 3. Chỉ ghi đè dữ liệu từ form vào các cột KHÔNG chứa công thức
        //    VÀ KHÔNG thuộc nhóm cột bị khóa (STT, Mã nhân viên, Thâm niên)
        for (var j = 0; j < numCols; j++) {
          if (formulas[j] === "" && !isCotKhoa[j]) {
            targetRange.getCell(1, j + 1).setValue(rowValues[j]);
          }
        }
      }

      return { success: true, message: "Thêm mới thành công!" };
    }

  } catch (err) {
    return { success: false, message: "Lỗi hệ thống: " + err.toString() };
  }
}

// 3. Xóa dòng nhân sự (tab DSGV, đúng file năm học)
function xoaNhanSuFromSheet(maNamHoc, rowNum) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("DSGV");
    if (sheet && rowNum > 1) {
      sheet.deleteRow(rowNum);
      return { success: true, message: "Xóa thành công!" };
    }
    return { success: false, message: "Lỗi dòng cần xóa!" };
  } catch (err) {
    return { success: false, message: "Lỗi hệ thống: " + err.toString() };
  }
}

/* ==========================================================================
   MODULE: TRÔNG TRƯA (Bảng chấm công trông trưa/trông ngủ theo tháng)
   --------------------------------------------------------------------------
   Dữ liệu lưu trong tab "TrongTrua" (tự tạo nếu chưa có) của đúng file năm
   học đang chọn, với 5 cột: Dấu thời gian | Ngày | Họ và tên | Bộ phận | Lớp.
   Mỗi dòng = 1 người trông 1 lớp vào 1 ngày cụ thể.
   ========================================================================== */

// Tìm chỉ số cột "Bộ phận" theo tiêu đề (không hardcode cột), để tự điền khi
// người dùng chọn Họ và tên. Nếu không tìm thấy, trả về -1 (không tự điền).
function timCotBoPhan(headers) {
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || "").toString().trim().toLowerCase();
    if (h.indexOf("bộ phận") !== -1 || h.indexOf("bo phan") !== -1) {
      return i;
    }
  }
  return -1;
}

// 1. Lấy danh sách nhân sự (Họ tên + Bộ phận) từ sheet DsTrongTrua và danh sách lớp
function layDSGVChoTrongTrua(maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheetGV = ss.getSheetByName("DsTrongTrua");
    var dsgv = [];

    if (sheetGV && sheetGV.getLastRow() >= 3) {
      // Lấy từ dòng 3 đến hết, lấy 2 cột (Cột A: Họ tên, Cột B: Bộ phận)
      var lastRow = sheetGV.getLastRow();
      var data = sheetGV.getRange(3, 1, lastRow - 2, 2).getDisplayValues();

      for (var i = 0; i < data.length; i++) {
        var hoTen = data[i][0];
        var boPhan = data[i][1];

        if (!hoTen) continue;
        dsgv.push({
          hoTen: hoTen,
          boPhan: boPhan || ""
        });
      }
    }

    var dsLop = layDanhSachLop(maNamHoc);
    return { dsgv: dsgv, dsLop: dsLop };
  } catch (err) {
    return { dsgv: [], dsLop: [], loi: err.toString() };
  }
}

// 2. Lấy danh sách người đã được xếp trông trưa của 1 NGÀY cụ thể (để hiển thị/sửa)
function layTrongTruaTheoNgay(ngayStr, maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrongTrua");
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var result = [];

    data.forEach(function(row) {
      var ngayFmt = layNgayDangYYYYMMDD(row[1], ss);
      if (ngayFmt === ngayStr) {
        result.push({ hoTen: row[2], boPhan: row[3] || "", lop: row[4] || "" });
      }
    });

    return result;
  } catch (err) {
    return [];
  }
}

// Hàm phụ trợ: chuẩn hóa 1 giá trị ngày (Date hoặc text) về dạng "yyyy-MM-dd"
function layNgayDangYYYYMMDD(giaTri, ss) {
  if (!giaTri) return "";
  if (giaTri instanceof Date) {
    return Utilities.formatDate(giaTri, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  }
  return giaTri.toString().trim();
}

// 2b. Lấy toàn bộ danh sách người đã được xếp trông trưa của 1 THÁNG, gom nhóm
//     theo từng Ngày. CHỈ trả về những ngày đã có người (bỏ qua ngày trống).
function layTrongTruaTheoThang(thangStr, maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrongTrua");
    if (!sheet || sheet.getLastRow() < 2) return [];

    var parts = thangStr.split("-"); // "yyyy-MM"
    var nam = parseInt(parts[0], 10);
    var thang = parseInt(parts[1], 10);

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var mapNgay = {}; // "yyyy-MM-dd" -> [{hoTen, boPhan, lop}]

    data.forEach(function(row) {
      var ngayFmt = layNgayDangYYYYMMDD(row[1], ss); // "yyyy-MM-dd"
      var p = ngayFmt.split("-");
      if (p.length !== 3) return;
      if (parseInt(p[0], 10) !== nam || parseInt(p[1], 10) !== thang) return;
      if (!row[2]) return;

      if (!mapNgay[ngayFmt]) mapNgay[ngayFmt] = [];
      mapNgay[ngayFmt].push({ hoTen: row[2], boPhan: row[3] || "", lop: row[4] || "" });
    });

    // Sắp xếp theo ngày tăng dần, mỗi phần tử = 1 ngày có người trông
    return Object.keys(mapNgay).sort().map(function(ngay) {
      var p = ngay.split("-");
      return {
        ngay: ngay,                                    // "yyyy-MM-dd" (dùng để đối chiếu/gán vào input date)
        ngayHienThi: p[2] + "/" + p[1] + "/" + p[0],    // "dd/MM/yyyy" (hiển thị)
        danhSach: mapNgay[ngay]
      };
    });
  } catch (err) {
    return [];
  }
}

// 3. Lưu (ghi đè) toàn bộ danh sách người trông trưa của 1 NGÀY cụ thể.
//    Xóa hết các dòng cũ thuộc đúng ngày đó rồi ghi lại từ danh sách mới nhất
//    do client gửi lên, để tránh trùng lặp khi người dùng sửa đi sửa lại.
function luuTrongTruaTheoNgay(ngayStr, maNamHoc, danhSach) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrongTrua");

    if (!sheet) {
      sheet = ss.insertSheet("TrongTrua");
      sheet.appendRow(["Dấu thời gian", "Ngày", "Họ và tên", "Bộ phận", "Lớp"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#fff3e0");
    }

    // Đọc toàn bộ dữ liệu hiện có 1 LẦN, giữ lại các dòng không thuộc ngày đang sửa
    // (thay vì deleteRow() từng dòng - rất chậm vì mỗi lệnh là 1 lượt gọi API riêng).
    var lastRow = sheet.getLastRow();
    var dataGiuLai = [];
    if (lastRow >= 2) {
      var duLieuCu = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      dataGiuLai = duLieuCu.filter(function(row) {
        var ngayFmt = layNgayDangYYYYMMDD(row[1], ss);
        return ngayFmt !== ngayStr;
      });
    }

    // Ghi ngày dưới dạng Date (chuẩn múi giờ) để đồng bộ định dạng với các module khác
    var parts = ngayStr.split("-");
    var ngayObj = ngayStr;
    if (parts.length === 3) {
      ngayObj = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
    }

    var dataMoi = (danhSach || [])
      .filter(function(item) { return item && item.hoTen; })
      .map(function(item) {
        return [new Date(), ngayObj, item.hoTen, item.boPhan || "", item.lop || ""];
      });

    var tongDuLieu = dataGiuLai.concat(dataMoi);

    // Ghi đè lại toàn bộ dữ liệu (trừ dòng tiêu đề) bằng 1-2 LỆNH GHI DUY NHẤT
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    if (tongDuLieu.length > 0) {
      sheet.getRange(2, 1, tongDuLieu.length, 5).setValues(tongDuLieu);
    }

    return "Đã lưu danh sách trông trưa ngày " + ngayStr + " thành công!";
  } catch (err) {
    return "Lỗi: " + err.toString();
  }
}

// 3b. Lấy toàn bộ danh sách người đã được xếp trông trưa trong 1 KHOẢNG NGÀY
//     (từ tuNgayStr đến denNgayStr, bao gồm cả 2 đầu mút). Trả về dạng PHẲNG
//     (mỗi phần tử = 1 người trông 1 lớp vào 1 ngày cụ thể, có kèm "ngay"),
//     dùng để nạp vào khung Thêm/Sửa khi người dùng chọn khoảng ngày.
function layTrongTruaTheoKhoangNgay(tuNgayStr, denNgayStr, maNamHoc) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrongTrua");
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var result = [];

    data.forEach(function(row) {
      var ngayFmt = layNgayDangYYYYMMDD(row[1], ss);
      if (!ngayFmt) return;
      if (ngayFmt >= tuNgayStr && ngayFmt <= denNgayStr) {
        if (!row[2]) return;
        result.push({ ngay: ngayFmt, hoTen: row[2], boPhan: row[3] || "", lop: row[4] || "" });
      }
    });

    // Sắp xếp theo Ngày tăng dần, cùng ngày thì theo Họ và tên
    result.sort(function(a, b) {
      if (a.ngay !== b.ngay) return a.ngay < b.ngay ? -1 : 1;
      return (a.hoTen || "").localeCompare(b.hoTen || "");
    });

    return result;
  } catch (err) {
    return [];
  }
}

// 3c. Lưu (ghi đè) toàn bộ danh sách người trông trưa của 1 KHOẢNG NGÀY cụ thể.
//     Xóa hết các dòng cũ có Ngày nằm trong khoảng [tuNgayStr, denNgayStr] rồi
//     ghi lại từ danh sách mới nhất do client gửi lên (mỗi phần tử tự mang
//     theo "ngay" riêng), để tránh trùng lặp khi người dùng sửa đi sửa lại.
function luuTrongTruaTheoKhoangNgay(tuNgayStr, denNgayStr, maNamHoc, danhSach) {
  try {
    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName("TrongTrua");

    if (!sheet) {
      sheet = ss.insertSheet("TrongTrua");
      sheet.appendRow(["Dấu thời gian", "Ngày", "Họ và tên", "Bộ phận", "Lớp"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#fff3e0");
    }

    // Đọc toàn bộ dữ liệu hiện có 1 LẦN DUY NHẤT, giữ lại những dòng có Ngày
    // KHÔNG nằm trong khoảng đang sửa (thay vì deleteRow() từng dòng một - rất chậm
    // vì mỗi deleteRow là 1 lượt gọi API riêng tới Google Sheets).
    var lastRow = sheet.getLastRow();
    var dataGiuLai = [];
    if (lastRow >= 2) {
      var duLieuCu = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      dataGiuLai = duLieuCu.filter(function(row) {
        var ngayFmt = layNgayDangYYYYMMDD(row[1], ss);
        return !(ngayFmt && ngayFmt >= tuNgayStr && ngayFmt <= denNgayStr);
      });
    }

    // Chuẩn bị dữ liệu mới từ danh sách client gửi lên
    var dataMoi = (danhSach || [])
      .filter(function(item) { return item && item.hoTen && item.ngay; })
      .map(function(item) {
        var parts = item.ngay.split("-");
        var ngayObj = item.ngay;
        if (parts.length === 3) {
          ngayObj = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
        }
        return [new Date(), ngayObj, item.hoTen, item.boPhan || "", item.lop || ""];
      });

    var tongDuLieu = dataGiuLai.concat(dataMoi);

    // Ghi đè lại toàn bộ dữ liệu (trừ dòng tiêu đề) bằng 1-2 LỆNH GHI DUY NHẤT
    // thay vì appendRow() từng dòng - giúp lưu nhanh hơn rất nhiều với khoảng ngày lớn.
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    if (tongDuLieu.length > 0) {
      sheet.getRange(2, 1, tongDuLieu.length, 5).setValues(tongDuLieu);
    }

    return "Đã lưu danh sách trông trưa thành công!";
  } catch (err) {
    return "Lỗi: " + err.toString();
  }
}

// 4. Xuất Excel "Bảng chấm công trông trưa" dạng ma trận (người x ngày làm việc)
//    cho cả THÁNG được chọn, đúng theo mẫu bảng chấm công hiện có của trường.
function xuatFileTrongTruaChuan(thangStr, maNamHoc) {
  var ss = moFileNamHoc(maNamHoc);
  var sheet = ss.getSheetByName("TrongTrua");

  // 1. Dọn dẹp sạch sẽ các Sheet tạm cũ còn sót lại (trong đúng file năm học này)
  xoaTatCaSheetTam(maNamHoc);

  var parts = thangStr.split("-"); // "yyyy-MM"
  var nam = parseInt(parts[0], 10);
  var thang = parseInt(parts[1], 10); // 1-12
  var soNgayTrongThang = new Date(nam, thang, 0).getDate();

  // 2. Danh sách các ngày làm việc (Thứ 2 -> Thứ 6) trong tháng đang xuất
  var tenThuMang = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"];
  var danhSachNgayLamViec = [];
  for (var d = 1; d <= soNgayTrongThang; d++) {
    var dateObj = new Date(nam, thang - 1, d);
    var thu = dateObj.getDay();
    if (thu >= 1 && thu <= 5) {
      danhSachNgayLamViec.push({ ngay: d, tenThu: tenThuMang[thu] });
    }
  }

  // 3. Đọc toàn bộ dữ liệu TrongTrua, gom nhóm theo Họ và tên, lọc đúng tháng đang xuất
  var mapNguoi = {};      // hoTen -> { boPhan, theoNgay: { ngay: lop } }
  var thuTuHienThi = [];  // giữ thứ tự xuất hiện lần đầu của mỗi người

  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    data.forEach(function(row) {
      var rawNgay = row[1];
      var dateObj2 = (rawNgay instanceof Date) ? rawNgay : new Date(rawNgay);
      if (isNaN(dateObj2.getTime())) return;
      if (dateObj2.getFullYear() !== nam || (dateObj2.getMonth() + 1) !== thang) return;

      var hoTen = (row[2] || "").toString().trim();
      if (!hoTen) return;
      var boPhan = row[3] || "";
      var lop = row[4] || "";
      var ngaySo = dateObj2.getDate();

      if (!mapNguoi[hoTen]) {
        mapNguoi[hoTen] = { boPhan: boPhan, theoNgay: {} };
        thuTuHienThi.push(hoTen);
      }
      mapNguoi[hoTen].theoNgay[ngaySo] = lop;
    });
  }

  // 4. Tạo Sheet tạm & đặt tên file xuất
  var tenThangHienThi = "T" + thang + "." + nam;
  var baseFileName = "ChamCong_TrongTrua_" + tenThangHienThi;
  var tempSheetName = "TEMP_" + baseFileName;
  var tempSheet = ss.insertSheet(tempSheetName);

  var soCotDuLieu = 3 + danhSachNgayLamViec.length + 2; // TT, Họ tên, Bộ phận, [ngày...], Tổng, Số buổi

  // Dòng 1: Tiêu đề lớn
  var rangeTieuDe = tempSheet.getRange(1, 1, 1, soCotDuLieu);
  rangeTieuDe.merge(); // 1. Gộp các ô dòng 1 lại

  tempSheet.getRange(1, 1).setValue("BẢNG CHẤM CÔNG TRÔNG TRƯA"); // 2. CHỈ gán giá trị vào riêng ô A1

  // 3. Định dạng cho tiêu đề
  rangeTieuDe
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Dòng 2: Tháng
  tempSheet.getRange(2, 2).setValue(tenThangHienThi).setFontWeight("bold");

  // Dòng 3: Tiêu đề cột (Thứ)
  var header3 = ["TT", "Họ và tên", "Bộ phận"];
  danhSachNgayLamViec.forEach(function(nv) { header3.push(nv.tenThu); });
  header3.push("Tổng");
  header3.push("Số buổi");
  tempSheet.getRange(3, 1, 1, header3.length).setValues([header3]);

  // Dòng 4: Số ngày trong tháng tương ứng từng cột Thứ
  var header4 = ["", "", ""];
  danhSachNgayLamViec.forEach(function(nv) { header4.push(nv.ngay); });
  header4.push("");
  header4.push("");
  tempSheet.getRange(4, 1, 1, header4.length).setValues([header4]);

  // 5. Đưa dữ liệu từng người, từ dòng 5
  var rowIdx = 5;
  thuTuHienThi.forEach(function(hoTen, idx) {
    var info = mapNguoi[hoTen];
    var rowArr = [idx + 1, hoTen, info.boPhan];
    var soBuoi = 0;

    danhSachNgayLamViec.forEach(function(nv) {
      var lop = info.theoNgay[nv.ngay] || "";
      if (lop) soBuoi++;
      rowArr.push(lop);
    });

    rowArr.push(soBuoi); // Tổng
    rowArr.push(soBuoi); // Số buổi

    tempSheet.getRange(rowIdx, 1, 1, rowArr.length).setValues([rowArr]);
    rowIdx++;
  });

  // 6. Định dạng khung bảng, độ rộng cột
  var lastRow = tempSheet.getLastRow();
  var tongSoCot = header3.length;

  if (lastRow >= 3) {
    tempSheet.getRange(3, 1, lastRow - 2, tongSoCot)
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    tempSheet.getRange(3, 1, 2, tongSoCot)
      .setFontWeight("bold")
      .setBackground("#fff3e0");

    if (lastRow > 4) {
      tempSheet.getRange(5, 2, lastRow - 4, 1).setHorizontalAlignment("left");
    }

    tempSheet.setColumnWidth(1, 36);
    tempSheet.setColumnWidth(2, 180);
    tempSheet.setColumnWidth(3, 90);
    for (var c = 4; c <= tongSoCot - 2; c++) {
      tempSheet.setColumnWidth(c, 46);
    }
    tempSheet.setColumnWidth(tongSoCot - 1, 55);
    tempSheet.setColumnWidth(tongSoCot, 55);

    tempSheet.setRowHeight(1, 38);
    tempSheet.setRowHeight(2, 24);
    tempSheet.setRowHeight(3, 26);
    tempSheet.setRowHeight(4, 22);
  }

  // 7. Khối ký tên (Ban giám hiệu / Tổ trưởng CSHS)
  var kyRow = lastRow + 3;
  var cotKy1 = 3;
  var cotKy2 = Math.max(cotKy1 + 3, tongSoCot - 3);

  tempSheet.getRange(kyRow, cotKy1).setValue("BAN GIÁM HIỆU").setFontWeight("bold");
  tempSheet.getRange(kyRow, cotKy2).setValue("TỔ TRƯỞNG CSHS").setFontWeight("bold");
  tempSheet.getRange(kyRow + 4, cotKy1).setValue("Lê Thanh Tâm").setFontWeight("bold");
  tempSheet.getRange(kyRow + 4, cotKy2).setValue("Nguyễn Thị Xuân Đào").setFontWeight("bold");

  SpreadsheetApp.flush();

  // 8. Gom các Sheet vào một File tạm thời để xuất Excel chứa nhiều Sheet
  var FILE_NGUON_ID = "1Vlm0tRvP1jEXaXcSvYN9bjTKHDuAcCwfoeEZX41wMA0"; // <-- Thay ID file nguồn vào đây
  var TEN_SHEET_1 = "PCN_lop1";                   // <-- Thay tên Sheet 1 cần chép
  var TEN_SHEET_2 = "PCN_1co2lop";                   // <-- Thay tên Sheet 2 cần chép

  // Tạo file Google Sheet tạm để chứa các sheet xuất
  var tempExportSS = SpreadsheetApp.create(baseFileName);
  var defaultSheet = tempExportSS.getSheets()[0]; // Sheet mặc định ban đầu

  // a. Copy Sheet Trông trưa sang file tạm
  var copiedSheetMain = tempSheet.copyTo(tempExportSS);
  copiedSheetMain.setName("Bảng chấm công"); // Đặt tên hiển thị trong Excel

  // b. Copy 2 Sheet từ file bên ngoài sang
  try {
    var sourceSS = SpreadsheetApp.openById(FILE_NGUON_ID);
    var sheet1 = sourceSS.getSheetByName(TEN_SHEET_1);
    var sheet2 = sourceSS.getSheetByName(TEN_SHEET_2);

    if (sheet1) {
      var c1 = sheet1.copyTo(tempExportSS);
      c1.setName(sheet1.getName());
    }
    if (sheet2) {
      var c2 = sheet2.copyTo(tempExportSS);
      c2.setName(sheet2.getName());
    }
  } catch (e) {
    Logger.log("Lỗi khi mở/copy file nguồn: " + e.toString());
  }

  // c. Xóa sheet trống mặc định lúc tạo file
  tempExportSS.deleteSheet(defaultSheet);

  // d. Tạo link xuất toàn bộ File (không dùng &gid để xuất đủ cả 3 sheets)
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempExportSS.getId() + "/export?format=xlsx";

  // e. Đưa file tạm vào Thùng rác Drive để tránh rác dung lượng (link tải vẫn hoạt động)
  DriveApp.getFileById(tempExportSS.getId()).setTrashed(true);

  return {
    success: true,
    downloadUrl: exportUrl,
    tempSheetName: tempSheetName,
    fileName: baseFileName + '.xlsx',
    maNamHoc: maNamHoc || MA_NAM_HOC_MAC_DINH
  };
}

/* ==========================================================================
   MODULE: ĐỀ NGHỊ VẬT TƯ (DNVT)
   --------------------------------------------------------------------------
   - Có 1 sheet MẪU tên "Template_DNVT" trong file năm học (giống mẫu Phiếu
     đề nghị cung cấp hàng hóa/dịch vụ).
   - Mỗi THÁNG chỉ có DUY NHẤT 1 phiếu: khi lưu lần đầu trong tháng, hệ thống
     COPY sheet Template_DNVT thành 1 sheet mới tên "Tháng 9.26" (Tháng
     <số tháng>.<2 số cuối năm>). Các lần lưu sau trong CÙNG tháng sẽ gộp
     thêm mặt hàng vào đúng sheet đó.
   - Các trường Người đề xuất / Trưởng phòng-bộ phận / Phòng-Bộ phận là CỐ
     ĐỊNH (khai báo bên dưới), không cho sửa trên form. Cô muốn đổi thì sửa
     3 hằng số DNVT_NGUOI_DE_NGHI / DNVT_TRUONG_PHONG / DNVT_PHONG_BO_PHAN.
   - Toàn bộ vị trí ô trên phiếu (Ngày:, Người đề xuất, Trưởng phòng/bộ
     phận, Phòng/Bộ phận, Mục đích sử dụng, Ngày sử dụng, bảng hàng hóa,
     dòng TỔNG) đều được TÌM THEO NHÃN CHỮ trong sheet mẫu, KHÔNG hard-code
     số dòng/cột. Vì vậy Template_DNVT bắt buộc phải còn nguyên các nhãn chữ
     đúng như trong ảnh mẫu cô gửi (không đổi tên nhãn, có thể đổi định
     dạng/màu sắc/độ rộng cột thoải mái).
   ========================================================================== */

var DNVT_TEN_SHEET_TEMPLATE = "Template_DNVT";

// ----- 3 giá trị cố định của phiếu, sửa trực tiếp tại đây nếu cần đổi -----
var DNVT_NGUOI_DE_NGHI = "Hoàng Thị Hà";
var DNVT_TRUONG_PHONG  = "Nguyễn Thị Xuân Đào";
var DNVT_PHONG_BO_PHAN = "Chăm sóc học sinh";

// Định dạng 1 số thành chuỗi tiền Việt Nam đồng, vd 26250000 -> "26.250.000 đ"
// (tự viết tay, không phụ thuộc toLocaleString('vi-VN') để chắc chắn chạy đúng
// trên mọi runtime của Apps Script)
function dvFormatTienVND(so) {
  var n = Math.round(Number(so) || 0);
  var am = n < 0;
  var chuoi = Math.abs(n).toString();
  var ketQua = '';
  for (var i = 0; i < chuoi.length; i++) {
    var viTriTuPhai = chuoi.length - i;
    ketQua += chuoi.charAt(i);
    if (viTriTuPhai > 1 && viTriTuPhai % 3 === 1) ketQua += '.';
  }
  return (am ? '-' : '') + ketQua + ' đ';
}

// Ngược lại: đọc 1 ô có thể đang là số thuần hoặc chuỗi đã format kiểu
// "26.250.000 đ" / "-2.000 đ", trả về số nguyên để tính toán
function dvParseTienVND(giaTri) {
  if (typeof giaTri === 'number') return giaTri;
  var s = (giaTri || '').toString().trim();
  if (!s) return 0;
  var am = s.charAt(0) === '-';
  var chiSo = s.replace(/[^0-9]/g, '');
  var n = parseInt(chiSo, 10);
  if (isNaN(n)) return 0;
  return am ? -n : n;
}

// Tên sheet của 1 tháng, ví dụ nam=2026, thang=9 -> "Tháng 9.26"
function tenSheetThangDNVT(nam, thang) {
  return "Tháng " + thang + "." + String(nam).slice(-2);
}

// Từ 2 số cuối năm lấy trên tên sheet (vd "26"), suy ra năm đầy đủ (vd 2026),
// dựa vào mã năm học (vd "2627" -> năm bắt đầu 2026, năm kết thúc 2027).
// Nhờ vậy so sánh mốc "năm*12+tháng" giữa các sheet trong CÙNG 1 file năm học
// luôn đúng thứ tự thời gian thực tế, kể cả khi năm học vắt qua 2 năm dương lịch.
function namDayDuDNVT(maNamHoc, haiSoCuoiNam) {
  var ma = (maNamHoc || MA_NAM_HOC_MAC_DINH).toString().trim();
  var namBatDau = 2000 + parseInt(ma.substring(0, 2), 10);
  var namKetThuc = 2000 + parseInt(ma.substring(2, 4), 10);
  var target = String(haiSoCuoiNam);
  if (String(namBatDau).slice(-2) === target) return namBatDau;
  if (String(namKetThuc).slice(-2) === target) return namKetThuc;
  return 2000 + parseInt(target, 10); // dự phòng, không nên xảy ra
}

// Tìm ô đầu tiên chứa/bắt đầu bằng "chuoiTim" trong 1 vùng hàng của sheet
// kieuKhop: 'batdau' (mặc định, an toàn hơn vì tránh trùng nhãn con) hoặc 'chua'
function timOTrongVungTheoNhan(sheet, chuoiTim, hangBatDau, hangKetThuc, kieuKhop) {
  var lastCol = sheet.getLastColumn();
  var soHang = hangKetThuc - hangBatDau + 1;
  if (lastCol < 1 || soHang < 1) return null;

  var values = sheet.getRange(hangBatDau, 1, soHang, lastCol).getValues();
  var target = boDauTiengVietGS(chuoiTim).trim();

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cell = boDauTiengVietGS(values[r][c]).trim();
      if (!cell) continue;
      var match = (kieuKhop === 'chua') ? (cell.indexOf(target) !== -1) : (cell.indexOf(target) === 0);
      if (match) return { row: hangBatDau + r, col: c + 1 };
    }
  }
  return null;
}

function timOTheoNhan(sheet, chuoiTim, kieuKhop) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;
  return timOTrongVungTheoNhan(sheet, chuoiTim, 1, lastRow, kieuKhop);
}

// Ô để GHI GIÁ TRỊ ngay sau 1 ô NHÃN (tự động bỏ qua vùng ô gộp của nhãn, nếu có)
function oGiaTriKeBenNhanDNVT(sheet, hang, cotNhan) {
  var oNhan = sheet.getRange(hang, cotNhan);
  var merged = oNhan.getMergedRanges();
  var cotBatDau = (merged && merged.length > 0) ? merged[0].getLastColumn() : cotNhan;
  return sheet.getRange(hang, cotBatDau + 1);
}

// Điền các trường cố định của phiếu ngay khi tạo sheet mới cho 1 tháng
function dienThongTinCoDinhDNVT(sheet, nam, thang) {
  var vt;

  vt = timOTheoNhan(sheet, "Ngày đề nghị:", 'batdau');
  if (vt) oGiaTriKeBenNhanDNVT(sheet, vt.row, vt.col)
    .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

  vt = timOTheoNhan(sheet, "Người đề xuất", 'batdau');
  if (vt) oGiaTriKeBenNhanDNVT(sheet, vt.row, vt.col).setValue(DNVT_NGUOI_DE_NGHI);

  vt = timOTheoNhan(sheet, "Trưởng phòng", 'batdau');
  if (vt) oGiaTriKeBenNhanDNVT(sheet, vt.row, vt.col).setValue(DNVT_TRUONG_PHONG);

  // "batdau": tránh trùng với nhãn "Trưởng phòng/bộ phận" ở dòng trên
  vt = timOTheoNhan(sheet, "Phòng/Bộ phận", 'batdau');
  if (vt) oGiaTriKeBenNhanDNVT(sheet, vt.row, vt.col).setValue(DNVT_PHONG_BO_PHAN);

  vt = timOTheoNhan(sheet, "Ngày sử dụng", 'batdau');
  if (vt) {
    var thangPad = (thang < 10 ? "0" : "") + thang;
    oGiaTriKeBenNhanDNVT(sheet, vt.row, vt.col).setValue("Tháng " + thangPad + "/" + nam);
  }
}

// Lấy sheet của đúng tháng; nếu chưa có thì copy từ Template_DNVT và điền sẵn
// các trường cố định
function layHoacTaoSheetDNVT(maNamHoc, nam, thang) {
  var ss = moFileNamHoc(maNamHoc);
  var tenSheetThang = tenSheetThangDNVT(nam, thang);
  var sheet = ss.getSheetByName(tenSheetThang);

  if (!sheet) {
    var template = ss.getSheetByName(DNVT_TEN_SHEET_TEMPLATE);
    if (!template) {
      throw new Error("Không tìm thấy sheet mẫu '" + DNVT_TEN_SHEET_TEMPLATE + "' trong file năm học. Vui lòng tạo sheet mẫu này trước.");
    }
    sheet = template.copyTo(ss);
    sheet.setName(tenSheetThang);

    // Đưa sheet mới ra ngay sau Template để danh sách tab gọn gàng
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(template.getIndex() + 1);

    dienThongTinCoDinhDNVT(sheet, nam, thang);
  }
  return sheet;
}

// Xác định cấu trúc bảng "Danh mục hàng yêu cầu" trong 1 sheet tháng
// (dòng tiêu đề, dòng TỔNG, các cột TT/Tên mặt hàng/ĐVT/SL/Đơn giá/Thành tiền/Ghi chú)
function layCauTrucBangDNVT(sheet) {
  var viTriHeader = timOTheoNhan(sheet, "TÊN MẶT HÀNG", 'chua');
  if (!viTriHeader) {
    throw new Error("Không tìm thấy tiêu đề cột 'TÊN MẶT HÀNG' trong sheet '" + sheet.getName() + "'. Kiểm tra lại cấu trúc Template_DNVT.");
  }

  var hangHeader = viTriHeader.row;
  var colTenMH = viTriHeader.col;
  var ct = {
    hangHeader: hangHeader,
    colTT: colTenMH - 1,
    colTenMH: colTenMH,
    colDVT: colTenMH + 1,
    colSL: colTenMH + 2,
    colDonGia: colTenMH + 3,
    colThanhTien: colTenMH + 4,
    colGhiChu: colTenMH + 5
  };

  // Đảm bảo sheet đủ cột tới cột Ghi chú (phòng khi mẫu chưa mở rộng đủ)
  if (sheet.getMaxColumns() < ct.colGhiChu) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), ct.colGhiChu - sheet.getMaxColumns());
  }

  var lastRow = sheet.getLastRow();
  var viTriTong = timOTrongVungTheoNhan(sheet, "TỔNG", hangHeader + 1, lastRow, 'batdau');
  if (!viTriTong) {
    throw new Error("Không tìm thấy dòng 'TỔNG' phía dưới bảng hàng hóa trong sheet '" + sheet.getName() + "'. Kiểm tra lại cấu trúc Template_DNVT.");
  }
  ct.hangTong = viTriTong.row;

  return ct;
}

// Tính tổng số tiền đã chi LŨY KẾ từ đầu file năm học cho tới hết tháng đang xem
// (bao gồm cả tháng đang xem), bằng cách cộng dòng TỔNG của mọi sheet tháng có
// mốc thời gian (năm*12+tháng) <= mốc thời gian tháng đang xem, trong CÙNG 1 file.
function tinhSoTienDaChiDNVT(maNamHoc, namChon, thangChon) {
  var ss = moFileNamHoc(maNamHoc);
  var sheets = ss.getSheets();
  var mocChon = namChon * 12 + thangChon;
  var re = /^Tháng\s+(\d{1,2})\.(\d{2})$/;
  var tong = 0;

  sheets.forEach(function(sh) {
    var m = re.exec(sh.getName());
    if (!m) return; // bỏ qua Template_DNVT và các sheet khác không đúng định dạng "Tháng X.YY"

    var thangSheet = parseInt(m[1], 10);
    var namSheet = namDayDuDNVT(maNamHoc, m[2]);
    var mocSheet = namSheet * 12 + thangSheet;
    if (mocSheet > mocChon) return; // chỉ tính các tháng <= tháng đang xem

    try {
      var ct = layCauTrucBangDNVT(sh);
      var giaTri = sh.getRange(ct.hangTong, ct.colThanhTien).getValue();
      tong += Number(giaTri) || 0;
    } catch (e) {
      // sheet lỗi cấu trúc (thiếu nhãn...) thì bỏ qua, không làm hỏng cả phép tính
    }
  });

  return tong;
}

// Lấy (và ghi luôn vào sheet, nếu sheet tháng đã tồn tại) 3 giá trị ngân sách:
// - budget: Chi phí ngân sách dự kiến (cố định, đọc từ sheet tháng nếu có, hoặc
//   từ Template_DNVT nếu tháng đó chưa tạo phiếu)
// - daChi: tổng đã chi lũy kế tính tới hết tháng đang xem (đã bao gồm tháng này)
// - conLai: budget - daChi
// Ghi giá trị Ngân sách dự trù (budget) MỚI cho CẢ NĂM HỌC. Lưu duy nhất 1 nơi
// (sheet mẫu Template_DNVT) làm nguồn chân lý (Single Source of Truth) - mỗi
// khi 1 tháng bất kỳ được xem/lưu, giá trị này sẽ tự đồng bộ xuống sheet tháng đó.
function luuNganSachDuTruDNVT(maNamHoc, giaTriMoi) {
  try {
    var giaTri = Number(giaTriMoi);
    if (isNaN(giaTri) || giaTri < 0) {
      return { success: false, message: "Giá trị ngân sách không hợp lệ!" };
    }

    var ss = moFileNamHoc(maNamHoc);
    var template = ss.getSheetByName(DNVT_TEN_SHEET_TEMPLATE);
    if (!template) {
      return { success: false, message: "Không tìm thấy sheet mẫu '" + DNVT_TEN_SHEET_TEMPLATE + "'!" };
    }

    var vtBudget = timOTheoNhan(template, "Chi phí ngân sách dự kiến", 'chua');
    if (!vtBudget) {
      return { success: false, message: "Không tìm thấy nhãn 'Chi phí ngân sách dự kiến' trong sheet mẫu!" };
    }
    oGiaTriKeBenNhanDNVT(template, vtBudget.row, vtBudget.col).setValue(dvFormatTienVND(giaTri));

    return { success: true, message: "Đã cập nhật ngân sách dự trù!", budget: giaTri };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

// Ghi chú: Ngân sách dự trù / Đã chi / Còn lại được lưu vào ô LIỀN KỀ ngay sau
// ô nhãn tương ứng (dùng oGiaTriKeBenNhanDNVT, tự xử lý đúng ô kể cả khi nhãn
// bị merge nhiều cột) - đúng cấu trúc bảng trong ảnh mẫu cô gửi. Giá trị được
// ghi dưới dạng chuỗi đã format tiền Việt Nam đồng (vd "26.250.000 đ"), và khi
// đọc lại để tính toán sẽ parse ngược lại thành số bằng dvParseTienVND.
// Ngân sách dự trù luôn được đọc từ Template_DNVT (nguồn chân lý duy nhất, xem
// hàm luuNganSachDuTruDNVT phía trên); mỗi lần gọi hàm này, giá trị đó cũng
// được đồng bộ (ghi đè) xuống sheet của tháng đang xem để khi in/xuất file vẫn
// đúng số liệu mới nhất.
function layNganSachDNVT(maNamHoc, nam, thang) {
  var ss = moFileNamHoc(maNamHoc);
  var tenSheetThang = tenSheetThangDNVT(nam, thang);
  var sheet = ss.getSheetByName(tenSheetThang);
  var template = ss.getSheetByName(DNVT_TEN_SHEET_TEMPLATE);

  var budget = 0;
  if (template) {
    var vtBudget = timOTheoNhan(template, "Chi phí ngân sách dự kiến", 'chua');
    if (vtBudget) {
      budget = dvParseTienVND(oGiaTriKeBenNhanDNVT(template, vtBudget.row, vtBudget.col).getValue());
    }
  }

  var daChi = tinhSoTienDaChiDNVT(maNamHoc, nam, thang);
  var conLai = budget - daChi;

  // Nếu phiếu tháng này đã tồn tại thì đồng bộ cả 3 giá trị (đã format) vào sheet
  if (sheet) {
    var vtBudgetThang = timOTheoNhan(sheet, "Chi phí ngân sách dự kiến", 'chua');
    if (vtBudgetThang) oGiaTriKeBenNhanDNVT(sheet, vtBudgetThang.row, vtBudgetThang.col).setValue(dvFormatTienVND(budget));

    var vtDaChi = timOTheoNhan(sheet, "Số tiền đã chi", 'chua');
    if (vtDaChi) oGiaTriKeBenNhanDNVT(sheet, vtDaChi.row, vtDaChi.col).setValue(dvFormatTienVND(daChi));

    var vtConLai = timOTheoNhan(sheet, "Ngân sách còn lại", 'chua');
    if (vtConLai) oGiaTriKeBenNhanDNVT(sheet, vtConLai.row, vtConLai.col).setValue(dvFormatTienVND(conLai));
  }

  return { budget: budget, daChi: daChi, conLai: conLai };
}

// Đánh lại STT liên tục cho các dòng mặt hàng đang có nội dung + cập nhật công thức TỔNG
function capNhatSttVaTongDNVT(sheet, ct) {
  var tt = 1;
  for (var r = ct.hangHeader + 1; r < ct.hangTong; r++) {
    var ten = sheet.getRange(r, ct.colTenMH).getValue();
    if (ten && ten.toString().trim() !== "") {
      sheet.getRange(r, ct.colTT).setValue(tt);
      tt++;
    } else {
      sheet.getRange(r, ct.colTT).clearContent();
    }
  }

  var colChuTong = sheet.getRange(1, ct.colThanhTien).getA1Notation().replace(/[0-9]/g, '');
  if (ct.hangTong - 1 >= ct.hangHeader + 1) {
    sheet.getRange(ct.hangTong, ct.colThanhTien)
      .setFormula("=SUM(" + colChuTong + (ct.hangHeader + 1) + ":" + colChuTong + (ct.hangTong - 1) + ")");
  } else {
    sheet.getRange(ct.hangTong, ct.colThanhTien).setValue(0);
  }
}

// Thêm danh sách mặt hàng mới vào đúng sheet tháng, tự chèn dòng nếu không đủ
// dòng trống có sẵn giữa dòng tiêu đề và dòng TỔNG
function themMatHangDNVT(sheet, ct, danhSachMoi) {
  var hangTrong = [];
  for (var r = ct.hangHeader + 1; r < ct.hangTong; r++) {
    var giaTri = sheet.getRange(r, ct.colTenMH).getValue();
    if (!giaTri || giaTri.toString().trim() === "") hangTrong.push(r);
  }

  var soDongTao = danhSachMoi.length - hangTrong.length;

  if (soDongTao > 0) {
    var hangMau = ct.hangTong - 1; // dòng ngay trên dòng TỔNG, để sao chép định dạng/viền
    var soCotMau = ct.colGhiChu - ct.colTT + 1;

    sheet.insertRowsBefore(ct.hangTong, soDongTao);

    var rangeMau = sheet.getRange(hangMau, ct.colTT, 1, soCotMau);
    for (var i = 0; i < soDongTao; i++) {
      var hangMoi = ct.hangTong + i; // các dòng vừa được chèn, nằm trước dòng TỔNG (đã bị đẩy xuống)
      rangeMau.copyTo(sheet.getRange(hangMoi, ct.colTT, 1, soCotMau));
      hangTrong.push(hangMoi);
    }
    ct.hangTong += soDongTao;
  }

  hangTrong.sort(function(a, b) { return a - b; });

  danhSachMoi.forEach(function(item, idx) {
    var hang = hangTrong[idx];
    var soLuong = Number(item.soLuong) || 0;
    var donGia = Number(item.donGia) || 0;
    sheet.getRange(hang, ct.colTenMH).setValue(item.tenMatHang || "");
    sheet.getRange(hang, ct.colDVT).setValue(item.dvt || "");
    sheet.getRange(hang, ct.colSL).setValue(soLuong);
    sheet.getRange(hang, ct.colDonGia).setValue(donGia);
    sheet.getRange(hang, ct.colThanhTien).setValue(soLuong * donGia);
    sheet.getRange(hang, ct.colGhiChu).setValue(item.ghiChu || "");
  });

  capNhatSttVaTongDNVT(sheet, ct);
}

// 1. LƯU (gộp thêm) danh sách mặt hàng vào phiếu của đúng tháng
//    data = { maNamHoc, thangSuDung: "yyyy-MM", danhSachMatHang: [{tenMatHang, dvt, soLuong, donGia, ghiChu}] }
function luuDeNghiVatTu(data) {
  try {
    if (!data || !data.thangSuDung) {
      return { success: false, message: "Vui lòng chọn Tháng sử dụng!" };
    }
    if (!data.danhSachMatHang || data.danhSachMatHang.length === 0) {
      return { success: false, message: "Vui lòng thêm ít nhất 1 mặt hàng trước khi lưu!" };
    }

    var parts = data.thangSuDung.split("-"); // "yyyy-MM"
    var nam = parseInt(parts[0], 10);
    var thang = parseInt(parts[1], 10);

    var sheet = layHoacTaoSheetDNVT(data.maNamHoc, nam, thang);
    var ct = layCauTrucBangDNVT(sheet);
    themMatHangDNVT(sheet, ct, data.danhSachMatHang);
    var nganSach = layNganSachDNVT(data.maNamHoc, nam, thang);

    return {
      success: true,
      message: "Đã lưu đề nghị vật tư tháng " + thang + "/" + nam + " thành công!",
      tenSheet: sheet.getName(),
      nganSach: nganSach
    };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

// 2. Lấy danh sách mặt hàng đã lưu của 1 tháng (để hiển thị khi mở lại/chọn tháng)
function layDeNghiVatTuTheoThang(maNamHoc, thangSuDung) {
  try {
    var parts = (thangSuDung || "").split("-");
    if (parts.length !== 2) return { danhSach: [], tong: 0, daTonTai: false };

    var nam = parseInt(parts[0], 10);
    var thang = parseInt(parts[1], 10);

    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName(tenSheetThangDNVT(nam, thang));
    if (!sheet) {
      // Chưa có phiếu tháng này -> vẫn tính/hiển thị được ngân sách (đọc budget
      // từ Template_DNVT, đã chi tính từ các tháng trước đó đã tồn tại)
      var nganSachChuaCoPhieu = layNganSachDNVT(maNamHoc, nam, thang);
      return { danhSach: [], tong: 0, daTonTai: false, nganSach: nganSachChuaCoPhieu };
    }

    var ct = layCauTrucBangDNVT(sheet);
    var danhSach = [];

    for (var r = ct.hangHeader + 1; r < ct.hangTong; r++) {
      var ten = sheet.getRange(r, ct.colTenMH).getValue();
      if (!ten || ten.toString().trim() === "") continue;
      danhSach.push({
        tenMatHang: ten,
        dvt: sheet.getRange(r, ct.colDVT).getValue(),
        soLuong: sheet.getRange(r, ct.colSL).getValue(),
        donGia: sheet.getRange(r, ct.colDonGia).getValue(),
        thanhTien: sheet.getRange(r, ct.colThanhTien).getValue(),
        ghiChu: sheet.getRange(r, ct.colGhiChu).getValue()
      });
    }

    var tongTien = sheet.getRange(ct.hangTong, ct.colThanhTien).getValue();
    var nganSach = layNganSachDNVT(maNamHoc, nam, thang);
    return { danhSach: danhSach, tong: tongTien || 0, daTonTai: true, nganSach: nganSach };
  } catch (err) {
    return { danhSach: [], tong: 0, daTonTai: false, loi: err.toString() };
  }
}

// 3. Xóa 1 mặt hàng đã lưu (viTriDong = chỉ số trong mảng trả về bởi layDeNghiVatTuTheoThang)
function xoaMatHangDNVT(maNamHoc, thangSuDung, viTriDong) {
  try {
    var parts = (thangSuDung || "").split("-");
    var nam = parseInt(parts[0], 10);
    var thang = parseInt(parts[1], 10);

    var ss = moFileNamHoc(maNamHoc);
    var sheet = ss.getSheetByName(tenSheetThangDNVT(nam, thang));
    if (!sheet) return { success: false, message: "Chưa có phiếu đề nghị của tháng này!" };

    var ct = layCauTrucBangDNVT(sheet);
    var danhSachHang = [];
    for (var r = ct.hangHeader + 1; r < ct.hangTong; r++) {
      var ten = sheet.getRange(r, ct.colTenMH).getValue();
      if (ten && ten.toString().trim() !== "") danhSachHang.push(r);
    }

    var hangCanXoa = danhSachHang[viTriDong];
    if (!hangCanXoa) return { success: false, message: "Không tìm thấy mặt hàng cần xóa!" };

    // Xóa HẲN dòng trên sheet (thay vì chỉ xóa nội dung) để các dòng phía dưới
    // tự động dịch lên, không để lại dòng trống xen giữa bảng.
    sheet.deleteRow(hangCanXoa);
    ct.hangTong -= 1; // dòng TỔNG cũng dịch lên 1 theo dòng vừa xóa

    capNhatSttVaTongDNVT(sheet, ct);
    var nganSach = layNganSachDNVT(maNamHoc, nam, thang);

    return { success: true, message: "Đã xóa mặt hàng khỏi phiếu!", nganSach: nganSach };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

// 4. Xuất Excel/PDF trực tiếp sheet của tháng đang chọn (không cần tạo sheet tạm
//    vì bản thân sheet tháng đã là phiếu hoàn chỉnh, đúng định dạng in ấn)
function xuatFileDNVT(thangSuDung, maNamHoc, dinhDang) {
  try {
    var parts = (thangSuDung || "").split("-");
    var nam = parseInt(parts[0], 10);
    var thang = parseInt(parts[1], 10);

    var ss = moFileNamHoc(maNamHoc);
    var tenSheetThang = tenSheetThangDNVT(nam, thang);
    var sheet = ss.getSheetByName(tenSheetThang);
    if (!sheet) {
      return { success: false, message: "Chưa có phiếu đề nghị vật tư của tháng " + thang + "/" + nam + "!" };
    }

    var ssId = ss.getId();
    var sheetId = sheet.getSheetId();
    var exportUrl;

    if (dinhDang === 'pdf') {
      exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&gid=" + sheetId +
        "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false";
    } else {
      exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx&gid=" + sheetId;
    }

    return {
      success: true,
      downloadUrl: exportUrl,
      fileName: tenSheetThang + (dinhDang === 'pdf' ? '.pdf' : '.xlsx')
    };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}
