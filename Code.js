var SPREADSHEET_ID = "1gWIntcQ2k_5mEjRfxrtZqoq6H63HBUZHfOvPPK4gn8E";
function doGet() {
  return HtmlService.createTemplateFromFile('Main').evaluate()
      .setTitle('HỆ THỐNG TRẮC NGHIỆM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getData(sheetName) { try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName(sheetName); return sheet ? sheet.getDataRange().getValues() : []; } catch (e) { return []; } }
function getTopicName(id) { var data = getData("ChuDe"); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() == String(id).trim()) return data[i][1]; } return "Đề thi"; }

// --- AUTH & USER ---
function kiemTraDangNhap(user, pass) {
  var data = getData("TaiKhoan"); var inputUser = String(user).trim().toLowerCase(); var inputPass = String(pass).trim();
  for (var i = 1; i < data.length; i++) { 
    var rowUser = String(data[i][0]).trim().toLowerCase(); var rowPass = String(data[i][1]).trim(); 
    if (rowUser === "" || rowPass === "") continue; 
    if (rowUser == inputUser && rowPass == inputPass) { 
      var role = data[i][3] ? String(data[i][3]).trim() : "Student"; 
      var allowChat = data[i][4] && String(data[i][4]).trim().toLowerCase() === "yes"; 
      try { capNhatThoiGianOnline(rowUser); } catch(e){} 
      return { success: true, msg: data[i][2], role: role, allowChat: allowChat }; 
    } 
  }
  return { success: false, msg: "Sai tài khoản hoặc mật khẩu!" };
}

// --- LOGIC SHOP & TIỀN TỆ ---
function getTuVi(userID) {
  var data = getData("DiemSo"); var total = 0; 
  var uid = String(userID).trim().toLowerCase(); 
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === uid) {
      var diem = parseFloat(data[i][5]); var tongCau = parseFloat(data[i][6]);
      if (tongCau > 0) total += (diem / tongCau) * 10;
    }
  }
  return total;
}

function layThongTinVatPham(userID) {
  var data = getData("VatPham"); var targetID = String(userID).trim().toLowerCase();
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var inventory = {}; var lastSpin = ""; var spentPoints = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetID) {
      lastSpin = data[i][1] ? Utilities.formatDate(new Date(data[i][1]), "GMT+7", "yyyy-MM-dd") : "";
      inventory = data[i][2] ? JSON.parse(data[i][2]) : {};
      spentPoints = data[i][3] ? parseFloat(data[i][3]) : 0;
      break;
    }
  }
  var totalTuVi = getTuVi(userID);
  var balance = Math.floor(totalTuVi - spentPoints);
  return { canSpin: (lastSpin !== today), inventory: inventory, balance: balance };
}

function muaVatPham(userID, itemCode, price) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("VatPham");
    var data = sheet.getDataRange().getValues(); var targetID = String(userID).trim().toLowerCase();
    var rowIndex = -1; var currentInv = {}; var spentPoints = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetID) {
        rowIndex = i + 1; currentInv = data[i][2] ? JSON.parse(data[i][2]) : {}; spentPoints = data[i][3] ? parseFloat(data[i][3]) : 0; break;
      }
    }
    var totalTuVi = getTuVi(userID); var balance = totalTuVi - spentPoints;
    if (balance < price) return { success: false, msg: "Không đủ TU VI! Hãy làm thêm bài tập." };
    
    spentPoints += price;
    if (!currentInv[itemCode]) currentInv[itemCode] = 0; currentInv[itemCode]++;
    
    if (rowIndex === -1) { sheet.appendRow([targetID, "", JSON.stringify(currentInv), spentPoints]); } 
    else { sheet.getRange(rowIndex, 3).setValue(JSON.stringify(currentInv)); sheet.getRange(rowIndex, 4).setValue(spentPoints); }
    
    return { success: true, msg: "Mua thành công!", inventory: currentInv, newBalance: Math.floor(totalTuVi - spentPoints) };
  } catch(e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); }
}

function tangVatPham(senderID, receiverID, itemCode, amount) {
  if (String(senderID).toLowerCase() === String(receiverID).toLowerCase()) return { success: false, msg: "Không thể tự tặng!" };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("VatPham"); var data = sheet.getDataRange().getValues();
    var sIdx = -1, rIdx = -1; var sData = {}, rData = {}; var sID = String(senderID).trim().toLowerCase(); var rID = String(receiverID).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var rowID = String(data[i][0]).trim().toLowerCase();
      if (rowID === sID) { sIdx = i + 1; sData = data[i][2] ? JSON.parse(data[i][2]) : {}; }
      if (rowID === rID) { rIdx = i + 1; rData = data[i][2] ? JSON.parse(data[i][2]) : {}; }
    }
    if (sIdx === -1) return { success: false, msg: "Bạn chưa có túi đồ!" };
    if (rIdx === -1) return { success: false, msg: "Người nhận chưa kích hoạt túi đồ!" };
    if (!sData[itemCode] || sData[itemCode] < amount) return { success: false, msg: "Không đủ vật phẩm!" };
    sData[itemCode] -= amount;
    if (!rData[itemCode]) rData[itemCode] = 0; rData[itemCode] += parseInt(amount);
    sheet.getRange(sIdx, 3).setValue(JSON.stringify(sData));
    sheet.getRange(rIdx, 3).setValue(JSON.stringify(rData));
    return { success: true, msg: "Đã gửi quà!", inventory: sData };
  } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); }
}

// --- QUAY THƯỞNG ---
var RATE_TABLE = [
  { code: "none", name: "Chúc may mắn lần sau", type: "text", val: 0, rate: 30 },
  { code: "help_5050", name: "Trợ giúp 50/50", type: "item", val: 1, rate: 15 },
  { code: "check_true", name: "Soi Đáp Án", type: "item", val: 1, rate: 10 },
  { code: "reveal_wrong", name: "Loại 1 câu sai", type: "item", val: 1, rate: 15 },
  { code: "time_5", name: "Thêm 5 phút", type: "item", val: 5, rate: 15 },
  { code: "time_10", name: "Thêm 10 phút", type: "item", val: 10, rate: 10 },
  { code: "time_15", name: "Thêm 15 phút", type: "item", val: 15, rate: 4 },
  { code: "anti_cheat", name: "Xóa 1 lần thoát Tab", type: "item", val: 1, rate: 1 }
];

function thucHienQuayThuong(userID) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("VatPham");
    if (!sheet) { sheet = ss.insertSheet("VatPham"); sheet.appendRow(["UserID", "LastSpinDate", "InventoryJSON", "SpentPoints"]); }
    var data = sheet.getDataRange().getValues(); var targetID = String(userID).trim().toLowerCase(); var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"); var rowIndex = -1; var currentInv = {};
    for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetID) { rowIndex = i + 1; var dbDate = data[i][1] ? Utilities.formatDate(new Date(data[i][1]), "GMT+7", "yyyy-MM-dd") : ""; if (dbDate === today) return { success: false, msg: "Hôm nay bạn đã quay rồi!" }; currentInv = data[i][2] ? JSON.parse(data[i][2]) : {}; break; } }
    var rand = Math.random() * 100; var accumulated = 0; var selectedItem = RATE_TABLE[0];
    for (var k = 0; k < RATE_TABLE.length; k++) { accumulated += RATE_TABLE[k].rate; if (rand <= accumulated) { selectedItem = RATE_TABLE[k]; break; } }
    if (selectedItem.code !== "none") { if (!currentInv[selectedItem.code]) currentInv[selectedItem.code] = 0; currentInv[selectedItem.code]++; }
    if (rowIndex === -1) { sheet.appendRow([targetID, today, JSON.stringify(currentInv), 0]); } else { sheet.getRange(rowIndex, 2).setValue(today); sheet.getRange(rowIndex, 3).setValue(JSON.stringify(currentInv)); }
    return { success: true, item: selectedItem, inventory: currentInv };
  } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); }
}

function normalizeTextCell(value) { return String(value == null ? "" : value).replace(/^'/, "").trim(); }
function getQuestionRowData(row, rowIndex) {
  if (!row) return null;
  var topicCol0 = normalizeTextCell(row[0]);
  var topicCol1 = normalizeTextCell(row[1]);
  var isNewFormat = topicCol1 !== "" && normalizeTextCell(row[2]) !== "" && /^[A-E]$/i.test(normalizeTextCell(row[8]));
  if (isNewFormat) {
    var optionsMap = {
      A: normalizeTextCell(row[3]), B: normalizeTextCell(row[4]), C: normalizeTextCell(row[5]),
      D: normalizeTextCell(row[6]), E: normalizeTextCell(row[7])
    };
    var correctLetter = normalizeTextCell(row[8]).toUpperCase();
    return {
      rowIndex: rowIndex, format: "new", id: normalizeTextCell(row[0]) || ("Q_" + rowIndex), maChuDe: topicCol1,
      question: normalizeTextCell(row[2]), options: [optionsMap.A, optionsMap.B, optionsMap.C, optionsMap.D, optionsMap.E],
      correctLetter: correctLetter, correctText: optionsMap[correctLetter] || "", explanation: normalizeTextCell(row[9])
    };
  }
  if (topicCol0 !== "" && normalizeTextCell(row[1]) !== "") {
    return {
      rowIndex: rowIndex, format: "old", id: "ROW_" + rowIndex, maChuDe: topicCol0, question: normalizeTextCell(row[1]),
      options: [normalizeTextCell(row[2]), normalizeTextCell(row[3]), normalizeTextCell(row[4]), normalizeTextCell(row[5]), ""],
      correctLetter: "", correctText: normalizeTextCell(row[6]), explanation: normalizeTextCell(row[7])
    };
  }
  return null;
}

function suDungVatPham(userID, itemCode, currentQuestionID) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("VatPham"); var data = sheet.getDataRange().getValues(); var targetID = String(userID).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetID) {
        var currentInv = data[i][2] ? JSON.parse(data[i][2]) : {};
        if (currentInv[itemCode] && currentInv[itemCode] > 0) {
          currentInv[itemCode]--; sheet.getRange(i + 1, 3).setValue(JSON.stringify(currentInv));
          var hintData = null;
          if ((itemCode === "help_5050" || itemCode === "check_true" || itemCode === "reveal_wrong") && currentQuestionID !== undefined) {
             var qData = getData("CauHoi"); var row = qData[currentQuestionID]; var questionData = getQuestionRowData(row, currentQuestionID);
             if (questionData && questionData.correctText) hintData = { ans: questionData.correctText };
          }
          return { success: true, msg: "Đã sử dụng!", remaining: currentInv[itemCode], hint: hintData };
        } else { return { success: false, msg: "Hết vật phẩm!" }; }
      }
    }
    return { success: false, msg: "Lỗi tìm túi đồ!" };
  } catch(e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); }
}

// --- CÁC HÀM CŨ ---
function capNhatThoiGianOnline(userId) { var lock = LockService.getScriptLock(); try { if (lock.tryLock(3000)) { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("TaiKhoan"); var data = sheet.getDataRange().getValues(); var targetID = String(userId).trim().toLowerCase(); if (sheet.getLastColumn() < 8) sheet.insertColumnsAfter(sheet.getLastColumn(), 8 - sheet.getLastColumn()); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetID) { sheet.getRange(i + 1, 8).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); return; } } } } catch (e) {} finally { lock.releaseLock(); } }
function layDanhSachTrangThaiOnline(reqID, reqPass) { if (!checkRole(reqID, reqPass, "admin")) return []; var data = getData("TaiKhoan"); var usersStatus = []; var now = new Date(); for (var i = 1; i < data.length; i++) { if (data[i][0]) { var lastActiveStr = data[i][7]; var status = "⚪ Offline"; var timeInfo = "Lâu quá rồi"; if (lastActiveStr) { var lastTime = new Date(lastActiveStr); if (!isNaN(lastTime.getTime())) { var diffMs = now - lastTime; var diffMins = Math.floor(diffMs / 60000); if (diffMins <= 5) { status = "🟢 Online"; timeInfo = "Vừa xong"; } else if (diffMins <= 60) { status = "🟡 Vừa thoát"; timeInfo = diffMins + " phút trước"; } else { status = "⚪ Offline"; timeInfo = Math.floor(diffMins/60) + " giờ trước"; } } } usersStatus.push({ id: data[i][0], name: data[i][2], role: data[i][3], status: status, time: timeInfo }); } } return usersStatus; }
function layDanhSachMetadata(userID) { var mon = getData("MonHoc"); var khoi = getData("Khoi"); var listMon = []; for(var i=1; i<mon.length; i++) if(mon[i][0]) listMon.push({ma: mon[i][0], ten: mon[i][1]}); var listKhoi = []; for(var j=1; j<khoi.length; j++) if(khoi[j][0]) listKhoi.push({ma: khoi[j][0], ten: khoi[j][1]}); if (!userID) return { mon: listMon, khoi: listKhoi }; var dataTK = getData("TaiKhoan"); var userPermissions = { mon: "ALL", khoi: "ALL", role: "Student" }; var targetID = String(userID).trim().toLowerCase(); for (var k = 1; k < dataTK.length; k++) { if (String(dataTK[k][0]).trim().toLowerCase() === targetID) { userPermissions.role = dataTK[k][3]; userPermissions.mon = dataTK[k][5] ? String(dataTK[k][5]).trim() : "ALL"; userPermissions.khoi = dataTK[k][6] ? String(dataTK[k][6]).trim() : "ALL"; break; } } if (String(userPermissions.role).toLowerCase() === 'admin') return { mon: listMon, khoi: listKhoi }; if (userPermissions.mon !== "ALL" && userPermissions.mon !== "") { var allowedMon = userPermissions.mon.split(',').map(function(s){ return s.trim().toLowerCase(); }); listMon = listMon.filter(function(m) { return allowedMon.indexOf(String(m.ma).toLowerCase()) !== -1; }); } if (userPermissions.khoi !== "ALL" && userPermissions.khoi !== "") { var allowedKhoi = userPermissions.khoi.split(',').map(function(s){ return s.trim().toLowerCase(); }); listKhoi = listKhoi.filter(function(k) { return allowedKhoi.indexOf(String(k.ma).toLowerCase()) !== -1; }); } return { mon: listMon, khoi: listKhoi }; }
function checkRole(uid, upass, roleCheck) { if (!uid || !upass) return false; var data = getData("TaiKhoan"); var targetID = String(uid).trim().toLowerCase(); var targetPass = String(upass).trim(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetID && String(data[i][1]).trim() === targetPass) { var role = data[i][3] ? String(data[i][3]).trim().toLowerCase() : ""; if (roleCheck === "admin") return role === "admin"; if (roleCheck === "teacher") return (role === "teacher" || role === "admin"); } } return false; }
function layDanhSachTaiKhoan(reqID, reqPass) { if (!checkRole(reqID, reqPass, "admin")) throw new Error("⛔ Không có quyền Admin!"); var data = getData("TaiKhoan"); var users = []; for (var i = 1; i < data.length; i++) { if (data[i][0]) users.push({ id: data[i][0], pass: "******", name: data[i][2], role: data[i][3], chat: data[i][4], mon: data[i][5] || "", khoi: data[i][6] || "" }); } return users; }
function luuTaiKhoan(userObj, reqID, reqPass) { if (!checkRole(reqID, reqPass, "admin")) return { success: false, msg: "Không có quyền Admin!" }; var lock = LockService.getScriptLock(); try { lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("TaiKhoan"); var data = sheet.getDataRange().getValues(); var targetID = String(userObj.id).trim().toLowerCase(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetID) { sheet.getRange(i + 1, 2, 1, 6).setValues([[userObj.pass, userObj.name, userObj.role, userObj.chat, userObj.mon, userObj.khoi]]); return { success: true, msg: "Đã cập nhật tài khoản!" }; } } sheet.appendRow([userObj.id, userObj.pass, userObj.name, userObj.role, userObj.chat, userObj.mon, userObj.khoi]); return { success: true, msg: "Đã thêm tài khoản mới!" }; } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); } }
function xoaTaiKhoan(userId, reqID, reqPass) { if (!checkRole(reqID, reqPass, "admin")) return { success: false, msg: "Không có quyền Admin!" }; if (String(userId).toLowerCase() === String(reqID).toLowerCase()) return { success: false, msg: "Không thể tự xóa!" }; var lock = LockService.getScriptLock(); try { lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("TaiKhoan"); var data = sheet.getDataRange().getValues(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === String(userId).trim().toLowerCase()) { sheet.deleteRow(i + 1); return { success: true, msg: "Đã xóa tài khoản!" }; } } return { success: false, msg: "Không tìm thấy ID!" }; } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); } }
function guiTinNhanTraoDoi(user, noiDung) { try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("TraoDoi"); var time = Utilities.formatDate(new Date(), "GMT+7", "HH:mm dd/MM"); sheet.appendRow([user, noiDung, time]); return { success: true }; } catch (e) { return { success: false }; } }
function layNoiDungTraoDoi() { try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheetTK = ss.getSheetByName("TaiKhoan"); var dataTK = sheetTK.getDataRange().getValues(); var userMap = {}; for (var i = 1; i < dataTK.length; i++) userMap[String(dataTK[i][0]).trim().toLowerCase()] = { name: dataTK[i][2], role: dataTK[i][3] }; var sheetChat = ss.getSheetByName("TraoDoi"); var dataChat = sheetChat.getDataRange().getValues(); var chats = []; var start = Math.max(1, dataChat.length - 50); for (var j = start; j < dataChat.length; j++) { var senderID = String(dataChat[j][0]).trim(); if (senderID) { var userInfo = userMap[senderID.toLowerCase()] || { name: senderID, role: "Student" }; chats.push({ user: userInfo.name, id: senderID, role: userInfo.role, text: dataChat[j][1], time: dataChat[j][2] }); } } return chats; } catch (e) { return []; } }
function layDanhSachChuDe(userID) { var data = getData("ChuDe"); var userRole = "Student"; var allowedMon = ["ALL"]; var allowedKhoi = ["ALL"]; if (userID) { var dataTK = getData("TaiKhoan"); var targetID = String(userID).trim().toLowerCase(); for (var k = 1; k < dataTK.length; k++) { if (String(dataTK[k][0]).trim().toLowerCase() === targetID) { userRole = dataTK[k][3] ? String(dataTK[k][3]).trim() : "Student"; var strMon = dataTK[k][5] ? String(dataTK[k][5]).trim() : ""; if (strMon !== "" && strMon !== "ALL") allowedMon = strMon.split(',').map(function(s){ return s.trim().toLowerCase(); }); var strKhoi = dataTK[k][6] ? String(dataTK[k][6]).trim() : ""; if (strKhoi !== "" && strKhoi !== "ALL") allowedKhoi = strKhoi.split(',').map(function(s){ return s.trim().toLowerCase(); }); break; } } } if (userRole.toLowerCase() === 'admin') { allowedMon = ["ALL"]; allowedKhoi = ["ALL"]; } var output = []; for (var i = 1; i < data.length; i++) { if (data[i][0] && String(data[i][0]).trim() !== "") { var maDe = String(data[i][0]).trim(); var monDe = data[i][5] ? String(data[i][5]).trim().toLowerCase() : ""; var khoiDe = data[i][6] ? String(data[i][6]).trim().toLowerCase() : ""; var checkMon = (allowedMon[0] === "ALL" || monDe === "" || allowedMon.indexOf(monDe) !== -1); var checkKhoi = (allowedKhoi[0] === "ALL" || khoiDe === "" || allowedKhoi.indexOf(khoiDe) !== -1); if (checkMon && checkKhoi) { output.push({ maChuDe: maDe, tenChuDe: data[i][1], matKhau: String(data[i][2]).trim(), thoiGian: data[i][3], moTa: data[i][4], maMon: data[i][5] ? String(data[i][5]).trim() : "", maKhoi: data[i][6] ? String(data[i][6]).trim() : "", timeStart: data[i][7] ? String(data[i][7]) : "", timeEnd: data[i][8] ? String(data[i][8]) : "", viewMode: data[i][9] ? String(data[i][9]).trim() : "detail", soCau: data[i][10] ? parseInt(data[i][10]) : 0 }); } } } return output; }
function themChuDeMoi(obj, reqID, reqPass) { if (!checkRole(reqID, reqPass, "teacher")) return { success: false, msg: "⛔ TỪ CHỐI!" }; var lock = LockService.getScriptLock(); try { lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("ChuDe"); var data = sheet.getDataRange().getValues(); for (var i = 1; i < data.length; i++) if (String(data[i][0]).trim().toLowerCase() == String(obj.ma).trim().toLowerCase()) return { success: false, msg: "Mã chủ đề này đã tồn tại!" }; sheet.appendRow(["'" + obj.ma, obj.ten, obj.pass, obj.time, obj.mota, obj.mon, obj.khoi, "'" + obj.timeStart, "'" + obj.timeEnd, obj.viewMode, obj.soCau]); SpreadsheetApp.flush(); return { success: true, msg: "Đã thêm chủ đề thành công!" }; } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); } }
function suaChuDe(obj, reqID, reqPass) { if (!checkRole(reqID, reqPass, "teacher")) return { success: false, msg: "⛔ TỪ CHỐI!" }; var lock = LockService.getScriptLock(); try { lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("ChuDe"); var data = sheet.getDataRange().getValues(); var targetID = String(obj.ma).trim().toLowerCase(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetID) { var rowData = [obj.ten, obj.pass, obj.time, obj.mota, obj.mon, obj.khoi, "'" + obj.timeStart, "'" + obj.timeEnd, obj.viewMode, obj.soCau]; sheet.getRange(i + 1, 2, 1, 10).setValues([rowData]); return { success: true, msg: "Đã cập nhật chủ đề!" }; } } return { success: false, msg: "Không tìm thấy mã chủ đề để sửa!" }; } catch (e) { return { success: false, msg: "Lỗi: " + e.toString() }; } finally { lock.releaseLock(); } }
function luuCauHoiTuLatex(maChuDe, listCauHoi, reqID, reqPass) { if (!checkRole(reqID, reqPass, "teacher")) return { success: false, msg: "⛔ TỪ CHỐI!" }; try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("CauHoi"); var dataToAppend = []; for (var i = 0; i < listCauHoi.length; i++) { var item = listCauHoi[i]; var opts = item.options || []; dataToAppend.push(["'" + maChuDe, item.question, opts[0]||"", opts[1]||"", opts[2]||"", opts[3]||"", item.correct, item.explain]); } if (dataToAppend.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, 8).setValues(dataToAppend); return { success: true, msg: "Đã thêm thành công " + dataToAppend.length + " câu hỏi!" }; } catch (e) { return { success: false, msg: "Lỗi Server: " + e.toString() }; } }
function luuCauHoiTrucTiep(maChuDe, listCauHoi, reqID, reqPass) {
  if (!checkRole(reqID, reqPass, "teacher")) return { success: false, msg: "⛔ TỪ CHỐI!" };
  var lock = LockService.getScriptLock();
  try {
    if (!listCauHoi || !listCauHoi.length) return { success: false, msg: "Danh sách câu hỏi trống!" };
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("CauHoi");
    if (sheet.getLastColumn() < 11) sheet.insertColumnsAfter(sheet.getLastColumn(), 11 - sheet.getLastColumn());
    var targetTopic = normalizeTextCell(maChuDe);
    var existing = sheet.getDataRange().getValues();
    var seq = 0;
    for (var i = 1; i < existing.length; i++) {
      var rowData = getQuestionRowData(existing[i], i);
      if (rowData && String(rowData.maChuDe).toLowerCase() === targetTopic.toLowerCase()) {
        seq++;
        var currentId = normalizeTextCell(rowData.id);
        var m = currentId.match(/_(\d+)$/);
        if (m) seq = Math.max(seq, parseInt(m[1], 10));
      }
    }
    var dataToAppend = [];
    for (var j = 0; j < listCauHoi.length; j++) {
      var item = listCauHoi[j] || {};
      var question = normalizeTextCell(item.question);
      var options = item.options || {};
      var optA = normalizeTextCell(options.A), optB = normalizeTextCell(options.B), optC = normalizeTextCell(options.C), optD = normalizeTextCell(options.D), optE = normalizeTextCell(options.E);
      var correct = normalizeTextCell(item.correctAnswer).toUpperCase();
      if (!question) continue;
      if (!/^[A-E]$/.test(correct)) return { success: false, msg: "Có câu hỏi chưa chọn đáp án đúng hợp lệ (A-E)." };
      var optionMap = { A: optA, B: optB, C: optC, D: optD, E: optE };
      if (!optionMap[correct]) return { success: false, msg: "Đáp án đúng không có nội dung. Vui lòng kiểm tra lại." };
      seq++;
      var autoId = targetTopic + "_" + ("000" + seq).slice(-3);
      var qID = normalizeTextCell(item.id) || autoId;
      dataToAppend.push(["'" + qID, "'" + targetTopic, question, optA, optB, optC, optD, optE, correct, normalizeTextCell(item.explanation), item.createdAt || new Date().toISOString()]);
    }
    if (!dataToAppend.length) return { success: false, msg: "Không có câu hỏi hợp lệ để lưu!" };
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, 11).setValues(dataToAppend);
    SpreadsheetApp.flush();
    return { success: true, msg: "Đã lưu thành công " + dataToAppend.length + " câu hỏi!" };
  } catch (e) {
    return { success: false, msg: "Lỗi Server: " + e.toString() };
  } finally { lock.releaseLock(); }
}
function layDanhSachCauHoi(maChuDeCanLay) { var data = getData("CauHoi"); var questions = []; var targetID = normalizeTextCell(maChuDeCanLay); var topics = getData("ChuDe"); var limit = 0; for(var t=1; t<topics.length; t++) { if(normalizeTextCell(topics[t][0]) === targetID) { limit = topics[t][10] ? parseInt(topics[t][10]) : 0; break; } } for (var i = 1; i < data.length; i++) { var rowData = getQuestionRowData(data[i], i); if (rowData && normalizeTextCell(rowData.maChuDe) === targetID) { var listOptions = rowData.options.filter(function(o) { return normalizeTextCell(o) !== ""; }); if (rowData.question && listOptions.length > 0) { for (var k = listOptions.length - 1; k > 0; k--) { var j = Math.floor(Math.random() * (k + 1)); var temp = listOptions[k]; listOptions[k] = listOptions[j]; listOptions[j] = temp; } questions.push({ id: i, question: rowData.question, options: listOptions }); } } } for (var i = questions.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var temp = questions[i]; questions[i] = questions[j]; questions[j] = temp; } if (limit > 0 && limit < questions.length) { questions = questions.slice(0, limit); } return questions; }

// --- CHẤM ĐIỂM ---
function xuLyChamDiem(baiLamUser, listID, maChuDe, userID, startTimeStr, violationCount) { 
  var topics = layDanhSachChuDe(); var currentTopic = topics.find(t => t.maChuDe == maChuDe); 
  if (currentTopic && currentTopic.timeEnd) { if (new Date() > new Date(new Date(currentTopic.timeEnd).getTime() + 60000)) return { success: false, msg: "Rất tiếc! Đã quá hạn nộp bài." }; } 
  var data = getData("CauHoi"); var results = []; var diem = 0; var logBaiLam = []; 
  for(var k=0; k < baiLamUser.length; k++) { var qID = listID[k]; var row = data[qID]; if(!row) continue; var questionData = getQuestionRowData(row, qID); if(!questionData) continue; var correctAns = normalizeTextCell(questionData.correctText); var isCorrect = (String(baiLamUser[k]).trim().toLowerCase() === correctAns.toLowerCase()); if(isCorrect) diem++; results.push({ isCorrect: isCorrect, dapAnDung: correctAns, giaiThich: questionData.explanation ? String(questionData.explanation) : "" }); logBaiLam.push({ q: questionData.question, a: baiLamUser[k], c: correctAns, ok: isCorrect }); } 
  var viewMode = currentTopic.viewMode || "detail"; if (viewMode === "hidden") results = []; else if (viewMode === "score_only") results.forEach(r => { r.dapAnDung = ""; r.giaiThich = ""; }); 
  var now = new Date(); var start = new Date(startTimeStr); var diffMs = Math.max(0, now - start); var durationStr = Math.floor(diffMs / 60000) + " phút " + Math.floor((diffMs % 60000) / 1000) + " giây"; 
  var lock = LockService.getScriptLock(); 
  try { lock.waitLock(30000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("DiemSo"); var tenChuDe = getTopicName(maChuDe); 
  var allData = sheet.getDataRange().getValues(); var lanThu = 1; for (var r = 1; r < allData.length; r++) if (String(allData[r][1]) == userID && String(allData[r][2]) == maChuDe) lanThu++; 
  sheet.appendRow([Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss"), userID, maChuDe, lanThu, tenChuDe, diem, listID.length, Utilities.formatDate(start, "GMT+7", "dd/MM/yyyy HH:mm:ss"), durationStr, violationCount || 0, JSON.stringify(logBaiLam)]); SpreadsheetApp.flush(); } catch(e) { return { success: false, msg: "Lỗi hệ thống: " + e.toString() }; } finally { lock.releaseLock(); } 
  if (viewMode === "hidden") return { success: true, viewMode: "hidden", msg: "Đã nộp bài thành công!" }; 
  return { tongDiem: diem, tongCau: listID.length, chiTiet: results, success: true, viewMode: viewMode }; 
}

function layLichSuLamBai(userName) { var data = getData("DiemSo"); var history = []; var targetUser = String(userName).trim().toLowerCase(); for (var i = data.length - 1; i >= 1; i--) { if (String(data[i][1]).trim().toLowerCase() == targetUser) history.push({ thoiGianNop: Utilities.formatDate(new Date(data[i][0]), "GMT+7", "HH:mm dd/MM"), maChuDe: String(data[i][2]).trim(), lanThu: data[i][3], tenChuDe: data[i][4], diem: data[i][5], tongCau: data[i][6], thoiGianBatDau: data[i][7] ? String(data[i][7]) : "", thoiGianLam: data[i][8] ? String(data[i][8]) : "Không xác định" }); } return history; };
function xoaChuDe(maChuDe, reqID, reqPass) { if (!checkRole(reqID, reqPass, "teacher")) return { success: false, msg: "⛔ TỪ CHỐI!" }; try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var targetID = normalizeTextCell(maChuDe).toLowerCase(); var sheetChuDe = ss.getSheetByName("ChuDe"); var dataChuDe = sheetChuDe.getDataRange().getValues(); var isDeleted = false; for (var i = dataChuDe.length - 1; i >= 1; i--) { if (normalizeTextCell(dataChuDe[i][0]).toLowerCase() == targetID) { sheetChuDe.deleteRow(i + 1); isDeleted = true; } } if (!isDeleted) return { success: false, msg: "Không tìm thấy mã chủ đề!" }; var sheetCauHoi = ss.getSheetByName("CauHoi"); var dataCauHoi = sheetCauHoi.getDataRange().getValues(); var countQ = 0; for (var j = dataCauHoi.length - 1; j >= 1; j--) { var rowTopicOld = normalizeTextCell(dataCauHoi[j][0]).toLowerCase(); var rowTopicNew = normalizeTextCell(dataCauHoi[j][1]).toLowerCase(); if (rowTopicOld == targetID || rowTopicNew == targetID) { sheetCauHoi.deleteRow(j + 1); countQ++; } } return { success: true, msg: "Đã xóa chủ đề và " + countQ + " câu hỏi liên quan!" }; } catch (e) { return { success: false, msg: "Lỗi Server: " + e.toString() }; } }
function layThongKe(maChuDe) {
  try {
    var data = getData("DiemSo");
    var targetID = String(maChuDe).trim();

    // --- BƯỚC 1: TẠO MAP TÊN (ID -> Tên thật) ---
    var accData = getData("TaiKhoan");
    var nameMap = {};
    if (accData && accData.length > 0) {
      for (var k = 1; k < accData.length; k++) {
        var accUser = String(accData[k][0]).trim().toLowerCase();
        if (accUser) nameMap[accUser] = accData[k][2];
      }
    }

    var scores = [];
    var details = [];
    var spectrum = new Array(11).fill(0);

    // --- BƯỚC 2: DUYỆT BẢNG ĐIỂM ---
    // Duyệt từ dưới lên để lấy bài mới nhất trước (nếu cần)
    for (var i = data.length - 1; i >= 1; i--) {
      // Kiểm tra đúng mã chủ đề (So sánh tương đối để tránh lỗi dư khoảng trắng)
      if (String(data[i][2]).trim() == targetID) {
        
        var diem = parseFloat(data[i][5]);
        var tongCau = parseFloat(data[i][6]);
        var thangDiem10 = (tongCau > 0) ? (diem / tongCau) * 10 : 0;
        
        scores.push(thangDiem10);

        // Biểu đồ phổ điểm
        var roundedScore = Math.round(thangDiem10);
        if (roundedScore >= 0 && roundedScore <= 10) spectrum[roundedScore]++;

        // Lấy ID gốc và Tên hiển thị
        var userID = String(data[i][1]).trim().toLowerCase();
        var displayName = nameMap[userID] ? nameMap[userID] : data[i][1];

        details.push({
          ngay: Utilities.formatDate(new Date(data[i][0]), "GMT+7", "HH:mm dd/MM"),
          uid: userID,        // <--- QUAN TRỌNG: ID gốc dùng để truy vấn chi tiết
          ten: displayName,   // Tên dùng để hiển thị lên bảng
          lan: data[i][3],
          diem: diem,
          tongCau: tongCau,
          diem10: thangDiem10.toFixed(2),
          tgLam: data[i][8] ? data[i][8] : "",
          viPham: data[i][9] ? data[i][9] : 0
        });
      }
    }

    if (scores.length === 0) return { success: false, msg: "Chưa có dữ liệu bài làm nào!" };

    // --- BƯỚC 3: TÍNH TOÁN THỐNG KÊ (AVG, MEDIAN, MODE...) ---
    scores.sort(function(a, b) { return a - b; });
    var min = scores[0];
    var max = scores[scores.length - 1];
    var sum = scores.reduce((a, b) => a + b, 0);
    var avg = sum / scores.length;
    
    // Tính Median
    var mid = Math.floor(scores.length / 2);
    var median = scores.length % 2 !== 0 ? scores[mid] : (scores[mid - 1] + scores[mid]) / 2;
    
    // Tính Mode
    var frequency = {};
    var maxFreq = 0;
    var modes = [];
    scores.forEach(function(val) {
      var key = val.toFixed(1);
      frequency[key] = (frequency[key] || 0) + 1;
      if (frequency[key] > maxFreq) maxFreq = frequency[key];
    });
    for (var key in frequency) if (frequency[key] === maxFreq) modes.push(key);

    // Phân loại
    var counts = { gioi: 0, kha: 0, tb: 0, yeu: 0 };
    scores.forEach(s => {
      if (s >= 8) counts.gioi++;
      else if (s >= 6.5) counts.kha++;
      else if (s >= 5) counts.tb++;
      else counts.yeu++;
    });

    return {
      success: true,
      total: scores.length,
      avg: avg.toFixed(2),
      median: median.toFixed(2),
      mode: modes.join(", "),
      min: min.toFixed(2),
      max: max.toFixed(2),
      dist: counts,
      spectrum: spectrum,
      list: details // Danh sách này giờ đã có field 'uid'
    };
  } catch (e) {
    return { success: false, msg: "Lỗi Server: " + e.toString() };
  }
}
function layChiTietBaiLamHocSinh(maChuDe, userID, lanThu) {
  var data = getData("DiemSo");
  
  // Chuẩn hóa dữ liệu đầu vào để tìm kiếm chính xác
  var targetUser = String(userID).trim().toLowerCase();
  var targetMa = String(maChuDe).trim().toLowerCase();
  var targetLan = String(lanThu).trim();

  for (var i = data.length - 1; i >= 1; i--) {
    // Lấy dữ liệu từ dòng hiện tại
    var rowUser = String(data[i][1]).trim().toLowerCase();
    var rowMa = String(data[i][2]).trim().toLowerCase();
    var rowLan = String(data[i][3]).trim();

    // So sánh: User ID + Mã Đề + Lần Thi
    if (rowMa === targetMa && rowUser === targetUser && rowLan === targetLan) {
      
      // Cột 10 (tức cột K trong Excel) là nơi lưu JSON chi tiết
      var jsonRaw = data[i][10]; 

      if (!jsonRaw || String(jsonRaw).trim() === "") {
        return { success: false, msg: "Không tìm thấy dữ liệu chi tiết (Log bài làm bị trống)." };
      }

      try {
        var parsedData = JSON.parse(jsonRaw);
        return { success: true, data: parsedData };
      } catch (e) {
        return { success: false, msg: "Dữ liệu bài làm bị lỗi định dạng (JSON Error)." };
      }
    }
  }
  return { success: false, msg: "Không tìm thấy bài thi khớp với thông tin yêu cầu." };
}
function doiMatKhau(user, matKhauCu, matKhauMoi) { var lock = LockService.getScriptLock(); try { lock.waitLock(10000); var ss = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = ss.getSheetByName("TaiKhoan"); var data = sheet.getDataRange().getValues(); var targetUser = String(user).trim().toLowerCase(); for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === targetUser) { if (String(data[i][1]).trim() !== String(matKhauCu).trim()) return { success: false, msg: "Mật khẩu cũ sai!" }; sheet.getRange(i + 1, 2).setValue(matKhauMoi); SpreadsheetApp.flush(); return { success: true, msg: "Đổi mật khẩu thành công!" }; } } return { success: false, msg: "Lỗi tài khoản!" }; } catch (e) { return { success: false, msg: "Lỗi hệ thống: " + e.toString() }; } finally { lock.releaseLock(); } };
function xuatPDF(maChuDe) { try { var stats = layThongKe(maChuDe); if (!stats.success) return { success: false, msg: "Không có dữ liệu!" }; var tenChuDe = getTopicName(maChuDe); var list = stats.list; var html = `<html><head><style>body{font-family:'Times New Roman',serif;padding:20px}h2{text-align:center;color:#2563eb}.info{text-align:center;margin-bottom:20px;color:#555;font-style:italic}table{width:100%;border-collapse:collapse;margin-top:10px;font-size:12px}th{background-color:#f3f4f6;padding:5px;border:1px solid #999;font-weight:bold}td{padding:5px;border:1px solid #999;text-align:center}.text-left{text-align:left}.pass{color:green;font-weight:bold}.fail{color:red;font-weight:bold}.warn{color:red;font-weight:bold}</style></head><body><h2>KẾT QUẢ THI TRẮC NGHIỆM</h2><div class="info">Chủ đề: <strong>${tenChuDe}</strong><br>Ngày xuất báo cáo: ${Utilities.formatDate(new Date(), "GMT+7", "HH:mm dd/MM/yyyy")}</div><table><thead><tr><th style="width:5%">STT</th><th style="width:25%">Họ và tên</th><th style="width:10%">Ngày nộp</th><th style="width:5%">Lần</th><th style="width:15%">Thời gian</th><th style="width:10%">Số câu</th><th style="width:10%">Điểm (10)</th><th style="width:10%;color:#d97706">Thoát Tab</th></tr></thead><tbody>`; for (var i = 0; i < list.length; i++) { var item = list[i]; var colorClass = (item.diem10 >= 5) ? 'pass' : 'fail'; var warnClass = (item.viPham > 0) ? 'warn' : ''; html += `<tr><td>${i + 1}</td><td class="text-left">${item.ten}</td><td>${item.ngay}</td><td>${item.lan}</td><td>${item.tgLam}</td><td>${item.diem}/${item.tongCau}</td><td class="${colorClass}">${item.diem10}</td><td class="${warnClass}">${item.viPham}</td></tr>`; } html += `</tbody></table><div style="margin-top:20px;text-align:right;font-size:11px"><em>Hệ thống trắc nghiệm tự động</em></div></body></html>`; var blob = Utilities.newBlob(html, MimeType.HTML, "temp.html"); var tempFile = DriveApp.createFile(blob); var pdfBlob = tempFile.getAs(MimeType.PDF); var base64 = Utilities.base64Encode(pdfBlob.getBytes()); tempFile.setTrashed(true); return { success: true, data: base64, filename: "Bao_cao_" + maChuDe + ".pdf" }; } catch (e) { return { success: false, msg: "Lỗi xuất PDF: " + e.toString() }; } };
function tinhCanhGioi(tongDiem) { var BASE_EXP = 240; var dsDaiCanhGioi = ["Luyện Khí", "Trúc Cơ", "Kết Đan", "Nguyên Anh", "Hóa Thần", "Luyện Hư", "Hợp Thể", "Đại Thừa", "Độ Kiếp"]; var dsTieuCanh = ["Sơ Kỳ", "Trung Kỳ", "Hậu Kỳ", "Viên Mãn"]; if (tongDiem <= 0) return { name: "Phàm Nhân", class: "rank-0", percent: 0 }; var currentThreshold = 0; var currentExpPool = BASE_EXP; for (var i = 0; i < dsDaiCanhGioi.length; i++) { var unit = currentExpPool / 15; var msSoKy = unit * 1; var msTrungKy= unit * 3; var msHauKy = unit * 7; var msVienMan= unit * 15; var diemTrongCap = tongDiem - currentThreshold; if (diemTrongCap < msVienMan) { var subIndex = 0; if (diemTrongCap < msSoKy) subIndex = 0; else if (diemTrongCap < msTrungKy) subIndex = 1; else if (diemTrongCap < msHauKy) subIndex = 2; else subIndex = 3; return { name: dsDaiCanhGioi[i] + " " + dsTieuCanh[subIndex], class: "rank-" + (i + 1), percent: ((diemTrongCap / msVienMan) * 100).toFixed(1) }; } currentThreshold += currentExpPool; currentExpPool *= 5; } return { name: "Phi Thăng Tiên Giới", class: "rank-9-4", percent: 100 }; }

// --- BXH: MAPPING ID RA TÊN & ĐIỂM THỰC TẾ ---
function layDuLieuBangXepHang() { 
  try { 
    var dataDiem = getData("DiemSo"); 
    var dataTK = getData("TaiKhoan"); 
    var dataVatPham = getData("VatPham"); 

    var idToName = {}; 
    var idToRole = {}; 
    for (var i = 1; i < dataTK.length; i++) { 
        var uid = String(dataTK[i][0]).trim().toLowerCase();
        if (uid) {
            idToName[uid] = dataTK[i][2]; 
            idToRole[uid] = dataTK[i][3] || "Student"; 
        }
    } 

    var idToSpent = {};
    for (var k = 1; k < dataVatPham.length; k++) {
        var uidVP = String(dataVatPham[k][0]).trim().toLowerCase();
        var spent = dataVatPham[k][3] ? parseFloat(dataVatPham[k][3]) : 0; 
        idToSpent[uidVP] = spent;
    }
    
    var stats = {}; 
    for (var j = 1; j < dataDiem.length; j++) { 
        var row = dataDiem[j]; 
        var uID = String(row[1]).trim().toLowerCase();
        var diem = parseFloat(row[5]); 
        var tongCau = parseFloat(row[6]); 
        
        if (!uID || tongCau == 0) continue; 
        
        var diem10 = (diem / tongCau) * 10; 
        
        if (!stats[uID]) {
            var displayName = idToName[uID] || uID;
            stats[uID] = { name: displayName, totalEarned: 0, count: 0, role: idToRole[uID] || "Khách" }; 
        }
        stats[uID].totalEarned += diem10; 
        stats[uID].count++; 
    } 
    
    var ranking = []; 
    for (var key in stats) { 
        var totalEarned = stats[key].totalEarned;
        var spent = idToSpent[key] || 0;
        var currentTuVi = Math.max(0, totalEarned - spent); 

        var realmInfo = tinhCanhGioi(currentTuVi); 
        
        ranking.push({ 
            name: stats[key].name, 
            role: stats[key].role, 
            examCount: stats[key].count, 
            totalScore: currentTuVi.toFixed(1), 
            realm: realmInfo.name, 
            rankClass: realmInfo.class, 
            percent: realmInfo.percent 
        }); 
    } 
    
    ranking.sort(function(a, b) { return b.totalScore - a.totalScore; }); 
    
    return ranking.slice(0, 50); 
  } catch (e) { return []; } 
}

function heartBeatSystem(userName) { if (!userName) return []; var cache = CacheService.getScriptCache(); var keyName = String(userName).trim(); cache.put("ONLINE_" + keyName, "1", 45); var bxh = layDuLieuBangXepHang(); var messages = []; var limit = Math.min(bxh.length, 3); for (var i = 0; i < limit; i++) { var vip = bxh[i]; var vipName = String(vip.name).trim(); var isOnline = cache.get("ONLINE_" + vipName); if (isOnline) { var rankTitle = "TOP " + (i + 1); var icon = (i===0) ? "👑" : ((i===1) ? "🥈" : "🥉"); messages.push(`${icon} ${rankTitle} - ${vipName} đang online!`); } } return messages; };
