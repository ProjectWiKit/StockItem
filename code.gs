function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var stockLogSheet = ss.getSheetByName("StockTransactionLogs") || ss.insertSheet("StockTransactionLogs");
  if (stockLogSheet.getLastRow() === 0) {
    stockLogSheet.appendRow(["Time", "User", "SiteID", "StockID", "ItemName", "Qty", "Unit", "ReturnType"]);
  }

  var stockLogs = [];
  if (stockLogSheet.getLastRow() > 1) {
    var slRows = stockLogSheet.getDataRange().getValues();
    for (var sl = 1; sl < slRows.length; sl++) {
      if (slRows[sl][0]) {
        stockLogs.push({
          time: slRows[sl][0],
          user: slRows[sl][1],
          siteId: slRows[sl][2] ? slRows[sl][2].toString().trim() : "",
          stockId: slRows[sl][3] ? slRows[sl][3].toString().trim() : "",
          itemName: slRows[sl][4] ? slRows[sl][4].toString().trim() : "",
          qty: parseInt(slRows[sl][5]) || 0,
          unit: slRows[sl][6] ? slRows[sl][6].toString().trim() : "ชิ้น",
          returnType: slRows[sl][7] ? slRows[sl][7].toString().trim() : ""
        });
      }
    }
  }

  var stockSheet = ss.getSheetByName("Stock") || ss.insertSheet("Stock");
  if (stockSheet.getLastRow() === 0) {
    stockSheet.appendRow(["ID", "Category", "Name", "Qty", "Unit"]);
  }
  
  var stockMap = {};
  var sRows = stockSheet.getDataRange().getValues();
  var categoryGroups = {};

  for (var j = 1; j < sRows.length; j++) {
    if (sRows[j][0]) {
      var sId = sRows[j][0].toString().trim();
      var cat = sRows[j][1] ? sRows[j][1].toString().trim() : "อุปกรณ์ทั่วไป";
      var name = sRows[j][2] ? sRows[j][2].toString().trim() : "";
      var qty = parseInt(sRows[j][3]) || 0;
      var unit = sRows[j][4] ? sRows[j][4].toString().trim() : "ชิ้น";

      stockMap[sId.toUpperCase()] = {
        id: sId,
        category: cat,
        name: name,
        qty: qty,
        unit: unit
      };

      if (!categoryGroups[cat]) {
        categoryGroups[cat] = [];
      }
      categoryGroups[cat].push([sId, cat, name, qty, unit]);
    }
  }

  var serialSheet = ss.getSheetByName("Serials") || ss.insertSheet("Serials");
  if (serialSheet.getLastRow() === 0) {
    serialSheet.appendRow(["StockID", "SerialNo", "Status", "Qty", "Unit"]);
  }
  var serialMap = {};
  var srRows = serialSheet.getDataRange().getValues();
  for (var k = 1; k < srRows.length; k++) {
    if (srRows[k][0]) {
      var stId = srRows[k][0].toString().trim().toUpperCase();
      var sNo = srRows[k][1] ? srRows[k][1].toString().trim() : "";
      var status = srRows[k][2] ? srRows[k][2].toString().trim() : "พร้อมใช้งาน";
      var sQty = parseInt(srRows[k][3]) || 1;
      var sUnit = srRows[k][4] ? srRows[k][4].toString().trim() : "ชิ้น";
      
      if (status === "พร้อมใช้งาน") {
        if (!serialMap[stId]) {
          serialMap[stId] = [];
        }
        serialMap[stId].push({
          serialNo: sNo,
          qty: sQty,
          unit: sUnit
        });
      }
    }
  }

  for (var catName in categoryGroups) {
    var subSheet = ss.getSheetByName(catName) || ss.insertSheet(catName);
    subSheet.clear();
    subSheet.appendRow(["ID", "Category", "Name", "Qty", "Unit"]);
    var itemsList = categoryGroups[catName];
    if (itemsList.length > 0) {
      subSheet.getRange(2, 1, itemsList.length, itemsList[0].length).setValues(itemsList);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ 
    stockLogs: stockLogs,
    stock: stockMap,
    serials: serialMap
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", message: "ระบบกำลังประมวลผลข้อมูล กรุณาลองใหม่อีกครั้ง"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "login") {
      var inputUser = (data.username || "").toString().trim();
      var inputPass = (data.password || "").toString().trim();

      var adminSheet = ss.getSheetByName("Admin");
      if (adminSheet && adminSheet.getLastRow() > 1) {
        var aRows = adminSheet.getDataRange().getValues();
        for (var i = 1; i < aRows.length; i++) {
          if (aRows[i][0].toString().trim() === inputUser && aRows[i][1].toString().trim() === inputPass) {
            return ContentService.createTextOutput(JSON.stringify({
              status: "success", role: "admin", name: aRows[i][2] ? aRows[i][2].toString().trim() : "Admin", username: inputUser
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var memberSheet = ss.getSheetByName("Members");
      if (memberSheet && memberSheet.getLastRow() > 1) {
        var mRows = memberSheet.getDataRange().getValues();
        for (var m = 1; m < mRows.length; m++) {
          if (mRows[m][0].toString().trim() === inputUser && mRows[m][2].toString().trim() === inputPass) {
            return ContentService.createTextOutput(JSON.stringify({
              status: "success", role: "technician", name: mRows[m][1].toString().trim(), username: inputUser
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "error", message: "รหัสผู้ใช้หรือรหัสผ่านไม่ถูกต้อง"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "register") {
      var memberSheet = ss.getSheetByName("Members") || ss.insertSheet("Members");
      if (memberSheet.getLastRow() === 0) {
        memberSheet.appendRow(["ID", "Name", "Password", "Phone", "Status"]);
      }
      memberSheet.appendRow([data.memberId, data.memberName, data.memberPass, data.memberPhone, "active"]);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success", message: "ส่งคำขอสมัครสมาชิกสำเร็จ"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "saveChecklist") {
      var stockSheet = ss.getSheetByName("Stock");
      if (!stockSheet || stockSheet.getLastRow() <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ไม่พบข้อมูลสต็อกในระบบ" })).setMimeType(ContentService.MimeType.JSON);
      }

      var sData = stockSheet.getDataRange().getValues();
      var stockMap = {};
      for (var r = 1; r < sData.length; r++) {
        if (sData[r][0]) {
          stockMap[sData[r][0].toString().trim().toUpperCase()] = {
            rowIndex: r + 1,
            qty: parseInt(sData[r][3]) || 0,
            name: sData[r][2],
            unit: sData[r][4],
            category: sData[r][1] ? sData[r][1].toString().trim() : "อุปกรณ์ทั่วไป"
          };
        }
      }

      var dataItems = data.usedItems || [];
      for (var u = 0; u < dataItems.length; u++) {
        var item = dataItems[u];
        var sIdKey = item.id.toUpperCase();
        if (stockMap[sIdKey]) {
          var currentStock = stockMap[sIdKey].qty;
          var requestedQty = parseInt(item.deductQty) || 0;
          if (currentStock < requestedQty) {
            return ContentService.createTextOutput(JSON.stringify({
              status: "error",
              message: "❌ อุปกรณ์ไม่พอเบิก! (" + stockMap[sIdKey].name + ") ในคลังเหลือ " + currentStock + " " + stockMap[sIdKey].unit + " แต่ต้องการเบิกถึง " + requestedQty + " " + stockMap[sIdKey].unit
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var stockLogSheet = ss.getSheetByName("StockTransactionLogs") || ss.insertSheet("StockTransactionLogs");
      if (stockLogSheet.getLastRow() === 0) {
        stockLogSheet.appendRow(["Time", "User", "SiteID", "StockID", "ItemName", "Qty", "Unit", "ReturnType"]);
      }

      var serialSheet = ss.getSheetByName("Serials");
      var serialData = serialSheet ? serialSheet.getDataRange().getValues() : [];

      var nowTime = new Date();
      dataItems.forEach(function(item) {
        var sIdKey = item.id.toUpperCase();
        var target = stockMap[sIdKey];
        if (target) {
          var deductVal = parseInt(item.deductQty) || 0;
          var newQty = Math.max(0, target.qty - deductVal);
          stockSheet.getRange(target.rowIndex, 4).setValue(newQty);
          target.qty = newQty;

          if (item.serial && item.serial !== "-") {
            for (var si = 1; si < serialData.length; si++) {
              var rowStockId = serialData[si][0] ? serialData[si][0].toString().trim().toUpperCase() : "";
              var rowSerialNo = serialData[si][1] ? serialData[si][1].toString().trim() : "";
              
              if (rowStockId === sIdKey && rowSerialNo === item.serial) {
                var currentSerialQty = parseInt(serialData[si][3]) || 1;
                var updatedSerialQty = currentSerialQty - deductVal;
                
                if (updatedSerialQty <= 0) {
                  serialSheet.getRange(si + 1, 3).setValue("ถูกใช้งาน");
                  serialSheet.getRange(si + 1, 4).setValue(0);
                } else {
                  serialSheet.getRange(si + 1, 4).setValue(updatedSerialQty);
                }
                break;
              }
            }
          }

          var returnTypeDesc = "ยืมอุปกรณ์ (ต้องนำกลับมาคืน)";
          if (target.category === "อุปกรณ์สายอากาศ") {
            returnTypeDesc = "ตัดสต็อกถาวร (ไม่ต้องคืน)";
          }

          var displayName = target.name;
          if (item.serial && item.serial !== "-") {
            displayName += " [รุ่น/SN: " + item.serial + "]";
          }

          stockLogSheet.appendRow([
            nowTime,
            data.userName || "Technician",
            data.siteId || "-",
            item.id,
            displayName,
            deductVal,
            target.unit,
            returnTypeDesc
          ]);
        }
      });

      updateCategorySheets(ss, stockSheet);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", message: "บันทึกผลการตรวจเช็คและอัปเดตสต็อก/Serial เรียบร้อยแล้ว"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "returnEquipment") {
      var stockSheet = ss.getSheetByName("Stock");
      if (!stockSheet || stockSheet.getLastRow() <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ไม่พบข้อมูลสต็อกในระบบ" })).setMimeType(ContentService.MimeType.JSON);
      }

      var sData = stockSheet.getDataRange().getValues();
      var stockMap = {};
      for (var r = 1; r < sData.length; r++) {
        if (sData[r][0]) {
          stockMap[sData[r][0].toString().trim().toUpperCase()] = {
            rowIndex: r + 1,
            qty: parseInt(sData[r][3]) || 0,
            name: sData[r][2],
            unit: sData[r][4],
            category: sData[r][1] ? sData[r][1].toString().trim() : "อุปกรณ์ทั่วไป"
          };
        }
      }

      var returnItems = data.returnItems || [];
      var stockLogSheet = ss.getSheetByName("StockTransactionLogs") || ss.insertSheet("StockTransactionLogs");
      if (stockLogSheet.getLastRow() === 0) {
        stockLogSheet.appendRow(["Time", "User", "SiteID", "StockID", "ItemName", "Qty", "Unit", "ReturnType"]);
      }

      var serialSheet = ss.getSheetByName("Serials");
      var serialData = serialSheet ? serialSheet.getDataRange().getValues() : [];

      var nowTime = new Date();
      returnItems.forEach(function(item) {
        var sIdKey = item.id.toUpperCase();
        var target = stockMap[sIdKey];
        if (target) {
          var returnQtyVal = parseInt(item.returnQty) || 0;
          var newQty = target.qty + returnQtyVal;
          stockSheet.getRange(target.rowIndex, 4).setValue(newQty);
          target.qty = newQty;

          if (serialSheet && serialData.length > 1) {
            for (var si = 1; si < serialData.length; si++) {
              var rowStockId = serialData[si][0] ? serialData[si][0].toString().trim().toUpperCase() : "";
              
              if (rowStockId === sIdKey) {
                var currentSerialQty = parseInt(serialData[si][3]) || 0;
                var updatedSerialQty = currentSerialQty + returnQtyVal;
                
                serialSheet.getRange(si + 1, 3).setValue("พร้อมใช้งาน");
                serialSheet.getRange(si + 1, 4).setValue(updatedSerialQty);
                break;
              }
            }
          }

          stockLogSheet.appendRow([
            nowTime,
            data.userName || "Technician",
            data.siteId || "-",
            item.id,
            target.name,
            returnQtyVal,
            target.unit,
            "📥 คืนอุปกรณ์เข้าคลัง (" + (data.note || "ปกติ") + ")"
          ]);
        }
      });

      updateCategorySheets(ss, stockSheet);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success", message: "บันทึกการคืนอุปกรณ์เข้าคลังและอัปเดต Serial เรียบร้อยแล้ว"
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error", message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function updateCategorySheets(ss, stockSheet) {
  var sRows = stockSheet.getDataRange().getValues();
  var categoryGroups = {};
  for (var j = 1; j < sRows.length; j++) {
    if (sRows[j][0]) {
      var sId = sRows[j][0].toString().trim();
      var cat = sRows[j][1] ? sRows[j][1].toString().trim() : "อุปกรณ์ทั่วไป";
      var name = sRows[j][2] ? sRows[j][2].toString().trim() : "";
      var qty = parseInt(sRows[j][3]) || 0;
      var unit = sRows[j][4] ? sRows[j][4].toString().trim() : "ชิ้น";

      if (!categoryGroups[cat]) {
        categoryGroups[cat] = [];
      }
      categoryGroups[cat].push([sId, cat, name, qty, unit]);
    }
  }

  for (var catName in categoryGroups) {
    var subSheet = ss.getSheetByName(catName) || ss.insertSheet(catName);
    subSheet.clear();
    subSheet.appendRow(["ID", "Category", "Name", "Qty", "Unit"]);
    var itemsList = categoryGroups[catName];
    if (itemsList.length > 0) {
      subSheet.getRange(2, 1, itemsList.length, itemsList[0].length).setValues(itemsList);
    }
  }
}