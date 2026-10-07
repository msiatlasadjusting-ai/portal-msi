// =========================================================
// NOTIFICATION & SLA MONITORING SERVICE (UPDATED FOR NEW 55 COLUMNS)
// Penanganan kalkulasi SLA penugasan, notifikasi portal in-app, dan email reminder harian
// =========================================================

function getAppSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}
  
  if (typeof SPREADSHEET_ID !== "undefined" && SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return null;
}

// Validasi Tanggal Presisi agar Teks seperti 'MSI', 'APB', 'PC', '-' Tidak Dibaca Sebagai Tanggal
function isValidDateStr(val) {
  if (!val || val === '-' || String(val).trim() === '') return false;
  var str = String(val).trim().toUpperCase();
  if (str === 'MSI' || str === 'APB' || str === 'PC' || str === 'ON PROCESS' || str === 'REPORT SENT') return false;

  if (val instanceof Date) return !isNaN(val.getTime());

  var d = new Date(val);
  return !isNaN(d.getTime());
}

function parseDateToYmdSla(val) {
  if (!isValidDateStr(val)) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
  }
  var str = String(val).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) return str.substring(0, 10);
  
  var d = new Date(str);
  return Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
}

function addDaysToDateStr(dateStr, days) {
  var parts = dateStr.split('-');
  var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
}

// =========================================================
// HELPER CALCULATE JOB STATUS BACKEND (100% IDENTIK DENGAN FRONTEND)
// =========================================================
function getCalculatedJobStatusBackend(item) {
  if (String(item["Status"]).trim() === "Cancelled") return "Cancelled";
  if (item["Paid Date"]) return "Invoice Paid";
  if (item["Invoice Date"]) return "Invoice Sent";
  if (item["LOP Approved Date"]) return "Awaiting Invoice";
  if (item["Report Sent Date"]) return "Report Sent";
  if (item["Report Start Date"]) return "Drafting Report";
  if (item["Survey Start Date"] || item["Survey Finish Date"]) return "In Site Survey";
  return "On Process";
}

// =========================================================
// CHECK AND GENERATE NOTIFICATIONS (IN-APP PORTAL)
// =========================================================
function checkAndGenerateNotifications() {
  try {
    var ss = getAppSpreadsheet();
    if (!ss) return;

    var sheetName = (typeof SHEET_NAME !== "undefined" && SHEET_NAME) ? SHEET_NAME : "Mastercase";
    var masterSheet = ss.getSheetByName(sheetName);
    var notifSheet = ss.getSheetByName("Notifications");

    if (!masterSheet || !notifSheet) {
      Logger.log("[NotifService] ERROR: Sheet Mastercase atau Notifications tidak ditemukan.");
      return;
    }

    var masterData = masterSheet.getDataRange().getValues();
    if (masterData.length <= 1) return;

    var headerRowIdx = -1;
    for (var r = 0; r < Math.min(masterData.length, 10); r++) {
      var rowStr = masterData[r].map(function (c) { return String(c).toLowerCase().trim(); }).join(" ");
      if (rowStr.indexOf("job number") !== -1 || rowStr.indexOf("instruction date") !== -1) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) headerRowIdx = 0; 

    var headers = masterData[headerRowIdx].map(function (c) { return String(c).trim().toLowerCase(); });

    var getColIdx = function (colName) {
      return headers.indexOf(String(colName).trim().toLowerCase());
    };

    var colJobNo    = getColIdx("Job Number");
    var colPic      = getColIdx("Lead PIC");
    if (colPic === -1) colPic = getColIdx("PIC");
    var colStatus   = getColIdx("Status");

    var colS2Finish = getColIdx("Survey Finish Date");
    var colS3Start  = getColIdx("Report Start Date");
    var colS3Finish = getColIdx("Report Sent Date");
    var colS5Start  = getColIdx("Invoice Date");
    var colS5Finish = getColIdx("Paid Date");

    var existingNotifs = notifSheet.getDataRange().getValues();
    var existingIds = {};
    for (var i = 1; i < existingNotifs.length; i++) {
      if (existingNotifs[i][0]) {
        existingIds[String(existingNotifs[i][0]).trim()] = true;
      }
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var todayStr = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd");

    var generatedCount = 0;

    for (var r = headerRowIdx + 1; r < masterData.length; r++) {
      var row = masterData[r];
      var jobNo = String(row[colJobNo] || '').trim();
      var pic = String(row[colPic] || '').trim();
      var rawStatus = (colStatus !== -1 && row[colStatus]) ? String(row[colStatus]).trim() : "";

      if (!jobNo || jobNo === '-' || jobNo.toLowerCase() === 'job number') continue;

      var s2Finish = parseDateToYmdSla(colS2Finish !== -1 ? row[colS2Finish] : "");
      var s3Start  = parseDateToYmdSla(colS3Start  !== -1 ? row[colS3Start]  : "");
      var s3Finish = parseDateToYmdSla(colS3Finish !== -1 ? row[colS3Finish] : "");
      var s5Start  = parseDateToYmdSla(colS5Start  !== -1 ? row[colS5Start]  : "");
      var s5Finish = parseDateToYmdSla(colS5Finish !== -1 ? row[colS5Finish] : "");

      if (s5Finish !== "" || rawStatus === "Invoice Paid" || rawStatus === "Cancelled") {
        continue;
      }

      var activeTargetDateStr = "";
      var stageTag = "";
      var stageTitle = "";

      if (!s2Finish && !s3Start) {
        continue;
      } 
      else if (!s3Finish) {
        var baseDate = s3Start ? s3Start : s2Finish;
        activeTargetDateStr = addDaysToDateStr(baseDate, 5);
        stageTag = "REPORT";
        stageTitle = "Interim Report";
      } 
      else if (s5Start) {
        activeTargetDateStr = addDaysToDateStr(s5Start, 7);
        stageTag = "INVOICE";
        stageTitle = "Fee Invoicing & Settlement";
      }

      if (activeTargetDateStr) {
        var parts = activeTargetDateStr.split('-');
        var targetDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        targetDate.setHours(0, 0, 0, 0);

        var diffTime = targetDate.getTime() - today.getTime();
        var daysLeft = Math.round(diffTime / (1000 * 60 * 60 * 24));

        var notifId = "";
        var title = "";
        var message = "";
        var type = "";

        if (daysLeft < 0) {
          notifId = "NOTIF-" + jobNo + "-" + stageTag + "-OVERDUE-" + todayStr;
          type = "Overdue";
          title = "Overdue: " + stageTitle;
          message = "Job #" + jobNo + " stage " + stageTitle + " telah terlewati " + Math.abs(daysLeft) + " hari.";
        } else if (daysLeft <= 3 && daysLeft >= 0) {
          notifId = "NOTIF-" + jobNo + "-" + stageTag + "-WARNING-" + todayStr;
          type = "Warning";
          title = "Mendekati Deadline " + stageTitle;
          message = (daysLeft === 0)
            ? "Job #" + jobNo + " memasuki hari H target deadline " + stageTitle + "."
            : "Job #" + jobNo + " tersisa " + daysLeft + " hari sebelum target " + stageTitle + ".";
        }

        if (notifId) {
          if (!existingIds[notifId]) {
            var nowStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
            notifSheet.appendRow([notifId, jobNo, type, title, message, "Unread", nowStr, pic]);
            existingIds[notifId] = true;
            generatedCount++;
            Logger.log("[NotifService] NOTIFIKASI BARU BERHASIL DIBUAT: " + notifId + " untuk Job #" + jobNo);
          }
        }
      }
    }

    Logger.log("[NotifService] Selesai. Total notifikasi baru dibuat: " + generatedCount);

  } catch (err) {
    Logger.log("Error checkAndGenerateNotifications: " + err.message);
  }
}

// =========================================================
// SEND DAILY SLA REMINDER EMAIL
// =========================================================
function sendDailySlaReminderEmail() {
  var ss = getAppSpreadsheet();
  if (!ss) return;

  var sheetName = (typeof SHEET_NAME !== "undefined" && SHEET_NAME) ? SHEET_NAME : "Mastercase";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0].map(function (c) { return String(c).trim().toLowerCase(); });

  var getColIdx = function (colName) {
    return headers.indexOf(String(colName).trim().toLowerCase());
  };

  var idxJobNo    = getColIdx("Job Number");
  var idxInsured  = getColIdx("Owner / Client's Customer");
  var idxPic      = getColIdx("Lead PIC");
  if (idxPic === -1) idxPic = getColIdx("PIC");

  var idxS2Finish = getColIdx("Survey Finish Date");
  var idxS3Start  = getColIdx("Report Start Date");
  var idxS3Finish = getColIdx("Report Sent Date");
  var idxS4App    = getColIdx("LOP Approved Date");
  var idxS5Start  = getColIdx("Invoice Date");
  var idxS5Finish = getColIdx("Paid Date");
  var idxStatus   = getColIdx("Status");

  var reminderList = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var jobNo = row[idxJobNo];
    if (!jobNo) continue;

    var item = {
      "Job Number": jobNo,
      "Survey Finish Date": parseDateToYmdSla(row[idxS2Finish]),
      "Report Start Date": parseDateToYmdSla(row[idxS3Start]),
      "Report Sent Date": parseDateToYmdSla(row[idxS3Finish]),
      "LOP Approved Date": parseDateToYmdSla(row[idxS4App]),
      "Invoice Date": parseDateToYmdSla(row[idxS5Start]),
      "Paid Date": parseDateToYmdSla(row[idxS5Finish]),
      "Status": row[idxStatus]
    };

    var currentStatus = getCalculatedJobStatusBackend(item);
    if (currentStatus === "Invoice Paid" || currentStatus === "Cancelled") continue;

    var activeTargetDateStr = "";
    if (item["Survey Finish Date"] && !item["Report Sent Date"]) {
      var baseDate = item["Report Start Date"] ? item["Report Start Date"] : item["Survey Finish Date"];
      activeTargetDateStr = addDaysToDateStr(baseDate, 5);
    } else if (item["Invoice Date"] && !item["Paid Date"]) {
      activeTargetDateStr = addDaysToDateStr(item["Invoice Date"], 7);
    }

    if (activeTargetDateStr) {
      var parts = activeTargetDateStr.split('-');
      var targetDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      targetDate.setHours(0, 0, 0, 0);

      var diffTime = targetDate.getTime() - today.getTime();
      var daysLeft = Math.round(diffTime / (1000 * 3600 * 24));

      var slaText = "";
      if (daysLeft < 0) {
        slaText = "Overdue " + Math.abs(daysLeft) + " Hari";
      } else if (daysLeft <= 3) {
        slaText = "H-" + daysLeft + " Target";
      }

      if (slaText !== "") {
        reminderList.push({
          jobNo: "#" + jobNo,
          insured: row[idxInsured] || "N/A",
          pic: row[idxPic] || "-",
          status: currentStatus,
          slaText: slaText
        });
      }
    }
  }

  if (reminderList.length === 0) return;

  var htmlBody = '<div style="font-family: Arial, sans-serif; color: #334155; max-width: 700px; margin: auto;">' +
    '<h2 style="color: #EC1F27; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px;">Daily SLA Case Summary Report</h2>' +
    '<p style="font-size: 13px;">Tanggal: <b>' + Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy") + '</b></p>' +
    '<p style="font-size: 13px;">Berikut adalah ringkasan pekerjaan yang memerlukan perhatian (Overdue / Mendekati Target SLA):</p>' +
    '<table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 15px;">' +
    '<thead>' +
    '<tr style="background-color: #2D323E; color: #ffffff; text-align: left;">' +
    '<th style="padding: 10px; border: 1px solid #ddd;">No</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Job Number</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Tertanggung (Insured)</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">PIC</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Status Job</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Reminder SLA</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>';

  reminderList.forEach(function (item, idx) {
    var bgRow = (idx % 2 === 0) ? '#ffffff' : '#f8fafc';
    htmlBody += '<tr style="background-color: ' + bgRow + ';">' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; text-align: center;">' + (idx + 1) + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold; color: #EC1F27;">' + item.jobNo + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0;">' + item.insured + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0;">' + item.pic + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">' + item.status + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; color: #dc2626; font-weight: bold;">' + item.slaText + '</td>' +
      '</tr>';
  });

  htmlBody += '</tbody></table>' +
    '<p style="font-size: 11px; color: #94a3b8; margin-top: 20px;">Email ini dikirimkan secara otomatis oleh MSI Portal System.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: "msi.atlasadjusting@gmail.com",
    subject: "[REMINDER SLA] Daily Case Report - " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"),
    htmlBody: htmlBody
  });
}

// =========================================================
// SEND DAILY SLA REMINDER EMAIL (FIXED DATA SYNCHRONIZATION)
// =========================================================
function sendDailySlaReminderEmail() {
  var ss = getAppSpreadsheet();
  if (!ss) return;

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0].map(function(c) { return String(c).trim().toLowerCase(); });
  
  var getColIdx = function(colName) {
    return headers.indexOf(String(colName).trim().toLowerCase());
  };

  var idxJobNo    = getColIdx("Job Number");
  var idxInsured  = getColIdx("Owner / Client's Customer");
  var idxPic      = getColIdx("Lead PIC");
  if (idxPic === -1) idxPic = getColIdx("PIC");

  var idxS2Finish = getColIdx("Survey Finish Date");
  var idxS3Start  = getColIdx("Report Start Date");
  var idxS3Finish = getColIdx("Report Sent Date");
  var idxS4App    = getColIdx("LOP Approved Date");
  var idxS5Start  = getColIdx("Invoice Date");
  var idxS5Finish = getColIdx("Paid Date");
  var idxStatus   = getColIdx("Status");

  var reminderList = [];
  var today = new Date();
  today.setHours(0,0,0,0);

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var jobNo = row[idxJobNo];
    if (!jobNo) continue;

    var item = {
      "Job Number": jobNo,
      "Survey Finish Date": parseDateToYmd(row[idxS2Finish]),
      "Report Start Date": parseDateToYmd(row[idxS3Start]),
      "Report Sent Date": parseDateToYmd(row[idxS3Finish]),
      "LOP Approved Date": parseDateToYmd(row[idxS4App]),
      "Invoice Date": parseDateToYmd(row[idxS5Start]),
      "Paid Date": parseDateToYmd(row[idxS5Finish]),
      "Status": row[idxStatus]
    };

    // 1. Dapatkan Status Dinamis Kasus
    var currentStatus = getCalculatedJobStatusBackend(item);

    // Abaikan Kasus yang Sudah Lunas atau Cancelled
    if (currentStatus === "Invoice Paid" || currentStatus === "Cancelled") continue;

    // 2. Evaluasi Target SLA
    var targetDate = null;
    if (item["Survey Finish Date"] && !item["Report Sent Date"]) {
      targetDate = new Date(item["Survey Finish Date"]);
      targetDate.setDate(targetDate.getDate() + 5); // Target Laporan 5 Hari
    } else if (item["Invoice Date"] && !item["Paid Date"]) {
      targetDate = new Date(item["Invoice Date"]);
      targetDate.setDate(targetDate.getDate() + 7); // Target Bayar 7 Hari
    }

    if (targetDate) {
      targetDate.setHours(0,0,0,0);
      var diffTime = targetDate.getTime() - today.getTime();
      var daysLeft = Math.round(diffTime / (1000 * 3600 * 24));

      var slaText = "";
      if (daysLeft < 0) {
        slaText = "Overdue " + Math.abs(daysLeft) + " Hari";
      } else if (daysLeft <= 3) {
        slaText = "H-" + daysLeft + " Target";
      }

      if (slaText !== "") {
        reminderList.push({
          jobNo: "#" + jobNo,
          insured: row[idxInsured] || "N/A",
          pic: row[idxPic] || "-",
          status: currentStatus,
          slaText: slaText
        });
      }
    }
  }

  if (reminderList.length === 0) return;

  // 3. Susun Email HTML Body
  var htmlBody = '<div style="font-family: Arial, sans-serif; color: #334155; max-width: 700px; margin: auto;">' +
    '<h2 style="color: #EC1F27; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px;">Daily SLA Case Summary Report</h2>' +
    '<p style="font-size: 13px;">Tanggal: <b>' + Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy") + '</b></p>' +
    '<p style="font-size: 13px;">Berikut adalah ringkasan pekerjaan yang memerlukan perhatian (Overdue / Mendekati Target SLA):</p>' +
    '<table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 15px;">' +
    '<thead>' +
    '<tr style="background-color: #2D323E; color: #ffffff; text-align: left;">' +
    '<th style="padding: 10px; border: 1px solid #ddd;">No</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Job Number</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Tertanggung (Insured)</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">PIC</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Status Job</th>' +
    '<th style="padding: 10px; border: 1px solid #ddd;">Reminder SLA</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>';

  reminderList.forEach(function(item, idx) {
    var bgRow = (idx % 2 === 0) ? '#ffffff' : '#f8fafc';
    htmlBody += '<tr style="background-color: ' + bgRow + ';">' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; text-align: center;">' + (idx + 1) + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold; color: #EC1F27;">' + item.jobNo + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0;">' + item.insured + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0;">' + item.pic + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; font-weight: bold;">' + item.status + '</td>' +
      '<td style="padding: 8px; border: 1px solid #e2e8f0; color: #dc2626; font-weight: bold;">' + item.slaText + '</td>' +
      '</tr>';
  });

  htmlBody += '</tbody></table>' +
    '<p style="font-size: 11px; color: #94a3b8; margin-top: 20px;">Email ini dikirimkan secara otomatis oleh MSI Portal System.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: "msi.atlasadjusting@gmail.com",
    subject: "[REMINDER SLA] Daily Case Report - " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"),
    htmlBody: htmlBody
  });
}

function getNotificationsData(currentUserEmail) {
  try {
    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "Spreadsheet tidak dapat diakses." };

    var sheet = ss.getSheetByName("Notifications");
    if (!sheet) return { success: true, data: [], unreadCount: 0 };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [], unreadCount: 0 };

    var userEmail = (currentUserEmail || "").toLowerCase().trim();
    var list = [];
    var unreadCount = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var notifId = String(data[i][0] || '').trim();
      if (!notifId) continue;

      var readByStr = String(data[i][5] || '').toLowerCase();
      var isReadByUser = false;

      if (userEmail && readByStr) {
        var readUsersList = readByStr.split(',').map(function (u) { return u.trim(); });
        if (readUsersList.indexOf(userEmail) !== -1) {
          isReadByUser = true;
        }
      }

      var statusDisplay = isReadByUser ? "Read" : "Unread";
      if (!isReadByUser) unreadCount++;

      var rawDate = data[i][6];
      var dateDisplay = rawDate;
      if (rawDate instanceof Date) {
        var tz = ss.getSpreadsheetTimeZone() || "Asia/Jakarta";
        dateDisplay = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd HH:mm");
      }

      list.push({
        rowIndex: i + 1,
        notifId: notifId,
        jobNo: String(data[i][1] || '-'),
        type: String(data[i][2] || 'Info'),
        title: String(data[i][3] || 'Notifikasi'),
        message: String(data[i][4] || ''),
        status: statusDisplay,
        createdAt: String(dateDisplay || '-'),
        recipientPic: String(data[i][7] || '-')
      });
    }

    return { success: true, data: list, unreadCount: unreadCount };
  } catch (err) {
    return { success: false, error: "Apps Script Error: " + err.message };
  }
}

function markNotificationAsRead(notifId, currentUserEmail) {
  try {
    if (!currentUserEmail) return { success: false, error: "User Email diperlukan." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false };

    var sheet = ss.getSheetByName("Notifications");
    if (!sheet) return { success: false };

    var userEmail = currentUserEmail.toLowerCase().trim();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(notifId).trim()) {
        var currentReadBy = String(data[i][5] || '').trim();
        var readUsersList = currentReadBy ? currentReadBy.split(',').map(function (u) { return u.trim().toLowerCase(); }) : [];

        if (readUsersList.indexOf(userEmail) === -1) {
          readUsersList.push(userEmail);
          var updatedReadBy = readUsersList.join(', ');
          sheet.getRange(i + 1, 6).setValue(updatedReadBy);
        }
        break;
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function markAllNotificationsAsRead(currentUserEmail) {
  try {
    if (!currentUserEmail) return { success: false, error: "User Email diperlukan." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false };

    var sheet = ss.getSheetByName("Notifications");
    if (!sheet) return { success: false };

    var userEmail = currentUserEmail.toLowerCase().trim();
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      var range = sheet.getRange(2, 6, lastRow - 1, 1);
      var values = range.getValues();

      for (var i = 0; i < values.length; i++) {
        var currentReadBy = String(values[i][0] || '').trim();
        var readUsersList = currentReadBy ? currentReadBy.split(',').map(function (u) { return u.trim().toLowerCase(); }) : [];

        if (readUsersList.indexOf(userEmail) === -1) {
          readUsersList.push(userEmail);
          values[i][0] = readUsersList.join(', ');
        }
      }
      range.setValues(values);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}