// =========================================================
// EMAIL NOTIFICATION & DIGEST SERVICE
// Pengiriman laporan ringkasan email harian dan notifikasi per-case kepada PIC serta Manager
// =========================================================

// =========================================================
// GET MANAGER EMAILS FROM FIRESTORE
// Mengambil daftar email seluruh user dengan role 'manager' atau 'admin' dari database Firestore
// =========================================================
function getManagerEmailsFromFirestore() {
  var firestoreUrl = "https://firestore.googleapis.com/v1/projects/" + (typeof FIREBASE_PROJECT_ID !== "undefined" ? FIREBASE_PROJECT_ID : "") + "/databases/(default)/documents:runQuery";

  try {
    var queryPayload = {
      "structuredQuery": {
        "from": [{ "collectionId": "users" }],
        "where": {
          "fieldFilter": {
            "field": { "fieldPath": "role" },
            "op": "EQUAL",
            "value": { "stringValue": "manager" }
          }
        }
      }
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(queryPayload),
      "headers": {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      },
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(firestoreUrl, options);
    var jsonResult = JSON.parse(response.getContentText());

    var managerEmails = [];

    if (Array.isArray(jsonResult)) {
      jsonResult.forEach(function (item) {
        if (item.document && item.document.fields) {
          var fields = item.document.fields;
          var emailVal = fields.email ? fields.email.stringValue : "";
          if (emailVal && emailVal.trim() !== "") {
            managerEmails.push(emailVal.trim().toLowerCase());
          }
        }
      });
    }

    if (managerEmails.length > 0) {
      return managerEmails.join(", ");
    }
  } catch (err) {
    Logger.log("Error fetch Firestore Manager Emails: " + err.message);
  }

  return (typeof DEFAULT_MANAGER_FALLBACK_EMAIL !== "undefined" && DEFAULT_MANAGER_FALLBACK_EMAIL) ? DEFAULT_MANAGER_FALLBACK_EMAIL : "msi.atlasadjusting@gmail.com";
}

// =========================================================
// GET PIC EMAIL BY NAME
// Mencari alamat email PIC yang terdaftar berdasarkan Nama atau Kode pada Sheet PIC
// =========================================================
function getPicEmailByName(picName) {
  try {
    var ss = getAppSpreadsheet();
    var picSheet = ss.getSheetByName("PIC");
    if (!picSheet) return null;

    var lastRow = picSheet.getLastRow();
    if (lastRow <= 1) return null;

    var data = picSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var target = String(picName || '').trim().toLowerCase();

    for (var i = 0; i < data.length; i++) {
      var nama = String(data[i][0] || '').trim().toLowerCase();
      var kode = String(data[i][1] || '').trim().toLowerCase();
      var email = String(data[i][2] || '').trim();

      if ((target === nama || target === kode) && email.indexOf("@") !== -1) {
        return email;
      }
    }
  } catch (e) {
    Logger.log("Error getPicEmailByName: " + e.message);
  }
  return null;
}

// =========================================================
// SEND SINGLE CASE EMAIL TO PIC
// Mengirimkan notifikasi email individual untuk 1 case yang mengalami alert SLA ke PIC
// =========================================================
function sendSingleCaseEmailToPIC(recipientEmail, caseObj) {
  if (!recipientEmail || !caseObj) {
    Logger.log("[EmailService] Warning: recipientEmail atau caseObj tidak valid.");
    return;
  }

  var isOverdue = caseObj.isOverdue || false;
  var reminderStatus = caseObj.reminderStatus || "Target SLA";
  var jobNo = caseObj.jobNo || "-";
  var insured = caseObj.insured || "-";
  var picName = caseObj.pic || "Adjuster";
  var jobStatus = caseObj.jobStatus || "On Process";

  var headerTitle = isOverdue ? "Peringatan SLA Case Overdue" : "Peringatan SLA Case Warning";
  var headerColor = isOverdue ? "#EC1F27" : "#D97706";
  var subject = "[" + String(reminderStatus).toUpperCase() + "] Job #" + jobNo + " - " + insured;

  var htmlBody =
    '<div style="font-family: Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.6;">' +
    '<h2 style="color: ' + headerColor + '; margin-bottom: 10px;">' + headerTitle + '</h2>' +
    '<p>Halo <strong>' + picName + '</strong>,</p>' +
    '<p>Berikut adalah pembaruan status SLA pekerjaan yang Anda tangani saat ini:</p>' +
    '<table style="border-collapse: collapse; width: 100%; max-width: 500px; margin: 15px 0; border: 1px solid #E5E7EB;">' +
    '<tr><td style="padding: 8px 12px; font-weight: bold; background-color: #F8FAFC; border-bottom: 1px solid #E5E7EB;">Job Number</td><td style="padding: 8px 12px; border-bottom: 1px solid #E5E7EB;">#' + jobNo + '</td></tr>' +
    '<tr><td style="padding: 8px 12px; font-weight: bold; background-color: #F8FAFC; border-bottom: 1px solid #E5E7EB;">Tertanggung (Insured)</td><td style="padding: 8px 12px; border-bottom: 1px solid #E5E7EB;">' + insured + '</td></tr>' +
    '<tr><td style="padding: 8px 12px; font-weight: bold; background-color: #F8FAFC; border-bottom: 1px solid #E5E7EB;">Status Job</td><td style="padding: 8px 12px; border-bottom: 1px solid #E5E7EB;"><span style="background-color: #F3E8FF; color: #6B21A8; padding: 2px 8px; border-radius: 6px; font-weight: bold; font-size: 11px;">' + jobStatus + '</span></td></tr>' +
    '<tr><td style="padding: 8px 12px; font-weight: bold; background-color: #F8FAFC; border-bottom: 1px solid #E5E7EB;">Reminder SLA</td><td style="padding: 8px 12px; color: ' + headerColor + '; font-weight: bold; border-bottom: 1px solid #E5E7EB;">' + reminderStatus + '</td></tr>' +
    '</table>' +
    '<p>Mohon segera melakukan pembaruan status pada Portal Web MSI.</p>' +
    '<br><p style="font-size: 11px; color: #94A3B8;">Email otomatis dari Sistem Portal Mahardhika Surveyor Indonesia.</p>' +
    '</div>';

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("[EmailService] Email berhasil dikirim ke: " + recipientEmail + " untuk Job #" + jobNo);
  } catch (e) {
    Logger.log("Gagal mengirim email ke " + recipientEmail + ": " + e.message);
  }
}

// =========================================================
// SEND SUMMARY EMAIL TO MANAGER
// Mengirimkan laporan rekapitulasi seluruh case overdue dan warning kepada manager
// =========================================================
function sendSummaryEmailToManager(managerEmails, alertList) {
  if (!managerEmails) return;

  var todayStr = Utilities.formatDate(new Date(), "GMT+7", "dd MMM yyyy");
  var subject = "[DAILY SUMMARY] Ringkasan Case SLA - " + todayStr + " (" + alertList.length + " Case)";

  var rowsHtml = "";
  alertList.forEach(function (item, idx) {
    var statusBg = item.isOverdue ? "#FEF2F2" : "#FFFBEB";
    var statusColor = item.isOverdue ? "#DC2626" : "#D97706";

    rowsHtml +=
      '<tr style="border-bottom: 1px solid #E5E7EB;">' +
      '<td style="padding: 10px; text-align: center;">' + (idx + 1) + '</td>' +
      '<td style="padding: 10px; font-weight: bold; color: #0F172A;">#' + item.jobNo + '</td>' +
      '<td style="padding: 10px;">' + item.insured + '</td>' +
      '<td style="padding: 10px; text-align: center;"><span style="background-color: #F1F5F9; padding: 2px 8px; border-radius: 6px; font-weight: bold;">' + item.pic + '</span></td>' +
      '<td style="padding: 10px; text-align: center;"><span style="background-color: #F3E8FF; color: #6B21A8; padding: 3px 8px; border-radius: 6px; font-weight: bold; font-size: 11px;">' + item.jobStatus + '</span></td>' +
      '<td style="padding: 10px; text-align: center;"><span style="background-color: ' + statusBg + '; color: ' + statusColor + '; padding: 4px 10px; border-radius: 8px; font-weight: bold; font-size: 11px;">' + item.reminderStatus + '</span></td>' +
      '</tr>';
  });

  var htmlBody =
    '<div style="font-family: Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.6;">' +
    '<h2 style="color: #1E232A; margin-bottom: 5px;">Daily SLA Case Summary Report</h2>' +
    '<p style="color: #64748B; font-size: 12px; margin-top: 0;">Tanggal: ' + todayStr + '</p>' +
    '<p>Berikut adalah ringkasan seluruh pekerjaan yang memerlukan perhatian (Overdue maupun Mendekati Target SLA):</p>' +
    '<table style="border-collapse: collapse; width: 100%; border: 1px solid #CBD5E1; margin-top: 15px;">' +
    '<thead>' +
    '<tr style="background-color: #1E232A; color: #FFFFFF; font-size: 12px; text-align: left;">' +
    '<th style="padding: 10px; text-align: center;">No</th>' +
    '<th style="padding: 10px;">Job Number</th>' +
    '<th style="padding: 10px;">Tertanggung (Insured)</th>' +
    '<th style="padding: 10px; text-align: center;">PIC</th>' +
    '<th style="padding: 10px; text-align: center;">Status Job</th>' +
    '<th style="padding: 10px; text-align: center;">Reminder SLA</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' +
    rowsHtml +
    '</tbody>' +
    '</table>' +
    '<br><p style="font-size: 11px; color: #94A3B8;">Laporan ini dibuat secara otomatis oleh Portal Management MSI.</p>' +
    '</div>';

  try {
    MailApp.sendEmail({
      to: managerEmails,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("[EmailService] Summary Email berhasil dikirim ke Manager: " + managerEmails);
  } catch (e) {
    Logger.log("Gagal mengirim summary email ke Manager: " + e.message);
  }
}

// =========================================================
// PROCESS DAILY EMAIL NOTIFICATIONS (SYNCHRONIZED WITH BACKEND & UI)
// =========================================================
function processDailyEmailNotifications() {
  if (typeof ENABLE_EMAIL_NOTIFICATIONS !== "undefined" && !ENABLE_EMAIL_NOTIFICATIONS) {
    Logger.log("[EmailService] Fitur email nonaktif (ENABLE_EMAIL_NOTIFICATIONS = false).");
    return;
  }

  try {
    var ss = getAppSpreadsheet();
    if (!ss) return;

    var sheetName = (typeof SHEET_NAME !== "undefined" && SHEET_NAME) ? SHEET_NAME : "Mastercase";
    var masterSheet = ss.getSheetByName(sheetName);
    var picSheet = ss.getSheetByName("PIC");

    if (!masterSheet) return;

    var masterData = masterSheet.getDataRange().getValues();
    if (masterData.length <= 1) return;

    // 1. CARI BARIS HEADER (BARIS 1 S/D 10 DENGAN MATCHING LOWERCASE)
    var headerRowIdx = -1;
    for (var r = 0; r < Math.min(masterData.length, 10); r++) {
      var rowStr = masterData[r].map(function (c) { return String(c).toLowerCase().trim(); }).join(" ");
      if (rowStr.indexOf("job number") !== -1 || rowStr.indexOf("instruction date") !== -1) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) headerRowIdx = 0;

    var rawHeaders = masterData[headerRowIdx].map(function (c) { return String(c).trim(); });
    var lowerHeaders = rawHeaders.map(function(h) { return h.toLowerCase().trim(); });

    var findColIndex = function(possibleNames) {
      for (var k = 0; k < possibleNames.length; k++) {
        var target = possibleNames[k].toLowerCase().trim();
        var idx = lowerHeaders.indexOf(target);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    var colJobNo    = findColIndex(["Job Number", "Job No"]);
    var colPic      = findColIndex(["Lead PIC", "PIC"]);
    var colInsured  = findColIndex(["Owner / Client's Customer", "Owner/ Client's Customer", "Owner", "Insured"]);
    var colStatus   = findColIndex(["Status"]);
    
    var colS2Finish = findColIndex(["Survey Finish Date"]);
    var colS3Start  = findColIndex(["Report Start Date"]);
    var colS3Finish = findColIndex(["Report Sent Date"]);
    var colS5Start  = findColIndex(["Invoice Date"]);
    var colS5Finish = findColIndex(["Paid Date"]);

    // 2. MAP EMAIL PIC DARI SHEET "PIC"
    var picEmailMap = {};
    if (picSheet) {
      var picData = picSheet.getDataRange().getValues();
      for (var u = 1; u < picData.length; u++) {
        var picNama   = String(picData[u][0] || '').trim().toUpperCase();
        var picKode   = String(picData[u][1] || '').trim().toUpperCase();
        var userEmail = String(picData[u][2] || '').trim().toLowerCase();

        if (userEmail && userEmail.indexOf("@") !== -1) {
          if (picNama) picEmailMap[picNama] = userEmail;
          if (picKode) picEmailMap[picKode] = userEmail;
        }
      }
    }

    var MANAGER_EMAILS = getManagerEmailsFromFirestore() || "msi.atlasadjusting@gmail.com";
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var alertCasesList = [];

    // 3. ITERASI TIAP KASUS PADA MASTERCASE
    for (var i = headerRowIdx + 1; i < masterData.length; i++) {
      var row = masterData[i];
      var jobNo = String(row[colJobNo] || '').trim();
      var picRaw = String(row[colPic] || '').trim();
      var picUpper = picRaw.toUpperCase();
      var insured = (colInsured !== -1 && row[colInsured]) ? String(row[colInsured]).trim() : "-";
      var rawStatus = (colStatus !== -1 && row[colStatus]) ? String(row[colStatus]).trim() : "";

      if (!jobNo || jobNo === '-' || jobNo.toLowerCase() === 'job number') continue;

      var s2Finish = parseDateToYmdSla(colS2Finish !== -1 ? row[colS2Finish] : "");
      var s3Start  = parseDateToYmdSla(colS3Start  !== -1 ? row[colS3Start]  : "");
      var s3Finish = parseDateToYmdSla(colS3Finish !== -1 ? row[colS3Finish] : "");
      var s5Start  = parseDateToYmdSla(colS5Start  !== -1 ? row[colS5Start]  : "");
      var s5Finish = parseDateToYmdSla(colS5Finish !== -1 ? row[colS5Finish] : "");

      // ABAIKAN KASUS LUNAS ATAU CANCELLED
      if (s5Finish !== "" || rawStatus === "Invoice Paid" || rawStatus === "Cancelled") {
        continue;
      }

      var activeTargetDateStr = "";
      var jobStatusDisplay = rawStatus || "On Process";

      // EVALUASI PERSIS SESUAI LOGIKA UI JOB DETAIL
      if (!s2Finish && !s3Start) {
        continue;
      } else if (!s3Finish) {
        var baseDate = s3Start ? s3Start : s2Finish;
        activeTargetDateStr = addDaysToDateStr(baseDate, 5);
        jobStatusDisplay = "Drafting Report";
      } else if (s5Start) {
        activeTargetDateStr = addDaysToDateStr(s5Start, 7);
        jobStatusDisplay = "Invoice Sent";
      }

      if (activeTargetDateStr) {
        var dateParts = activeTargetDateStr.split('-');
        var targetDate = new Date(parseInt(dateParts[0], 10), parseInt(dateParts[1], 10) - 1, parseInt(dateParts[2], 10));
        targetDate.setHours(0, 0, 0, 0);

        var diffTime = targetDate.getTime() - today.getTime();
        var daysLeft = Math.round(diffTime / (1000 * 60 * 60 * 24));

        var reminderStatus = "";
        var isAlert = false;

        if (daysLeft < 0) {
          reminderStatus = "Overdue " + Math.abs(daysLeft) + " Hari";
          isAlert = true;
        } else if (daysLeft <= 3 && daysLeft >= 0) {
          reminderStatus = (daysLeft === 0) ? "Hari Ini Target SLA" : "H-" + daysLeft + " Target";
          isAlert = true;
        }

        if (isAlert) {
          var caseObj = {
            jobNo: jobNo,
            pic: picRaw,
            insured: insured,
            jobStatus: jobStatusDisplay,
            reminderStatus: reminderStatus,
            isOverdue: daysLeft < 0
          };

          alertCasesList.push(caseObj);
          var targetPicEmail = picEmailMap[picUpper] || MANAGER_EMAILS;
          
          if (targetPicEmail) {
            sendSingleCaseEmailToPIC(targetPicEmail, caseObj);
          }
        }
      }
    }

    if (alertCasesList.length > 0) {
      sendSummaryEmailToManager(MANAGER_EMAILS, alertCasesList);
    } else {
      Logger.log("[EmailService] Tidak ada kasus yang memenuhi kriteria alert SLA hari ini.");
    }

  } catch (err) {
    Logger.log("Error processDailyEmailNotifications: " + err.message);
  }
}

// =========================================================
// SEND DAILY EMAIL DIGEST
// Cronjob trigger pengiriman digest email harian
// =========================================================
function sendDailyEmailDigest() {
  processDailyEmailNotifications();
}

// =========================================================
// FUNGSI DEBUG DENGAN TEST MAIL DIRECT
// =========================================================
function testEmailDebug() {
  Logger.log("=== MULAI DEBUG EMAIL SERVICE ===");
  
  var testEmailRecipient = "msi.atlasadjusting@gmail.com";
  Logger.log("Mencoba kirim email percobaan langsung ke: " + testEmailRecipient);
  
  try {
    MailApp.sendEmail({
      to: testEmailRecipient,
      subject: "[TESTING PORTAL] Tes Pengiriman Email Portal MSI",
      htmlBody: "<h3 style='color:green;'>Berhasil! Sistem email Apps Script bekerja normal.</h3>"
    });
    Logger.log("SUCCESS: Email test berhasil terkirim!");
  } catch (e) {
    Logger.log("ERROR GAGAL KIRIM EMAIL: " + e.message);
  }
  
  Logger.log("=== SELESAI DEBUG ===");
}