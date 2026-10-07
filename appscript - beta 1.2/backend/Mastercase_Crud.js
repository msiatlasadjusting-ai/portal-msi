// =========================================================
// MASTERCASE CRUD & LIFECYCLE MANAGEMENT
// Modul pengelolaan operasi CRUD data Mastercase Sheet, sinkronisasi milestone stage, dan folder Drive
// =========================================================

// =========================================================
// GET MASTERCASE SHEET
// Mengambil instance Sheet Mastercase dari Spreadsheet
// =========================================================
function getMastercaseSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan!");
  }
  return sheet;
}

// =========================================================
// PARSE MONEY
// Helper konversi nilai uang ke format numerik murni
// =========================================================
function parseMoney(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  const cleanStr = String(val).replace(/\./g, '').replace(/,/g, '').replace(/[^0-9\-]/g, '');
  const num = Number(cleanStr);
  return isNaN(num) ? 0 : num;
}

// =========================================================
// GET NEXT JOB NUMBER
// Menghitung nomor Job berikutnya secara otomatis berbasis sequence numerik
// =========================================================
function getNextJobNumber() {
  try {
    const sheet = getMastercaseSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return 8001;
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let maxNo = 8000;
    for (let i = 0; i < values.length; i++) {
      const val = parseInt(values[i][0], 10);
      if (!isNaN(val) && val > maxNo) {
        maxNo = val;
      }
    }
    return maxNo >= 8001 ? maxNo + 1 : 8001;
  } catch (e) {
    return 8001;
  }
}

// =========================================================
// CREATE JOB DRIVE FOLDER
// Membuat folder Google Drive untuk penampung berkas case baru
// =========================================================
function createJobDriveFolder(refMsi) {
  try {
    if (!DRIVE_PARENT_FOLDER_ID || DRIVE_PARENT_FOLDER_ID === "GANTI_DENGAN_FOLDER_ID_KAMU") {
      console.warn("Folder ID induk Google Drive belum diset.");
      return null;
    }

    const parentFolder = DriveApp.getFolderById(DRIVE_PARENT_FOLDER_ID);
    const folderName = String(refMsi || "Mastercase_New").replace(/[\:\*\?\"\\\,\<\>\|]/g, "").trim();

    const existingFolders = parentFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      return existingFolders.next().getUrl();
    }
    
    const newFolder = parentFolder.createFolder(folderName);
    return newFolder.getUrl();
  } catch (error) {
    console.error("Gagal membuat folder Google Drive:", error);
    return null;
  }
}

// =========================================================
// GET DROPDOWN MASTER DATA
// Mengambil opsi referensi PIC dan Agent/Broker dari Sheet pendukung
// =========================================================
function getDropdownMasterData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let picList = [];
    const picSheet = ss.getSheetByName("PIC");
    if (picSheet && picSheet.getLastRow() > 1) {
      const picValues = picSheet.getRange(2, 1, picSheet.getLastRow() - 1, 2).getValues();
      picList = picValues.map(row => {
        const nama = String(row[0] || "").trim();
        const kode = String(row[1] || "").trim();
        const formattedLabel = (kode && nama) ? `${kode} - ${nama}` : (kode || nama);
        return {
          code: kode || nama,
          name: formattedLabel
        };
      }).filter(item => item.code !== "");
    }

    let agentList = [];
    const agentSheet = ss.getSheetByName("Agent") || ss.getSheetByName("Broker");
    if (agentSheet && agentSheet.getLastRow() > 1) {
      const agentCols = agentSheet.getLastColumn();
      const agentValues = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentCols).getValues();
      agentList = agentValues.map(row => {
        if (row.length > 1 && !isNaN(Number(row[0])) && row[1]) {
          return String(row[1]).trim();
        }
        return String(row[0]).trim();
      }).filter(val => val !== "");
    }

    return {
      success: true,
      pic: picList,
      agent: agentList
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// APPLY CURRENCY FORMATS
// Menerapkan format mata uang Rupiah pada kolom keuangan di baris sheet
// =========================================================
function applyCurrencyFormats(sheet, rowIdx) {
  try {
    const currencyFormat = '"Rp "#,##0;("Rp "#,##0);"-"';
    sheet.getRange(rowIdx, 35, 1, 8).setNumberFormat(currencyFormat);
    sheet.getRange(rowIdx, 47).setNumberFormat(currencyFormat);
    sheet.getRange(rowIdx, 50, 1, 4).setNumberFormat(currencyFormat);
  } catch (e) {
    console.error("Gagal apply format currency:", e);
  }
}

// =========================================================
// GET MASTERCASE DATA
// Membaca seluruh data register Mastercase dan menggabungkannya dengan milestone stage
// =========================================================
function getMastercaseData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getMastercaseSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 55);

    if (lastRow <= 1) {
      return { success: true, data: [], nextJobNumber: 8001 };
    }

    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    const CANONICAL_KEYS = [
      "Job Number", 
      "Instruction Date", 
      "Lead PIC", 
      "Ref MSI", 
      "Ref Client",
      "Type of Survey", 
      "LOB", "Client", 
      "Broker / Agent", 
      "Owner / Client's Customer",
      "Location", 
      "Object / Vessel / Unit", 
      "Accident", 
      "DOL", 
      "Survey Date",
      "Survey Start Date", 
      "Survey Finish Date", 
      "Survey PIC Handler", 
      "Survey Remarks",
      "Report Start Date", 
      "Report Sent Date", 
      "Report PIC Handler", 
      "Report Remarks",
      "Location Client", 
      "LOP Start Date", 
      "LOP Approved Date", 
      "LOP PIC Handler",
      "LOP Remarks", 
      "PIC Client", 
      "Invoice Date", 
      "Paid Date", 
      "Invoice PIC Handler",
      "Invoice Remarks", 
      "Remarks FAT", 
      "Fee (IDR)", 
      "TAX Basis (IDR)", 
      "IOU (IDR)",
      "Expenses Act. (IDR)", 
      "Admin Cost (IDR)", 
      "Expenses", "VAT 11% (IDR)",
      "Invoice Amount (IDR)", 
      "Comm.", 
      "Load", 
      "Nomor Jurnal Load", 
      "No. Faktur Pajak",
      "PPH 23 (IDR)", 
      "Nomor Bukti Potong PPh 23", 
      "Tanggal Bukti Potong PPh 23",
      "Invoice Payment (net PPH23)", 
      "Client Payment", "Selisih", 
      "Reimburse Settled (+/-)",
      "Status", 
      "Folder"
    ];

    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(allData.length, 5); r++) {
      const rowStr = allData[r].map(c => String(c).toLowerCase()).join(" ");
      if (rowStr.includes("job") || rowStr.includes("instruction") || rowStr.includes("pic")) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = allData[headerRowIdx];
    const cleanedHeaders = headerRow.map(h =>
      String(h).replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim()
    );

    let milestoneMap = {};
    try {
      const milestoneSheet = ss.getSheetByName("Job_Milestones");
      if (milestoneSheet && milestoneSheet.getLastRow() > 1) {
        // Membaca 11 kolom (A1:K)
        const mData = milestoneSheet.getRange(2, 1, milestoneSheet.getLastRow() - 1, 11).getValues();
        mData.forEach(mRow => {
          const jobNo = String(mRow[1] || "").trim();
          if (!jobNo) return;
          if (!milestoneMap[jobNo]) milestoneMap[jobNo] = [];
          
          let sDate = mRow[4];
          if (sDate instanceof Date) sDate = Utilities.formatDate(sDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
          
          let tDate = mRow[5];
          if (tDate instanceof Date) tDate = Utilities.formatDate(tDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
          
          let fDate = mRow[6];
          if (fDate instanceof Date) fDate = Utilities.formatDate(fDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");

          milestoneMap[jobNo].push({
            milestoneId: String(mRow[0] || ""),
            jobNumber: jobNo,
            stageCode: String(mRow[2] || ""),
            stageName: String(mRow[3] || ""),
            startDate: sDate ? String(sDate) : "",
            targetDate: tDate ? String(tDate) : "",
            finishDate: fDate ? String(fDate) : "",
            duration: mRow[7] !== "" && mRow[7] !== null && !isNaN(Number(mRow[7])) ? Number(mRow[7]) : 0,
            pic: String(mRow[8] || ""),
            status: String(mRow[9] || ""),
            remarks: String(mRow[10] || "")
          });
        });
      }
    } catch (e) {
      console.error("Gagal membaca Job_Milestones:", e);
    }

    const rawRows = allData.slice(headerRowIdx + 1);
    const result = [];

    rawRows.forEach((row, idx) => {
      const hasAnyData = row.some(cell => cell !== "" && cell !== null && cell !== undefined);
      if (!hasAnyData) return;

      const actualRowIndex = headerRowIdx + 2 + idx;
      let rowObj = { _rowIndex: actualRowIndex };

      CANONICAL_KEYS.forEach((canonKey, colIdx) => {
        let value = row[colIdx];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
        } else if (value === null || value === undefined) {
          value = "";
        }
        rowObj[canonKey] = value;
      });

      cleanedHeaders.forEach((header, colIdx) => {
        if (header && colIdx < row.length) {
          let value = row[colIdx];
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
          } else if (value === null || value === undefined) {
            value = "";
          }
          rowObj[header] = value;
        }
      });

      const currentJobNo = String(rowObj["Job Number"] || "");
      rowObj.milestones = milestoneMap[currentJobNo] || [];

      result.push(rowObj);
    });

    const nextJobNo = getNextJobNumber();
    return { success: true, data: result, nextJobNumber: nextJobNo };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// AUTO DETECT STATUS
// Menentukan status milestone berdasarkan kelengkapan start dan finish date
// =========================================================
function autoDetectStatus(startDate, finishDate, isStage1) {
  if (isStage1) return "Completed";
  if (finishDate && String(finishDate).trim() !== "" && String(finishDate) !== "-") return "Completed";
  if (startDate && String(startDate).trim() !== "" && String(startDate) !== "-") return "In Progress";
  return "Pending";
}

// =========================================================
// CALCULATE DAYS BETWEEN
// Menghitung selisih hari kalender antara dua tanggal
// =========================================================
function calculateDaysBetween(d1Input, d2Input) {
  try {
    if (!d1Input || !d2Input || d1Input === "-" || d2Input === "-") return 0;
    
    let d1 = (d1Input instanceof Date) ? d1Input : new Date(d1Input);
    let d2 = (d2Input instanceof Date) ? d2Input : new Date(d2Input);

    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return 0;

    d1.setHours(0,0,0,0);
    d2.setHours(0,0,0,0);

    const diffTime = d2.getTime() - d1.getTime();
    const days = Math.round(diffTime / (1000 * 60 * 60 * 24));
    return isNaN(days) ? 0 : (days < 0 ? 0 : days);
  } catch (e) {
    return 0;
  }
}

// =========================================================
// INIT JOB MILESTONES
// Menginisialisasi 5 Milestone tahap penugasan di Sheet Job_Milestones
// =========================================================
function initJobMilestones(ss, jobNumber, formData) {
  let milestoneSheet = ss.getSheetByName("Job_Milestones");
  if (!milestoneSheet) {
    milestoneSheet = ss.insertSheet("Job_Milestones");
    milestoneSheet.appendRow([
      "Milestone ID", "Job Number", "Stage Code", "Stage Name", 
      "Start Date", "Target Date", "Finish Date", "Duration (Days)", 
      "PIC Handler", "Status Stage", "Remarks / Evidence"
    ]);
  }

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
  const instructionDate = formData.instructionDate || todayStr;
  const leadPic = formData.pic || "Admin";

  const st1Start = instructionDate;
  const st1Target = instructionDate;
  const st1Finish = instructionDate;

  const st2Start = formData.surveyStartDate || formData.surveyDate || "";
  const st2Finish = formData.surveyFinishDate || "";
  const st2Target = st2Finish || (st2Start ? st2Start : "");
  const st2Pic = formData.stage2PicHandler || formData.surveyPicHandler || leadPic;
  const st2Remarks = (formData.remarks || formData.surveyRemarks || formData.initialMilestoneRemarks || "").trim();

  const st3Start = formData.reportPreparationStartDate || formData.reportStartDate || st2Finish;
  const st3Finish = formData.surveyReportSentOn || formData.reportSentDate || "";
  const st3Target = st3Finish || (st2Finish ? addDaysHelper(st2Finish, 5) : "");
  const st3Pic = formData.stage3PicHandler || formData.reportPicHandler || leadPic;
  const st3Remarks = (formData.stage3Remarks || formData.reportRemarks || "").trim();

  const st4Start = formData.lopStartDate || st3Finish;
  const st4Finish = formData.lopApprovedDate || formData.invoiceDate || "";
  const st4Target = st4Finish || (st3Finish ? addDaysHelper(st3Finish, 3) : "");
  const st4Pic = formData.stage4PicHandler || formData.lopPicHandler || leadPic;
  const st4Remarks = (formData.stage4Remarks || formData.lopRemarks || "").trim();

  const st5Start = formData.invoiceDate || "";
  const st5Finish = formData.paidOn || formData.paidDate || "";
  const st5Target = st5Finish || (st5Start ? addDaysHelper(st5Start, 7) : "");
  const st5Pic = formData.stage5PicHandler || formData.invoicePicHandler || "Finance";
  const st5Remarks = (formData.stage5Remarks || formData.invoiceRemarks || formData.remarksFat || "").trim();

  const defaultStages = [
    {
      code: "STAGE_1",
      name: "Assignment & Insured Contact",
      startDate: st1Start,
      targetDate: st1Target,
      finishDate: st1Finish,
      status: "Completed",
      pic: leadPic,
      remarks: "Surat penugasan diterima & tercatat."
    },
    {
      code: "STAGE_2",
      name: "Site Inspection & Joint Survey",
      startDate: st2Start,
      targetDate: st2Target,
      finishDate: st2Finish,
      status: autoDetectStatus(st2Start, st2Finish, false),
      pic: st2Pic,
      remarks: st2Remarks
    },
    {
      code: "STAGE_3",
      name: "Interim Report & Cause Analysis",
      startDate: st3Start,
      targetDate: st3Target,
      finishDate: st3Finish,
      status: autoDetectStatus(st3Start, st3Finish, false),
      pic: st3Pic,
      remarks: st3Remarks
    },
    {
      code: "STAGE_4",
      name: "Final Adjustment & LOP",
      startDate: st4Start,
      targetDate: st4Target,
      finishDate: st4Finish,
      status: autoDetectStatus(st4Start, st4Finish, false),
      pic: st4Pic,
      remarks: st4Remarks
    },
    {
      code: "STAGE_5",
      name: "Fee Invoicing & Settlement",
      startDate: st5Start,
      targetDate: st5Target,
      finishDate: st5Finish,
      status: autoDetectStatus(st5Start, st5Finish, false),
      pic: st5Pic,
      remarks: st5Remarks
    }
  ];

  defaultStages.forEach((stage, idx) => {
    const milestoneId = "M-" + jobNumber + "-" + (idx + 1);
    const duration = (stage.startDate && stage.finishDate) ? calculateDaysBetween(stage.startDate, stage.finishDate) : 0;
    
    milestoneSheet.appendRow([
      milestoneId,         // A: Milestone ID
      jobNumber,           // B: Job Number
      stage.code,          // C: Stage Code
      stage.name,          // D: Stage Name
      stage.startDate,     // E: Start Date
      stage.targetDate,    // F: Target Date
      stage.finishDate,    // G: Finish Date
      duration,            // H: Duration (Days)
      stage.pic,           // I: PIC Handler
      stage.status,        // J: Status Stage
      stage.remarks        // K: Remarks / Evidence
    ]);
  });
}

function addDaysHelper(dateStr, days) {
  try {
    if (!dateStr || dateStr === "-") return "";
    let d = new Date(dateStr);
    if (isNaN(d.getTime())) return "";
    d.setDate(d.getDate() + days);
    return Utilities.formatDate(d, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
  } catch (e) {
    return "";
  }
}

// =========================================================
// SYNC MILESTONE DATES
// Memperbarui milestone stage yang bersesuaian saat form diubah
// =========================================================
function syncMilestoneDates(ss, jobNumber, formData) {
  const milestoneSheet = ss.getSheetByName("Job_Milestones");
  if (!milestoneSheet) {
    initJobMilestones(ss, jobNumber, formData);
    return;
  }

  const data = milestoneSheet.getDataRange().getValues();
  let foundRows = false;
  const leadPic = formData.pic || "Admin";

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(jobNumber)) {
      foundRows = true;
      const stageCode = String(data[i][2]);
      let startDate = "";
      let targetDate = "";
      let finishDate = "";
      let status = "Pending";
      let picHandler = leadPic;
      let remarks = "";

      if (stageCode === "STAGE_1") {
        startDate = formData.instructionDate || "";
        targetDate = formData.instructionDate || "";
        finishDate = formData.instructionDate || "";
        status = "Completed";
        picHandler = leadPic;
        remarks = "Surat penugasan diterima & tercatat.";
      } else if (stageCode === "STAGE_2") {
        startDate = formData.surveyStartDate || formData.surveyDate || "";
        finishDate = formData.surveyFinishDate || "";
        targetDate = finishDate || startDate;
        status = autoDetectStatus(startDate, finishDate, false);
        picHandler = formData.stage2PicHandler || formData.surveyPicHandler || leadPic;
        remarks = (formData.remarks || formData.surveyRemarks || formData.initialMilestoneRemarks || "").trim();
      } else if (stageCode === "STAGE_3") {
        startDate = formData.reportPreparationStartDate || formData.reportStartDate || formData.surveyFinishDate || ""; 
        finishDate = formData.surveyReportSentOn || formData.reportSentDate || "";
        targetDate = finishDate || (formData.surveyFinishDate ? addDaysHelper(formData.surveyFinishDate, 5) : "");
        status = autoDetectStatus(startDate, finishDate, false);
        picHandler = formData.stage3PicHandler || formData.reportPicHandler || leadPic;
        remarks = (formData.stage3Remarks || formData.reportRemarks || "").trim();
      } else if (stageCode === "STAGE_4") {
        startDate = formData.lopStartDate || formData.surveyReportSentOn || formData.reportSentDate || ""; 
        finishDate = formData.lopApprovedDate || formData.invoiceDate || "";
        targetDate = finishDate || (formData.surveyReportSentOn ? addDaysHelper(formData.surveyReportSentOn, 3) : "");
        status = autoDetectStatus(startDate, finishDate, false);
        picHandler = formData.stage4PicHandler || formData.lopPicHandler || leadPic;
        remarks = (formData.stage4Remarks || formData.lopRemarks || "").trim();
      } else if (stageCode === "STAGE_5") {
        startDate = formData.invoiceDate || "";
        finishDate = formData.paidOn || formData.paidDate || "";
        targetDate = finishDate || (formData.invoiceDate ? addDaysHelper(formData.invoiceDate, 7) : "");
        status = autoDetectStatus(startDate, finishDate, false);
        picHandler = formData.stage5PicHandler || formData.invoicePicHandler || "Finance";
        remarks = (formData.stage5Remarks || formData.invoiceRemarks || formData.remarksFat || "").trim();
      }

      if (startDate instanceof Date) startDate = Utilities.formatDate(startDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
      if (targetDate instanceof Date) targetDate = Utilities.formatDate(targetDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
      if (finishDate instanceof Date) finishDate = Utilities.formatDate(finishDate, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");

      const duration = (startDate && finishDate && startDate !== "-" && finishDate !== "-") 
        ? calculateDaysBetween(startDate, finishDate) 
        : 0;

      milestoneSheet.getRange(i + 1, 5).setValue(startDate);    // E: Start Date
      milestoneSheet.getRange(i + 1, 6).setValue(targetDate);   // F: Target Date
      milestoneSheet.getRange(i + 1, 7).setValue(finishDate);   // G: Finish Date
      milestoneSheet.getRange(i + 1, 8).setValue(duration);     // H: Duration (Days)
      milestoneSheet.getRange(i + 1, 9).setValue(picHandler);   // I: PIC Handler
      milestoneSheet.getRange(i + 1, 10).setValue(status);      // J: Status Stage
      milestoneSheet.getRange(i + 1, 11).setValue(remarks);     // K: Remarks / Evidence
    }
  }

  if (!foundRows) {
    initJobMilestones(ss, jobNumber, formData);
  }
}
// =========================================================
// CALCULATE MASTERCASE STATUS
// Menghitung status global penugasan Mastercase berdasarkan field stage
// =========================================================
function calculateMastercaseStatus(formData) {
  if (formData.status === "Cancelled" || formData.jobStatus === "Cancelled") {
    return "Cancelled";
  }

  if (formData.paidOn && String(formData.paidOn).trim() !== "" && String(formData.paidOn) !== "-") {
    return "Invoice Paid";
  }
  if (formData.invoiceDate && String(formData.invoiceDate).trim() !== "" && String(formData.invoiceDate) !== "-") {
    return "Invoice Sent";
  }

  if ((formData.lopApprovedDate && String(formData.lopApprovedDate).trim() !== "" && String(formData.lopApprovedDate) !== "-") ||
      (formData.lopStartDate && String(formData.lopStartDate).trim() !== "" && String(formData.lopStartDate) !== "-")) {
    return "Awaiting Invoice";
  }

  if (formData.surveyReportSentOn && String(formData.surveyReportSentOn).trim() !== "" && String(formData.surveyReportSentOn) !== "-") {
    return "Report Sent";
  }
  if (formData.reportPreparationStartDate && String(formData.reportPreparationStartDate).trim() !== "" && String(formData.reportPreparationStartDate) !== "-") {
    return "Drafting Report";
  }

  if ((formData.surveyStartDate && String(formData.surveyStartDate).trim() !== "" && String(formData.surveyStartDate) !== "-") ||
      (formData.surveyDate && String(formData.surveyDate).trim() !== "" && String(formData.surveyDate) !== "-") ||
      (formData.surveyFinishDate && String(formData.surveyFinishDate).trim() !== "" && String(formData.surveyFinishDate) !== "-")) {
    return "In Site Survey";
  }

  return "On Process";
}

// =========================================================
// CREATE MASTERCASE
// Menambahkan record Mastercase baru ke Sheet dan membuatkan folder Drive & Milestone
// =========================================================
function createMastercase(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getMastercaseSheet();

    if (!formData.instructionDate || !formData.pic || !formData.lob) {
      throw new Error("Gagal menyimpan: Instruction Date, Lead PIC, dan LOB wajib diisi!");
    }

    const nextJobNo = getNextJobNumber();
    const jobNumber = formData.jobNumber ? parseInt(formData.jobNumber, 10) : nextJobNo;

    const fee = parseMoney(formData.feeIdr);
    const iou = parseMoney(formData.iouIdr);
    const expAct = parseMoney(formData.expensesAct);
    const adminCost = parseMoney(formData.adminCost);
    const expenses = parseMoney(formData.expenses);
    const load = parseMoney(formData.load);
    const clientPayment = parseMoney(formData.clientPayment);

    const taxBasis = Math.round(fee * (11 / 12));
    const vat11 = Math.round(fee * 0.11);
    const invoiceAmount = fee + expenses + vat11;
    const pph23 = Math.round(fee * 0.02);
    const netPayment = invoiceAmount - pph23;
    const selisih = clientPayment - netPayment;
    const reimburseSettled = iou - expAct;

    const refMsi = formData.refMsi && formData.refMsi !== "[ Preview Ref MSI ]" 
      ? formData.refMsi 
      : String(jobNumber);
    const driveFolderUrl = createJobDriveFolder(refMsi) || "";
    const surveyDate = formData.surveyStartDate || formData.surveyDate || "";
    const autoStatus = calculateMastercaseStatus(formData);

    const newRow = [
      jobNumber,
      formData.instructionDate || "",
      formData.pic || "",
      refMsi,
      formData.refClient || "",
      formData.typeOfSurvey || "",
      formData.lob || "",
      formData.client || "",
      formData.broker || "",
      formData.ownerCustomer || "",
      formData.location || "",
      formData.objectVesselUnit || "",
      formData.accident || "",
      formData.dol || "",
      surveyDate || "",
      formData.surveyStartDate || "",
      formData.surveyFinishDate || "",
      formData.stage2PicHandler || formData.surveyPicHandler || formData.pic || "",
      formData.remarks || formData.surveyRemarks || "",
      formData.reportPreparationStartDate || formData.reportStartDate || formData.surveyFinishDate || "",
      formData.surveyReportSentOn || formData.reportSentDate || "",
      formData.stage3PicHandler || formData.reportPicHandler || formData.pic || "",
      formData.stage3Remarks || formData.reportRemarks || "",
      formData.clientOfficeLocation || formData.locationClient || "",
      formData.lopStartDate || formData.surveyReportSentOn || "",
      formData.lopApprovedDate || "",
      formData.stage4PicHandler || formData.lopPicHandler || formData.pic || "",
      formData.stage4Remarks || formData.lopRemarks || "",
      formData.clientPicPerson || formData.picClient || "",
      formData.invoiceDate || "",
      formData.paidOn || formData.paidDate || "",
      formData.stage5PicHandler || formData.invoicePicHandler || "Finance",
      formData.stage5Remarks || formData.invoiceRemarks || "",
      formData.remarksFat || "",
      fee,
      taxBasis,
      iou,
      expAct,
      adminCost,
      expenses,
      vat11,
      invoiceAmount,
      formData.comm || "",
      load,
      formData.nomorJurnalLoad || "",
      formData.noFakturPajak || "",
      pph23,
      formData.nomorBuktiPotongPph23 || "",
      formData.tanggalBuktiPotongPph23 || "",
      netPayment,
      clientPayment,
      selisih,
      reimburseSettled,
      autoStatus,
      driveFolderUrl
    ];

    sheet.appendRow(newRow);
    const newRowIdx = sheet.getLastRow();
    applyCurrencyFormats(sheet, newRowIdx);

    initJobMilestones(ss, jobNumber, formData);

    return {
      success: true,
      message: "Mastercase #" + jobNumber + " berhasil ditambahkan beserta 5 Milestone awal!",
      jobNumber: jobNumber,
      nextJobNumber: getNextJobNumber()
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// UPDATE MASTERCASE
// Memperbarui record Mastercase yang sudah ada pada baris spesifik
// =========================================================
function updateMastercase(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getMastercaseSheet();

    const rowIndex = parseInt(formData.rowIndex, 10);
    if (!rowIndex || rowIndex < 2) {
      throw new Error("Row index tidak valid untuk update!");
    }

    const jobNumber = formData.jobNumber;

    const fee = parseMoney(formData.feeIdr);
    const iou = parseMoney(formData.iouIdr);
    const expAct = parseMoney(formData.expensesAct);
    const adminCost = parseMoney(formData.adminCost);
    const expenses = parseMoney(formData.expenses);
    const load = parseMoney(formData.load);
    const clientPayment = parseMoney(formData.clientPayment);

    const taxBasis = Math.round(fee * (11 / 12));
    const vat11 = Math.round(fee * 0.11);
    const invoiceAmount = fee + expenses + vat11;
    const pph23 = Math.round(fee * 0.02);
    const netPayment = invoiceAmount - pph23;
    const selisih = clientPayment - netPayment;
    const reimburseSettled = iou - expAct;

    const refMsi = formData.refMsi || String(jobNumber);
    const surveyDate = formData.surveyStartDate || formData.surveyDate || "";
    const autoStatus = calculateMastercaseStatus(formData);
    const existingFolderUrl = sheet.getRange(rowIndex, 55).getValue() || createJobDriveFolder(refMsi) || "";

    const updatedRow = [
      jobNumber,
      formData.instructionDate || "",
      formData.pic || "",
      refMsi,
      formData.refClient || "",
      formData.typeOfSurvey || "",
      formData.lob || "",
      formData.client || "",
      formData.broker || "",
      formData.ownerCustomer || "",
      formData.location || "",
      formData.objectVesselUnit || "",
      formData.accident || "",
      formData.dol || "",
      surveyDate || "",
      formData.surveyStartDate || "",
      formData.surveyFinishDate || "",
      formData.stage2PicHandler || formData.surveyPicHandler || formData.pic || "",
      formData.remarks || formData.surveyRemarks || "",
      formData.reportPreparationStartDate || formData.reportStartDate || formData.surveyFinishDate || "",
      formData.surveyReportSentOn || formData.reportSentDate || "",
      formData.stage3PicHandler || formData.reportPicHandler || formData.pic || "",
      formData.stage3Remarks || formData.reportRemarks || "",
      formData.clientOfficeLocation || formData.locationClient || "",
      formData.lopStartDate || formData.surveyReportSentOn || "",
      formData.lopApprovedDate || "",
      formData.stage4PicHandler || formData.lopPicHandler || formData.pic || "",
      formData.stage4Remarks || formData.lopRemarks || "",
      formData.clientPicPerson || formData.picClient || "",
      formData.invoiceDate || "",
      formData.paidOn || formData.paidDate || "",
      formData.stage5PicHandler || formData.invoicePicHandler || "Finance",
      formData.stage5Remarks || formData.invoiceRemarks || "",
      formData.remarksFat || "",
      fee,
      taxBasis,
      iou,
      expAct,
      adminCost,
      expenses,
      vat11,
      invoiceAmount,
      formData.comm || "",
      load,
      formData.nomorJurnalLoad || "",
      formData.noFakturPajak || "",
      pph23,
      formData.nomorBuktiPotongPph23 || "",
      formData.tanggalBuktiPotongPph23 || "",
      netPayment,
      clientPayment,
      selisih,
      reimburseSettled,
      autoStatus,
      existingFolderUrl
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    applyCurrencyFormats(sheet, rowIndex);

    syncMilestoneDates(ss, jobNumber, formData);

    return {
      success: true,
      message: "Mastercase #" + jobNumber + " berhasil diperbarui!"
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// DELETE MASTERCASE
// Menghapus baris Mastercase, folder Drive terkait, catatan milestone, dan notifikasi
// =========================================================
function deleteMastercase(rowIndex) {
  try {
    const sheet = getMastercaseSheet();
    const ss = sheet.getParent();
    const rIdx = parseInt(rowIndex, 10);
    if (isNaN(rIdx) || rIdx < 2) throw new Error("Row index tidak valid.");

    const jobNo = String(sheet.getRange(rIdx, 1).getValue()).trim();
    const refMsi = String(sheet.getRange(rIdx, 4).getValue()).trim();
    const driveFolderUrl = String(sheet.getRange(rIdx, 55).getValue()).trim();

    deleteJobDriveFolder(driveFolderUrl, refMsi);

    sheet.deleteRow(rIdx);

    if (jobNo) {
      deleteJobMilestonesByJobNo(ss, jobNo);
      deleteJobNotificationsByJobNo(ss, jobNo); // <-- TAMBAHAN BARU
    }

    return { 
      success: true, 
      message: "Mastercase #" + jobNo + " beserta Folder Drive, Milestones & Notifikasi berhasil dihapus!", 
      nextJobNumber: getNextJobNumber() 
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// DELETE JOB DRIVE FOLDER
// Memindahkan folder Google Drive milik job ke Trash
// =========================================================
function deleteJobDriveFolder(folderUrl, refMsi) {
  try {
    if (folderUrl && folderUrl.indexOf("drive.google.com") !== -1) {
      const match = folderUrl.match(/[-\w]{25,}/);
      if (match && match[0]) {
        const folder = DriveApp.getFolderById(match[0]);
        folder.setTrashed(true);
        return;
      }
    }

    if (refMsi && refMsi !== "-") {
      const parentFolderId = DRIVE_PARENT_FOLDER_ID;
      const folderName = String(refMsi).replace(/[\/\:\*\?\"\\\,\<\>\|]/g, "-");
      const parentFolder = DriveApp.getFolderById(parentFolderId);
      const folders = parentFolder.getFoldersByName(folderName);
      
      while (folders.hasNext()) {
        const folder = folders.next();
        folder.setTrashed(true);
      }
    }
  } catch (e) {
    console.error("Gagal menghapus folder Google Drive untuk Ref MSI " + refMsi + ":", e);
  }
}

// =========================================================
// DELETE JOB MILESTONES BY JOB NO
// Menghapus seluruh baris milestone milik nomor job tertentu di Sheet Job_Milestones
// =========================================================
function deleteJobMilestonesByJobNo(ss, jobNo) {
  try {
    const milestoneSheet = ss.getSheetByName("Job_Milestones");
    if (!milestoneSheet) return;

    const lastRow = milestoneSheet.getLastRow();
    if (lastRow <= 1) return;

    const data = milestoneSheet.getRange(1, 1, lastRow, 2).getValues(); 

    for (let i = data.length - 1; i >= 1; i--) {
      const currentJobNo = String(data[i][1]).trim(); 
      if (currentJobNo === String(jobNo).trim()) {
        milestoneSheet.deleteRow(i + 1);
      }
    }
  } catch (e) {
    console.error("Gagal menghapus Job_Milestones untuk Job #" + jobNo + ":", e);
  }
}

// =========================================================
// DELETE JOB NOTIFICATIONS BY JOB NO
// Menghapus seluruh baris notifikasi milik nomor job tertentu di Sheet Notifications
// =========================================================
function deleteJobNotificationsByJobNo(ss, jobNo) {
  try {
    const notifSheet = ss.getSheetByName("Notifications");
    if (!notifSheet) return;

    const lastRow = notifSheet.getLastRow();
    if (lastRow <= 1) return;

    const data = notifSheet.getRange(1, 2, lastRow, 1).getValues(); 

    for (let i = data.length - 1; i >= 1; i--) {
      const currentJobNo = String(data[i][0] || '').trim(); 
      if (currentJobNo === String(jobNo).trim()) {
        notifSheet.deleteRow(i + 1);
      }
    }
  } catch (e) {
    console.error("Gagal menghapus Notifications untuk Job #" + jobNo + ":", e);
  }
}
