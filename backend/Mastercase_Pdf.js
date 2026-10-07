// =========================================================
// MASTERCASE PDF GENERATOR
// Pembuatan dokumen PDF ringkasan case dari template Google Docs & pengiriman lampiran email
// =========================================================

function formatRupiahGS(num) {
  if (num === null || num === undefined || num === '' || isNaN(Number(num))) {
    return '0';
  }
  var n = Number(num);
  var isNegative = n < 0;
  var absolute = Math.abs(Math.round(n));
  var formatted = absolute.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return isNegative ? '-' + formatted : formatted;
}

function generateMastercasePDF(jobData) {
  try {
    var templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
    var targetFolder = DriveApp.getFolderById(PDF_FOLDER_ID);

    var refMsiValue = jobData.refMsi || jobData.jobNo || "Detail";
    var pdfFileName = "Invoice_" + refMsiValue.replace(/[\/\:\*\?\"\\\,\<\>\|]/g, "-") + ".pdf";
    var tempDocName = "TEMP_INV_" + (jobData.jobNo || "Detail") + "_" + new Date().getTime();

    var copiedFile = templateFile.makeCopy(tempDocName, targetFolder);
    var doc = DocumentApp.openById(copiedFile.getId());
    var body = doc.getBody();

    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd MMMM yyyy");
    
    body.replaceText("{{PRINT_DATE}}", todayStr);
    body.replaceText("{{REF_MSI}}", jobData.refMsi || "-");
    body.replaceText("{{REF_CLIENT}}", jobData.refClient || "-");
    body.replaceText("{{CLIENT}}", jobData.client || "-");
    var locClientVal = jobData.clientOfficeLocation || jobData.locationClient || "";
    var picClientVal = jobData.clientPicPerson || jobData.picClient || "";
    body.replaceText("{{Location Client}}", locClientVal.trim() !== "" ? locClientVal : "-");
    body.replaceText("{{PIC number}}", picClientVal.trim() !== "" ? picClientVal : "-");

    body.replaceText("{{SURVEY_TYPE}}", jobData.surveyType || "-");
    body.replaceText("{{LOB}}", jobData.lob || "-");
    
    body.replaceText("{{INSURED}}", jobData.owner || "-"); 
    
    body.replaceText("{{DOL}}", jobData.dol || "-");
    body.replaceText("{{LOCATION}}", jobData.location || "-");
    body.replaceText("{{OBJECT_UNIT}}", jobData.objectUnit || "-");

    body.replaceText("{{FEE_SURVEY}}", formatRupiahGS(jobData.fee));
    body.replaceText("{{TAX_BASIS}}", formatRupiahGS(jobData.taxBasis));
    body.replaceText("{{VAT_11}}", formatRupiahGS(jobData.vat11));
    body.replaceText("{{TOTAL_EXPENSES}}", formatRupiahGS(jobData.expenses));
    body.replaceText("{{INVOICE_AMOUNT}}", formatRupiahGS(jobData.invoiceAmount));

    doc.saveAndClose();

    var pdfBlob = copiedFile.getAs(MimeType.PDF).setName(pdfFileName);
    var pdfFile = targetFolder.createFile(pdfBlob);

    DriveApp.getFileById(copiedFile.getId()).setTrashed(true);

    var recipientEmail = jobData.userEmail || "";

    if (recipientEmail) {
      var emailSubject = "[Invoice MSI] " + (jobData.refMsi || "#" + jobData.jobNo) + " - " + (jobData.client || "Client");

      var emailBodyHtml = `
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #333; max-width: 600px; line-height: 1.6;">
          <h2 style="color: #EC1F27; margin-bottom: 5px;">PT Mahardhika Surveyor Indonesia</h2>
          <p style="font-size: 11px; color: #777; margin-top: 0;">Official Invoice Notification</p>
          <hr style="border: none; border-top: 1px solid #eee; margin: 15px 0;">
          
          <p>Halo <strong>${jobData.userName || "User"}</strong>,</p>
          <p>Dokumen Invoice untuk penugasan berikut telah berhasil dicetak dan dilampirkan:</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 12px;">
            <tr style="background-color: #f8f9fa;">
              <td style="padding: 8px; font-weight: bold; width: 140px; border: 1px solid #e9ecef;">Our Ref (Ref MSI):</td>
              <td style="padding: 8px; border: 1px solid #e9ecef; font-weight: bold; color: #EC1F27;">${jobData.refMsi || "-"}</td>
            </tr>
            <tr>
              <td style="padding: 8px; font-weight: bold; border: 1px solid #e9ecef;">Client Name:</td>
              <td style="padding: 8px; border: 1px solid #e9ecef;">${jobData.client || "-"}</td>
            </tr>
            <tr style="background-color: #f8f9fa;">
              <td style="padding: 8px; font-weight: bold; border: 1px solid #e9ecef;">Total Invoice:</td>
              <td style="padding: 8px; border: 1px solid #e9ecef; font-weight: bold;">IDR ${formatRupiahGS(jobData.invoiceAmount)}</td>
            </tr>
          </table>

          <p>Dokumen PDF terlampir juga telah tersimpan otomatis di Google Drive portal MSI.</p>
          
          <div style="margin-top: 25px; padding-top: 15px; border-top: 1px solid #eee; font-size: 11px; color: #888;">
            <p>© 2026 PT Mahardhika Surveyor Indonesia. All rights reserved.</p>
          </div>
        </div>
      `;

      MailApp.sendEmail({
        to: recipientEmail,
        subject: emailSubject,
        htmlBody: emailBodyHtml,
        attachments: [pdfFile.getAs(MimeType.PDF)]
      });
    }

    return {
      success: true,
      url: pdfFile.getUrl(),
      recipientEmail: recipientEmail
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}