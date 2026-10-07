// =========================================================
// USER PROFILE & PIC MANAGEMENT SERVICE
// Pengelolaan penyimpanan foto profil di Google Drive dan sinkronisasi database PIC Spreadsheet
// =========================================================

// =========================================================
// UPLOAD PROFILE PHOTO TO DRIVE
// Mendekode foto profil Base64, menghapus foto versi sebelumnya, dan menyimpan file baru ke Drive
// =========================================================
function uploadProfilePhotoToDrive(base64Data, fileName) {
  try {
    if (!base64Data) {
      throw new Error("Data foto Base64 tidak boleh kosong.");
    }

    const folderName = "MSI_Profile_Photos";
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const parts = fileName.split("_");
    if (parts.length >= 2) {
      const targetUid = parts[1]; 
      const searchPrefix = "PROFILE_" + targetUid + "_";

      try {
        const existingFiles = folder.getFiles();
        while (existingFiles.hasNext()) {
          const file = existingFiles.next();
          const curName = file.getName();

          if (curName.indexOf(searchPrefix) === 0 || curName.includes("PROFILE_" + targetUid)) {
            try {
              file.setTrashed(true);
            } catch (trashErr) {
              Logger.log("Gagal memindahkan file lama ke trash: " + trashErr.toString());
            }
          }
        }
      } catch (searchErr) {
        Logger.log("Gagal mencari file lama: " + searchErr.toString());
      }
    }

    const splitData = base64Data.split(",");
    const contentType = (splitData[0].match(/:(.*?);/) || [])[1] || "image/jpeg";
    const base64String = splitData.length > 1 ? splitData[1] : splitData[0];
    const bytes = Utilities.base64Decode(base64String);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const newFile = folder.createFile(blob);

    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("Domain membatasi setSharing otomatis: " + shareErr.toString());
    }

    const fileId = newFile.getId();
    const directUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    return {
      success: true,
      url: directUrl,
      fileId: fileId
    };

  } catch (error) {
    Logger.log("Error Upload Drive: " + error.toString());
    return {
      success: false,
      error: error.message || String(error)
    };
  }
}

// =========================================================
// SYNC USER PIC TO SPREADSHEET
// Sinkronisasi data identitas PIC (Nama, Kode PIC, Email) dari Firestore ke Sheet PIC
// =========================================================
function syncUserPicToSpreadsheet(payload) {
  try {
    if (!SPREADSHEET_ID) {
      throw new Error("SPREADSHEET_ID belum dikonfigurasi.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let picSheet = ss.getSheetByName("PIC");

    if (!picSheet) {
      picSheet = ss.insertSheet("PIC");
      picSheet.appendRow(["Nama", "Kode", "Email"]);
    }

    const name = String(payload.name || "").trim();
    const code = String(payload.picCode || "").trim().toUpperCase();
    const email = String(payload.email || "").trim().toLowerCase();

    if (!code) {
      return { success: false, error: "Kode PIC tidak boleh kosong." };
    }

    const lastRow = picSheet.getLastRow();
    let foundRow = -1;

    if (lastRow > 1) {
      const picValues = picSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < picValues.length; i++) {
        const rowNama = String(picValues[i][0] || "").trim(); 
        const rowKode = String(picValues[i][1] || "").trim().toUpperCase(); 

        if (rowKode === code || (name && rowNama.toUpperCase() === name.toUpperCase())) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > 0) {
      picSheet.getRange(foundRow, 1).setValue(name);
      picSheet.getRange(foundRow, 2).setValue(code);
      if (email) picSheet.getRange(foundRow, 3).setValue(email);
    } else {
      picSheet.appendRow([name, code, email]);
    }

    return {
      success: true,
      message: "Data Kode PIC & Email berhasil diperbarui di Spreadsheet!"
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =========================================================
// FORCE DRIVE AUTHORIZATION
// Trigger akses DriveApp untuk memicu otorisasi OAuth Google Drive
// =========================================================
function forceDriveAuthorization() {
  const folders = DriveApp.getFoldersByName("MSI_Profile_Photos");
  if (folders.hasNext()) {
    const folder = folders.next();
    const files = folder.getFiles();
    if (files.hasNext()) {
      const file = files.next();
      Logger.log("File ditemukan: " + file.getName());
    }
  } else {
    Logger.log("Folder tidak ditemukan.");
  }
}
