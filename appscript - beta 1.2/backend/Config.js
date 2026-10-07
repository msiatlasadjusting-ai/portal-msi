// =========================================================
// GLOBAL CONFIGURATION & CONSTANTS
// Konfigurasi ID Spreadsheet, Folder Google Drive, Template Dokumen, dan Enum Sistem
// =========================================================

var SPREADSHEET_ID = "1WA-HYCbMWQBd3amBY_h1slClJ9PTBlavV69FLTbsLPI";
var SHEET_NAME = "Mastercase";
var DOC_TEMPLATE_ID = "1-o_szYE0whI3voOAf7jjefPBjbqx7HgBzNk6Owoyx2I";
var PDF_FOLDER_ID = "1AxuGBjLZGeu5RrCJP0BqoOKUPmDQuXOs";
var DRIVE_PARENT_FOLDER_ID = "1aXiYVRM9jtxRjlSLu_THtriGiBExdzaF";
var FIREBASE_PROJECT_ID = "msi-atlas-adjusting";

var ENABLE_EMAIL_NOTIFICATIONS = true;
var DEFAULT_MANAGER_FALLBACK_EMAIL = "msi.atlasadjusting@gmail.com";

var SLA_STAGE_LIMITS = {
  STAGE_1_REGISTER: 0, 
  STAGE_2_SURVEY: 0,   
  STAGE_3_REPORT: 5,   
  STAGE_4_LOP: 3,      
  STAGE_5_INVOICE: 7 
};

var STAGE_DEFAULT_STATUS = {
  1: 'On Process',
  2: 'In Site Survey',
  3: 'Drafting Report',
  4: 'Awaiting Invoice',
  5: 'Invoice Sent'
};
