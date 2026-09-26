// ===============================================================
// Code.gs - updated to sync edited sales prices back to Products
// ===============================================================

const aSheet = SpreadsheetApp.getActiveSpreadsheet();
const reportsSheet = aSheet.getSheetByName("Reports");
const salesSheet = aSheet.getSheetByName("sales");
const salesOfCompetitorSheet = aSheet.getSheetByName("salesOfCompetitor");
const expensesSheet = aSheet.getSheetByName("expenses");
const promotersSheet = aSheet.getSheetByName("promoters");
// V42: صفحة Users أُلغيت — حسابات الموظفين (id, username, password, role)
// أصبحت ضمن صفحة "Employees" بالأعمدة: A=id، B=name، C=username، D=password،
// E=role (admin/manager/user)، F=jobPosition (المنصب الوظيفي)، G=mgr (اسم المدير المباشر).
const ATTENDANCE_SHEET_NAME = "attendance";
const ATTENDANCE_HEADERS = ["timestamp","username","status","attendanceStatement","notes","medicalReportUrl","approvalStatus","approvedBy","approvedAt","rejectionReason","rejectionBy"];
const STATUS_WT_SHEET_NAME = "statusWT";
const MEDICAL_REPORT_FOLDER_NAME = "التقارير الطبية - الدوام";
const SICK_LEAVE_STATUS_LABEL = "اجازة مرضية";
const cache = CacheService.getScriptCache();
const CACHE_EXPIRATION_SECONDS = 3600;
const INITIAL_DATA_CACHE_VERSION = 'v26-employees-merged';
const PRODUCT_LOOKUP_CACHE_VERSION = 'v25-product-lookup';

// ===============================================================
// V67: نظام الجلسات (التوثيق). سابقاً كان العميل يحدد هويته ودوره بنفسه في كل
// طلب، ما كان يسمح لأي شخص يعرف رابط التطبيق بقراءة كل البيانات أو التلاعب بها.
// الآن يُصدر doLogin رمز token، وتتحقق كل دوال القراءة/الكتابة من الرمز ثم
// تستخرج الهوية والدور من ملف Employees (المصدر الموثوق) — أي قيمة يرسلها
// العميل لهوية/دور تُتجاهل وتُستبدل بقيمة الخادم.
// مدة الجلسة: 6 ساعات (الحد الأقصى المسموح لـ ScriptCache TTL).
// ===============================================================
const SESSION_CACHE_PREFIX_ = 'sess_';
const SESSION_DURATION_SECONDS_ = 6 * 60 * 60;
const SESSION_STORE_ = CacheService.getScriptCache();

// خطأ توثيق يحمل code='unauthorized' حتى يميزه العميل عن بقية الأخطاء.
function unauthorizedError_(message) {
  const err = new Error(message || 'Unauthorized');
  err.code = 'unauthorized';
  return err;
}

// قراءة صف الموظف من ملف Employees — المصدر الموثوق للهوية والدور والمنصب.
function getEmployeeByIdOrUsername_(value) {
  const sheet = aSheet.getSheetByName("Employees");
  if (!sheet || sheet.getLastRow() < 2) return null;
  const v = String(value || '').trim();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowId = String(row[0] || '').trim();
    const rowUsername = String(row[2] || '').trim();
    if (rowId === v || rowUsername.toLowerCase() === v.toLowerCase()) {
      return {
        id: rowId || rowUsername.toLowerCase(),
        username: rowUsername,
        systemRole: String(row[4] || '').trim().toLowerCase(),
        jobPosition: String(row[5] || '').trim(),
        mgr: String(row[6] || '').trim(),
        name: String(row[1] || '').trim() || rowUsername
      };
    }
  }
  return null;
}

// إنشاء جلسة وإرجاع الـ token الجديد.
function createSession_(employee) {
  const token = Utilities.getUuid().replace(/-/g, '');
  const expiresAt = Date.now() + SESSION_DURATION_SECONDS_ * 1000;
  try { SESSION_STORE_.put(SESSION_CACHE_PREFIX_ + token, employee.id + '|' + expiresAt, SESSION_DURATION_SECONDS_); } catch (e) {}
  return { token, expiresAt };
}

// إنهاء الجلسة فوراً (زر خروج) — لا يصلح الـ token بعدها.
function revokeSession_(token) {
  const t = String(token || '').trim();
  if (!t) return;
  try { SESSION_STORE_.remove(SESSION_CACHE_PREFIX_ + t); } catch (e) {}
}

// تحويل token صالح إلى كائن الموظف الموثّق — أي طلب بلا token أو بجلسة منتهية يُرفض.
function authenticateRequest_(token) {
  const t = String(token || '').trim();
  if (!t) throw unauthorizedError_('لا توجد جلسة دخول — سجّل الدخول أولاً');
  let entry = null;
  try { entry = SESSION_STORE_.get(SESSION_CACHE_PREFIX_ + t); } catch (e) {}
  if (!entry) throw unauthorizedError_('انتهت الجلسة أو غير صالحة — سجّل الدخول مجدداً');
  const parts = String(entry).split('|');
  if (Number(parts[1]) < Date.now()) {
    try { SESSION_STORE_.remove(SESSION_CACHE_PREFIX_ + t); } catch (e) {}
    throw unauthorizedError_('انتهت مدة الجلسة — سجّل الدخول مجدداً');
  }
  const emp = getEmployeeByIdOrUsername_(parts[0]);
  if (!emp) throw unauthorizedError_('الحساب غير موجود');
  return emp;
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    // V67: كل أفعال GET تتطلب جلسة موثّقة؛ الهوية والدور من ملف Employees على الخادم.
    const emp = authenticateRequest_(e.parameter.token);
    let response;
    switch (action) {
      case "getInitialData": response = getInitialData(e.parameter.forceRefresh === "1"); break;
      case "getReports": response = getReports(emp.id, emp.systemRole, emp.name, e.parameter.targetUserId); break;
      case "getReportById": response = getReportById(e.parameter.id, emp); break;
      case "findProductByBarcode": response = findProductByBarcode(e.parameter.barcode, e.parameter.campaign); break;
      case "getCompetitorProducts": response = getCompetitorProductsFromSheet(); break;
      case "getUserFestivalMovements": response = getUserFestivalMovements(emp.id, emp.systemRole, e.parameter.targetUserId); break;
      case "getTeamOptions": response = getTeamOptions(emp.id, emp.systemRole); break;
      case "getAttendance": response = getAttendance(emp.id, emp.systemRole, emp.name, e.parameter.targetUserId); break;
      case "getStatusOptions": response = getStatusOptions(); break;
      case "getDashboardData": response = getDashboardData(emp.id, emp.systemRole); break;
      case "getPromoterGoals": response = getPromoterGoals(e.parameter.month, emp.id, emp.systemRole); break;
      case "getUsers": response = getUsersAdmin(emp); break;
      default: response = {status:"error", message:"Invalid GET Action"};
    }
    return createJsonResponse(response);
  } catch (err) {
    return createJsonResponse({status:"error", code:err.code||'', message:err.message});
  }
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const action = request.action;
    let response;

    if (action === "doLogin") {
      response = doLogin(request.payload);
    } else if (action === "logout") {
      // V67: إلغاء الجلسة فوراً — يُرسل العميل الـ token الحالي ليُبطَل.
      revokeSession_((request.payload || {}).token);
      response = {status:"success"};
    } else {
      const raw = request.payload || {};
      // V67: هوية ودور موثّقان من ملف Employees — نستبدل أي قيم يرسلها العميل.
      const emp = authenticateRequest_(raw.token);
      const payload = Object.assign({}, raw);
      delete payload.token;
      payload.userId = emp.id;
      payload.userName = emp.name;
      payload.role = emp.systemRole;
      payload.mgr = emp.mgr;
      payload.jobPosition = emp.jobPosition;
      // ملاحظة: payload.username لا يُكتب هنا — فهو يحمل معنيين مختلفين:
      // هوية مُسجِّل الدوام (submitAttendance) أو معرّف سجل الدوام المستهدف (approveAttendance).
      // يُضبط لمجال الدوام فقط عند الحاجة أدناه.
      // أسماء/معرّفات المُنشئ والمُعدّل تُفرض من الخادم (لا يمكن تزويرها).
      payload.createdById = emp.id;
      payload.createdByName = emp.name;
      payload.editedBy = emp.name;

      switch (action) {
        case "submitReport":
          response = handleReportSubmission(payload);
          break;

        case "addLocation":
          response = addLocation(payload);
          break;

        case "updateProductPrice":
          response = updateProductPrice(payload);
          break;

        case "refreshCache":
          response = refreshCache();
          break;

        case "addFestivalMovement":
          response = addFestivalMovement(payload);
          break;

        case "closeOutUserTallyToSales":
          response = closeOutUserTallyToSales(payload);
          break;

        case "submitAttendance":
          // V67: هوية مُسجِّل الدوام من الحساب الموثّق حصراً — لا يمكن التبليغ عن الدوام بصيغة مغايرة.
          payload.username = emp.username;
          response = submitAttendance(payload);
          break;

        case "archiveOldReports":
          response = archiveOldReports(payload);
          break;

        // V50: اعتماد التقارير — دور «مدقق» (auditor) أو admin فقط.
        case "approveReport":
          response = approveReport(payload);
          break;

        case "approveAllReports":
          response = approveAllReports(payload);
          break;

        // V61: حذف ناعم لتقرير — admin/manager/auditor.
        case "deleteReport":
          response = deleteReport(payload);
          break;

        // V66: اعتماد/رفض سجل دوام.
        case "approveAttendance":
          response = approveAttendance(payload);
          break;

        case "approveAllAttendance":
          response = approveAllAttendance(payload);
          break;

        // V66: اعتماد/رفض حركة سحب/مرتجع.
        case "approveMovement":
          response = approveMovement(payload);
          break;

        case "approveAllMovements":
          response = approveAllMovements(payload);
          break;

        // V69: تعديل حركة مادة مرفوضة/قيد المراجعة من صاحبها فقط — المعتمدة لا تُعدَّل.
        case "editFestivalMovement":
          response = editFestivalMovement(payload, emp);
          break;

        // V69: تعديل سجل دوام مرفوض/قيد المراجعة من صاحبه فقط — المعتمد لا يُعدَّل.
        case "editAttendance":
          response = editAttendance(payload, emp);
          break;

        // V63: أهداف المروجين الشهرية.
        case "savePromoterGoals":
          response = savePromoterGoals(payload);
          break;

        // V63: النسخ الاحتياطي الأسبوعي.
        case "setupWeeklyBackup":
          response = setupWeeklyBackup(payload);
          break;
        case "backupNow":
          response = backupNow(payload);
          break;

        // V68: إدارة المستخدمين والصلاحيات — للإداري (admin) فقط.
        case "saveUser":
          // «المنصب الوظيفي» (jobPosition) و«المدير المباشر» (manager) حقلان من نموذج
          // إدارة المستخدمين — نستعيدهما من الطلب الخام لأن doPost يفرض قيمهما
          // تلقائياً لصاحب الطلب نفسه.
          payload.jobPosition = String(raw.jobPosition || '').trim();
          payload.actorUsername = emp.username;
          response = saveUser(payload);
          break;
        case "deleteUser":
          payload.actorUsername = emp.username;
          response = deleteUser(payload);
          break;

        default:
          response = {
            status: "error",
            message: "Invalid POST Action: " + action
          };
      }
    }

    return createJsonResponse(response);

  } catch (err) {
    return createJsonResponse({
      status: "error",
      code: err.code || '',
      message: err.message
    });
  }
}

    function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getAttendanceSheet_() {
  let sheet = aSheet.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sheet) sheet = aSheet.insertSheet(ATTENDANCE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,ATTENDANCE_HEADERS.length).setValues([ATTENDANCE_HEADERS]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < ATTENDANCE_HEADERS.length) {
    // Upgrade older sheets (created before the medicalReportUrl column existed)
    // by filling in any missing header cells without touching existing data.
    const existingHeaders = sheet.getRange(1,1,1,ATTENDANCE_HEADERS.length).getValues()[0];
    ATTENDANCE_HEADERS.forEach((h,i) => { if (!existingHeaders[i]) sheet.getRange(1,i+1).setValue(h); });
  }
  return sheet;
}

// Fallback list used only if the "statusWT" sheet cannot be created/read for
// some reason (e.g. permissions). Also used as the seed data the very first
// time the "statusWT" sheet is created.
const ATTENDANCE_ALLOWED_STATUSES = ["بداية دوام","نهاية دوام","عطلة أسبوعية","عطلة رسمية","اجازة إدارية","اجازة مرضية","حضور إضافي"];

// Returns the "statusWT" sheet, creating it (with a header row + the default
// status list already filled in) the very first time the app runs.
function getStatusWTSheet_() {
  let sheet = aSheet.getSheetByName(STATUS_WT_SHEET_NAME);
  if (!sheet) {
    sheet = aSheet.insertSheet(STATUS_WT_SHEET_NAME);
    sheet.getRange(1,1,1,1).setValue("حالة الدوام");
    sheet.setFrozenRows(1);
    sheet.getRange(2,1,ATTENDANCE_ALLOWED_STATUSES.length,1).setValues(ATTENDANCE_ALLOWED_STATUSES.map(s => [s]));
  }
  return sheet;
}

// Reads the list of attendance-status options from the "statusWT" sheet
// (column A, skipping the header row). Creates the sheet with the default
// options on first run so the dropdowns always have data to show.
function getStatusWTOptions_() {
  const sheet = getStatusWTSheet_();
  if (sheet.getLastRow() < 2) return ATTENDANCE_ALLOWED_STATUSES.slice();
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  const options = values.map(row => String(row[0] || '').trim()).filter(v => v);
  return options.length ? options : ATTENDANCE_ALLOWED_STATUSES.slice();
}

// GET action: returns the current dropdown options so the front-end (attendance
// form select + history filter) always mirrors the "statusWT" sheet.
function getStatusOptions() {
  return {status:"success", options: getStatusWTOptions_()};
}

// Returns the Drive folder used for medical reports, creating it once and
// remembering its ID in Script Properties. This intentionally avoids
// DriveApp.getFoldersByName(), which needs the broad "See, edit, create and
// delete ALL your Drive files" permission (drive / drive.readonly scope).
// Only ever creating/opening a folder the script itself made keeps Apps
// Script's auto-detected scope down to the narrower drive.file permission.
const MEDICAL_REPORT_FOLDER_ID_PROP = "MEDICAL_REPORT_FOLDER_ID";
function getMedicalReportFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(MEDICAL_REPORT_FOLDER_ID_PROP);
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) { /* folder was moved/deleted - recreate below */ }
  }
  const folder = DriveApp.createFolder(MEDICAL_REPORT_FOLDER_NAME);
  props.setProperty(MEDICAL_REPORT_FOLDER_ID_PROP, folder.getId());
  return folder;
}

// Saves an uploaded medical report (PDF or image, sent as base64) to Drive and
// returns a shareable link to store alongside the attendance row.
function saveMedicalReport_(base64Data, fileName, mimeType) {
  let folder;
  try {
    folder = getMedicalReportFolder_();
  } catch (e) {
    throw new Error(driveAuthErrorMessage_(e));
  }
  const safeMime = String(mimeType || '').trim() || 'application/octet-stream';
  const safeName = String(fileName || 'medical-report').trim();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), safeMime, safeName);
  let file;
  try {
    file = folder.createFile(blob);
  } catch (e) {
    throw new Error(driveAuthErrorMessage_(e));
  }
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* sharing may be restricted by domain policy */ }
  return file.getUrl();
}

// Turns Google's raw "you do not have permission to call DriveApp..." error
// into a clear, actionable Arabic message pointing at the one-time fix
// (see setupDriveAccess() below and the README).
function driveAuthErrorMessage_(e) {
  const raw = String((e && e.message) || e || '');
  if (/DriveApp|drive\.readonly|permission/i.test(raw)) {
    return "لم يتم منح صلاحية Google Drive لهذا المشروع بعد. افتح محرر Apps Script، شغّل الدالة setupDriveAccess مرة واحدة ووافق على الصلاحيات المطلوبة، ثم انشر نسخة جديدة من الويب أب (Deploy > Manage deployments > New version). راجع ملف الـ README المرفق لخطوات مفصّلة.";
  }
  return raw || "حدث خطأ غير متوقع أثناء حفظ التقرير الطبي على Drive";
}

/**
 * ⚠️ شغّل هاي الدالة يدوياً من محرر Apps Script (زر ▶ Run) مرة واحدة فقط بعد
 * نسخ هذا الكود لأول مرة، أو أي مرة يُطلب منك فيها صلاحية Drive من جديد.
 * هدفها الوحيد هو إظهار نافذة "Authorization required" ليوافق مالك المشروع
 * عليها؛ بدون هذه الموافقة اليدوية (لمرة واحدة) ترفض Google أي استدعاء
 * لـ DriveApp من الويب أب برسالة "You do not have permission...".
 * بعد الموافقة، لازم تنشر نسخة جديدة من الويب أب حتى يعمل رفع التقارير
 * الطبية: Deploy > Manage deployments > قلم التعديل > Version: New version > Deploy.
 */
function setupDriveAccess() {
  const folder = getMedicalReportFolder_();
  Logger.log('تم منح صلاحية Drive بنجاح. رابط مجلد التقارير الطبية: ' + folder.getUrl());
}

function submitAttendance(payload) {
  const sheet = getAttendanceSheet_();
  const username = String(payload && (payload.username || payload.userName) || '').trim();
  const status = String(payload && payload.status || '').trim();
  const statement = String(payload && (payload.attendanceStatement || payload.statement) || '').trim();
  const notes = String(payload && payload.notes || '').trim();
  if (!username) throw new Error("اسم المستخدم مطلوب");
  const allowedStatuses = getStatusWTOptions_();
  if (allowedStatuses.indexOf(status) === -1) throw new Error("حالة الدوام غير صحيحة");
  if (!statement) throw new Error("بيان الدوام مطلوب");

  let medicalReportUrl = '';
  const fileBase64 = payload && payload.medicalReportBase64;
  if (status === SICK_LEAVE_STATUS_LABEL) {
    if (!fileBase64) throw new Error("يجب إرفاق التقرير الطبي (PDF أو صورة) عند اختيار حالة اجازة مرضية");
    medicalReportUrl = saveMedicalReport_(fileBase64, payload.medicalReportFileName, payload.medicalReportMimeType);
  }

  sheet.appendRow([new Date(),username,status,statement,notes,medicalReportUrl]);
  sheet.getRange(sheet.getLastRow(),1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
  return {status:"success",message:"تم تسجيل الدوام بنجاح"};
}

function getAttendance(userId, userRole, userName, targetUserId) {
  const sheet = getAttendanceSheet_();
  if (sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), ATTENDANCE_HEADERS.length);
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  const headersAll = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());
  const colOf = (name) => { const i = headersAll.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(name)); return i >= 0 ? i : -1; };
  const iApproval = colOf('approvalStatus'), iApprovedBy = colOf('approvedBy'), iApprovedAt = colOf('approvedAt'), iRejReason = colOf('rejectionReason'), iRejBy = colOf('rejectionBy');
  const role = String(userRole || '').trim().toLowerCase();
  const requesterId = String(userId || '').trim();
  const selectedId = String(targetUserId || '').trim();
  const allowedIds = getTeamMemberIds(requesterId,role);
  let allowedNames = null;
  const users = getUsersLite();
  if (allowedIds !== null) {
    const allowed = users.filter(u => allowedIds.indexOf(u.id) !== -1);
    allowedNames = allowed.map(u => String(u.username || u.name || ''));
    if (selectedId) {
      const selected = allowed.find(u => u.id === selectedId);
      allowedNames = selected ? [String(selected.username || selected.name || '')] : [];
    }
  } else if (selectedId) {
    const selected = users.find(u => u.id === selectedId);
    allowedNames = selected ? [String(selected.username || selected.name || '')] : [];
  }
  return values.filter(row => allowedNames === null || allowedNames.indexOf(String(row[1] || '')) !== -1)
    .reverse().map(row => ({
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0],Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm:ss") : String(row[0] || ''),
      username:String(row[1] || ''), status:String(row[2] || ''),
      attendanceStatement:String(row[3] || ''), notes:String(row[4] || ''),
      medicalReportUrl:String(row[5] || ''),
      approvalStatus: iApproval >= 0 ? String(row[iApproval] || '') : '',
      approvedBy: iApprovedBy >= 0 ? String(row[iApprovedBy] || '') : '',
      approvedAt: iApprovedAt >= 0 ? String(row[iApprovedAt] || '') : '',
      rejectionReason: iRejReason >= 0 ? String(row[iRejReason] || '') : '',
      rejectionBy: iRejBy >= 0 ? String(row[iRejBy] || '') : ''
    }));
}

// V66: اعتماد/رفض سجل دوام — admin/manager/auditor. يحدد الصف عبر timestamp+username.
function approveAttendance(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager' && role !== 'auditor') {
    throw new Error('صلاحية الاعتماد متاحة للمدقق (auditor) والإداري (admin) والمدير (manager) فقط');
  }
  const timestamp = String(payload.timestamp || '').trim();
  const username = String(payload.username || '').trim();
  if (!timestamp || !username) throw new Error('معرّف سجل الدوام (الوقت والمستخدم) مطلوب');
  const status = String(payload.status || 'approved').trim().toLowerCase();
  if (status !== 'approved' && status !== 'rejected') throw new Error('حالة غير صالحة (approved أو rejected)');
  const sheet = getAttendanceSheet_();
  const lastCol = Math.max(sheet.getLastColumn(), ATTENDANCE_HEADERS.length);
  const headersAll = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h || '').trim());
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  const tz = Session.getScriptTimeZone();
  let targetRow = -1;
  for (let i = 0; i < data.length; i++) {
    const cellTime = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], tz, 'yyyy-MM-dd HH:mm:ss') : String(data[i][0] || '');
    if (cellTime === timestamp && String(data[i][1] || '').trim() === username) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) throw new Error('سجل الدوام غير موجود');
  const col = (name, fallback) => {
    const idx = headersAll.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(name));
    return idx >= 0 ? idx + 1 : fallback;
  };
  const by = String(payload.approvedBy || payload.userName || '');
  const now = String(payload.approvedAt || new Date().toISOString());
  sheet.getRange(targetRow, col('approvalStatus', 7)).setValue(status);
  sheet.getRange(targetRow, col('approvedBy', 8)).setValue(by);
  sheet.getRange(targetRow, col('approvedAt', 9)).setValue(now);
  sheet.getRange(targetRow, col('rejectionReason', 10)).setValue(status === 'rejected' ? String(payload.reason || '') : '');
  sheet.getRange(targetRow, col('rejectionBy', 11)).setValue(status === 'rejected' ? by : '');
  return { status: 'success', message: status === 'approved' ? 'تم اعتماد سجل الدوام.' : 'تم رفض سجل الدوام.' };
}

// V66: اعتماد/رفض حركة سحب/مرتجع/صرف/مبيعات — admin/manager/auditor. يحدد الصف عبر المعرّف id.
function approveMovement(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager' && role !== 'auditor') {
    throw new Error('صلاحية الاعتماد متاحة للمدقق (auditor) والإداري (admin) والمدير (manager) فقط');
  }
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('معرّف الحركة مطلوب');
  const status = String(payload.status || 'approved').trim().toLowerCase();
  if (status !== 'approved' && status !== 'rejected') throw new Error('حالة غير صالحة (approved أو rejected)');
  const sheet = ensureFestivalMovementSheet();
  const lastCol = Math.max(sheet.getLastColumn(), FESTIVAL_MOVEMENT_HEADERS.length);
  const headersAll = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h || '').trim());
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  let targetRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][8] || '').trim() === id) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) throw new Error('الحركة غير موجودة');
  const col = (name) => {
    const idx = headersAll.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(name));
    return idx >= 0 ? idx + 1 : -1;
  };
  const by = String(payload.approvedBy || payload.userName || '');
  const now = String(payload.approvedAt || new Date().toISOString());
  const sets = { approvalStatus: status, approvedBy: by, approvedAt: now, rejectionReason: status === 'rejected' ? String(payload.reason || '') : '', rejectionBy: status === 'rejected' ? by : '' };
  Object.keys(sets).forEach(name => {
    const c = col(name);
    if (c >= 0) sheet.getRange(targetRow, c).setValue(sets[name]);
  });
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
  return { status: 'success', message: status === 'approved' ? 'تم اعتماد الحركة.' : 'تم رفض الحركة.' };
}

// V69: تعديل حركة سحب/مرتجع غير معتمدة — صاحبها (أو admin/manager/auditor).
// المعتمدة لا تُعدَّل، وبعد التعديل تعود الحركة «قيد المراجعة» ويُمحى أثر الاعتماد/الرفض.
function editFestivalMovement(payload, emp) {
  payload = payload || {};
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('معرّف الحركة مطلوب');

  const item = String(payload.item || '').trim();
  const quantity = Number(payload.quantity);
  const operation = String(payload.operation || '').trim();
  const invoiceNumber = String(payload.invoiceNumber || '').trim();
  if (!item) throw new Error('اسم المادة مطلوب');
  if (!Number.isFinite(quantity) || quantity <= 0) throw new Error('الكمية يجب أن تكون رقماً أكبر من 0');
  if (operation !== 'سحب' && operation !== 'مرتجع') throw new Error('نوع العملية يجب أن يكون "سحب" أو "مرتجع"');

  const requesterId = String((emp && emp.id) || payload.createdById || '').trim();
  const requesterRole = String((emp && emp.systemRole) || '').trim().toLowerCase();

  const sheet = ensureFestivalMovementSheet();
  const lastCol = Math.max(sheet.getLastColumn(), FESTIVAL_MOVEMENT_HEADERS.length);
  const headersAll = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h || '').trim());
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  let targetRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][8] || '').trim() === id) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) throw new Error('الحركة غير موجودة');

  const ownerId = String(data[targetRow - 2][5] || '').trim();
  const canEdit = requesterRole === 'admin' || requesterRole === 'manager' || requesterRole === 'auditor' || (ownerId !== '' && ownerId === requesterId);
  if (!canEdit) throw new Error('لا تملك صلاحية تعديل هذه الحركة');

  if (String(data[targetRow - 2][9] || '').trim() !== 'rejected') throw new Error('لا يمكن تعديل الحركة إلا إذا كانت مرفوضة');

  const col = (name) => {
    const idx = headersAll.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(name));
    return idx >= 0 ? idx + 1 : -1;
  };
  const sets = {
    item: item, quantity: quantity, operation: operation, invoiceNumber: invoiceNumber,
    approvalStatus: '', approvedBy: '', approvedAt: '', rejectionReason: '', rejectionBy: ''
  };
  Object.keys(sets).forEach(name => {
    const c = col(name);
    if (c >= 0) sheet.getRange(targetRow, c).setValue(sets[name]);
  });
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
  return { status: 'success', message: 'تم تعديل الحركة وأصبحت قيد المراجعة مجدداً.' };
}

// V69: تعديل سجل دوام غير معتمد — صاحبه (أو admin/manager/auditor).
// المعتمد لا يُعدَّل، وبعد التعديل يعود السجل «قيد المراجعة» ويُمحى أثر الاعتماد/الرفض.
function editAttendance(payload, emp) {
  payload = payload || {};
  const timestamp = String(payload.timestamp || '').trim();
  const username = String(payload.username || '').trim();
  if (!timestamp || !username) throw new Error('معرّف سجل الدوام (الوقت والمستخدم) مطلوب');

  const status = String(payload.status || '').trim();
  const statement = String(payload.attendanceStatement || payload.statement || '').trim();
  const notes = String(payload.notes || '').trim();
  const allowedStatuses = getStatusWTOptions_();
  if (allowedStatuses.indexOf(status) === -1) throw new Error('حالة الدوام غير صالحة');
  if (!statement) throw new Error('بيان الدوام مطلوب');

  const requesterRole = String((emp && emp.systemRole) || '').trim().toLowerCase();
  const requesterUsername = String((emp && emp.username) || payload.editedBy || '').trim();

  const sheet = getAttendanceSheet_();
  const lastCol = Math.max(sheet.getLastColumn(), ATTENDANCE_HEADERS.length);
  const headersAll = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h || '').trim());
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  const tz = Session.getScriptTimeZone();
  let targetRow = -1;
  for (let i = 0; i < data.length; i++) {
    const cellTime = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], tz, 'yyyy-MM-dd HH:mm:ss') : String(data[i][0] || '');
    if (cellTime === timestamp && String(data[i][1] || '').trim() === username) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) throw new Error('سجل الدوام غير موجود');

  const canEdit = requesterRole === 'admin' || requesterRole === 'manager' || requesterRole === 'auditor' || (username !== '' && username === requesterUsername);
  if (!canEdit) throw new Error('لا تملك صلاحية تعديل سجل الدوام هذا');

  if (String(data[targetRow - 2][6] || '').trim() !== 'rejected') throw new Error('لا يمكن تعديل سجل الدوام إلا إذا كان مرفوضاً');

  const col = (name) => {
    const idx = headersAll.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(name));
    return idx >= 0 ? idx + 1 : -1;
  };
  const sets = {
    status: status, attendanceStatement: statement, notes: notes,
    approvalStatus: '', approvedBy: '', approvedAt: '', rejectionReason: '', rejectionBy: ''
  };
  Object.keys(sets).forEach(name => {
    const c = col(name);
    if (c >= 0) sheet.getRange(targetRow, c).setValue(sets[name]);
  });
  return { status: 'success', message: 'تم تعديل سجل الدوام وأصبح قيد المراجعة مجدداً.' };
}

function doLogin(payload) {
  const sheet = aSheet.getSheetByName("Employees");
  if (!sheet || sheet.getLastRow() < 2) throw new Error("لا توجد بيانات موظفين");
  // الأعمدة: A=id(0)، B=name(1)، C=username(2)، D=password(3)، E=role(4) ...
  const usersData = sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
  const inputUsernameLower = String(payload.username || '').toLowerCase();
  const inputPassword = String(payload.password || '');

  // V42: تحقق آمن من كلمة المرور.
  // 1) إذا كانت المخزنة بصيغة hash (تبدأ بـ "pwd$") نقارن بالـ hash.
  // 2) وإلا (نص عادي من أنظمة أقدم) نقارن مباشرة، وعند النجاح نرقّي المخزنة إلى hash تلقائياً
  //    حتى لا تبقى أي كلمة مرور نصية في الـ Sheet.
  for (let i = 0; i < usersData.length; i++) {
    const row = usersData[i];
    if (String(row[2]).toLowerCase() !== inputUsernameLower) continue;
    const stored = String(row[3] || '');
    if (!stored) continue;
    let valid = false;
    if (stored.indexOf('pwd$') === 0) {
      // صيغة مخزنة: pwd$<salt>$<hash> حيث hash = HMAC(salt+password)
      const parts = stored.split('$');
      if (parts.length === 3) {
        const salt = parts[1];
        const expected = parts[2];
        const actual = computePasswordHash_(inputPassword, salt);
        valid = (actual === expected);
      }
    } else {
      // نص عادي قديم: نتأكد من التطابق، ومن ثم نرقّيها لاحقاً.
      valid = (stored === inputPassword);
    }
    if (!valid) continue;
    // ترقية تلقائية لكلمات المرور النصية القديمة إلى hash (مرة واحدة لكل مستخدم).
    if (stored.indexOf('pwd$') !== 0) {
      const { salt, hash } = hashPassword_(inputPassword);
      sheet.getRange(i + 2, 4).setValue('pwd$' + salt + '$' + hash);
    }
    // V42: المعرّف للواجهة = عمود id، وإن كان فارغاً نعتمد على username كبديل حتى
    // تُبنى صلاحيات المدير (الفريق) بشكل صحيح حتى لمن لم يملأ عمود id.
    const rawId = String(row[0] || '').trim();
    const rawName = String(row[1] || '').trim();
    const rawUsername = String(row[2] || '').trim();
    // V67: إنشاء جلسة وإرجاع الـ token — كل الطلبات اللاحقة تتطلب هذا الـ token.
    const session = createSession_({ id: rawId || rawUsername });
    // V67: نُسابق قاعدة البيانات الأساسية ضمن استجابة الدخول — دخول أسرع بطلب واحد.
    let initDb = null;
    try { initDb = getInitialData(false); } catch (e) { initDb = null; }
    return {status:"success", token: session.token, expiresAt: session.expiresAt, db: initDb, user:{
      id: rawId || rawUsername,
      username: rawUsername,
      role: String(row[4] || '').trim().toLowerCase(),
      jobPosition: String(row[5] || '').trim(),
      name: rawName || rawUsername
    }};
  }
  throw new Error("Invalid credentials");
}

// توليد salt عشوائي (طول 16 بايت) وتشفير كلمة المرور به.
// نستخدم Utilities.computeHmacSha256Signature لعدم الحاجة لأي مكتبة خارجية،
// والأهم أنه لا يمكن استرجاع كلمة المرور من الـ hash المخزّن.
function hashPassword_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const hash = computePasswordHash_(password, salt);
  return { salt, hash };
}

function computePasswordHash_(password, salt) {
  const sig = Utilities.computeHmacSha256Signature(String(password), String(salt));
  return sig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// دالة مساعدة لتوليد hash يدوياً لأي مستخدم جديد في الـ Sheet:
// تُستخدم عند إضافة مستخدم جديد لنوع كلمة مرور نصية قديمة؛
// لا تُستدعى تلقائياً في هذا الملف.
function hashPasswordForUser(password) {
  const { salt, hash } = hashPassword_(String(password || ''));
  return 'pwd$' + salt + '$' + hash;
}

// معرّف = لأغراض التوثيق: الصيغة المخزنة الناتجة من hashPasswordForUser
// تكون 'pwd$<salt>$<hash>' مثل: pwd$abc123def4567890$a1b2c3...


// ===============================================================
// نظام الصلاحيات: admin (كل شيء) / manager (فريقه فقط) / user (نفسه فقط)
// ===============================================================

// V42: تطبيع اسم لمقارنة متسامحة (إزالة مسافات مكررة + تجاهل حالة الأحرف) لتجنّب
// فشل التطابق بسبب مسافة إضافية أو فرق في الأحرف الكبيرة/الصغيرة.
function normalizeName_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

// V42: قراءة بيانات حسابات الموظفين من صفحة "Employees" (لم تعد صفحة Users موجودة).
// الأعمدة: A=id، B=name، C=username، D=password، E=role، F=jobPosition، G=mgr.
function getUsersLite() {
  const sheet = aSheet.getSheetByName('Employees');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
  return data
    .map(row => {
      const rawId = String(row[0] ?? '').trim();
      const rawName = String(row[1] ?? '').trim();
      const rawUsername = String(row[2] ?? '').trim();
      return {
        id: rawId || rawUsername, // V42: احتياط — إن كان عمود id فارغاً نعتمد على username
        name: rawName,
        username: rawUsername,
        role: String(row[4] ?? '').trim().toLowerCase(),
        jobPosition: String(row[5] ?? '').trim(),
        mgr: String(row[6] ?? '').trim(),
        nname: normalizeName_(row[1]),
        nusername: normalizeName_(row[2])
      };
    })
    // نحتفظ بالصفوف التي تملك معرّفاً أو اسماً أو اسم مستخدم — لا نحذف من فراغ المعرّف.
    .filter(u => u.id || u.name || u.username);
}

// V42: مصدر الحقيقة لهرمية الفريق هو صفحة "Employees" وعمود (G / mgr) الذي يحوي
// اسم المدير المباشر لكل موظف. نعيد خريطة: الاسم المطبّع -> اسم مديره المطبّع.
function getEmployeeMgrMap() {
  const sheet = aSheet.getSheetByName('Employees');
  if (!sheet || sheet.getLastRow() < 2) return {};
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1] ?? '').trim();   // B = name
    const mgr = String(rows[i][6] ?? '').trim();     // G = mgr
    if (name) map[normalizeName_(name)] = normalizeName_(mgr);
  }
  return map;
}

// يجد اسم الموظف من صفحة Employees حسب المعرّف، مع احتياط لمطابقة اسم المستخدم.
function getUserNameById(userId) {
  const id = String(userId || '').trim();
  if (!id) return '';
  const users = getUsersLite();
  let u = users.find(x => x.id === id);
  if (!u) u = users.find(x => x.username !== '' && normalizeName_(x.username) === normalizeName_(id));
  return u ? (u.name || u.username) : '';
}

// V42: يجمع أسماء كل من يتبع (مباشرة أو عبر التسلسل الهرمي) اسماً معيّناً،
// اعتماداً على صفحة Employees (عمود G / mgr). الـ BFS يغطي عدة مستويات.
function collectSubordinateNames(mgrName, empMgrMap) {
  const result = new Set();
  const start = normalizeName_(mgrName);
  if (!start) return result;
  const queue = [start];
  let guard = 0;
  while (queue.length && guard++ < 2000) {
    const current = queue.shift();
    Object.keys(empMgrMap).forEach(name => {
      if (empMgrMap[name] === current && !result.has(name)) {
        result.add(name);
        queue.push(name);
      }
    });
  }
  return result;
}

// يجمع معرّفات أعضاء الفريق الكامل (التسلسل الهرمي متعدد المستويات من صفحة Employees)
// لمدير معيّن: نفسه + كل من يتبعه مباشرة أو عبر مستوى أدنى.
function getManagerTeamIds(requesterId) {
  const empMgrMap = getEmployeeMgrMap();
  const requesterName = getUserNameById(requesterId);
  const users = getUsersLite();
  // نضمن إدراج المدير نفسه (بالمعرّف أو باسم المستخدم إن كان المعرّف فراغاً).
  const idSet = new Set([String(requesterId).trim()]);
  const requesterNname = normalizeName_(requesterName);
  users.forEach(u => {
    if (u.id === String(requesterId).trim()) idSet.add(u.id);
    if (requesterNname && (u.nname === requesterNname || u.nusername === requesterNname)) idSet.add(u.id);
  });

  if (requesterName && Object.keys(empMgrMap).length) {
    const subNames = collectSubordinateNames(requesterName, empMgrMap);
    users.forEach(u => {
      if (u.nname && subNames.has(u.nname)) idSet.add(u.id);
    });
  }

  return Array.from(idSet);
}

// يحدد معرّفات المستخدمين التي يحق لصاحب الطلب رؤية بياناتهم.
// يعيد null إذا كانت الصلاحية شاملة (admin)، أو مصفوفة معرّفات محصورة (manager/user).
function getTeamMemberIds(requesterId, requesterRole) {
  const role = String(requesterRole || '').trim().toLowerCase();
  const id = String(requesterId || '').trim();
  if (role === 'admin') return null;
  if (role === 'manager') return getManagerTeamIds(id);
  return [id];
}

// قائمة أسماء الموظفين لتعبئة القائمة المنسدلة: admin يرى الجميع، manager يرى فريقه (كل المستويات).
function getTeamOptions(requesterId, requesterRole) {
  const role = String(requesterRole || '').trim().toLowerCase();
  const id = String(requesterId || '').trim();
  if (role !== 'admin' && role !== 'manager') return { status:'success', options: [] };

  const users = getUsersLite();
  let list;
  if (role === 'admin') {
    list = users.filter(u => u.id !== id);
  } else {
    const allowedIds = new Set(getManagerTeamIds(id));
    list = users.filter(u => u.id !== id && allowedIds.has(u.id));
  }

  list.sort((a,b) => String(a.name).localeCompare(String(b.name), 'ar'));
  return { status:'success', options: list.map(u => ({ id:u.id, name:u.name || u.username })) };
}

function normalizeHeaderName(value) {
  return String(value ?? '').trim().replace(/\s+/g,' ').toLowerCase();
}

function getColumnIndex(headers, names, fallback) {
  const normalized = headers.map(normalizeHeaderName);
  for (const name of names) {
    const idx = normalized.indexOf(normalizeHeaderName(name));
    if (idx !== -1) return idx;
  }
  return fallback;
}

function isCancelledValue(value) {
  if (value === true) return true;
  const text = String(value ?? '').trim().toLowerCase();
  return ['true','1','yes','نعم'].includes(text);
}


function getCompetitorProductsFromSheet() {
  const sheet = aSheet.getSheetByName("ProductsOfCompetitor");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];

  const headers = values[0].map(h => String(h ?? '').trim());
  const productCol = getColumnIndex(headers, ['product','المادة','اسم المادة'], 0);
  const barcodeCol = getColumnIndex(headers, ['barcode','باركود','الباركود'], 1);
  const cancelledCol = getColumnIndex(headers, ['إلغاء المادة','cancelled','canceled','الغاء المادة'], 2);

  return values.slice(1).map(row => {
    const name = String(row[productCol] ?? '').trim();
    const barcode = String(row[barcodeCol] ?? '').trim();
    const cancelled = isCancelledValue(row[cancelledCol]);
    return { name, barcode, cancelled };
  }).filter(p => p.name && !p.cancelled);
}

function getInitialData(forceRefresh) {
  const cacheKey = "initialData_" + INITIAL_DATA_CACHE_VERSION;
  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
  } else {
    cache.remove(cacheKey);
  }

  // Read only the sheets required by the application instead of scanning every
  // spreadsheet tab on each cache rebuild. This is noticeably faster on large files.
  const requiredSheetNames = ["Locations","Products","ProductsOfCompetitor","Employees","ExpenseItems"];
  const sheetData = {};
  requiredSheetNames.forEach(sheetName => {
    const sheet = aSheet.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 0) {
      sheetData[sheetName] = sheet.getDataRange().getValues();
    }
  });

  const locations = (sheetData.Locations || [[]]).slice(1).map(row => ({gov:row[0], region:row[1], market:row[2]}));
  // V42: صفحة Employees الموحّدة الأعمدة: A=id، B=name، C=username، D=password،
  // E=role (admin/manager/user)، F=jobPosition (المنصب الوظيفي)، G=mgr.
  // "role" على الواجهة = المنصب الوظيفي (مسؤول جرد/منسق نقطة/مروج) كما كانت سابقاً.
  const employees = (sheetData.Employees || [[]]).slice(1).map(row => ({
    id: String(row[0] ?? '').trim(),
    name: String(row[1] ?? '').trim(),
    username: String(row[2] ?? '').trim(),
    systemRole: String(row[4] ?? '').trim().toLowerCase(),
    role: String(row[5] ?? '').trim(),            // jobPosition = المنصب الوظيفي
    mgr: String(row[6] ?? '').trim()
  })).filter(e => e.name || e.id);

  const products = {};
  const productRows = sheetData.Products || [];
  if (productRows.length > 1) {
    const headers = productRows[0];
    const campaignCol = getColumnIndex(headers,['campaign','الحملة'],0);
    const productCol = getColumnIndex(headers,['product','المادة','اسم المادة'],1);
    const priceCol = getColumnIndex(headers,['price','السعر'],2);
    const companyCol = getColumnIndex(headers,['الشركة','company'],3);
    const barcodeCol = getColumnIndex(headers,['barcode','باركود','الباركود'],4);
    const categoryCol = getColumnIndex(headers,['تصنيف المادة','category','التصنيف'],5);
    const cancelledCol = getColumnIndex(headers,['إلغاء المادة','cancelled','canceled','الغاء المادة'],6);
    productRows.slice(1).forEach(row => {
      const campaign = String(row[campaignCol] ?? '').trim();
      const product = String(row[productCol] ?? '').trim();
      if (!campaign || !product || isCancelledValue(row[cancelledCol])) return;
      if (!products[campaign]) products[campaign] = [];
      products[campaign].push({name:product, price:Number(row[priceCol]) || 0, company:String(row[companyCol] ?? '').trim(), barcode:String(row[barcodeCol] ?? '').trim(), category:String(row[categoryCol] ?? '').trim(), cancelled:false});
    });
  }

  const expenseItems = {};
  (sheetData.ExpenseItems || [[]]).slice(1).forEach(([campaign,item]) => {
    if (!expenseItems[campaign]) expenseItems[campaign] = [];
    expenseItems[campaign].push(item);
  });

  const initialData = {locations, products, competitorProducts: buildCompetitorProductsFromRows(sheetData.ProductsOfCompetitor || []), employees, expenseItems};
  try { cache.put(cacheKey, JSON.stringify(initialData), CACHE_EXPIRATION_SECONDS); } catch (e) { /* local/browser cache remains available */ }
  return initialData;
}

function findProductByBarcode(barcode, campaign) {
  const code = String(barcode || '').trim();
  if (!code) return {status:"error", message:"Barcode is required"};

  // V25: never scan the entire Products sheet for every barcode. Reuse the
  // already-built initial-data cache and search only the relevant campaign.
  const initialData = getInitialData(false);
  const productsByCampaign = initialData.products || {};
  const wantedCampaign = String(campaign || '').trim();
  const campaigns = (!wantedCampaign || wantedCampaign === 'شاملة')
    ? Object.keys(productsByCampaign)
    : [wantedCampaign];

  for (const camp of campaigns) {
    const list = productsByCampaign[camp] || [];
    const match = list.find(p => String(p.barcode || '').trim() === code && String(p.category || '').trim() === 'مادة بيعية' && !p.cancelled);
    if (match) {
      return {status:"success", product:{campaign:camp, name:match.name, price:Number(match.price)||0, company:match.company||'', barcode:String(match.barcode||'').trim(), category:match.category||'', cancelled:false}};
    }
  }
  return {status:"error", message:"Product not found"};
}

function clearReportDetails(reportId) {
  const sheetsToProcess = [
    {sheet:salesSheet, columns:4},
    {sheet:salesOfCompetitorSheet, columns:4},
    {sheet:expensesSheet, columns:3},
    {sheet:promotersSheet, columns:3}
  ];
  sheetsToProcess.forEach(obj => {
    if (!obj.sheet || obj.sheet.getLastRow() < 2) return;
    const data = obj.sheet.getRange(1,1,obj.sheet.getLastRow(),obj.columns).getValues();
    const headers = data.shift();
    const dataToKeep = data.filter(row => String(row[0]) !== String(reportId));
    obj.sheet.clearContents();
    obj.sheet.getRange(1,1,1,headers.length).setValues([headers]);
    if (dataToKeep.length) obj.sheet.getRange(2,1,dataToKeep.length,dataToKeep[0].length).setValues(dataToKeep);
  });
}

function addReportDetails(reportData) {
  const reportId = String(reportData.id);
  if (reportData.sales && reportData.sales.length) {
    const salesRows = reportData.sales.map(s => [reportId,s.product,Number(s.price)||0,Number(s.quantity)||0]);
    salesSheet.getRange(salesSheet.getLastRow()+1,1,salesRows.length,4).setValues(salesRows);
  }
  if (reportData.salesOfCompetitor && reportData.salesOfCompetitor.length && salesOfCompetitorSheet) {
    const rows = reportData.salesOfCompetitor.map(s => [reportId,s.product,Number(s.price)||0,Number(s.quantity)||0]);
    salesOfCompetitorSheet.getRange(salesOfCompetitorSheet.getLastRow()+1,1,rows.length,4).setValues(rows);
  }
  if (reportData.expenses && reportData.expenses.length) {
    const expenseRows = reportData.expenses.map(e => [reportId,e.item,Number(e.quantity)||0]);
    expensesSheet.getRange(expensesSheet.getLastRow()+1,1,expenseRows.length,3).setValues(expenseRows);
  }
  if (reportData.promoters && reportData.promoters.length) {
    const promoterRows = reportData.promoters.filter(p => p && String(p).trim()).map(p => [reportId,p,reportData.date]);
    if (promoterRows.length) promotersSheet.getRange(promotersSheet.getLastRow()+1,1,promoterRows.length,3).setValues(promoterRows);
  }
}

function buildCompetitorProductsFromRows(rows) {
  const result = [];
  if (!Array.isArray(rows) || rows.length < 2) return result;
  const headers = rows[0];
  const productCol = getColumnIndex(headers,['product','المادة','اسم المادة'],0);
  const barcodeCol = getColumnIndex(headers,['barcode','باركود','الباركود'],1);
  const cancelledCol = getColumnIndex(headers,['إلغاء المادة','cancelled','canceled','الغاء المادة'],2);
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const product = String(row[productCol] ?? '').trim();
    if (!product || isCancelledValue(row[cancelledCol])) continue;
    result.push({product, name: product, barcode: String(row[barcodeCol] ?? '').trim(), cancelled:false});
  }
  return result;
}

// Updates master prices without rewriting the whole Products sheet.
function updateProductPrices(reportData) {
  const sheet = aSheet.getSheetByName('Products');
  if (!sheet || sheet.getLastRow() < 2 || !reportData || !Array.isArray(reportData.sales)) return;
  if (String(reportData.event || '').trim() !== 'ترويج وبيع مباشر') return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const productCol = getColumnIndex(headers,['product','المادة','اسم المادة'],1);
  const priceCol = getColumnIndex(headers,['price','السعر'],2);
  const campaignCol = getColumnIndex(headers,['campaign','الحملة'],0);

  const products = sheet.getRange(2, productCol + 1, lastRow - 1, 1).getValues();
  const campaigns = sheet.getRange(2, campaignCol + 1, lastRow - 1, 1).getValues();
  const prices = sheet.getRange(2, priceCol + 1, lastRow - 1, 1).getValues();

  const reportCampaign = String(reportData.campaign ?? '').trim();
  const requested = new Map();
  reportData.sales.forEach(s => {
    const name = String(s.product ?? '').trim();
    const price = Number(s.price);
    if (!name || !Number.isFinite(price)) return;
    const campaign = String(s.campaign ?? reportCampaign).trim();
    requested.set(`${campaign}|||${name}`, price);
  });
  if (!requested.size) return;

  let changed = false;
  for (let i = 0; i < products.length; i++) {
    const name = String(products[i][0] ?? '').trim();
    const rowCampaign = String(campaigns[i][0] ?? '').trim();
    if (!name) continue;
    let newPrice = requested.get(`${rowCampaign}|||${name}`);
    if (newPrice === undefined && reportCampaign === 'شاملة') {
      const matches = [...requested.entries()].filter(([k]) => k.endsWith(`|||${name}`));
      if (matches.length === 1) newPrice = matches[0][1];
    }
    if (newPrice !== undefined && Number(prices[i][0]) !== newPrice) {
      prices[i][0] = newPrice;
      changed = true;
    }
  }
  if (changed) {
    // One write for the price column instead of rewriting every Products cell.
    sheet.getRange(2, priceCol + 1, prices.length, 1).setValues(prices);
    SpreadsheetApp.flush();
    cache.remove('initialData_' + INITIAL_DATA_CACHE_VERSION);
    cache.remove('allReports_v3_user_history');
  }
}

function ensureReportsColumn(columnName) {
  const headers = reportsSheet.getRange(1,1,1,reportsSheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName(columnName));
  if (idx !== -1) return idx + 1;
  const newCol = reportsSheet.getLastColumn() + 1;
  reportsSheet.getRange(1,newCol).setValue(columnName);
  return newCol;
}

function updateProductPrice(payload) {
  payload = payload || {};
  // V67: تغيير الأسعار للإداري/المدير فقط.
  const updaterRole = String(payload.role || '').trim().toLowerCase();
  if (updaterRole !== 'admin' && updaterRole !== 'manager') {
    throw new Error('صلاحية تعديل أسعار المواد متاحة للإداري (admin) والمدير (manager) فقط');
  }
  const productName = String(payload.product || '').trim();
  const campaign = String(payload.campaign || '').trim();
  const price = Number(payload.price);
  const barcode = String(payload.barcode || '').trim();
  if (!productName) throw new Error('اسم المادة مطلوب');
  if (!campaign) throw new Error('نوع الحملة مطلوب');
  if (!Number.isFinite(price) || price < 0) throw new Error('السعر غير صالح');

  const sheet = aSheet.getSheetByName('Products');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Products sheet is empty');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const productCol = getColumnIndex(headers,['product','المادة','اسم المادة'],1);
  const priceCol = getColumnIndex(headers,['price','السعر'],2);
  const campaignCol = getColumnIndex(headers,['campaign','الحملة'],0);
  const barcodeCol = getColumnIndex(headers,['barcode','باركود','الباركود'],4);

  const names = sheet.getRange(2, productCol + 1, lastRow - 1, 1).getValues();
  const campaigns = sheet.getRange(2, campaignCol + 1, lastRow - 1, 1).getValues();
  const barcodes = sheet.getRange(2, barcodeCol + 1, lastRow - 1, 1).getValues();
  const prices = sheet.getRange(2, priceCol + 1, lastRow - 1, 1).getValues();

  let matches = 0;
  for (let i = 0; i < names.length; i++) {
    const rowName = String(names[i][0] ?? '').trim();
    const rowCampaign = String(campaigns[i][0] ?? '').trim();
    const rowBarcode = String(barcodes[i][0] ?? '').trim();
    const campaignMatches = campaign === 'شاملة' ? rowName === productName : (rowCampaign === campaign && rowName === productName);
    if (!campaignMatches) continue;
    if (barcode && rowBarcode && rowBarcode !== barcode) continue;
    prices[i][0] = price;
    matches++;
  }
  if (!matches) throw new Error('المادة غير موجودة ضمن الحملة المحددة');

  // Only the price column is written. This is dramatically faster than setValues on the entire sheet.
  sheet.getRange(2, priceCol + 1, prices.length, 1).setValues(prices);
  SpreadsheetApp.flush();
  cache.remove('initialData_' + INITIAL_DATA_CACHE_VERSION);
  cache.remove('allReports_v3_user_history');
  getInitialData(true);

  return {status:'success', updated:matches, price:price, cacheRefreshed:true};
}

function refreshCache() {
  cache.removeAll([
    'initialData_' + INITIAL_DATA_CACHE_VERSION,
    'allReports',
    'allReports_v2',
    'allReports_v3_user_history',
    FESTIVAL_MOVEMENTS_CACHE_KEY
  ]);

  // إعادة بناء بيانات الموقع فوراً من الـSheet، وليس فقط حذف الكاش.
  const initialData = getInitialData(true);
  return {
    status:'success',
    message:'تم تحديث جميع بيانات الموقع والكاش بنجاح',
    initialData: initialData,
    refreshedAt: new Date().toISOString()
  };
}

function handleReportSubmission(reportData) {
  reportData = reportData || {};
  ensureReportsColumn('createdById');
  ensureReportsColumn('createdByName');
  // V50: أعمدة الإثباتات وحالة الاعتماد (تُنشأ تلقائياً عند الحاجة).
  ['photo', 'signature', 'approvalStatus', 'approvedBy', 'approvedAt'].forEach(ensureReportsColumn);
  // V61: أعمدة أثر التعديل والموقع الجغرافي والحذف الناعم.
  ['editedAt', 'editedBy', 'latitude', 'longitude', 'deletedAt', 'deletedBy'].forEach(ensureReportsColumn);
  const submittedUserId = String(reportData.userId || reportData.createdById || '').trim();
  const submittedUserName = String(reportData.userName || reportData.createdByName || '').trim();
  const eventType = String(reportData.event || '').trim();
  const phoneNumber = String(reportData.phoneNumber || '').trim();
  if (eventType === 'ترويج وبيع غير مباشر' && !phoneNumber) {
    throw new Error('رقم هاتف صاحب المحل مطلوب لهذا النوع من الحدث');
  }
  ensureReportsColumn('phoneNumber');
  const headers = reportsSheet.getRange(1,1,1,reportsSheet.getLastColumn()).getValues()[0];
  const reportId = String(reportData.id);
  const reportIds = reportsSheet.getRange('A2:A').getValues().flat().map(String);
  const existingRowIndex = reportIds.indexOf(reportId);
  // عند التعديل نُبقي مالك التقرير الأصلي كما هو.
  if (existingRowIndex !== -1) {
    const existingRow = reportsSheet.getRange(existingRowIndex + 2, 1, 1, headers.length).getValues()[0];
    const ownerIdCol = headers.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('createdById'));
    const ownerNameCol = headers.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('createdByName'));
    if (ownerIdCol >= 0 && String(existingRow[ownerIdCol] || '').trim()) reportData.createdById = String(existingRow[ownerIdCol]);
    if (ownerNameCol >= 0 && String(existingRow[ownerNameCol] || '').trim()) reportData.createdByName = String(existingRow[ownerNameCol]);
    // V67: عند التعديل يُسمح للحفظ لمالك التقرير أو لإداري/مدير/مدقق فقط.
    const ownerId = ownerIdCol >= 0 ? String(existingRow[ownerIdCol] || '').trim() : '';
    const ownerName = ownerNameCol >= 0 ? String(existingRow[ownerNameCol] || '').trim() : '';
    const requesterRole = String(reportData.role || '').trim().toLowerCase();
    const isOwner = ownerId !== '' && submittedUserId !== '' && ownerId === submittedUserId;
    if (!isOwner && requesterRole !== 'admin' && requesterRole !== 'manager' && requesterRole !== 'auditor') {
      const ownerLabel = ownerName || ownerId || 'شخص آخر';
      throw new Error('لا يمكنك تعديل تقرير أنشأه ' + ownerLabel + ' — هذه الصلاحية لصاحب التقرير أو للإداري/المدير فقط');
    }
  }

  // ملكية التقرير: تُحفظ عند الإنشاء، وعند التعديل يتم الحفاظ على المالك الأصلي أعلاه.
  if (existingRowIndex === -1) {
    reportData.createdById = submittedUserId;
    reportData.createdByName = submittedUserName;
  } else {
    // V61: سجل آخر تعديل — القيمة تُرسل من الواجهة وإلا نأخذ هوية المستخدم الحالي.
    reportData.editedAt = String(reportData.editedAt || new Date().toISOString());
    reportData.editedBy = String(reportData.editedBy || submittedUserName || submittedUserId || '');
  }

  const newRow = headers.map(header => reportData[header] === undefined ? '' : (Array.isArray(reportData[header]) ? JSON.stringify(reportData[header]) : reportData[header]));

  if (existingRowIndex !== -1) {
    reportsSheet.getRange(existingRowIndex+2,1,1,newRow.length).setValues([newRow]);
    clearReportDetails(reportId);
  } else {
    reportsSheet.appendRow(newRow);
  }

  addReportDetails(reportData);
  updateProductPrices(reportData);
  syncReportExpensesToFestivalMovement(reportData);
  cache.removeAll(['allReports','allReports_v2','allReports_v3_user_history','initialData_' + INITIAL_DATA_CACHE_VERSION]);
  return {status:'success', reportId:reportId};
}

// تحويل قيم التاريخ والوقت القادمة من Google Sheets إلى صيغة مناسبة
// مباشرة لحقول HTML من نوع date / time.
// مهم عند التعديل: نقرأ القيمة الفعلية من Reports ثم نعيدها كـ yyyy-MM-dd و HH:mm.
function normalizeReportDateTimeFields(reportObj) {
  if (!reportObj) return reportObj;

  const timezone = aSheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'GMT';

  const formatDateValue = (value) => {
    if (value === null || value === undefined || value === '') return '';
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
    }
    const text = String(value).trim();
    if (!text) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

    // إذا كانت القيمة النصية ISO/تاريخاً قابلاً للتحويل، نعيدها بصيغة حقل date.
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd');
    }
    return text;
  };

  const formatTimeValue = (value) => {
    if (value === null || value === undefined || value === '') return '';
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timezone, 'HH:mm');
    }

    const text = String(value).trim();
    if (!text) return '';

    // قيم الوقت المعتادة: 09:30 أو 09:30:00
    let match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?/);
    if (match) {
      return `${String(Number(match[1])).padStart(2,'0')}:${match[2]}`;
    }

    // قد تكون القيمة ISO أو نصاً ناتجاً عن Date.
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timezone, 'HH:mm');
    }
    return text;
  };

  reportObj.date = formatDateValue(reportObj.date);
  reportObj.timeFrom = formatTimeValue(reportObj.timeFrom);
  reportObj.timeTo = formatTimeValue(reportObj.timeTo);

  return reportObj;
}

function getReports(userId, userRole, userName, targetUserId) {
  // v2 حتى لا نستخدم بيانات allReports القديمة المخزنة قبل إصلاح التاريخ والوقت.
  const cacheKey = 'allReports_v3_user_history';
  const cachedData = cache.get(cacheKey);
  let allReports;
  if (cachedData) {
    allReports = JSON.parse(cachedData);
  } else {
    if (reportsSheet.getLastRow() < 2) return [];
  const reportsData = reportsSheet.getDataRange().getValues();
  const salesData = salesSheet.getLastRow()>1 ? salesSheet.getDataRange().getValues() : [];
  const competitorSalesData = salesOfCompetitorSheet && salesOfCompetitorSheet.getLastRow()>1 ? salesOfCompetitorSheet.getDataRange().getValues() : [];
  const expensesData = expensesSheet.getLastRow()>1 ? expensesSheet.getDataRange().getValues() : [];
  const promotersData = promotersSheet.getLastRow()>1 ? promotersSheet.getDataRange().getValues() : [];
  const reportHeaders = reportsData.shift();
  const reportsMap = new Map();
  reportsData.forEach(row => {
    if (!row[0]) return;
    const reportObj = {};
    reportHeaders.forEach((header,index) => {
      if (header === 'participants' || header === 'promoters') {
        try { reportObj[header] = JSON.parse(row[index] || '[]'); } catch(e) { reportObj[header]=[]; }
      } else if (header === 'photo' || header === 'signature') {
        // V50: لا تُحمَّل الصور/التوقيعات مع قوائم التقارير (تضخيم غير ضروري).
        // تُقرأ عند الحاجة فقط عبر getReportById.
      } else reportObj[header]=row[index];
    });
    reportObj.sales=[]; reportObj.salesOfCompetitor=[]; reportObj.expenses=[];
    normalizeReportDateTimeFields(reportObj);
    reportsMap.set(String(reportObj.id),reportObj);
  });
  if (salesData.length) {
    salesData.shift();
    salesData.forEach(([id,product,price,quantity]) => { const report=reportsMap.get(String(id)); if(report) report.sales.push({product,price,quantity}); });
  }
  if (competitorSalesData.length) {
    competitorSalesData.shift();
    competitorSalesData.forEach(([id,product,price,quantity]) => { const report=reportsMap.get(String(id)); if(report) report.salesOfCompetitor.push({product,price,quantity}); });
  }
  if (expensesData.length) {
    expensesData.shift();
    expensesData.forEach(([id,item,quantity]) => { const report=reportsMap.get(String(id)); if(report) report.expenses.push({item,quantity}); });
  }
  if (promotersData.length) {
    promotersData.shift();
    promotersData.forEach(([id,name]) => { const report=reportsMap.get(String(id)); if(report){ if(!Array.isArray(report.promoters)) report.promoters=[]; if(name) report.promoters.push(name); } });
  }
    allReports=Array.from(reportsMap.values());
    // V61: استبعاد التقارير المحذوفة (حذف ناعم — deletedAt ممتلئ) من كل القوائم والكاش.
    allReports = allReports.filter(r => !String(r.deletedAt || '').trim());
    try { cache.put(cacheKey,JSON.stringify(allReports),CACHE_EXPIRATION_SECONDS); } catch(e) {}
  }

  // سجل المستخدم: نعتمد على createdById، مع fallback إلى الاسم للتقارير القديمة/التي لا تحتوي المعرّف.
  const normalizedRole = String(userRole || '').trim().toLowerCase();
  if (!userId) return allReports;
  const uid = String(userId).trim();
  const uname = String(userName || '').trim();

  // نطاق الصلاحية: null = الكل (admin)، أو مصفوفة معرّفات مسموحة (manager: فريقه، user: نفسه فقط).
  let allowedIds = getTeamMemberIds(uid, normalizedRole);

  // إذا طلب المستخدم عرض موظف/تقارير محددة عبر القائمة المنسدلة، نُضيّق النطاق إليه
  // شرط أن يكون ضمن صلاحياته أصلاً (حماية من تجاوز الصلاحيات).
  const target = String(targetUserId || '').trim();
  if (target && target !== 'all' && target !== 'الكل') {
    if (allowedIds === null || allowedIds.indexOf(target) !== -1) {
      allowedIds = [target];
    }
  }

  if (allowedIds === null) return allReports; // admin بدون تضييق إضافي

  const idsSet = new Set(allowedIds);
  // V42: نطابق أيضاً بالأسماء/أسماء المستخدمين — لأي تقرير يفتقد createdById،
  // نسمح به لو كان createdByName لأحد أعضاء الفريق (لا نقصر على المستخدم نفسه).
  const allowedNameSet = new Set();
  getUsersLite().forEach(u => {
    if (allowedIds.indexOf(u.id) !== -1) {
      if (u.nname) allowedNameSet.add(u.nname);
      if (u.nusername) allowedNameSet.add(normalizeName_(u.nusername));
    }
  });

  return allReports.filter(r => {
    const ownerId = String(r.createdById || '').trim();
    if (ownerId && idsSet.has(ownerId)) return true;
    const ownerName = String(r.createdByName || '').trim();
    return !!ownerName && allowedNameSet.has(normalizeName_(ownerName));
  });
}

function getReportById(id, requester) {
  // V67: البحث داخل نطاق صلاحية الطالب فقط (لا يكشف تقارير المستخدمين الآخرين).
  const req = requester || {};
  const report = getReports(req.id, req.systemRole, req.name, '').find(r => String(r.id) === String(id));
  if (!report) return {status:'error',message:'Report not found'};
  // V50: قراءة الصورة والتوقيع لهذا التقرير فقط (مستثناة من قوائم getReports).
  try {
    const headers = reportsSheet.getRange(1,1,1,reportsSheet.getLastColumn()).getValues()[0];
    const idx = reportsSheet.getRange('A2:A').getValues().flat().map(String).indexOf(String(id));
    if (idx !== -1) {
      const values = reportsSheet.getRange(idx + 2, 1, 1, headers.length).getValues()[0];
      headers.forEach((h,i) => { if (h === 'photo' || h === 'signature') report[h] = values[i] || ''; });
    }
  } catch (e) { /* تجاهل أخطاء القراءة التفصيلية */ }
  return {status:'success',report:report};
}

// ===============================================================
// V42: بيانات لوحة التحليلات (dashboard.html) — إحصائيات مجمّعة سريعة
// ===============================================================
function getDashboardData(userId, userRole) {
  // نعيد تقارير النطاق المسموح به فقط (نفس منطق getReports لكن بدون تفاصيل المبيعات الكاملة).
  const role = String(userRole || '').trim().toLowerCase();
  const uid = String(userId || '').trim();
  const all = getReports(uid, role, '', 'all'); // خاضع للصلاحيات admin/manager/user
  if (!Array.isArray(all)) return { status:'success', kpis:{}, daily:{}, campaigns:[], governorates:[], employees:[] };

  // ---- KPIs ----
  let totalSales = 0, totalQty = 0;
  const employeesSet = new Set();
  all.forEach(r => {
    (r.sales || []).forEach(s => {
      totalSales += (Number(s.price) || 0) * (Number(s.quantity) || 0);
      totalQty += Number(s.quantity) || 0;
    });
    if (r.createdByName) employeesSet.add(String(r.createdByName));
  });

  // ---- Daily sales (آخر 30 يوم) ----
  const dailyMap = {};
  const today = new Date();
  for (let i = 29; i >= 0; i--) {
    const d = new Date(today); d.setDate(d.getDate() - i);
    const key = Utilities.formatDate(d, aSheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone(), 'yyyy-MM-dd');
    dailyMap[key] = 0;
  }
  all.forEach(r => {
    if (!r.date) return;
    const key = String(r.date).slice(0, 10);
    if (!(key in dailyMap)) return;
    let sum = 0;
    (r.sales || []).forEach(s => sum += (Number(s.price) || 0) * (Number(s.quantity) || 0));
    dailyMap[key] += sum;
  });

  // ---- Campaign totals ----
  const campMap = {};
  all.forEach(r => {
    const c = String(r.campaign || 'غير محدد');
    if (!campMap[c]) campMap[c] = 0;
    (r.sales || []).forEach(s => campMap[c] += (Number(s.price) || 0) * (Number(s.quantity) || 0));
  });

  // ---- Governorate totals ----
  const govMap = {};
  all.forEach(r => {
    const g = String(r.governorate || 'غير محدد');
    if (!govMap[g]) govMap[g] = 0;
    (r.sales || []).forEach(s => govMap[g] += (Number(s.price) || 0) * (Number(s.quantity) || 0));
  });

  // ---- Employee performance ----
  const empMap = {};
  all.forEach(r => {
    const name = String(r.createdByName || 'غير معروف');
    if (!empMap[name]) empMap[name] = { name, reports: 0, sales: 0 };
    empMap[name].reports++;
    (r.sales || []).forEach(s => empMap[name].sales += (Number(s.price) || 0) * (Number(s.quantity) || 0));
  });
  const employees = Object.values(empMap).sort((a,b) => b.sales - a.sales).slice(0, 15);

  const zip = (obj, max = 100) => Object.entries(obj).sort((a,b) => b[1]-a[1]).slice(0, max).map(([label, value]) => ({ label, value }));

  return {
    status: 'success',
    kpis: {
      totalReports: all.length,
      totalSales: Math.round(totalSales * 100) / 100,
      totalQuantity: totalQty,
      activeEmployees: employeesSet.size
    },
    daily: Object.keys(dailyMap).map(k => ({ date: k, value: dailyMap[k] })),
    campaigns: zip(campMap),
    governorates: zip(govMap, 10),
    employees
  };
}

// ===============================================================
// V44: أرشفة التقارير القديمة يدوياً (تُستخدم عادة مرة شهرياً) — تنقل التقارير
// الأقدم من عدد الأشهر المحدد، وكل بياناتها المرتبطة (sales/salesOfCompetitor/
// expenses/promoters/festivalMovement) إلى أوراق أرشيف منفصلة (تُنشأ تلقائياً
// إن لم تكن موجودة)، ثم تُفرغ كل الكاش فوراً حتى لا يبقى أي فرق بين البيانات
// الحقيقية والمعروضة. متاحة لحساب admin فقط، لأنها عملية شاملة على كل التقارير
// بغض النظر عن الفريق (لا معنى لأرشفة جزئية لفريق واحد فقط).
// ===============================================================
function ensureArchiveSheet(name, headers) {
  let sheet = aSheet.getSheetByName(name);
  if (!sheet) {
    sheet = aSheet.insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  } else if (sheet.getLastRow() < 1) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sheet;
}

function appendRowsToArchive(archiveSheet, rows) {
  if (!rows.length) return;
  archiveSheet.getRange(archiveSheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

function archiveOldReports(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin') throw new Error('أرشفة التقارير القديمة متاحة فقط لحساب admin');

  const monthsToKeep = Number(payload.monthsToKeep);
  if (!Number.isFinite(monthsToKeep) || monthsToKeep <= 0) throw new Error('عدد الأشهر يجب أن يكون رقماً أكبر من صفر');

  const timezone = aSheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'GMT';
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - monthsToKeep);
  const cutoffStr = Utilities.formatDate(cutoff, timezone, 'yyyy-MM-dd');

  if (reportsSheet.getLastRow() < 2) return { status:'success', archivedReports:0, cutoff: cutoffStr };

  // --- 1) تقسيم ورقة Reports إلى (تُبقى) و(تُؤرشف) حسب عمود date ---
  const reportsValues = reportsSheet.getDataRange().getValues();
  const reportHeaders = reportsValues[0];
  const dateColIdx = reportHeaders.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('date'));
  const idColIdx = reportHeaders.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('id'));
  if (dateColIdx === -1 || idColIdx === -1) throw new Error('تعذر تحديد عمودي id/date في ورقة Reports');

  const dataRows = reportsValues.slice(1);
  const toArchive = [];
  const toKeep = [];
  const archivedIds = new Set();
  dataRows.forEach(row => {
    const rowDate = String(row[dateColIdx] || '').slice(0, 10);
    if (rowDate && rowDate < cutoffStr) {
      toArchive.push(row);
      archivedIds.add(String(row[idColIdx]));
    } else {
      toKeep.push(row);
    }
  });

  if (!toArchive.length) {
    return { status:'success', archivedReports:0, cutoff: cutoffStr, message: 'لا توجد تقارير أقدم من هذا التاريخ لأرشفتها' };
  }

  // --- 2) نقل ورقة Reports نفسها ---
  const reportsArchiveSheet = ensureArchiveSheet('Reports_Archive', reportHeaders);
  appendRowsToArchive(reportsArchiveSheet, toArchive);
  reportsSheet.clearContents();
  reportsSheet.getRange(1,1,1,reportHeaders.length).setValues([reportHeaders]);
  if (toKeep.length) reportsSheet.getRange(2,1,toKeep.length,reportHeaders.length).setValues(toKeep);

  // --- 3) نفس المنطق لبقية الأوراق المرتبطة (مفتاحها reportId في العمود الأول) ---
  const relatedSheets = [
    { sheet: salesSheet, archiveName: 'sales_Archive' },
    { sheet: salesOfCompetitorSheet, archiveName: 'salesOfCompetitor_Archive' },
    { sheet: expensesSheet, archiveName: 'expenses_Archive' },
    { sheet: promotersSheet, archiveName: 'promoters_Archive' }
  ];
  relatedSheets.forEach(({ sheet, archiveName }) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rows = values.slice(1);
    const keep = [];
    const archive = [];
    rows.forEach(row => {
      if (archivedIds.has(String(row[0]))) archive.push(row); else keep.push(row);
    });
    if (!archive.length) return;
    const archiveSheet = ensureArchiveSheet(archiveName, headers);
    appendRowsToArchive(archiveSheet, archive);
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    if (keep.length) sheet.getRange(2,1,keep.length,headers.length).setValues(keep);
  });

  // --- 4) نفس المنطق لورقة festivalMovement — الحركات اليدوية (سحب/مرتجع) التي لا
  //         تحمل reportId (فارغة) لا تُؤرشف أبداً لأنها غير مرتبطة بتقرير محدد أصلاً ---
  const movementSheet = aSheet.getSheetByName(FESTIVAL_MOVEMENT_SHEET_NAME);
  if (movementSheet && movementSheet.getLastRow() > 1) {
    const values = movementSheet.getDataRange().getValues();
    const headers = values[0];
    const rows = values.slice(1);
    const keep = [];
    const archive = [];
    rows.forEach(row => {
      const rid = String(row[0] || '').trim();
      if (rid && archivedIds.has(rid)) archive.push(row); else keep.push(row);
    });
    if (archive.length) {
      const archiveSheet = ensureArchiveSheet('festivalMovement_Archive', headers);
      appendRowsToArchive(archiveSheet, archive);
      movementSheet.clearContents();
      movementSheet.getRange(1,1,1,headers.length).setValues([headers]);
      if (keep.length) movementSheet.getRange(2,1,keep.length,headers.length).setValues(keep);
    }
  }

  // --- 5) تفريغ كل الكاش فوراً حتى لا يبقى أي فرق بين البيانات الحقيقية والمعروضة ---
  try {
    cache.removeAll(['allReports','allReports_v2','allReports_v3_user_history','initialData_' + INITIAL_DATA_CACHE_VERSION]);
    cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
  } catch (e) {}

  return { status:'success', archivedReports: toArchive.length, cutoff: cutoffStr };
}

function addLocation(payload) {
  payload = payload || {};
  // V67: إضافة المواقع للإداري (admin) فقط.
  const adderRole = String(payload.role || '').trim().toLowerCase();
  if (adderRole !== 'admin') throw new Error('صلاحية إضافة المواقع متاحة للإداري (admin) فقط');
  const sheet = aSheet.getSheetByName("Locations");

  if (!sheet) {
    throw new Error("Sheet 'Locations' غير موجودة");
  }

  payload = payload || {};

  // نوع الإضافة
  const type = String(
    payload.type ||
    payload.locationType ||
    payload.kind ||
    ""
  ).trim().toLowerCase();

  // نقبل أكثر من اسم حتى تتوافق مع script.js الحالي
  let gov = String(
    payload.gov ??
    payload.governorate ??
    payload.governor ??
    payload.province ??
    ""
  ).trim();

  let region = String(
    payload.region ??
    payload.area ??
    ""
  ).trim();

  let market = String(
    payload.market ??
    payload.marketName ??
    payload.shop ??
    payload.name ??
    ""
  ).trim();

  // إذا كانت الإضافة محافظة، فالقيمة قد تكون مرسلة كـ name أو value
  if (
    (type === "governorate" || type === "gov" || type === "province") &&
    !gov
  ) {
    gov = String(
      payload.name ??
      payload.value ??
      payload.text ??
      ""
    ).trim();
  }

  if (!type) {
    throw new Error("نوع البيانات المطلوب إضافتها غير محدد");
  }

  let newGov = gov;
  let newRegion = region;
  let newMarket = market;

  // =========================================================
  // إضافة محافظة
  // =========================================================
  if (
    type === "governorate" ||
    type === "gov" ||
    type === "province"
  ) {
    if (!newGov) {
      throw new Error("اسم المحافظة مطلوب");
    }

    newRegion = "";
    newMarket = "";
  }

  // =========================================================
  // إضافة منطقة
  // =========================================================
  else if (
    type === "region" ||
    type === "area"
  ) {
    if (!newGov) {
      throw new Error("يجب اختيار المحافظة أولاً");
    }

    if (!newRegion) {
      newRegion = String(
        payload.name ??
        payload.value ??
        payload.text ??
        ""
      ).trim();
    }

    if (!newRegion) {
      throw new Error("اسم المنطقة مطلوب");
    }

    newMarket = "";
  }

  // =========================================================
  // إضافة محل
  // =========================================================
  else if (
    type === "market" ||
    type === "shop"
  ) {
    if (!newGov) {
      throw new Error("يجب اختيار المحافظة أولاً");
    }

    if (!newRegion) {
      throw new Error("يجب اختيار المنطقة أولاً");
    }

if (!newMarket) {
  newMarket = String(
    payload.name ??
    payload.value ??
    payload.text ??
    ""
  ).trim();
}

if (!newMarket) {
  throw new Error("اسم المحل مطلوب");
}
  }

  else {
    throw new Error("نوع الموقع غير صحيح: " + type);
  }

  // =========================================================
  // منع التكرار
  // =========================================================

  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {

    const data = sheet
      .getRange(2, 1, lastRow - 1, 3)
      .getValues();

    const duplicate = data.some(row => {

      const rowGov = String(row[0] ?? "").trim();
      const rowRegion = String(row[1] ?? "").trim();
      const rowMarket = String(row[2] ?? "").trim();

      return (
        rowGov === newGov &&
        rowRegion === newRegion &&
        rowMarket === newMarket
      );
    });

    if (duplicate) {
      throw new Error("هذه البيانات موجودة مسبقاً");
    }
  }

  // =========================================================
  // إضافة إلى Locations
  // =========================================================

  sheet.appendRow([
    newGov,
    newRegion,
    newMarket
  ]);

  // =========================================================
  // تحديث الكاش
  // =========================================================

  try {
    cache.remove("initialData_" + INITIAL_DATA_CACHE_VERSION);
  } catch (e) {}

  return {
    status: "success",
    message: "تمت إضافة البيانات بنجاح",
    location: {
      gov: newGov,
      region: newRegion,
      market: newMarket
    }
  };
}

// ===============================================================
// سحب / مرتجع مواد (festivalMovement) — مبنية على المستخدم (createdById)
// أنواع العمليات في العمود operation:
//   "سحب" / "مرتجع"  -> تُدخل يدوياً من صفحة سحب/مرتجع المواد.
//   "صرف"            -> تُسجَّل تلقائياً من جدول المصاريف عند حفظ تقرير "ترويج وبيع مباشر".
//   "مبيعات"         -> تُسجَّل فقط يدوياً عبر زر "إرسال صافي المحصلة إلى المبيعات"
//                       (لم تعد تلقائية)، وتُصفّر محصلة المستخدم لحظة إرسالها.
// ===============================================================

const FESTIVAL_MOVEMENT_SHEET_NAME = "festivalMovement";
const FESTIVAL_MOVEMENT_HEADERS = ["reportId","item","quantity","operation","date","createdById","createdByName","invoiceNumber","id","approvalStatus","approvedBy","approvedAt","rejectionReason","rejectionBy"];
const DIRECT_SALE_EVENT = "ترويج وبيع مباشر";
const FESTIVAL_MOVEMENT_SALE_OP = "مبيعات";
const FESTIVAL_MOVEMENT_EXPENSE_OP = "صرف";

function ensureFestivalMovementSheet() {
  let sheet = aSheet.getSheetByName(FESTIVAL_MOVEMENT_SHEET_NAME);
  if (!sheet) {
    sheet = aSheet.insertSheet(FESTIVAL_MOVEMENT_SHEET_NAME);
    sheet.getRange(1,1,1,FESTIVAL_MOVEMENT_HEADERS.length).setValues([FESTIVAL_MOVEMENT_HEADERS]);
    return sheet;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1,1,1,FESTIVAL_MOVEMENT_HEADERS.length).setValues([FESTIVAL_MOVEMENT_HEADERS]);
    return sheet;
  }
  // ترقية تلقائية وآمنة لأي عمود جديد ناقص (مثل invoiceNumber) دون المساس بالبيانات الحالية.
  const existingHeaders = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => String(h ?? '').trim());
  FESTIVAL_MOVEMENT_HEADERS.forEach((headerName, idx) => {
    const exists = existingHeaders.some(h => normalizeHeaderName(h) === normalizeHeaderName(headerName));
    if (!exists) sheet.getRange(1, idx + 1).setValue(headerName);
  });
  // V66: معرّف حتمي لكل صف بلا معرّف: mv-<رقم صف البيانات> (مستقر لأن إعادة الكتابة تحافظ عليه).
  if (sheet.getLastRow() >= 2) {
    const lastCol = Math.max(sheet.getLastColumn(), FESTIVAL_MOVEMENT_HEADERS.length);
    const allHeaders = sheet.getRange(1,1,1,lastCol).getValues()[0];
    const idIdx = allHeaders.findIndex(h => normalizeHeaderName(String(h || '')) === normalizeHeaderName('id'));
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
    for (let i = 0; i < data.length; i++) {
      if (idIdx >= 0 && !String(data[i][idIdx] || '').trim()) {
        sheet.getRange(i + 2, idIdx + 1).setValue('mv-' + (i + 2));
      }
    }
  }
  return sheet;
}

const FESTIVAL_MOVEMENTS_CACHE_KEY = 'allFestivalMovements_v2';

// قراءة كل حركات الشيت مرة واحدة وتخزينها في الكاش (سريعة)، ثم نُفلترها بالذاكرة
// حسب صلاحية كل طالب بدل قراءة الشيت من جديد في كل طلب.
function getAllFestivalMovementsCached() {
  const cached = cache.get(FESTIVAL_MOVEMENTS_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const sheet = ensureFestivalMovementSheet();
  let movements = [];
  if (sheet.getLastRow() >= 2) {
    const timezone = aSheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'GMT';
    const lastCol = Math.max(sheet.getLastColumn(), FESTIVAL_MOVEMENT_HEADERS.length);
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
    movements = data.map(row => {
      const dateValue = row[4];
      const dateText = (dateValue instanceof Date && !isNaN(dateValue.getTime()))
        ? Utilities.formatDate(dateValue, timezone, 'yyyy-MM-dd HH:mm')
        : String(dateValue || '');
      return {
        reportId: String(row[0] || ''),
        item: String(row[1]),
        quantity: Number(row[2]) || 0,
        operation: String(row[3]),
        date: dateText,
        rawDate: (dateValue instanceof Date) ? dateValue.getTime() : 0,
        createdById: String(row[5] || ''),
        createdByName: String(row[6] || ''),
        invoiceNumber: String(row[7] || ''),
        id: String(row[8] || ''),
        approvalStatus: String(row[9] || ''),
        approvedBy: String(row[10] || ''),
        approvedAt: String(row[11] || ''),
        rejectionReason: String(row[12] || ''),
        rejectionBy: String(row[13] || '')
      };
    });
  }
  try { cache.put(FESTIVAL_MOVEMENTS_CACHE_KEY, JSON.stringify(movements), CACHE_EXPIRATION_SECONDS); } catch (e) {}
  return movements;
}

// كل حركات مستخدم/فريق ضمن صلاحية الطالب (سحب/مرتجع يدوي + صرف تلقائي + مبيعات يدوية عبر زر التصفير)، الأحدث أولاً.
// admin: يرى الجميع أو يضيّق لموظف عبر targetUserId. manager: فريقه فقط (نفسه + كامل التسلسل الهرمي من صفحة Employees: عمود mgr).
// user: نفسه فقط دائماً بغض النظر عن targetUserId.
function getUserFestivalMovements(userId, userRole, targetUserId) {
  const id = String(userId || '').trim();
  if (!id) throw new Error('معرّف المستخدم مطلوب');
  const normalizedRole = String(userRole || '').trim().toLowerCase();

  let allowedIds = getTeamMemberIds(id, normalizedRole); // null = الكل (admin)
  const target = String(targetUserId || '').trim();
  if (target && target !== 'all' && target !== 'الكل') {
    if (allowedIds === null || allowedIds.indexOf(target) !== -1) {
      allowedIds = [target];
    }
  }

  const all = getAllFestivalMovementsCached();
  const movements = (allowedIds === null)
    ? all.slice()
    : (() => { const s = new Set(allowedIds); return all.filter(m => s.has(m.createdById)); })();

  movements.sort((a,b) => b.rawDate - a.rawDate);
  const cleaned = movements.map(m => { const c = Object.assign({}, m); delete c.rawDate; return c; });
  return { status:'success', movements: cleaned };
}

// يحسب لكل مادة: إجمالي السحب/المرتجع/الصرف التلقائي/المبيعات اليدوية، والمتبقي الصافي
// = سحب - مرتجع - صرف - مبيعات، لمستخدم معيّن.
function computeUserFestivalSummary(userId) {
  const id = String(userId || '').trim();
  if (!id) return [];
  const sheet = ensureFestivalMovementSheet();
  if (sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,FESTIVAL_MOVEMENT_HEADERS.length).getValues();
  const map = new Map();
  data.filter(row => String(row[5]).trim() === id).forEach(row => {
    const item = String(row[1] || '').trim();
    if (!item) return;
    if (!map.has(item)) map.set(item, { item, withdrawn:0, returned:0, expensed:0, sold:0 });
    const entry = map.get(item);
    const qty = Number(row[2]) || 0;
    const operation = String(row[3] || '').trim();
    if (operation === 'سحب') entry.withdrawn += qty;
    else if (operation === 'مرتجع') entry.returned += qty;
    else if (operation === FESTIVAL_MOVEMENT_EXPENSE_OP) entry.expensed += qty;
    else if (operation === FESTIVAL_MOVEMENT_SALE_OP) entry.sold += qty;
  });

  return Array.from(map.values()).map(e => ({ ...e, remaining: e.withdrawn - e.returned - e.expensed - e.sold }));
}

// سجل يدوي لسحب/مرتجع مواد مرتبط بالمستخدم الحالي (وليس بتقرير معيّن بالضرورة)، مع رقم فاتورة اختياري.
function addFestivalMovement(payload) {
  payload = payload || {};
  const createdById = String(payload.createdById || '').trim();
  if (!createdById) throw new Error('تعذر تحديد هوية المستخدم');
  const createdByName = String(payload.createdByName || '').trim();
  const reportId = String(payload.reportId || '').trim(); // اختياري: يبقى فارغاً للحركات اليدوية

  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) throw new Error('يرجى إضافة مادة واحدة على الأقل');

  const now = new Date();
  const rows = items.map(it => {
    const item = String((it && it.item) || '').trim();
    const quantity = Number(it && it.quantity);
    const operation = String((it && it.operation) || '').trim();
    const invoiceNumber = String((it && it.invoiceNumber) || '').trim();
    if (!item) throw new Error('اسم المادة مطلوب');
    if (!Number.isFinite(quantity) || quantity <= 0) throw new Error('الكمية يجب أن تكون رقماً أكبر من صفر (المادة: ' + item + ')');
    if (operation !== 'سحب' && operation !== 'مرتجع') throw new Error('العملية يجب أن تكون "سحب" أو "مرتجع" (المادة: ' + item + ')');
    return [reportId, item, quantity, operation, now, createdById, createdByName, invoiceNumber];
  });

  const sheet = ensureFestivalMovementSheet();
  const startRow = sheet.getLastRow() + 1;
  const rowsFull = rows.map((r, i) => r.concat(['mv-' + (startRow + i), '', '', '', '', '']));
  sheet.getRange(startRow, 1, rowsFull.length, FESTIVAL_MOVEMENT_HEADERS.length).setValues(rowsFull);
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);

  return { status:'success', added: rowsFull.length };
}

// تُستدعى تلقائياً من handleReportSubmission عند حفظ أو تعديل أي تقرير.
// تُسجّل كميات جدول المصاريف كحركة "صرف" باسم صاحب التقرير، فتُطرح تلقائياً من محصلته.
// عند تعديل تقرير موجود، تُستبدل حركات "صرف" السابقة لنفس التقرير بالقيم الجديدة (بدون تكرار).
function syncReportExpensesToFestivalMovement(reportData) {
  reportData = reportData || {};

  const reportId = String(reportData.id || '').trim();
  if (!reportId) return;

  const ownerId = String(reportData.createdById || '').trim();
  const ownerName = String(reportData.createdByName || '').trim();
  const expenses = Array.isArray(reportData.expenses) ? reportData.expenses : [];

  const sheet = ensureFestivalMovementSheet();
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2,1,lastRow-1,FESTIVAL_MOVEMENT_HEADERS.length).getValues() : [];

  // إزالة أي حركات "صرف" مسجلة سابقاً لنفس التقرير (لتفادي تكرارها عند التعديل).
  const kept = existingRows.filter(row => !(String(row[0]).trim() === reportId && String(row[3]).trim() === FESTIVAL_MOVEMENT_EXPENSE_OP));

  const now = new Date();
  const newExpenseRows = expenses
    .map(e => ({ item: String((e && e.item) || '').trim(), quantity: Number(e && e.quantity) || 0 }))
    .filter(e => e.item && e.quantity > 0)
    .map(e => [reportId, e.item, e.quantity, FESTIVAL_MOVEMENT_EXPENSE_OP, now, ownerId, ownerName, '', '', '', '', '', '', '']);

  const finalRows = kept.concat(newExpenseRows);

  sheet.clearContents();
  sheet.getRange(1,1,1,FESTIVAL_MOVEMENT_HEADERS.length).setValues([FESTIVAL_MOVEMENT_HEADERS]);
  if (finalRows.length) {
    sheet.getRange(2,1,finalRows.length,FESTIVAL_MOVEMENT_HEADERS.length).setValues(finalRows);
  }
  // V66: تملئة معرّفات أي صفوف جديدة بلا معرّف (اضطراب/ديمومة معرّفات الصفوف الباقية محفوظة).
  ensureFestivalMovementSheet();
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
}

// تُستدعى يدوياً فقط (زر "إرسال صافي المحصلة إلى المبيعات") لكل مادة معها صافٍ موجب:
// تُضاف حركة "مبيعات" بقيمة الصافي الحالي بالضبط، فتصبح محصلة تلك المادة صفراً فور الإرسال.
function closeOutUserTallyToSales(payload) {
  payload = payload || {};
  const createdById = String(payload.createdById || '').trim();
  if (!createdById) throw new Error('تعذر تحديد هوية المستخدم');
  const createdByName = String(payload.createdByName || '').trim();
  const selectedReportId = String(payload.reportId || '').trim();
  if (!selectedReportId) throw new Error('يرجى اختيار التقرير الذي ستضاف إليه المبيعات');

  const summary = computeUserFestivalSummary(createdById);
  const itemsToClose = summary.filter(e => e.remaining > 0);
  if (!itemsToClose.length) return { status:'success', added: 0, message: 'لا توجد كميات متبقية (صافي موجب) لإرسالها' };

  // نقرأ مبيعات التقرير المختار للحصول على السعر الصحيح لكل مادة.
  // إذا لم تكن المادة موجودة في التقرير، نستخدم السعر الحالي من Products كاحتياط.
  const report = getReports('', 'admin', '').find(r => String(r.id) === selectedReportId);
  if (!report) throw new Error('التقرير المحدد غير موجود');

  const reportPriceMap = new Map();
  (Array.isArray(report.sales) ? report.sales : []).forEach(s => {
    const product = String(s.product || '').trim();
    if (!product) return;
    const price = Number(s.price);
    if (Number.isFinite(price)) reportPriceMap.set(product, price);
  });

  const productsSheet = aSheet.getSheetByName('Products');
  const productPriceMap = new Map();
  if (productsSheet && productsSheet.getLastRow() > 1) {
    const values = productsSheet.getDataRange().getValues();
    const headers = values[0] || [];
    const productCol = getColumnIndex(headers, ['product','المادة','اسم المادة'], 1);
    const priceCol = getColumnIndex(headers, ['price','السعر'], 2);
    const cancelledCol = getColumnIndex(headers, ['إلغاء المادة','cancelled','canceled','الغاء المادة'], 6);
    values.slice(1).forEach(row => {
      const product = String(row[productCol] || '').trim();
      if (!product || isCancelledValue(row[cancelledCol])) return;
      const price = Number(row[priceCol]);
      if (Number.isFinite(price) && !productPriceMap.has(product)) productPriceMap.set(product, price);
    });
  }

  const now = new Date();
  const salesRows = itemsToClose.map(e => {
    const price = reportPriceMap.has(e.item) ? reportPriceMap.get(e.item) : (productPriceMap.get(e.item) || 0);
    return [selectedReportId, e.item, Number(price) || 0, e.remaining];
  });

  // 1) تسجَّل المبيعات في sales مع reportId للتقرير الذي اختاره المستخدم.
  if (!salesSheet) throw new Error("Sheet 'sales' غير موجودة");
  salesSheet.getRange(salesSheet.getLastRow() + 1, 1, salesRows.length, 4).setValues(salesRows);

  // 2) وتسجَّل في festivalMovement كحركة مبيعات مرتبطة بنفس التقرير.
  const movementSheet = ensureFestivalMovementSheet();
  const movementRows = itemsToClose.map((e, i) => [selectedReportId, e.item, e.remaining, FESTIVAL_MOVEMENT_SALE_OP, now, createdById, createdByName, '', 'mv-' + (movementSheet.getLastRow() + 1 + i), '', '', '', '', '']);
  movementSheet.getRange(movementSheet.getLastRow() + 1, 1, movementRows.length, FESTIVAL_MOVEMENT_HEADERS.length).setValues(movementRows);

  SpreadsheetApp.flush();
  cache.remove('allReports_v3_user_history');
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
  return { status:'success', added: movementRows.length, salesAdded: salesRows.length, reportId: selectedReportId };
}

// ===================================================================
// V50: اعتماد التقارير — يرصد تقريراً كـ «معتمد» مع اسم المدقق ووقته.
// ملاحظة: تتطلب هذه الدالة إعادة نشر (Deploy) لمشروع Apps Script بعد
// تحديث code.gs حتى يعمل زر «اعتماد» في صفحة سجل التعديلات.
// (الصيانة المحلية: role auditor أو admin فقط؛ أي دور آخر مرفوض).
// ===================================================================

// V70: اعتماد جماعي للتقارير — يعتمد كل ما هو قيد المراجعة ضمن نطاق المستخدم، ويترك المرفوض كما هو.
function approveAllReports(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager') throw new Error('اعتماد الكل متاح للإداري (admin) والمدير (manager) فقط');
  const reports = getReports(String(payload.userId || ''), role, String(payload.userName || ''), String(payload.targetUserId || '')) || [];
  const pending = reports.filter(r => String(r.approvalStatus || '').trim().toLowerCase() !== 'approved' && String(r.approvalStatus || '').trim().toLowerCase() !== 'rejected');
  let approved = 0;
  pending.forEach(r => { approveReport({ id: String(r.id), role, approvedBy: payload.userName || '', approvedAt: new Date().toISOString() }); approved++; });
  return { status:'success', approved, skippedRejected: reports.filter(r => String(r.approvalStatus || '').trim().toLowerCase() === 'rejected').length, message: approved ? `تم اعتماد ${approved} تقرير.` : 'لا توجد تقارير قيد المراجعة للاعتماد.' };
}

// V70: اعتماد جماعي للدوام — يعتمد كل ما هو قيد المراجعة ضمن نطاق المستخدم، ويترك المرفوض كما هو.
function approveAllAttendance(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager') throw new Error('اعتماد الكل متاح للإداري (admin) والمدير (manager) فقط');
  const rows = getAttendance(String(payload.userId || ''), role, String(payload.userName || ''), String(payload.targetUserId || '')) || [];
  const pending = rows.filter(r => String(r.approvalStatus || '').trim().toLowerCase() !== 'approved' && String(r.approvalStatus || '').trim().toLowerCase() !== 'rejected');
  let approved = 0;
  pending.forEach(r => { approveAttendance({ timestamp:String(r.timestamp), username:String(r.username), role, approvedBy:payload.userName || '', approvedAt:new Date().toISOString(), status:'approved' }); approved++; });
  return { status:'success', approved, skippedRejected: rows.filter(r => String(r.approvalStatus || '').trim().toLowerCase() === 'rejected').length, message: approved ? `تم اعتماد ${approved} سجل دوام.` : 'لا توجد سجلات دوام قيد المراجعة للاعتماد.' };
}

// V70: اعتماد جماعي للحركات — يعتمد كل ما هو قيد المراجعة ضمن نطاق المستخدم، ويترك المرفوض كما هو.
function approveAllMovements(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager') throw new Error('اعتماد الكل متاح للإداري (admin) والمدير (manager) فقط');
  const result = getUserFestivalMovements(String(payload.userId || ''), role, String(payload.targetUserId || '')) || { movements:[] };
  const movements = Array.isArray(result.movements) ? result.movements : [];
  const pending = movements.filter(m => String(m.approvalStatus || '').trim().toLowerCase() !== 'approved' && String(m.approvalStatus || '').trim().toLowerCase() !== 'rejected');
  let approved = 0;
  pending.forEach(m => { approveMovement({ id:String(m.id), role, approvedBy:payload.userName || '', approvedAt:new Date().toISOString(), status:'approved' }); approved++; });
  cache.remove(FESTIVAL_MOVEMENTS_CACHE_KEY);
  return { status:'success', approved, skippedRejected: movements.filter(m => String(m.approvalStatus || '').trim().toLowerCase() === 'rejected').length, message: approved ? `تم اعتماد ${approved} حركة.` : 'لا توجد حركات قيد المراجعة للاعتماد.' };
}

function approveReport(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'auditor' && role !== 'admin' && role !== 'manager') {
    throw new Error('صلاحية الاعتماد متاحة للمدقق (auditor) والإداري (admin) والمدير (manager) فقط');
  }
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('معرّف التقرير مطلوب');
  // V63: دعم الحالتين معتمد/مرفوض + سبب الرفض.
  const status = String(payload.status || payload.approvalStatus || 'approved').trim().toLowerCase();
  if (status !== 'approved' && status !== 'rejected') throw new Error('حالة غير صالحة (approved أو rejected)');
  ensureReportsColumn('approvalStatus');
  ensureReportsColumn('approvedBy');
  ensureReportsColumn('approvedAt');
  ensureReportsColumn('rejectionReason');
  ensureReportsColumn('rejectionBy');
  const headers = reportsSheet.getRange(1, 1, 1, reportsSheet.getLastColumn()).getValues()[0];
  const row = reportsSheet.getRange('A2:A').getValues().flat().map(String).indexOf(id);
  if (row === -1) throw new Error('التقرير غير موجود');
  const targetRow = row + 2;
  const by = String(payload.approvedBy || payload.userName || '');
  const now = String(payload.approvedAt || new Date().toISOString());
  const valuesByHeader = {
    approvalStatus: status,
    approvedBy: by,
    approvedAt: now,
    rejectionReason: status === 'rejected' ? String(payload.reason || payload.rejectionReason || '') : '',
    rejectionBy: status === 'rejected' ? by : ''
  };
  Object.keys(valuesByHeader).forEach(h => {
    const col = headers.findIndex(x => normalizeHeaderName(x) === normalizeHeaderName(h));
    if (col >= 0) reportsSheet.getRange(targetRow, col + 1).setValue(valuesByHeader[h]);
  });
  cache.remove('allReports_v3_user_history');
  cache.remove('allReports_v2');
  cache.remove('allReports');
  return { status: 'success', reportId: id, approvalStatus: status };
}

// V61: حذف ناعم لتقرير — متاح للإداري (admin) والمدير (manager) والمدقق (auditor).
// يُملأ deletedAt/deletedBy بدل إزالة الصف، وتُفرَّغ كواش التقارير فوراً.
function deleteReport(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager' && role !== 'auditor') {
    throw new Error('صلاحية الحذف متاحة للإداري (admin) والمدير (manager) والمدقق (auditor) فقط');
  }
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('معرّف التقرير مطلوب');
  const ids = reportsSheet.getRange('A2:A').getValues().flat().map(String);
  const row = ids.indexOf(id);
  if (row === -1) throw new Error('التقرير غير موجود');
  const targetRow = row + 2;
  ensureReportsColumn('deletedAt');
  ensureReportsColumn('deletedBy');
  const freshHeaders = reportsSheet.getRange(1, 1, 1, reportsSheet.getLastColumn()).getValues()[0];
  const colAt = freshHeaders.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('deletedAt'));
  const colBy = freshHeaders.findIndex(h => normalizeHeaderName(h) === normalizeHeaderName('deletedBy'));
  if (colAt >= 0) reportsSheet.getRange(targetRow, colAt + 1).setValue(new Date().toISOString());
  if (colBy >= 0) reportsSheet.getRange(targetRow, colBy + 1).setValue(String(payload.userName || payload.userId || ''));
  clearReportDetails(id);
  cache.removeAll(['allReports_v3_user_history', 'allReports_v2', 'allReports']);
  return { status: 'success', reportId: id };
}
// ---------- V63: أهداف المروجين الشهرية ----------
function getPromoterGoalsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('PromoterGoals');
  if (!sh) {
    sh = ss.insertSheet('PromoterGoals');
    sh.getRange('A1:D1').setValues([['month', 'promoter', 'points', 'pieces']]);
    sh.getRange('A1:D1').setFontWeight('bold');
  }
  return sh;
}

function getPromoterGoals(month, userId, role) {
  const m = String(month || '').trim();
  if (!m) throw new Error('الشهر مطلوب بصيغة YYYY-MM');
  const sh = getPromoterGoalsSheet_();
  const values = sh.getRange('A2:D').getValues().filter(r => String(r[0] || '') === m);
  return {
    status: 'success',
    goals: values.map(r => ({ promoter: String(r[1] || ''), points: Number(r[2]) || 0, pieces: Number(r[3]) || 0 }))
  };
}

function savePromoterGoals(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'manager') {
    throw new Error('صلاحية تحديد الأهداف متاحة للإداري (admin) والمدير (manager) فقط');
  }
  const m = String(payload.month || '').trim();
  if (!m) throw new Error('الشهر مطلوب بصيغة YYYY-MM');
  const sh = getPromoterGoalsSheet_();
  sh.getRange('A2:A' + Math.max(2, sh.getLastRow())).clearContent();
  const entries = Array.isArray(payload.entries) ? payload.entries : [];
  if (entries.length && entries.length > 0) {
    sh.getRange(2, 1, entries.length, 4).setValues(entries.map(e => [
      m,
      String(e.promoter || '').trim(),
      Number(e.points) || 0,
      Number(e.pieces) || 0
    ]));
  }
  return { status: 'success', message: 'تم حفظ أهداف الشهر.', month: m, saved: entries.length };
}

// ---------- V63: النسخ الاحتياطي الأسبوعي ----------
function createBackupSpreadsheet_() {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), source.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HH-mm');
  const backup = SpreadsheetApp.create('FieldReports_BACKUP_' + stamp);
  source.getSheets().forEach(function (sh) {
    if (sh.getName() === 'PromoterGoals') return;
    let target = null;
    try {
      target = backup.getSheetByName(sh.getName()) || backup.insertSheet(sh.getName());
    } catch (err) {
      target = backup.getSheetByName('Sheet1') || backup.insertSheet('Sheet1');
    }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      target.getRange(1, 1, lastRow, lastCol).setValues(sh.getRange(1, 1, lastRow, lastCol).getValues());
    }
    target.getRange(1, 1, 1, Math.max(1, lastCol)).setFontWeight('bold');
  });
  try { backup.deleteSheet(backup.getSheetByName('Sheet1')); } catch (err) {}
  return { url: backup.getUrl() };
}

function setupWeeklyBackup(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin') throw new Error('صلاحية تفعيل النسخ الاحتياطي متاحة للإداري (admin) فقط');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('weeklyBackupEnabled', 'true');
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyBackupTrigger_');
  if (triggers.length === 0) {
    ScriptApp.newTrigger('weeklyBackupTrigger_').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  }
  const active = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyBackupTrigger_').length;
  return { status: 'success', message: 'تم تفعيل النسخ الاحتياطي الأسبوعي (كل يوم أحد الساعة 2 صباحًا).', activeTriggers: active };
}

function weeklyBackupTrigger_() {
  try {
    const enabled = PropertiesService.getScriptProperties().getProperty('weeklyBackupEnabled');
    if (enabled !== 'true') return;
    createBackupSpreadsheet_();
  } catch (err) {
    console.error('weekly backup failed: ' + err.message);
  }
}

function backupNow(payload) {
  payload = payload || {};
  const role = String(payload.role || '').trim().toLowerCase();
  if (role !== 'admin') throw new Error('صلاحية إنشاء نسخة احتياطية متاحة للإداري (admin) فقط');
  const result = createBackupSpreadsheet_();
  return { status: 'success', message: 'تم إنشاء النسخة الاحتياطية.', url: result.url };
}

// ===============================================================
// V68: إدارة المستخدمين والصلاحيات — للإداري (admin) فقط.
// الأعمدة: A=id، B=name، C=username، D=password (hash)، E=systemRole،
// F=jobPosition (المنصب الوظيفي)، G=mgr (اسم المدير المباشر).
// لا تُكشف كلمات المرور أبداً — admin يضبطها فقط (تُخزَّن hash).
// ===============================================================

// قراءة صفوف الموظفين (المعرّف + اسم الدخول + الدور) لفحص «آخر إداري».
function employeesRows_() {
  const sheet = aSheet.getSheetByName('Employees');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
}

function invalidateInitialDataCache_() {
  try { cache.remove('initialData_' + INITIAL_DATA_CACHE_VERSION); } catch (e) {}
}

function getUsersAdmin(requester) {
  requester = requester || {};
  if (String(requester.systemRole || '').trim().toLowerCase() !== 'admin') {
    throw new Error('صلاحية إدارة المستخدمين مخصصة للإداري (admin) فقط');
  }
  const sheet = aSheet.getSheetByName('Employees');
  if (!sheet || sheet.getLastRow() < 2) return { status: 'success', users: [] };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return {
    status: 'success',
    users: data.map(row => ({
      id: String(row[0] ?? '').trim(),
      name: String(row[1] ?? '').trim(),
      username: String(row[2] ?? '').trim(),
      systemRole: String(row[4] ?? '').trim().toLowerCase(),
      jobPosition: String(row[5] ?? '').trim(),
      mgr: String(row[6] ?? '').trim(),
      hasPassword: !!String(row[3] ?? '').trim()
    })).filter(u => u.name || u.username || u.id)
  };
}

// إنشاء مستخدم جديد أو تعديل مستخدم موجود (editUsername = اسم الدخول الأصلي).
function saveUser(payload) {
  payload = payload || {};
  const actorRole = String(payload.role || '').trim().toLowerCase();
  if (actorRole !== 'admin') throw new Error('صلاحية إدارة المستخدمين مخصصة للإداري (admin) فقط');
  const norm = s => String(s || '').trim().toLowerCase();
  const actorUsername = String(payload.actorUsername || '').trim().toLowerCase();
  const editUsername = String(payload.editUsername || '').trim().toLowerCase();
  const name = String(payload.name || '').trim();
  const username = String(payload.username || '').trim().toLowerCase();
  if (!name) throw new Error('اسم الموظف مطلوب');
  if (!username) throw new Error('اسم الدخول (username) مطلوب');
  if (!/^[a-zA-Z0-9._\-]{3,50}$/.test(username)) throw new Error('اسم الدخول يجب أن يكون 3-50 حرفاً إنجليزياً/رقمياً بدون مسافات');
  const systemRole = String(payload.systemRole || '').trim().toLowerCase();
  const allowedRoles = ['admin', 'manager', 'auditor', 'user'];
  if (allowedRoles.indexOf(systemRole) === -1) throw new Error('صلاحية غير معروفة — اختر من القائمة');
  const jobPosition = String(payload.jobPosition || '').trim();
  const mgr = String(payload.manager || '').trim();
  const newPassword = String(payload.password || '').trim();

  const rows = employeesRows_();
  let targetIndex = -1;
  if (editUsername) {
    for (let i = 0; i < rows.length; i++) {
      if (norm(rows[i][2]) === editUsername) { targetIndex = i; break; }
    }
    if (targetIndex === -1) throw new Error('الحساب المطلوب تعديله غير موجود');
  }

  // فحص تفرد اسم الدخول (مع استثناء الصف نفسه عند التعديل).
  for (let i = 0; i < rows.length; i++) {
    if (i === targetIndex) continue;
    if (norm(rows[i][2]) === username) throw new Error('اسم الدخول مستخدم من قبل — اختر اسماً آخر');
  }

  // حماية: آخر إداري لا يُحذف ولا يُسحب، والإداري لا يسحب صلاحيته عن نفسه.
  const countAdmins = rows.filter(r => norm(r[4]) === 'admin').length;
  const targetRow = targetIndex >= 0 ? rows[targetIndex] : null;
  if (targetRow !== null) {
    const isEditSelf = actorUsername !== '' && norm(targetRow[2]) === actorUsername;
    if (norm(targetRow[4]) === 'admin' && systemRole !== 'admin' && countAdmins <= 1) {
      throw new Error('لا يمكن إزالة آخر حساب إداري — أضف إدارياً آخر أولاً');
    }
    if (isEditSelf && systemRole !== 'admin') throw new Error('لا يمكنك سحب صلاحية الإداري عن حسابك');
  }

  // كلمة المرور: مطلوبة للمستخدم الجديد، اختيارية عند التعديل (فارغة = إبقاء الحالية).
  let newHash = null;
  if (targetIndex === -1) {
    if (!newPassword) throw new Error('كلمة المرور مطلوبة لإنشاء حساب جديد');
    if (newPassword.length < 4) throw new Error('كلمة المرور لا تقل عن 4 أحرف');
    newHash = hashPassword_(newPassword);
  } else if (newPassword) {
    if (newPassword.length < 4) throw new Error('كلمة المرور لا تقل عن 4 أحرف');
    newHash = hashPassword_(newPassword);
  }

  const sheet = aSheet.getSheetByName('Employees');
  if (targetIndex === -1) {
    const rowNum = sheet.getLastRow() + 1;
    const hash = newHash ? ('pwd$' + newHash.salt + '$' + newHash.hash) : '';
    // A=id(فارغ)، B=name، C=username، D=password(hash)، E=systemRole، F=jobPosition، G=mgr.
    sheet.getRange(rowNum, 1, 1, 7).setValues([['', name, username, hash, systemRole, jobPosition, mgr]]);
    invalidateInitialDataCache_();
    return { status: 'success', message: 'تم إنشاء الحساب «' + name + '» بنجاح.' };
  }
  const rowNum = targetIndex + 2;
  // V68: نكتب الأعمدة الصحيحة فقط، مع ترك عمود كلمة المرور (D) دون مسّ ما لم تُحدَّث:
  // B=name، C=username ثم E=systemRole، F=jobPosition، G=mgr (نكسر الاستمرارية لتجنّب الكتابة على D).
  sheet.getRange(rowNum, 2, 1, 2).setValues([[name, username]]);
  sheet.getRange(rowNum, 5, 1, 3).setValues([[systemRole, jobPosition, mgr]]);
  if (newHash) sheet.getRange(rowNum, 4).setValue('pwd$' + newHash.salt + '$' + newHash.hash);
  invalidateInitialDataCache_();
  return { status: 'success', message: 'تم تحديث الحساب «' + name + '» بنجاح.' };
}

function deleteUser(payload) {
  payload = payload || {};
  const actorRole = String(payload.role || '').trim().toLowerCase();
  if (actorRole !== 'admin') throw new Error('صلاحية إدارة المستخدمين مخصصة للإداري (admin) فقط');
  const norm = s => String(s || '').trim().toLowerCase();
  const actorUsername = String(payload.actorUsername || '').trim().toLowerCase();
  const editUsername = String(payload.editUsername || '').trim().toLowerCase();
  if (!editUsername) throw new Error('معرّف الحساب مطلوب');
  const rows = employeesRows_();
  let targetIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (norm(rows[i][2]) === editUsername) { targetIndex = i; break; }
  }
  if (targetIndex === -1) throw new Error('الحساب غير موجود');
  if (actorUsername !== '' && norm(rows[targetIndex][2]) === actorUsername) throw new Error('لا يمكنك حذف حسابك بنفسك');
  const countAdmins = rows.filter(r => norm(r[4]) === 'admin').length;
  if (norm(rows[targetIndex][4]) === 'admin' && countAdmins <= 1) throw new Error('لا يمكن حذف آخر حساب إداري');
  aSheet.getSheetByName('Employees').deleteRow(targetIndex + 2);
  invalidateInitialDataCache_();
  return { status: 'success', message: 'تم حذف الحساب.' };
}