/** =========================================================
 *  HR DEPARTMENT CERTIFICATE OF EMPLOYMENT (COE) GENERATOR
 *  ========================================================= */

const CONFIG = {
  TEMPLATE_DOC_ID: '18-WDBFOX1vl3dmok5bBT-0X66Ikug88byJCzK7sUEP8',
  WORKFORCE_SHEET_ID: '1-JXPPzcD1JU6MlBKhLD7PNqJ4SM18RLASw3_tsivReg',
  WORKFORCE_SHEET_NAME: 'Main Office',
  LOG_SHEET_ID: '1xPBKdfm-z0ZKM_thflck1VJXKg6I52SGKDu2UdLrl94',
  LOG_SHEET_NAME: 'COE_Issuance_Log',

  OUTPUT_FOLDER_ID: '12IEqqmsyTGuUO0ycGieKlq77C02OKZzP',
  KEEP_DOC_COPY: false, 

  COE_PREFIX: 'COE',

  SIGNATORY: {
    name: 'JOHN M. CRUZ', 
    title1: 'Department Head',
    title2: 'Human Resource Manager'
  },

  // TODO: Replace these with your exact wording.
  // Available tokens: <<FULL_NAME>> <<POSITION>> <<DEPARTMENT>> <<STATUS>>
  //                    <<DATE_HIRED>> <<SALARY>>
  CLAUSES: {
    withCompensation:
      'is a bona fide employee of D.N. Kairos, holding the position of ' +
      '<<POSITION>> in the <<DEPARTMENT>>, under <<STATUS>> status, since <<DATE_HIRED>> up to Present, ' +
      'with a monthly salary of <<SALARY>>.',
    withoutCompensation:
      'is a bona fide employee of D.N. Kairos, holding the position of ' +
      '<<POSITION>> in the <<DEPARTMENT>>, under <<STATUS>> status, since <<DATE_HIRED>> up to Present. ' 
  }
};
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Certificate of Employment Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** ---------- MENU / SIDEBAR ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('COE Generator')
    .addItem('Open Generator', 'showApp')
    .addToUi();
}

function showApp() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(950)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Certificate of Employment Generator');
}

/** ---------- SHEET HELPERS ---------- */

function getSheetAsObjects_(spreadsheetId, sheetName) {
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

/** ---------- EMPLOYEE LOOKUP ---------- */

function getEmployees() {
  const rows = getSheetAsObjects_(CONFIG.WORKFORCE_SHEET_ID, CONFIG.WORKFORCE_SHEET_NAME);
  return rows.map(r => ({
    id: r.EmployeeID,
    fullName: r.FullName,
    position: r.Position,
    department: r.Department
  }));
}

function getEmployeeById_(employeeId) {
  const rows = getSheetAsObjects_(CONFIG.WORKFORCE_SHEET_ID, CONFIG.WORKFORCE_SHEET_NAME);
  return rows.find(r => String(r.EmployeeID) === String(employeeId)) || null;
}

/** ---------- TEXT / DATE HELPERS ---------- */

function fillTokens_(template, emp) {
  const salaryFormatted = emp.MonthlySalary
    ? '₱' + Number(emp.MonthlySalary).toLocaleString('en-PH', { minimumFractionDigits: 2 })
    : 'N/A';
  const dateHiredFormatted = emp.DateHired
    ? Utilities.formatDate(new Date(emp.DateHired), 'GMT+8', 'MMMM d, yyyy')
    : 'N/A';

  const tokenMap = {
    FULL_NAME: emp.FullName || '',
    POSITION: emp.Position || '',
    DEPARTMENT: emp.Department || '',
    STATUS: emp.EmploymentStatus || '',
    DATE_HIRED: dateHiredFormatted,
    SALARY: salaryFormatted
  };

  let output = template;
  Object.keys(tokenMap).forEach(key => {
    output = output.split(`<<${key}>>`).join(tokenMap[key]);
  });
  return output;
}

function formatOrdinalDate_(date) {
  const day = date.getDate();
  const suffix = (d) => {
    if (d > 3 && d < 21) return 'th';
    switch (d % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  };
  const monthYear = Utilities.formatDate(date, 'GMT+8', 'MMMM yyyy');
  return `${day}${suffix(day)} day of ${monthYear}`;
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** ---------- COE NUMBER (race-condition safe) ---------- */

function getNextCoeNo_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const year = new Date().getFullYear();
    const key = 'COE_COUNTER_' + year;
    const count = parseInt(props.getProperty(key) || '0', 10) + 1;
    props.setProperty(key, String(count));
    const padded = String(count).padStart(4, '0');
    return `${CONFIG.COE_PREFIX}-${year}-${padded}`;
  } finally {
    lock.releaseLock();
  }
}

/** ---------- LOGGING ---------- */

function logIssuance_(entry) {
  const sheet = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  const rowMap = {
    'Date Generated': entry.dateGenerated,
    'Requester Name': entry.requesterName,
    'COE Variant': entry.variant,
    'OR Number': entry.coeNo,     // TODO: rename this header to "COE No." in the sheet for clarity
    'Purpose': entry.purpose,
    'Signatory': entry.signatory,
    'Processor': entry.processor
  };

  const row = headers.map(h => (rowMap[h] !== undefined ? rowMap[h] : ''));
  sheet.appendRow(row);
}

/** ---------- MAIN ENTRY POINT (called from index.html) ---------- */

function generateCOE(formData) {
  try {
    const emp = getEmployeeById_(formData.employeeId);
    if (!emp) throw new Error('Employee not found. Please re-select from the list.');

    const coeNo = getNextCoeNo_();
    const doneDate = formatOrdinalDate_(new Date());

    const clauseTemplate = formData.variant === 'withComp'
      ? CONFIG.CLAUSES.withCompensation
      : CONFIG.CLAUSES.withoutCompensation;
    const bodyClause = fillTokens_(clauseTemplate, emp);

    const shortName = formData.requesterName || `${emp.FirstName} ${emp.LastName}`;
    const purposeClause = formData.purpose
      ? ` for ${formData.purpose}.`
      : ' for whatever legal purpose it may serve.';

    const placeholders = {
      NAME: emp.FullName,
      BODY_CLAUSE: bodyClause,
      SHORT_NAME: shortName,
      PURPOSE_CLAUSE: purposeClause,
      DONE_DATE: doneDate,
      SIG_NAME: CONFIG.SIGNATORY.name,
      SIG_TITLE1: CONFIG.SIGNATORY.title1,
      SIG_TITLE2: CONFIG.SIGNATORY.title2,
      OR_NO: coeNo,
      ISSUED_ON: formData.issuedOn || doneDate,
      ISSUED_AT: formData.issuedAt || 'Butuan City',
      AMOUNT_PAID: formData.amountPaid || 'N/A'
    };

    const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    const docName = `COE_${emp.FullName}_${coeNo}`;

    const copyFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID).makeCopy(docName, folder);
    const doc = DocumentApp.openById(copyFile.getId());
    const body = doc.getBody();

    Object.keys(placeholders).forEach(key => {
      body.replaceText(escapeRegex_(`{{${key}}}`), String(placeholders[key]));
    });
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copyFile.getId()).getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(docName + '.pdf');

    if (!CONFIG.KEEP_DOC_COPY) {
      DriveApp.getFileById(copyFile.getId()).setTrashed(true);
    }

    logIssuance_({
      dateGenerated: new Date(),
      requesterName: shortName,
      variant: formData.variant === 'withComp' ? 'With Compensation' : 'Without Compensation',
      coeNo: coeNo,
      purpose: formData.purpose || 'General purpose',
      signatory: CONFIG.SIGNATORY.name,
      processor: formData.processorName || Session.getActiveUser().getEmail() || 'Unspecified'
    });

    return {
      success: true,
      coeNo: coeNo,
      docUrl: CONFIG.KEEP_DOC_COPY ? copyFile.getUrl() : null,
      pdfUrl: pdfFile.getUrl()
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
/** ---------- DASHBOARD DATA ---------- */

function getDashboardStats() {
  const rows = getSheetAsObjects_(CONFIG.LOG_SHEET_ID, CONFIG.LOG_SHEET_NAME);
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  let thisMonth = 0, withComp = 0, withoutComp = 0;

  rows.forEach(r => {
    const d = new Date(r['Date Generated']);
    if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) thisMonth++;
    if (String(r['COE Variant']).toLowerCase().includes('without')) withoutComp++;
    else withComp++;
  });

  return {
    total: rows.length,
    thisMonth: thisMonth,
    withComp: withComp,
    withoutComp: withoutComp
  };
}

function getRecentIssuances() {
  const rows = getSheetAsObjects_(CONFIG.LOG_SHEET_ID, CONFIG.LOG_SHEET_NAME);
  return rows
    .sort((a, b) => new Date(b['Date Generated']) - new Date(a['Date Generated']))
    .slice(0, 8)
    .map(r => ({
      requesterName: r['Requester Name'],
      coeNo: r['OR Number'],
      variant: r['COE Variant'],
      date: Utilities.formatDate(new Date(r['Date Generated']), 'GMT+8', 'MMM d, yyyy')
    }));
}
