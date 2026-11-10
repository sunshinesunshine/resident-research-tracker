/**
 * =========================================================================
 * ระบบติดตามความก้าวหน้างานวิจัยแพทย์ประจำบ้าน (Resident Research Tracker)
 * ภาควิชาจักษุวิทยา (Department of Ophthalmology)
 * 
 * Google Apps Script Backend (Code.gs)
 * ทำหน้าที่เป็น Serverless Controller เชื่อมต่อ Google Sheets 100% ฟรี
 * =========================================================================
 */

const SHEET_NAMES = {
  RESIDENTS: 'Residents',
  CURRICULUM: 'Curriculum_Config',
  FORUMS: 'Forums',
  ADVISORS: 'Advisors',
  CONFIG: 'App_Config'
};


const SPREADSHEET_ID = '18f71XkkgACgGwhWGii2wIA0BnAlsnox0_sy61VEkFXk';

function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID && SPREADSHEET_ID.length > 10) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  } catch (err) {
    Logger.log('openById error, falling back to getActive: ' + err);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
const DEFAULT_PIN = '1234';

/**
 * เปิดหน้าเว็บ Web App
 */
function doGet(e) {
  try {
    checkAndInitializeDatabase();
  } catch(err) {
    Logger.log('Init database on doGet error: ' + err);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบติดตามงานวิจัย Resident | ภาควิชาจักษุวิทยา')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * ดึงข้อมูลทั้งหมดในครั้งเดียว (Single Batch GET - 0 ms UI Latency)
 */
function getInitialData() {
  try {
    checkAndInitializeDatabase();
    const ss = getSpreadsheet();
    
    // 1. App Config
    const config = getAppConfigData(ss);
    
    // 2. Curriculum Template (6 ขั้นตอน)
    const curriculumTemplate = getCurriculumData(ss);
    
    // 3. Forums
    const forums = getForumsData(ss);
    
    // 4. Advisors (Staff List)
    const staffList = getAdvisorsData(ss);
    
    // 5. Residents
    const residents = getResidentsData(ss);

    return {
      success: true,
      data: {
        academicYearConfig: {
          autoAcademicYear: config.autoAcademicYear !== false,
          currentYear: parseInt(config.currentYear, 10) || 2569
        },
        dashboardSectionOrder: config.dashboardSectionOrder || ['kpi', 'radar', 'progression', 'forum', 'design'],
        adminPin: config.adminPin || DEFAULT_PIN,
        curriculumTemplate: curriculumTemplate,
        forums: forums,
        staff: staffList,
        staffList: staffList,
        residents: residents
      }
    };
  } catch (err) {
    Logger.log('Error in getInitialData: ' + err.toString());
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * บันทึก/อัปเดตข้อมูล Resident รายบุคคล
 */
function saveResidentData(resident) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
    if (!sheet) throw new Error('Sheet Residents not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIdx = headers.indexOf('id');
    if (idColIdx === -1) throw new Error('Column "id" not found in Residents sheet');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]) === String(resident.id)) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    const rowValues = serializeResidentRow(headers, resident);

    if (rowIndex > 0) {
      // อัปเดตแถวเดิม
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // เพิ่มแถวใหม่
      sheet.appendRow(rowValues);
    }

    return { success: true, residentId: resident.id };
  } catch (err) {
    Logger.log('Error saveResidentData: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * ลบ Resident ออกจากระบบ
 */
function deleteResident(residentId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
    if (!sheet) throw new Error('Sheet Residents not found');

    const data = sheet.getDataRange().getValues();
    const idColIdx = data[0].indexOf('id');
    if (idColIdx === -1) throw new Error('Column "id" not found');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]) === String(residentId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }

    return { success: false, error: 'Resident not found' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * บันทึกกรอบเวลาหลักสูตร (Curriculum Template 6 ขั้นตอน)
 */
function saveCurriculumTemplate(templateList) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.CURRICULUM);
    if (!sheet) throw new Error('Curriculum sheet not found');

    sheet.clearContents();
    const headers = ['stage', 'name', 'shortName', 'detail', 'startMonth', 'startCurriculumYear', 'endMonth', 'endCurriculumYear'];
    const rows = [headers];

    templateList.forEach(t => {
      rows.push([
        t.stage,
        t.name,
        t.shortName,
        t.detail,
        t.startMonth,
        t.startCurriculumYear,
        t.endMonth,
        t.endCurriculumYear
      ]);
    });

    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    formatHeaderRow(sheet);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * บันทึกกำหนดการรอบ Forum
 */
function saveForums(forumsList) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.FORUMS);
    if (!sheet) throw new Error('Forums sheet not found');

    sheet.clearContents();
    const headers = ['id', 'dateStr', 'displayDate', 'title', 'focus'];
    const rows = [headers];

    forumsList.forEach(f => {
      rows.push([f.id, f.dateStr, f.displayDate, f.title, f.focus]);
    });

    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    formatHeaderRow(sheet);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * บันทึกรายชื่ออาจารย์ที่ปรึกษา (Advisors)
 */
function saveAdvisors(staffList) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ADVISORS);
    if (!sheet) throw new Error('Advisors sheet not found');

    sheet.clearContents();
    const headers = ['id', 'title', 'name', 'role', 'subspecialty'];
    const rows = [headers];

    (staffList || []).forEach(s => {
      rows.push([
        s.id || '',
        s.title || '',
        s.name || '',
        s.role || s.field || 'อาจารย์ที่ปรึกษา',
        s.subspecialty || s.field || ''
      ]);
    });

    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    formatHeaderRow(sheet);
    return { success: true };
  } catch (err) {
    Logger.log('saveAdvisors error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * บันทึกการตั้งค่าระบบ (App Config)
 */
function saveAppConfig(config) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    if (!sheet) throw new Error('Config sheet not found');

    const configMap = {
      autoAcademicYear: config.autoAcademicYear,
      currentYear: config.currentYear,
      dashboardSectionOrder: JSON.stringify(config.dashboardSectionOrder || ['kpi', 'radar', 'progression', 'forum', 'design']),
      adminPin: config.adminPin || DEFAULT_PIN,
      lastUpdated: new Date().toISOString()
    };

    sheet.clearContents();
    const rows = [['Key', 'Value']];
    for (const k in configMap) {
      rows.push([k, String(configMap[k])]);
    }

    sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    formatHeaderRow(sheet);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * รีเซ็ตระบบกลับเป็นค่าเริ่มต้น (Seed Clean Data)
 */
function resetSystemToDefault(pin) {
  try {
    const currentConfig = getAppConfigData(getSpreadsheet());
    const validPin = currentConfig.adminPin || DEFAULT_PIN;
    if (String(pin).trim() !== String(validPin).trim()) {
      return { success: false, error: 'รหัสผ่าน Admin PIN ไม่ถูกต้อง' };
    }

    initializeAllSheets(true);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// =========================================================================
// HELPER SERIALIZERS & SHEET READERS
// =========================================================================

function getAppConfigData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = data[i][1];
    if (key === 'autoAcademicYear') config[key] = (val === 'true' || val === true);
    else if (key === 'currentYear') config[key] = parseInt(val, 10) || 2569;
    else if (key === 'dashboardSectionOrder') {
      try { config[key] = JSON.parse(val); } catch(e) { config[key] = ['kpi', 'radar', 'progression', 'forum', 'design']; }
    } else {
      config[key] = val;
    }
  }
  return config;
}

function getCurriculumData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CURRICULUM);
  if (!sheet) return getDefaultCurriculumTemplate();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getDefaultCurriculumTemplate();

  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    list.push({
      stage: parseInt(row[headers.indexOf('stage')], 10) || (i),
      name: String(row[headers.indexOf('name')] || ''),
      shortName: String(row[headers.indexOf('shortName')] || ''),
      detail: String(row[headers.indexOf('detail')] || ''),
      startMonth: parseInt(row[headers.indexOf('startMonth')], 10),
      startCurriculumYear: parseInt(row[headers.indexOf('startCurriculumYear')], 10),
      endMonth: parseInt(row[headers.indexOf('endMonth')], 10),
      endCurriculumYear: parseInt(row[headers.indexOf('endCurriculumYear')], 10)
    });
  }
  return list;
}

function getForumsData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.FORUMS);
  if (!sheet) return getDefaultForums();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getDefaultForums();

  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    list.push({
      id: row[headers.indexOf('id')],
      dateStr: String(row[headers.indexOf('dateStr')] || ''),
      displayDate: String(row[headers.indexOf('displayDate')] || ''),
      title: String(row[headers.indexOf('title')] || ''),
      focus: String(row[headers.indexOf('focus')] || '')
    });
  }
  return list;
}

function getAdvisorsData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISORS);
  if (!sheet) return getDefaultStaffList();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[headers.indexOf('name')] || '').trim();
    if (name) {
      const field = String(row[headers.indexOf('subspecialty')] || row[headers.indexOf('role')] || '');
      list.push({
        id: String(row[headers.indexOf('id')] || ('staff_' + i)),
        title: String(row[headers.indexOf('title')] || ''),
        name: name,
        role: String(row[headers.indexOf('role')] || 'อาจารย์ที่ปรึกษา'),
        subspecialty: String(row[headers.indexOf('subspecialty')] || ''),
        field: field
      });
    }
  }
  return list;
}

function getResidentsData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[headers.indexOf('id')];
    const name = row[headers.indexOf('name')];
    if (!id || !name) continue;

    const milestones = [
      row[headers.indexOf('stage1_done')] === true,
      row[headers.indexOf('stage2_done')] === true,
      row[headers.indexOf('stage3_done')] === true,
      row[headers.indexOf('stage4_done')] === true,
      row[headers.indexOf('stage5_done')] === true,
      row[headers.indexOf('stage6_done')] === true
    ];

    list.push({
      id: String(id),
      name: String(name),
      nickname: String(row[headers.indexOf('nickname')] || ''),
      batchYear: parseInt(row[headers.indexOf('batchYear')], 10) || 2569,
      subspecialty: String(row[headers.indexOf('subspecialty')] || ''),
      title: String(row[headers.indexOf('title')] || ''),
      studyDesign: String(row[headers.indexOf('studyDesign')] || ''),
      advisor: String(row[headers.indexOf('advisor')] || ''),
      coAdvisor: String(row[headers.indexOf('coAdvisor')] || ''),
      targetN: String(row[headers.indexOf('targetN')] || ''),
      irbStatus: String(row[headers.indexOf('irbStatus')] || ''),
      funding: String(row[headers.indexOf('funding')] || ''),
      driveLink: String(row[headers.indexOf('driveLink')] || ''),
      notes: String(row[headers.indexOf('notes')] || ''),
      milestones: milestones,
      isAlumni: row[headers.indexOf('isAlumni')] === true,
      lastUpdated: String(row[headers.indexOf('lastUpdated')] || '')
    });
  }
  return list;
}

function serializeResidentRow(headers, r) {
  const ms = r.milestones || [false, false, false, false, false, false];
  const map = {
    id: r.id,
    name: r.name,
    nickname: r.nickname || '',
    batchYear: r.batchYear || 2569,
    subspecialty: r.subspecialty || '',
    title: r.title || '',
    studyDesign: r.studyDesign || '',
    advisor: r.advisor || '',
    coAdvisor: r.coAdvisor || '',
    targetN: r.targetN || '',
    irbStatus: r.irbStatus || '',
    funding: r.funding || '',
    driveLink: r.driveLink || '',
    notes: r.notes || '',
    stage1_done: ms[0] === true,
    stage2_done: ms[1] === true,
    stage3_done: ms[2] === true,
    stage4_done: ms[3] === true,
    stage5_done: ms[4] === true,
    stage6_done: ms[5] === true,
    isAlumni: r.isAlumni === true,
    lastUpdated: new Date().toISOString()
  };

  return headers.map(h => (map.hasOwnProperty(h) ? map[h] : ''));
}

// =========================================================================
// DATABASE AUTO-INITIALIZER (FIRST RUN SETUP)
// =========================================================================

function checkAndInitializeDatabase() {
  const ss = getSpreadsheet();
  let initialized = false;
  for (const key in SHEET_NAMES) {
    if (!ss.getSheetByName(SHEET_NAMES[key])) {
      initialized = true;
      break;
    }
  }
  if (initialized) {
    initializeAllSheets(false);
  }

  // Auto-sync official 33 staff members to Advisors sheet if old list without subspecialties is detected
  try {
    const advSheet = ss.getSheetByName(SHEET_NAMES.ADVISORS);
    if (advSheet) {
      const data = advSheet.getDataRange().getValues();
      const hasCornea = data.some(row => String(row[4] || '') === 'Cornea');
      if (!hasCornea || data.length <= 5) {
        saveAdvisors(getDefaultStaffList());
      }
    }
  } catch (e) {
    Logger.log('Auto-update advisors error: ' + e);
  }
}

function initializeAllSheets(forceReset) {
  const ss = getSpreadsheet();

  // 1. Config Sheet
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!configSheet || forceReset) {
    if (!configSheet) configSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    saveAppConfig({
      autoAcademicYear: true,
      currentYear: 2569,
      dashboardSectionOrder: ['kpi', 'radar', 'progression', 'forum', 'design'],
      adminPin: DEFAULT_PIN
    });
  }

  // 2. Curriculum Config Sheet
  let currSheet = ss.getSheetByName(SHEET_NAMES.CURRICULUM);
  if (!currSheet || forceReset) {
    if (!currSheet) currSheet = ss.insertSheet(SHEET_NAMES.CURRICULUM);
    saveCurriculumTemplate(getDefaultCurriculumTemplate());
  }

  // 3. Forums Sheet
  let forumSheet = ss.getSheetByName(SHEET_NAMES.FORUMS);
  if (!forumSheet || forceReset) {
    if (!forumSheet) forumSheet = ss.insertSheet(SHEET_NAMES.FORUMS);
    saveForums(getDefaultForums());
  }

  // 4. Advisors Sheet
  let advSheet = ss.getSheetByName(SHEET_NAMES.ADVISORS);
  if (!advSheet || forceReset) {
    if (!advSheet) advSheet = ss.insertSheet(SHEET_NAMES.ADVISORS);
    saveAdvisors(getDefaultStaffList());
  }

  // 5. Residents Sheet
  let resSheet = ss.getSheetByName(SHEET_NAMES.RESIDENTS);
  if (!resSheet || forceReset) {
    if (!resSheet) resSheet = ss.insertSheet(SHEET_NAMES.RESIDENTS);
    resSheet.clearContents();
    const headers = [
      'id', 'name', 'nickname', 'batchYear', 'subspecialty', 'title', 'studyDesign',
      'advisor', 'coAdvisor', 'targetN', 'irbStatus', 'funding', 'driveLink', 'notes',
      'stage1_done', 'stage2_done', 'stage3_done', 'stage4_done', 'stage5_done', 'stage6_done',
      'isAlumni', 'lastUpdated'
    ];
    resSheet.appendRow(headers);
    formatHeaderRow(resSheet);

    // Populate Sample Test Data (ชั้นปีละ 2 คน)
    const samples = getDefaultSampleResidents();
    samples.forEach(r => {
      resSheet.appendRow(serializeResidentRow(headers, r));
    });
  }
}

function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#102A43'); // Navy 900
  headerRange.setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

// =========================================================================
// DEFAULT DATA PROVIDERS
// =========================================================================

function getDefaultCurriculumTemplate() {
  return [
    {
      stage: 1,
      name: "1. ทบทวนวรรณกรรม",
      shortName: "ทบทวนวรรณกรรม",
      detail: "ทบทวนวรรณกรรม & กำหนดกรอบแนวคิด PICO",
      startMonth: 6, // ก.ค.
      startCurriculumYear: 1,
      endMonth: 9,   // ต.ค.
      endCurriculumYear: 1
    },
    {
      stage: 2,
      name: "2. ยื่นโครงร่างพิจารณา IRB",
      shortName: "ยื่น IRB",
      detail: "ยื่นโครงร่างวิจัยพิจารณาจริยธรรมการวิจัย (IRB)",
      startMonth: 10, // พ.ย.
      startCurriculumYear: 1,
      endMonth: 1,    // ก.พ.
      endCurriculumYear: 1
    },
    {
      stage: 3,
      name: "3. ขอทุนสนับสนุนวิจัย & แก้ไข",
      shortName: "ขอทุน / แก้ไข",
      detail: "ขอทุนสนับสนุนงานวิจัย และยื่นแก้ไขพิจารณาตามข้อเสนอแนะกรรมการ",
      startMonth: 2, // มี.ค.
      startCurriculumYear: 1,
      endMonth: 5,   // มิ.ย.
      endCurriculumYear: 1
    },
    {
      stage: 4,
      name: "4. เก็บรวบรวมข้อมูล",
      shortName: "เก็บข้อมูล",
      detail: "เก็บรวบรวมข้อมูลงานวิจัย / ติดตามเคสผู้ป่วย / ลงบันทึก CRF",
      startMonth: 6, // ก.ค.
      startCurriculumYear: 2,
      endMonth: 9,   // ต.ค.
      endCurriculumYear: 2
    },
    {
      stage: 5,
      name: "5. วิเคราะห์ทางสถิติ",
      shortName: "วิเคราะห์สถิติ",
      detail: "ทำความสะอาดข้อมูล (Data Cleaning) & วิเคราะห์ผลการวิจัยทางสถิติ",
      startMonth: 10, // พ.ย.
      startCurriculumYear: 2,
      endMonth: 0,    // ม.ค.
      endCurriculumYear: 2
    },
    {
      stage: 6,
      name: "6. จัดทำรูปเล่ม & เผยแพร่",
      shortName: "รูปเล่ม & เผยแพร่",
      detail: "จัดทำรูปเล่มฉบับสมบูรณ์ นำเสนอใน Research Forum และส่งตีพิมพ์ผลงาน",
      startMonth: 1, // ก.พ.
      startCurriculumYear: 2,
      endMonth: 5,   // มิ.ย.
      endCurriculumYear: 3
    }
  ];
}

function getDefaultForums() {
  return [
    {
      id: 1,
      dateStr: "2027-01-19",
      displayDate: "19 มกราคม 2570",
      title: "Research Forum ครั้งที่ 1 / 2570",
      focus: "R1: นำเสนอโครงร่าง PICO | R2: รายงานความก้าวหน้าการเก็บข้อมูล | R3: ซ้อมนำเสนอผลงานฉบับสมบูรณ์"
    },
    {
      id: 2,
      dateStr: "2027-04-20",
      displayDate: "20 เมษายน 2570",
      title: "Research Forum ครั้งที่ 2 / 2570",
      focus: "R1: สรุปแบบยื่นจริยธรรม IRB | R2: เริ่มวิเคราะห์ผลสถิติเบื้องต้น | R3: ตรวจร่าง Manuscript รูปเล่ม"
    },
    {
      id: 3,
      dateStr: "2027-06-15",
      displayDate: "15 มิถุนายน 2570",
      title: "Annual Ophthalmology Resident Research Day",
      focus: "R3: Final Presentation & ประกวดงานวิจัยยอดเยี่ยม | R2: ส่งฐานข้อมูลวิจัย | R1: รับรองผลการยื่น IRB"
    }
  ];
}

function getDefaultStaffList() {
  return [
    { id: "staff_1", title: "", name: "นวลจิรา ประกายรุ้งทอง", role: "อาจารย์ที่ปรึกษา", subspecialty: "Cornea", field: "Cornea" },
    { id: "staff_2", title: "", name: "กนกทิพย์ มันตโชติ", role: "อาจารย์ที่ปรึกษา", subspecialty: "Cornea", field: "Cornea" },
    { id: "staff_3", title: "", name: "เบญจมา กีรติดำเกิงสกุล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Cornea", field: "Cornea" },
    { id: "staff_4", title: "", name: "วีรภัทร อุดมวงศ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Cornea", field: "Cornea" },
    { id: "staff_5", title: "", name: "ชิดชนก ตันติพัฒน์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Cornea", field: "Cornea" },
    { id: "staff_6", title: "", name: "ดวงดาว ทัศณรงค์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Glaucoma", field: "Glaucoma" },
    { id: "staff_7", title: "", name: "กุลวรรณ โรจนเนืองนิตย์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Glaucoma", field: "Glaucoma" },
    { id: "staff_8", title: "", name: "อรอร ธงอินเนตร", role: "อาจารย์ที่ปรึกษา", subspecialty: "Glaucoma", field: "Glaucoma" },
    { id: "staff_9", title: "", name: "สุพัตรา จามรสุวรรณ", role: "อาจารย์ที่ปรึกษา", subspecialty: "Glaucoma", field: "Glaucoma" },
    { id: "staff_10", title: "", name: "ปิติพงศ์ สุรเมธากุล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Glaucoma", field: "Glaucoma" },
    { id: "staff_11", title: "", name: "ปิยวดี ชัยมงคลตระกูล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Neuro", field: "Neuro" },
    { id: "staff_12", title: "", name: "ปรเมศวร์ เลาวณาภิบาล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Neuro", field: "Neuro" },
    { id: "staff_13", title: "", name: "จิเมธัส เฉลิมพงษ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Neuro", field: "Neuro" },
    { id: "staff_14", title: "", name: "สุณิสา สินธุวงศ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "OPRC", field: "OPRC" },
    { id: "staff_15", title: "", name: "มิ่งขวัญ ลำยองเสถียร", role: "อาจารย์ที่ปรึกษา", subspecialty: "OPRC", field: "OPRC" },
    { id: "staff_16", title: "", name: "อรวีณัฏฐ์ นิมิตรวงศ์สกุล", role: "อาจารย์ที่ปรึกษา", subspecialty: "OPRC", field: "OPRC" },
    { id: "staff_17", title: "", name: "สรรเสริญ วัฒนะพานิช", role: "อาจารย์ที่ปรึกษา", subspecialty: "OPRC", field: "OPRC" },
    { id: "staff_18", title: "", name: "อาภา พรเศรษฐ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Ped", field: "Ped" },
    { id: "staff_19", title: "", name: "จิราภา ทรงเพ็ชร์มงคล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Ped", field: "Ped" },
    { id: "staff_20", title: "", name: "สุดาวดี สมบูรณ์ธนกิจ", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_21", title: "", name: "วงศ์ศิริ ทวีบรรจงสิน", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_22", title: "", name: "วรภัทร วงษ์สวัสดิ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_23", title: "", name: "อุมาภรณ์ จิตตวัฒนรัตน์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_24", title: "", name: "พิริญา สุ่มสวัสดิ์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_25", title: "", name: "เอกชัย อารยางกูร", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_26", title: "", name: "กาญจน์ธิดา จันทน์วิมล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_27", title: "", name: "วรินทร์ สมิทธิเมธินทร์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_28", title: "", name: "นวลพร ศรีม่วง", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_29", title: "", name: "ทิฆัมพร จันทร์ภู่", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_30", title: "", name: "ปาณิศา สิงหเนตร", role: "อาจารย์ที่ปรึกษา", subspecialty: "Retina", field: "Retina" },
    { id: "staff_31", title: "", name: "โยธิน ฐิตวัฒนกุล", role: "อาจารย์ที่ปรึกษา", subspecialty: "Uveitis", field: "Uveitis" },
    { id: "staff_32", title: "", name: "นพคุณ ปัญญายิ่งยง", role: "อาจารย์ที่ปรึกษา", subspecialty: "Uveitis", field: "Uveitis" },
    { id: "staff_33", title: "", name: "ณิศวรรณ วัฒนศัพท์", role: "อาจารย์ที่ปรึกษา", subspecialty: "Uveitis", field: "Uveitis" }
  ];
}

function getDefaultSampleResidents() {
  return [
    // R1 (Batch 2569)
    {
      id: "res_2569_01",
      name: "พญ.วิชญา รุ่งเรือง",
      nickname: "หมอนก",
      batchYear: 2569,
      subspecialty: "Glaucoma",
      title: "ความแม่นยำของการตรวจ Corneal Hysteresis ในการทำนายภาวะ Steroid Response ในผู้ป่วยหลังผ่าตัดต้อกระจก",
      studyDesign: "Prospective Cohort Study",
      advisor: "ศ.คลินิก นพ. ชัยพร ศิริวัฒน์",
      coAdvisor: "รศ. พญ. วราภรณ์ สุวรรณเวช",
      targetN: "120 ตา",
      irbStatus: "ยื่นเสนอขอจริยธรรมแล้ว (อยู่ระหว่างพิจารณา)",
      funding: "ทุนวิจัยคณะแพทยศาสตร์",
      driveLink: "https://drive.google.com",
      notes: "ทบทวนวรรณกรรมเสร็จสมบูรณ์ กำลังปรับแก้โครงร่างวิจัยตามข้อเสนอแนะเบื้องต้นของที่ปรึกษา",
      milestones: [true, false, false, false, false, false],
      isAlumni: false
    },
    {
      id: "res_2569_02",
      name: "นพ.พีรวิชญ์ เมธาธรณ์",
      nickname: "หมอพีท",
      batchYear: 2569,
      subspecialty: "Retina",
      title: "การศึกษาผลของการฉีด Anti-VEGF ต่อนความหนาของจอประสาทตาและหลอดเลือดในผู้ป่วย Diabetic Macular Edema",
      studyDesign: "Prospective Cohort Study",
      advisor: "รศ. พญ. วราภรณ์ สุวรรณเวช",
      coAdvisor: "",
      targetN: "80 คน",
      irbStatus: "ร่างโครงร่างเสร็จสิ้น เตรียมยื่นส่งเอกสาร",
      funding: "ทุนภาควิชาจักษุวิทยา",
      driveLink: "https://drive.google.com",
      notes: "กำลังตรวจสอบเครื่องมือตรวจวัด OCTA ร่วมกับฝ่ายวิจัย",
      milestones: [true, false, false, false, false, false],
      isAlumni: false
    },
    // R2 (Batch 2568)
    {
      id: "res_2568_01",
      name: "นพ.ธนวัฒน์ เกียรติสกุล",
      nickname: "หมอบาส",
      batchYear: 2568,
      subspecialty: "Cornea & Refractive",
      title: "การเปรียบเทียบผลลัพธ์สายตาและค่าความโค้งกระจกตาระหว่าง Toric IOL กับ Non-Toric ร่วมกับ Astigmatic Keratotomy",
      studyDesign: "Randomized Controlled Trial (RCT)",
      advisor: "ผศ. นพ. กิตติพงษ์ ภัทรศรัณย์",
      coAdvisor: "",
      targetN: "90 ตา",
      irbStatus: "ผ่านการรับรองแล้ว (EC-68-042)",
      funding: "ทุนอุดหนุนการวิจัยประจำปี",
      driveLink: "https://drive.google.com",
      notes: "เก็บรวบรวมข้อมูลเคสครบตามเป้าหมายแล้ว อยู่ระหว่าง Data Cleaning และวิเคราะห์สถิติด้วย R",
      milestones: [true, true, true, true, false, false],
      isAlumni: false
    },
    {
      id: "res_2568_02",
      name: "พญ.ศิรินภา ธีระวานิช",
      nickname: "หมอส้ม",
      batchYear: 2568,
      subspecialty: "Glaucoma",
      title: "ประสิทธิผลของ Micropulse Transscleral Cyclophotocoagulation ในผู้ป่วยต้อหินขั้นรุนแรง: การศึกษาติดตามผล 1 ปี",
      studyDesign: "Retrospective Chart Review",
      advisor: "ศ.คลินิก นพ. ชัยพร ศิริวัฒน์",
      coAdvisor: "",
      targetN: "100 ตา",
      irbStatus: "ผ่านการรับรองแล้ว (EC-68-015)",
      funding: "ทุนภาควิชา",
      driveLink: "https://drive.google.com",
      notes: "บันทึกข้อมูล eCRF ไปแล้ว 65% เกิดความล่าช้าในการดึงแฟ้มเวชระเบียนเก่า กำลังประสานงานห้องเวชระเบียน",
      milestones: [true, true, true, false, false, false],
      isAlumni: false
    },
    // R3 (Batch 2567)
    {
      id: "res_2567_01",
      name: "พญ.กานต์พิชชา วัฒนดำรง",
      nickname: "หมอเกรซ",
      batchYear: 2567,
      subspecialty: "Pediatric & Strabismus",
      title: "ความสัมพันธ์ของมุมเหล่และความสามารถในการมองเห็นภาพสามมิติหลังผ่าตัดในเด็กที่มีภาวะ Exotropia",
      studyDesign: "Prospective Cohort Study",
      advisor: "อ. พญ. นพวรรณ รัตนพิเชษฐ์",
      coAdvisor: "",
      targetN: "60 คน",
      irbStatus: "ผ่านการรับรองแล้ว (EC-67-089)",
      funding: "ทุนมูลนิธิโรคตา",
      driveLink: "https://drive.google.com",
      notes: "วิเคราะห์ผลสถิติเสร็จสิ้นแล้ว กำลังเรียบเรียง Manuscript ภาษาอังกฤษเตรียมส่งตีพิมพ์ใน TJO",
      milestones: [true, true, true, true, true, false],
      isAlumni: false
    },
    {
      id: "res_2567_02",
      name: "นพ.ชยพล วงศ์สวัสดิ์",
      nickname: "หมอท็อป",
      batchYear: 2567,
      subspecialty: "Retina",
      title: "การประเมินความแม่นยำของแบบจำลอง Deep Learning ในการตรวจคัดกรองเบาหวานขึ้นจอประสาทตาจากภาพถ่าย Fundus",
      studyDesign: "Diagnostic Accuracy Study",
      advisor: "รศ. พญ. วราภรณ์ สุวรรณเวช",
      coAdvisor: "",
      targetN: "1,500 ภาพ",
      irbStatus: "ผ่านการรับรองแล้ว (EC-67-033)",
      funding: "ทุนวิจัยพัฒนานวัตกรรม",
      driveLink: "https://drive.google.com",
      notes: "จัดทำรูปเล่มฉบับสมบูรณ์เรียบร้อยแล้ว เตรียมนำเสนอในงาน Annual Research Forum",
      milestones: [true, true, true, true, true, true],
      isAlumni: false
    }
  ];
}
