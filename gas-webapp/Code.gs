const PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
const PROP_REMOTE_AUTH_SALT = 'REMOTE_AUTH_SALT';
const PROP_REMOTE_AUTH_HASH = 'REMOTE_AUTH_HASH';
const SETTINGS_SUFFIX = '_設定';
const STUDENTS_SUFFIX = '_學生名單';
const SCHEDULE_SUFFIX = '_自動排程';
const BACKUP_SHEET_NAME = '_系統備份';
const BACKUP_TRIGGER_HANDLER = 'createTimedBackupSnapshot';
const DEFAULT_SPREADSHEET_NAME = '二階甄選智慧排程雲端資料';
const SHEET_NAME_MAX_LENGTH = 90;
const APP_SCHEMA_VERSION = '2026-06-13';
const SECURE_STORAGE_MODE = 'encrypted_v1';

const SETTINGS_KEYS = [
  ['schema_version', '資料格式版本'],
  ['storage_mode', '儲存模式'],
  ['department_id', '科系識別碼'],
  ['department_name', '科系名稱'],
  ['department_sub', '副標題'],
  ['department_icon', '圖示'],
  ['cfg_written_duration', '筆試時間（分）'],
  ['cfg_oral_duration', '口試時間（分）'],
  ['cfg_oral_teams_per_ladder', '一梯幾隊'],
  ['cfg_oral_team_group_count', '一隊幾組'],
  ['cfg_team_capacity', '單間筆試教室上限'],
  ['cfg_oral_group_capacity', '口試每組人數'],
  ['cfg_scheduler_strategy', '排法代碼'],
  ['cfg_school_grouping', '同校分組偏好'],
  ['written_rooms_json', '筆試教室 JSON'],
  ['oral_rooms_json', '口試教室 JSON'],
  ['sessions_json', '甄試時段 JSON'],
  ['students_count', '學生人數'],
  ['scheduled_count', '已排程人數'],
  ['updated_at_iso', '最後更新 ISO 時間']
];

const STUDENT_HEADERS = [
  'student_id',
  'student_name',
  'birth_date',
  'school',
  'preferred_session_id',
  'preferred_session_name',
  'conflict_time',
  'assigned_session_id',
  'assigned_session_name',
  'assigned_written_room_idx',
  'assigned_written_room_name',
  'assigned_oral_room_idx',
  'assigned_oral_room_name',
  'assigned_order',
  'assigned_oral_group_idx',
  'written_start',
  'written_end',
  'oral_start',
  'oral_end',
  'updated_at_iso'
];

const SCHEDULE_HEADERS = [
  'student_id',
  'student_name',
  'birth_date',
  'school',
  'preferred_session',
  'assigned_session',
  'written_room',
  'written_time',
  'oral_room',
  'oral_group',
  'oral_time',
  'conflict_time',
  'updated_at_iso'
];

const BACKUP_HEADERS = [
  'saved_at_iso',
  'source',
  'schema_version',
  'department_count',
  'data_json'
];

const SECURE_STUDENT_HEADERS = [
  'student_token',
  'student_mask',
  'birth_mask',
  'school_mask',
  'preferred_session_name',
  'assigned_session_name',
  'written_room_name',
  'oral_room_name',
  'oral_group_label',
  'written_time',
  'oral_time',
  'conflict_mask',
  'encrypted_payload',
  'updated_at_iso'
];

const SECURE_SCHEDULE_HEADERS = [
  'student_token',
  'student_mask',
  'assigned_session_name',
  'written_room_name',
  'written_time',
  'oral_room_name',
  'oral_group_label',
  'oral_time',
  'status_note',
  'updated_at_iso'
];

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('四技二專甄選二階智慧團體流水線排程系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;padding:24px;color:#334155;">' +
      '<h2 style="margin:0 0 12px;">GitHub 安全同步橋接已啟用</h2>' +
      '<p style="margin:0;">這個 Apps Script Web App 正在提供 GitHub Pages 前端的加密同步服務。</p>' +
      '</body></html>'
    ).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function doPost(e) {
  const payload = parseRemoteRequestPayload_(e);
  const requestId = safe_(payload.requestId);
  const clientOrigin = safe_(payload.clientOrigin);

  let bridgePayload;
  try {
    const result = withScriptLock_(function () {
      return handleRemoteRequest_(payload);
    });
    bridgePayload = { ok: true, result: result };
  } catch (error) {
    bridgePayload = {
      ok: false,
      error: error && error.message ? error.message : String(error || 'Unknown error')
    };
  }

  return createBridgeResponseHtml_(requestId, clientOrigin, bridgePayload);
}

function getBootstrapData() {
  return withScriptLock_(function () {
    const ss = openOrCreateSpreadsheet_();
    ensureInfrastructure_(ss);
    if (isSecureRemoteConfigured_()) {
      return {
        departments: [],
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        spreadsheetName: ss.getName(),
        schemaVersion: APP_SCHEMA_VERSION,
        secureMode: SECURE_STORAGE_MODE
      };
    }
    return {
      departments: loadDepartmentWorkspaces_(ss),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      spreadsheetName: ss.getName(),
      schemaVersion: APP_SCHEMA_VERSION
    };
  });
}

function saveDepartmentWorkspaces(payload) {
  return withScriptLock_(function () {
    const ss = openOrCreateSpreadsheet_();
    ensureInfrastructure_(ss);
    if (isSecureRemoteConfigured_()) {
      throw new Error('目前已啟用 GitHub 安全同步模式，請改由 GitHub 前端進行資料寫入。');
    }

    const inputDepartments = payload && Array.isArray(payload.departments) ? payload.departments : [];
    const departments = inputDepartments.map(function (dept) {
      return normalizeDepartmentWorkspace_(dept);
    });

    syncDepartmentBundles_(ss, departments);
    appendBackupSnapshot_(ss, payload && payload.source ? String(payload.source) : 'webapp_save', departments);

    return {
      ok: true,
      savedAt: new Date().toISOString(),
      departmentCount: departments.length,
      spreadsheetUrl: ss.getUrl()
    };
  });
}

function ensureAutoBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function (trigger) {
    return trigger.getHandlerFunction() === BACKUP_TRIGGER_HANDLER;
  });

  if (!exists) {
    ScriptApp.newTrigger(BACKUP_TRIGGER_HANDLER)
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .create();
  }

  return { ok: true, triggerReady: true };
}

function createTimedBackupSnapshot() {
  return withScriptLock_(function () {
    const ss = openOrCreateSpreadsheet_();
    ensureInfrastructure_(ss);
    if (isSecureRemoteConfigured_()) {
      const secureDepartments = loadSecureDepartmentWorkspaces_(ss);
      appendSecureBackupSnapshot_(ss, 'time_trigger_secure', secureDepartments);
      return { ok: true, departmentCount: secureDepartments.length, savedAt: new Date().toISOString(), secureMode: SECURE_STORAGE_MODE };
    }
    const departments = loadDepartmentWorkspaces_(ss);
    appendBackupSnapshot_(ss, 'time_trigger', departments);
    return { ok: true, departmentCount: departments.length, savedAt: new Date().toISOString() };
  });
}

function runManualBackupNow() {
  return withScriptLock_(function () {
    const ss = openOrCreateSpreadsheet_();
    ensureInfrastructure_(ss);
    if (isSecureRemoteConfigured_()) {
      const secureDepartments = loadSecureDepartmentWorkspaces_(ss);
      appendSecureBackupSnapshot_(ss, 'manual_backup_secure', secureDepartments);
      return { ok: true, departmentCount: secureDepartments.length, savedAt: new Date().toISOString(), secureMode: SECURE_STORAGE_MODE };
    }
    const departments = loadDepartmentWorkspaces_(ss);
    appendBackupSnapshot_(ss, 'manual_backup', departments);
    return { ok: true, departmentCount: departments.length, savedAt: new Date().toISOString() };
  });
}

function handleRemoteRequest_(payload) {
  const ss = openOrCreateSpreadsheet_();
  ensureInfrastructure_(ss);

  const action = safe_(payload.action);
  if (!action) {
    throw new Error('缺少 action。');
  }

  if (action === 'connect_secure') {
    const authState = verifyOrInitializeSecurePassphrase_(safe_(payload.passphrase));
    return {
      connected: true,
      newlyConfigured: authState.newlyConfigured,
      departments: loadSecureDepartmentWorkspaces_(ss),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      spreadsheetName: ss.getName(),
      schemaVersion: APP_SCHEMA_VERSION,
      secureMode: SECURE_STORAGE_MODE
    };
  }

  if (action === 'save_secure_workspaces') {
    verifyOrInitializeSecurePassphrase_(safe_(payload.passphrase));
    const inputDepartments = Array.isArray(payload.departments) ? payload.departments : [];
    const departments = inputDepartments.map(normalizeSecureDepartmentPayload_);
    syncSecureDepartmentBundles_(ss, departments);
    appendSecureBackupSnapshot_(ss, safe_(payload.source) || 'github_secure_save', departments);
    return {
      savedAt: new Date().toISOString(),
      departmentCount: departments.length,
      spreadsheetUrl: ss.getUrl(),
      secureMode: SECURE_STORAGE_MODE
    };
  }

  throw new Error('不支援的 action: ' + action);
}

function parseRemoteRequestPayload_(e) {
  const parameterPayload = e && e.parameter && e.parameter.payload ? safe_(e.parameter.payload) : '';
  if (parameterPayload) {
    try {
      return JSON.parse(parameterPayload);
    } catch (error) {
      throw new Error('payload JSON 格式錯誤。');
    }
  }

  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (error) {
      // Ignore and fall through to empty payload.
    }
  }

  return {};
}

function createBridgeResponseHtml_(requestId, clientOrigin, bridgePayload) {
  const safeTargetOrigin = sanitizePostMessageOrigin_(clientOrigin);
  const messageJson = JSON.stringify({
    __schedulerBridge: true,
    requestId: requestId,
    payload: bridgePayload
  }).replace(/</g, '\\u003c');

  const html = [
    '<!DOCTYPE html>',
    '<html><head><meta charset="UTF-8"></head><body>',
    '<script>',
    'const message = ' + messageJson + ';',
    'const targetOrigin = ' + JSON.stringify(safeTargetOrigin || '*') + ';',
    'if (window.parent && window.parent !== window) {',
    '  window.parent.postMessage(message, targetOrigin);',
    '}',
    'document.body.textContent = "bridge-ready";',
    '</script>',
    '</body></html>'
  ].join('');

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sanitizePostMessageOrigin_(origin) {
  const raw = safe_(origin);
  if (raw === '*') return '*';
  if (/^https?:\/\/[A-Za-z0-9._:-]+$/.test(raw)) return raw;
  return '*';
}

function isSecureRemoteConfigured_() {
  const props = PropertiesService.getScriptProperties();
  return !!props.getProperty(PROP_REMOTE_AUTH_HASH);
}

function verifyOrInitializeSecurePassphrase_(passphrase) {
  const secret = safe_(passphrase);
  if (secret.length < 8) {
    throw new Error('保護密碼至少需要 8 碼。');
  }

  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty(PROP_REMOTE_AUTH_SALT);
  let hash = props.getProperty(PROP_REMOTE_AUTH_HASH);

  if (!salt || !hash) {
    salt = Utilities.getUuid();
    hash = hashSecurePassphrase_(secret, salt);
    props.setProperty(PROP_REMOTE_AUTH_SALT, salt);
    props.setProperty(PROP_REMOTE_AUTH_HASH, hash);
    return { newlyConfigured: true };
  }

  if (hashSecurePassphrase_(secret, salt) !== hash) {
    throw new Error('保護密碼不正確。');
  }

  return { newlyConfigured: false };
}

function hashSecurePassphrase_(passphrase, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '::' + String(passphrase),
    Utilities.Charset.UTF_8
  );
  return bytesToHex_(digest);
}

function bytesToHex_(bytes) {
  return bytes.map(function (byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function withScriptLock_(runner) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return runner();
  } finally {
    lock.releaseLock();
  }
}

function openOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();

  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      props.setProperty(PROP_SPREADSHEET_ID, active.getId());
      return active;
    }
  } catch (error) {
    // Ignore and continue to fallback.
  }

  const savedId = props.getProperty(PROP_SPREADSHEET_ID);
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (error) {
      // Continue and create a new spreadsheet.
    }
  }

  const spreadsheet = SpreadsheetApp.create(DEFAULT_SPREADSHEET_NAME);
  props.setProperty(PROP_SPREADSHEET_ID, spreadsheet.getId());
  return spreadsheet;
}

function ensureInfrastructure_(ss) {
  const backupSheet = getOrCreateSheet_(ss, BACKUP_SHEET_NAME);
  ensureBackupSheet_(backupSheet);
  if (!backupSheet.isSheetHidden()) {
    backupSheet.hideSheet();
  }
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureBackupSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, BACKUP_HEADERS.length).setValues([BACKUP_HEADERS]);
    return;
  }

  const headerValues = sheet.getRange(1, 1, 1, BACKUP_HEADERS.length).getValues()[0];
  const headerText = headerValues.join('|');
  if (headerText !== BACKUP_HEADERS.join('|')) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, BACKUP_HEADERS.length).setValues([BACKUP_HEADERS]);
  }
}

function scanDepartmentBundles_(ss) {
  return ss.getSheets()
    .filter(function (sheet) {
      return endsWith_(sheet.getName(), SETTINGS_SUFFIX);
    })
    .map(function (settingsSheet) {
      const baseName = settingsSheet.getName().slice(0, -SETTINGS_SUFFIX.length);
      const settingsMap = readSettingsMap_(settingsSheet);
      return {
        baseName: baseName,
        departmentId: safe_(settingsMap.department_id),
        settingsSheet: settingsSheet,
        studentsSheet: ss.getSheetByName(baseName + STUDENTS_SUFFIX),
        scheduleSheet: ss.getSheetByName(baseName + SCHEDULE_SUFFIX)
      };
    });
}

function syncDepartmentBundles_(ss, departments) {
  const existingBundles = scanDepartmentBundles_(ss);
  const existingById = {};
  existingBundles.forEach(function (bundle) {
    if (bundle.departmentId) {
      existingById[bundle.departmentId] = bundle;
    }
  });

  const keepDepartmentIds = {};
  departments.forEach(function (dept) {
    const savedBundle = upsertDepartmentBundle_(ss, dept, existingById[dept.id] || null);
    keepDepartmentIds[savedBundle.departmentId] = true;
  });

  existingBundles.forEach(function (bundle) {
    if (bundle.departmentId && !keepDepartmentIds[bundle.departmentId]) {
      deleteDepartmentBundle_(ss, bundle);
    }
  });
}

function syncSecureDepartmentBundles_(ss, departments) {
  const existingBundles = scanDepartmentBundles_(ss);
  const existingById = {};
  existingBundles.forEach(function (bundle) {
    if (bundle.departmentId) {
      existingById[bundle.departmentId] = bundle;
    }
  });

  const keepDepartmentIds = {};
  departments.forEach(function (departmentPayload) {
    const savedBundle = upsertSecureDepartmentBundle_(ss, departmentPayload, existingById[departmentPayload.settings.id] || null);
    keepDepartmentIds[savedBundle.departmentId] = true;
  });

  existingBundles.forEach(function (bundle) {
    if (bundle.departmentId && !keepDepartmentIds[bundle.departmentId]) {
      deleteDepartmentBundle_(ss, bundle);
    }
  });
}

function upsertSecureDepartmentBundle_(ss, departmentPayload, existingBundle) {
  const settings = normalizeSecureDepartmentSettings_(departmentPayload.settings);
  const studentRows = normalizeSecureStudentRows_(departmentPayload.students);
  const targetBaseName = makeUniqueBundleBaseName_(ss, settings.name || settings.id, existingBundle ? existingBundle.baseName : '');

  let settingsSheet = existingBundle ? existingBundle.settingsSheet : null;
  let studentsSheet = existingBundle ? existingBundle.studentsSheet : null;
  let scheduleSheet = existingBundle ? existingBundle.scheduleSheet : null;

  if (existingBundle && existingBundle.baseName !== targetBaseName) {
    if (settingsSheet) settingsSheet.setName(targetBaseName + SETTINGS_SUFFIX);
    if (studentsSheet) studentsSheet.setName(targetBaseName + STUDENTS_SUFFIX);
    if (scheduleSheet) scheduleSheet.setName(targetBaseName + SCHEDULE_SUFFIX);
  }

  settingsSheet = settingsSheet || getOrCreateSheet_(ss, targetBaseName + SETTINGS_SUFFIX);
  studentsSheet = studentsSheet || getOrCreateSheet_(ss, targetBaseName + STUDENTS_SUFFIX);
  scheduleSheet = scheduleSheet || getOrCreateSheet_(ss, targetBaseName + SCHEDULE_SUFFIX);

  writeSecureSettingsSheet_(settingsSheet, settings, studentRows);
  writeSecureStudentsSheet_(studentsSheet, studentRows);
  writeSecureScheduleSheet_(scheduleSheet, studentRows);

  return {
    departmentId: settings.id,
    baseName: targetBaseName
  };
}

function upsertDepartmentBundle_(ss, departmentInput, existingBundle) {
  const dept = normalizeDepartmentWorkspace_(departmentInput);
  const targetBaseName = makeUniqueBundleBaseName_(ss, dept.name || dept.id, existingBundle ? existingBundle.baseName : '');

  let settingsSheet = existingBundle ? existingBundle.settingsSheet : null;
  let studentsSheet = existingBundle ? existingBundle.studentsSheet : null;
  let scheduleSheet = existingBundle ? existingBundle.scheduleSheet : null;

  if (existingBundle && existingBundle.baseName !== targetBaseName) {
    if (settingsSheet) settingsSheet.setName(targetBaseName + SETTINGS_SUFFIX);
    if (studentsSheet) studentsSheet.setName(targetBaseName + STUDENTS_SUFFIX);
    if (scheduleSheet) scheduleSheet.setName(targetBaseName + SCHEDULE_SUFFIX);
  }

  settingsSheet = settingsSheet || getOrCreateSheet_(ss, targetBaseName + SETTINGS_SUFFIX);
  studentsSheet = studentsSheet || getOrCreateSheet_(ss, targetBaseName + STUDENTS_SUFFIX);
  scheduleSheet = scheduleSheet || getOrCreateSheet_(ss, targetBaseName + SCHEDULE_SUFFIX);

  writeSettingsSheet_(settingsSheet, dept);
  writeStudentsSheet_(studentsSheet, dept);
  writeScheduleSheet_(scheduleSheet, dept);

  return {
    departmentId: dept.id,
    baseName: targetBaseName
  };
}

function deleteDepartmentBundle_(ss, bundle) {
  [bundle.settingsSheet, bundle.studentsSheet, bundle.scheduleSheet].forEach(function (sheet) {
    if (sheet) ss.deleteSheet(sheet);
  });
}

function loadDepartmentWorkspaces_(ss) {
  return scanDepartmentBundles_(ss)
    .map(function (bundle) {
      const settingsMap = readSettingsMap_(bundle.settingsSheet);
      return buildDepartmentFromBundle_(bundle, settingsMap);
    })
    .filter(function (dept) {
      return !!dept;
    })
    .sort(function (a, b) {
      return safe_(a.name).localeCompare(safe_(b.name), 'zh-Hant');
    });
}

function loadSecureDepartmentWorkspaces_(ss) {
  return scanDepartmentBundles_(ss)
    .map(function (bundle) {
      const settingsMap = readSettingsMap_(bundle.settingsSheet);
      if (safe_(settingsMap.storage_mode) !== SECURE_STORAGE_MODE) return null;
      return buildSecureDepartmentFromBundle_(bundle, settingsMap);
    })
    .filter(function (dept) {
      return !!dept;
    })
    .sort(function (a, b) {
      return safe_(a.settings && a.settings.name).localeCompare(safe_(b.settings && b.settings.name), 'zh-Hant');
    });
}

function buildSecureDepartmentFromBundle_(bundle, settingsMap) {
  return {
    settings: normalizeSecureDepartmentSettings_({
      id: safe_(settingsMap.department_id) || ('dept-' + bundle.baseName),
      name: safe_(settingsMap.department_name) || bundle.baseName,
      sub: safe_(settingsMap.department_sub),
      icon: safe_(settingsMap.department_icon) || 'cpu',
      sessions: parseJsonArray_(settingsMap.sessions_json),
      writtenRooms: parseJsonArray_(settingsMap.written_rooms_json),
      oralRooms: parseJsonArray_(settingsMap.oral_rooms_json),
      cfgWrittenDuration: safeNumber_(settingsMap.cfg_written_duration, 30),
      cfgOralDuration: safeNumber_(settingsMap.cfg_oral_duration, 20),
      cfgOralTeamsPerLadder: safeNumber_(settingsMap.cfg_oral_teams_per_ladder, 2),
      cfgOralTeamGroupCount: safeNumber_(settingsMap.cfg_oral_team_group_count, 2),
      cfgTeamCapacity: safeNumber_(settingsMap.cfg_team_capacity, 30),
      cfgOralGroupCapacity: safeNumber_(settingsMap.cfg_oral_group_capacity, 5),
      cfgSchedulerStrategy: safe_(settingsMap.cfg_scheduler_strategy) || 'preference_wave',
      cfgSchoolGrouping: safe_(settingsMap.cfg_school_grouping) || 'none',
      storageMode: SECURE_STORAGE_MODE
    }),
    students: readSecureStudentsSheet_(bundle.studentsSheet)
  };
}

function buildDepartmentFromBundle_(bundle, settingsMap) {
  if (safe_(settingsMap.storage_mode) === SECURE_STORAGE_MODE) {
    return normalizeDepartmentWorkspace_({
      id: safe_(settingsMap.department_id) || ('dept-' + bundle.baseName),
      name: safe_(settingsMap.department_name) || bundle.baseName,
      sub: safe_(settingsMap.department_sub),
      icon: safe_(settingsMap.department_icon) || 'cpu',
      sessions: parseJsonArray_(settingsMap.sessions_json),
      writtenRooms: parseJsonArray_(settingsMap.written_rooms_json),
      oralRooms: parseJsonArray_(settingsMap.oral_rooms_json),
      cfgWrittenDuration: safeNumber_(settingsMap.cfg_written_duration, 30),
      cfgOralDuration: safeNumber_(settingsMap.cfg_oral_duration, 20),
      cfgOralTeamsPerLadder: safeNumber_(settingsMap.cfg_oral_teams_per_ladder, 2),
      cfgOralTeamGroupCount: safeNumber_(settingsMap.cfg_oral_team_group_count, 2),
      cfgTeamCapacity: safeNumber_(settingsMap.cfg_team_capacity, 30),
      cfgOralGroupCapacity: safeNumber_(settingsMap.cfg_oral_group_capacity, 5),
      cfgSchedulerStrategy: safe_(settingsMap.cfg_scheduler_strategy) || 'preference_wave',
      cfgSchoolGrouping: safe_(settingsMap.cfg_school_grouping) || 'none',
      students: []
    });
  }

  const sessions = parseJsonArray_(settingsMap.sessions_json);
  const writtenRooms = parseJsonArray_(settingsMap.written_rooms_json);
  const oralRooms = parseJsonArray_(settingsMap.oral_rooms_json);
  const students = readStudentsSheet_(bundle.studentsSheet);

  return normalizeDepartmentWorkspace_({
    id: safe_(settingsMap.department_id) || ('dept-' + bundle.baseName),
    name: safe_(settingsMap.department_name) || bundle.baseName,
    sub: safe_(settingsMap.department_sub),
    icon: safe_(settingsMap.department_icon) || 'cpu',
    sessions: sessions,
    writtenRooms: writtenRooms,
    oralRooms: oralRooms,
    cfgWrittenDuration: safeNumber_(settingsMap.cfg_written_duration, 30),
    cfgOralDuration: safeNumber_(settingsMap.cfg_oral_duration, 20),
    cfgOralTeamsPerLadder: safeNumber_(settingsMap.cfg_oral_teams_per_ladder, 2),
    cfgOralTeamGroupCount: safeNumber_(settingsMap.cfg_oral_team_group_count, 2),
    cfgTeamCapacity: safeNumber_(settingsMap.cfg_team_capacity, 30),
    cfgOralGroupCapacity: safeNumber_(settingsMap.cfg_oral_group_capacity, 5),
    cfgSchedulerStrategy: safe_(settingsMap.cfg_scheduler_strategy) || 'preference_wave',
    cfgSchoolGrouping: safe_(settingsMap.cfg_school_grouping) || 'none',
    students: students
  });
}

function writeSettingsSheet_(sheet, dept) {
  const updatedAt = new Date().toISOString();
  const scheduledCount = dept.students.filter(function (student) {
    return !!safe_(student.assignedSessionId);
  }).length;

  const kv = {
    schema_version: APP_SCHEMA_VERSION,
    storage_mode: safe_(dept.storageMode),
    department_id: dept.id,
    department_name: dept.name,
    department_sub: dept.sub || '',
    department_icon: dept.icon || 'cpu',
    cfg_written_duration: String(dept.cfgWrittenDuration),
    cfg_oral_duration: String(dept.cfgOralDuration),
    cfg_oral_teams_per_ladder: String(dept.cfgOralTeamsPerLadder),
    cfg_oral_team_group_count: String(dept.cfgOralTeamGroupCount),
    cfg_team_capacity: String(dept.cfgTeamCapacity),
    cfg_oral_group_capacity: String(dept.cfgOralGroupCapacity),
    cfg_scheduler_strategy: dept.cfgSchedulerStrategy || 'preference_wave',
    cfg_school_grouping: dept.cfgSchoolGrouping || 'none',
    written_rooms_json: JSON.stringify(dept.writtenRooms || []),
    oral_rooms_json: JSON.stringify(dept.oralRooms || []),
    sessions_json: JSON.stringify(dept.sessions || []),
    students_count: String((dept.students || []).length),
    scheduled_count: String(scheduledCount),
    updated_at_iso: updatedAt
  };

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'note']]);

  const rows = SETTINGS_KEYS.map(function (pair) {
    const key = pair[0];
    return [key, key in kv ? kv[key] : '', pair[1]];
  });

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  autoResizeSheet_(sheet, 3);
}

function writeStudentsSheet_(sheet, dept) {
  const sessionMap = buildSessionMap_(dept.sessions || []);
  const writtenRooms = Array.isArray(dept.writtenRooms) ? dept.writtenRooms : [];
  const oralRooms = Array.isArray(dept.oralRooms) ? dept.oralRooms : [];
  const updatedAt = new Date().toISOString();

  sheet.clearContents();
  sheet.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]);

  const rows = (dept.students || []).map(function (studentInput) {
    const student = normalizeStudent_(studentInput);
    const preferredSession = sessionMap[student.pref] || null;
    const assignedSession = sessionMap[student.assignedSessionId] || null;
    const writtenRoomName = student.assignedWrittenRoomIdx !== null && writtenRooms[student.assignedWrittenRoomIdx] !== undefined
      ? writtenRooms[student.assignedWrittenRoomIdx]
      : '';
    const oralRoomName = student.assignedOralRoomIdx !== null && oralRooms[student.assignedOralRoomIdx] !== undefined
      ? oralRooms[student.assignedOralRoomIdx]
      : '';

    return [
      student.id,
      student.name,
      normalizeBirthDate_(student.birthDate),
      safe_(student.school),
      safe_(student.pref),
      preferredSession ? preferredSession.name : '',
      safe_(student.conflict),
      safe_(student.assignedSessionId),
      assignedSession ? assignedSession.name : '',
      student.assignedWrittenRoomIdx === null ? '' : String(student.assignedWrittenRoomIdx),
      writtenRoomName,
      student.assignedOralRoomIdx === null ? '' : String(student.assignedOralRoomIdx),
      oralRoomName,
      student.assignedOrder === null ? '' : String(student.assignedOrder),
      student.assignedOralGroupIdx === null ? '' : String(student.assignedOralGroupIdx),
      safe_(student.writtenStart),
      safe_(student.writtenEnd),
      safe_(student.oralStart),
      safe_(student.oralEnd),
      updatedAt
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, STUDENT_HEADERS.length).setValues(rows);
  }

  autoResizeSheet_(sheet, STUDENT_HEADERS.length);
}

function writeScheduleSheet_(sheet, dept) {
  const sessionMap = buildSessionMap_(dept.sessions || []);
  const writtenRooms = Array.isArray(dept.writtenRooms) ? dept.writtenRooms : [];
  const oralRooms = Array.isArray(dept.oralRooms) ? dept.oralRooms : [];
  const updatedAt = new Date().toISOString();

  sheet.clearContents();
  sheet.getRange(1, 1, 1, SCHEDULE_HEADERS.length).setValues([SCHEDULE_HEADERS]);

  const rows = (dept.students || [])
    .filter(function (student) {
      return !!safe_(student.assignedSessionId);
    })
    .map(function (studentInput) {
      const student = normalizeStudent_(studentInput);
      const preferredSession = sessionMap[student.pref] || null;
      const assignedSession = sessionMap[student.assignedSessionId] || null;
      const writtenRoomName = student.assignedWrittenRoomIdx !== null && writtenRooms[student.assignedWrittenRoomIdx] !== undefined
        ? writtenRooms[student.assignedWrittenRoomIdx]
        : '';
      const oralRoomName = student.assignedOralRoomIdx !== null && oralRooms[student.assignedOralRoomIdx] !== undefined
        ? oralRooms[student.assignedOralRoomIdx]
        : '';

      return [
        student.id,
        student.name,
        normalizeBirthDate_(student.birthDate),
        safe_(student.school),
        preferredSession ? preferredSession.name : '',
        assignedSession ? assignedSession.name : '',
        writtenRoomName,
        [safe_(student.writtenStart), safe_(student.writtenEnd)].filter(Boolean).join(' ~ '),
        oralRoomName,
        student.assignedOralGroupIdx === null ? '' : ('第 ' + (parseInt(student.assignedOralGroupIdx, 10) + 1) + ' 組'),
        [safe_(student.oralStart), safe_(student.oralEnd)].filter(Boolean).join(' ~ '),
        safe_(student.conflict),
        updatedAt
      ];
    });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SCHEDULE_HEADERS.length).setValues(rows);
  }

  autoResizeSheet_(sheet, SCHEDULE_HEADERS.length);
}

function writeSecureSettingsSheet_(sheet, settings, studentRows) {
  const normalized = normalizeSecureDepartmentSettings_(settings);
  const rows = Array.isArray(studentRows) ? studentRows : [];
  const scheduledCount = rows.filter(function (row) {
    return !!safe_(row.assignedSessionName);
  }).length;
  const kv = {
    schema_version: APP_SCHEMA_VERSION,
    storage_mode: SECURE_STORAGE_MODE,
    department_id: normalized.id,
    department_name: normalized.name,
    department_sub: normalized.sub || '',
    department_icon: normalized.icon || 'cpu',
    cfg_written_duration: String(normalized.cfgWrittenDuration),
    cfg_oral_duration: String(normalized.cfgOralDuration),
    cfg_oral_teams_per_ladder: String(normalized.cfgOralTeamsPerLadder),
    cfg_oral_team_group_count: String(normalized.cfgOralTeamGroupCount),
    cfg_team_capacity: String(normalized.cfgTeamCapacity),
    cfg_oral_group_capacity: String(normalized.cfgOralGroupCapacity),
    cfg_scheduler_strategy: normalized.cfgSchedulerStrategy || 'preference_wave',
    cfg_school_grouping: normalized.cfgSchoolGrouping || 'none',
    written_rooms_json: JSON.stringify(normalized.writtenRooms || []),
    oral_rooms_json: JSON.stringify(normalized.oralRooms || []),
    sessions_json: JSON.stringify(normalized.sessions || []),
    students_count: String(rows.length),
    scheduled_count: String(scheduledCount),
    updated_at_iso: new Date().toISOString()
  };

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'note']]);
  const outputRows = SETTINGS_KEYS.map(function (pair) {
    const key = pair[0];
    return [key, key in kv ? kv[key] : '', pair[1]];
  });
  sheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);
  autoResizeSheet_(sheet, 3);
}

function writeSecureStudentsSheet_(sheet, studentRows) {
  const rows = normalizeSecureStudentRows_(studentRows);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, SECURE_STUDENT_HEADERS.length).setValues([SECURE_STUDENT_HEADERS]);

  const outputRows = rows.map(function (row) {
    return [
      row.studentToken,
      row.studentMask,
      row.birthMask,
      row.schoolMask,
      row.preferredSessionName,
      row.assignedSessionName,
      row.writtenRoomName,
      row.oralRoomName,
      row.oralGroupLabel,
      row.writtenTime,
      row.oralTime,
      row.conflictMask,
      row.encryptedPayload,
      row.updatedAtIso
    ];
  });

  if (outputRows.length > 0) {
    sheet.getRange(2, 1, outputRows.length, SECURE_STUDENT_HEADERS.length).setValues(outputRows);
  }

  autoResizeSheet_(sheet, SECURE_STUDENT_HEADERS.length);
}

function writeSecureScheduleSheet_(sheet, studentRows) {
  const rows = normalizeSecureStudentRows_(studentRows)
    .filter(function (row) {
      return !!safe_(row.assignedSessionName);
    })
    .map(function (row) {
      return [
        row.studentToken,
        row.studentMask,
        row.assignedSessionName,
        row.writtenRoomName,
        row.writtenTime,
        row.oralRoomName,
        row.oralGroupLabel,
        row.oralTime,
        row.conflictMask,
        row.updatedAtIso
      ];
    });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, SECURE_SCHEDULE_HEADERS.length).setValues([SECURE_SCHEDULE_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SECURE_SCHEDULE_HEADERS.length).setValues(rows);
  }
  autoResizeSheet_(sheet, SECURE_SCHEDULE_HEADERS.length);
}

function readSettingsMap_(sheet) {
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  values.forEach(function (row) {
    const key = safe_(row[0]);
    if (!key) return;
    map[key] = row[1];
  });
  return map;
}

function readStudentsSheet_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, STUDENT_HEADERS.length).getValues();
  return values
    .map(function (row) {
      const item = {};
      STUDENT_HEADERS.forEach(function (header, idx) {
        item[header] = row[idx];
      });

      return normalizeStudent_({
        id: safe_(item.student_id),
        name: safe_(item.student_name),
        birthDate: normalizeBirthDate_(item.birth_date),
        school: safe_(item.school),
        pref: safe_(item.preferred_session_id),
        conflict: safe_(item.conflict_time),
        assignedSessionId: safe_(item.assigned_session_id),
        assignedWrittenRoomIdx: parseNullableNumber_(item.assigned_written_room_idx),
        assignedOralRoomIdx: parseNullableNumber_(item.assigned_oral_room_idx),
        assignedOrder: parseNullableNumber_(item.assigned_order),
        assignedOralGroupIdx: parseNullableNumber_(item.assigned_oral_group_idx),
        writtenStart: safe_(item.written_start),
        writtenEnd: safe_(item.written_end),
        oralStart: safe_(item.oral_start),
        oralEnd: safe_(item.oral_end)
      });
    })
    .filter(function (student) {
      return !!student.id || !!student.name;
    });
}

function readSecureStudentsSheet_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headerWidth = SECURE_STUDENT_HEADERS.length;
  const values = sheet.getRange(2, 1, lastRow - 1, headerWidth).getValues();
  return values.map(function (row) {
    const item = {};
    SECURE_STUDENT_HEADERS.forEach(function (header, idx) {
      item[header] = row[idx];
    });
    return normalizeSecureStudentRow_(item);
  }).filter(function (row) {
    return !!row.encryptedPayload;
  });
}

function appendBackupSnapshot_(ss, source, departments) {
  const sheet = getOrCreateSheet_(ss, BACKUP_SHEET_NAME);
  ensureBackupSheet_(sheet);

  const payload = JSON.stringify({
    departments: departments,
    spreadsheetId: ss.getId(),
    schemaVersion: APP_SCHEMA_VERSION
  });

  sheet.appendRow([
    new Date().toISOString(),
    safe_(source) || 'unknown',
    APP_SCHEMA_VERSION,
    String((departments || []).length),
    payload
  ]);

  trimBackupSheet_(sheet, 200);
}

function appendSecureBackupSnapshot_(ss, source, departments) {
  const sheet = getOrCreateSheet_(ss, BACKUP_SHEET_NAME);
  ensureBackupSheet_(sheet);

  const payload = JSON.stringify({
    storageMode: SECURE_STORAGE_MODE,
    departments: departments,
    spreadsheetId: ss.getId(),
    schemaVersion: APP_SCHEMA_VERSION
  });

  sheet.appendRow([
    new Date().toISOString(),
    safe_(source) || 'secure_unknown',
    APP_SCHEMA_VERSION,
    String((departments || []).length),
    payload
  ]);

  trimBackupSheet_(sheet, 200);
}

function trimBackupSheet_(sheet, maxEntries) {
  const lastRow = sheet.getLastRow();
  const dataRows = Math.max(0, lastRow - 1);
  if (dataRows <= maxEntries) return;

  const deleteCount = dataRows - maxEntries;
  sheet.deleteRows(2, deleteCount);
}

function makeUniqueBundleBaseName_(ss, rawName, exceptBaseName) {
  const desiredBase = sanitizeSheetBaseName_(rawName);
  const exceptNames = exceptBaseName ? getBundleSheetNames_(exceptBaseName) : [];
  let candidate = desiredBase || '未命名系所';
  let suffix = 1;

  while (!canUseBundleBaseName_(ss, candidate, exceptNames)) {
    suffix += 1;
    const tail = '_' + suffix;
    candidate = (desiredBase || '系所').slice(0, Math.max(1, SHEET_NAME_MAX_LENGTH - tail.length)) + tail;
  }

  return candidate;
}

function canUseBundleBaseName_(ss, baseName, exceptSheetNames) {
  const targetNames = getBundleSheetNames_(baseName);
  return targetNames.every(function (sheetName) {
    if (exceptSheetNames.indexOf(sheetName) !== -1) return true;
    return !ss.getSheetByName(sheetName);
  });
}

function getBundleSheetNames_(baseName) {
  return [
    baseName + SETTINGS_SUFFIX,
    baseName + STUDENTS_SUFFIX,
    baseName + SCHEDULE_SUFFIX
  ];
}

function sanitizeSheetBaseName_(name) {
  const cleaned = safe_(name)
    .replace(/[\[\]\*\?\/\\:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const fallback = cleaned || '未命名系所';
  return fallback.slice(0, SHEET_NAME_MAX_LENGTH);
}

function buildSessionMap_(sessions) {
  const map = {};
  (Array.isArray(sessions) ? sessions : []).forEach(function (session) {
    if (!session || !session.id) return;
    map[session.id] = session;
  });
  return map;
}

function normalizeDepartmentWorkspace_(dept) {
  const safeDept = dept || {};
  return {
    id: safe_(safeDept.id) || ('dept-' + new Date().getTime()),
    name: safe_(safeDept.name) || '未命名系所',
    sub: safe_(safeDept.sub),
    icon: safe_(safeDept.icon) || 'cpu',
    writtenRooms: Array.isArray(safeDept.writtenRooms) && safeDept.writtenRooms.length ? safeDept.writtenRooms.map(safe_) : ['筆試一試室'],
    oralRooms: Array.isArray(safeDept.oralRooms) && safeDept.oralRooms.length ? safeDept.oralRooms.map(safe_) : ['口試一教室'],
    sessions: normalizeSessions_(safeDept.sessions),
    cfgWrittenDuration: safeNumber_(safeDept.cfgWrittenDuration, 30),
    cfgOralDuration: safeNumber_(safeDept.cfgOralDuration, 20),
    cfgOralTeamsPerLadder: safeNumber_(safeDept.cfgOralTeamsPerLadder, 2),
    cfgOralTeamGroupCount: safeNumber_(safeDept.cfgOralTeamGroupCount, 2),
    cfgTeamCapacity: safeNumber_(safeDept.cfgTeamCapacity, 30),
    cfgOralGroupCapacity: safeNumber_(safeDept.cfgOralGroupCapacity, 5),
    cfgSchedulerStrategy: safe_(safeDept.cfgSchedulerStrategy) || 'preference_wave',
    cfgSchoolGrouping: safe_(safeDept.cfgSchoolGrouping) || 'none',
    students: Array.isArray(safeDept.students) ? safeDept.students.map(normalizeStudent_) : []
  };
}

function normalizeSecureDepartmentPayload_(departmentPayload) {
  const payload = departmentPayload || {};
  return {
    settings: normalizeSecureDepartmentSettings_(payload.settings),
    students: normalizeSecureStudentRows_(payload.students)
  };
}

function normalizeSecureDepartmentSettings_(settings) {
  const normalized = normalizeDepartmentWorkspace_(settings || {});
  return {
    id: normalized.id,
    name: normalized.name,
    sub: normalized.sub,
    icon: normalized.icon,
    writtenRooms: normalized.writtenRooms,
    oralRooms: normalized.oralRooms,
    sessions: normalized.sessions,
    cfgWrittenDuration: normalized.cfgWrittenDuration,
    cfgOralDuration: normalized.cfgOralDuration,
    cfgOralTeamsPerLadder: normalized.cfgOralTeamsPerLadder,
    cfgOralTeamGroupCount: normalized.cfgOralTeamGroupCount,
    cfgTeamCapacity: normalized.cfgTeamCapacity,
    cfgOralGroupCapacity: normalized.cfgOralGroupCapacity,
    cfgSchedulerStrategy: normalized.cfgSchedulerStrategy,
    cfgSchoolGrouping: normalized.cfgSchoolGrouping,
    storageMode: SECURE_STORAGE_MODE
  };
}

function normalizeSecureStudentRows_(rows) {
  return (Array.isArray(rows) ? rows : []).map(normalizeSecureStudentRow_);
}

function normalizeSecureStudentRow_(row) {
  const safeRow = row || {};
  return {
    studentToken: safe_(safeRow.studentToken || safeRow.student_token),
    studentMask: safe_(safeRow.studentMask || safeRow.student_mask),
    birthMask: safe_(safeRow.birthMask || safeRow.birth_mask),
    schoolMask: safe_(safeRow.schoolMask || safeRow.school_mask),
    preferredSessionName: safe_(safeRow.preferredSessionName || safeRow.preferred_session_name),
    assignedSessionName: safe_(safeRow.assignedSessionName || safeRow.assigned_session_name),
    writtenRoomName: safe_(safeRow.writtenRoomName || safeRow.written_room_name),
    oralRoomName: safe_(safeRow.oralRoomName || safeRow.oral_room_name),
    oralGroupLabel: safe_(safeRow.oralGroupLabel || safeRow.oral_group_label),
    writtenTime: safe_(safeRow.writtenTime || safeRow.written_time),
    oralTime: safe_(safeRow.oralTime || safeRow.oral_time),
    conflictMask: safe_(safeRow.conflictMask || safeRow.conflict_mask),
    encryptedPayload: safe_(safeRow.encryptedPayload || safeRow.encrypted_payload),
    updatedAtIso: safe_(safeRow.updatedAtIso || safeRow.updated_at_iso) || new Date().toISOString()
  };
}

function normalizeSessions_(sessions) {
  const list = Array.isArray(sessions) ? sessions : [];
  if (!list.length) {
    return [{ id: 's1', name: '時段一', startTime: '09:00' }];
  }

  return list.map(function (session, idx) {
    return {
      id: safe_(session && session.id) || ('s' + (idx + 1)),
      name: safe_(session && session.name) || ('時段' + (idx + 1)),
      startTime: safe_(session && session.startTime) || '09:00'
    };
  });
}

function normalizeStudent_(student) {
  const safeStudent = student || {};
  return {
    id: safe_(safeStudent.id),
    name: safe_(safeStudent.name),
    birthDate: normalizeBirthDate_(safeStudent.birthDate || safeStudent.birth_date || safeStudent.birthday || safeStudent.dob),
    school: safe_(safeStudent.school),
    pref: safe_(safeStudent.pref || safeStudent.preferred_session_id),
    conflict: safe_(safeStudent.conflict || safeStudent.conflict_time),
    assignedSessionId: safe_(safeStudent.assignedSessionId || safeStudent.assigned_session_id),
    assignedWrittenRoomIdx: parseNullableNumber_(safeStudent.assignedWrittenRoomIdx),
    assignedOralRoomIdx: parseNullableNumber_(safeStudent.assignedOralRoomIdx),
    assignedOrder: parseNullableNumber_(safeStudent.assignedOrder),
    assignedOralGroupIdx: parseNullableNumber_(safeStudent.assignedOralGroupIdx),
    writtenStart: safe_(safeStudent.writtenStart),
    writtenEnd: safe_(safeStudent.writtenEnd),
    oralStart: safe_(safeStudent.oralStart),
    oralEnd: safe_(safeStudent.oralEnd)
  };
}

function normalizeBirthDate_(value) {
  const raw = safe_(value).trim();
  if (!raw) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  const digits = raw.replace(/[^\d]/g, '');
  if (digits.length === 8) {
    return digits.slice(0, 4) + '-' + digits.slice(4, 6) + '-' + digits.slice(6, 8);
  }
  if (digits.length === 7) {
    const year = String(parseInt(digits.slice(0, 3), 10) + 1911);
    return year + '-' + digits.slice(3, 5) + '-' + digits.slice(5, 7);
  }
  if (digits.length === 6) {
    const year2 = String(parseInt(digits.slice(0, 2), 10) + 1911);
    return year2 + '-' + digits.slice(2, 4) + '-' + digits.slice(4, 6);
  }

  return raw;
}

function parseJsonArray_(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function parseNullableNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const num = parseInt(value, 10);
  return isNaN(num) ? null : num;
}

function safeNumber_(value, fallback) {
  const num = parseInt(value, 10);
  return isNaN(num) ? fallback : num;
}

function safe_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function endsWith_(value, suffix) {
  return String(value).slice(-suffix.length) === suffix;
}

function autoResizeSheet_(sheet, columnCount) {
  try {
    sheet.autoResizeColumns(1, columnCount);
  } catch (error) {
    // Ignore resize failures in Apps Script quotas.
  }
}
