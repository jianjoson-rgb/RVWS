/**
 * Shopify Global Accounts Action Registry
 * Code.gs | Version 2.16.6
 */

const SPREADSHEET_ID = '1q736D72n1E74ngWOG-TARRLG4XuidaJ8-_BCgQ1ir3o';

const ACTIONTRACKER_VERSION = '2.16.6';

const REGISTRY_SECURITY_HEADERS = Object.freeze([
  'Event_ID', 'Occurred_At_UTC', 'Actor_Email', 'Operation', 'Requested_Record_ID',
  'Outcome', 'Reason_Code', 'App_Version'
]);

function registrySecurityError_(message, code) {
  const error = new Error(message);
  error.registrySecurityCode = code;
  return error;
}

function registryAuditOperation_(operation, requestedRecordId, work, logSuccess, scheduled) {
  try {
    const result = work();
    if (logSuccess) registryWriteSecurityEvent_(operation, requestedRecordId, 'SUCCESS', 'OPERATION_COMPLETED', scheduled);
    return result;
  } catch (error) {
    const denied = ['IDENTITY_UNVERIFIED', 'ACCOUNT_INACTIVE', 'PERMISSION_REQUIRED',
      'OWNER_REQUIRED', 'APPROVAL_SITE_REQUIRED', 'FOLLOWUP_ASSIGNEE_REQUIRED',
      'TRIGGER_IDENTITY_REQUIRED', 'TRIGGER_OWNER_REQUIRED', 'PLAN_COMPLETED'];
    const code = error && error.registrySecurityCode;
    const reason = denied.indexOf(code) >= 0 || code === 'STALE_VERSION' ? code : 'OPERATION_FAILED';
    registryWriteSecurityEvent_(operation, requestedRecordId,
      denied.indexOf(code) >= 0 ? 'DENIED' : 'FAILED', reason, scheduled);

    throw error;
  }
}

function registryWriteSecurityEvent_(operation, requestedRecordId, outcome, reason, scheduled) {
  const event = {
    Event_ID: '', Occurred_At_UTC: new Date().toISOString(), Actor_Email: 'UNVERIFIED',
    Operation: operation, Requested_Record_ID: '',
    Outcome: outcome, Reason_Code: reason, App_Version: ACTIONTRACKER_VERSION
  };

  if (typeof requestedRecordId === 'string' && requestedRecordId) {
    event.Requested_Record_ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,159}$/.test(requestedRecordId) ?
      requestedRecordId : '[INVALID_ID]';
  }
  let lock = null, acquired = false;
  try {

    const user = scheduled ? Session.getEffectiveUser() : Session.getActiveUser();
    const email = String(user.getEmail() || '').trim().toLowerCase();
    if (/^[A-Z0-9._%+\-]+@taskus\.com$/i.test(email)) event.Actor_Email = email;
    event.Event_ID = 'SEC-' + Utilities.getUuid();
    lock = LockService.getScriptLock();
    lock.waitLock(5000); acquired = true;
    const ss = getTrackerSpreadsheet_();
    let sheet = ss.getSheetByName('Security_Log');
    if (!sheet) sheet = ss.insertSheet('Security_Log');
    if (!sheet.getLastRow()) {
      if (sheet.getMaxColumns() < REGISTRY_SECURITY_HEADERS.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), REGISTRY_SECURITY_HEADERS.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, REGISTRY_SECURITY_HEADERS.length).setValues([REGISTRY_SECURITY_HEADERS.slice()]);
      sheet.setFrozenRows(1);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const map = requireHeaders_(headers, REGISTRY_SECURITY_HEADERS, 'Security_Log');
    appendNamedLog_({ sheet: sheet, headers: headers, map: map }, event);
    SpreadsheetApp.flush();
  } catch (loggingError) {

    try { console.warn(JSON.stringify({ registrySecurityEvent: event, delivery: 'EXECUTION_LOG_FALLBACK' })); }
    catch (ignored) {   }
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch (ignored) {   }
    }
  }
}

function submitNewAction(payload, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitNewAction', '', function() {
    const result = submitNewActionCore_.apply(null, args);
    return registryQueueCommittedEvent_('REGISTERED', '', requestToken, result);
  }, false, false);
}

function submitActionUpdate(actionPlanId, currentPerformance, actionStatus, comment, requestToken, expectedRecordVersion) {
  const args = arguments;
  return registryAuditOperation_('submitActionUpdate', actionPlanId, function() {
    const result = submitActionUpdateCore_.apply(null, args);
    return registryQueueCommittedEvent_('EXECUTION_UPDATED', actionPlanId, requestToken, result);
  }, false, false);
}

function submitLeadershipComment(actionPlanId, commentText, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitLeadershipComment', actionPlanId, function() {
    const result = submitLeadershipCommentCore_.apply(null, args);
    return registryQueueCommittedEvent_('FEEDBACK_ADDED', actionPlanId, requestToken, result);
  }, false, false);
}

function uploadActionEvidence(formObject) {
  const args = arguments;
  return registryAuditOperation_('uploadActionEvidence', (formObject && formObject.actionPlanId) || '', function() {
    return uploadActionEvidenceCore_.apply(null, args);
  }, false, false);
}

function submitReviewDecision(actionPlanId, decision, comment, expectedPlanVersion, expectedRecordVersion, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitReviewDecision', actionPlanId, function() {
    const result = submitReviewDecisionCore_.apply(null, args);
    return registryQueueCommittedEvent_('REVIEW_DECIDED', actionPlanId, requestToken, result);
  }, false, false);
}

function submitPlanRevision(actionPlanId, payload, revisionComment, expectedPlanVersion, expectedRecordVersion, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitPlanRevision', actionPlanId, function() {
    const result = submitPlanRevisionCore_.apply(null, args);
    return registryQueueCommittedEvent_('PLAN_REVISED', actionPlanId, requestToken, result);
  }, false, false);
}

function submitEffectivenessAssessment(actionPlanId, payload, expectedPlanVersion, expectedRecordVersion, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitEffectivenessAssessment', actionPlanId, function() {
    const result = submitEffectivenessAssessmentCore_.apply(null, args);
    return registryQueueCommittedEvent_('IMPACT_ASSESSED', actionPlanId, requestToken, result);
  }, false, false);
}

function submitRegistryFollowUpEvent(followUpId, payload, expectedVersion, requestToken) {
  const args = arguments;
  return registryAuditOperation_('submitRegistryFollowUpEvent', followUpId, function() {
    const result = submitRegistryFollowUpEventCore_.apply(null, args);
    return registryQueueCommittedEvent_('FOLLOWUP_UPDATED', '', requestToken, result);
  }, false, false);
}

function setupRegistryAssessments() {
  const args = arguments;
  return registryAuditOperation_('setupRegistryAssessments', '', function() {
    return setupRegistryAssessmentsCore_.apply(null, args);
  }, true, false);
}

function setupRegistryFollowUps() {
  const args = arguments;
  return registryAuditOperation_('setupRegistryFollowUps', '', function() {
    return setupRegistryFollowUpsCore_.apply(null, args);
  }, true, false);
}

function resolveRegistryNotification(notificationId, outcome, notes) {
  const args = arguments;
  return registryAuditOperation_('resolveRegistryNotification', notificationId, function() {
    return resolveRegistryNotificationCore_.apply(null, args);
  }, true, false);
}

function setRegistryLiveEmailMode(mode, requestToken) {
  const args = arguments;
  return registryAuditOperation_('setRegistryLiveEmailMode', '', function() {
    return setRegistryLiveEmailModeCore_.apply(null, args);
  }, true, false);
}

function sendRegistryFollowUpEmailsNow(kind, requestToken) {
  const args = arguments;
  return registryAuditOperation_('sendRegistryFollowUpEmailsNow', '', function() {
    return sendRegistryFollowUpEmailsNowCore_.apply(null, args);
  }, true, false);
}

function installRegistryFollowUpReminderTrigger_() {
  const args = arguments;
  return registryAuditOperation_('installRegistryFollowUpReminderTrigger_', '', function() {
    return installRegistryFollowUpReminderTriggerCore_.apply(null, args);
  }, true, false);
}

function runRegistryFollowUpReminders_() {
  const args = arguments;
  return registryAuditOperation_('runRegistryFollowUpReminders_', '', function() {
    return runRegistryFollowUpRemindersCore_.apply(null, args);
  }, true, true);
}

const ACCESS_SHEET_NAME = 'Access_Control';
const FRAMEWORK_SHEET_NAME = 'Action_Framework';
const COMMENT_LOG_SHEET_NAME = 'Comment_Log';
const EVIDENCE_LOG_SHEET_NAME = 'Evidence_Log';
const ACTIONTRACKER_MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
const ACTIONTRACKER_EXTENSION_HEADERS = Object.freeze([
  'Department', 'Framework_Category', 'Action_Type', 'Activity',
  'Plan_Version', 'Record_Version', 'Review_Status', 'Latest_Review_ID',
  'Effectiveness_Status', 'Latest_Assessment_ID',
  'Approval_Required', 'Channel', 'LOB', 'Cohort', 'Cohort_Details'
]);
const ACTIONTRACKER_COMMENT_HEADERS = Object.freeze([
  'Comment_ID', 'Action_Plan_ID', 'Commenter_Email', 'Commenter_Name',
  'Commented_At', 'Comment_Text', 'Request_ID'
]);
const ACTIONTRACKER_EVIDENCE_HEADERS = Object.freeze([
  'Evidence_ID', 'Action_Plan_ID', 'Plan_Version', 'File_ID', 'File_Name',
  'File_URL', 'Evidence_Description', 'Uploaded_By', 'Uploaded_At', 'Request_ID'
]);

const FORM_SHEET_NAME = 'Form_Responses';
const MASTER_SHEET_NAME = 'Action_Master';
const UPDATE_LOG_SHEET_NAME = 'Update_Log';
const SUBMISSION_LOG_SHEET_NAME = 'Submission_Log';
const SETTINGS_SHEET_NAME = 'Settings';

const ACTIONTRACKER_ALLOWED_STATUSES = Object.freeze([
  'Not Started',
  'Ongoing',
  'On Hold',
  'Completed',
  'Cancelled'
]);

const ACTIONTRACKER_DEFAULT_SITES = Object.freeze([
  'Global',
  'ARN',
  'HTW',
  'MTQ',
  'PHX'
]);

const ACTIONTRACKER_ACTION_CATEGORIES = Object.freeze([
  'Coaching',
  'Training',
  'Process Improvement',
  'QA Intervention',
  'WFM Intervention',
  'Performance Management',
  'Knowledge / Policy',
  'Technology / Tool',
  'Security & Compliance',
  'Other'
]);

const ACTIONTRACKER_MASTER_FORMULA_COLUMNS = Object.freeze({
  'Days Remaining': 16,
  'Performance Change': 17,
  'Gap to Goal': 18,
  'Goal Progress %': 19,
  'Schedule Health': 20
});

function getTrackerSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function normalizeHeaderKey_(value) {
  const key = String(value == null ? '' : value).trim().toLowerCase();

  return /^site\s*\/\s*department$/.test(key) ? 'site' : key;
}

function buildHeaderMapStrict_(headers, sheetName) {
  const map = {};
  const duplicates = [];

  headers.forEach(function(header, index) {
    const key = normalizeHeaderKey_(header);
    if (!key) return;

    if (Object.prototype.hasOwnProperty.call(map, key)) {
      duplicates.push(String(header).trim());
      return;
    }

    map[key] = index;
  });

  if (duplicates.length) {
    throw new Error(
      sheetName + ' has duplicate header(s): ' +
      Array.from(new Set(duplicates)).join(', ')
    );
  }

  return map;
}

function requireHeaders_(headers, requiredHeaders, sheetName) {
  const map = buildHeaderMapStrict_(headers, sheetName);
  const missing = requiredHeaders.filter(function(header) {
    return !Object.prototype.hasOwnProperty.call(
      map,
      normalizeHeaderKey_(header)
    );
  });

  if (missing.length) {
    throw new Error(
      sheetName + ' is missing required header(s): ' + missing.join(', ')
    );
  }

  return map;
}

function assertMasterFormulaSchema_(headers) {
  const map = requireHeaders_(
    headers,
    Object.keys(ACTIONTRACKER_MASTER_FORMULA_COLUMNS),
    MASTER_SHEET_NAME
  );

  Object.keys(ACTIONTRACKER_MASTER_FORMULA_COLUMNS).forEach(function(header) {
    const actualColumn = map[normalizeHeaderKey_(header)] + 1;
    const expectedColumn = ACTIONTRACKER_MASTER_FORMULA_COLUMNS[header];

    if (actualColumn !== expectedColumn) {
      throw new Error(
        MASTER_SHEET_NAME + ' schema mismatch: "' + header +
        '" must remain in column ' + expectedColumn +
        ' but is currently in column ' + actualColumn + '.'
      );
    }
  });
}

function requireNonBlankText_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    throw new Error(label + ' is required.');
  }
  return text;
}

function normalizeActionTrackerSite_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';

  if (text.toUpperCase() === 'MYQ') {
    return 'MTQ';
  }

  const canonical = ACTIONTRACKER_DEFAULT_SITES.find(function(site) {
    return site.toLowerCase() === text.toLowerCase();
  });

  return canonical || text;
}

function dateToClientIso_(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return '';
  }

  return value.toISOString();
}

function validateTaskUsOwnerEmail_(value) {
  const email = requireNonBlankText_(value, 'Action Owner').toLowerCase();

  if (!/^[a-z0-9._%+-]+@taskus\.com$/i.test(email)) {
    throw new Error(
      'Action Owner must be a valid @taskus.com email address.'
    );
  }

  return email;
}

function parseStoredMetricNumber_(value, unit) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const text = String(value == null ? '' : value).trim();
  if (!text) return null;

  const normalized = text.replace(/,/g, '');

  if (/^-?\d{1,3}:\d{2}:\d{2}$/.test(normalized)) {
    const negative = normalized.charAt(0) === '-';
    const clean = negative ? normalized.slice(1) : normalized;
    const parts = clean.split(':').map(Number);

    if (
      parts.length !== 3 ||
      parts.some(function(part) { return !Number.isFinite(part); }) ||
      parts[1] > 59 ||
      parts[2] > 59
    ) {
      return null;
    }

    const minutes = parts[0] * 60 + parts[1] + parts[2] / 60;
    return negative ? -minutes : minutes;
  }

  const number = Number(normalized.replace('%', ''));
  if (!Number.isFinite(number)) return null;

  if (String(unit || '').trim() === '%' && normalized.indexOf('%') !== -1) {
    return number / 100;
  }

  return number;
}

function registryMetricDelta_(baselineValue, currentValue, goalValue, direction, unit, gap) {
  const current = parseStoredMetricNumber_(currentValue, unit);
  const reference = parseStoredMetricNumber_(gap ? goalValue : baselineValue, unit);
  if (current === null || reference === null) return null;
  if (!gap) return current - reference;
  const normalized = String(direction || '').trim().toLowerCase().replace(/\s+is better$/, '');
  if (normalized !== 'higher' && normalized !== 'lower') return null;
  return normalized === 'lower' ? current - reference : reference - current;
}

function deriveOutcomeStatus_(baselineValue, currentValue, goalValue, direction, unit) {
  const baseline = parseStoredMetricNumber_(baselineValue, unit);
  const current = parseStoredMetricNumber_(currentValue, unit);
  const goal = parseStoredMetricNumber_(goalValue, unit);
  const normalizedDirection = String(direction || '').trim().toLowerCase().replace(/\s+is better$/, '');

  if (
    baseline === null ||
    current === null ||
    goal === null ||
    (normalizedDirection !== 'higher' && normalizedDirection !== 'lower')
  ) {
    return 'Insufficient Data';
  }

  if (
    (normalizedDirection === 'higher' && current >= goal) ||
    (normalizedDirection === 'lower' && current <= goal)
  ) {
    return 'Goal Met';
  }

  const tolerance = Math.max(
    Math.abs(goal - baseline) * 0.01,
    Math.max(Math.abs(baseline), Math.abs(goal), 1) * 1e-9
  );

  if (Math.abs(current - baseline) <= tolerance) {
    return 'Unchanged';
  }

  const improving = normalizedDirection === 'higher'
    ? current > baseline
    : current < baseline;

  return improving ? 'Improving' : 'Worsening';
}

function getSettingsOptionValues_(headerName) {
  const ss = getTrackerSpreadsheet_();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) {
    return [];
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  const map = buildHeaderMapStrict_(headers, SETTINGS_SHEET_NAME);
  const index = map[normalizeHeaderKey_(headerName)];

  if (index === undefined) {
    return [];
  }

  return sheet
    .getRange(2, index + 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .map(function(value) {
      return String(value || '').trim();
    })
    .filter(function(value) {
      return value !== '';
    });
}

function protectLiteralText_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function setLiteralCellValue_(range, value) {
  range.setValue(protectLiteralText_(value));
}

function isAhtKpiServer_(kpiName) {
  const text = String(kpiName || '').trim().toLowerCase();
  return (
    /(^|\W)aht($|\W)/.test(text) ||
    text.indexOf('average handle time') !== -1 ||
    text.indexOf('average handling time') !== -1
  );
}

function parseKpiMetricServer_(value, kpiName, kpiSetting, label) {
  if (
    value === null ||
    value === undefined ||
    typeof value === 'boolean' ||
    String(value).trim() === ''
  ) {
    throw new Error(label + ' is required.');
  }

  const unit = String((kpiSetting && kpiSetting.unit) || '').trim();
  const rawText = String(value).trim().replace(/,/g, '');
  let numericValue;

  if (isAhtKpiServer_(kpiName)) {
    if (rawText.indexOf(':') !== -1) {
      const match = rawText.match(/^(\d+):(\d{2}):(\d{2})$/);
      if (!match) {
        throw new Error(label + ' must use HH:MM:SS for AHT.');
      }

      const hours = Number(match[1]);
      const minutes = Number(match[2]);
      const seconds = Number(match[3]);

      if (minutes >= 60 || seconds >= 60) {
        throw new Error(label + ' contains an invalid AHT time.');
      }

      numericValue = hours * 60 + minutes + seconds / 60;
    } else {
      numericValue = Number(rawText);
    }

    if (!Number.isFinite(numericValue) || numericValue < 0) {
      throw new Error(label + ' must be a non-negative AHT value.');
    }

    return numericValue;
  }

  if (unit === '%') {
    if (rawText.endsWith('%')) {
      if (!rawText.slice(0, -1).trim()) throw new Error(label + ' must contain a percentage value.');
      numericValue = Number(rawText.slice(0, -1)) / 100;
    } else {
      numericValue = Number(rawText);
    }

    if (!Number.isFinite(numericValue)) {
      throw new Error(label + ' must be numeric.');
    }

    if (numericValue < 0 || numericValue > 1) {
      throw new Error(
        label + ' must be from 0 to 1, or entered with a % sign ' +
        '(for example 82% or 0.82).'
      );
    }

    return numericValue;
  }

  numericValue = Number(rawText);

  if (!Number.isFinite(numericValue)) {
    throw new Error(label + ' must be numeric.');
  }

  if (unit.toLowerCase() === 'count' && (!Number.isSafeInteger(numericValue) || numericValue < 0)) {
    throw new Error(label + ' must be a non-negative whole number for a Count KPI.');
  }
  return numericValue;
}

function validateKpiPerformanceServer_(value, kpiName, kpiSetting) {
  return parseKpiMetricServer_(
    value,
    kpiName,
    kpiSetting,
    'Current Performance'
  );
}

function validateAllowedStatus_(status) {
  const clean = String(status == null ? '' : status).trim();

  if (!clean) {
    throw new Error('Action Status is required.');
  }

  if (ACTIONTRACKER_ALLOWED_STATUSES.indexOf(clean) === -1) {
    throw new Error('Invalid Action Status: ' + clean);
  }

  return clean;
}

function parseDateInputServer_(value, label) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  const text = String(value == null ? '' : value).trim();
  if (!text) {
    throw new Error(label + ' is required.');
  }

  let year;
  let month;
  let day;
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (match) {
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  } else {
    match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) {
      throw new Error(label + ' must be a valid date.');
    }
    month = Number(match[1]);
    day = Number(match[2]);
    year = Number(match[3]);
  }

  const parsed = new Date(year, month - 1, day);

  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    throw new Error(label + ' must be a valid calendar date.');
  }

  return parsed;
}

function requireAuthenticatedWriter_() {
  return requireActiveActor_().email;
}

function getAccessEntry_(email) {
  const table = readTable_(ACCESS_SHEET_NAME,
    ['Email', 'Name', 'Role', 'Active', 'Can_Submit', 'Can_Comment', 'Can_Audit', 'Can_Approve', 'Approval_Sites']);
  const matches = table.rows.filter(function(row) {
    return String(row[table.map.email] || '').trim().toLowerCase() === email;
  });
  if (matches.length > 1) throw new Error('Duplicate email in Access_Control. Ask the maintainer to correct it.');
  if (!matches.length) return null;
  const row = matches[0];
  return {
    email: email, name: String(row[table.map.name] || '').trim(),
    active: row[table.map.active] === true,
    canSubmit: row[table.map.can_submit] === true,
    canComment: row[table.map.can_comment] === true,
    canAudit: row[table.map.can_audit] === true,
    canAdmin: table.map.can_admin !== undefined && row[table.map.can_admin] === true,
    canApprove: row[table.map.can_approve] === true,
    approvalSites: parseApprovalSites_(row[table.map.approval_sites])
  };
}

function requireActiveActor_(flag) {

  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email || !/^[A-Z0-9._%+\-]+@taskus\.com$/i.test(email)) {
    throw registrySecurityError_('Your signed-in work identity could not be verified. Writes are blocked. Open the app using your authorized TaskUs account.', 'IDENTITY_UNVERIFIED');
  }
  const actor = getAccessEntry_(email);
  if (!actor || !actor.active) throw registrySecurityError_('Your account is not active in Access_Control. Writes are blocked.', 'ACCOUNT_INACTIVE');
  if (flag && !actor[flag]) throw registrySecurityError_(flag === 'canAdmin' ?
    'This operation requires Can_Admin checked in Access_Control.' :
    'You do not have permission for this operation.', 'PERMISSION_REQUIRED');
  return actor;
}

function getAccessContext() {
  try {
    const actor = requireActiveActor_();
    return Object.assign({}, actor, { verified: true, reviewEnabled: true, assessmentEnabled: true, admin: false });
  } catch (error) {
    return { verified: false, active: false, canSubmit: false, canComment: false,
      canAudit: false, canApprove: false, approvalSites: [], reviewEnabled: true, assessmentEnabled: true,
      admin: false, message: error.message };
  }
}

function readTable_(sheetName, requiredHeaders) {
  const sheet = getTrackerSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + ' sheet not found.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = requireHeaders_(headers, requiredHeaders, sheetName);
  const count = Math.max(0, sheet.getLastRow() - 1);
  return { sheet: sheet, headers: headers, map: map,
    rows: count ? sheet.getRange(2, 1, count, headers.length).getValues() : [] };
}

function getFrameworkRows_() {
  const table = readTable_(FRAMEWORK_SHEET_NAME,
    ['Category', 'Type', 'Definition', 'Activity_Examples', 'Active', 'Sort_Order']);
  const seen = {};
  return table.rows.filter(function(row) { return row[table.map.active] === true; })
    .map(function(row) {
      const category = requireNonBlankText_(row[table.map.category], 'Framework Category');
      const type = requireNonBlankText_(row[table.map.type], 'Framework Type');
      if (['Intervention', 'Escalation', 'Standardization', 'Scaling', 'Engagement'].indexOf(category) < 0) {
        throw new Error('Unsupported active category in Action_Framework: ' + category);
      }
      const key = registryFrameworkLabelKey_(category) + '\n' + registryFrameworkLabelKey_(type);
      if (seen[key]) throw new Error('Duplicate active Category / Type in Action_Framework (including equivalent hyphen/spacing variants): ' + category + ' / ' + type);
      seen[key] = true;
      return { category: category, type: type,
        definition: String(row[table.map.definition] || ''),
        examples: String(row[table.map.activity_examples] || ''),
        sortOrder: Number(row[table.map.sort_order]) || 999 };
    }).sort(function(a, b) { return a.sortOrder - b.sortOrder || a.type.localeCompare(b.type); });
}

function validateClassification_(category, type) {
  const match = getFrameworkRows_().find(function(row) {
    return row.category === category && row.type === type;
  });
  if (!match) throw new Error('Select an active, allowed Category / Type combination. Refresh the form if the framework changed.');
  return match;
}

function boundedText_(value, label, maxLength, required) {
  const text = String(value == null ? '' : value).trim();
  if (required && !text) throw new Error(label + ' is required.');
  if (text.length > maxLength) throw new Error(label + ' must be at most ' + maxLength + ' characters.');
  return text;
}

function storedVersion_(value) {
  if (value === '' || value === null || value === undefined) return 1;
  const version = Number(value);
  if (!Number.isSafeInteger(version) || version < 1) throw new Error('Invalid stored version. Ask the maintainer to review this record.');
  return version;
}

function requireVersion_(expected, actual) {
  if (!Number.isSafeInteger(Number(expected)) || Number(expected) < 1 || Number(expected) !== actual) {
    throw registrySecurityError_('This action changed since you opened it. Reload its details before saving.', 'STALE_VERSION');
  }
}

function findActionRecord_(actionPlanId) {
  const table = readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS));
  assertMasterFormulaSchema_(table.headers);
  const indices = [];
  table.rows.forEach(function(row, index) {
    if (String(row[table.map['action plan id']] || '').trim() === actionPlanId) indices.push(index);
  });
  if (indices.length !== 1) throw new Error(indices.length ? 'Duplicate Action Plan ID. Writes blocked.' : 'Action Plan not found.');
  table.rowNumber = indices[0] + 2;
  table.row = table.rows[indices[0]];
  table.planVersion = storedVersion_(table.row[table.map.plan_version]);
  table.recordVersion = storedVersion_(table.row[table.map.record_version]);
  return table;
}

function requireOwner_(actor, action) {
  const owner = String(action.row[action.map['action owner']] || '').trim().toLowerCase();
  if (actor.email !== owner) throw registrySecurityError_('Only the active action owner may update execution or upload evidence.', 'OWNER_REQUIRED');
}

function digest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)
    .map(function(byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function normalizeRequestToken_(requestToken) {
  const token = String(requestToken == null ? '' : requestToken).trim();

  if (!token) {
    throw new Error('Request token is required. Refresh the app and try again.');
  }

  if (!/^[A-Za-z0-9._:-]{8,160}$/.test(token)) {
    throw new Error('Invalid request token.');
  }

  return token;
}

function getOperationPropertyKey_(kind, requestToken) {

  return 'ACTIONTRACKER_' + kind + '_' + digest_(requireAuthenticatedWriter_() + '\n' + requestToken);
}

function beginIdempotentOperation_(kind, requestToken, fingerprint) {
  const properties = PropertiesService.getScriptProperties();
  const key = getOperationPropertyKey_(kind, requestToken);
  const existingRaw = properties.getProperty(key);

  if (existingRaw) {
    let existing = null;

    try {
      existing = JSON.parse(existingRaw);
    } catch (error) {
      existing = null;
    }

    if (existing && existing.status === 'COMPLETE') {
      if (existing.fingerprint !== fingerprint) throw new Error('Request token was reused for different data. Reload before submitting.');
      return {
        duplicate: true,
        result: existing.result || null,
        key: key
      };
    }

    throw new Error(
      'This request is already being processed. Do not submit it again.'
    );
  }

  const entries = properties.getProperties();
  const usedBytes = Object.keys(entries).reduce(function(total, name) {
    return total + (name.length + String(entries[name]).length) * 2;
  }, 0);
  if (usedBytes > 400000) throw new Error('Request storage requires maintenance. No changes were saved. Contact the maintainer.');
  properties.setProperty(
    key,
    JSON.stringify({
      status: 'IN_PROGRESS',
      fingerprint: fingerprint,
      startedAt: new Date().toISOString()
    })
  );

  return {
    duplicate: false,
    result: null,
    key: key
  };
}

function completeIdempotentOperation_(operationKey, result) {
  const previous = JSON.parse(PropertiesService.getScriptProperties().getProperty(operationKey) || '{}');
  PropertiesService.getScriptProperties().setProperty(
    operationKey,
    JSON.stringify({
      status: 'COMPLETE',
      fingerprint: previous.fingerprint,
      completedAt: new Date().toISOString(),
      result: result
    })
  );
}

function abandonIdempotentOperation_(operationKey) {
  if (!operationKey) return;
  PropertiesService.getScriptProperties().deleteProperty(operationKey);
}

function markOperationNeedsReview_(operationKey, detail) {
  if (!operationKey) return;

  PropertiesService.getScriptProperties().setProperty(
    operationKey,
    JSON.stringify({
      status: 'FAILED_NEEDS_REVIEW',
      failedAt: new Date().toISOString(),
      detail: detail
    })
  );
}

function deterministicFormRequestToken_(sheet, row, timestamp) {
  const seed = [
    sheet.getSheetId(),
    row,
    timestamp instanceof Date ? timestamp.toISOString() : String(timestamp || '')
  ].join('|');

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    seed,
    Utilities.Charset.UTF_8
  );

  const hex = digest.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');

  return 'FORM-' + hex;
}

function findHeaderIndex_(headers, headerName) {

  const target =
    String(headerName)
      .trim()
      .toLowerCase();

  return headers.findIndex(
    header =>
      String(header)
        .trim()
        .toLowerCase() === target
  );
}

function getLastDataRowByColumn_(sheet, columnNumber) {

  const maxRows = sheet.getMaxRows();

  if (maxRows < 2) {
    return 1;
  }

  const values =
    sheet
      .getRange(
        2,
        columnNumber,
        maxRows - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = values.length - 1;
    i >= 0;
    i--
  ) {

    if (
      String(values[i][0]).trim() !== ''
    ) {
      return i + 2;
    }
  }

  return 1;
}

function getKpiSettings_() {

  const ss =
    getTrackerSpreadsheet_();

  const sheet =
    ss.getSheetByName(
      SETTINGS_SHEET_NAME
    );

  const settings = {};

  if (!sheet) {
    return settings;
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return settings;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .getDisplayValues();

  values.forEach(
    row => {

      const kpi =
        String(row[0] || '')
          .trim();

      if (!kpi) {
        return;
      }

      if (settings[kpi.toLowerCase()]) throw new Error('Duplicate KPI in Settings: ' + kpi);
      settings[
        kpi.toLowerCase()
      ] = {

        name:
          kpi,

        direction:
          String(row[1] || '')
            .trim(),

        unit:
          String(row[2] || '')
            .trim()

      };

    }
  );

  return settings;
}

function getRequiredMasterHeaders_() {
  return [
    'Action Plan ID',
    'Created Date',
    'Created By',
    'Site',
    'Action Plan Name',
    'Action Description',
    'Target KPI / Behavior',
    'Action Owner',
    'Action Category',
    'Baseline Performance',
    'Current Performance',
    'End Goal / Target',
    'Start Date',
    'Target End Date',
    'Action Status',
    'Days Remaining',
    'Performance Change',
    'Gap to Goal',
    'Goal Progress %',
    'Schedule Health',
    'Latest Comment',
    'Last Updated Date',
    'Last Updated By',
    'Completion Date'
  ];
}

function getRequiredSubmissionLogHeaders_() {
  return [
    'Submission ID',
    'Action Plan ID',
    'Submission Date',
    'Submission Source',
    'Submitted By',
    'Site',
    'Action Plan Name',
    'Action Description',
    'Target KPI / Behavior',
    'Action Owner',
    'Action Category',
    'Baseline Performance',
    'End Goal / Target',
    'Start Date',
    'Target End Date',
    'Action Status', 'Channel', 'LOB', 'Cohort', 'Cohort_Details',
    'Plan_Version', 'Request_ID', 'Revision_Comment', 'Plan_Snapshot'
  ];
}

function generateSubmissionId_(submissionSheet, submissionHeaders) {
  const map = requireHeaders_(
    submissionHeaders,
    ['Submission ID'],
    SUBMISSION_LOG_SHEET_NAME
  );

  const idColumn = map[normalizeHeaderKey_('Submission ID')] + 1;
  const lastRow = getLastDataRowByColumn_(submissionSheet, idColumn);
  let highestNumber = 0;

  if (lastRow >= 2) {
    submissionSheet
      .getRange(2, idColumn, lastRow - 1, 1)
      .getDisplayValues()
      .flat()
      .forEach(function(id) {
        const match = String(id).trim().match(/^SUB-(\d+)$/);
        if (!match) return;
        const number = Number(match[1]);
        if (number > highestNumber) highestNumber = number;
      });
  }

  return 'SUB-' + String(highestNumber + 1).padStart(5, '0');
}

function createActionPlan_(payload, context) {
  payload = payload || {};
  context = context || {};
  requireActiveActor_('canSubmit');

  const ss = context.spreadsheet || getTrackerSpreadsheet_();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const submissionSheet = ss.getSheetByName(SUBMISSION_LOG_SHEET_NAME);

  if (!masterSheet) {
    throw new Error('Action_Master sheet not found.');
  }

  if (!submissionSheet) {
    throw new Error('Submission_Log sheet not found.');
  }

  let site = normalizeActionTrackerSite_(
    requireNonBlankText_(payload.site, 'Site / Department')
  );
  const actionPlanName = boundedText_(payload.actionPlanName, 'Action Plan Name', 200, true);
  const actionDescription = boundedText_(payload.actionDescription, 'Action Steps', 5000, true);
  const measurement=registryMeasurementFields_(payload),governance=measurement.Measurement_Method==='Governance outcome';
  const kpi = governance?'':requireNonBlankText_(payload.kpi, 'Target KPI / Behavior');
  const owner = validateTaskUsOwnerEmail_(payload.owner);
  const category = requireNonBlankText_(payload.category, 'Action Category');
  const actionType = requireNonBlankText_(payload.actionType, 'Action Type');
  const activity = boundedText_(payload.activity, 'Activity', 2000, true);
  const department = boundedText_(payload.department, 'Department', 100, false);
  validateClassification_(category, actionType);
  const ownerEntry = getAccessEntry_(owner);
  if (!ownerEntry || !ownerEntry.active) throw new Error('Action Owner must be an active, listed user.');
  if (department && getSettingsOptionValues_('Department').indexOf(department) < 0) {
    throw new Error('Department is not configured in Settings.');
  }

  const creationOptions = getCreationOptions();
  const allowedSites = (creationOptions.sites || []).map(function(value) {
    return String(value).trim().toLowerCase();
  });
  const allowedCategories = (creationOptions.categories || []).map(function(value) {
    return String(value).trim().toLowerCase();
  });

  if (allowedSites.indexOf(site.toLowerCase()) === -1) {
    throw new Error('Site / Department is not configured in ActionTracker.');
  }

  site = creationOptions.sites.find(function(choice) { return registrySiteKey_(choice) === registrySiteKey_(site); });

  if (allowedCategories.indexOf(category.toLowerCase()) === -1) {
    throw new Error('Action Category is not configured in ActionTracker.');
  }

  const actionStatus = validateAllowedStatus_(payload.actionStatus);
  if (actionStatus !== 'Not Started') throw new Error('New plans start Not Started until an assigned approver approves the plan.');
  const scope = Object.assign(validateRegistryScope_(payload), validateRegistryGlide_(payload));
  const startDate = parseDateInputServer_(payload.startDate, 'Start Date');
  const endDate = parseDateInputServer_(payload.endDate, 'Target End Date');

  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('Target End Date cannot be earlier than Start Date.');
  }

  const kpiSettings = getKpiSettings_();
  const kpiKey = kpi.toLowerCase();
  const kpiSetting = kpiSettings[kpiKey];

  if (!governance && !kpiSetting) {
    throw new Error('KPI "' + kpi + '" is not configured in Settings.');
  }

  const baseline = governance?'':parseKpiMetricServer_(
    payload.baseline,
    kpi,
    kpiSetting,
    'Baseline Performance'
  );

  const goal = governance?'':parseKpiMetricServer_(
    payload.goal,
    kpi,
    kpiSetting,
    'End Goal / Target'
  );

  const source = requireNonBlankText_(
    context.source || 'ActionTracker App',
    'Submission Source'
  );

  const submittedBy = requireActiveActor_('canSubmit').email;

  const requestToken = normalizeRequestToken_(context.requestToken);
  const createdDate = context.createdDate instanceof Date
    ? context.createdDate
    : new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let operation = null;
  let masterRollback = [];
  let submissionLogRow = null;

  try {
    requireActiveActor_('canSubmit');
    validateClassification_(category, actionType);
    validateRegistryScope_(payload);
    validateRegistryGlide_(payload);
    readTable_('Review_Log', REGISTRY_REVIEW_HEADERS);
    const latestOwner = getAccessEntry_(owner);
    if (!latestOwner || !latestOwner.active) throw new Error('Action Owner is no longer active.');
    operation = beginIdempotentOperation_('CREATE', requestToken, digest_(JSON.stringify([
      site, actionPlanName, actionDescription, kpi, owner, category, actionType,
      activity, department, baseline, goal, startDate.toISOString(), endDate.toISOString(), actionStatus, scope, measurement
    ])));

    if (operation.duplicate) {
      return operation.result;
    }

    ensureRegistryGovernanceHeaders_();
    ensureRegistryGlideHeaders_();
    const masterHeaders = masterSheet
      .getRange(1, 1, 1, masterSheet.getLastColumn())
      .getDisplayValues()[0];

    const masterMap = requireHeaders_(
      masterHeaders,
      getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS),
      MASTER_SHEET_NAME
    );

    assertMasterFormulaSchema_(masterHeaders);

    ensureSubmissionAuditHeaders_(submissionSheet);
    const submissionHeaders = submissionSheet
      .getRange(1, 1, 1, submissionSheet.getLastColumn())
      .getDisplayValues()[0];

    const submissionMap = requireHeaders_(
      submissionHeaders,
      getRequiredSubmissionLogHeaders_(),
      SUBMISSION_LOG_SHEET_NAME
    );

    const actionPlanId = generateActionPlanId_(
      masterSheet,
      site,
      createdDate
    );

    const submissionId = generateSubmissionId_(
      submissionSheet,
      submissionHeaders
    );

    const idColumn = masterMap[normalizeHeaderKey_('Action Plan ID')] + 1;
    const newRow = getLastDataRowByColumn_(masterSheet, idColumn) + 1;

    function masterColumn(headerName) {
      return masterMap[normalizeHeaderKey_(headerName)] + 1;
    }

    function writeMasterField(headerName, value, literalText) {
      const column = masterColumn(headerName);

      if (column >= 16 && column <= 20) {
        throw new Error(
          'ActionTracker blocked a write into controlled formula column ' +
          column + ' (' + headerName + ').'
        );
      }

      const range = masterSheet.getRange(newRow, column);
      masterRollback.push({
        range: range,
        oldValue: range.getValue(),
        literalText: literalText
      });

      if (literalText) {
        setLiteralCellValue_(range, value);
      } else {
        range.setValue(value);
      }
    }

    writeMasterField('Action Plan ID', actionPlanId, true);
    writeMasterField('Created Date', createdDate, false);
    writeMasterField('Created By', submittedBy, true);
    writeMasterField('Site', site, true);
    writeMasterField('Action Plan Name', actionPlanName, true);
    writeMasterField('Action Description', actionDescription, true);
    writeMasterField('Target KPI / Behavior', kpi, true);
    writeMasterField('Action Owner', owner, true);

    writeMasterField('Action Category', '', true);
    writeMasterField('Department', department, true);
    writeMasterField('Framework_Category', category, true);
    writeMasterField('Action_Type', actionType, true);
    writeMasterField('Activity', activity, true);
    writeMasterField('Plan_Version', 1, false);
    writeMasterField('Record_Version', 1, false);
    writeMasterField('Review_Status', 'Pending Review', true);
    writeMasterField('Latest_Review_ID', '', true);
    writeMasterField('Effectiveness_Status', 'Not Yet Assessed', true);
    writeMasterField('Latest_Assessment_ID', '', true);
    writeMasterField('Approval_Required', true, false);
    Object.keys(measurement).forEach(function(header){writeMasterField(header,measurement[header],true);});
    Object.keys(scope).forEach(function(header) { writeMasterField(header, scope[header], true); });
    writeMasterField('Baseline Performance', baseline, false);
    writeMasterField('Current Performance', baseline, false);
    writeMasterField('End Goal / Target', goal, false);
    writeMasterField('Start Date', startDate, false);
    writeMasterField('Target End Date', endDate, false);
    writeMasterField('Action Status', actionStatus, true);
    writeMasterField('Latest Comment', '', true);
    writeMasterField('Last Updated Date', createdDate, false);
    writeMasterField('Last Updated By', submittedBy, true);

    if (actionStatus === 'Completed') {
      writeMasterField('Completion Date', createdDate, false);
    } else {
      writeMasterField('Completion Date', '', false);
    }

    SpreadsheetApp.flush();

    const logOutput = new Array(submissionHeaders.length).fill('');

    function setSubmissionField(headerName, value) {
      logOutput[submissionMap[normalizeHeaderKey_(headerName)]] = value;
    }

    setSubmissionField('Submission ID', protectLiteralText_(submissionId));
    setSubmissionField('Action Plan ID', protectLiteralText_(actionPlanId));
    setSubmissionField('Submission Date', createdDate);
    setSubmissionField('Submission Source', protectLiteralText_(source));
    setSubmissionField('Submitted By', protectLiteralText_(submittedBy));
    setSubmissionField('Site', protectLiteralText_(site));
    setSubmissionField('Action Plan Name', protectLiteralText_(actionPlanName));
    setSubmissionField('Action Description', protectLiteralText_(actionDescription));
    setSubmissionField('Target KPI / Behavior', protectLiteralText_(kpi));
    setSubmissionField('Action Owner', protectLiteralText_(owner));
    setSubmissionField('Action Category', protectLiteralText_(category));
    setSubmissionField('Baseline Performance', baseline);
    setSubmissionField('End Goal / Target', goal);
    setSubmissionField('Start Date', startDate);
    setSubmissionField('Target End Date', endDate);
    setSubmissionField('Action Status', protectLiteralText_(actionStatus));
    Object.keys(measurement).forEach(function(header){setSubmissionField(header,protectLiteralText_(measurement[header]));});
    Object.keys(scope).forEach(function(header) { setSubmissionField(header, protectLiteralText_(scope[header])); });
    setSubmissionField('Plan_Version', 1);
    setSubmissionField('Request_ID', protectLiteralText_(requestToken));
    setSubmissionField('Revision_Comment', 'Initial registration');
    const createdSnapshot = masterSheet.getRange(newRow, 1, 1, masterHeaders.length).getValues()[0];
    setSubmissionField('Plan_Snapshot', JSON.stringify(snapshotNamedRow_(masterHeaders, createdSnapshot)));

    submissionLogRow = submissionSheet.getLastRow() + 1;
    submissionSheet
      .getRange(submissionLogRow, 1, 1, logOutput.length)
      .setValues([logOutput]);

    SpreadsheetApp.flush();

    const result = {
      success: true,
      version: ACTIONTRACKER_VERSION,
      submissionId: submissionId,
      actionPlanId: actionPlanId,
      createdDate: createdDate.toISOString()
    };

    completeIdempotentOperation_(operation.key, result);
    masterRollback = [];
    submissionLogRow = null;

    return result;

  } catch (error) {
    let rollbackError = null;

    try {
      if (submissionLogRow) {
        submissionSheet
          .getRange(submissionLogRow, 1, 1, submissionSheet.getLastColumn())
          .clearContent();
      }

      masterRollback.slice().reverse().forEach(function(item) {
        if (item.literalText && typeof item.oldValue === 'string') {
          setLiteralCellValue_(item.range, item.oldValue);
        } else if (
          item.oldValue === '' ||
          item.oldValue === null ||
          item.oldValue === undefined
        ) {
          item.range.clearContent();
        } else {
          item.range.setValue(item.oldValue);
        }
      });

      SpreadsheetApp.flush();
    } catch (restoreError) {
      rollbackError = restoreError;
    }

    if (operation && operation.key) {
      if (rollbackError) {
        markOperationNeedsReview_(operation.key, {
          operation: 'CREATE',
          error: String(error && error.message ? error.message : error),
          rollbackError: String(
            rollbackError && rollbackError.message
              ? rollbackError.message
              : rollbackError
          )
        });
      } else {
        abandonIdempotentOperation_(operation.key);
      }
    }

    if (rollbackError) {
      throw new Error(
        'Creation failed and automatic rollback also failed. ' +
        'Do not retry until Action_Master and Submission_Log are reviewed. ' +
        'Original error: ' +
        (error && error.message ? error.message : String(error))
      );
    }

    throw error;

  } finally {
    lock.releaseLock();
  }
}

function handleFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error(
      'handleFormSubmit must run from the spreadsheet form-submit trigger.'
    );
  }

  const responseSheet = e.range.getSheet();

  if (responseSheet.getName() !== FORM_SHEET_NAME) {
    return;
  }

  throw new Error('Legacy Google Form intake is paused. Submit through the ActionTracker app. No Action_Master record was written.');
}

function submitNewActionCore_(payload, requestToken) {
  return createActionPlan_(
    payload,
    {
      spreadsheet: getTrackerSpreadsheet_(),
      source: 'ActionTracker App',
      createdDate: new Date(),
      requestToken: requestToken,
      requireAuthenticatedWriter: true
    }
  );
}

function getCreationOptions() {
  const framework = getFrameworkRows_();
  const settings = getKpiSettings_();
  const kpis = Object.keys(settings)
    .map(function(key) {
      return {
        name: settings[key].name || key,
        unit: settings[key].unit || '',
        direction: settings[key].direction || ''
      };
    })
    .sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });

  const sites = new Set(ACTIONTRACKER_DEFAULT_SITES);
  const categories = new Set(ACTIONTRACKER_ACTION_CATEGORIES);

  getSettingsOptionValues_('Site').forEach(function(site) {
    const normalized = normalizeActionTrackerSite_(site);
    if (normalized) sites.add(normalized);
  });

  getSettingsOptionValues_('Action Category').forEach(function(category) {
    if (category) categories.add(category);
  });

  const ss = getTrackerSpreadsheet_();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (masterSheet && masterSheet.getLastRow() >= 2) {
    const headers = masterSheet
      .getRange(1, 1, 1, masterSheet.getLastColumn())
      .getDisplayValues()[0];
    const map = buildHeaderMapStrict_(headers, MASTER_SHEET_NAME);
    const siteIndex = map[normalizeHeaderKey_('Site')];
    const categoryIndex = map[normalizeHeaderKey_('Action Category')];
    const idIndex = map[normalizeHeaderKey_('Action Plan ID')];

    if (
      siteIndex !== undefined &&
      categoryIndex !== undefined &&
      idIndex !== undefined
    ) {
      const idColumn = idIndex + 1;
      const lastRow = getLastDataRowByColumn_(masterSheet, idColumn);

      if (lastRow >= 2) {
        const rows = masterSheet
          .getRange(2, 1, lastRow - 1, masterSheet.getLastColumn())
          .getDisplayValues();

        rows.forEach(function(row) {
          const site = normalizeActionTrackerSite_(row[siteIndex]);
          const category = String(row[categoryIndex] || '').trim();
          if (site) sites.add(site);
          if (category) categories.add(category);
        });
      }
    }
  }

  const siteOrder = ACTIONTRACKER_DEFAULT_SITES.reduce(function(map, site, index) {
    map[site] = index;
    return map;
  }, {});

  const categoryOrder = ACTIONTRACKER_ACTION_CATEGORIES.reduce(
    function(map, category, index) {
      map[category] = index;
      return map;
    },
    {}
  );

  return {
    version: ACTIONTRACKER_VERSION,
    kpis: kpis,
    sites: Array.from(sites).sort(function(a, b) {
      const ai = Object.prototype.hasOwnProperty.call(siteOrder, a)
        ? siteOrder[a]
        : 999;
      const bi = Object.prototype.hasOwnProperty.call(siteOrder, b)
        ? siteOrder[b]
        : 999;
      return ai !== bi ? ai - bi : a.localeCompare(b);
    }),
    categories: Array.from(new Set(framework.map(function(row) { return row.category; }))),
    framework: framework,
    categoryIntents: REGISTRY_SPONSOR_INTENTS,
    departments: getSettingsOptionValues_('Department'),
    channels: getSettingsOptionValues_('Channel'),
    lobs: getSettingsOptionValues_('LOB'),
    cohorts: getSettingsOptionValues_('Cohort'),
    permissions: getAccessContext(),
    legacyCategories: Array.from(categories).sort(function(a, b) {
      const ai = Object.prototype.hasOwnProperty.call(categoryOrder, a)
        ? categoryOrder[a]
        : 999;
      const bi = Object.prototype.hasOwnProperty.call(categoryOrder, b)
        ? categoryOrder[b]
        : 999;
      return ai !== bi ? ai - bi : a.localeCompare(b);
    }),
    statuses: ['Not Started']
  };
}

function generateActionPlanId_(
  masterSheet,
  site,
  dateCreated
) {
  const headers = masterSheet
    .getRange(1, 1, 1, masterSheet.getLastColumn())
    .getDisplayValues()[0];
  const map = requireHeaders_(
    headers,
    ['Action Plan ID'],
    MASTER_SHEET_NAME
  );
  const idColumn = map[normalizeHeaderKey_('Action Plan ID')] + 1;

  const siteCode = String(site)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');

  const year = dateCreated.getFullYear();
  const prefix = siteCode + '-' + year + '-';
  const lastRow = getLastDataRowByColumn_(masterSheet, idColumn);
  let highestNumber = 0;

  if (lastRow >= 2) {
    masterSheet
      .getRange(2, idColumn, lastRow - 1, 1)
      .getDisplayValues()
      .flat()
      .forEach(function(id) {
        const text = String(id).trim();
        const prefixes = siteCode === 'MTQ' ? [prefix, 'MYQ-' + year + '-'] : [prefix];
        const matchedPrefix = prefixes.find(function(p) { return text.startsWith(p); });
        if (!matchedPrefix) return;
        const numberPart = Number(text.slice(matchedPrefix.length));
        if (Number.isInteger(numberPart) && numberPart > highestNumber) {
          highestNumber = numberPart;
        }
      });
  }

  return prefix + String(highestNumber + 1).padStart(4, '0');
}

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile(
      'Index'
    )
    .setTitle(
      'Shopify Global Accounts Action Registry'
    );
}

function registryReportingContext_() {
  const result = { reviewers: [], submissions: {}, reviews: {}, assessments: {}, assessmentDetails: {}, reviewHistoryLoaded: false, assessmentHistoryLoaded: false, warnings: [] };

  try {
    const table = readTable_(ACCESS_SHEET_NAME, ['Email', 'Name', 'Active', 'Can_Approve', 'Approval_Sites']);
    table.rows.forEach(function(row) {
      if (row[table.map.active] !== true || row[table.map.can_approve] !== true) return;
      const email = String(row[table.map.email] || '').trim().toLowerCase();
      if (!/^[A-Z0-9._%+\-]+@taskus\.com$/i.test(email)) return;
      result.reviewers.push({ email: email, name: String(row[table.map.name] || email),
        approvalSites: parseApprovalSites_(row[table.map.approval_sites]), canApprove: true });
    });
  } catch (e) { result.warnings.push('Reviewer assignments could not be loaded.'); }
  try {
    const table = readTable_(SUBMISSION_LOG_SHEET_NAME, ['Action Plan ID', 'Submission Date']);
    table.rows.forEach(function(row) {
      const id = String(row[table.map['action plan id']] || '').trim();
      const version = table.map.plan_version === undefined ? 1 : Number(row[table.map.plan_version]) || 1;
      const date = dateToClientIso_(row[table.map['submission date']]);
      if (id && date) result.submissions[id + ':' + version] = date;
    });
  } catch (e) { result.warnings.push('Submission dates could not be loaded; revised-plan waiting times may be unavailable.'); }
  try {
    const table = readTable_('Review_Log', REGISTRY_REVIEW_HEADERS), seen = {};
    table.rows.forEach(function(row) {
      const id = String(row[table.map.review_id] || '');
      if (!id) return;
      if (seen[id]) { result.reviews[id] = null; result.warnings.push('A duplicate review ID was excluded from reviewer activity.'); return; }
      seen[id] = true;
      result.reviews[id] = { actionPlanId: String(row[table.map.action_plan_id]),
        planVersion: Number(row[table.map.plan_version]), recordVersionReviewed: Number(row[table.map.record_version_reviewed]), decision: String(row[table.map.decision]),
        email: String(row[table.map.reviewed_by]), name: String(row[table.map.reviewer_name] || row[table.map.reviewed_by]),
        reviewedAt: dateToClientIso_(row[table.map.reviewed_at]) };
    });
    result.reviewHistoryLoaded = true;
  } catch (e) { result.warnings.push('Review history could not be loaded; reviewer activity may be unavailable.'); }
  try {
    if (getTrackerSpreadsheet_().getSheetByName(REGISTRY_ASSESSMENT_SHEET)) {
      const table = readTable_(REGISTRY_ASSESSMENT_SHEET, ['Action_Plan_ID', 'Plan_Version', 'Assessment_ID']);
      const seen=Object.create(null),evidenceIndex=registryStabilityEvidenceIndex_();
      table.rows.forEach(function(row) {
        const id=String(row[table.map.assessment_id] || '').trim();if(!id)return;
        result.assessments[String(row[table.map.action_plan_id]).trim()+':'+Number(row[table.map.plan_version])] = true;
        if(seen[id]){result.assessmentDetails[id]=null;result.warnings.push('Duplicate assessment pointer excluded from stability evidence.');return;}seen[id]=true;
        const period=row[table.map.period_end];
        result.assessmentDetails[id]={actionPlanId:String(row[table.map.action_plan_id] || '').trim(),planVersion:Number(row[table.map.plan_version]),recordVersionReviewed:Number(row[table.map.record_version_reviewed]),
          status:String(row[table.map.effectiveness_status] || ''),periodEnd:period instanceof Date?Utilities.formatDate(period,'Asia/Manila','yyyy-MM-dd'):String(period || '').trim(),
          sustained:String(row[table.map.sustained_result] || ''),expectations:String(row[table.map.expectations_result] || ''),relationship:String(row[table.map.relationship] || ''),
          hasEvidence:String(row[table.map.evidence_ids] || '').split(/\r?\n/).some(function(evidenceId){const item=evidenceIndex[evidenceId.trim()];return !!item && item.actionPlanId===String(row[table.map.action_plan_id] || '').trim() && item.planVersion===Number(row[table.map.plan_version]);}) || String(row[table.map.evidence_urls] || '').split(/\r?\n/).some(function(url){return isRegistryGlideUrl_(url.trim());})};
      });
    }
    result.assessmentHistoryLoaded = true;
  } catch (e) { result.assessmentDetails = {}; result.warnings.push('Assessment history unavailable; first-assessment queue cannot be confirmed.'); }
  return result;
}
function registryEligibleReviewers_(context, site) {
  const seen = {};
  return context.reviewers.filter(function(actor) {
    if (seen[actor.email] || !registryActorApprovesSite_(actor, site)) return false;
    seen[actor.email] = true; return true;
  }).map(function(actor) { return { name: actor.name, email: actor.email }; });
}

function getDashboardData() {
  const ss = getTrackerSpreadsheet_();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) {
    throw new Error('Action_Master sheet not found.');
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const headerMap = requireHeaders_(
    headers,
    getRequiredMasterHeaders_(),
    MASTER_SHEET_NAME
  );

  assertMasterFormulaSchema_(headers);

  const idColumn = headerMap[normalizeHeaderKey_('Action Plan ID')] + 1;
  const lastRow = getLastDataRowByColumn_(sheet, idColumn);

  if (lastRow < 2) {
    return [];
  }

  const range = sheet.getRange(
    2,
    1,
    lastRow - 1,
    sheet.getLastColumn()
  );

  const displayData = range.getDisplayValues();
  const rawData = range.getValues();
  const kpiSettings = getKpiSettings_();
  const reporting = registryReportingContext_();
  const dashboardIdCounts=Object.create(null);
  displayData.forEach(function(row){const id=String(row[headerMap['action plan id']] || '').trim();if(id)dashboardIdCounts[id]=(dashboardIdCounts[id] || 0)+1;});

  function displayValue(row, headerName) {
    const index = headerMap[normalizeHeaderKey_(headerName)];
    return index === undefined ? '' : row[index];
  }

  function rawValue(row, headerName) {
    const index = headerMap[normalizeHeaderKey_(headerName)];
    return index === undefined ? '' : row[index];
  }

  return displayData
    .map(function(row, rowIndex) {
      const rawRow = rawData[rowIndex];
      const id = displayValue(row, 'Action Plan ID');

      if (!String(id || '').trim()) {
        return null;
      }

      const kpi = displayValue(row, 'Target KPI / Behavior');
      const setting =
        resolveRegistryKpiSetting_(kpiSettings, kpi) || {
          direction: '',
          unit: ''
        };

      return registryGovernanceRead_({
        id: id,
        site: normalizeActionTrackerSite_(displayValue(row, 'Site')),
        name: displayValue(row, 'Action Plan Name'),
        description: displayValue(row, 'Action Description'),
        kpi: kpi,
        unit: setting.unit,
        direction: setting.direction,
        owner: displayValue(row, 'Action Owner'),
        category: displayValue(row, 'Framework_Category') || displayValue(row, 'Action Category'),
        legacyCategory: displayValue(row, 'Action Category'),
        frameworkCategory: displayValue(row, 'Framework_Category'),
        actionType: displayValue(row, 'Action_Type'),
        activity: displayValue(row, 'Activity'),
        department: displayValue(row, 'Department'),
        channel: displayValue(row, 'Channel'), lob: displayValue(row, 'LOB'),
        cohort: displayValue(row, 'Cohort'), cohortDetails: displayValue(row, 'Cohort_Details'),
        glideAvailable: displayValue(row, 'Glide_Path_Available') || 'Not Recorded',
        glideUrl: displayValue(row, 'Glide_Path_URL'),
        approvalRequired: rawValue(rawRow, 'Approval_Required') === true,
        planVersion: Number(rawValue(rawRow, 'Plan_Version')) || 1,
        eligibleReviewers: registryEligibleReviewers_(reporting, displayValue(row, 'Site')),
        reviewSubmittedAtIso: reporting.submissions[String(id).trim() + ':' + (Number(rawValue(rawRow, 'Plan_Version')) || 1)] ||
          (Number(rawValue(rawRow, 'Plan_Version')) <= 1 ? dateToClientIso_(rawValue(rawRow, 'Created Date')) : ''),
        latestReview: reporting.reviews[String(rawValue(rawRow, 'Latest_Review_ID') || '')] || null,
        reportingWarnings: reporting.warnings,
        reviewStatus: displayValue(row, 'Review_Status') || 'Not Recorded',
        effectivenessStatus: displayValue(row, 'Effectiveness_Status') || 'Not Yet Assessed',
        awaitingFirstAssessment: registryAwaitingAssessment_(rawRow, headerMap, reporting),
        assessmentHistoryLoaded: reporting.assessmentHistoryLoaded,
        stability: dashboardIdCounts[String(id).trim()]>1?{approved:false,assessment:null,issue:'Duplicate Action ID'}:registryStabilityFacts_(rawRow, headerMap, reporting),
        baseline: parseStoredMetricNumber_(rawValue(rawRow, 'Baseline Performance'), setting.unit),
        current: parseStoredMetricNumber_(rawValue(rawRow, 'Current Performance'), setting.unit),
        goal: parseStoredMetricNumber_(rawValue(rawRow, 'End Goal / Target'), setting.unit),
        outcome: deriveOutcomeStatus_(
          rawValue(rawRow, 'Baseline Performance'),
          rawValue(rawRow, 'Current Performance'),
          rawValue(rawRow, 'End Goal / Target'),
          setting.direction,
          setting.unit
        ),
        status: displayValue(row, 'Action Status'),
        health: displayValue(row, 'Schedule Health'),
        daysRemaining: displayValue(row, 'Days Remaining'),
        progress: displayValue(row, 'Goal Progress %'),
        startDate: displayValue(row, 'Start Date'),
        endDate: displayValue(row, 'Target End Date'),
        createdDateIso: dateToClientIso_(rawValue(rawRow, 'Created Date')),
        lastUpdatedDateIso: dateToClientIso_(rawValue(rawRow, 'Last Updated Date')),
        completionDateIso: dateToClientIso_(rawValue(rawRow, 'Completion Date')),
        endDateIso: dateToClientIso_(rawValue(rawRow, 'Target End Date'))
      },rawRow,headerMap);
    })
    .filter(function(row) {
      return row !== null;
    });
}

function getActionPlans() {

  const rows =
    getDashboardData();

  return rows.map(
    row => ({

      id:
        row.id,

      site:
        row.site,

      name:
        row.name,

      status:
        row.status,

      currentPerformance:
        row.current

    })
  );
}

function getActionPlanDetails(actionPlanId) {
  const ss = getTrackerSpreadsheet_();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) {
    throw new Error('Action_Master sheet not found.');
  }

  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const headerMap = requireHeaders_(
    headers,
    getRequiredMasterHeaders_(),
    MASTER_SHEET_NAME
  );

  assertMasterFormulaSchema_(headers);

  const idColumn = headerMap[normalizeHeaderKey_('Action Plan ID')] + 1;
  const lastRow = getLastDataRowByColumn_(sheet, idColumn);

  if (lastRow < 2) {
    throw new Error('No actions found.');
  }

  const range = sheet.getRange(
    2,
    1,
    lastRow - 1,
    sheet.getLastColumn()
  );

  const displayData = range.getDisplayValues();
  const rawData = range.getValues();
  const idIndex = headerMap[normalizeHeaderKey_('Action Plan ID')];

  const rowIndex = displayData.findIndex(function(row) {
    return String(row[idIndex] || '').trim() === actionPlanId;
  });

  if (rowIndex === -1) {
    throw new Error('Action Plan not found: ' + actionPlanId);
  }

  const row = displayData[rowIndex];
  const rawRow = rawData[rowIndex];

  function value(headerName) {
    const index = headerMap[normalizeHeaderKey_(headerName)];
    return index === undefined ? '' : row[index];
  }

  function rawValue(headerName) {
    const index = headerMap[normalizeHeaderKey_(headerName)];
    return index === undefined ? '' : rawRow[index];
  }

  const kpi = value('Target KPI / Behavior');
  const kpiSettings = getKpiSettings_();
  const setting =
    resolveRegistryKpiSetting_(kpiSettings, kpi) || {
      unit: '',
      direction: ''
    };

  const permissions = getAccessContext();
  permissions.canUpdate = permissions.verified && permissions.email === String(value('Action Owner') || '').trim().toLowerCase();
  const approvalAction = findActionRecord_(actionPlanId);
  const approval = getRegistryApprovalState_(approvalAction);
  const isOwner = permissions.canUpdate;
  const executionLocked = registryPlanCompleted_(approvalAction);
  permissions.canUpdate = isOwner && approval.executionAllowed && !executionLocked;
  permissions.canUpload = isOwner && approval.executionAllowed && !executionLocked;
  permissions.canReview = permissions.verified && permissions.canApprove &&
    registryActorApprovesSite_(permissions, value('Site')) &&
    approval.required && value('Review_Status') === 'Pending Review' && !executionLocked;
  permissions.canRevise = isOwner && approval.required &&
    ['Pending Review', 'Returned for Revision', 'Approved'].indexOf(value('Review_Status')) >= 0 &&
    ['Completed', 'Cancelled'].indexOf(value('Action Status')) < 0;
  const evidence = getActionEvidence_(actionPlanId);
  const planVersion = storedVersion_(rawValue('Plan_Version'));
  const assessmentEligibility = getRegistryAssessmentEligibility_(approvalAction, approval);
  permissions.canAssess = permissions.verified && permissions.active && permissions.canAudit && assessmentEligibility.allowed;
  const assessments = getRegistryAssessments_(actionPlanId);

  return registryGovernanceRead_({
    actionPlanId: value('Action Plan ID'),
    site: normalizeActionTrackerSite_(value('Site')),
    actionPlanName: value('Action Plan Name'),
    actionDescription: value('Action Description'),
    kpi: kpi,
    kpiUnit: setting.unit,
    kpiDirection: setting.direction,
    owner: value('Action Owner'),
    category: value('Framework_Category') || value('Action Category'),
    legacyCategory: value('Action Category'),
    frameworkCategory: value('Framework_Category'),
    actionType: value('Action_Type'), activity: value('Activity'),
    department: value('Department'),
    channel: value('Channel'), lob: value('LOB'), cohort: value('Cohort'),
    cohortDetails: value('Cohort_Details'),
    glideAvailable: value('Glide_Path_Available') || 'Not Recorded', glideUrl: value('Glide_Path_URL'),
    executionLocked: executionLocked,
    approvalRequired: approval.required, executionAllowed: approval.executionAllowed,
    approvalMessage: approval.message, reviews: getRegistryReviews_(actionPlanId),
    startDateInput: registryDateInput_(rawValue('Start Date')),
    endDateInput: registryDateInput_(rawValue('Target End Date')),
    baselineRaw: rawValue('Baseline Performance'), goalRaw: rawValue('End Goal / Target'),
    planVersion: planVersion,
    recordVersion: storedVersion_(rawValue('Record_Version')),
    reviewStatus: value('Review_Status') || 'Not Recorded',
    effectivenessStatus: value('Effectiveness_Status') || 'Not Yet Assessed',
    permissions: permissions,
    evidence: evidence,
    assessments: assessments,
    assessmentDefaults: registryAssessmentDefaults_(approvalAction, assessments),
    assessmentMessage: !assessmentEligibility.allowed ? assessmentEligibility.message :
      (!permissions.canAssess ? 'Assessment requires an active account with Can_Audit checked.' : 'Assessment permitted for this plan version.'),
    comments: getActionComments_(actionPlanId),
    hasCurrentEvidence: evidence.some(function(item) { return item.planVersion === planVersion; }),
    uploadConfigured: !!PropertiesService.getScriptProperties().getProperty('EVIDENCE_FOLDER_ID'),
    baseline: parseStoredMetricNumber_(rawValue('Baseline Performance'), setting.unit),
    current: parseStoredMetricNumber_(rawValue('Current Performance'), setting.unit),
    goal: parseStoredMetricNumber_(rawValue('End Goal / Target'), setting.unit),
    outcome: deriveOutcomeStatus_(
      rawValue('Baseline Performance'),
      rawValue('Current Performance'),
      rawValue('End Goal / Target'),
      setting.direction,
      setting.unit
    ),
    startDate: value('Start Date'),
    endDate: value('Target End Date'),
    actionStatus: value('Action Status'),
    daysRemaining: value('Days Remaining'),
    performanceChange: registryMetricDelta_(rawValue('Baseline Performance'), rawValue('Current Performance'), rawValue('End Goal / Target'), setting.direction, setting.unit, false),
    gapToGoal: registryMetricDelta_(rawValue('Baseline Performance'), rawValue('Current Performance'), rawValue('End Goal / Target'), setting.direction, setting.unit, true),
    goalProgress: value('Goal Progress %'),
    scheduleHealth: value('Schedule Health'),
    latestComment: value('Latest Comment'),
    lastUpdatedDate: value('Last Updated Date'),
    lastUpdatedBy: value('Last Updated By'),
    completionDate: value('Completion Date')
  },rawRow,headerMap);
}

function submitActionUpdateCore_(
  actionPlanId,
  currentPerformance,
  actionStatus,
  comment,
  requestToken,
  expectedRecordVersion
) {
  const ss = getTrackerSpreadsheet_();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const logSheet = ss.getSheetByName(UPDATE_LOG_SHEET_NAME);

  if (!masterSheet) {
    throw new Error('Action_Master sheet not found.');
  }

  if (!logSheet) {
    throw new Error('Update_Log sheet not found.');
  }

  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');
  actionStatus = validateAllowedStatus_(actionStatus);
  comment = boundedText_(comment, 'Execution comment', 5000, false);
  requestToken = normalizeRequestToken_(requestToken);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let operation = null;
  let rollbackState = [];
  let logRow = null;

  try {
    const actor = requireActiveActor_();
    const action = findActionRecord_(actionPlanId);
    requireOwner_(actor, action);
    operation = beginIdempotentOperation_('UPDATE', requestToken, digest_(JSON.stringify([
      actionPlanId, currentPerformance, actionStatus, comment, expectedRecordVersion
    ])));

    if (operation.duplicate) {
      return operation.result;
    }
    requireRegistryPlanEditable_(action);
    requireVersion_(expectedRecordVersion, action.recordVersion);
    requireRegistryExecutionApproved_(action);
    const governance=registryMeasurementMethod_(action.row[action.map.measurement_method])==='Governance outcome';
    if(governance && String(currentPerformance == null?'':currentPerformance).trim())throw new Error('Governance updates do not accept numeric performance.');
    const oldStatus = String(action.row[action.map['action status']] || '').trim();
    if (actionStatus === 'Completed' && oldStatus !== 'Completed') {
      requireExecutionEvidence_(actionPlanId, action.planVersion);
    }

    const now = new Date();
    const updatedBy = requireAuthenticatedWriter_();

    ensureRegistryGovernanceHeaders_();
    const masterHeaders = masterSheet
      .getRange(1, 1, 1, masterSheet.getLastColumn())
      .getDisplayValues()[0];

    const masterMap = requireHeaders_(
      masterHeaders,
      [
        'Action Plan ID', 'Action Owner', 'Record_Version', 'Plan_Version',
        'Target KPI / Behavior',
        'Current Performance',
        'Action Status',
        'Latest Comment',
        'Last Updated Date',
        'Last Updated By',
        'Completion Date',
        'Days Remaining',
        'Performance Change',
        'Gap to Goal',
        'Goal Progress %',
        'Schedule Health'
      ],
      MASTER_SHEET_NAME
    );

    assertMasterFormulaSchema_(masterHeaders);

    const logHeaders = logSheet
      .getRange(1, 1, 1, logSheet.getLastColumn())
      .getDisplayValues()[0];

    const logMap = requireHeaders_(
      logHeaders,
      [
        'Update ID',
        'Action Plan ID',
        'Update Date',
        'Updated By',
        'Current Performance',
        'Action Status',
        'Comment'
      ],
      UPDATE_LOG_SHEET_NAME
    );

    function masterColumn(headerName) {
      return masterMap[normalizeHeaderKey_(headerName)] + 1;
    }

    const idColumn = masterColumn('Action Plan ID');
    const lastMasterRow = getLastDataRowByColumn_(masterSheet, idColumn);

    if (lastMasterRow < 2) {
      throw new Error('No actions found.');
    }

    const actionIds = masterSheet
      .getRange(2, idColumn, lastMasterRow - 1, 1)
      .getDisplayValues()
      .flat();
    if (actionIds.filter(function(id) { return String(id).trim() === actionPlanId; }).length > 1) {
      throw new Error('Duplicate Action Plan ID. Writes blocked.');
    }

    const relativeIndex = actionIds.findIndex(function(id) {
      return String(id).trim() === actionPlanId;
    });

    if (relativeIndex === -1) {
      throw new Error('Action Plan ID not found: ' + actionPlanId);
    }

    const masterRow = relativeIndex + 2;
    const kpiName = masterSheet
      .getRange(masterRow, masterColumn('Target KPI / Behavior'))
      .getDisplayValue();
    const kpiSettings = getKpiSettings_();
    const kpiSetting = resolveRegistryKpiSetting_(kpiSettings, kpiName);

    if (!governance && !kpiSetting) {
      throw new Error('KPI "' + kpiName + '" is not configured in Settings.');
    }

    const numericPerformance = governance?'':validateKpiPerformanceServer_(
      currentPerformance,
      kpiName,
      kpiSetting
    );

    const trackedFields = [
      { header: 'Current Performance', literal: false },
      { header: 'Action Status', literal: true },
      { header: 'Latest Comment', literal: true },
      { header: 'Last Updated Date', literal: false },
      { header: 'Last Updated By', literal: true },
      { header: 'Completion Date', literal: false },
      { header: 'Record_Version', literal: false }
    ];

    rollbackState = trackedFields.map(function(item) {
      const range = masterSheet.getRange(
        masterRow,
        masterColumn(item.header)
      );
      return {
        range: range,
        oldValue: range.getValue(),
        literal: item.literal
      };
    });

    function setMasterValue(headerName, value, literalText) {
      const range = masterSheet.getRange(
        masterRow,
        masterColumn(headerName)
      );
      if (literalText) {
        setLiteralCellValue_(range, value);
      } else {
        range.setValue(value);
      }
    }

    setMasterValue('Current Performance', numericPerformance, false);
    setMasterValue('Action Status', actionStatus, true);
    setMasterValue('Latest Comment', comment, true);
    setMasterValue('Last Updated Date', now, false);
    setMasterValue('Last Updated By', updatedBy, true);
    setMasterValue('Record_Version', action.recordVersion + 1, false);

    const completionCell = masterSheet.getRange(
      masterRow,
      masterColumn('Completion Date')
    );

    if (actionStatus === 'Completed') {
      if (!completionCell.getValue()) {
        completionCell.setValue(now);
      }
    } else {
      completionCell.clearContent();
    }

    SpreadsheetApp.flush();

    const updateId = generateUpdateId_(logSheet);
    const logOutput = new Array(logHeaders.length).fill('');

    function setLogField(headerName, value) {
      logOutput[logMap[normalizeHeaderKey_(headerName)]] = value;
    }

    setLogField('Update ID', protectLiteralText_(updateId));
    setLogField('Action Plan ID', protectLiteralText_(actionPlanId));
    setLogField('Update Date', now);
    setLogField('Updated By', protectLiteralText_(updatedBy));
    setLogField('Current Performance', numericPerformance);
    setLogField('Action Status', protectLiteralText_(actionStatus));
    setLogField('Comment', protectLiteralText_(comment));
    setLogField('Measurement_Method',protectLiteralText_(governance?'Governance outcome':'Numeric KPI'));
    setLogField('Plan_Version',action.planVersion);

    logRow = logSheet.getLastRow() + 1;
    logSheet
      .getRange(logRow, 1, 1, logOutput.length)
      .setValues([logOutput]);

    SpreadsheetApp.flush();

    const result = {
      success: true,
      version: ACTIONTRACKER_VERSION,
      updateId: updateId,
      actionPlanId: actionPlanId
    };

    completeIdempotentOperation_(operation.key, result);
    rollbackState = [];
    logRow = null;

    return result;

  } catch (error) {
    let rollbackError = null;

    try {
      if (logRow) {
        logSheet
          .getRange(logRow, 1, 1, logSheet.getLastColumn())
          .clearContent();
      }

      rollbackState.slice().reverse().forEach(function(item) {
        if (
          item.oldValue === '' ||
          item.oldValue === null ||
          item.oldValue === undefined
        ) {
          item.range.clearContent();
        } else if (item.literal && typeof item.oldValue === 'string') {
          setLiteralCellValue_(item.range, item.oldValue);
        } else {
          item.range.setValue(item.oldValue);
        }
      });

      SpreadsheetApp.flush();
    } catch (restoreError) {
      rollbackError = restoreError;
    }

    if (operation && operation.key) {
      if (rollbackError) {
        markOperationNeedsReview_(operation.key, {
          operation: 'UPDATE',
          actionPlanId: actionPlanId,
          error: String(error && error.message ? error.message : error),
          rollbackError: String(
            rollbackError && rollbackError.message
              ? rollbackError.message
              : rollbackError
          )
        });
      } else {
        abandonIdempotentOperation_(operation.key);
      }
    }

    if (rollbackError) {
      throw new Error(
        'Update failed and automatic rollback also failed. ' +
        'Do not retry until Action_Master and Update_Log are reviewed. ' +
        'Original error: ' +
        (error && error.message ? error.message : String(error))
      );
    }

    throw error;

  } finally {
    lock.releaseLock();
  }
}

function getActionComments_(actionPlanId) {
  const table = readTable_(COMMENT_LOG_SHEET_NAME, ACTIONTRACKER_COMMENT_HEADERS);
  return table.rows.filter(function(row) {
    return String(row[table.map.action_plan_id] || '').trim() === actionPlanId;
  }).map(function(row) {
    return { commentId: String(row[table.map.comment_id] || ''),
      email: String(row[table.map.commenter_email] || ''),
      name: String(row[table.map.commenter_name] || ''),
      commentedAt: dateToClientIso_(row[table.map.commented_at]),
      text: String(row[table.map.comment_text] || '') };
  }).reverse();
}

function getActionEvidence_(actionPlanId) {
  const table = readTable_(EVIDENCE_LOG_SHEET_NAME, ACTIONTRACKER_EVIDENCE_HEADERS);
  return table.rows.filter(function(row) {
    return String(row[table.map.action_plan_id] || '').trim() === actionPlanId;
  }).map(function(row) {
    const fileId = String(row[table.map.file_id] || '').trim();
    return { evidenceId: String(row[table.map.evidence_id] || ''),
      planVersion: storedVersion_(row[table.map.plan_version]), fileId: fileId,
      fileName: String(row[table.map.file_name] || ''),

      fileUrl: /^[A-Za-z0-9_-]+$/.test(fileId) ? 'https://drive.google.com/file/d/' + fileId + '/view' : '',
      description: String(row[table.map.evidence_description] || ''),
      uploadedBy: String(row[table.map.uploaded_by] || ''),
      uploadedAt: dateToClientIso_(row[table.map.uploaded_at]) };
  }).reverse();
}

function getEvidenceFolder_() {
  const id = String(PropertiesService.getScriptProperties().getProperty('EVIDENCE_FOLDER_ID') || '').trim();
  if (!/^[A-Za-z0-9_-]{10,}$/.test(id)) throw new Error('Evidence uploads are not configured. The maintainer must set EVIDENCE_FOLDER_ID.');
  return DriveApp.getFolderById(id);
}

function requireExecutionEvidence_(actionPlanId, planVersion) {
  const evidence = getActionEvidence_(actionPlanId).filter(function(item) { return item.planVersion === planVersion; });
  if (!evidence.length) throw new Error('Upload proof of execution for this plan before marking it Completed.');
  const folderId = getEvidenceFolder_().getId();
  const present = evidence.some(function(item) {
    try {
      const file = DriveApp.getFileById(item.fileId);
      if (file.isTrashed()) return false;
      const parents = file.getParents();
      while (parents.hasNext()) if (parents.next().getId() === folderId) return true;
    } catch (error) { return false; }
    return false;
  });
  if (!present) throw new Error('Execution proof is missing or unavailable in the evidence folder. Upload an accessible artifact before completing this plan.');

}

function appendNamedLog_(table, fields) {
  const output = new Array(table.headers.length).fill('');
  Object.keys(fields).forEach(function(header) {
    const value = fields[header];
    output[table.map[normalizeHeaderKey_(header)]] = typeof value === 'string' ? protectLiteralText_(value) : value;
  });
  const row = table.sheet.getLastRow() + 1;
  table.sheet.getRange(row, 1, 1, output.length).setValues([output]);
  return row;
}

function submitLeadershipCommentCore_(actionPlanId, commentText, requestToken) {
  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');
  const text = boundedText_(commentText, 'Leadership comment', 5000, true);
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let operation = null;
  let table = null;
  let logRow = null;
  try {
    const actor = requireActiveActor_('canComment');
    findActionRecord_(actionPlanId);
    table = readTable_(COMMENT_LOG_SHEET_NAME, ACTIONTRACKER_COMMENT_HEADERS);
    operation = beginIdempotentOperation_('COMMENT', requestToken, digest_(JSON.stringify([actionPlanId, text])));
    if (operation.duplicate) return operation.result;
    const id = 'COM-' + Utilities.getUuid();
    const now = new Date();

    logRow = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, { Comment_ID: id, Action_Plan_ID: actionPlanId,
      Commenter_Email: actor.email, Commenter_Name: actor.name,
      Commented_At: now, Comment_Text: text, Request_ID: requestToken });
    SpreadsheetApp.flush();
    const result = { success: true, commentId: id, actionPlanId: actionPlanId };
    completeIdempotentOperation_(operation.key, result);
    logRow = null;
    return result;
  } catch (error) {
    try {
      if (logRow) table.sheet.getRange(logRow, 1, 1, table.headers.length).clearContent();
      SpreadsheetApp.flush();
      if (operation && !operation.duplicate) abandonIdempotentOperation_(operation.key);
    } catch (rollbackError) {
      if (operation) markOperationNeedsReview_(operation.key, { error: error.message, rollbackError: rollbackError.message });
      throw new Error('Comment save/rollback failed. Do not retry until Comment_Log is reviewed.');
    }
    throw error;
  } finally { lock.releaseLock(); }
}

function uploadActionEvidenceCore_(formObject) {

  const form = formObject || {};
  const actionPlanId = requireNonBlankText_(form.actionPlanId, 'Action Plan ID');
  const description = boundedText_(form.evidenceDescription, 'Evidence description', 2000, true);
  const token = normalizeRequestToken_(form.requestToken);
  requireActiveActor_();
  const blob = form.evidenceFile;
  if (!blob || typeof blob.getBytes !== 'function') throw new Error('Select a file to upload.');
  const bytes = blob.getBytes();
  if (!bytes.length || bytes.length > ACTIONTRACKER_MAX_UPLOAD_BYTES) throw new Error('File must be non-empty and at most 10 MB.');
  const name = boundedText_(blob.getName(), 'File name', 200, true).replace(/[\\/\u0000-\u001f]/g, '_');
  const extension = name.split('.').pop().toLowerCase();
  const allowed = {
    pdf: ['application/pdf'], png: ['image/png'], jpg: ['image/jpeg'], jpeg: ['image/jpeg'],
    txt: ['text/plain'], csv: ['text/csv', 'application/vnd.ms-excel', 'text/plain'],
    docx: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    xlsx: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    pptx: ['application/vnd.openxmlformats-officedocument.presentationml.presentation']
  };
  if (!allowed[extension] || allowed[extension].indexOf(String(blob.getContentType() || '').toLowerCase()) < 0) {
    throw new Error('Use PDF, PNG, JPEG, TXT, CSV, DOCX, XLSX or PPTX with its matching file type.');
  }
  const fingerprint = digest_(JSON.stringify([actionPlanId, form.planVersion, form.recordVersion, name, description, digest_(bytes)]));
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let operation = null;
  let file = null;
  let logRow = null;
  let table = null;
  let versionCell = null;
  let oldVersion = null;
  try {
    const actor = requireActiveActor_();
    const action = findActionRecord_(actionPlanId);
    requireOwner_(actor, action);
    operation = beginIdempotentOperation_('EVIDENCE', token, fingerprint);
    if (operation.duplicate) return operation.result;
    requireRegistryPlanEditable_(action);
    requireVersion_(form.recordVersion, action.recordVersion);
    requireVersion_(form.planVersion, action.planVersion);
    requireRegistryExecutionApproved_(action);
    table = readTable_(EVIDENCE_LOG_SHEET_NAME, ACTIONTRACKER_EVIDENCE_HEADERS);
    const folder = getEvidenceFolder_();
    const id = 'EVD-' + Utilities.getUuid();
    const now = new Date();
    file = folder.createFile(blob.copyBlob().setName(actionPlanId + '_' + id + '_' + name));
    logRow = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, { Evidence_ID: id, Action_Plan_ID: actionPlanId,
      Plan_Version: action.planVersion, File_ID: file.getId(), File_Name: name,
      File_URL: file.getUrl(), Evidence_Description: description,
      Uploaded_By: actor.email, Uploaded_At: now, Request_ID: token });
    versionCell = action.sheet.getRange(action.rowNumber, action.map.record_version + 1);
    oldVersion = versionCell.getValue();
    versionCell.setValue(action.recordVersion + 1);
    SpreadsheetApp.flush();
    const result = { success: true, evidenceId: id, actionPlanId: actionPlanId,
      fileUrl: file.getUrl(), recordVersion: action.recordVersion + 1 };
    completeIdempotentOperation_(operation.key, result);
    logRow = null; versionCell = null; file = null;
    return result;
  } catch (error) {
    try {
      if (logRow) table.sheet.getRange(logRow, 1, 1, table.headers.length).clearContent();
      if (versionCell) versionCell.setValue(oldVersion);

      if (file) file.setTrashed(true);
      SpreadsheetApp.flush();
      if (operation && !operation.duplicate) abandonIdempotentOperation_(operation.key);
    } catch (rollbackError) {
      if (operation) markOperationNeedsReview_(operation.key, { error: error.message, rollbackError: rollbackError.message });
      throw new Error('Upload save/rollback failed. Do not retry until Evidence_Log and the newly created file are reviewed.');
    }
    throw error;
  } finally { lock.releaseLock(); }
}

function generateUpdateId_(logSheet) {
  const headers = logSheet
    .getRange(1, 1, 1, logSheet.getLastColumn())
    .getDisplayValues()[0];
  const map = requireHeaders_(
    headers,
    ['Update ID'],
    UPDATE_LOG_SHEET_NAME
  );
  const idColumn = map[normalizeHeaderKey_('Update ID')] + 1;
  const lastRow = getLastDataRowByColumn_(logSheet, idColumn);
  let highestNumber = 0;

  if (lastRow >= 2) {
    logSheet
      .getRange(2, idColumn, lastRow - 1, 1)
      .getDisplayValues()
      .flat()
      .forEach(function(id) {
        const match = String(id).trim().match(/^UPD-(\d+)$/);
        if (!match) return;
        const number = Number(match[1]);
        if (number > highestNumber) highestNumber = number;
      });
  }

  return 'UPD-' + String(highestNumber + 1).padStart(5, '0');
}

function getActionUpdateHistory(
  actionPlanId
) {

  const ss =
    getTrackerSpreadsheet_();

  const sheet =
    ss.getSheetByName(
      UPDATE_LOG_SHEET_NAME
    );

  if (!sheet) {

    throw new Error(
      'Update_Log sheet not found.'
    );
  }

  actionPlanId =
    String(
      actionPlanId || ''
    ).trim();

  if (!actionPlanId) {

    throw new Error(
      'Action Plan ID is required.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {
    return [];
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();

  const idIndex =
    findHeaderIndex_(
      headers,
      'Action Plan ID'
    );

  if (idIndex === -1) {

    throw new Error(
      'Action Plan ID column not found in Update_Log.'
    );
  }

  function getValue(
    row,
    headerName
  ) {

    const index =
      findHeaderIndex_(
        headers,
        headerName
      );

    return index !== -1
      ? row[index]
      : '';
  }

  return data

    .filter(
      row =>
        String(
          row[idIndex]
        ).trim() ===
        actionPlanId
    )

    .map(
      row => ({

        updateId:
          getValue(
            row,
            'Update ID'
          ),

        updateDate:
          getValue(
            row,
            'Update Date'
          ),

        updatedBy:
          getValue(
            row,
            'Updated By'
          ),

        measurementMethod:getValue(row,'Measurement_Method') || 'Numeric KPI',
        planVersion:getValue(row,'Plan_Version'),
        currentPerformance:
          getValue(
            row,
            'Current Performance'
          ),

        actionStatus:
          getValue(
            row,
            'Action Status'
          ),

        comment:
          getValue(
            row,
            'Comment'
          )

      })
    )

    .reverse();
}

const REGISTRY_REVIEW_HEADERS = Object.freeze([
  'Review_ID', 'Action_Plan_ID', 'Plan_Version', 'Record_Version_Reviewed',
  'Decision', 'Review_Comment', 'Reviewed_By', 'Reviewer_Name', 'Reviewed_At', 'Request_ID'
]);
const REGISTRY_SUBMISSION_AUDIT_HEADERS = Object.freeze([
  'Plan_Version', 'Request_ID', 'Revision_Comment', 'Plan_Snapshot'
]);

function registrySiteKey_(value) {
  return normalizeActionTrackerSite_(value).toLowerCase();
}

function getConfiguredRegistrySites_() {
  const choices = ACTIONTRACKER_DEFAULT_SITES.concat(getSettingsOptionValues_('Site'));
  const seen = new Set();
  return choices.map(normalizeActionTrackerSite_).filter(function(site) {
    const key = registrySiteKey_(site);
    if (!key || seen.has(key)) return false;
    seen.add(key); return true;
  });
}

function parseApprovalSites_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return [];
  const configured = getConfiguredRegistrySites_();
  const parts = text.split(',').map(function(site) {
    const key = registrySiteKey_(site);
    return configured.find(function(choice) { return registrySiteKey_(choice) === key; });
  });

  if (parts.some(function(site) { return !site; })) return [];
  return Array.from(new Set(parts));
}

function registryActorApprovesSite_(actor, site) {
  const key = registrySiteKey_(site);
  return !!key && actor.canApprove && actor.approvalSites.some(function(assigned) {
    return registrySiteKey_(assigned) === key;
  });
}

const REGISTRY_GLIDE_HEADERS = Object.freeze(['Glide_Path_Available', 'Glide_Path_URL']);

function validateRegistryGlide_(payload) {
  const available = boundedText_(payload.glideAvailable, 'Time-phased glide path available', 10, true);
  if (['Yes', 'No'].indexOf(available) < 0) throw new Error('Select Yes or No for time-phased glide path availability.');
  const url = boundedText_(payload.glideUrl, 'Glide Path Link', 2000, available === 'Yes');
  if (available === 'No' && url) throw new Error('Remove the glide path link when availability is No.');
  if (available === 'Yes' && !isRegistryGlideUrl_(url)) {
    throw new Error('Use an HTTPS Google Drive file or Google Docs, Sheets, or Slides link for the glide path.');
  }
  return { Glide_Path_Available: available, Glide_Path_URL: url };
}

function isRegistryGlideUrl_(url) {
  if (/[\s<>\x00-\x1f]/.test(String(url || ''))) return false;
  return /^https:\/\/(?:docs\.google\.com\/(?:spreadsheets|document|presentation)\/d\/[A-Za-z0-9_-]{10,}(?:[/?#].*)?|drive\.google\.com\/(?:file\/d\/[A-Za-z0-9_-]{10,}(?:[/?#].*)?|open\?id=[A-Za-z0-9_-]{10,}(?:&.*)?))$/.test(String(url || ''));
}

function ensureRegistryGlideHeaders_() {
  const ss = getTrackerSpreadsheet_();
  const master = ss.getSheetByName(MASTER_SHEET_NAME), submission = ss.getSheetByName(SUBMISSION_LOG_SHEET_NAME);
  if (!master || !submission) throw new Error('Action_Master and Submission_Log are required.');

  const mh = master.getRange(1, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  requireHeaders_(mh, getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS), MASTER_SHEET_NAME);
  assertMasterFormulaSchema_(mh);
  const sh = submission.getRange(1, 1, 1, submission.getLastColumn()).getDisplayValues()[0];
  requireHeaders_(sh, getRequiredSubmissionLogHeaders_().filter(function(h) {
    return REGISTRY_SUBMISSION_AUDIT_HEADERS.indexOf(h) < 0;
  }), SUBMISSION_LOG_SHEET_NAME);
  ensureSubmissionAuditHeaders_(submission);
  [master, submission].forEach(function(sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const map = buildHeaderMapStrict_(headers, sheet.getName());
    const missing = REGISTRY_GLIDE_HEADERS.filter(function(h) { return map[normalizeHeaderKey_(h)] === undefined; });
    if (!missing.length) return;
    if (sheet.getMaxColumns() < headers.length + missing.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length + missing.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  });
}

function validateRegistryScope_(payload) {
  const scope = {};
  [['Channel', 'channel'], ['LOB', 'lob'], ['Cohort', 'cohort']].forEach(function(pair) {
    const text = boundedText_(payload[pair[1]], pair[0], 300, true);
    if (getSettingsOptionValues_(pair[0]).indexOf(text) < 0) {
      throw new Error(pair[0] + ' is not an allowed Settings choice. Refresh the form.');
    }
    scope[pair[0]] = text;
  });
  scope.Cohort_Details = boundedText_(payload.cohortDetails, 'Cohort Details', 2000,
    /^targeted(?:\s|\(|$)/i.test(scope.Cohort));
  return scope;
}

function resolveRegistryKpiSetting_(settings, name) {
  const key = String(name || '').trim().toLowerCase();
  if (settings[key]) return settings[key];

  const aliases = { 'smiley rate': 'smiley', 'qa': 'quality', 'productivity': 'productivity (tph/spd)' };
  return settings[aliases[key]] || null;
}

function registryDateInput_(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, getTrackerSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function snapshotNamedRow_(headers, row) {
  const snapshot = {};
  headers.forEach(function(header, index) {
    if (String(header).trim()) snapshot[header] = row[index] instanceof Date ? row[index].toISOString() : row[index];
  });
  return snapshot;
}

function ensureSubmissionAuditHeaders_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  const map = requireHeaders_(headers,
    getRequiredSubmissionLogHeaders_().filter(function(h) { return REGISTRY_SUBMISSION_AUDIT_HEADERS.indexOf(h) < 0; }),
    SUBMISSION_LOG_SHEET_NAME);
  const missing = REGISTRY_SUBMISSION_AUDIT_HEADERS.filter(function(h) { return map[normalizeHeaderKey_(h)] === undefined; });
  if (!missing.length) return;
  const first = headers.length + 1;
  if (sheet.getMaxColumns() < headers.length + missing.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length + missing.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, first, 1, missing.length).setValues([missing]);
}

function getRegistryReviews_(actionPlanId) {
  const table = readTable_('Review_Log', REGISTRY_REVIEW_HEADERS);
  return table.rows.filter(function(row) {
    return String(row[table.map.action_plan_id] || '').trim() === actionPlanId;
  }).map(function(row) {
    return { reviewId: String(row[table.map.review_id] || ''),
      planVersion: storedVersion_(row[table.map.plan_version]),
      recordVersionReviewed: storedVersion_(row[table.map.record_version_reviewed]),
      decision: String(row[table.map.decision] || ''), text: String(row[table.map.review_comment] || ''),
      email: String(row[table.map.reviewed_by] || ''), name: String(row[table.map.reviewer_name] || ''),
      reviewedAt: dateToClientIso_(row[table.map.reviewed_at]) };
  }).reverse();
}

function getRegistryApprovalState_(action) {
  const flag = action.row[action.map.approval_required];
  if (flag === false || flag === '' || flag == null) {
    return { required: false, executionAllowed: true,
      message: 'Existing plan: approval requirement is not enabled for this record.' };
  }
  if (flag !== true) return { required: true, executionAllowed: false,
    message: 'Approval_Required must be a checkbox boolean. Ask the maintainer to correct it.' };
  const status = String(action.row[action.map.review_status] || '').trim();
  const latestId = String(action.row[action.map.latest_review_id] || '').trim();
  const id = String(action.row[action.map['action plan id']] || '').trim();
  const table = readTable_('Review_Log', REGISTRY_REVIEW_HEADERS);
  const matching = table.rows.filter(function(row) { return String(row[table.map.review_id] || '').trim() === latestId; });
  let approved = false;
  if (latestId && matching.length === 1) {
    const review = matching[0];
    approved = status === 'Approved' && String(review[table.map.action_plan_id] || '').trim() === id &&
      Number(review[table.map.plan_version]) === action.planVersion &&
      Number.isSafeInteger(Number(review[table.map.record_version_reviewed])) &&
      Number(review[table.map.record_version_reviewed]) >= 1 &&
      Number(review[table.map.record_version_reviewed]) < action.recordVersion &&
      review[table.map.decision] === 'Approved';
  }
  return { required: true, executionAllowed: approved,
    message: approved ? 'Current plan version is approved for execution.' :
      'Execution is blocked until an assigned approver approves the current plan version. Review status: ' + (status || 'Not Recorded') + '.' };
}

function registryPlanCompleted_(action) {
  return String(action.row[action.map['action status']] || '').trim().toLowerCase() === 'completed';
}

function requireRegistryPlanEditable_(action) {
  if (registryPlanCompleted_(action)) throw registrySecurityError_(
    'Completed plans are read-only. Execution, plan details and execution evidence cannot be changed; effectiveness reviews remain available.',
    'PLAN_COMPLETED');
}

function requireRegistryExecutionApproved_(action) {
  const approval = getRegistryApprovalState_(action);
  if (!approval.executionAllowed) throw new Error(approval.message);
}

function requireRegistryReviewer_(actor, action) {
  const site = normalizeActionTrackerSite_(action.row[action.map.site]);
  if (!registryActorApprovesSite_(actor, site)) {
    throw registrySecurityError_('Approval requires Can_Approve checked and this action site/department listed in Approval_Sites.', 'APPROVAL_SITE_REQUIRED');
  }
}

function registryMasterChanges_(action, fields) {
  return Object.keys(fields).map(function(header) {
    const index = action.map[normalizeHeaderKey_(header)];
    if (index === undefined || (index >= 15 && index <= 19)) throw new Error('Controlled or missing master field: ' + header);
    return { range: action.sheet.getRange(action.rowNumber, index + 1), oldValue: action.row[index], value: fields[header] };
  });
}

function applyRegistryChanges_(changes) {
  changes.forEach(function(item) {
    if (typeof item.value === 'string') setLiteralCellValue_(item.range, item.value);
    else item.range.setValue(item.value);
  });
}

function rollbackRegistryWrite_(changes, table, row, operation, error) {
  try {
    if (row && table) table.sheet.getRange(row, 1, 1, table.headers.length).clearContent();
    changes.slice().reverse().forEach(function(item) {
      if (item.oldValue === '' || item.oldValue == null) item.range.clearContent();
      else if (typeof item.oldValue === 'string') setLiteralCellValue_(item.range, item.oldValue);
      else item.range.setValue(item.oldValue);
    });
    SpreadsheetApp.flush();
    if (operation && !operation.duplicate) abandonIdempotentOperation_(operation.key);
  } catch (restoreError) {
    if (operation) markOperationNeedsReview_(operation.key, { error: error.message, rollbackError: restoreError.message });
    throw new Error('Save and rollback could not be confirmed. Do not retry until the master and audit log are checked.');
  }
}

function submitReviewDecisionCore_(actionPlanId, decision, comment, expectedPlanVersion, expectedRecordVersion, requestToken) {
  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');
  if (['Approved', 'Returned for Revision', 'Declined'].indexOf(decision) < 0) throw new Error('Select a valid review decision.');
  comment = boundedText_(comment, 'Review Comment', 5000, false);
  if (!comment && decision === 'Returned for Revision') throw new Error('Required: Clearly state what needs to be revised before the plan can be resubmitted for approval.');
  if (!comment && decision === 'Declined') throw new Error('Required: Explain why the plan is being declined and why it should not proceed.');
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let operation = null, table = null, logRow = null, changes = [];
  try {
    const actor = requireActiveActor_('canApprove');
    const action = findActionRecord_(actionPlanId);
    requireRegistryReviewer_(actor, action);
    operation = beginIdempotentOperation_('REVIEW', requestToken, digest_(JSON.stringify([
      actionPlanId, decision, comment, expectedPlanVersion, expectedRecordVersion
    ])));
    if (operation.duplicate) return operation.result;
    requireRegistryPlanEditable_(action);
    requireVersion_(expectedPlanVersion, action.planVersion);
    requireVersion_(expectedRecordVersion, action.recordVersion);
    if (action.row[action.map.approval_required] !== true) throw new Error('This existing plan is not enrolled in the approval process.');
    if (action.row[action.map.review_status] !== 'Pending Review') throw new Error('Only a Pending Review plan may receive a decision.');
    if (action.row[action.map['action status']] !== 'Not Started') throw new Error('Plans awaiting proposal approval must be Not Started.');
    table = readTable_('Review_Log', REGISTRY_REVIEW_HEADERS);
    const reviewId = 'REV-' + Utilities.getUuid(), now = new Date();
    changes = registryMasterChanges_(action, { Review_Status: decision, Latest_Review_ID: reviewId,
      Record_Version: action.recordVersion + 1, 'Last Updated Date': now, 'Last Updated By': actor.email,
      'Action Status': decision === 'Approved' ? 'Ongoing' : decision === 'Declined' ? 'Cancelled' : 'Not Started' });
    logRow = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, { Review_ID: reviewId, Action_Plan_ID: actionPlanId, Plan_Version: action.planVersion,
      Record_Version_Reviewed: action.recordVersion, Decision: decision, Review_Comment: comment,
      Reviewed_By: actor.email, Reviewer_Name: actor.name, Reviewed_At: now, Request_ID: requestToken });
    applyRegistryChanges_(changes); SpreadsheetApp.flush();
    const result = { success: true, reviewId: reviewId, decision: decision, actionPlanId: actionPlanId,
      planVersion: action.planVersion, recordVersion: action.recordVersion + 1,
      actionStatus: decision === 'Approved' ? 'Ongoing' : decision === 'Declined' ? 'Cancelled' : 'Not Started' };
    completeIdempotentOperation_(operation.key, result); changes = []; logRow = null;
    return result;
  } catch (error) { rollbackRegistryWrite_(changes, table, logRow, operation, error); throw error; }
  finally { lock.releaseLock(); }
}

function validateRegistryRevision_(payload) {
  const siteKey = registrySiteKey_(requireNonBlankText_(payload.site, 'Site / Department'));
  const site = getConfiguredRegistrySites_().find(function(choice) { return registrySiteKey_(choice) === siteKey; });
  if (!site) throw new Error('Site / Department is not configured.');
  const scope = Object.assign(validateRegistryScope_(payload), validateRegistryGlide_(payload));
  const owner = validateTaskUsOwnerEmail_(payload.owner);
  const ownerEntry = getAccessEntry_(owner);
  if (!ownerEntry || !ownerEntry.active) throw new Error('Action Owner must be an active listed account.');
  const category = requireNonBlankText_(payload.category, 'Action Category');
  const type = requireNonBlankText_(payload.actionType, 'Action Type');
  validateClassification_(category, type);
  const department = boundedText_(payload.department, 'Department', 100, false);
  if (department && getSettingsOptionValues_('Department').indexOf(department) < 0) throw new Error('Department is not configured.');
  const measurement=registryMeasurementFields_(payload),governance=measurement.Measurement_Method==='Governance outcome';
  const kpi = governance?'':requireNonBlankText_(payload.kpi, 'KPI');
  const setting = getKpiSettings_()[kpi.toLowerCase()];
  if (!governance && !setting) throw new Error('Select a currently configured KPI.');
  const start = parseDateInputServer_(payload.startDate, 'Start Date');
  const end = parseDateInputServer_(payload.endDate, 'Target End Date');
  if (end < start) throw new Error('Target End Date cannot be earlier than Start Date.');
  const baseline = governance?'':parseKpiMetricServer_(payload.baseline, kpi, setting, 'Baseline Performance');
  return Object.assign({}, scope, measurement, { Site: site, 'Action Plan Name': boundedText_(payload.actionPlanName, 'Action Plan Name', 200, true),
    'Action Description': boundedText_(payload.actionDescription, 'Action Steps', 5000, true),
    'Target KPI / Behavior': kpi, 'Action Owner': owner, Framework_Category: category,
    Action_Type: type, Activity: boundedText_(payload.activity, 'Activity', 2000, true), Department: department,
    'Baseline Performance': baseline, 'End Goal / Target': governance?'':parseKpiMetricServer_(payload.goal, kpi, setting, 'End Goal / Target'),
    'Start Date': start, 'Target End Date': end });
}

function submitPlanRevisionCore_(actionPlanId, payload, revisionComment, expectedPlanVersion, expectedRecordVersion, requestToken) {
  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');
  revisionComment = boundedText_(revisionComment, 'Revision Comment', 2000, true);
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let operation = null, table = null, logRow = null, changes = [];
  try {
    const actor = requireActiveActor_();
    let action = findActionRecord_(actionPlanId); requireOwner_(actor, action);
    operation = beginIdempotentOperation_('REVISION', requestToken, digest_(JSON.stringify([
      actionPlanId, payload, revisionComment, expectedPlanVersion, expectedRecordVersion
    ])));
    if (operation.duplicate) return operation.result;
    requireRegistryPlanEditable_(action);
    requireVersion_(expectedPlanVersion, action.planVersion); requireVersion_(expectedRecordVersion, action.recordVersion);
    if (action.row[action.map.approval_required] !== true ||
      ['Pending Review', 'Returned for Revision', 'Approved'].indexOf(action.row[action.map.review_status]) < 0 ||
      ['Completed', 'Cancelled'].indexOf(action.row[action.map['action status']]) >= 0) {
      throw new Error('This plan cannot be revised. Only an open plan enrolled in approval may be revised by its owner.');
    }
    const fields = validateRegistryRevision_(payload);
    ensureRegistryGovernanceHeaders_();
    ensureRegistryGlideHeaders_();
    action = findActionRecord_(actionPlanId);
    const now = new Date(), nextPlan = action.planVersion + 1;
    const oldKpi = String(action.row[action.map['target kpi / behavior']] || '');
    Object.assign(fields, { Plan_Version: nextPlan, Record_Version: action.recordVersion + 1,
      Review_Status: 'Pending Review', Latest_Review_ID: '', Approval_Required: true,
      'Action Status': 'Not Started', 'Completion Date': '', 'Last Updated Date': now, 'Last Updated By': actor.email,
      Effectiveness_Status: 'Not Yet Assessed', Latest_Assessment_ID: '' });

    const oldSetting = resolveRegistryKpiSetting_(getKpiSettings_(), oldKpi);
    if (fields.Measurement_Method==='Governance outcome' || registryMeasurementRead_(action.row,action.map).measurementMethod!==fields.Measurement_Method || !oldSetting || oldSetting.name !== fields['Target KPI / Behavior']) fields['Current Performance'] = fields['Baseline Performance'];
    changes = registryMasterChanges_(action, fields);
    const previous = snapshotNamedRow_(action.headers, action.row);
    const next = Object.assign({}, previous);
    Object.keys(fields).forEach(function(header) {
      const actualHeader = action.headers[action.map[normalizeHeaderKey_(header)]];
      next[actualHeader] = fields[header] instanceof Date ? fields[header].toISOString() : fields[header];
    });
    const snapshot = JSON.stringify({ previous: previous, current: next });
    if (snapshot.length > 45000) throw new Error('Plan snapshot is too large for its audit cell. No revision was saved.');
    const sheet = getTrackerSpreadsheet_().getSheetByName(SUBMISSION_LOG_SHEET_NAME);
    if (!sheet) throw new Error('Submission_Log sheet not found.');
    ensureSubmissionAuditHeaders_(sheet);
    table = readTable_(SUBMISSION_LOG_SHEET_NAME, getRequiredSubmissionLogHeaders_());
    const submissionId = generateSubmissionId_(sheet, table.headers);
    const log = { 'Submission ID': submissionId, 'Action Plan ID': actionPlanId, 'Submission Date': now,
      'Submission Source': 'Registry Plan Revision', 'Submitted By': actor.email, 'Action Status': 'Not Started',
      Plan_Version: nextPlan, Request_ID: requestToken, Revision_Comment: revisionComment, Plan_Snapshot: snapshot,
      'Action Category': fields.Framework_Category };
    ['Site', 'Action Plan Name', 'Action Description', 'Target KPI / Behavior', 'Action Owner', 'Baseline Performance',
      'End Goal / Target', 'Start Date', 'Target End Date', 'Channel', 'LOB', 'Cohort', 'Cohort_Details', 'Glide_Path_Available', 'Glide_Path_URL'].forEach(function(header) { log[header] = fields[header]; });
    logRow = table.sheet.getLastRow() + 1; appendNamedLog_(table, log);
    applyRegistryChanges_(changes); SpreadsheetApp.flush();
    const result = { success: true, actionPlanId: actionPlanId, planVersion: nextPlan,
      recordVersion: action.recordVersion + 1, submissionId: submissionId };
    completeIdempotentOperation_(operation.key, result); changes = []; logRow = null;
    return result;
  } catch (error) { rollbackRegistryWrite_(changes, table, logRow, operation, error); throw error; }
  finally { lock.releaseLock(); }
}

const REGISTRY_ASSESSMENT_SHEET = 'Assessment_Log';
const REGISTRY_EFFECTIVENESS_STATUSES = Object.freeze([
  'Validated', 'Partially Effective', 'No Observable Effect', 'Insufficient Evidence'
]);
const REGISTRY_ASSESSMENT_DECISIONS = Object.freeze(['Continue', 'Adjust', 'Stop', 'Recommend Scaling']);
const REGISTRY_ASSESSMENT_HEADERS = Object.freeze([
  'Assessment_ID', 'Action_Plan_ID', 'Plan_Version', 'Record_Version_Reviewed',
  'Effectiveness_Status', 'Period_Start', 'Period_End', 'Expected_Impact',
  'Behavior_Changes', 'Observed_Impact', 'Expectations_Result', 'Sustained_Result',
  'Relationship', 'Other_Factors', 'Assessment_Rationale', 'Evidence_IDs', 'Evidence_URLs',
  'Decision', 'Decision_Rationale', 'Decision_Owner', 'Next_Review_Date',
  'Repeatability', 'Transferability', 'Risk_Result', 'Risk_Notes',
  'Assessor_Email', 'Assessor_Name', 'Assessed_At', 'Metric_Snapshot', 'Request_ID'
]);

function getRegistryAssessmentEligibility_(action, approval) {
  approval = approval || getRegistryApprovalState_(action);
  if (!approval.executionAllowed) return { allowed: false, message: 'Approve the current plan version before assessing its effectiveness.' };
  const status = String(action.row[action.map['action status']] || '').trim();
  if (['Ongoing', 'On Hold', 'Completed'].indexOf(status) < 0) {
    return { allowed: false, message: 'Effectiveness assessment is available for Ongoing, On Hold, or Completed actions.' };
  }
  return { allowed: true, message: '' };
}

function ensureRegistryAssessmentLog_() {
  const ss = getTrackerSpreadsheet_();
  let sheet = ss.getSheetByName(REGISTRY_ASSESSMENT_SHEET);
  if (!sheet) sheet = ss.insertSheet(REGISTRY_ASSESSMENT_SHEET);
  if (!sheet.getLastRow()) {
    if (sheet.getMaxColumns() < REGISTRY_ASSESSMENT_HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), REGISTRY_ASSESSMENT_HEADERS.length - sheet.getMaxColumns());
    sheet.getRange(1, 1, 1, REGISTRY_ASSESSMENT_HEADERS.length).setValues([REGISTRY_ASSESSMENT_HEADERS.slice()]);
    sheet.setFrozenRows(1);
  }

  return readTable_(REGISTRY_ASSESSMENT_SHEET, REGISTRY_ASSESSMENT_HEADERS);
}

function setupRegistryAssessmentsCore_() {
  requireActiveActor_('canAdmin');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const table = ensureRegistryAssessmentLog_();
    SpreadsheetApp.flush();
    const result = { success: true, sheet: REGISTRY_ASSESSMENT_SHEET, headers: table.headers.length, version: ACTIONTRACKER_VERSION };
      return result;
  } finally { lock.releaseLock(); }
}

function getRegistryAssessments_(actionPlanId) {
  if (!getTrackerSpreadsheet_().getSheetByName(REGISTRY_ASSESSMENT_SHEET)) return [];
  const table = readTable_(REGISTRY_ASSESSMENT_SHEET, REGISTRY_ASSESSMENT_HEADERS);
  function text(row, key) { return String(row[table.map[normalizeHeaderKey_(key)]] || ''); }
  function calendar(row, key) {
    const value = row[table.map[normalizeHeaderKey_(key)]];
    return value instanceof Date ? Utilities.formatDate(value, getTrackerSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(value || '');
  }
  return table.rows.filter(function(row) { return text(row, 'Action_Plan_ID').trim() === actionPlanId; })
    .map(function(row) {
      let snapshot = null;
      try { snapshot = JSON.parse(text(row, 'Metric_Snapshot')); } catch (error) {   }
      return {
        assessmentId: text(row, 'Assessment_ID'), planVersion: storedVersion_(row[table.map.plan_version]),
        recordVersionReviewed: storedVersion_(row[table.map.record_version_reviewed]),
        status: text(row, 'Effectiveness_Status'), periodStart: calendar(row, 'Period_Start'), periodEnd: calendar(row, 'Period_End'),
        expectedImpact: text(row, 'Expected_Impact'), behaviorChanges: text(row, 'Behavior_Changes'), observedImpact: text(row, 'Observed_Impact'),
        expectationsResult: text(row, 'Expectations_Result'), sustainedResult: text(row, 'Sustained_Result'),
        relationship: text(row, 'Relationship'), otherFactors: text(row, 'Other_Factors'), rationale: text(row, 'Assessment_Rationale'),
        evidenceIds: text(row, 'Evidence_IDs').split('\n').filter(Boolean), evidenceUrls: text(row, 'Evidence_URLs').split('\n').filter(Boolean),
        governanceCriteriaResult:text(row,'Governance_Criteria_Result'),
        decision: text(row, 'Decision'), decisionRationale: text(row, 'Decision_Rationale'),
        decisionOwner: text(row, 'Decision_Owner'), nextReviewDate: calendar(row, 'Next_Review_Date'),
        repeatability: text(row, 'Repeatability'), transferability: text(row, 'Transferability'),
        riskResult: text(row, 'Risk_Result'), riskNotes: text(row, 'Risk_Notes'),
        assessorEmail: text(row, 'Assessor_Email'), assessorName: text(row, 'Assessor_Name'),
        assessedAt: dateToClientIso_(row[table.map.assessed_at]), snapshot: snapshot
      };
    }).reverse();
}

function registryAssessmentChoice_(value, label, allowed) {
  const result = boundedText_(value, label, 100, true);
  if (allowed.indexOf(result) < 0) throw new Error('Select a valid ' + label + '.');
  return result;
}

function registryAssessmentDate_(value, label) {
  const result = boundedText_(value, label, 10, true);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(result)) throw new Error(label + ' must use YYYY-MM-DD.');
  parseDateInputServer_(result, label);
  return result;
}

function validateRegistryAssessment_(payload, action) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('Assessment details are required.');
  const fields = {};
  fields.status = registryAssessmentChoice_(payload.status, 'effectiveness status', REGISTRY_EFFECTIVENESS_STATUSES);
  fields.decision = registryAssessmentChoice_(payload.decision, 'next decision', REGISTRY_ASSESSMENT_DECISIONS);
  fields.periodStart = registryAssessmentDate_(payload.periodStart, 'Assessment period start');
  fields.periodEnd = registryAssessmentDate_(payload.periodEnd, 'Assessment period end');
  fields.nextReviewDate = registryAssessmentDate_(payload.nextReviewDate, 'Next review date');
  const today = registryToday_();
  if (fields.periodStart > today) throw new Error('Observation period start cannot be after today (' + today + ', Manila). Future dates belong under Next effectiveness review date.');
  if (fields.periodEnd > today) throw new Error('Observation period end cannot be after today (' + today + ', Manila). Use the latest date covered by your evidence.');
  if (fields.periodStart > fields.periodEnd) throw new Error('Observation period start must be on or before observation period end.');
  if (fields.nextReviewDate < today || fields.nextReviewDate < fields.periodEnd) throw new Error('Next review date must be today or later and cannot precede the assessment period.');
  [ ['expectedImpact', 'Expected impact'], ['behaviorChanges', 'Behavior changes'], ['observedImpact', 'Observed impact'],
    ['otherFactors', 'Other contributing factors'], ['rationale', 'Assessment rationale'], ['decisionRationale', 'Decision rationale']
  ].forEach(function(pair) { fields[pair[0]] = boundedText_(payload[pair[0]], pair[1], 3000, true); });
  fields.expectationsResult = registryAssessmentChoice_(payload.expectationsResult, 'expectations result', ['Met', 'Partly Met', 'Not Met', 'Too Early']);
  fields.sustainedResult = registryAssessmentChoice_(payload.sustainedResult, 'sustained result', ['Sustained', 'Not Sustained', 'Too Early']);
  fields.relationship = registryAssessmentChoice_(payload.relationship, 'relationship to the action', ['Supported', 'Plausible', 'Unknown', 'Unrelated']);
  fields.decisionOwner = validateTaskUsOwnerEmail_(payload.decisionOwner);
  const owner = getAccessEntry_(fields.decisionOwner);
  if (!owner || !owner.active) throw new Error('Decision owner must be an active listed TaskUs account.');
  if (payload.evidenceIds != null && !Array.isArray(payload.evidenceIds)) throw new Error('Select evidence from this plan.');
  fields.evidenceIds = Array.from(new Set(payload.evidenceIds || []));
  if (fields.evidenceIds.length > 10) throw new Error('Select at most 10 evidence files.');
  const evidence = getActionEvidence_(String(action.row[action.map['action plan id']]).trim());
  fields.evidenceIds.forEach(function(id) {
    const item = evidence.find(function(entry) { return entry.evidenceId === id && entry.planVersion === action.planVersion; });
    if (typeof id !== 'string' || !item) {
      throw new Error('Assessment evidence must belong to this action and its current plan version.');
    }
    try {
      if (DriveApp.getFileById(item.fileId).isTrashed()) throw new Error('Trashed evidence');
    } catch (error) { throw new Error('A selected evidence file is missing or unavailable. Select available proof or use a reviewer-accessible source link.'); }
  });
  const rawUrls = boundedText_(payload.evidenceUrls, 'Supporting source links', 10000, false);
  fields.evidenceUrls = Array.from(new Set(rawUrls.split(/\r?\n/).map(function(value) { return value.trim(); }).filter(Boolean)));
  if (fields.evidenceUrls.length > 5 || fields.evidenceUrls.some(function(url) { return !isRegistryGlideUrl_(url); })) {
    throw new Error('Use up to 5 HTTPS Google Drive, Docs, Sheets, or Slides source links, one per line.');
  }
  const hasEvidence = fields.evidenceIds.length + fields.evidenceUrls.length > 0;
  if (fields.status !== 'Insufficient Evidence' && !hasEvidence) throw new Error('Attach evidence or a source link for this effectiveness rating, or select Insufficient Evidence.');
  if (fields.status === 'Validated' && (fields.expectationsResult !== 'Met' || fields.relationship !== 'Supported')) {
    throw new Error('Validated requires expectations Met and a Supported relationship explained by the evidence.');
  }
  if (fields.status === 'No Observable Effect' && (action.row[action.map['action status']] !== 'Completed' || fields.expectationsResult !== 'Not Met')) {
    throw new Error('No Observable Effect requires a Completed action with expectations Not Met.');
  }
  fields.governanceCriteriaResult=registryMeasurementRead_(action.row,action.map).measurementMethod==='Governance outcome'?boundedText_(payload.governanceCriteriaResult,'Evidence against each acceptance criterion',3000,true):'';
  fields.repeatability = ''; fields.transferability = ''; fields.riskResult = ''; fields.riskNotes = '';
  if (fields.decision === 'Recommend Scaling') {
    if (fields.status !== 'Validated' || fields.sustainedResult !== 'Sustained') throw new Error('Recommend Scaling requires Validated impact and Sustained results.');
    fields.repeatability = boundedText_(payload.repeatability, 'Repeatability evidence', 3000, true);
    fields.transferability = boundedText_(payload.transferability, 'Transferability and adaptations', 3000, true);
    fields.riskResult = registryAssessmentChoice_(payload.riskResult, 'risk review', ['Acceptable', 'Unacceptable', 'Not Reviewed']);
    fields.riskNotes = boundedText_(payload.riskNotes, 'Risk review rationale', 3000, true);
    if (fields.riskResult !== 'Acceptable') throw new Error('Resolve or review the risks before recommending scaling.');
  }
  return fields;
}

function registryAssessmentSnapshot_(action) {
  function value(header) { return action.row[action.map[normalizeHeaderKey_(header)]]; }
  const kpi = String(value('Target KPI / Behavior') || '');
  const setting = resolveRegistryKpiSetting_(getKpiSettings_(), kpi) || { unit: '', direction: '' };
  return Object.assign(registryMeasurementRead_(action.row,action.map), { kpi: kpi, unit: setting.unit, direction: setting.direction,
    baseline: parseStoredMetricNumber_(value('Baseline Performance'), setting.unit),
    current: parseStoredMetricNumber_(value('Current Performance'), setting.unit),
    goal: parseStoredMetricNumber_(value('End Goal / Target'), setting.unit),
    actionStatus: String(value('Action Status') || ''), site: String(value('Site') || ''),
    channel: String(value('Channel') || ''), lob: String(value('LOB') || ''), cohort: String(value('Cohort') || '') });
}

function submitEffectivenessAssessmentCore_(actionPlanId, payload, expectedPlanVersion, expectedRecordVersion, requestToken) {
  actionPlanId = requireNonBlankText_(actionPlanId, 'Action Plan ID');
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let operation = null, table = null, logRow = null, changes = [], followTable = null, followRow = null, followCount = 0;
  try {
    const actor = requireActiveActor_('canAudit');
    const action = findActionRecord_(actionPlanId);
    const eligibility = getRegistryAssessmentEligibility_(action);
    if (!eligibility.allowed) throw new Error(eligibility.message);
    operation = beginIdempotentOperation_('ASSESSMENT', requestToken, digest_(JSON.stringify([
      actionPlanId, payload, expectedPlanVersion, expectedRecordVersion
    ])));
    if (operation.duplicate) return operation.result;
    requireVersion_(expectedPlanVersion, action.planVersion);
    requireVersion_(expectedRecordVersion, action.recordVersion);
    const fields = validateRegistryAssessment_(payload, action);
    const commitment = validateRegistryFollowUpCommitment_(payload, fields, actor);
    table = ensureRegistryAssessmentLog_();
    followTable = registryEnsureLog_('Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS);
    registryAppendHeaders_(table.sheet,['Governance_Criteria_Result']);table=readTable_(REGISTRY_ASSESSMENT_SHEET,REGISTRY_ASSESSMENT_HEADERS);
    const id = 'ASM-' + Utilities.getUuid(), now = new Date();
    changes = registryMasterChanges_(action, { Effectiveness_Status: fields.status, Latest_Assessment_ID: id,
      Record_Version: action.recordVersion + 1 });
    logRow = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, { Assessment_ID: id, Action_Plan_ID: actionPlanId, Plan_Version: action.planVersion,
      Record_Version_Reviewed: action.recordVersion, Effectiveness_Status: fields.status,
      Period_Start: fields.periodStart, Period_End: fields.periodEnd, Expected_Impact: fields.expectedImpact,
      Behavior_Changes: fields.behaviorChanges, Observed_Impact: fields.observedImpact,
      Expectations_Result: fields.expectationsResult, Sustained_Result: fields.sustainedResult,
      Relationship: fields.relationship, Other_Factors: fields.otherFactors, Assessment_Rationale: fields.rationale,
      Evidence_IDs: fields.evidenceIds.join('\n'), Evidence_URLs: fields.evidenceUrls.join('\n'),
      Governance_Criteria_Result:fields.governanceCriteriaResult,
      Decision: fields.decision, Decision_Rationale: fields.decisionRationale, Decision_Owner: fields.decisionOwner,
      Next_Review_Date: fields.nextReviewDate, Repeatability: fields.repeatability, Transferability: fields.transferability,
      Risk_Result: fields.riskResult, Risk_Notes: fields.riskNotes, Assessor_Email: actor.email,
      Assessor_Name: actor.name, Assessed_At: now, Metric_Snapshot: JSON.stringify(Object.assign(registryAssessmentSnapshot_(action), { observationDefaults: registryAssessmentDefaults_(action, getRegistryAssessments_(actionPlanId)) })), Request_ID: requestToken });
    const followUpId = 'FUP-' + Utilities.getUuid();
    followRow = followTable.sheet.getLastRow() + 1;
    followCount = 1;
    registryCreateFollowUp_(followTable, followUpId, action, id, fields, commitment, actor, requestToken, now, fields.decision === 'Continue' && commitment.due === fields.nextReviewDate ? 'Effectiveness Review' : 'Decision Delivery');
    let reviewFollowUpId = '';
    if (fields.decision !== 'Continue' || commitment.due !== fields.nextReviewDate) {
      reviewFollowUpId = 'FUP-' + Utilities.getUuid(); followCount = 2;
      registryCreateFollowUp_(followTable, reviewFollowUpId, action, id, fields, { due: fields.nextReviewDate, verifier: commitment.verifier, deliverable: 'Complete the next effectiveness review: behavior/KPI movement, expectations, sustained effect, other contributing factors, and the next decision. Record a new assessment with supporting evidence.' }, actor, requestToken, now, 'Effectiveness Review');
    }
    applyRegistryChanges_(changes); SpreadsheetApp.flush();
    const result = { success: true, followUpId: followUpId, reviewFollowUpId: reviewFollowUpId, assessmentId: id, actionPlanId: actionPlanId, status: fields.status,
      decision: fields.decision, planVersion: action.planVersion, recordVersion: action.recordVersion + 1 };
    completeIdempotentOperation_(operation.key, result); changes = []; logRow = null; followRow = null;
    return result;
  } catch (error) {
    if (followRow) {
      try { followTable.sheet.getRange(followRow, 1, followCount, followTable.headers.length).clearContent(); }
      catch (restoreError) { if (operation) markOperationNeedsReview_(operation.key, { error: error.message, followUpRollbackError: restoreError.message }); throw new Error('Assessment/follow-up rollback needs maintainer review. Do not retry.'); }
    }
    rollbackRegistryWrite_(changes, table, logRow, operation, error); throw error;
  }
  finally { lock.releaseLock(); }
}

function registryValidateSecuritySetup_(warnings) {
  const access = readTable_(ACCESS_SHEET_NAME, [
    'Email', 'Name', 'Role', 'Active', 'Can_Submit', 'Can_Comment',
    'Can_Audit', 'Can_Approve', 'Approval_Sites', 'Can_Admin'
  ]);
  const maintainers = access.rows.filter(function(row) {
    return row[access.map.active] === true && row[access.map.can_admin] === true;
  });
  if (!maintainers.length) throw new Error('Setup incomplete: check Active and Can_Admin for at least one authorized maintainer in Access_Control.');
  maintainers.forEach(function(row) {
    const email = String(row[access.map.email] || '').trim().toLowerCase();
    if (!/^[A-Z0-9._%+\-]+@taskus\.com$/i.test(email)) throw new Error('Setup incomplete: an active Can_Admin row does not have a valid TaskUs email.');
    getAccessEntry_(email);
  });
  const sheet = getTrackerSpreadsheet_().getSheetByName('Security_Log');
  if (!sheet || !sheet.getLastRow() || !sheet.getLastColumn()) {
    throw new Error('Setup incomplete: Security_Log is missing or empty. Run setupRegistryFollowUps using an active Can_Admin account.');
  }
  requireHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0],
    REGISTRY_SECURITY_HEADERS, 'Security_Log');
  const owner = String(PropertiesService.getScriptProperties().getProperty('FOLLOWUP_TRIGGER_OWNER') || '');
  const checks = { activeMaintainerCount: maintainers.length, securityLogReady: true,
    reminderOwnerAuthorized: false, reminderTriggerStatus: 'NOT_CONFIGURED', notificationChannel: 'GOOGLE_CHAT', chatConfigured: !!registryChatWebhookUrl_(false) };
  if (registryReminderSettings_().mode === 'LIVE') registryChatWebhookUrl_(true);
  else if (!checks.chatConfigured) warnings.push('Google Chat is paused and its webhook is not configured.');
  if (!owner) {
    warnings.push('Daily reminder trigger owner is not recorded. Automatic reminders are not configured; manual Google Chat sending remains available to authorized maintainers once configured.');
    return checks;
  }
  if (owner !== owner.trim().toLowerCase() || !/^[A-Z0-9._%+\-]+@taskus\.com$/i.test(owner)) {
    throw new Error('Setup incomplete: FOLLOWUP_TRIGGER_OWNER must contain the lowercase TaskUs email of the installed trigger owner.');
  }
  const actor = getAccessEntry_(owner);
  if (!actor || !actor.active || !actor.canAdmin) {
    throw new Error('Setup incomplete: the recorded reminder-trigger owner must have Active and Can_Admin checked in Access_Control.');
  }
  checks.reminderOwnerAuthorized = true;

  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (effective !== owner) {
    checks.reminderTriggerStatus = 'OWNER_CHECK_REQUIRED';
    warnings.push('Reminder owner authorization is valid. Run validateRegistrySetup from that owner\'s Apps Script account to verify the installed reminder trigger.');
    return checks;
  }
  const triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === 'runRegistryFollowUpReminders_';
  });
  if (triggers.length !== 1) {
    throw new Error('Setup incomplete: the recorded owner must have exactly one runRegistryFollowUpReminders_ trigger installed. Review the owner\'s Apps Script triggers before relying on automatic delivery.');
  }
  checks.reminderTriggerStatus = 'INSTALLED';
  return checks;
}

function validateRegistrySetup() {
  requireActiveActor_();
  readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS));
  assertMasterFormulaSchema_(readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_()).headers);
  readTable_('Review_Log', REGISTRY_REVIEW_HEADERS);
  readTable_(SUBMISSION_LOG_SHEET_NAME, getRequiredSubmissionLogHeaders_().filter(function(h) { return REGISTRY_SUBMISSION_AUDIT_HEADERS.indexOf(h) < 0; }));
  ['Channel', 'LOB', 'Cohort', 'Site'].forEach(function(header) {
    if (!getSettingsOptionValues_(header).length) throw new Error('Settings.' + header + ' has no choices.');
  });
  const settings = getKpiSettings_();
  Object.keys(settings).forEach(function(key) {
    if (!/^(higher|lower)( is better)?$/i.test(settings[key].direction)) throw new Error('Invalid KPI direction: ' + settings[key].name);
  });
  const warnings = [];
  const securityChecks = registryValidateSecuritySetup_(warnings);
  if (getTrackerSpreadsheet_().getSheetByName(REGISTRY_ASSESSMENT_SHEET)) readTable_(REGISTRY_ASSESSMENT_SHEET, REGISTRY_ASSESSMENT_HEADERS);
  else warnings.push('Assessment_Log will be created by setupRegistryAssessments or the first authorized assessment save.');
  [['Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS], ['Notification_Log', REGISTRY_NOTIFICATION_HEADERS], ['Follow_Up_Routing', REGISTRY_ROUTING_HEADERS]].forEach(function(pair) {
    if (getTrackerSpreadsheet_().getSheetByName(pair[0])) readTable_(pair[0], pair[1]);
    else warnings.push(pair[0] + ' will be created by setupRegistryFollowUps.');
  });
  registryReminderSettings_();
  const master = readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_());
  master.rows.forEach(function(row) {
    const id = String(row[master.map['action plan id']] || '').trim();
    if (!id) return;
    const kpi = String(row[master.map['target kpi / behavior']] || '');
    const setting = resolveRegistryKpiSetting_(settings, kpi);
    if (registryMeasurementRead_(row,master.map).measurementMethod==='Governance outcome') return;
    if (!setting) warnings.push(id + ': historical KPI is not configured (' + kpi + ').');
    else if (setting.unit === '%') ['Baseline Performance', 'Current Performance', 'End Goal / Target'].forEach(function(h) {
      const value = parseStoredMetricNumber_(row[master.map[normalizeHeaderKey_(h)]], '%');
      if (value !== null && (value < 0 || value > 1)) warnings.push(id + ': check percentage storage in ' + h + '.');
    });
  });
  const result = { success: true, version: ACTIONTRACKER_VERSION, warnings: warnings, securityChecks: securityChecks };
  result.chatEventChecks = registryValidateChatEventSetup_(warnings);
    return result;
}

const REGISTRY_FOLLOWUP_HEADERS = Object.freeze([
  'Event_ID', 'Follow_Up_ID', 'Assessment_ID', 'Action_Plan_ID', 'Plan_Version',
  'Follow_Up_Version', 'Decision', 'Commitment_Type', 'Deliverable', 'Owner_Email', 'Verifier_Email',
  'Original_Due_Date', 'Due_Date', 'Next_Review_Date', 'Status', 'Update_Notes',
  'Evidence_URLs', 'Submitted_At', 'Verification_Due_Date', 'Closed_At',
  'Extension_Requested_Date', 'Extension_Reason', 'Event_Type', 'Actor_Email', 'Event_At', 'Request_ID'
]);
const REGISTRY_NOTIFICATION_HEADERS = Object.freeze([
  'Notification_ID', 'Recipient_Email', 'Subject', 'Targets_JSON', 'Status',
  'Prepared_At', 'Sent_At', 'Error', 'Run_ID'
]);
const REGISTRY_ROUTING_HEADERS = Object.freeze(['Owner_Email', 'Manager_Email', 'Executive_Email', 'Active']);

function registryEnsureLog_(name, headers) {
  const ss = getTrackerSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (!sheet.getLastRow()) {
    if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]); sheet.setFrozenRows(1);
  }
  return readTable_(name, headers);
}
function setupRegistryFollowUpsCore_() {
  requireActiveActor_('canAdmin');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    ensureRegistryAssessmentLog_();
    registryEnsureLog_('Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS);
    registryEnsureLog_('Notification_Log', REGISTRY_NOTIFICATION_HEADERS);
    registryEnsureLog_('Follow_Up_Routing', REGISTRY_ROUTING_HEADERS);
    registryEnsureLog_('Security_Log', REGISTRY_SECURITY_HEADERS);
    SpreadsheetApp.flush();
    return { success: true, mode: registryReminderSettings_().mode, version: ACTIONTRACKER_VERSION };
  } finally { lock.releaseLock(); }
}
function registryCalendar_(value) {
  return value instanceof Date ? Utilities.formatDate(value, getTrackerSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(value || '');
}
function registryToday_(now) { return Utilities.formatDate(now || new Date(), 'Asia/Manila', 'yyyy-MM-dd'); }
function registryAddDays_(date, count) {
  return new Date(Date.parse(date + 'T00:00:00Z') + count * 86400000).toISOString().slice(0, 10);
}
function registryReminderSettings_() {
  const p = PropertiesService.getScriptProperties();
  const raw = Number(p.getProperty('FOLLOWUP_VERIFICATION_DAYS') || 2);
  if (!Number.isInteger(raw) || raw < 1 || raw > 30) throw new Error('FOLLOWUP_VERIFICATION_DAYS must be 1–30.');
  const stored = p.getProperty('REGISTRY_CHAT_MODE');
  if (stored && !['LIVE', 'DEV', 'PAUSED'].includes(stored)) throw new Error('Set REGISTRY_CHAT_MODE to LIVE or PAUSED.');
  const previous = p.getProperty('FOLLOWUP_EMAIL_MODE');
  const paused = stored ? stored !== 'LIVE' : (previous === 'PAUSED' || previous === 'DEV');
  return { mode: paused ? 'PAUSED' : 'LIVE', verificationDays: raw, channel: 'GOOGLE_CHAT',
    webhookConfigured: !!registryChatWebhookUrl_(false), appUrl: String(p.getProperty('FOLLOWUP_APP_URL') || '') };
}
function registryVerifier_(email, owner) {
  email = validateTaskUsOwnerEmail_(email);
  const person = getAccessEntry_(email);
  if (!person || !person.active || !person.canAudit) throw new Error('Verifier must be active with Can_Audit checked.');
  if (email === owner) throw new Error('Select a verifier other than the follow-up owner.');
  return email;
}
function validateRegistryFollowUpCommitment_(payload, fields, actor) {
  const due = registryAssessmentDate_(payload.followUpDueDate, 'Follow-up completion deadline');
  if (due < registryToday_()) throw new Error('Follow-up completion deadline must be today or later.');
  return { due: due, deliverable: boundedText_(payload.followUpDeliverable, 'Follow-up deliverable', 3000, true),
    verifier: registryVerifier_(payload.followUpVerifier || actor.email, fields.decisionOwner) };
}
function registryCreateFollowUp_(table, id, action, assessmentId, fields, commitment, actor, token, now, type) {
  return appendNamedLog_(table, {
    Event_ID: 'FUE-' + Utilities.getUuid(), Follow_Up_ID: id, Assessment_ID: assessmentId,
    Action_Plan_ID: String(action.row[action.map['action plan id']]), Plan_Version: action.planVersion,
    Follow_Up_Version: 1, Decision: fields.decision, Commitment_Type: type, Deliverable: commitment.deliverable,
    Owner_Email: fields.decisionOwner, Verifier_Email: commitment.verifier,
    Original_Due_Date: commitment.due, Due_Date: commitment.due, Next_Review_Date: fields.nextReviewDate,
    Status: 'Open', Event_Type: 'Created', Actor_Email: actor.email, Event_At: now, Request_ID: token
  });
}
function registryFollowUpRows_() {
  if (!getTrackerSpreadsheet_().getSheetByName('Follow_Up_Log')) return { items: [], history: [] };
  const t = readTable_('Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS), history = [], latest = {};
  t.rows.forEach(function(row) {
    function v(h) { return row[t.map[normalizeHeaderKey_(h)]]; }
    const id = String(v('Follow_Up_ID') || '').trim(); if (!id) return;
    const item = { id: id, assessmentId: String(v('Assessment_ID') || ''), actionPlanId: String(v('Action_Plan_ID') || ''),
      planVersion: storedVersion_(v('Plan_Version')), version: storedVersion_(v('Follow_Up_Version')),
      decision: String(v('Decision') || ''), type: String(v('Commitment_Type') || ''), deliverable: String(v('Deliverable') || ''),
      owner: String(v('Owner_Email') || ''), verifier: String(v('Verifier_Email') || ''),
      originalDue: registryCalendar_(v('Original_Due_Date')), due: registryCalendar_(v('Due_Date')),
      nextReviewDate: registryCalendar_(v('Next_Review_Date')), status: String(v('Status') || ''),
      notes: String(v('Update_Notes') || ''), evidenceUrls: String(v('Evidence_URLs') || '').split('\n').filter(Boolean),
      submittedAt: dateToClientIso_(v('Submitted_At')), verificationDue: registryCalendar_(v('Verification_Due_Date')),
      closedAt: dateToClientIso_(v('Closed_At')), extensionDate: registryCalendar_(v('Extension_Requested_Date')),
      extensionReason: String(v('Extension_Reason') || ''), event: String(v('Event_Type') || ''),
      actor: String(v('Actor_Email') || ''), eventAt: dateToClientIso_(v('Event_At')) };
    if (latest[id] && item.version <= latest[id].version) throw new Error('Follow_Up_Log contains duplicate or unordered versions. Ask the maintainer to review it.');
    latest[id] = item; history.push(item);
  });
  return { items: Object.keys(latest).map(function(id) { return latest[id]; }), history: history };
}
function registryNotificationRows_() {
  if (!getTrackerSpreadsheet_().getSheetByName('Notification_Log')) return [];
  const t = readTable_('Notification_Log', REGISTRY_NOTIFICATION_HEADERS);
  return t.rows.filter(function(row) { return row[t.map.notification_id]; }).map(function(row) {
    let targets; try { targets = JSON.parse(String(row[t.map.targets_json] || '[]')); } catch (e) { throw new Error('Notification_Log has invalid target data. Reminder processing is blocked.'); }
    if (!Array.isArray(targets)) throw new Error('Notification_Log has invalid target data.');
    return { id: String(row[t.map.notification_id]), recipient: String(row[t.map.recipient_email]),
      status: String(row[t.map.status]), targets: targets, sentAt: dateToClientIso_(row[t.map.sent_at]),
      preparedAt: dateToClientIso_(row[t.map.prepared_at]), error: String(row[t.map.error] || '') };
  });
}

function registryAcknowledgedNotificationIds_(notifications) {
  const ids=new Set();
  notifications.filter(function(n){return n.status==='EXCEPTION_ACKNOWLEDGED';}).forEach(function(n){
    n.targets.forEach(function(t){if(t.notificationId)ids.add(String(t.notificationId));});
  });
  return ids;
}
function acknowledgeRegistryNotificationFailures(notificationIds, notes, requestToken) {
  return registryAuditOperation_('acknowledgeRegistryNotificationFailures','',function(){
    const actor=requireActiveActor_('canAdmin');
    if(!Array.isArray(notificationIds) || !notificationIds.length || notificationIds.length>100)throw new Error('Select between 1 and 100 failed notices to clear.');
    const ids=notificationIds.map(function(id){return boundedText_(id,'Notification ID',100,true);});
    if(new Set(ids).size!==ids.length)throw new Error('Notification IDs must be unique.');
    ids.sort();notes=boundedText_(notes,'Clearance notes',3000,true);requestToken=normalizeRequestToken_(requestToken);
    const lock=LockService.getScriptLock();lock.waitLock(30000);let op=null,table=null,row=null;
    try{
      const notifications=registryNotificationRows_();
      ids.forEach(function(id){
        const matches=notifications.filter(function(n){return n.id===id;});
        if(matches.length!==1 || matches[0].status!=='FAILED')throw new Error('Only recorded FAILED notices can be cleared. Uncertain sends require reconciliation.');
      });
      op=beginIdempotentOperation_('CHAT_CLEAR_FAILURES',requestToken,digest_(JSON.stringify([ids,notes])));if(op.duplicate)return op.result;
      const acknowledged=registryAcknowledgedNotificationIds_(notifications),pending=ids.filter(function(id){return !acknowledged.has(id);});
      if(pending.length){
        table=registryEnsureLog_('Notification_Log',REGISTRY_NOTIFICATION_HEADERS);row=table.sheet.getLastRow()+1;
        appendNamedLog_(table,{Notification_ID:'ACK-'+Utilities.getUuid(),Recipient_Email:actor.email,Subject:'Failed notice display acknowledged',
          Targets_JSON:JSON.stringify(pending.map(function(id){return {notificationId:id};})),Status:'EXCEPTION_ACKNOWLEDGED',Prepared_At:new Date(),
          Error:'Cleared from active view by '+actor.email+': '+notes,Run_ID:requestToken});SpreadsheetApp.flush();
      }
      const result={success:true,clearedCount:pending.length};completeIdempotentOperation_(op.key,result);row=null;return result;
    }catch(error){rollbackRegistryWrite_([],table,row,op,error);throw error;}
    finally{lock.releaseLock();}
  },true,false);
}
function registryChatCardText_(value, limit) {
  return String(value==null?'':value).replace(/[\u0000-\u001f\u007f]/g,' ').replace(/\s+/g,' ').trim().slice(0,limit || 400)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function registryChatPayload_(digest, plan) {
  const kind=plan.kind || 'EVENTS',entries=digest.entries || [],sections=[];
  const title=kind==='ASSIGNMENT'?'New commitments assigned':kind==='REMINDER'?'Follow-ups requiring action':'Plan update';
  const summary=title+' · '+(entries.length ? entries.length+' commitment(s)' : registryChatText_(digest.subject,120));
  function field(label,value,limit){return {decoratedText:{topLabel:label,text:registryChatCardText_(value,limit),wrapText:true}};}
  entries.slice(0,10).forEach(function(x){
    const widgets=[field('Status / decision',registryEmailStatus_(x)+' · '+x.decision+' · '+x.type,180),
      field('Accountable owner',x.responsibleName+' ('+x.responsible+')',240),field('Due',registryEmailDate_(x.due),60),
      field('Next step',x.deliverable,400)];
    if(x.escalation)widgets.push(field('Escalation attention',x.escalation,400));
    if(x.helpNeeded)widgets.push(field('Support needed',x.helpNeeded,300));
    sections.push({header:registryChatCardText_(x.name,150)+' · '+registryChatCardText_(x.actionPlanId,100),widgets:widgets});
  });
  if(!entries.length){
    const body=String(digest.body || '').replace(/<https:\/\/script\.google\.com\/[^>]+\|Open Action Registry>/g,'');
    const text=body.split('\n').filter(function(line){return line.trim();}).slice(0,18).map(function(line){return registryChatCardText_(line,350);}).join('<br>');
    sections.push({header:'Latest saved update',widgets:[{textParagraph:{text:text || 'Open the registry for the saved update.'}}]});
  }
  const footer=[{textParagraph:{text:'Update progress or review evidence in the registry. Completion and demonstrated impact are assessed separately.'}}];
  const url=registryOfficialAppUrl_();
  if(url)footer.push({buttonList:{buttons:[{text:'Open Action Registry',onClick:{openLink:{url:url}}}]}});
  sections.push({widgets:footer});
  return {text:'Shopify Global Accounts | Action Registry\n'+summary,
    cardsV2:[{cardId:'action-registry',card:{header:{title:'Action Registry',subtitle:registryChatCardText_(title+' · '+registryEmailDate_(plan.today || registryToday_()),150)},sections:sections}}]};
}

function registryFollowUpHealth_(item, notifications, today) {
  const closed = item.status === 'Closed' || item.status === 'Withdrawn';
  const phase = item.status === 'Submitted for Verification' ? 'verifier' : 'owner';
  const responsible = phase === 'verifier' ? item.verifier : item.owner;
  const deadline = phase === 'verifier' ? item.verificationDue : item.due;
  let strikes = 0, lastSent = '', uncertain = false;
  notifications.forEach(function(n) {
    n.targets.forEach(function(target) {
      if (target.id !== item.id || target.phase !== phase || target.kind !== 'OVERDUE' || target.role !== 'accountable') return;
      if (n.status === 'SENDING') uncertain = true;
      if (n.status === 'SENT') { strikes = Math.max(strikes, Number(target.strike) || 0); if (n.sentAt > lastSent) lastSent = n.sentAt; }
    });
  });
  const overdue = !closed && !!deadline && deadline < today;
  const daysOverdue = overdue ? Math.floor((Date.parse(today) - Date.parse(deadline)) / 86400000) : 0;
  return Object.assign({}, item, { phase: phase, responsible: responsible, deadline: deadline, overdue: overdue,
    daysOverdue: daysOverdue, strikes: Math.min(3, strikes), escalated: overdue && strikes >= 3,
    lastReminderAt: lastSent, notificationUncertain: uncertain,
    health: closed ? item.status : (overdue ? (strikes >= 3 ? 'Escalated' : 'Overdue') : (deadline <= registryAddDays_(today, 1) ? 'Due soon' : 'On track')) });
}
function getRegistryFollowUpPortfolio() {
  const actor = requireActiveActor_(), raw = registryFollowUpRows_(), notifications = registryNotificationRows_();
  const master = readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS));
  const byId = {};
  master.rows.forEach(function(row) { const id = String(row[master.map['action plan id']] || '').trim(); if (id) byId[id] = row; });
  const today = registryToday_();
  const items = raw.items.map(function(item) {
    const x = registryFollowUpHealth_(item, notifications, today), row = byId[item.actionPlanId];
    x.name = row ? String(row[master.map['action plan name']] || '') : item.actionPlanId;
    x.site = row ? normalizeActionTrackerSite_(row[master.map.site]) : 'Unknown';
    x.earlierPlan = !row || storedVersion_(row[master.map.plan_version]) !== item.planVersion;
    x.canUpdate = actor.email === item.owner && !['Closed', 'Withdrawn', 'Submitted for Verification'].includes(item.status);
    x.canVerify = actor.canAudit && actor.email === item.verifier && actor.email !== item.owner;
    x.history = raw.history.filter(function(e) { return e.id === item.id; }).reverse();
    return x;
  });
  const linked = {}; items.forEach(function(x) { linked[x.assessmentId] = true; });
  let untrackedAssessments = 0;
  if (getTrackerSpreadsheet_().getSheetByName(REGISTRY_ASSESSMENT_SHEET)) {
    const table = readTable_(REGISTRY_ASSESSMENT_SHEET, REGISTRY_ASSESSMENT_HEADERS);
    untrackedAssessments = table.rows.filter(function(r) { return r[table.map.assessment_id] && !linked[String(r[table.map.assessment_id])]; }).length;
  }
  const routing = registryReminderRouting_();
  const routingExceptions = routing.errors.slice();
  items.filter(function(item) { return !['Closed', 'Withdrawn'].includes(item.status) && !routing.routes[item.responsible]; }).forEach(function(item) { routingExceptions.push(item.id + ': missing manager/executive mapping for ' + item.responsible); });
  const acknowledgedNotifications=registryAcknowledgedNotificationIds_(notifications);
  return { items: items, today: today, mode: registryReminderSettings_().mode, channel: 'GOOGLE_CHAT', webhookConfigured: !!registryChatWebhookUrl_(false), eventScheduleRecorded: !!PropertiesService.getScriptProperties().getProperty('REGISTRY_CHAT_TRIGGER_OWNER'), dailyScheduleRecorded: !!PropertiesService.getScriptProperties().getProperty('FOLLOWUP_TRIGGER_OWNER'), untrackedAssessments: untrackedAssessments, routingExceptions: routingExceptions,
    acknowledgedFailureCount: notifications.filter(function(n){return n.status==='FAILED' && acknowledgedNotifications.has(n.id);}).length,
    notificationExceptions: notifications.filter(function(n) { return n.status === 'SENDING' || (n.status === 'FAILED' && !acknowledgedNotifications.has(n.id)); }).map(function(n) {
      const target=n.targets.find(function(t){return t.eventId && t.kind==='EVENT';});
      return { id: n.id, status: n.status, recipient: n.recipient, error: n.error, preparedAt: n.preparedAt, eventId: target ? target.eventId : '' };
    }) };
}
function registryFollowUpFields_(item) {
  return { Follow_Up_ID: item.id, Assessment_ID: item.assessmentId, Action_Plan_ID: item.actionPlanId,
    Plan_Version: item.planVersion, Follow_Up_Version: item.version + 1, Decision: item.decision, Commitment_Type: item.type,
    Deliverable: item.deliverable, Owner_Email: item.owner, Verifier_Email: item.verifier,
    Original_Due_Date: item.originalDue, Due_Date: item.due, Next_Review_Date: item.nextReviewDate,
    Status: item.status, Update_Notes: item.notes, Evidence_URLs: item.evidenceUrls.join('\n'),
    Submitted_At: item.submittedAt ? new Date(item.submittedAt) : '', Verification_Due_Date: item.verificationDue,
    Closed_At: item.closedAt ? new Date(item.closedAt) : '', Extension_Requested_Date: item.extensionDate,
    Extension_Reason: item.extensionReason };
}
function submitRegistryFollowUpEventCore_(followUpId, payload, expectedVersion, requestToken) {
  followUpId = requireNonBlankText_(followUpId, 'Follow-up ID'); requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let operation = null, table = null, row = null;
  try {
    const actor = requireActiveActor_();
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('Follow-up details are required.');
    const event = registryAssessmentChoice_(payload.event, 'follow-up operation', ['Update', 'Submit', 'Verify', 'Return', 'Request Extension', 'Approve Extension', 'Reject Extension', 'Reopen', 'Withdraw']);
    const item = registryFollowUpRows_().items.find(function(x) { return x.id === followUpId; });
    if (!item) throw new Error('Follow-up not found.');
    const ownerOp = ['Update', 'Submit', 'Request Extension'].includes(event);
    if (ownerOp ? actor.email !== item.owner : (!actor.canAudit || actor.email !== item.verifier || actor.email === item.owner)) throw registrySecurityError_('Only the assigned owner or authorized independent verifier can perform this operation.', 'FOLLOWUP_ASSIGNEE_REQUIRED');
    operation = beginIdempotentOperation_('FOLLOWUP', requestToken, digest_(JSON.stringify([followUpId, payload, expectedVersion])));
    if (operation.duplicate) return operation.result;
    requireVersion_(expectedVersion, item.version);
    const fields = registryFollowUpFields_(item), now = new Date(), today = registryToday_(now);
    const closed = ['Closed', 'Withdrawn'].includes(item.status), submitted = item.status === 'Submitted for Verification';
    if (ownerOp && (closed || submitted)) throw new Error('This commitment is closed or waiting for verification.');
    if (closed && event !== 'Reopen') throw new Error('Reopen the commitment before updating it.');
    const note = boundedText_(payload.notes, 'Progress / verification notes', 3000, true);
    fields.Update_Notes = note;
    if (event === 'Update') fields.Status = registryAssessmentChoice_(payload.status, 'progress status', ['In Progress', 'Blocked']);
    if (event === 'Request Extension') {
      fields.Extension_Requested_Date = registryAssessmentDate_(payload.extensionDate, 'Requested completion date');
      if (fields.Extension_Requested_Date <= item.due || fields.Extension_Requested_Date < today) throw new Error('Requested date must extend the current deadline and be today or later.');
      fields.Extension_Reason = note;
    }
    if (event === 'Approve Extension' || event === 'Reject Extension') {
      if (!item.extensionDate) throw new Error('No extension request is pending.');
      if (event === 'Approve Extension') {
        if (item.extensionDate < today) throw new Error('The requested date has passed. Return for a new extension request.');
        fields.Due_Date = item.extensionDate;
      }
      fields.Extension_Requested_Date = ''; fields.Extension_Reason = '';
    }
    if (event === 'Submit') {
      const links = boundedText_(payload.evidenceUrls, 'Completion evidence links', 10000, true).split(/\r?\n/).map(function(x) { return x.trim(); }).filter(Boolean);
      if (!links.length || links.length > 5 || links.some(function(x) { return !isRegistryGlideUrl_(x); })) throw new Error('Use 1–5 reviewer-accessible Google Drive, Docs, Sheets, or Slides evidence links.');
      fields.Evidence_URLs = Array.from(new Set(links)).join('\n'); fields.Status = 'Submitted for Verification';
      fields.Submitted_At = now; fields.Verification_Due_Date = registryAddDays_(today, registryReminderSettings_().verificationDays);
    }
    if (event === 'Verify' || event === 'Return') {
      if (!submitted) throw new Error('The owner must submit evidence for verification first.');
      if (event === 'Verify' && (!item.evidenceUrls.length || item.evidenceUrls.some(function(url) { return !isRegistryGlideUrl_(url); }))) throw new Error('Valid completion evidence is required before verified closure.');
      fields.Status = event === 'Verify' ? 'Closed' : 'In Progress';
      if (event === 'Verify') fields.Closed_At = now;
      else { fields.Submitted_At = ''; fields.Verification_Due_Date = ''; }
    }
    if (event === 'Withdraw') { fields.Status = 'Withdrawn'; fields.Closed_At = now; }
    if (event === 'Reopen') {
      if (!closed) throw new Error('Only a closed or withdrawn commitment can be reopened.');
      fields.Status = 'Open'; fields.Closed_At = ''; fields.Submitted_At = ''; fields.Verification_Due_Date = '';
    }

    fields.Event_ID = 'FUE-' + Utilities.getUuid(); fields.Event_Type = event;
    fields.Actor_Email = actor.email; fields.Event_At = now; fields.Request_ID = requestToken;
    table = registryEnsureLog_('Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS); row = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, fields); SpreadsheetApp.flush();
    const result = { success: true, id: item.id, version: item.version + 1, status: fields.Status };
    completeIdempotentOperation_(operation.key, result); row = null; return result;
  } catch (e) { rollbackRegistryWrite_([], table, row, operation, e); throw e; }
  finally { lock.releaseLock(); }
}

function registryReminderRouting_() {
  const routes = {}, errors = [];
  if (!getTrackerSpreadsheet_().getSheetByName('Follow_Up_Routing')) return { routes: routes, errors: ['Follow_Up_Routing is missing. Run setupRegistryFollowUps and populate recipient mappings.'] };
  const t = readTable_('Follow_Up_Routing', REGISTRY_ROUTING_HEADERS);
  t.rows.filter(function(r) { return r[t.map.active] === true; }).forEach(function(r) {
    const owner = String(r[t.map.owner_email] || '').trim().toLowerCase();
    if (routes[owner]) throw new Error('Duplicate active owner in Follow_Up_Routing.');
    try {
      validateTaskUsOwnerEmail_(owner);
      const manager = validateTaskUsOwnerEmail_(r[t.map.manager_email]), executive = validateTaskUsOwnerEmail_(r[t.map.executive_email]);
      [owner, manager, executive].forEach(function(email) { const a = getAccessEntry_(email); if (!a || !a.active) throw new Error('Recipient is not active in Access_Control: ' + email); });
      routes[owner] = { manager: manager, executive: executive };
    } catch (e) { errors.push(owner + ': ' + e.message); }
  });
  return { routes: routes, errors: errors };
}
function registryReminderPlan_(now) {
  return registryChatPlan_('REMINDER', now);
}
function registryEmailDate_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) return String(date || 'Not recorded');
  const parts = date.split('-'), months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return Number(parts[2]) + ' ' + months[Number(parts[1]) - 1] + ' ' + parts[0];
}
function registryEmailSummary_(digest) {
  const entries = digest.entries || [], count = entries.length;
  const noun = count === 1 ? 'follow-up' : 'follow-ups';
  const executive = entries.some(function(x) { return x.role === 'executive'; });
  const manager = entries.some(function(x) { return x.role === 'manager'; });
  const overdue = entries.filter(function(x) { return x.daysOverdue > 0 || x.kind === 'OVERDUE'; }).length;
  const assigned = entries.filter(function(x) { return x.kind === 'ASSIGNMENT'; }).length;
  const dueSoon = entries.filter(function(x) { return x.kind === 'DUE_SOON'; }).length;
  const headline = executive ? 'Executive attention required' : (manager ? 'Manager attention required' :
    (overdue ? 'Overdue follow-ups need action' : (assigned === count ? 'New follow-ups assigned' : 'Upcoming follow-ups need action')));
  const earliest = entries.map(function(x) { return x.due; }).filter(Boolean).sort()[0];
  return { headline: headline, count: count, overdue: overdue, assigned: assigned, dueSoon: dueSoon,
    earliest: earliest, caption: count + ' ' + noun + ' · ' + overdue + ' overdue' + (earliest ? ' · Earliest due ' + registryEmailDate_(earliest) : ''),
    subject: 'Action Registry | ' + (executive ? 'Executive attention' : manager ? 'Manager attention' : overdue ? overdue + ' overdue' : assigned === count ? count + ' new assignment' + (count === 1 ? '' : 's') : count + ' due soon') };
}
function registryEmailStatus_(entry) {
  const overdue = entry.daysOverdue > 0 || entry.kind === 'OVERDUE';
  let label = overdue ? 'Overdue' + (entry.daysOverdue ? ' · ' + entry.daysOverdue + ' day' + (entry.daysOverdue === 1 ? '' : 's') : '') :
    (entry.kind === 'ASSIGNMENT' ? 'New assignment' : 'Due soon');
  if (entry.role === 'executive') label += ' · Executive escalation';
  else if (entry.role === 'manager') label += ' · Manager escalation';
  if (entry.kind === 'OVERDUE' && entry.strike) label += ' · Reminder ' + entry.strike + ' of 3';
  return label;
}
function registryComposeDigest_(group, today, settings) {
  const summary = registryEmailSummary_(group);
  group.subject = summary.subject + ' | ' + registryEmailDate_(today);
  group.body = 'Shopify Global Accounts\n' + summary.headline + '\nAs of ' + registryEmailDate_(today) + ' (Asia/Manila)\n' + summary.caption + '\n\n' + group.entries.map(function(x) {
    return x.name + '\n' + x.actionPlanId + ' · ' + x.decision + ' · ' + x.type + '\n' + registryEmailStatus_(x) +
      '\nAccountable: ' + (x.responsibleName || x.responsible) + (x.responsibleName && x.responsibleName !== x.responsible ? ' (' + x.responsible + ')' : '') +
      '\nDue: ' + registryEmailDate_(x.due) + '\nNext action: ' + x.deliverable +
      (x.phase === 'verifier' ? '\nVerification: Review submitted evidence and confirm completion or return it with guidance.' : '') +
      (x.helpNeeded ? '\nSupport needed: ' + x.helpNeeded : '');
  }).join('\n\n') + '\n\nOpen Effectiveness & Decisions to update progress, submit evidence or verify completion.\nCompletion and demonstrated impact are assessed separately.';
  const appUrl = registryOfficialAppUrl_(); if (appUrl) group.body += '\nOpen Action Registry: ' + appUrl;
}

function previewRegistryFollowUpReminders() {
  const actor=requireActiveActor_();
  if (!actor.canAudit && !actor.canAdmin) throw registrySecurityError_('Preview requires Can_Audit or Can_Admin.', 'PERMISSION_REQUIRED');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try { return registryReminderPlan_(new Date()); } finally { lock.releaseLock(); }
}

function installRegistryFollowUpReminderTriggerCore_() {
  const actor = requireActiveActor_('canAdmin');
  const effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (effective !== actor.email) throw registrySecurityError_('Install the trigger from the script editor using your own authorized account.', 'TRIGGER_IDENTITY_REQUIRED');
  setupRegistryFollowUpsCore_();
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'runRegistryFollowUpReminders_'; }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  const p = PropertiesService.getScriptProperties(); p.setProperty('FOLLOWUP_TRIGGER_OWNER', actor.email);
  const trigger = ScriptApp.newTrigger('runRegistryFollowUpReminders_').timeBased().atHour(9).everyDays(1).inTimezone('Asia/Manila').create();
  return { success: true, mode: registryReminderSettings_().mode, triggerId: trigger.getUniqueId(), window: '09:00–10:00 Asia/Manila daily' };
}
function runRegistryFollowUpRemindersCore_() {

  const p = PropertiesService.getScriptProperties();
  const email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const actor = getAccessEntry_(email);
  if (email !== p.getProperty('FOLLOWUP_TRIGGER_OWNER') || !actor || !actor.active || !actor.canAdmin) throw registrySecurityError_('Reminder trigger owner is not authorized. Check Active and Can_Admin for the recorded trigger owner.', 'TRIGGER_OWNER_REQUIRED');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const plan = registryReminderPlan_(new Date());
    if (plan.mode !== 'LIVE') {   return plan; }
    return registryExecuteReminderPlan_(plan, email);
  } finally { lock.releaseLock(); }
}

function resolveRegistryNotificationCore_(notificationId, outcome, notes) {
  const actor = requireActiveActor_('canAdmin');
  if (!['Confirmed Sent', 'Confirmed Not Sent'].includes(outcome)) throw new Error('Select a confirmed send outcome.');
  notes = boundedText_(notes, 'Reconciliation evidence / notes', 3000, true);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const t = readTable_('Notification_Log', REGISTRY_NOTIFICATION_HEADERS);
    const matches = []; t.rows.forEach(function(r, i) { if (String(r[t.map.notification_id]) === notificationId) matches.push(i); });
    if (matches.length !== 1) throw new Error('Notification not uniquely found.');
    const index = matches[0], data = t.rows[index];
    if (data[t.map.status] !== 'SENDING') throw new Error('Only an uncertain send can be reconciled.');
    const row = index + 2;
    setLiteralCellValue_(t.sheet.getRange(row, t.map.error + 1), 'Reconciled by ' + actor.email + ' on ' + new Date().toISOString() + ': ' + notes);
    if (outcome === 'Confirmed Sent' && !data[t.map.sent_at]) t.sheet.getRange(row, t.map.sent_at + 1).setValue(new Date());
    t.sheet.getRange(row, t.map.status + 1).setValue(outcome === 'Confirmed Sent' ? 'SENT' : 'FAILED');
    SpreadsheetApp.flush(); return { success: true };
  } finally { lock.releaseLock(); }
}

function registryAssessmentDefaults_(action, assessments) {
  const today = registryToday_(), planStartRaw = action.row[action.map['start date']];
  let plannedStart = registryDateInput_(planStartRaw);
  if (!plannedStart && /^\d{4}-\d{2}-\d{2}$/.test(String(planStartRaw || ''))) {
    try { plannedStart = registryAssessmentDate_(planStartRaw, 'Planned start'); } catch (e) {   }
  }
  const latest = (assessments || []).find(function(item) { return item.planVersion === action.planVersion; });
  let periodStart = '', startSource = '', message = '';
  if (latest && /^\d{4}-\d{2}-\d{2}$/.test(latest.periodEnd)) {
    const next = registryAddDays_(latest.periodEnd, 1);
    if (next <= today) { periodStart = next; startSource = 'Day after previous observation end'; }
    else { startSource = 'Previous assessment already covers today or later'; message = 'Select an overlapping past period if needed to assess sustained impact. No new completed observation dates are available.'; }
  } else if (plannedStart && plannedStart <= today) {
    periodStart = plannedStart; startSource = 'Original planned start — confirm against the evidence';
  } else {
    startSource = 'No valid past planned start'; message = 'Enter the actual past period covered by your evidence. A future planned start cannot be an observation date.';
  }
  return { periodStart: periodStart, periodEnd: today, today: today, originalPlannedStart: plannedStart,
    startSource: startSource, endSource: 'Today in Manila — change to evidence data-through date if earlier',
    previousAssessmentId: latest ? latest.assessmentId : '', message: message };
}

function registryExecuteReminderPlan_(plan, replyTo) {
  return registryExecuteChatPlan_(plan);
}

function registryOfficialAppUrl_() {
  const url = registryReminderSettings_().appUrl;
  return /^https:\/\/script\.google\.com\/(?:a\/macros\/taskus\.com\/|macros\/)s\/[A-Za-z0-9_-]+\/exec$/.test(url) ? url : '';
}

function registryAssignmentPlan_(now) {
  return registryChatPlan_('ASSIGNMENT', now);
}
function previewRegistryAssignmentEmails() {
  const actor=requireActiveActor_();
  if (!actor.canAudit && !actor.canAdmin) throw registrySecurityError_('Preview requires Can_Audit or Can_Admin.', 'PERMISSION_REQUIRED');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try { return registryAssignmentPlan_(); } finally { lock.releaseLock(); }
}
function setRegistryLiveEmailModeCore_(mode, requestToken) {
  if (!['LIVE', 'PAUSED', 'DEV'].includes(mode)) throw new Error('Select enable or pause Google Chat delivery.');
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let op = null, table = null, row = null, previous = null, changed = false;
  try {
    const actor = requireActiveActor_('canAdmin');
    if (mode === 'LIVE') registryChatWebhookUrl_(true);
    op = beginIdempotentOperation_('CHAT_MODE', requestToken, digest_(mode)); if (op.duplicate) return op.result;

    ensureRegistryAssessmentLog_(); registryEnsureLog_('Follow_Up_Log', REGISTRY_FOLLOWUP_HEADERS);
    registryEnsureLog_('Follow_Up_Routing', REGISTRY_ROUTING_HEADERS);
    table = registryEnsureLog_('Notification_Log', REGISTRY_NOTIFICATION_HEADERS);
    const props = PropertiesService.getScriptProperties(); previous = props.getProperty('REGISTRY_CHAT_MODE');
    row = table.sheet.getLastRow() + 1;
    appendNamedLog_(table, { Notification_ID: 'CFG-' + Utilities.getUuid(), Recipient_Email: actor.email,
      Subject: 'Google Chat delivery ' + (mode === 'LIVE' ? 'enabled' : 'paused') + ' by ' + actor.email, Targets_JSON: '[]', Status: 'CONFIG', Prepared_At: new Date(),
      Error: 'Authorized configuration change; no Chat messages sent.', Run_ID: requestToken });
    props.setProperty('REGISTRY_CHAT_MODE', mode === 'LIVE' ? 'LIVE' : 'PAUSED'); changed = true; SpreadsheetApp.flush();
    const result = { success: true, mode: mode === 'LIVE' ? 'LIVE' : 'PAUSED', dailyScheduleRecorded: !!props.getProperty('FOLLOWUP_TRIGGER_OWNER') };
    completeIdempotentOperation_(op.key, result); row = null; return result;
  } catch (error) {
    if (changed) { const props = PropertiesService.getScriptProperties(); if (previous == null) props.deleteProperty('REGISTRY_CHAT_MODE'); else props.setProperty('REGISTRY_CHAT_MODE', previous); }
    rollbackRegistryWrite_([], table, row, op, error); throw error;
  } finally { lock.releaseLock(); }
}
function sendRegistryFollowUpEmailsNowCore_(kind, requestToken) {
  if (!['ASSIGNMENT', 'REMINDER', 'EVENTS'].includes(kind)) throw new Error('Select plan updates, assignment notices or eligible due reminders.');
  requestToken = normalizeRequestToken_(requestToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000); let op = null, startedBatch = false;
  try {
    const actor = requireActiveActor_('canAdmin');
    if (registryReminderSettings_().mode !== 'LIVE') throw new Error('Resume Google Chat delivery before sending Chat messages.');
    registryChatWebhookUrl_(true);
    op = beginIdempotentOperation_('SEND_CHAT', requestToken, digest_(kind)); if (op.duplicate) return op.result;
    const plan = kind === 'EVENTS' ? registryEventChatPlan_() : kind === 'ASSIGNMENT' ? registryAssignmentPlan_() : registryReminderPlan_();
    startedBatch = true;
    const result = registryExecuteReminderPlan_(plan, actor.email);
    completeIdempotentOperation_(op.key, result); return result;
  } catch (error) {

    if (op && !op.duplicate) { if (startedBatch) markOperationNeedsReview_(op.key, { error: error.message }); else abandonIdempotentOperation_(op.key); }
    throw error;
  } finally { lock.releaseLock(); }
}

function registryChatWebhookUrl_(required) {
  const value = String(PropertiesService.getScriptProperties().getProperty('REGISTRY_CHAT_WEBHOOK_URL') || '').trim();
  const valid = /^https:\/\/chat\.googleapis\.com\/v1\/spaces\/[A-Za-z0-9_-]+\/messages\?(?:key=[A-Za-z0-9_-]+&token=[A-Za-z0-9_-]+|token=[A-Za-z0-9_-]+&key=[A-Za-z0-9_-]+)$/.test(value);
  if (!valid && required) throw new Error('Configure a valid REGISTRY_CHAT_WEBHOOK_URL in Script Properties. Webhook credentials are not shown in the app or logs.');
  return valid ? value : '';
}
function registryChatText_(value, limit) {
  return String(value == null ? '' : value).replace(/[<>*`~]/g, '').replace(/\s+/g, ' ').trim().slice(0, limit || 400);
}
function registryChatBody_(digest, today, kind) {
  const heading = kind === 'ASSIGNMENT' ? 'New commitments assigned' : 'Follow-ups requiring action';
  const body = ['Shopify Global Accounts | Action Registry', heading + ' · ' + registryEmailDate_(today),
    'Shared Chat space · ' + digest.entries.length + ' commitment(s)', ''];
  digest.entries.forEach(function(x) {
    body.push(registryChatText_(x.name, 150) + ' · ' + registryChatText_(x.actionPlanId, 100));
    body.push(registryChatText_(x.decision, 80) + ' · ' + registryChatText_(x.type, 80) + ' · ' + registryEmailStatus_(x));
    body.push('Accountable: ' + registryChatText_(x.responsibleName, 100) + ' (' + registryChatText_(x.responsible, 180) + ')');
    body.push('Due: ' + registryEmailDate_(x.due) + ' · Next action: ' + registryChatText_(x.deliverable, 400));
    if (x.escalation) body.push('Escalation attention: ' + registryChatText_(x.escalation, 400));
    if (x.helpNeeded) body.push('Support needed: ' + registryChatText_(x.helpNeeded, 300));
    body.push('');
  });
  body.push('Update progress, submit evidence or verify completion in the registry. Completion and demonstrated impact are assessed separately.');
  const url = registryOfficialAppUrl_();
  if (url) body.push('<' + url + '|Open Action Registry>');
  return body.join('\n');
}
function registryChatPlan_(kind, now) {
  now = now || new Date();
  const today = registryToday_(now), notifications = registryNotificationRows_(), routing = registryReminderRouting_();
  const settings = registryReminderSettings_(), errors = routing.errors.slice(), entries = [];
  const master = readTable_(MASTER_SHEET_NAME, getRequiredMasterHeaders_()), names = {};
  master.rows.forEach(function(row) { names[String(row[master.map['action plan id']])] = String(row[master.map['action plan name']] || ''); });
  registryFollowUpRows_().items.forEach(function(raw) {
    const x = registryFollowUpHealth_(raw, notifications, today);
    if (['Closed', 'Withdrawn'].includes(x.status)) return;
    const actor = getAccessEntry_(x.responsible);
    if (!actor || !actor.active || (x.phase === 'verifier' && !actor.canAudit)) { errors.push(x.id + ': accountable account is inactive or lacks verification access.'); return; }
    if (!x.deadline || !/^\d{4}-\d{2}-\d{2}$/.test(x.deadline)) { errors.push(x.id + ': missing or invalid deadline.'); return; }
    const related = notifications.filter(function(n) { return n.targets.some(function(t) { return t.id === x.id && t.phase === x.phase && t.role === 'accountable'; }); });
    if (related.some(function(n) { return n.status === 'SENDING'; })) { errors.push(x.id + ': uncertain Chat/send outcome requires reconciliation before retrying.'); return; }
    if (related.some(function(n) { return n.status === 'SENT' && Date.parse(n.sentAt) && now.getTime() - Date.parse(n.sentAt) < 86400000; })) return;
    let noticeKind = 'ASSIGNMENT', strike = 0;
    if (kind === 'ASSIGNMENT') {
      if (related.some(function(n) { return n.status === 'SENT' && n.targets.some(function(t) { return t.id === x.id && t.phase === x.phase && t.kind === 'ASSIGNMENT' && t.role === 'accountable'; }); })) return;
    } else {
      noticeKind = x.overdue ? 'OVERDUE' : (x.deadline <= registryAddDays_(today, 1) ? 'DUE_SOON' : '');
      if (!noticeKind) return;
      if (noticeKind === 'DUE_SOON' && related.some(function(n) { return n.status === 'SENT' && n.targets.some(function(t) { return t.id === x.id && t.phase === x.phase && t.kind === noticeKind && t.due === x.deadline; }); })) return;
      strike = x.overdue ? Math.min(3, x.strikes + 1) : 0;
    }
    const target = { id: x.id, phase: x.phase, kind: noticeKind, due: x.deadline, strike: strike, role: 'accountable' };
    const targets = [target], attention = [], route = routing.routes[x.responsible];
    if (strike >= 2) {
      if (!route) errors.push(x.id + ': add active manager/executive routing for escalation attention.');
      else {
        attention.push('Manager ' + route.manager); targets.push(Object.assign({}, target, {role:'manager'}));
        if (strike >= 3) { attention.push('Executive ' + route.executive); targets.push(Object.assign({}, target, {role:'executive'})); }
      }
    }
    entries.push({targets:targets, entry:{actionPlanId:x.actionPlanId, name:names[x.actionPlanId] || x.actionPlanId,
      decision:x.decision, type:x.type, responsible:x.responsible, responsibleName:actor.name || x.responsible,
      due:x.deadline, daysOverdue:x.daysOverdue, deliverable:x.deliverable, phase:x.phase,
      kind:noticeKind, strike:strike, role:'accountable', escalation:attention.join('; '),
      helpNeeded:x.status === 'Blocked' || x.extensionDate ? x.notes : ''}});
  });
  entries.sort(function(a,b) { return b.entry.daysOverdue-a.entry.daysOverdue || a.entry.due.localeCompare(b.entry.due); });
  const digests = [];

  if (entries.length > 100) errors.push((entries.length-100) + ' commitment(s) deferred to the next run (100 per run maximum).');
  function pack(batch) {
    const digest={recipient:'Google Chat space',entries:batch.map(function(x){return x.entry;}),
      targets:batch.reduce(function(all,x){return all.concat(x.targets);},[]),subject:'Action Registry | '+(kind==='ASSIGNMENT'?'Assignments':'Due reminders')};
    digest.body=registryChatBody_(digest,today,kind); return digest;
  }
  let batch=[];
  entries.slice(0,100).forEach(function(item) {
    const trial=pack(batch.concat([item]));
    if (batch.length && (batch.length>=10 || Utilities.newBlob(JSON.stringify(registryChatPayload_(trial,{kind:kind,today:today}))).getBytes().length>28000)) {
      digests.push(pack(batch)); batch=[];
    }
    batch.push(item);
  });
  if (batch.length) digests.push(pack(batch));
  if (digests.length>10) {
    const deferred=digests.slice(10).reduce(function(total,d){return total+d.entries.length;},0);
    errors.push(deferred+' commitment(s) deferred to the next run (10 messages per run maximum).');
    digests.splice(10);
  }
  if (!registryChatWebhookUrl_(false)) errors.push('Google Chat webhook is not configured in Script Properties.');
  return {mode:settings.mode,today:today,digests:digests,exceptions:Array.from(new Set(errors)),kind:kind,
    channel:'GOOGLE_CHAT',webhookConfigured:!!registryChatWebhookUrl_(false)};
}

function registryChatDiagnosticText_(value, webhookUrl) {
  let text=String(value == null ? '' : value);
  const secrets=[];
  String(webhookUrl || '').replace(/[?&](?:key|token)=([^&]+)/g,function(match,secret){secrets.push(secret);return match;});
  secrets.forEach(function(secret){
    [secret,encodeURIComponent(secret)].forEach(function(form){if(form)text=text.split(form).join('[redacted]');});
  });
  return text.replace(/https?:\/\/[^\s<>"']+/gi,'[redacted URL]')
    .replace(/\b(?:key|token|authorization|api[_-]?key)\s*[=:]\s*[^\s,;]+/gi,'[redacted credential]')
    .replace(/[A-Za-z0-9_-]{25,}/g,'[redacted identifier]')
    .replace(/[\u0000-\u001f\u007f]/g,' ').replace(/\s+/g,' ').trim().slice(0,800);
}
function registryChatResponseDiagnostic_(response, webhookUrl) {
  try {
    const raw=response.getContentText();
    if(typeof raw!=='string' || raw.length>32768)return 'Google did not return a bounded structured error.';
    const data=JSON.parse(raw), error=data && (data.error || data);
    if(!error || typeof error!=='object' || Array.isArray(error))return 'Google did not return a structured error.';
    const status=typeof error.status==='string' && /^[A-Z][A-Z_]{0,59}$/.test(error.status)?registryChatDiagnosticText_(error.status,webhookUrl):'';
    const message=typeof error.message==='string'?registryChatDiagnosticText_(error.message,webhookUrl):'';
    return [status,message].filter(Boolean).join(': ') || 'Google did not provide an error reason.';
  } catch(error) { return 'Google returned an unreadable error response.'; }
}

function registryExecuteChatPlan_(plan) {
  if (registryReminderSettings_().mode !== 'LIVE' || plan.mode !== 'LIVE') throw new Error('Resume Google Chat delivery before sending messages.');
  const url=registryChatWebhookUrl_(true), table=registryEnsureLog_('Notification_Log',REGISTRY_NOTIFICATION_HEADERS);
  const runId='RUN-'+Utilities.getUuid(), sent=[], failures=[];
  for (let i=0;i<plan.digests.length;i++) {
    const digest=plan.digests[i], payload=JSON.stringify(registryChatPayload_(digest,plan));
    if (Utilities.newBlob(payload).getBytes().length > 30000) { failures.push('Chat message exceeds the size limit; no message was sent for this batch.'); continue; }
    if (i) Utilities.sleep(1100);
    const id='NTF-'+Utilities.getUuid(), row=appendNamedLog_(table,{Notification_ID:id,Recipient_Email:'Google Chat space',
      Subject:digest.subject,Targets_JSON:JSON.stringify(digest.targets),Status:'SENDING',Prepared_At:new Date(),Run_ID:runId});
    SpreadsheetApp.flush();
    let code=0, response=null;
    try {

      response=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json; charset=UTF-8',
        payload:payload,muteHttpExceptions:true,followRedirects:false});
      code=response.getResponseCode();
    } catch (e) {

      setLiteralCellValue_(table.sheet.getRange(row,table.map.error+1),'Chat transport outcome unknown. Confirm in the Chat space before reconciling.');
      SpreadsheetApp.flush(); failures.push(id+': Chat outcome unknown; reconciliation required.'); break;
    }
    if (code<200 || code>=300) {
      const uncertain=code===408 || code>=500;
      table.sheet.getRange(row,table.map.status+1).setValue(uncertain?'SENDING':'FAILED');
      const message='Google Chat HTTP '+code+(uncertain?': acceptance unknown; reconciliation required.':': request rejected.')+' Google response: '+registryChatResponseDiagnostic_(response,url);
      setLiteralCellValue_(table.sheet.getRange(row,table.map.error+1),message); SpreadsheetApp.flush(); failures.push(id+': '+message); break;
    }
    let name='';
    try { const data=JSON.parse(response.getContentText()); if (/^spaces\/[A-Za-z0-9_-]+\/messages\/[A-Za-z0-9._-]+$/.test(data.name || '')) name=data.name; } catch(e) {   }

    table.sheet.getRange(row,table.map.sent_at+1).setValue(new Date());
    setLiteralCellValue_(table.sheet.getRange(row,table.map.error+1),name?'Chat accepted: '+name:'Google Chat accepted the message; read status is not available.');
    table.sheet.getRange(row,table.map.status+1).setValue('SENT'); SpreadsheetApp.flush(); sent.push(id);
  }
  return {success:failures.length===0,recipientCount:sent.length,messageCount:sent.length,notificationIds:sent,
    channel:'GOOGLE_CHAT',exceptions:(plan.exceptions || []).concat(failures)};
}

const REGISTRY_CHAT_OUTBOX_HEADERS = Object.freeze(['Event_ID','Action_Plan_ID','Event_Type','Queued_At','Body','Actor_Email']);

function registryAwaitingAssessment_(row, map, reporting) {
  if (!reporting.assessmentHistoryLoaded) return false;
  const id=String(row[map['action plan id']] || '').trim();let version,record;
  try{version=storedVersion_(row[map.plan_version]);record=storedVersion_(row[map.record_version]);}catch(error){return false;}
  if (!['Ongoing','On Hold','Completed'].includes(String(row[map['action status']] || '').trim())) return false;
  const flag=row[map.approval_required];
  if (flag !== false && flag !== '' && flag != null) {
    const review=reporting.reviews[String(row[map.latest_review_id] || '')];
    if(flag !== true || String(row[map.review_status]) !== 'Approved' || !review || review.actionPlanId !== id ||
      review.planVersion !== version || review.decision !== 'Approved' || !Number.isSafeInteger(review.recordVersionReviewed) ||
      review.recordVersionReviewed < 1 || review.recordVersionReviewed >= record) return false;
  }
  return !reporting.assessments[id+':'+version];
}

function registryQueueCommittedEvent_(type, actionPlanId, token, result) {

  if (!result || result.success !== true) return result;
  let lock=null, acquired=false;
  try {
    actionPlanId=String(actionPlanId || result.actionPlanId || '');
    if (!actionPlanId && (result.followUpId || result.id)) {
      const item=registryFollowUpRows_().items.find(function(x){return x.id===(result.followUpId || result.id);});
      if(item)actionPlanId=item.actionPlanId;
    }
    if(!actionPlanId)throw new Error('No action reference');
    const actor=String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const eventId='EVT-'+digest_(actor+'\n'+type+'\n'+String(token || '')+'\n'+actionPlanId);
    lock=LockService.getScriptLock();lock.waitLock(5000);acquired=true;
    const table=registryEnsureLog_('Registry_Chat_Outbox',REGISTRY_CHAT_OUTBOX_HEADERS);
    if(table.rows.some(function(row){return String(row[table.map.event_id])===eventId;}))return result;
    const action=findActionRecord_(actionPlanId);
    function value(header){return action.row[action.map[normalizeHeaderKey_(header)]];}
    const headings={REGISTERED:'New plan requires review',EXECUTION_UPDATED:'Execution updated',FEEDBACK_ADDED:'Leadership feedback added',
      REVIEW_DECIDED:'Plan review completed',PLAN_REVISED:'Revised plan requires review',IMPACT_ASSESSED:'Effectiveness assessed',FOLLOWUP_UPDATED:'Decision follow-up updated'};
    const body=['Shopify Global Accounts | Action Registry',headings[type] || 'Plan updated',
      registryChatText_(value('Action Plan Name'),150)+' · '+registryChatText_(actionPlanId,100),
      'Scope: '+[value('Site'),value('Department'),value('LOB'),value('Channel')].filter(Boolean).map(function(x){return registryChatText_(x,80);}).join(' · '),
      'Execution: '+registryChatText_(value('Action Status'),80)+' · Review: '+registryChatText_(value('Review_Status') || 'Not Recorded',80),
      'Impact: '+registryChatText_(value('Effectiveness_Status') || 'Not Yet Assessed',80),
      'Plan version: '+action.planVersion+' · Record: '+action.recordVersion,
      registryMeasurementRead_(action.row,action.map).measurementMethod==='Governance outcome'?'Governance outcome: '+registryChatText_(value('Governance_Expected_Outcome'),250)+' · Acceptance criteria: '+registryChatText_(value('Governance_Acceptance_Criteria'),300):'KPI: '+registryChatText_(value('Target KPI / Behavior'),100)+' · Current: '+registryChatText_(action.sheet.getRange(action.rowNumber,action.map['current performance']+1).getDisplayValue(),60)+' · Goal: '+registryChatText_(action.sheet.getRange(action.rowNumber,action.map['end goal / target']+1).getDisplayValue(),60),
      'Action steps (excerpt): '+registryChatText_(value('Action Description'),200)];
    if(type==='IMPACT_ASSESSED')body.push('Decision: '+registryChatText_(result.decision,80));
    if(type==='FOLLOWUP_UPDATED')body.push('Follow-up: '+registryChatText_(result.followUpId || result.id,100)+' · '+registryChatText_(result.status,80));
    const next=type==='REGISTERED'||type==='PLAN_REVISED'?'Assigned approver: review the current proposal.':type==='REVIEW_DECIDED'?
      (result.decision==='Approved'?'Owner: execute the approved steps and record evidence.':result.decision==='Returned for Revision'?'Owner: address review feedback and resubmit.':'Owner: review the decline reason.'):
      type==='IMPACT_ASSESSED'?'Accountable owner: complete the recorded decision commitment.':type==='FEEDBACK_ADDED'?'Owner: read leadership feedback in the registry.':'Review the saved update and next commitment in the registry.';
    body.push('Next action: '+next,'Changed by: '+registryChatText_(actor,180),'Recorded: '+new Date().toISOString());
    const url=registryOfficialAppUrl_();if(url)body.push('<'+url+'|Open Action Registry>');
    appendNamedLog_(table,{Event_ID:eventId,Action_Plan_ID:actionPlanId,Event_Type:type,Queued_At:new Date(),Body:body.join('\n'),Actor_Email:actor});
    SpreadsheetApp.flush();
    return result;
  } catch(e) {

    result.notificationWarning='Saved successfully, but its Chat notice could not be queued. Ask the maintainer to check Registry_Chat_Outbox; do not resubmit the action.';
    try {console.warn('REGISTRY_CHAT_QUEUE_FAILED');}catch(ignored){}
    return result;
  } finally {if(acquired)try{lock.releaseLock();}catch(ignored){}}
}

function registryValidateChatEventSetup_(warnings) {
  const owner=String(PropertiesService.getScriptProperties().getProperty('REGISTRY_CHAT_TRIGGER_OWNER') || '');
  if(getTrackerSpreadsheet_().getSheetByName('Registry_Chat_Outbox'))readTable_('Registry_Chat_Outbox',REGISTRY_CHAT_OUTBOX_HEADERS);
  if(!owner){warnings.push('Plan-update Chat trigger not installed. Run installRegistryChatTrigger_ from the authorized maintainer account.');return {status:'NOT_CONFIGURED'};}
  const actor=getAccessEntry_(owner);
  if(!actor || !actor.active || !actor.canAdmin)throw new Error('Queued Chat trigger owner requires Active and Can_Admin.');
  const effective=String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if(effective!==owner){warnings.push('Run setup validation from the queued Chat trigger owner to confirm its trigger.');return {status:'OWNER_CHECK_REQUIRED'};}
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='runRegistryQueuedChat_';});
  if(triggers.length!==1)throw new Error('Queued Chat owner must have exactly one runRegistryQueuedChat_ trigger.');
  return {status:'INSTALLED'};
}
function registryEventChatPlan_() {
  const plan={mode:registryReminderSettings_().mode,today:registryToday_(),digests:[],exceptions:[],kind:'EVENTS',channel:'GOOGLE_CHAT',webhookConfigured:!!registryChatWebhookUrl_(false)};
  if(!getTrackerSpreadsheet_().getSheetByName('Registry_Chat_Outbox'))return plan;
  const table=readTable_('Registry_Chat_Outbox',REGISTRY_CHAT_OUTBOX_HEADERS), notifications=registryNotificationRows_();
  const seen=new Set();
  table.rows.forEach(function(row){
    const id=String(row[table.map.event_id] || '');if(!id)return;
    if(seen.has(id)){plan.exceptions.push('Duplicate queued event excluded: '+id);return;}seen.add(id);
    const related=notifications.filter(function(n){return n.targets.some(function(t){return t.eventId===id;});});
    if(related.some(function(n){return n.status==='SENT';}))return;
    if(related.some(function(n){return n.status==='SENDING';})){plan.exceptions.push(id+': send outcome unknown; reconcile in reminder controls.');return;}
    let retryIndex=-1;related.forEach(function(n,i){if(n.status==='RETRY_AUTHORIZED')retryIndex=i;});
    if(related.slice(retryIndex+1).filter(function(n){return n.status==='FAILED';}).length>=3){plan.exceptions.push(id+': rejected three times; maintainer investigation and explicit retry required.');return;}
    const body=String(row[table.map.body] || '');
    if(!body || Utilities.newBlob(JSON.stringify({text:body})).getBytes().length>28000){plan.exceptions.push(id+': invalid queued message.');return;}
    plan.digests.push({recipient:'Google Chat space',subject:'Action Registry | '+String(row[table.map.event_type]),body:body,
      targets:[{eventId:id,id:String(row[table.map.action_plan_id]),kind:'EVENT',role:'event',phase:'event',strike:0}]});
  });
  if(plan.digests.length>10){plan.exceptions.push((plan.digests.length-10)+' plan notice(s) deferred to the next run.');plan.digests.splice(10);}
  if(!plan.webhookConfigured)plan.exceptions.push('Configure REGISTRY_CHAT_WEBHOOK_URL in Script Properties.');
  return plan;
}
function previewRegistryPlanEvents() {
  requireActiveActor_();return registryEventChatPlan_();
}
function retryRegistryPlanNotice(eventId, notes, requestToken) {
  return registryAuditOperation_('retryRegistryPlanNotice',eventId,function(){
    const actor=requireActiveActor_('canAdmin');
    eventId=boundedText_(eventId,'Queued event ID',100,true);notes=boundedText_(notes,'Retry reason',3000,true);
    requestToken=normalizeRequestToken_(requestToken);
    const lock=LockService.getScriptLock();lock.waitLock(30000);let op=null;
    try{
      const table=readTable_('Registry_Chat_Outbox',REGISTRY_CHAT_OUTBOX_HEADERS);
      if(table.rows.filter(function(row){return String(row[table.map.event_id])===eventId;}).length!==1)throw new Error('Queued event must exist exactly once.');
      const related=registryNotificationRows_().filter(function(n){return n.targets.some(function(t){return t.eventId===eventId;});});
      if(related.some(function(n){return n.status==='SENDING';}))throw new Error('Reconcile the uncertain send before authorizing a retry.');
      if(related.some(function(n){return n.status==='SENT';}))throw new Error('An accepted notice cannot be resent.');
      if(!related.some(function(n){return n.status==='FAILED';}))throw new Error('This event has no rejected send to retry.');
      op=beginIdempotentOperation_('CHAT_EVENT_RETRY',requestToken,digest_(eventId+'\n'+notes));if(op.duplicate)return op.result;
      const log=registryEnsureLog_('Notification_Log',REGISTRY_NOTIFICATION_HEADERS);
      appendNamedLog_(log,{Notification_ID:'NTF-'+Utilities.getUuid(),Recipient_Email:'Google Chat space',Subject:'Plan notice retry authorized',
        Targets_JSON:JSON.stringify([{eventId:eventId,kind:'EVENT_RETRY',role:'event',phase:'event',strike:0}]),Status:'RETRY_AUTHORIZED',Prepared_At:new Date(),
        Error:'Retry authorized by '+actor.email+': '+notes,Run_ID:requestToken});SpreadsheetApp.flush();
      const result={success:true,eventId:eventId};completeIdempotentOperation_(op.key,result);return result;
    }catch(e){if(op&&!op.duplicate)markOperationNeedsReview_(op.key,{reason:'CHAT_RETRY_NOT_CONFIRMED'});throw e;}
    finally{lock.releaseLock();}
  },true,false);
}

function setupRegistryChatSchedules() {
  const actor=requireActiveActor_('canAdmin');
  const effective=String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if(actor.email!==effective)throw registrySecurityError_('Run setup from your own authorized Apps Script editor account.','TRIGGER_IDENTITY_REQUIRED');
  registryChatWebhookUrl_(true);
  const chat=installRegistryChatTrigger_();
  const reminders=installRegistryFollowUpReminderTrigger_();
  const checks=validateRegistrySetup();
  const result={success:true,chatSchedule:chat,dueReminderSchedule:reminders,checks:checks};

  return result;
}

function installRegistryChatTrigger_() {
  const actor=requireActiveActor_('canAdmin'), email=String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if(actor.email!==email)throw registrySecurityError_('Install from the signed-in authorized maintainer account.','TRIGGER_IDENTITY_REQUIRED');
  registryChatWebhookUrl_(true);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    registryEnsureLog_('Registry_Chat_Outbox',REGISTRY_CHAT_OUTBOX_HEADERS);
    const handler='runRegistryQueuedChat_';
    ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===handler;}).forEach(function(t){ScriptApp.deleteTrigger(t);});
    ScriptApp.newTrigger(handler).timeBased().everyMinutes(15).create();
    PropertiesService.getScriptProperties().setProperty('REGISTRY_CHAT_TRIGGER_OWNER',email);
    return {success:true,handler:handler,minutes:15};
  } finally {lock.releaseLock();}
}
function runRegistryQueuedChat_() {
  const props=PropertiesService.getScriptProperties(),email=String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(),actor=getAccessEntry_(email);
  if(email!==props.getProperty('REGISTRY_CHAT_TRIGGER_OWNER') || !actor || !actor.active || !actor.canAdmin)throw registrySecurityError_('Queued Chat trigger owner is not authorized.','TRIGGER_OWNER_REQUIRED');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    if(registryReminderSettings_().mode!=='LIVE')return {success:true,paused:true};
    const events=registryExecuteChatPlan_(registryEventChatPlan_());
    if(!events.success)return {success:false,events:events,assignmentsDeferred:true};

    if(events.messageCount)Utilities.sleep(1100);
    const assignments=registryExecuteChatPlan_(registryAssignmentPlan_());
    return {success:events.success && assignments.success,events:events,assignments:assignments};
  } finally {lock.releaseLock();}
}

const REGISTRY_STABILITY_SHEET = 'Stability_Rules';
const REGISTRY_STABILITY_DEFAULTS = Object.freeze([
  ['Rules_Status','Proposed','Proposed or Approved. Change only after business review of the rules.'],
  ['Rules_Version','DRAFT-1','Visible version of the classification rules.'],
  ['Minimum_Core_Plans',5,'Minimum eligible plans in the four core categories.'],
  ['Dominance_Share',0.5,'Intervention or Escalation share required for dominance (0–1).'],
  ['LockedIn_Mature_Share',0.7,'Minimum Standardization + Scaling share for Locked-In (0–1).'],
  ['Minimum_Validated_Share',0.6,'Minimum core share with recent evidence-backed Validated assessments (0–1).'],
  ['Minimum_Sustained_Validated_Share',0.5,'Minimum core share with recent Validated AND Sustained assessments (0–1).'],
  ['Evidence_Max_Age_Days',90,'Maximum age of assessment observation end, in Manila calendar days.'],
  ['Completed_Lookback_Days',90,'Include approved Completed plans only within this many Manila calendar days.']
]);
const REGISTRY_SPONSOR_FRAMEWORK = Object.freeze([
  ['Intervention','Behavior Realignment','Direct, tactical engagements focused on correcting individual or squad-level execution variance.','1:1 Targeted Coaching; QA Side-by-Side Shadowing; Outlier Performance Plan; Disciplinary Action / Removals'],
  ['Intervention','Capability Building','Rapid-response learning interventions deployed when knowledge or skill gaps are identified as root drivers.','Refresher Training Blitz; Knowledge Packet Delivery; Nested Incubation Support; Huddle Upskilling'],
  ['Intervention','Tactical Capacity & Operational Control','Short-term workforce or queue management adjustments to protect service levels and prevent burnt-out operations.','Concurrency Override; Mandatory Overtime / Shrinkage Clawback; Temporary Channel Floor Shifting; Queue Threshold Triggering'],
  ['Escalation','Product & Bug Escalation (For Long standing issues only)','Formal friction logs and tickets raised to client product or engineering teams when systems or software break.','Product Incident Ticket; Merchant Journey Friction Report; System Outage Alerting'],
  ['Escalation','Policy & Knowledge Clarification','Requests sent to client stakeholders to resolve ambiguous policies, broken macros, or conflicting guidelines.','Escalation for Policy Clarification; Macro/Help Center Article Gap Request; Edge Case Guidance Request'],
  ['Escalation','Resource & Cross-Functional Request','Formal appeals for executive decisions, additional budget, headcount, or cross-department intervention.','Leadership Attention Request; Capacity Expansion Proposal; Cross-Functional SLA Escalation'],
  ['Standardization','SOP & Workflow Institutionalization','Updating core operating procedures and training materials so improvements survive past initial focus periods.','Knowledge Base Article / Macro Update; SOP / Workflow Redesign; QA Rubric / Calibration Standard Adjustment'],
  ['Standardization','Baseline Governance & Routine Lock-in','Establishing permanent operational checkpoints and audit mechanisms to ensure execution remains discipline-tight.','Audit Routine Creation; Team Leader Execution Standard Update; WFM Baseline Requirement Lock'],
  ['Standardization','Governance Projects','Launching a new project like WBR streamline or new dashboards.','BIDS, QA, LX, SLT, or Global Ops initiatives that may not be tied directly to performance but improve overall account governance'],
  ['Scaling','Replication & Pilot Rollout','Expanding a locally validated best practice to additional LOBs, squads, or sites under controlled monitoring.','Multi-Site Pilot Launch; LOB Cross-Replication; Best-Practice Playbook Distribution'],
  ['Scaling','Global Institutionalization','Formally adopting a proven local innovation as the global standard across the entire account network.','Global Standard Operating Model Update; Account-Wide Training Curriculum Release; Global Tooling / Automation Rollout'],
  ['Engagement','Incentivized Activity','Activities that increase morale, participation, and drive teams. Complementary, not the main driver for KPIs.','Town Halls; Blitz Incentive; Team Building'],
  ['Engagement','Non - Incentivized Activity','Activities that increase morale, participation, and drive teams. Complementary, not the main driver for KPIs.','']
]);
const REGISTRY_SPONSOR_OPTIONS = Object.freeze({
  Department:['Quality','Training','WFM'],
  Channel:['Chat','Phones','ASYNC','Overall'],
  LOB:['MS Core','MS Core - Premium','Tier 2 - Money/Live Assist/TMS','Plus','Retail','App Review','Biz Ops','Buyer Support','All LOBs'],
  Cohort:['Targeted (Specify in the comments)','0-3 Months','3-6 Months','6-9 Months','9-12 Months','12 Months +','Peak (PACE)','Achieving (PACE)','Concerning (PACE)','Escalated (PACE)','Overall - Teammates','Leadership - Team Leaders','Leadership - Operations Managers','Leadership - SLT']
});
const REGISTRY_SPONSOR_KPIS = Object.freeze(['Smiley','AFTR','Quality','AHT','Productivity (TPH/SPD)','Attrition (Retention)','Shrinkage (OOO)','Interval Compliance','Schedule Adherence','Compliance (CCA/ZTP)','Compliance (FRM)','Hygiene (Aux Toggling, Credits, Cherry Picking etc)']);
const REGISTRY_SPONSOR_INTENTS = Object.freeze({
  Intervention:'Corrective actions to fix performance variance, KPI drops, or execution gaps within existing operational bounds.',
  Escalation:'Upward or cross-functional actions when an issue cannot be resolved within frontline operations alone, including product, policy, or capacity blockers.',
  Standardization:'Embed a verified fix or baseline standard into everyday routines, documentation, or workflows.',
  Scaling:'Replicate a validated local best practice across teams, LOBs, or global sites.',
  Engagement:'Increase morale and participation. Complementary to, not the main driver of, KPI improvement.'
});

function registryAppendConfigRows_(name, headers, rows, keyFields) {
  const table=registryEnsureLog_(name,headers), seen={};
  function key(row){return keyFields.map(function(h){const value=row[table.map[normalizeHeaderKey_(h)]];return name===FRAMEWORK_SHEET_NAME?registryFrameworkLabelKey_(value):String(value || '').trim().toLowerCase();}).join('\n');}
  table.rows.forEach(function(row){seen[key(row)]=true;});
  let added=0;
  rows.forEach(function(values){
    const row=Array(table.headers.length).fill('');
    headers.forEach(function(h,i){row[table.map[normalizeHeaderKey_(h)]]=typeof values[i]==='string'?protectLiteralText_(values[i]):values[i];});
    const k=key(row);if(seen[k])return;
    table.sheet.getRange(table.sheet.getLastRow()+1,1,1,row.length).setValues([row]);seen[k]=true;added++;
  });return added;
}
function getRegistryStabilityConfig() {
  const response={ready:false,status:'Unavailable',version:'',today:registryToday_(),values:{},errors:[]};
  try {
    const table=readTable_(REGISTRY_STABILITY_SHEET,['Key','Value','Meaning']),seen={};
    table.rows.forEach(function(row){const key=String(row[table.map.key] || '').trim();if(!key)return;if(seen[key])throw new Error('Duplicate Stability_Rules key: '+key);seen[key]=true;response.values[key]=row[table.map.value];});
    REGISTRY_STABILITY_DEFAULTS.forEach(function(row){if(!seen[row[0]])throw new Error('Missing Stability_Rules key: '+row[0]);});
    response.status=String(response.values.Rules_Status).trim();response.version=String(response.values.Rules_Version).trim();
    if(!['Proposed','Approved'].includes(response.status) || !response.version)throw new Error('Rules_Status must be Proposed or Approved, with a Rules_Version.');
    ['Minimum_Core_Plans','Evidence_Max_Age_Days','Completed_Lookback_Days'].forEach(function(k){const raw=response.values[k],n=Number(raw);if(!['string','number'].includes(typeof raw) || String(raw).trim()==='' || !Number.isSafeInteger(n) || n<1 || n>3650)throw new Error('Invalid positive integer: '+k);response.values[k]=n;});
    ['Dominance_Share','LockedIn_Mature_Share','Minimum_Validated_Share','Minimum_Sustained_Validated_Share'].forEach(function(k){const raw=response.values[k],n=Number(raw);if(!['string','number'].includes(typeof raw) || String(raw).trim()==='' || !Number.isFinite(n) || n<=0 || n>1)throw new Error('Invalid share (0–1): '+k);response.values[k]=n;});
    if(response.values.Dominance_Share<0.5 || response.values.LockedIn_Mature_Share<0.5 || response.values.Minimum_Sustained_Validated_Share>response.values.Minimum_Validated_Share)throw new Error('Dominance and maturity shares must be at least 0.5; sustained threshold cannot exceed validated threshold.');
    response.ready=true;
  } catch(error){response.errors.push(String(error.message || 'Stability configuration unavailable.'));}
  return response;
}
function setupRegistrySponsorFeatures() {
  return registryAuditOperation_('setupRegistrySponsorFeatures','',function(){
    requireActiveActor_('canAdmin');const lock=LockService.getScriptLock();lock.waitLock(30000);
    try {
      const settingsPreflight=readTable_(SETTINGS_SHEET_NAME,['KPI','Direction','Unit']);
      if(['kpi','direction','unit'].some(function(key,i){return settingsPreflight.map[key]!==i;}))throw new Error('Settings must begin with KPI, Direction, Unit in that order. Correct the configuration layout before setup.');
      getKpiSettings_();
      if(getTrackerSpreadsheet_().getSheetByName(FRAMEWORK_SHEET_NAME))getFrameworkRows_();
      const addedRules=registryAppendConfigRows_(REGISTRY_STABILITY_SHEET,['Key','Value','Meaning'],REGISTRY_STABILITY_DEFAULTS,['Key']);
      const addedFramework=registryAppendConfigRows_(FRAMEWORK_SHEET_NAME,['Category','Type','Definition','Activity_Examples','Active','Sort_Order'],REGISTRY_SPONSOR_FRAMEWORK.map(function(row,i){return row.concat([true,i+1]);}),['Category','Type']);
      const ss=getTrackerSpreadsheet_(),sheet=ss.getSheetByName(SETTINGS_SHEET_NAME);
      if(!sheet || !sheet.getLastRow())throw new Error('Existing Settings sheet with KPI / Direction / Unit headers is required.');
      let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0],map=buildHeaderMapStrict_(headers,SETTINGS_SHEET_NAME),addedOptions={};
      Object.keys(REGISTRY_SPONSOR_OPTIONS).forEach(function(h){
        let index=map[normalizeHeaderKey_(h)];
        if(index===undefined){index=headers.length;if(sheet.getMaxColumns()<=index)sheet.insertColumnsAfter(sheet.getMaxColumns(),index+1-sheet.getMaxColumns());sheet.getRange(1,index+1).setValue(h);headers.push(h);map[normalizeHeaderKey_(h)]=index;}
        const existing=getSettingsOptionValues_(h),seen=new Set(existing.map(function(x){return x.toLowerCase().replace(/\s+/g,' ').replace('in the comment)','in the comments)');}));
        let last=1;if(sheet.getLastRow()>1)sheet.getRange(2,index+1,sheet.getLastRow()-1,1).getDisplayValues().forEach(function(row,i){if(String(row[0] || '').trim())last=i+2;});
        addedOptions[h]=0;REGISTRY_SPONSOR_OPTIONS[h].forEach(function(value){const key=value.toLowerCase().replace(/\s+/g,' ');if(seen.has(key))return;sheet.getRange(++last,index+1).setValue(protectLiteralText_(value));seen.add(key);addedOptions[h]++;});
      });

      const review=registrySponsorConfigurationReview_(),missingKpis=review.missingKpis;
      SpreadsheetApp.flush();
      const stability=getRegistryStabilityConfig();
      const runtimeConfigurationReady=stability.ready && !missingKpis.length && !review.frameworkError && !review.invalidMetricSettings.length && review.frameworkCoverage.activeReferenceTypes===review.frameworkCoverage.expectedTypes;
      const result={success:true,runtimeConfigurationReady:runtimeConfigurationReady,configurationComplete:stability.ready && !missingKpis.length && !review.frameworkDifferences.length && !review.frameworkError && !review.invalidMetricSettings.length,version:ACTIONTRACKER_VERSION,addedRules:addedRules,addedFramework:addedFramework,addedOptions:addedOptions,stability:stability,missingKpis:missingKpis,configurationReview:review,
        warnings:(missingKpis.length?['Add the missing KPI names with the approved Direction and Unit in Settings. Existing KPI configuration was preserved.']:[]).concat(review.frameworkDifferences.length?['Framework reference differences remain; inspect configurationReview.frameworkDifferenceDetails for row numbers, fields and current/reference values. Harmless formatting is normalized; different wording still requires business review. Existing rows were not overwritten.']:[]).concat(review.invalidMetricSettings.length?['Review configurationReview.invalidMetricSettings before using those KPIs.']:[]).concat(review.frameworkError?[review.frameworkError]:[]).concat(stability.errors)};
       return result;
    } finally {lock.releaseLock();}
  },true,false);
}

function registryFrameworkLabelKey_(value) {
  return String(value == null ? '' : value).normalize('NFKC')
    .replace(/[\u2010-\u2015\u2212]/g, '-').replace(/\s+/g, ' ')
    .replace(/\s*-\s*/g, '-').trim().toLowerCase();
}
function registryFrameworkProseKey_(value) {
  return registryFrameworkLabelKey_(String(value == null ? '' : value)
    .replace(/[\u2018\u2019]/g, "'").replace(/[\u201c\u201d]/g, '"'))
    .replace(/\s*\/\s*/g, '/').replace(/[.]$/, '').trim();
}
function registryFrameworkExamplesKey_(value) {
  return String(value == null ? '' : value).split(/[;\r\n\u2022\u00b7\u25cf\u25aa]+/)
    .map(function(item) { return registryFrameworkProseKey_(item.replace(/^\s*[-*]\s+/, '')); })
    .filter(Boolean).sort().join('\n');
}
function registrySponsorConfigurationReview_() {
  let frameworkError = '';
  try { getFrameworkRows_(); } catch (error) { frameworkError = String(error.message); }
  const kpis = getKpiSettings_();
  const table = readTable_(FRAMEWORK_SHEET_NAME, ['Category','Type','Definition','Activity_Examples','Active','Sort_Order']);
  const differences = [], details = [], duplicates = [], grouped = Object.create(null);
  table.rows.forEach(function(row, index) {
    const category = registryFrameworkLabelKey_(row[table.map.category]);
    const type = registryFrameworkLabelKey_(row[table.map.type]);
    if (!category || !type) return;
    const key = category + '\n' + type;
    (grouped[key] || (grouped[key] = [])).push({ row: row, sheetRow: index + 2 });
  });
  Object.keys(grouped).forEach(function(key) {
    const matches = grouped[key];
    if (matches.length > 1) duplicates.push({
      category: String(matches[0].row[table.map.category]),
      type: String(matches[0].row[table.map.type]),
      sheetRows: matches.map(function(item) { return item.sheetRow; }),
      activeRows: matches.filter(function(item) { return item.row[table.map.active] === true; })
        .map(function(item) { return item.sheetRow; })
    });
  });
  let activeReferenceTypes = 0;
  REGISTRY_SPONSOR_FRAMEWORK.forEach(function(reference) {
    const key = registryFrameworkLabelKey_(reference[0]) + '\n' + registryFrameworkLabelKey_(reference[1]);
    const matches = grouped[key] || [];
    const active = matches.filter(function(item) { return item.row[table.map.active] === true; });
    if (active.length === 1) activeReferenceTypes++;

    const match = active[0] || matches[0];
    const issues = [];
    if (!match) {
      issues.push({ field: 'Category / Type', reason: 'Missing reference type',
        current: '', expected: reference[0] + ' / ' + reference[1] });
    } else {
      if (!active.length) issues.push({ field: 'Active', reason: 'Reference type has no active row',
        current: match.row[table.map.active], expected: true });
      if (active.length > 1) issues.push({ field: 'Active', reason: 'Multiple equivalent active rows',
        current: active.map(function(item) { return item.sheetRow; }), expected: 'One active row' });
      [['Definition', 'definition', reference[2], registryFrameworkProseKey_],
       ['Activity_Examples', 'activity_examples', reference[3], registryFrameworkExamplesKey_]]
        .forEach(function(check) {
          const current = String(match.row[table.map[check[1]]] || '');
          if (check[3](current) !== check[3](check[2])) issues.push({
            field: check[0], reason: 'Wording differs after formatting normalization; review business meaning',
            current: current.slice(0, 4000), currentTruncated: current.length > 4000, expected: check[2]
          });
        });
    }
    if (issues.length) {
      differences.push(reference[0] + ' / ' + reference[1]);
      details.push({ category: reference[0], type: reference[1],
        sheetRows: matches.map(function(item) { return item.sheetRow; }),
        comparedSheetRow: match ? match.sheetRow : null, issues: issues });
    }
  });
  return { frameworkError: frameworkError,
    missingKpis: REGISTRY_SPONSOR_KPIS.filter(function(k) { return !resolveRegistryKpiSetting_(kpis, k); }),
    frameworkDifferences: differences, frameworkDifferenceDetails: details,
    frameworkEquivalentRows: duplicates,
    frameworkCoverage: { expectedTypes: REGISTRY_SPONSOR_FRAMEWORK.length, activeReferenceTypes: activeReferenceTypes },
    invalidMetricSettings: Object.keys(kpis).filter(function(k) {
      return !/^(higher|lower)( is better)?$/i.test(kpis[k].direction) || !String(kpis[k].unit || '').trim();
    }).map(function(k) { return kpis[k].name; }) };
}

function registryStabilityFacts_(row,map,reporting) {
  const id=String(row[map['action plan id']] || '').trim();let version,record;
  try{version=storedVersion_(row[map.plan_version]);record=storedVersion_(row[map.record_version]);}catch(error){return {approved:false,assessment:null,issue:'Invalid version metadata'};}
  const flag=row[map.approval_required];
  if(flag!==true && flag!==false && flag!=='' && flag!=null)return {approved:false,assessment:null,issue:'Invalid approval metadata'};
  if(flag===true && reporting.reviewHistoryLoaded===false)return {approved:false,assessment:null,issue:'Review history unavailable'};
  const review=reporting.reviews[String(row[map.latest_review_id] || '').trim()];
  const approved=row[map.approval_required]===true && String(row[map.review_status])==='Approved' && !!review && review.actionPlanId===id && review.planVersion===version && review.decision==='Approved' && Number.isSafeInteger(review.recordVersionReviewed) && review.recordVersionReviewed>=1 && review.recordVersionReviewed<record;
  if(flag===true && String(row[map.review_status])==='Approved' && !approved)return {approved:false,assessment:null,issue:'Approved plan has no matching current-version review'};
  const assessment=(reporting.assessmentDetails || {})[String(row[map.latest_assessment_id] || '')];
  const confirmed=!!assessment && assessment.actionPlanId===id && assessment.planVersion===version && Number.isSafeInteger(assessment.recordVersionReviewed) && assessment.recordVersionReviewed>=1 && assessment.recordVersionReviewed<record && assessment.status===String(row[map.effectiveness_status]);
  return {approved:approved,assessment:confirmed?{status:assessment.status,periodEnd:assessment.periodEnd,sustained:assessment.sustained,hasEvidence:assessment.hasEvidence,expectations:assessment.expectations,relationship:assessment.relationship}:null};
}

function registryStabilityEvidenceIndex_() {
  const index=Object.create(null),seen=Object.create(null);
  if(!getTrackerSpreadsheet_().getSheetByName(EVIDENCE_LOG_SHEET_NAME))return index;
  const table=readTable_(EVIDENCE_LOG_SHEET_NAME,ACTIONTRACKER_EVIDENCE_HEADERS);
  table.rows.forEach(function(row){
    const id=String(row[table.map.evidence_id] || '').trim();if(!id)return;
    if(seen[id]){index[id]=null;return;}seen[id]=true;
    const version=Number(row[table.map.plan_version]),file=String(row[table.map.file_id] || '').trim(),action=String(row[table.map.action_plan_id] || '').trim();
    index[id]=Number.isSafeInteger(version) && version>=1 && /^[A-Za-z0-9_-]+$/.test(file) && action?{actionPlanId:action,planVersion:version}:null;
  });return index;
}

const REGISTRY_GOVERNANCE_HEADERS = Object.freeze(['Measurement_Method','Governance_Current_Problem','Governance_Expected_Outcome','Governance_Acceptance_Criteria','Governance_Verification_Method']);
function registryMeasurementMethod_(value) {
  const method=String(value == null?'':value).trim() || 'Numeric KPI';
  if(!['Numeric KPI','Governance outcome'].includes(method))throw new Error('Select Numeric KPI or Governance outcome.');
  return method;
}
function registryMeasurementFields_(payload) {
  const method=registryMeasurementMethod_(payload.measurementMethod),governance=method==='Governance outcome';
  if(governance && (payload.category!=='Standardization' || payload.actionType!=='Governance Projects'))throw new Error('Governance outcome is available only for Standardization / Governance Projects.');
  const fields={Measurement_Method:method};
  [['Governance_Current_Problem','governanceCurrentProblem','Current governance problem'],['Governance_Expected_Outcome','governanceExpectedOutcome','Expected governance outcome'],['Governance_Acceptance_Criteria','governanceAcceptanceCriteria','Acceptance criteria'],['Governance_Verification_Method','governanceVerificationMethod','Verification method']].forEach(function(pair){fields[pair[0]]=governance?boundedText_(payload[pair[1]],pair[2],3000,true):'';});
  if(governance && [payload.kpi,payload.baseline,payload.goal].some(function(value){return value!=null && String(value).trim()!=='';}))throw new Error('Governance outcomes do not use numeric KPI fields. Clear KPI, baseline and target.');
  return fields;
}
function registryAppendHeaders_(sheet,required) {
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0],map=buildHeaderMapStrict_(headers,sheet.getName());
  const missing=required.filter(function(h){return map[normalizeHeaderKey_(h)]===undefined;});
  if(missing.length){if(sheet.getMaxColumns()<headers.length+missing.length)sheet.insertColumnsAfter(sheet.getMaxColumns(),headers.length+missing.length-sheet.getMaxColumns());sheet.getRange(1,headers.length+1,1,missing.length).setValues([missing]);}
  return missing;
}
function ensureRegistryGovernanceHeaders_() {
  const ss=getTrackerSpreadsheet_(),master=ss.getSheetByName(MASTER_SHEET_NAME),submission=ss.getSheetByName(SUBMISSION_LOG_SHEET_NAME),updates=ss.getSheetByName(UPDATE_LOG_SHEET_NAME);

  readTable_(MASTER_SHEET_NAME,getRequiredMasterHeaders_().concat(ACTIONTRACKER_EXTENSION_HEADERS));
  assertMasterFormulaSchema_(master.getRange(1,1,1,master.getLastColumn()).getDisplayValues()[0]);
  readTable_(SUBMISSION_LOG_SHEET_NAME,getRequiredSubmissionLogHeaders_().filter(function(h){return REGISTRY_SUBMISSION_AUDIT_HEADERS.indexOf(h)<0;}));
  readTable_(UPDATE_LOG_SHEET_NAME,['Update ID','Action Plan ID','Update Date','Updated By','Current Performance','Action Status','Comment']);
  const assessment=ss.getSheetByName(REGISTRY_ASSESSMENT_SHEET);if(assessment)readTable_(REGISTRY_ASSESSMENT_SHEET,REGISTRY_ASSESSMENT_HEADERS);
  const added={};[master,submission].forEach(function(sheet){added[sheet.getName()]=registryAppendHeaders_(sheet,REGISTRY_GOVERNANCE_HEADERS);});
  added[UPDATE_LOG_SHEET_NAME]=registryAppendHeaders_(updates,['Measurement_Method','Plan_Version']);
  if(assessment)added[REGISTRY_ASSESSMENT_SHEET]=registryAppendHeaders_(assessment,['Governance_Criteria_Result']);
  return added;
}
function setupRegistryGovernanceOutcomes() {
  return registryAuditOperation_('setupRegistryGovernanceOutcomes','',function(){const actor=requireActiveActor_('canAdmin'),lock=LockService.getScriptLock();lock.waitLock(30000);try{const added=ensureRegistryGovernanceHeaders_();SpreadsheetApp.flush();const result={success:true,version:ACTIONTRACKER_VERSION,added:added,historicalValuesPreserved:true}; return result;}finally{lock.releaseLock();}},true,false);
}
function registryMeasurementRead_(row,map) {
  function text(h){return String(row[map[normalizeHeaderKey_(h)]] || '');}
  return {measurementMethod:text('Measurement_Method') || 'Numeric KPI',governanceCurrentProblem:text('Governance_Current_Problem'),governanceExpectedOutcome:text('Governance_Expected_Outcome'),governanceAcceptanceCriteria:text('Governance_Acceptance_Criteria'),governanceVerificationMethod:text('Governance_Verification_Method')};
}
function registryGovernanceRead_(data,row,map) {
  Object.assign(data,registryMeasurementRead_(row,map));
  if(data.measurementMethod==='Governance outcome')Object.assign(data,{kpi:'Governance outcome',unit:'',direction:'',kpiUnit:'',kpiDirection:'',baseline:null,current:null,goal:null,baselineRaw:'',goalRaw:'',performanceChange:null,gapToGoal:null,progress:'',goalProgress:'',outcome:'Assess against criteria'});
  return data;
}

