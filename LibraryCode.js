/**
 * ============================================================================
 * MERGE STUDIO — REUSABLE GOOGLE APPS SCRIPT LIBRARY (LibraryCode.gs)
 * ============================================================================
 * A free, self-hosted AutoCrat alternative built on Google Apps Script.
 * Designed as a standalone Library project so you can maintain one codebase
 * and attach it to any number of Google Sheets without Marketplace review.
 *
 * Features:
 * - One merge configuration per spreadsheet (stored in DocumentProperties)
 * - Multi-sheet support
 * - <<Tag>> scanning across Doc body, header, and footer
 * - Branching logic: row merge condition + value-based branch rules to
 *   override Template Doc, Destination Folder, and Email Templates
 * - Two-stage Application Review & Multi-State Decision workflow
 * - PDF, Google Doc, or Both output formats + automated triggers
 * ============================================================================
 */

var CONFIG_KEY_PREFIX_ = 'MERGE_STUDIO_CONFIG_V1_';

var STATUS_COL_NAME_ = 'Merge Status';
var URL_COL_NAME_ = 'Merged File URL';
var TIMESTAMP_COL_NAME_ = 'Last Merged';
var DECISION_STATUS_COL_NAME_ = 'Decision Email Status';

var TRIGGER_HANDLERS_ = {
  ON_FORM_SUBMIT: 'onFormSubmitMerge',
  EVERY_15_MIN: 'runMergeFromTrigger'
};

/**
 * Returns spreadsheet context: sheet tabs, active sheet, headers, and the
 * single merge configuration for this spreadsheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Object} Context payload for the Sidebar UI.
 */
function getSpreadsheetContext(ss) {
  var spreadsheet = resolveSpreadsheet_(ss);
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(function (s) {
    return s.getName();
  });
  var activeSheet = spreadsheet.getActiveSheet();
  var activeSheetName = activeSheet ? activeSheet.getName() : sheetNames[0];
  var config = loadConfig_(spreadsheet, activeSheetName);
  var targetSheetName =
    config.sheetName && sheetNames.indexOf(config.sheetName) !== -1
      ? config.sheetName
      : activeSheetName;

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    sheetNames: sheetNames,
    activeSheetName: activeSheetName,
    headers: getSheetHeaders(spreadsheet, targetSheetName),
    config: config
  };
}

/**
 * Reads Row 1 headers for a specific sheet tab.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @return {string[]} Non-empty trimmed header names in column order.
 */
function getSheetHeaders(ss, sheetName) {
  var spreadsheet = resolveSpreadsheet_(ss);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet tab "' + sheetName + '" was not found.');
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return [];
  }
  var rawHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  return rawHeaders
    .map(function (h) {
      return String(h || '').trim();
    })
    .filter(function (h) {
      return h.length > 0;
    });
}

/**
 * Opens a Google Doc template by ID (or URL) and extracts unique <<Tag>>
 * placeholders across the document body, header, and footer.
 *
 * @param {string} docIdOrUrl Google Doc ID or full URL.
 * @return {Object} { docId, docTitle, tags }
 */
function getTemplateTags(docIdOrUrl) {
  var docId = extractDriveId_(docIdOrUrl);
  if (!docId) {
    throw new Error('Please enter a valid Google Doc Template ID or URL.');
  }

  var doc;
  try {
    doc = DocumentApp.openById(docId);
  } catch (err) {
    throw new Error(
      'Unable to open Google Doc "' +
        docId +
        '". Verify the ID and ensure you have at least View access. Details: ' +
        err.message
    );
  }

  var seen = {};
  var tags = [];

  var collectFromText = function (text) {
    if (!text) return;
    var regex = /<<([^<>]+?)>>/g;
    var match;
    while ((match = regex.exec(text)) !== null) {
      var tagName = String(match[1]).trim();
      if (tagName && !seen[tagName]) {
        seen[tagName] = true;
        tags.push(tagName);
      }
    }
  };

  var body = doc.getBody();
  if (body) collectFromText(body.getText());

  var header = doc.getHeader();
  if (header) collectFromText(header.getText());

  var footer = doc.getFooter();
  if (footer) collectFromText(footer.getText());

  return {
    docId: docId,
    docTitle: doc.getName(),
    tags: tags
  };
}

/**
 * Validates and persists the spreadsheet's merge configuration in DocumentProperties.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} config
 * @return {Object} { success: true, config: Object }
 */
function saveConfig(ss, config) {
  var spreadsheet = resolveSpreadsheet_(ss);
  if (!config || typeof config !== 'object') {
    throw new Error('Invalid configuration object.');
  }

  var sheetName = String(config.sheetName || '').trim();
  if (!sheetName || !spreadsheet.getSheetByName(sheetName)) {
    throw new Error('Selected sheet tab "' + sheetName + '" does not exist.');
  }


  var templateDocId = extractDriveId_(config.templateDocId);
  var destinationFolderId = extractDriveId_(config.destinationFolderId);

  var allowedFormats = { PDF: true, DOC: true, BOTH: true };
  var outputFormat = String(config.outputFormat || 'PDF').toUpperCase();
  if (!allowedFormats[outputFormat]) {
    throw new Error('Output format must be PDF, DOC, or BOTH.');
  }

  var allowedTriggers = { NONE: true, ON_FORM_SUBMIT: true, EVERY_15_MIN: true };
  var triggerMode = String(config.triggerMode || 'NONE').toUpperCase();
  if (!allowedTriggers[triggerMode]) {
    triggerMode = 'NONE';
  }

  var normalized = {
    sheetName: sheetName,
    templateDocId: templateDocId,
    destinationFolderId: destinationFolderId,
    mappings: normalizeMappings_(config.mappings),
    filenamePattern: String(config.filenamePattern || 'Merged - <<Timestamp>>').trim(),
    outputFormat: outputFormat,
    mergeCondition: normalizeCondition_(config.mergeCondition),
    branchRules: normalizeBranchRules_(config.branchRules),
    emailEnabled: Boolean(config.emailEnabled),
    emailRecipient: String(config.emailRecipient || '').trim(),
    emailSubject: String(config.emailSubject || '').trim(),
    emailBody: String(config.emailBody || '').trim(),
    reviewEmailEnabled: Boolean(config.reviewEmailEnabled),
    reviewEmailRecipient: String(config.reviewEmailRecipient || '').trim(),
    reviewEmailSubject: String(config.reviewEmailSubject || '').trim(),
    reviewEmailBody: String(config.reviewEmailBody || '').trim(),
    reviewAttachFile: config.reviewAttachFile !== undefined ? Boolean(config.reviewAttachFile) : true,
    decisionEmailEnabled: Boolean(config.decisionEmailEnabled),
    decisionColumn: String(config.decisionColumn || '').trim(),
    decisionRecipient: String(config.decisionRecipient || '').trim(),
    decisionStates: normalizeDecisionStates_(config),
    triggerMode: triggerMode,
    forceRemerge: Boolean(config.forceRemerge),
    updatedAt: new Date().toISOString()
  };

  if (normalized.emailEnabled && !normalized.emailRecipient) {
    throw new Error('Email recipient (column name or email address) is required when email is enabled.');
  }

  if (normalized.reviewEmailEnabled && !normalized.reviewEmailRecipient) {
    throw new Error('Reviewer email address is required when Application Review Email is enabled.');
  }

  if (normalized.decisionEmailEnabled) {
    if (!normalized.decisionColumn) {
      throw new Error('Dedicated Review Decision Column is required when Stage 2 decision emails are enabled.');
    }
    if (!normalized.decisionRecipient) {
      throw new Error('Applicant Recipient is required when Stage 2 decision emails are enabled.');
    }
    if (!normalized.decisionStates || normalized.decisionStates.length === 0) {
      throw new Error('At least one Decision State (e.g. Accepted, Rejected) is required.');
    }
  }

  saveConfig_(spreadsheet, normalized);

  syncProjectTriggers_(spreadsheet, normalized);

  return {
    success: true,
    config: normalized
  };
}

/**
 * Executes the spreadsheet's single merge configuration.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object=} options { forceRemerge?: boolean }
 * @return {Object} Execution summary.
 */
function runMerge(ss, options) {
  var spreadsheet = resolveSpreadsheet_(ss);
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('A merge is currently running on this spreadsheet. Please try again shortly.');
  }

  try {
    var opts = options || {};
    var activeSheet = spreadsheet.getActiveSheet();
    var defaultSheet = activeSheet ? activeSheet.getName() : spreadsheet.getSheets()[0].getName();
    var config = loadConfig_(spreadsheet, defaultSheet);

    return executeSingleJob_(spreadsheet, config, opts);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates the trigger mode and synchronizes project triggers for this spreadsheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} mode 'NONE' | 'ON_FORM_SUBMIT' | 'EVERY_15_MIN'
 * @return {Object} { success: true, mode: string, message: string }
 */
function configureTrigger(ss, mode) {
  var spreadsheet = resolveSpreadsheet_(ss);
  var normalizedMode = String(mode || 'NONE').toUpperCase();

  var allowed = { NONE: true, ON_FORM_SUBMIT: true, EVERY_15_MIN: true };
  if (!allowed[normalizedMode]) {
    throw new Error('Invalid trigger mode "' + mode + '". Expected NONE, ON_FORM_SUBMIT, or EVERY_15_MIN.');
  }

  var activeSheet = spreadsheet.getActiveSheet();
  var defaultSheet = activeSheet ? activeSheet.getName() : spreadsheet.getSheets()[0].getName();
  var config = loadConfig_(spreadsheet, defaultSheet);
  config.triggerMode = normalizedMode;
  config.updatedAt = new Date().toISOString();
  saveConfig_(spreadsheet, config);

  syncProjectTriggers_(spreadsheet, config);

  var msg =
    normalizedMode === 'NONE'
      ? 'Trigger disabled.'
      : normalizedMode === 'ON_FORM_SUBMIT'
      ? 'On Form Submit trigger enabled.'
      : 'Every 15 Minutes trigger enabled.';

  return {
    success: true,
    mode: normalizedMode,
    message: msg
  };
}

/**
 * Clears status column values for the configured target sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Object} { success: true, clearedColumns: string[], message: string }
 */
function clearStatuses(ss) {
  var spreadsheet = resolveSpreadsheet_(ss);
  var activeSheet = spreadsheet.getActiveSheet();
  var defaultSheet = activeSheet ? activeSheet.getName() : spreadsheet.getSheets()[0].getName();
  var config = loadConfig_(spreadsheet, defaultSheet);
  var sheet =
    config && config.sheetName
      ? spreadsheet.getSheetByName(config.sheetName)
      : spreadsheet.getActiveSheet();

  if (!sheet) {
    throw new Error('Target sheet tab was not found.');
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return {
      success: true,
      clearedColumns: [],
      message: 'No status rows to clear on "' + sheet.getName() + '".'
    };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
    return String(h || '').trim();
  });

  var targetCols = [
    STATUS_COL_NAME_,
    URL_COL_NAME_,
    TIMESTAMP_COL_NAME_,
    DECISION_STATUS_COL_NAME_
  ];
  var cleared = [];

  for (var i = 0; i < targetCols.length; i++) {
    var idx = headers.indexOf(targetCols[i]);
    if (idx !== -1) {
      sheet.getRange(2, idx + 1, lastRow - 1, 1).clearContent();
      cleared.push(targetCols[i]);
    }
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    clearedColumns: cleared,
    message:
      cleared.length > 0
        ? 'Cleared values on "' + sheet.getName() + '" in: ' + cleared.join(', ') + '.'
        : 'No Merge Studio status columns found on "' + sheet.getName() + '".'
  };
}

/* ============================================================================
 * INTERNAL EXECUTION & STORAGE HELPERS
 * ============================================================================ */

function executeSingleJob_(spreadsheet, config, opts) {
  var forceRemerge =
    opts.forceRemerge !== undefined ? Boolean(opts.forceRemerge) : Boolean(config.forceRemerge);

  var sheet = spreadsheet.getSheetByName(config.sheetName);
  if (!sheet) {
    throw new Error('Configured sheet "' + config.sheetName + '" was not found.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      totalRows: 0,
      processed: 0,
      decisionsSent: 0,
      skipped: 0,
      errors: 0,
      errorDetails: [],
      message: 'No data rows found below Row 1 headers.'
    };
  }

  var colMap = ensureTrackingColumns_(sheet);
  var lastCol = sheet.getLastColumn();
  var allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = allValues[0].map(function (h) {
    return String(h || '').trim();
  });

  var statusColIdx = colMap[STATUS_COL_NAME_];
  var urlColIdx = colMap[URL_COL_NAME_];
  var timeColIdx = colMap[TIMESTAMP_COL_NAME_];
  var decisionStatusColIdx = colMap[DECISION_STATUS_COL_NAME_];

  var timeZone = spreadsheet.getSpreadsheetTimeZone() || 'UTC';
  var fileCache = {};
  var folderCache = {};
  var processed = 0;
  var decisionsSent = 0;
  var skipped = 0;
  var errors = 0;
  var errorDetails = [];

  for (var r = 1; r < allValues.length; r++) {
    var rowValues = allValues[r];
    var sheetRowNum = r + 1;

    if (isRowEmpty_(rowValues, [statusColIdx, urlColIdx, timeColIdx, decisionStatusColIdx])) {
      skipped++;
      continue;
    }

    var currentStatus = String(rowValues[statusColIdx] || '').trim();
    var alreadyMerged = currentStatus === 'Merged' || currentStatus.indexOf('Merged (') === 0;
    var rowDidWork = false;

    try {
      var rowData = buildRowDataMap_(headers, rowValues);
      var timestampStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
      rowData['Timestamp'] = rowData['Timestamp'] || timestampStr;

      var primaryFileUrl = String(rowValues[urlColIdx] || '').trim();
      var createdPdfFile = null;
      var copyFile = null;

      // STAGE 1: Document Merge + Immediate Email + Reviewer Notification
      if (!alreadyMerged || forceRemerge) {
        if (config.mergeCondition && config.mergeCondition.enabled) {
          if (!evaluateCondition_(rowData, config.mergeCondition)) {
            sheet.getRange(sheetRowNum, statusColIdx + 1).setValue('Pending (Condition not met)');
            skipped++;
            continue;
          }
        }

        var activeBranch = findMatchingBranch_(rowData, config.branchRules);

        var effectiveTemplateId =
          (activeBranch && activeBranch.templateDocId) || config.templateDocId;
        var effectiveFolderId =
          (activeBranch && activeBranch.destinationFolderId) || config.destinationFolderId;

        if (!effectiveTemplateId) {
          throw new Error('Template Google Doc ID is missing (default and branch).');
        }
        if (!effectiveFolderId) {
          throw new Error('Destination Drive Folder ID is missing (default and branch).');
        }

        if (!fileCache[effectiveTemplateId]) {
          fileCache[effectiveTemplateId] = DriveApp.getFileById(effectiveTemplateId);
        }
        if (!folderCache[effectiveFolderId]) {
          folderCache[effectiveFolderId] = DriveApp.getFolderById(effectiveFolderId);
        }

        var templateFile = fileCache[effectiveTemplateId];
        var destFolder = folderCache[effectiveFolderId];

        var resolvedFilename = resolvePlaceholders_(
          config.filenamePattern || 'Merged Document - ' + sheetRowNum,
          rowData,
          config.mappings
        );
        resolvedFilename = sanitizeFilename_(resolvedFilename) || 'Merged_Row_' + sheetRowNum;

        copyFile = templateFile.makeCopy(resolvedFilename, destFolder);
        var copyDoc = DocumentApp.openById(copyFile.getId());

        replaceTagsInDocument_(copyDoc, config.mappings, rowData);
        copyDoc.saveAndClose();

        primaryFileUrl = copyFile.getUrl();

        if (config.outputFormat === 'PDF' || config.outputFormat === 'BOTH') {
          var pdfBlob = copyFile.getAs(MimeType.PDF).setName(resolvedFilename + '.pdf');
          createdPdfFile = destFolder.createFile(pdfBlob);
          primaryFileUrl = createdPdfFile.getUrl();
        }

        if (config.outputFormat === 'PDF') {
          copyFile.setTrashed(true);
          copyFile = null;
        }

        if (config.emailEnabled) {
          var emailCfg = {
            emailRecipient:
              (activeBranch && activeBranch.emailRecipient) || config.emailRecipient,
            emailSubject:
              (activeBranch && activeBranch.emailSubject) || config.emailSubject,
            emailBody:
              (activeBranch && activeBranch.emailBody) || config.emailBody,
            mappings: config.mappings
          };
          sendMergeEmail_(emailCfg, rowData, primaryFileUrl, createdPdfFile, copyFile);
        }

        if (config.reviewEmailEnabled) {
          var reviewCfg = {
            emailRecipient: config.reviewEmailRecipient,
            emailSubject:
              config.reviewEmailSubject || 'Application Review Requested: <<Timestamp>>',
            emailBody:
              config.reviewEmailBody ||
              'A new application has been merged for your review.\n\nDocument: <<Merged File URL>>',
            mappings: config.mappings
          };
          var attachPdf = config.reviewAttachFile !== false ? createdPdfFile : null;
          sendMergeEmail_(reviewCfg, rowData, primaryFileUrl, attachPdf, copyFile);
        }

        var mergedAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
        var statusLabel =
          activeBranch && activeBranch.name
            ? 'Merged (' + activeBranch.name + ')'
            : 'Merged';

        sheet.getRange(sheetRowNum, statusColIdx + 1).setValue(statusLabel);
        sheet.getRange(sheetRowNum, urlColIdx + 1).setValue(primaryFileUrl);
        sheet.getRange(sheetRowNum, timeColIdx + 1).setValue(mergedAt);
        rowData[URL_COL_NAME_] = primaryFileUrl;
        processed++;
        rowDidWork = true;
      }

      // STAGE 2: Multi-State Application Decision Email
      if (config.decisionEmailEnabled && config.decisionColumn) {
        var rawDecisionVal = rowData.hasOwnProperty(config.decisionColumn)
          ? String(rowData[config.decisionColumn] || '').trim()
          : '';

        if (rawDecisionVal) {
          var matchedState = findMatchingDecisionState_(rawDecisionVal, config.decisionStates);
          if (matchedState) {
            var currentDecisionStatus = String(rowValues[decisionStatusColIdx] || '').trim();
            var stateTag = 'Sent: ' + matchedState.name;
            var alreadySentThisState = currentDecisionStatus.indexOf(stateTag) === 0;

            if (!alreadySentThisState || forceRemerge) {
              var attachDecisionPdf = null;
              if (matchedState.attachFile) {
                if (createdPdfFile) {
                  attachDecisionPdf = createdPdfFile;
                } else if (primaryFileUrl) {
                  var existingId = extractDriveId_(primaryFileUrl);
                  if (existingId) {
                    try {
                      attachDecisionPdf = DriveApp.getFileById(existingId);
                    } catch (ignoreAttachErr) {
                      attachDecisionPdf = null;
                    }
                  }
                }
              }

              var decisionCfg = {
                emailRecipient: config.decisionRecipient,
                emailSubject:
                  matchedState.subject ||
                  'Application Update (' + matchedState.name + '): <<Name>>',
                emailBody:
                  matchedState.body ||
                  'Hello <<Name>>,\n\nYour application status has been updated to: ' +
                    matchedState.name +
                    '.',
                mappings: config.mappings
              };

              sendMergeEmail_(
                decisionCfg,
                rowData,
                primaryFileUrl,
                attachDecisionPdf,
                copyFile
              );

              var sentAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
              sheet
                .getRange(sheetRowNum, decisionStatusColIdx + 1)
                .setValue(stateTag + ' (' + sentAt + ')');
              decisionsSent++;
              rowDidWork = true;
            }
          }
        }
      }

      if (!rowDidWork) {
        skipped++;
      }
    } catch (rowErr) {
      var errMsg = rowErr && rowErr.message ? rowErr.message : String(rowErr);
      var failedAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(sheetRowNum, statusColIdx + 1).setValue('Error: ' + errMsg);
      sheet.getRange(sheetRowNum, timeColIdx + 1).setValue(failedAt);
      errors++;
      errorDetails.push('Row ' + sheetRowNum + ': ' + errMsg);
    }
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    totalRows: allValues.length - 1,
    processed: processed,
    decisionsSent: decisionsSent,
    skipped: skipped,
    errors: errors,
    errorDetails: errorDetails,
    message:
      'Complete. Merged: ' +
      processed +
      ', Decisions Sent: ' +
      decisionsSent +
      ', Skipped: ' +
      skipped +
      ', Errors: ' +
      errors +
      '.'
  };
}

function syncProjectTriggers_(spreadsheet, config) {
  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    var handler = existingTriggers[i].getHandlerFunction();
    if (
      handler === TRIGGER_HANDLERS_.ON_FORM_SUBMIT ||
      handler === TRIGGER_HANDLERS_.EVERY_15_MIN
    ) {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
  }

  if (config.triggerMode === 'ON_FORM_SUBMIT') {
    ScriptApp.newTrigger(TRIGGER_HANDLERS_.ON_FORM_SUBMIT)
      .forSpreadsheet(spreadsheet)
      .onFormSubmit()
      .create();
  }

  if (config.triggerMode === 'EVERY_15_MIN') {
    ScriptApp.newTrigger(TRIGGER_HANDLERS_.EVERY_15_MIN)
      .timeBased()
      .everyMinutes(15)
      .create();
  }
}

function resolveSpreadsheet_(ss) {
  var spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('No active Google Spreadsheet reference provided.');
  }
  return spreadsheet;
}

function getConfigKey_(ss) {
  return CONFIG_KEY_PREFIX_ + ss.getId();
}

function loadConfig_(ss, defaultSheetName) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(getConfigKey_(ss));

  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        return parsed;
      }
    } catch (e) {
      // Fall through to migration / default creation
    }
  }


  var config = buildDefaultConfig_(defaultSheetName);
  saveConfig_(ss, config);
  return config;
}

function saveConfig_(ss, config) {
  PropertiesService.getDocumentProperties().setProperty(
    getConfigKey_(ss),
    JSON.stringify(config)
  );
}

function buildDefaultConfig_(sheetName) {
  return {
    sheetName: sheetName || '',
    templateDocId: '',
    destinationFolderId: '',
    mappings: {},
    filenamePattern: '<<Name>> - Merged',
    outputFormat: 'PDF',
    mergeCondition: {
      enabled: false,
      column: '',
      operator: 'equals',
      value: ''
    },
    branchRules: [],
    emailEnabled: false,
    emailRecipient: '',
    emailSubject: 'Your merged document is ready',
    emailBody: 'Hello,\n\nPlease find your merged document attached or linked below:\n<<Merged File URL>>',
    reviewEmailEnabled: false,
    reviewEmailRecipient: '',
    reviewEmailSubject: 'Application Review Requested: <<Name>>',
    reviewEmailBody:
      'Hello,\n\nA new application has been merged and is ready for your review.\n\nDocument Link: <<Merged File URL>>',
    reviewAttachFile: true,
    decisionEmailEnabled: false,
    decisionColumn: '',
    decisionRecipient: '',
    decisionStates: [
      {
        name: 'Accepted',
        triggerValues: 'Accepted, Approved',
        subject: 'Congratulations! Your application has been accepted',
        body: 'Hello <<Name>>,\n\nWe are pleased to inform you that your application has been accepted.\n\nReference Document: <<Merged File URL>>',
        attachFile: false
      },
      {
        name: 'Waitlisted',
        triggerValues: 'Waitlisted, On Hold',
        subject: 'Application Status Update: Waitlisted',
        body: 'Hello <<Name>>,\n\nThank you for your patience. Your application is currently placed on our waitlist and we will update you as soon as a spot opens.',
        attachFile: false
      },
      {
        name: 'Rejected / Regarded',
        triggerValues: 'Rejected, Regarded, Declined',
        subject: 'Update regarding your application',
        body: 'Hello <<Name>>,\n\nThank you for your application. After careful review, we are unable to move forward at this time.',
        attachFile: false
      }
    ],
    triggerMode: 'NONE',
    forceRemerge: false
  };
}

function normalizeMappings_(mappings) {
  var out = {};
  if (!mappings || typeof mappings !== 'object') return out;
  var keys = Object.keys(mappings);
  for (var i = 0; i < keys.length; i++) {
    var tag = String(keys[i]).trim();
    var col = String(mappings[keys[i]] || '').trim();
    if (tag && col) {
      out[tag] = col;
    }
  }
  return out;
}

function normalizeCondition_(cond) {
  if (!cond || typeof cond !== 'object') {
    return { enabled: false, column: '', operator: 'equals', value: '' };
  }
  return {
    enabled: Boolean(cond.enabled),
    column: String(cond.column || '').trim(),
    operator: String(cond.operator || 'equals').trim(),
    value: String(cond.value || '').trim()
  };
}

function normalizeBranchRules_(rules) {
  if (!Array.isArray(rules)) return [];
  var out = [];
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (!r || typeof r !== 'object') continue;
    var col = String(r.column || '').trim();
    if (!col) continue;
    out.push({
      name: String(r.name || 'Branch ' + (i + 1)).trim(),
      column: col,
      operator: String(r.operator || 'equals').trim(),
      value: String(r.value || '').trim(),
      templateDocId: extractDriveId_(r.templateDocId),
      destinationFolderId: extractDriveId_(r.destinationFolderId),
      emailRecipient: String(r.emailRecipient || '').trim(),
      emailSubject: String(r.emailSubject || '').trim(),
      emailBody: String(r.emailBody || '')
    });
  }
  return out;
}

function normalizeDecisionStates_(config) {
  if (Array.isArray(config.decisionStates) && config.decisionStates.length > 0) {
    var states = [];
    for (var i = 0; i < config.decisionStates.length; i++) {
      var st = config.decisionStates[i];
      if (!st || typeof st !== 'object') continue;
      var triggers = String(st.triggerValues || '').trim();
      if (!triggers) continue;
      states.push({
        name: String(st.name || 'State ' + (i + 1)).trim(),
        triggerValues: triggers,
        subject: String(st.subject || '').trim(),
        body: String(st.body || ''),
        attachFile: Boolean(st.attachFile)
      });
    }
    return states;
  }
  return [];
}

function findMatchingDecisionState_(cellValue, decisionStates) {
  if (!Array.isArray(decisionStates) || decisionStates.length === 0) return null;
  var actualLower = String(cellValue || '').trim().toLowerCase();
  if (!actualLower) return null;

  for (var i = 0; i < decisionStates.length; i++) {
    var st = decisionStates[i];
    var tokens = String(st.triggerValues || '').split(',');
    for (var t = 0; t < tokens.length; t++) {
      var expected = String(tokens[t] || '').trim().toLowerCase();
      if (expected && actualLower === expected) {
        return st;
      }
    }
  }
  return null;
}

function evaluateCondition_(rowData, cond) {
  if (!cond || !cond.column) return true;
  var actual = rowData.hasOwnProperty(cond.column) ? String(rowData[cond.column]) : '';
  var actualTrim = actual.trim();
  var actualLower = actualTrim.toLowerCase();
  var expectedTrim = String(cond.value || '').trim();
  var expectedLower = expectedTrim.toLowerCase();

  switch (cond.operator) {
    case 'equals':
      return actualLower === expectedLower;
    case 'not_equals':
      return actualLower !== expectedLower;
    case 'contains':
      return expectedLower !== '' && actualLower.indexOf(expectedLower) !== -1;
    case 'not_contains':
      return expectedLower === '' || actualLower.indexOf(expectedLower) === -1;
    case 'greater_than': {
      var aNum = parseFloat(actualTrim.replace(/[^0-9.-]/g, ''));
      var eNum = parseFloat(expectedTrim.replace(/[^0-9.-]/g, ''));
      return !isNaN(aNum) && !isNaN(eNum) && aNum > eNum;
    }
    case 'less_than': {
      var aNum2 = parseFloat(actualTrim.replace(/[^0-9.-]/g, ''));
      var eNum2 = parseFloat(expectedTrim.replace(/[^0-9.-]/g, ''));
      return !isNaN(aNum2) && !isNaN(eNum2) && aNum2 < eNum2;
    }
    case 'is_empty':
      return actualTrim === '';
    case 'is_not_empty':
      return actualTrim !== '';
    default:
      return false;
  }
}

function findMatchingBranch_(rowData, branchRules) {
  if (!Array.isArray(branchRules) || branchRules.length === 0) return null;
  for (var i = 0; i < branchRules.length; i++) {
    if (evaluateCondition_(rowData, branchRules[i])) {
      return branchRules[i];
    }
  }
  return null;
}

function extractDriveId_(input) {
  var str = String(input || '').trim();
  if (!str) return '';
  var match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

function ensureTrackingColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers =
    lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
          return String(h || '').trim();
        })
      : [];

  var required = [
    STATUS_COL_NAME_,
    URL_COL_NAME_,
    TIMESTAMP_COL_NAME_,
    DECISION_STATUS_COL_NAME_
  ];
  var colMap = {};

  for (var i = 0; i < required.length; i++) {
    var name = required[i];
    var idx = headers.indexOf(name);
    if (idx === -1) {
      idx = headers.length;
      headers.push(name);
      var cell = sheet.getRange(1, idx + 1);
      cell.setValue(name).setFontWeight('bold');
    }
    colMap[name] = idx;
  }

  return colMap;
}

function isRowEmpty_(rowValues, ignoreIndices) {
  for (var i = 0; i < rowValues.length; i++) {
    if (ignoreIndices.indexOf(i) !== -1) continue;
    if (String(rowValues[i] || '').trim() !== '') {
      return false;
    }
  }
  return true;
}

function buildRowDataMap_(headers, rowValues) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h) {
      map[h] = rowValues[i] !== undefined ? String(rowValues[i]) : '';
    }
  }
  return map;
}

function replaceTagsInDocument_(doc, mappings, rowData) {
  var sections = [doc.getBody(), doc.getHeader(), doc.getFooter()];
  var replacedTags = {};

  var tags = Object.keys(mappings || {});
  for (var t = 0; t < tags.length; t++) {
    var tag = tags[t];
    var colName = mappings[tag];
    var value = colName && rowData.hasOwnProperty(colName) ? rowData[colName] : '';
    replacedTags[tag] = true;
    applyTagReplacement_(sections, tag, value);
  }

  var headers = Object.keys(rowData || {});
  for (var h = 0; h < headers.length; h++) {
    var header = headers[h];
    if (!replacedTags[header]) {
      applyTagReplacement_(sections, header, rowData[header]);
    }
  }
}

function applyTagReplacement_(sections, tag, value) {
  var searchPattern = escapeRegex_('<<' + tag + '>>');
  var safeReplacement = escapeReplacement_(value);
  for (var s = 0; s < sections.length; s++) {
    if (sections[s]) {
      sections[s].replaceText(searchPattern, safeReplacement);
    }
  }
}

function resolvePlaceholders_(templateStr, rowData, mappings) {
  var text = String(templateStr || '');
  return text.replace(/<<([^<>]+?)>>/g, function (_full, rawKey) {
    var key = String(rawKey).trim();
    if (rowData.hasOwnProperty(key)) {
      return rowData[key];
    }
    if (mappings && mappings[key] && rowData.hasOwnProperty(mappings[key])) {
      return rowData[mappings[key]];
    }
    return '';
  });
}

function sendMergeEmail_(config, rowData, primaryFileUrl, attachmentFile, docFile) {
  var recipientSpec = String(config.emailRecipient || '').trim();
  var recipient = '';

  if (/<<[^<>]+>>/.test(recipientSpec)) {
    recipient = resolvePlaceholders_(recipientSpec, rowData, config.mappings).trim();
  } else if (rowData.hasOwnProperty(recipientSpec)) {
    recipient = String(rowData[recipientSpec] || '').trim();
  } else {
    recipient = recipientSpec;
  }

  if (!recipient || recipient.indexOf('@') === -1) {
    throw new Error('Invalid or empty email recipient: "' + recipient + '".');
  }

  var enrichedRowData = {};
  var keys = Object.keys(rowData);
  for (var i = 0; i < keys.length; i++) {
    enrichedRowData[keys[i]] = rowData[keys[i]];
  }
  enrichedRowData['Merged File URL'] = primaryFileUrl;

  var subject = resolvePlaceholders_(
    config.emailSubject || 'Your merged document',
    enrichedRowData,
    config.mappings
  );
  var body = resolvePlaceholders_(
    config.emailBody || 'Your document is ready: ' + primaryFileUrl,
    enrichedRowData,
    config.mappings
  );

  var mailOptions = {};
  if (attachmentFile) {
    mailOptions.attachments = [attachmentFile.getBlob()];
  } else if (docFile && primaryFileUrl && body.indexOf(primaryFileUrl) === -1) {
    body += '\n\nDocument Link: ' + primaryFileUrl;
  }

  GmailApp.sendEmail(recipient, subject, body, mailOptions);
}

function sanitizeFilename_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeRegex_(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeReplacement_(str) {
  return String(str || '')
    .replace(/\\/g, '\\\\')
    .replace(/\$/g, '\\$');
}