/** LodashGS library: 1SQ0PlSMwndIuOAgtVJdjxsuXueECtY9OGejVDS37ckSVbMll73EXf2PW */

const APP_KEY = '69091b790253df048876a1e8';
const SERVER_ENDPOINT = 'https://app.augmentir.com';
const SETUP_SHEET_NAME = 'SETUP';
const SETUP_NEW_SHEET_NAME = 'SETUP_NEW';
const DATA_SHEET_NAME = 'DATA';
const DOMO_CREDENTIAL_SPREADSHEET_ID = '1LwX3FWLJOlssznZKAH2Xu6C3j3WAq1yhMU-xOb9k_Rs';
const DOMO_CREDENTIAL_SHEET_NAME = 'Domo';
const LOCAL_TIME_ZONE = 'Asia/Ho_Chi_Minh';
const LIMIT = 1000;
const GOOGLE_SHEET_CHUNK_ROWS = 5000;
const DEFAULT_HEADERS = [
  "jobId",
  "jobEventUrl",
  "procedureId",
  "title",
  "publishState",
  "procedureVersion",
  "status",

  "createdAt",
  "startedAt",
  "updatedAt",
  "completedAt",

  "stationId",
  "isExpired",
  "minutesRemaining",
  "approxMinutesElapsed",

  "cardInTask",
  "cardsInTask",

  "unitNumber",
  "unitCount",
];
const _ = LodashGS.load();
const CONFIG_RANGE = {
  PROCEDURE: {
    NAME: 'C2',
    ID: 'C3',
  },
  PROCEDURE_TABLE: {
    START_ROW: 3,
    START_COLUMN: 1,
    COLUMN_COUNT: 21,
  },
  STATUS: {
    CREATED: 'C4',
    INPROGRESS: 'C5',
    PAUSED: 'C6',
    CANCELLED: 'C7',
    COMPLETED: 'C8',
  },
  PUBLISH_STATE: {
    PRODUCTION: 'C9',
    PREPRODUCTION: 'C10',
    DRAFT: 'C11',
  },
  START_DATE: 'C12',
  END_DATE: 'C13',
  LOCK_HEADERS: 'C14',
  INCLUDE_EXPIRED_JOBS: 'C15',
};

const ss = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Helper Planner 🚀')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Augmentir')
    .addItem('Update', 'update')
    .addItem('Create SETUP_NEW', 'setupMultiProcedureTable')
    .addItem('Create README', 'createReadmeSheet')
    .addItem('Run Diagnostics', 'runDiagnostic')
    .addSeparator()
    .addSubMenu(ui.createMenu('Clear')
      .addItem('Clear Updated From', 'clearUpdatedFrom')
      .addItem('Clear Data', 'clearData')
      .addItem('Clear Headers', 'clearHeaders'))
    .addToUi();

  ui.createMenu('Helper Planner 🚀')
    .addItem('Sync Master Data', 'syncHelperPlannerMasterData')
    .addItem('Setup Database Sheets', 'setupHelperPlannerSheets')
    .addToUi();
}

function clearUpdatedFrom() {
  const newConfigSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME);
  if (newConfigSheet) {
    const table = CONFIG_RANGE.PROCEDURE_TABLE;
    const lastRow = newConfigSheet.getLastRow();
    if (lastRow >= table.START_ROW) {
      const fromColumn = newConfigSheet.getRange(2, 2).getValue() === 'Platform' ? 17 : 13;
      newConfigSheet.getRange(table.START_ROW, fromColumn, lastRow - table.START_ROW + 1, 1).clearContent();
      return;
    }
  }

  const configSheet = ss.getSheetByName(SETUP_SHEET_NAME);
  configSheet.getRange(CONFIG_RANGE.START_DATE).activate();
  configSheet.getRange(CONFIG_RANGE.START_DATE).clearContent();
}

function clearData() {
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  dataSheet.activate();
  if (dataSheet.getLastRow() < 2) return;
  dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).clear();
}

function clearHeaders() {
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  dataSheet.activate();
  dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).clear();
}

function update() {
  const configSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME) || ss.getSheetByName(SETUP_SHEET_NAME);
  if (!configSheet) throw new Error(`Không tìm thấy Sheet '${SETUP_NEW_SHEET_NAME}' hoặc '${SETUP_SHEET_NAME}'`);

  const procedureConfigs = _getProcedureConfigs(configSheet);
  if (procedureConfigs.length === 0) {
    throw new Error(`Chưa cấu hình Procedure nào trong Sheet '${configSheet.getName()}'`);
  }

  for (const procedureConfig of procedureConfigs) {
    try {
      if (procedureConfig.platform === 'DOMO') {
        _updateDomoDataset(procedureConfig);
      } else if (procedureConfig.platform === 'GOOGLE SHEET') {
        _copyGoogleSheetData(procedureConfig);
      } else {
        _updateProcedure(procedureConfig);
      }
    } catch (err) {
      const sourceName = procedureConfig.dataSheetName || procedureConfig.procedureId || procedureConfig.procedureName || procedureConfig.platform;
      throw new Error(`Update source '${sourceName}' failed: ${err.message || err}`);
    }
  }
}

function _getProcedureConfigs(configSheet) {
  const tableConfigs = configSheet.getName() === SETUP_NEW_SHEET_NAME ? _getProcedureConfigsFromTable(configSheet) : [];
  if (tableConfigs.length > 0) return tableConfigs;

  const procedureName = _cleanConfigText(configSheet.getRange(CONFIG_RANGE.PROCEDURE.NAME).getValue());
  const procedureId = _cleanConfigText(configSheet.getRange(CONFIG_RANGE.PROCEDURE.ID).getValue());

  if (!procedureName && !procedureId) return [];
  if (procedureName && procedureId) {
    throw new Error(`Không thể chỉ định đồng thời 'Procedure Name' và 'Procedure ID'`);
  }

  return [{
    procedureName,
    procedureId,
    platform: 'AUGMENTIR',
    dataSheetName: DATA_SHEET_NAME,
    status: {
      created: configSheet.getRange(CONFIG_RANGE.STATUS.CREATED).getValue(),
      inProgress: configSheet.getRange(CONFIG_RANGE.STATUS.INPROGRESS).getValue(),
      paused: configSheet.getRange(CONFIG_RANGE.STATUS.PAUSED).getValue(),
      cancelled: configSheet.getRange(CONFIG_RANGE.STATUS.CANCELLED).getValue(),
      completed: configSheet.getRange(CONFIG_RANGE.STATUS.COMPLETED).getValue(),
    },
    publishState: {
      production: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PRODUCTION).getValue(),
      preProduction: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PREPRODUCTION).getValue(),
      draft: configSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.DRAFT).getValue(),
    },
    startDate: configSheet.getRange(CONFIG_RANGE.START_DATE).getValue(),
    endDate: configSheet.getRange(CONFIG_RANGE.END_DATE).getValue(),
    lockHeaders: configSheet.getRange(CONFIG_RANGE.LOCK_HEADERS).getValue(),
    includeExpiredJobs: configSheet.getRange(CONFIG_RANGE.INCLUDE_EXPIRED_JOBS).getValue(),
    configSheet,
    updatedFromCell: CONFIG_RANGE.START_DATE,
    lastUpdatedCell: CONFIG_RANGE.START_DATE,
  }];
}

function _cleanConfigText(value) {
  if (value === null || value === undefined) return '';
  return value.toString().trim();
}

function _getProcedureConfigsFromTable(configSheet) {
  const table = CONFIG_RANGE.PROCEDURE_TABLE;
  const lastRow = configSheet.getLastRow();
  if (lastRow < table.START_ROW) return [];

  const isNewLayout = configSheet.getRange(2, 2).getValue() === 'Platform';
  const values = configSheet
    .getRange(table.START_ROW, table.START_COLUMN, lastRow - table.START_ROW + 1, table.COLUMN_COUNT)
    .getValues();

  return values
    .map((row, index) => {
      const rowNumber = table.START_ROW + index;
      return isNewLayout
        ? _parseSetupNewRow(configSheet, row, rowNumber)
        : _parseLegacySetupNewRow(configSheet, row, rowNumber);
    })
    .filter(Boolean);
}

function _parseLegacySetupNewRow(configSheet, row, rowNumber) {
  const enabled = row[0];
  const dataSheetName = _cleanConfigText(row[1]) || DATA_SHEET_NAME;
  const procedureName = _cleanConfigText(row[2]);
  const procedureId = _cleanConfigText(row[3]);
  const updatedFrom = row[12];
  const updatedTo = row[13];

  if (!_isEnabled(enabled)) return null;
  if (!procedureName && !procedureId) return null;
  if (procedureName && procedureId) {
    throw new Error(`Dòng ${rowNumber}: chỉ nhập một trong hai cột Procedure Name hoặc Procedure ID`);
  }

  return {
    platform: 'AUGMENTIR',
    procedureName,
    procedureId,
    dataSheetName,
    status: {
      created: row[4],
      inProgress: row[5],
      paused: row[6],
      cancelled: row[7],
      completed: row[8],
    },
    publishState: {
      production: row[9],
      preProduction: row[10],
      draft: row[11],
    },
    startDate: updatedFrom,
    endDate: updatedTo,
    lockHeaders: row[14],
    includeExpiredJobs: row[15],
    configSheet,
    updatedFromCell: configSheet.getRange(rowNumber, 13).getA1Notation(),
    lastUpdatedCell: configSheet.getRange(rowNumber, 17).getA1Notation(),
  };
}

function _parseSetupNewRow(configSheet, row, rowNumber) {
  const enabled = row[0];
  const platform = (_cleanConfigText(row[1]) || 'AUGMENTIR').toUpperCase();
  const dataSheetName = _cleanConfigText(row[2]) || DATA_SHEET_NAME;
  const procedureName = _cleanConfigText(row[3]);
  const procedureId = _cleanConfigText(row[4]);
  const dateColumn = _cleanConfigText(row[5]);
  const dateType = (row[6] || 'AUTO').toString().trim().toUpperCase();
  const query = row[7];
  const updatedFrom = row[16];
  const updatedTo = row[17];
  const updatedFromDisplay = configSheet.getRange(rowNumber, 17).getDisplayValue();
  const updatedToDisplay = configSheet.getRange(rowNumber, 18).getDisplayValue();

  if (!_isEnabled(enabled)) return null;
  if (platform !== 'AUGMENTIR' && platform !== 'DOMO' && platform !== 'GOOGLE SHEET') {
    throw new Error(`Dòng ${rowNumber}: Platform chỉ được nhập Augmentir, Domo hoặc Google sheet`);
  }
  if (!procedureId && !procedureName) return null;
  if (platform === 'AUGMENTIR' && procedureName && procedureId) {
    throw new Error(`Dòng ${rowNumber}: Augmentir chỉ nhập một trong hai cột Procedure Name hoặc Procedure ID`);
  }
  if (platform === 'DOMO' && !procedureId) {
    throw new Error(`Dòng ${rowNumber}: Domo cần Dataset ID trong cột Procedure ID / Dataset ID`);
  }
  if (platform === 'GOOGLE SHEET' && (!procedureName || !procedureId)) {
    throw new Error(`Dòng ${rowNumber}: Google sheet cần Procedure Name là tên sheet nguồn và Procedure ID / Dataset ID là Spreadsheet ID nguồn`);
  }

  return {
    platform,
    procedureName,
    procedureId,
    datasetId: procedureId,
    dataSheetName,
    dateColumn,
    dateType,
    query,
    status: {
      created: row[8],
      inProgress: row[9],
      paused: row[10],
      cancelled: row[11],
      completed: row[12],
    },
    publishState: {
      production: row[13],
      preProduction: row[14],
      draft: row[15],
    },
    startDate: platform === 'DOMO' ? (updatedFromDisplay || updatedFrom) : updatedFrom,
    endDate: platform === 'DOMO' ? (updatedToDisplay || updatedTo) : updatedTo,
    lockHeaders: row[18],
    includeExpiredJobs: row[19],
    configSheet,
    updatedFromCell: configSheet.getRange(rowNumber, 17).getA1Notation(),
    lastUpdatedCell: configSheet.getRange(rowNumber, 21).getA1Notation(),
  };
}

function _isEnabled(value) {
  if (value === false) return false;
  if (typeof value === 'string' && ['false', 'no', 'n', '0'].includes(value.toLowerCase())) return false;
  return true;
}

function _updateProcedure(procedureConfig) {
  let posibleMore = true;
  let lastupdated = _parseFlexibleDate(procedureConfig.startDate) || new Date(1999, 1, 1);
  let latestAcceptedUpdatedAt = null;
  _ensureDataSheet(procedureConfig.dataSheetName);

  while (posibleMore) {
    const response = _getJobsFromProcedure({
      procedureId: procedureConfig.procedureId,
      procedureName: procedureConfig.procedureName,
      startDate: new Date(lastupdated.getTime() + 1).toISOString(),
      endDate: procedureConfig.endDate ? (_parseFlexibleDate(procedureConfig.endDate) || new Date(procedureConfig.endDate)).toISOString() : undefined,
    });

    const actualResults = response.actualResults || [];
    posibleMore = response.possiblyMore || false;
    Logger.log(`Augmentir rows returned for ${procedureConfig.dataSheetName}: ${actualResults.length}, possiblyMore=${posibleMore}`);

    for (const job of actualResults) {
      const _updatedAt = new Date(job.updatedAt);
      if (_updatedAt > lastupdated) lastupdated = new Date(_updatedAt);
    }

    if (actualResults.length === 0) break;
    const writeResult = writeData(actualResults, procedureConfig);
    Logger.log(`Augmentir write result for ${procedureConfig.dataSheetName}: filtered=${writeResult.filteredRows}, generated=${writeResult.generatedRows}, updated=${writeResult.updatedRows}, added=${writeResult.addedRows}`);
    if (writeResult.latestAcceptedUpdatedAt && (!latestAcceptedUpdatedAt || writeResult.latestAcceptedUpdatedAt > latestAcceptedUpdatedAt)) {
      latestAcceptedUpdatedAt = writeResult.latestAcceptedUpdatedAt;
    }
  }

  if (latestAcceptedUpdatedAt && procedureConfig.lastUpdatedCell) {
    if (procedureConfig.updatedFromCell) {
      _setConfigDateValue(procedureConfig.configSheet, procedureConfig.updatedFromCell, latestAcceptedUpdatedAt);
    }
    _setConfigDateValue(procedureConfig.configSheet, procedureConfig.lastUpdatedCell, latestAcceptedUpdatedAt);
  }
}

function _setConfigDateValue(configSheet, cellA1, value) {
  const date = _parseFlexibleDate(value);
  if (!date) {
    configSheet.getRange(cellA1).setValue(value);
    return;
  }
  configSheet.getRange(cellA1).setValue(Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'));
}

function _ensureDataSheet(dataSheetName) {
  let sheet = ss.getSheetByName(dataSheetName);
  if (!sheet) sheet = ss.insertSheet(dataSheetName);
  return sheet;
}

function _updateDomoDataset(config) {
  if (!config.dateColumn) {
    throw new Error(`Domo dataset '${config.datasetId}' chưa có Date column`);
  }

  const sql = _buildDomoQuery(config.query, config.dateColumn, config.startDate, config.endDate, config.dateType);
  Logger.log('Domo SQL: ' + sql);
  const data = _queryDomoData(config.datasetId, sql);
  Logger.log('Domo rows returned: ' + ((data.rows || []).length));
  _writeDomoDataToSheet(data, config);

  const latestDate = _getMaxDateFromDomoRows(data, config.dateColumn);
  if (latestDate && config.updatedFromCell) {
    _setConfigDateValue(config.configSheet, config.updatedFromCell, latestDate);
    _setConfigDateValue(config.configSheet, config.lastUpdatedCell, latestDate);
  }
}

function _buildDomoQuery(baseQuery, dateColumn, startDate, endDate, dateType) {
  let sql = (baseQuery || 'SELECT * FROM table').toString().trim();
  sql = sql.replace(/;+\s*$/g, '');

  const filters = [];
  const dateExpression = _formatDomoColumn(dateColumn);
  if (startDate) filters.push(`${dateExpression} > '${_formatDomoDateLiteral(startDate, dateType)}'`);
  if (endDate) filters.push(`${dateExpression} <= '${_formatDomoDateLiteral(endDate, dateType)}'`);
  if (filters.length === 0) return sql;

  const orderByMatch = sql.match(/\s+order\s+by\s+/i);
  const orderByClause = orderByMatch ? sql.slice(orderByMatch.index) : '';
  const queryWithoutOrderBy = orderByMatch ? sql.slice(0, orderByMatch.index) : sql;
  const joiner = /\swhere\s/i.test(queryWithoutOrderBy) ? ' AND ' : ' WHERE ';
  return queryWithoutOrderBy + joiner + filters.join(' AND ') + orderByClause;
}

function _formatDomoColumn(columnName) {
  const name = columnName.toString().trim();
  if (/^`.*`$/.test(name) || /^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return '`' + name.replace(/`/g, '``') + '`';
}

function _formatDomoDateLiteral(value, dateType) {
  const date = _parseFlexibleDate(value);
  if (date && dateType === 'DATE') return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd');
  if (date && dateType === 'DATETIME') return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  if (date) return Utilities.formatDate(date, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss') + '.0000000 +07:00';
  return value.toString().trim();
}

function _getDomoAccessToken() {
  const credentialSpreadsheet = SpreadsheetApp.openById(DOMO_CREDENTIAL_SPREADSHEET_ID);
  const credentialSheet = credentialSpreadsheet.getSheetByName(DOMO_CREDENTIAL_SHEET_NAME);
  if (!credentialSheet) throw new Error(`Không tìm thấy Sheet '${DOMO_CREDENTIAL_SHEET_NAME}' chứa Domo credential`);

  const clientId = credentialSheet.getRange('B1').getValue();
  const clientSecret = credentialSheet.getRange('B2').getValue();
  const options = {
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
    },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch('https://api.domo.com/oauth/token?grant_type=client_credentials&scope=data', options);
  if (response.getResponseCode() >= 300) {
    throw new Error('Không lấy được Domo access token: ' + response.getContentText());
  }
  return JSON.parse(response.getContentText()).access_token;
}

function _queryDomoData(datasetId, query) {
  const accessToken = _getDomoAccessToken();
  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      Authorization: 'bearer ' + accessToken,
    },
    payload: JSON.stringify({ sql: query }),
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch('https://api.domo.com/v1/datasets/query/execute/' + datasetId, options);
  if (response.getResponseCode() >= 300) {
    throw new Error('Không query được Domo dataset ' + datasetId + ': ' + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

function _writeDomoDataToSheet(data, config) {
  const sheet = _ensureDataSheet(config.dataSheetName);
  const columns = data.columns || [];
  const rows = data.rows || [];

  if (columns.length === 0) return;

  const keyIndex = columns.indexOf('Id');
  if (keyIndex < 0) throw new Error(`Domo dataset '${config.datasetId}' không có cột Id để upsert dữ liệu`);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (JSON.stringify(currentHeaders.slice(0, columns.length)) !== JSON.stringify(columns)) {
      throw new Error(`Header sheet '${config.dataSheetName}' không khớp Domo dataset, dừng để tránh ghi sai dữ liệu lịch sử`);
    }
  }

  const currentRowCount = Math.max(sheet.getLastRow() - 1, 0);
  const currentRows = currentRowCount > 0
    ? sheet.getRange(2, 1, currentRowCount, columns.length).getValues()
    : [];
  const rowIndexById = {};
  currentRows.forEach((row, index) => {
    const id = row[keyIndex];
    if (id !== '' && id != null) rowIndexById[id.toString()] = index;
  });

  const newRows = [];
  let updatedCount = 0;
  rows.forEach(row => {
    row.length = columns.length;
    const id = row[keyIndex];
    const targetIndex = id !== '' && id != null ? rowIndexById[id.toString()] : null;
    if (targetIndex !== null && targetIndex !== undefined) {
      currentRows[targetIndex] = row;
      updatedCount++;
    } else {
      newRows.push(row);
    }
  });

  if (currentRowCount > 0 && updatedCount > 0) {
    sheet.getRange(2, 1, currentRowCount, columns.length).setValues(currentRows);
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, columns.length).setValues(newRows);
  }
  sheet.getRange(1, 1, 1, columns.length)
    .setBackground('#b9f79c')
    .setFontColor('blue')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function _copyGoogleSheetData(config) {
  const sourceSpreadsheet = SpreadsheetApp.openById(config.procedureId);
  const sourceSheet = sourceSpreadsheet.getSheetByName(config.procedureName);
  if (!sourceSheet) {
    throw new Error(`Không tìm thấy sheet nguồn '${config.procedureName}' trong Spreadsheet '${config.procedureId}'`);
  }

  const targetSheet = _ensureDataSheet(config.dataSheetName);
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) return;

  const sourceTz = sourceSpreadsheet.getSpreadsheetTimeZone();
  const headers = sourceSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const dateIndex = config.dateColumn ? _findHeaderIndex(headers, config.dateColumn) : -1;
  if (config.dateColumn && dateIndex < 0) {
    throw new Error(`Không tìm thấy cột ngày '${config.dateColumn}' trong sheet nguồn '${config.procedureName}'`);
  }

  const startDate = _parseFlexibleDate(config.startDate);
  const endDate = _parseFlexibleDate(config.endDate);
  let latestDate = null;

  let isRelativeWorkingDays = false;
  let thresholdStartDate = startDate;

  // 1. Resolve relative working days starting from the bottom up (highly optimized)
  if (config.dateColumn && dateIndex >= 0) {
    const rawStart = config.startDate;
    if (rawStart !== null && rawStart !== undefined && rawStart !== '') {
      const num = Number(rawStart);
      if (!isNaN(num) && Number.isInteger(num)) {
        isRelativeWorkingDays = true;
        const workingDaysCount = Math.abs(num);

        const seenDates = new Set();
        const uniqueDates = [];
        const DATE_CHUNK_SIZE = 5000;
        let dateStop = false;

        // Scan from bottom to top in chunks of 5000 rows
        for (let currentEndRow = lastRow; currentEndRow >= 2; currentEndRow -= DATE_CHUNK_SIZE) {
          const startRow = Math.max(2, currentEndRow - DATE_CHUNK_SIZE + 1);
          const rowCount = currentEndRow - startRow + 1;
          const dateValues = sourceSheet.getRange(startRow, dateIndex + 1, rowCount, 1).getValues();

          // Scan within chunk bottom-to-top
          for (let i = dateValues.length - 1; i >= 0; i--) {
            const val = dateValues[i][0];
            const parsed = _parseFlexibleDate(val, sourceTz);
            if (parsed) {
              const dateStr = Utilities.formatDate(parsed, sourceTz, "yyyy-MM-dd");
              if (!seenDates.has(dateStr)) {
                seenDates.add(dateStr);
                uniqueDates.push(parsed);
                if (seenDates.size >= workingDaysCount) {
                  dateStop = true;
                  break;
                }
              }
            }
          }
          if (dateStop) break;
        }

        if (uniqueDates.length > 0) {
          uniqueDates.sort((a, b) => b.getTime() - a.getTime()); // newest first
          const targetIndex = Math.min(workingDaysCount - 1, uniqueDates.length - 1);
          thresholdStartDate = new Date(uniqueDates[targetIndex]);
          thresholdStartDate.setHours(0, 0, 0, 0);
        }
      }
    }
  }

  targetSheet.clearContents();
  _writeGoogleSheetRows(targetSheet, 1, _sanitizeValuesForWrite([headers], sourceTz));

  // 2. Fetch and filter row data bottom-up with early termination
  const allRowsToWrite = [];
  const CHUNK_SIZE = 5000;
  let shouldStop = false;

  for (let currentEndRow = lastRow; currentEndRow >= 2; currentEndRow -= CHUNK_SIZE) {
    const startRow = Math.max(2, currentEndRow - CHUNK_SIZE + 1);
    const rowCount = currentEndRow - startRow + 1;
    const rows = sourceSheet.getRange(startRow, 1, rowCount, lastColumn).getValues();

    // Process chunk bottom-to-top
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      if (config.dateColumn && dateIndex >= 0) {
        const rowDate = _parseFlexibleDate(row[dateIndex], sourceTz);
        if (!rowDate) continue; // Skip blank or malformed date rows

        // Early termination: since rows are chronologically sorted, if row date is older than
        // our start limit, all remaining rows higher up are also older. Stop scanning completely!
        if (isRelativeWorkingDays) {
          if (thresholdStartDate && rowDate < thresholdStartDate) {
            shouldStop = true;
            break;
          }
        } else {
          if (startDate && rowDate <= startDate) {
            shouldStop = true;
            break;
          }
        }

        // Skip rows newer than end date (but keep scanning since older matching rows reside above)
        if (endDate && rowDate > endDate) {
          continue;
        }

        if (!latestDate || rowDate > latestDate) {
          latestDate = rowDate;
        }
      }

      allRowsToWrite.push(row);
    }

    if (shouldStop) break;
  }

  // 3. Restore original chronological top-to-bottom order
  allRowsToWrite.reverse();

  if (allRowsToWrite.length > 0) {
    const sanitizedRows = _sanitizeValuesForWrite(allRowsToWrite, sourceTz);
    _writeGoogleSheetRows(targetSheet, 2, sanitizedRows);
  }

  targetSheet.getRange(1, 1, 1, lastColumn)
    .setBackground('#b9f79c')
    .setFontColor('blue')
    .setFontWeight('bold');
  targetSheet.setFrozenRows(1);

  const syncTime = latestDate || new Date();
  if (config.lastUpdatedCell) _setConfigDateValue(config.configSheet, config.lastUpdatedCell, syncTime);
}

function _writeGoogleSheetRows(sheet, startRow, values) {
  if (values.length === 0 || values[0].length === 0) return;
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

function _findHeaderIndex(headers, expectedHeader) {
  const expected = _normalizeHeader(expectedHeader);
  return headers.findIndex(header => _normalizeHeader(header) === expected);
}

function _normalizeHeader(value) {
  return _cleanConfigText(value).replace(/\s+/g, ' ').toLowerCase();
}

function _sanitizeValuesForWrite(values, sourceTz) {
  const ssTz = sourceTz || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return values.map(row => row.map(cell => {
    if (cell instanceof Date && !isNaN(cell.getTime())) {
      const timeStr = Utilities.formatDate(cell, ssTz, "HH:mm:ss");
      if (timeStr === "00:00:00") {
        return Utilities.formatDate(cell, ssTz, "yyyy-MM-dd");
      }
      return Utilities.formatDate(cell, ssTz, "yyyy-MM-dd HH:mm:ss");
    }
    return cell;
  }));
}

function _filterGoogleSheetValuesByDate(values, config, sourceTz) {
  if (!config.dateColumn) return values;

  const headers = values[0];
  const dateIndex = headers.indexOf(config.dateColumn);
  if (dateIndex < 0) {
    throw new Error(`Không tìm thấy cột ngày '${config.dateColumn}' trong sheet nguồn '${config.procedureName}'`);
  }

  const startDate = _parseFlexibleDate(config.startDate);
  const endDate = _parseFlexibleDate(config.endDate);
  const filteredRows = values.slice(1).filter(row => {
    const rowDate = _parseFlexibleDate(row[dateIndex], sourceTz);
    if (!rowDate) return false;
    if (startDate && rowDate <= startDate) return false;
    if (endDate && rowDate > endDate) return false;
    return true;
  });

  return [headers].concat(filteredRows);
}

function _getMaxDateFromTableRows(values, dateColumn, tz) {
  if (values.length < 2) return null;
  const headers = values[0];
  const dateIndex = headers.indexOf(dateColumn);
  if (dateIndex < 0) return null;

  let maxDate = null;
  for (const row of values.slice(1)) {
    const parsedDate = _parseFlexibleDate(row[dateIndex], tz);
    if (parsedDate && (!maxDate || parsedDate > maxDate)) maxDate = parsedDate;
  }
  return maxDate;
}

function _getMaxDateFromDomoRows(data, dateColumn) {
  const columns = data.columns || [];
  const rows = data.rows || [];
  const dateIndex = columns.indexOf(dateColumn);
  if (dateIndex < 0) {
    throw new Error(`Không tìm thấy cột ngày '${dateColumn}' trong dữ liệu Domo`);
  }

  let maxDate = null;
  for (const row of rows) {
    const parsedDate = _parseFlexibleDate(row[dateIndex]);
    if (parsedDate && (!maxDate || parsedDate > maxDate)) maxDate = parsedDate;
  }
  return maxDate;
}

function _parseFlexibleDate(value, tz) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    const ssTz = tz || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const str = Utilities.formatDate(value, ssTz, "yyyy-MM-dd HH:mm:ss");
    return new Date(str.replace(/-/g, "/"));
  }

  const text = value.toString().trim();
  if (!text) return null;

  const domoTextDateTime = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?(?:\s*([+-]\d{2}):?(\d{2}))?$/);
  if (domoTextDateTime) {
    const year = Number(domoTextDateTime[1]);
    const month = Number(domoTextDateTime[2]) - 1;
    const day = Number(domoTextDateTime[3]);
    const hour = Number(domoTextDateTime[4] || 0);
    const minute = Number(domoTextDateTime[5] || 0);
    const second = Number(domoTextDateTime[6] || 0);
    const offsetHourText = domoTextDateTime[7];
    const offsetMinute = Number(domoTextDateTime[8] || 0);

    if (offsetHourText) {
      const offsetHour = Number(offsetHourText);
      const utcMillis = Date.UTC(year, month, day, hour, minute, second) - ((offsetHour * 60) + Math.sign(offsetHour) * offsetMinute) * 60000;
      return new Date(utcMillis);
    }

    return new Date(year, month, day, hour, minute, second);
  }

  const localDateTime = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/);
  if (localDateTime) {
    const year = Number(localDateTime[1]);
    const month = Number(localDateTime[2]) - 1;
    const day = Number(localDateTime[3]);
    const hour = Number(localDateTime[4] || 0);
    const minute = Number(localDateTime[5] || 0);
    const second = Number(localDateTime[6] || 0);
    return new Date(year, month, day, hour, minute, second);
  }

  const nativeDate = new Date(text);
  if (!isNaN(nativeDate.getTime())) return nativeDate;

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    const hour = Number(match[4] || 0);
    const minute = Number(match[5] || 0);
    const second = Number(match[6] || 0);
    return new Date(year, month, day, hour, minute, second);
  }

  return null;
}

function setupMultiProcedureTable() {
  const legacyConfigSheet = ss.getSheetByName(SETUP_SHEET_NAME);
  let configSheet = ss.getSheetByName(SETUP_NEW_SHEET_NAME);
  const isNewSheet = !configSheet;
  if (!configSheet) configSheet = ss.insertSheet(SETUP_NEW_SHEET_NAME);
  const wasNewLayout = configSheet.getRange(2, 2).getValue() === 'Platform';
  const migratedRows = wasNewLayout ? [] : _getMigratedSetupRows(configSheet);

  configSheet.getRange('A1:U1')
    .setValues([['MULTI PLATFORM JOB REPORT', '', '', '', '', 'Domo / Google sheet', '', '', 'Status', '', '', '', '', 'Publish state', '', '', 'Updated', '', 'Options', '', 'Sync']])
    .setBackground('#008b18')
    .setFontColor('white')
    .setFontWeight('bold');

  const headers = [
    'Enabled',
    'Platform',
    'Data sheet',
    'Procedure Name / Source sheet',
    'Procedure ID / Dataset ID / Spreadsheet ID',
    'Date column',
    'Date type',
    'Query',
    'created',
    'in-progress',
    'paused',
    'cancelled',
    'completed',
    'production',
    'pre-production',
    'draft',
    'from',
    'to',
    'Lock headers',
    'Include expired jobs',
    'Last Updated',
  ];

  configSheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#4caf50')
    .setFontColor('white')
    .setFontWeight('bold');

  const defaultRow = [
    true,
    'Augmentir',
    DATA_SHEET_NAME,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PROCEDURE.NAME).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PROCEDURE.ID).getValue() : '',
    '',
    'AUTO',
    '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.CREATED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.INPROGRESS).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.PAUSED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.CANCELLED).getValue() : false,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.STATUS.COMPLETED).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PRODUCTION).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.PREPRODUCTION).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.PUBLISH_STATE.DRAFT).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.START_DATE).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.END_DATE).getValue() : '',
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.LOCK_HEADERS).getValue() : true,
    legacyConfigSheet ? legacyConfigSheet.getRange(CONFIG_RANGE.INCLUDE_EXPIRED_JOBS).getValue() : false,
    '',
  ];

  if (migratedRows.length > 0) {
    configSheet.getRange(3, 1, Math.max(configSheet.getLastRow() - 2, 1), headers.length).clearContent();
    configSheet.getRange(3, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }

  const hasProcedure = configSheet.getLastRow() >= 3
    && configSheet.getRange(3, 4, Math.max(configSheet.getLastRow() - 2, 1), 2).getValues().some(row => row[0] || row[1]);
  if ((isNewSheet || !hasProcedure) && migratedRows.length === 0) {
    configSheet.getRange(3, 1, 1, headers.length).setValues([defaultRow]);
  }
  configSheet.getRange(3, 1, 50, 1).insertCheckboxes();
  configSheet.getRange(3, 9, 50, 8).insertCheckboxes();
  configSheet.getRange(3, 19, 50, 2).insertCheckboxes();
  configSheet.setFrozenRows(2);
  configSheet.autoResizeColumns(1, headers.length);
}

function _getMigratedSetupRows(configSheet) {
  if (configSheet.getLastRow() < 3) return [];
  const oldRows = configSheet.getRange(3, 1, configSheet.getLastRow() - 2, 17).getValues();
  return oldRows
    .filter(row => row[2] || row[3])
    .map(row => [
      row[0],
      'Augmentir',
      row[1] || DATA_SHEET_NAME,
      row[2],
      row[3],
      '',
      'AUTO',
      '',
      row[4],
      row[5],
      row[6],
      row[7],
      row[8],
      row[9],
      row[10],
      row[11],
      row[12],
      row[13],
      row[14],
      row[15],
      row[16],
    ]);
}

function createReadmeSheet() {
  const sheetName = 'Readme';
  let readmeSheet = ss.getSheetByName(sheetName);
  if (!readmeSheet) readmeSheet = ss.insertSheet(sheetName);
  readmeSheet.clear();

  const rows = [
    ['CONTROL ROOM - MIXING AUTOMATION README', '', ''],
    ['Last generated', Utilities.formatDate(new Date(), LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'), 'Timezone: ' + LOCAL_TIME_ZONE],
    ['', '', ''],
    ['1. Mục đích tổng thể', 'File này gom dữ liệu từ nhiều nguồn về các sheet DATA, sau đó tổng hợp KPI và gửi thông báo Google Chat.', 'Nguồn hiện hỗ trợ: Augmentir, Domo, Google sheet.'],
    ['2. Sheet cấu hình chính', 'SETUP_NEW', 'Mỗi dòng trong SETUP_NEW là một nguồn dữ liệu cần cập nhật. Bật/tắt bằng cột Enabled.'],
    ['3. Luồng chạy chính', 'Menu Augmentir > Update hoặc trigger gọi hàm update().', 'update() đọc SETUP_NEW, duyệt từng dòng enabled, rồi gọi logic theo Platform.'],
    ['', '', ''],
    ['SETUP_NEW - Cột', 'Ý nghĩa', 'Ghi chú'],
    ['Enabled', 'TRUE/FALSE để bật hoặc bỏ qua dòng cấu hình.', 'FALSE thì dòng đó không chạy.'],
    ['Platform', 'Loại nguồn dữ liệu.', 'Augmentir, Domo, hoặc Google sheet.'],
    ['Data sheet', 'Sheet đích trong file hiện tại.', 'Nếu chưa có, code sẽ tự tạo.'],
    ['Procedure Name / Source sheet', 'Augmentir: Procedure Name. Google sheet: tên sheet nguồn.', 'Domo thường để trống.'],
    ['Procedure ID / Dataset ID / Spreadsheet ID', 'Augmentir: Procedure ID. Domo: Dataset ID. Google sheet: Spreadsheet ID nguồn.', 'Google sheet cần ID file nguồn.'],
    ['Date column', 'Tên cột ngày để lọc dữ liệu.', 'Domo và Google sheet dùng được. Augmentir không cần.'],
    ['Date type', 'Kiểu ngày.', 'AUTO, TEXT, DATE, DATETIME. Domo AUTO/TEXT dùng format yyyy-MM-dd HH:mm:ss.0000000 +07:00.'],
    ['Query', 'SQL dùng cho Domo.', 'Ví dụ: SELECT * FROM table. Để trống sẽ dùng SELECT * FROM table.'],
    ['created/in-progress/paused/cancelled/completed', 'Bộ lọc status cho Augmentir.', 'Domo và Google sheet bỏ qua.'],
    ['production/pre-production/draft', 'Bộ lọc publish state cho Augmentir.', 'Domo và Google sheet bỏ qua.'],
    ['from', 'Mốc bắt đầu lấy dữ liệu incremental.', 'Sau khi chạy thành công sẽ tự cập nhật lên mốc mới nhất; riêng Domo vẫn giữ lịch sử trên sheet bằng upsert.'],
    ['to', 'Mốc kết thúc lấy dữ liệu.', 'Có thể để trống để lấy tới hiện tại/dữ liệu mới nhất.'],
    ['Lock headers', 'Khóa header với Augmentir.', 'TRUE thì không tự thêm cột động mới từ Augmentir.'],
    ['Include expired jobs', 'Có lấy job expired của Augmentir không.', 'TRUE thì lấy cả expired jobs.'],
    ['Last Updated', 'Mốc cập nhật gần nhất để quan sát.', 'Cũng dùng timezone Việt Nam.'],
    ['', '', ''],
    ['Platform: Augmentir', 'Lấy job từ API Augmentir theo Procedure Name hoặc Procedure ID.', 'Dữ liệu được ghi theo cấu trúc DEFAULT_HEADERS và các cột data/unitData động.'],
    ['Augmentir - Incremental', 'Code dùng cột from làm startDate, gọi API theo updatedAt.', 'Timestamp Augmentir là UTC và được ghi về giờ Việt Nam; ngày/ca vận hành lấy từ data.Ngày thực hiện/data.Ca khi có.'],
    ['Augmentir - DATA', 'Nếu Data sheet là DATA, KPI hiện tại sẽ đọc được.', 'KPI script đang đọc sheet DATA cố định.'],
    ['', '', ''],
    ['Platform: Domo', 'Lấy dữ liệu từ Domo Dataset API bằng Dataset ID.', 'Credential đọc từ spreadsheet credential, sheet Domo, ô B1/B2.'],
    ['Domo - Query', 'Code tự thêm WHERE theo Date column/from/to vào Query.', 'Nếu Query có ORDER BY, WHERE sẽ được chèn trước ORDER BY.'],
    ['Domo - Ghi dữ liệu', 'Domo ghi theo upsert mode bằng cột Id để giữ lịch sử giao dịch trong ca.', 'Không clear sheet đích; nếu header lệch sẽ dừng để tránh ghi sai dữ liệu lịch sử.'],
    ['Domo - Ngày', 'Date column ví dụ CreatedDate hoặc ModifiedDate.', 'Domo đọc mốc from/to theo giá trị hiển thị trên SETUP_NEW để tránh lệch giờ; dùng ModifiedDate nếu cần bắt thay đổi trạng thái record cũ.'],
    ['', '', ''],
    ['Platform: Google sheet', 'Copy dữ liệu từ một Spreadsheet khác.', 'Procedure Name / Source sheet là tên sheet nguồn; ID là Spreadsheet ID nguồn.'],
    ['Google sheet - Filter ngày', 'Nếu Date column trống thì copy toàn bộ.', 'Nếu có Date column thì chỉ copy dòng có ngày > from và <= to.'],
    ['Google sheet - Ghi dữ liệu', 'Clear sheet đích rồi ghi header + dòng đã lọc.', 'from cập nhật theo ngày lớn nhất đã copy; nếu không filter ngày thì cập nhật theo thời gian chạy.'],
    ['', '', ''],
    ['KPI_New', 'Sheet KPI được tạo/cập nhật bởi Update KPI (cHÍNH).js.', 'Hàm syncHistoricalKPI_Fast() đọc sheet DATA, Master_data, rồi ghi KPI_New.'],
    ['Google Chat', 'triggerSyncKPI() gọi update(), flush, sau đó sync KPI và gửi thông báo theo ALERT_HOURS.', 'Thông báo phân loại: chưa kiểm tra, đang thực hiện, chờ duyệt, hoàn thành.'],
    ['Ca làm việc', 'Ca 1: 05:55-13:54:59, Ca 2: 13:55-21:54:59, Ca 3: 21:55-05:54:59.', 'Sau 0h tới 05:54:59 vẫn tính về ngày ca trước.'],
    ['', '', ''],
    ['Menu Augmentir', 'Update', 'Chạy toàn bộ nguồn enabled trong SETUP_NEW.'],
    ['Menu Augmentir', 'Create SETUP_NEW', 'Tạo/cập nhật layout SETUP_NEW, migrate layout cũ nếu cần.'],
    ['Menu Augmentir', 'Create README', 'Tạo/cập nhật sheet Readme này.'],
    ['Menu Augmentir > Clear', 'Clear Updated From / Clear Data / Clear Headers.', 'Clear Updated From xóa mốc from trong SETUP_NEW.'],
    ['', '', ''],
    ['Lưu ý vận hành', 'Không đổi tên/cấu trúc header sheet DATA nếu KPI đang phụ thuộc vào DATA.', 'Nếu dùng Data sheet khác, KPI hiện tại sẽ không đọc trừ khi sửa KPI.'],
    ['Lưu ý bảo mật', 'Code hiện có API key/webhook/credential reference.', 'Không chia sẻ project hoặc file code ra ngoài nếu không cần thiết.'],
  ];

  readmeSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  readmeSheet.getRange(1, 1, 1, 3)
    .setBackground('#008b18')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(8, 1, 1, 3)
    .setBackground('#4caf50')
    .setFontColor('white')
    .setFontWeight('bold');
  readmeSheet.getRange(1, 1, rows.length, 3).setWrap(true).setVerticalAlignment('top');
  readmeSheet.setFrozenRows(2);
  readmeSheet.setColumnWidths(1, 1, 220);
  readmeSheet.setColumnWidths(2, 1, 420);
  readmeSheet.setColumnWidths(3, 1, 520);
}

function readSheetList() {
  return ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn(),
  }));
}

function readSheetChunk(sheetName, startRow, rowCount) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Không tìm thấy sheet '${sheetName}'`);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0 || startRow > lastRow) return [];

  const numRows = Math.min(rowCount, lastRow - startRow + 1);
  return sheet.getRange(startRow, 1, numRows, lastColumn).getDisplayValues();
}

function exportAiSheetSnapshots() {
  const folderName = 'AI Sheet Snapshots';
  const folder = _getOrCreateSiblingFolder(folderName);
  const maxRows = 200;

  for (const sheet of ss.getSheets()) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow === 0 || lastColumn === 0) continue;

    const rowCount = Math.min(lastRow, maxRows);
    const values = sheet.getRange(1, 1, rowCount, lastColumn).getDisplayValues();
    const csv = _valuesToCsv(values);
    const fileName = _safeFileName(sheet.getName()) + '.csv';
    _replaceFileInFolder(folder, fileName, csv, MimeType.CSV);
  }

  const manifest = {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    generatedAt: Utilities.formatDate(new Date(), LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
    maxRowsPerSheet: maxRows,
    sheets: readSheetList(),
  };
  _replaceFileInFolder(folder, 'manifest.json', JSON.stringify(manifest, null, 2), MimeType.PLAIN_TEXT);
}

function _getOrCreateSiblingFolder(folderName) {
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}

function _replaceFileInFolder(folder, fileName, content, mimeType) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) files.next().setTrashed(true);
  folder.createFile(fileName, content, mimeType);
}

function _valuesToCsv(values) {
  return values.map(row => row.map(value => {
    const text = value == null ? '' : value.toString();
    return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }).join(',')).join('\r\n');
}

function _safeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

function writeData(actualResults, procedureConfig) {
  const dataSet = actualResults.filter(x => [
    procedureConfig.status.created ? 'created' : undefined,
    procedureConfig.status.inProgress ? 'in-progress' : undefined,
    procedureConfig.status.paused ? 'paused' : undefined,
    procedureConfig.status.cancelled ? 'cancelled' : undefined,
    procedureConfig.status.completed ? 'completed' : undefined,
  ].includes(x.status))
    .filter(x => [
      procedureConfig.publishState.production ? 'production' : undefined,
      procedureConfig.publishState.preProduction ? 'pre-production' : undefined,
      procedureConfig.publishState.draft ? 'draft' : undefined,
      procedureConfig.publishState.draft ? null : undefined,
    ].includes(x.publishState))
    .filter(x => {
      if (procedureConfig.includeExpiredJobs) return true;
      return x.isExpired === false;
    });
  let latestAcceptedUpdatedAt = null;
  dataSet.forEach(job => {
    const updatedAt = new Date(job.updatedAt);
    if (!isNaN(updatedAt.getTime()) && (!latestAcceptedUpdatedAt || updatedAt > latestAcceptedUpdatedAt)) {
      latestAcceptedUpdatedAt = updatedAt;
    }
  });

  const sheet = _ensureDataSheet(procedureConfig.dataSheetName);
  const lockHeaders = procedureConfig.lockHeaders;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  let headers = [...DEFAULT_HEADERS];

  if (lastRow > 0 && lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  if (JSON.stringify(headers.slice(0, DEFAULT_HEADERS.length))
    !== JSON.stringify(DEFAULT_HEADERS)) {
    throw new Error('Headers mặc định đã bị thay đổi, không thể tiếp tục');
  }

  const newData = [];

  const formatValue = (v, key) => {
    if (v instanceof Array) return v.join('\n');
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) {
        // Cố định múi giờ hiển thị cho các cột hệ thống của Augmentir 
        // để tránh lỗi Google Sheet tự đổi theo timezone của file
        if (['createdAt', 'startedAt', 'updatedAt', 'completedAt'].includes(key)) {
          return Utilities.formatDate(d, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
        }
        return d;
      }
    }
    return v;
  }

  for (const job of dataSet) {

    const unitDataList = job.data.unitData?.length ? job.data.unitData : [{}];
    for (let i = 0; i < unitDataList.length; i++) {
      const row = [];
      const unitData = unitDataList[i];
      const obj = {
        ..._.mapValues(_.pick(job, DEFAULT_HEADERS), formatValue),
        unitCount: i + 1,
        ..._.mapValues(_.mapKeys(unitData, (v, k) => `unitData.${k}`), formatValue),
        ..._.mapValues(_.mapKeys(job.data, (v, k) => `data.${k}`), formatValue),
      };
      delete obj.unitData;

      const pickByDefaultHeaders = _.pick(obj, headers);

      for (const h of Object.keys(pickByDefaultHeaders)) {
        const index = headers.indexOf(h);
        if (index >= 0) row[index] = pickByDefaultHeaders[h];
      }

      if (!lockHeaders) {
        const pickByNewHeaders = _.omit(obj, [...headers, 'data.unitData']);
        headers.push(...Object.keys(pickByNewHeaders));
        row.push(...Object.values(pickByNewHeaders));
      }

      newData.push(row);

    }

  }

  const headersRange = sheet.getRange(1, 1, 1, headers.length);
  const defaultHeaderRange = sheet.getRange(1, 1, 1, DEFAULT_HEADERS.length);

  headersRange.setValues([headers]).setBackground('#b9f79c').setFontColor('blue').setFontWeight('bold');
  defaultHeaderRange.setBackground('black').setFontColor('white').setFontWeight('bold');
  _formatAugmentirDateColumns(sheet, headers);
  sheet.setFrozenRows(1);

  let currentData = [];
  if (sheet.getLastRow() >= 2) currentData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  const newRowList = [];
  let updatedCount = 0;

  const jobIdIndex = headers.indexOf('jobId');
  const updatedAtIndex = headers.indexOf('updatedAt');
  const unitCountIndex = headers.indexOf('unitCount');

  for (const _newRow of newData) {
    const jobId = _newRow[jobIdIndex];
    const updatedAt = _newRow[updatedAtIndex];
    const unitCount = _newRow[unitCountIndex];
    const currentRowIndex = currentData.findIndex(x => x[jobIdIndex] == jobId && x[unitCountIndex] == unitCount);
    if (currentRowIndex >= 0) {
      const currentRow = currentData[currentRowIndex];
      const currentRowUpdatedAt = currentRow[updatedAtIndex];

      if (new Date(currentRowUpdatedAt).getTime() !== new Date(updatedAt).getTime()) {
        const paddedRow = [..._newRow];
        paddedRow.length = headers.length;
        currentData[currentRowIndex] = paddedRow;
        updatedCount++;
      }
      continue;
    }

    newRowList.push(_newRow);
  }

  const finalCurrentData = currentData.map(row => {
    if (row.length < headers.length) {
      const padded = [...row];
      padded.length = headers.length;
      return padded;
    }
    return row;
  });

  if (finalCurrentData.length > 0 && updatedCount > 0) {
    sheet.getRange(2, 1, finalCurrentData.length, headers.length).setValues(finalCurrentData);
    console.log('=> ' + updatedCount + ' row(s) updated in bulk!');
  }

  newRowList.forEach(x => x.length = headers.length);

  if (newRowList.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRowList.length, headers.length).setValues(newRowList);
    console.log('=> ' + newRowList.length + ' new row(s) added!')
  }

  return {
    receivedRows: actualResults.length,
    filteredRows: dataSet.length,
    generatedRows: newData.length,
    updatedRows: needToUpdatedList.length,
    addedRows: newRowList.length,
    latestAcceptedUpdatedAt,
  };
}

function _formatAugmentirDateColumns(sheet, headers) {
  ['createdAt', 'startedAt', 'updatedAt', 'completedAt'].forEach(header => {
    const index = headers.indexOf(header);
    if (index < 0 || sheet.getMaxRows() < 2) return;
    sheet.getRange(2, index + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
}


function _getJobsFromProcedure({
  procedureId,
  procedureName,
  startDate,
  endDate,
}) {
  const options =
  {
    method: "POST",
    excludeArchived: true,
    payload: JSON.stringify({
      procedureId: procedureId || '',
      procedureName: procedureName || '',
      startDate: startDate || '',
      endDate: endDate || '',
      status: '',
      excludeArchived: true,
      limit: LIMIT,
      pagingupdate: true,
    }),
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { "X-aug-api-key": APP_KEY, "Content-Type": "application/json" }
  };
  try {
    const target = procedureId ? `procedureId=${procedureId}` : `procedureName=${procedureName || '(blank)'}`;
    Logger.log(`Augmentir request: ${target}, startDate=${startDate || '(blank)'}, endDate=${endDate || '(blank)'}`);
    const response = UrlFetchApp.fetch(SERVER_ENDPOINT + '/rest/v1/GetJobsFromProcedure', options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (statusCode >= 300) {
      throw new Error(`Augmentir API HTTP ${statusCode} for ${target}. Body: ${responseText || '(empty)'}`);
    }
    return JSON.parse(responseText);
  } catch (err) {
    if (err && err.message) throw new Error(err.message);
    throw new Error('Không thể lấy dữ liệu, vui lòng kiểm tra lại thông tin procedure (tên hoặc id)')
  }
}
