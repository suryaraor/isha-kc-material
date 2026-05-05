// Google Apps Script backend for the Isha KC materials board.
// Deploy as a web app and point the UI at the deployment URL.

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action;
  const callback = params.callback;

  try {
    let result;

    if (action === 'getMaterials') {
      result = getMaterials(params.sheetId, params.sheetName);
    } else if (action === 'updatePerson') {
      result = updatePerson(params.sheetId, params.sheetName, params.rowIndex, params.person);
    } else {
      result = { success: false, error: 'Unknown action' };
    }

    return createResponse(result, callback);
  } catch (error) {
    return createResponse({ success: false, error: error.toString() }, callback);
  }
}

function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = params.callback;

  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }

    let result;
    if (payload.action === 'updatePerson' || params.action === 'updatePerson') {
      result = updatePerson(
        payload.sheetId || params.sheetId,
        payload.sheetName || params.sheetName,
        payload.rowIndex,
        payload.person
      );
    } else {
      result = { success: false, error: 'Unknown action' };
    }

    return createResponse(result, callback);
  } catch (error) {
    return createResponse({ success: false, error: error.toString() }, callback);
  }
}

function createResponse(result, callback) {
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMaterials(sheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = getTargetSheet(spreadsheet, sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, error: 'No data found' };
  }

  const headers = data[0];
  const columns = getMaterialColumnMap(headers);
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = normalizeCell(row[columns.item]);
    const group = normalizeCell(row[columns.group]);
    const person = normalizeCell(row[columns.person]);
    const asOfDate = formatCell(row[columns.asOfDate]);

    if (!item && !group && !person && !asOfDate) {
      continue;
    }

    rows.push({
      rowIndex: i + 1,
      item: item,
      group: group,
      person: person,
      asOfDate: asOfDate
    });
  }

  return {
    success: true,
    data: rows,
    total: rows.length,
    people: uniqueValues(rows.map(row => row.person)),
    groups: uniqueValues(rows.map(row => row.group))
  };
}

function updatePerson(sheetId, sheetName, rowIndexFromClient, personName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = getTargetSheet(spreadsheet, sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return { success: false, error: 'No data found' };
  }

  const rowIndex = Number(rowIndexFromClient);
  if (!Number.isFinite(rowIndex) || rowIndex < 2 || rowIndex > data.length) {
    return { success: false, error: 'Invalid row index' };
  }

  const nextPerson = normalizeCell(personName);
  if (!nextPerson) {
    return { success: false, error: 'Person name is required' };
  }

  const columns = getMaterialColumnMap(data[0]);
  sheet.getRange(rowIndex, columns.person + 1).setValue(nextPerson);

  return {
    success: true,
    rowIndex: rowIndex,
    person: nextPerson,
    sheetName: sheet.getName()
  };
}

function getTargetSheet(spreadsheet, sheetName) {
  if (sheetName) {
    const namedSheet = spreadsheet.getSheetByName(sheetName);
    if (!namedSheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }
    return namedSheet;
  }

  const sheets = spreadsheet.getSheets();
  if (!sheets || sheets.length === 0) {
    throw new Error('No sheets found in spreadsheet');
  }

  return sheets[0];
}

function getMaterialColumnMap(headers) {
  const normalizedHeaders = headers.map(header => normalizeHeader(header));
  const itemIndex = normalizedHeaders.indexOf('item');
  const groupIndex = normalizedHeaders.indexOf('group');
  const personIndex = normalizedHeaders.indexOf('person');
  const asOfDateIndex = normalizedHeaders.findIndex(header => header === 'asofdate' || header === 'asof');

  if (itemIndex === -1 || groupIndex === -1 || personIndex === -1 || asOfDateIndex === -1) {
    throw new Error('Missing required columns: Item, Group, Person, As of date');
  }

  return {
    item: itemIndex,
    group: groupIndex,
    person: personIndex,
    asOfDate: asOfDateIndex
  };
}

function normalizeHeader(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

function normalizeCell(value) {
  return String(value || '').trim();
}

function formatCell(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value).trim();
}

function uniqueValues(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((left, right) => {
    return String(left).localeCompare(String(right), undefined, { sensitivity: 'base' });
  });
}

function testGetMaterials() {
  const result = getMaterials('19yApUs-AasdcS1oerxjN-OsngO9vhkavaxXKUjqpuUY');
  Logger.log(JSON.stringify(result, null, 2));
}
