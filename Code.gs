/**
 * SK Tech interview portal server-side application.
 *
 * Public functions in this file are called either by the web application
 * (doGet/getInterviewData) or by an installed spreadsheet form-submit trigger
 * (onFormSubmit).
 */

const CONFIG = Object.freeze({
  spreadsheetId: '1VQF3CI8MTBz4vAC4R-13_XEfY774h1Q7vxhBsSxA2tU',
  interviewSheetName: 'Pannel List',
  responseSheetName: 'Form Responses 3_Candiate details',
  candidateSheetName: 'Candiate Details',
  timeZone: 'Asia/Kolkata',
  supportEmail: 'support@mysktech.com',
  responseEmailColumn: 2,
  responseNameColumn: 3,
  headerRow: 1,
  interviewColumnCount: 9,
  formLockTimeoutMs: 30000
});

const INTERVIEW_HEADERS = Object.freeze([
  'Sk Tech Register ID',
  'Full Name',
  'Round',
  'Technologies Required*',
  'Interview Company',
  'Interview Date',
  'Interview Time (From) or If Time Not confirmed plz select 00:00 like Assessment',
  'Interview Time (To) or If Time Not confirmed plz select 00:00 like Assessment',
  'Batch'
]);

/**
 * Serves the portal shell. Interview data is loaded by getInterviewData().
 *
 * @return {HtmlOutput}
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Html')
    .setTitle('SK Tech Interview Panel Sheet')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns the current and upcoming interview rows for the portal.
 * Values returned to the browser are display-safe strings; no Date objects are
 * sent over google.script.run.
 *
 * @return {{headers: string[], rows: Object[], generatedAt: string, timeZone: string}}
 */
function getInterviewData() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(CONFIG.interviewSheetName);

  if (!sheet) {
    throw new Error(
      'Interview sheet "' + CONFIG.interviewSheetName + '" was not found. ' +
      'Available sheets: ' + spreadsheet.getSheets().map(function(item) {
        return item.getName();
      }).join(', ')
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.headerRow) {
    throw new Error(
      'Interview sheet "' + CONFIG.interviewSheetName + '" is empty. ' +
      'Add the required header row before loading the portal.'
    );
  }

  const actualHeaders = sheet
    .getRange(CONFIG.headerRow, 1, 1, CONFIG.interviewColumnCount)
    .getDisplayValues()[0];
  validateInterviewHeaders_(actualHeaders);

  const result = {
    headers: INTERVIEW_HEADERS.slice(),
    rows: [],
    generatedAt: Utilities.formatDate(
      new Date(),
      CONFIG.timeZone,
      'M/d/yyyy hh:mm:ss a'
    ),
    timeZone: CONFIG.timeZone
  };

  if (lastRow === CONFIG.headerRow) {
    return result;
  }

  const rowCount = lastRow - CONFIG.headerRow;
  const range = sheet.getRange(
    CONFIG.headerRow + 1,
    1,
    rowCount,
    CONFIG.interviewColumnCount
  );
  const rawRows = range.getValues();
  const displayRows = range.getDisplayValues();
  const todayKey = Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyy-MM-dd');
  const preparedRows = [];

  rawRows.forEach(function(rawRow, index) {
    const displayRow = displayRows[index];
    if (isBlankRow_(displayRow)) {
      return;
    }

    const interviewDate = parseDate_(rawRow[5], displayRow[5]);
    if (!interviewDate || interviewDate.key < todayKey) {
      return;
    }

    const startTime = parseTime_(rawRow[6], displayRow[6]);
    const endTime = parseTime_(rawRow[7], displayRow[7]);
    const values = displayRow.slice(0, CONFIG.interviewColumnCount).map(function(value) {
      return String(value == null ? '' : value);
    });

    values[5] = interviewDate.display;
    values[6] = startTime ? startTime.display : values[6];
    values[7] = endTime ? endTime.display : values[7];

    preparedRows.push({
      values: values,
      dateKey: interviewDate.key,
      startSeconds: startTime ? startTime.seconds : null,
      endSeconds: endTime ? endTime.seconds : null,
      sourceOrder: index
    });
  });

  preparedRows.sort(function(left, right) {
    if (left.dateKey !== right.dateKey) {
      return left.dateKey < right.dateKey ? -1 : 1;
    }

    const leftStart = left.startSeconds == null ? Number.POSITIVE_INFINITY : left.startSeconds;
    const rightStart = right.startSeconds == null ? Number.POSITIVE_INFINITY : right.startSeconds;
    if (leftStart !== rightStart) {
      return leftStart - rightStart;
    }

    return left.sourceOrder - right.sourceOrder;
  });

  // Preserve the existing panel behavior: the earliest appointment keeps its
  // normal status and only later appointments that overlap it are conflicts.
  const previousIntervalsByDate = Object.create(null);
  result.rows = preparedRows.map(function(row) {
    let status = row.dateKey === todayKey ? 'today' : 'upcoming';
    const isValidInterval =
      row.startSeconds != null &&
      row.endSeconds != null &&
      row.endSeconds > row.startSeconds;
    const previousIntervals = previousIntervalsByDate[row.dateKey] || [];

    if (isValidInterval) {
      const overlapsEarlierRow = previousIntervals.some(function(interval) {
        return row.startSeconds < interval.end && row.endSeconds > interval.start;
      });

      if (overlapsEarlierRow) {
        status = 'conflict';
      }

      previousIntervals.push({
        start: row.startSeconds,
        end: row.endSeconds
      });
      previousIntervalsByDate[row.dateKey] = previousIntervals;
    }

    return {
      values: row.values,
      status: status
    };
  });

  return result;
}

/**
 * Copies a qualifying form response to the candidate sheet and emails the
 * candidate their registration ID. Configure this as an installed form-submit
 * trigger; a simple trigger cannot authorize GmailApp.
 *
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function onFormSubmit(e) {
  if (!e || !e.range || typeof e.range.getSheet !== 'function') {
    throw new Error(
      'onFormSubmit requires an installed spreadsheet form-submit trigger event.'
    );
  }

  const submittedSheet = e.range.getSheet();
  if (submittedSheet.getName() !== CONFIG.responseSheetName) {
    console.log('Ignored form submission from: ' + submittedSheet.getName());
    return;
  }

  const eventSpreadsheet = submittedSheet.getParent();
  if (eventSpreadsheet.getId() !== CONFIG.spreadsheetId) {
    console.warn('Ignored form submission from an unexpected spreadsheet.');
    return;
  }

  const responseRow = e.range.getRow();
  if (responseRow <= CONFIG.headerRow) {
    console.warn('Ignored form submission event for the header row.');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.formLockTimeoutMs)) {
    throw new Error(
      'Registration is busy processing another submission. Please retry this response.'
    );
  }

  try {
    const responseValues = submittedSheet
      .getRange(responseRow, CONFIG.responseEmailColumn, 1, 2)
      .getDisplayValues()[0];
    const email = String(responseValues[0] || '').trim();
    const name = String(responseValues[1] || '').trim();

    if (!email || !name) {
      console.warn('Skipped response row ' + responseRow + ': name or email is blank.');
      return;
    }

    if (!isPlausibleEmail_(email)) {
      console.warn('Skipped response row ' + responseRow + ': invalid email address.');
      return;
    }

    const candidateSheet = eventSpreadsheet.getSheetByName(CONFIG.candidateSheetName);
    if (!candidateSheet) {
      throw new Error(
        'Candidate sheet "' + CONFIG.candidateSheetName + '" was not found.'
      );
    }

    if (candidateSheet.getLastRow() < CONFIG.headerRow) {
      throw new Error(
        'Candidate sheet "' + CONFIG.candidateSheetName + '" must contain a header row.'
      );
    }

    const skId = getNextNumericSkId_(candidateSheet);
    const nextRow = candidateSheet.getLastRow() + 1;
    candidateSheet.getRange(nextRow, 1, 1, 2).setValues([
      [skId, protectSheetText_(name)]
    ]);
    SpreadsheetApp.flush();

    const subject = 'SK Tech Registration Successful';
    const body = [
      'Hello ' + name + ',',
      '',
      'Welcome to SK Tech.',
      '',
      'Your registration was successful.',
      '',
      'Your SK Tech ID: SK' + skId,
      '',
      'Candidate Name: ' + name,
      '',
      'Please use this ID for future communication.',
      '',
      'Regards,',
      'SK Tech Team'
    ].join('\n');

    GmailApp.sendEmail(email, subject, body, getMailOptions_());
    console.log('Registered response row ' + responseRow + ' as SK' + skId + '.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Logs and returns all sheet names to help verify deployment configuration.
 *
 * @return {string[]}
 */
function checkSheetNames() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const names = spreadsheet.getSheets().map(function(sheet) {
    return sheet.getName();
  });

  names.forEach(function(name) {
    console.log('Sheet Name: [' + name + ']');
  });

  return names;
}

function validateInterviewHeaders_(actualHeaders) {
  const problems = [];

  INTERVIEW_HEADERS.forEach(function(expected, index) {
    const actual = actualHeaders[index] == null ? '' : String(actualHeaders[index]);
    if (normalizeHeader_(actual) !== normalizeHeader_(expected)) {
      problems.push(
        'column ' + columnLetter_(index + 1) + ': expected "' + expected +
        '", found "' + (actual || '(blank)') + '"'
      );
    }
  });

  if (problems.length) {
    throw new Error(
      'The first nine headers in "' + CONFIG.interviewSheetName +
      '" do not match the interview portal. Please keep these columns in the ' +
      'required order. ' + problems.join('; ') + '.'
    );
  }
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/\s*\*\s*/g, '*')
    .toLowerCase();
}

function columnLetter_(columnNumber) {
  let value = columnNumber;
  let result = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}

function isBlankRow_(row) {
  return row.every(function(value) {
    return String(value == null ? '' : value).trim() === '';
  });
}

function parseDate_(rawValue, displayValue) {
  if (isValidDate_(rawValue)) {
    const key = Utilities.formatDate(rawValue, CONFIG.timeZone, 'yyyy-MM-dd');
    return {
      key: key,
      display: formatDateKey_(key)
    };
  }

  const candidates = [rawValue, displayValue];
  for (let i = 0; i < candidates.length; i += 1) {
    const parsed = parseDateString_(candidates[i]);
    if (parsed) {
      return parsed;
    }
  }

  return null;
}

function parseDateString_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    return null;
  }

  let year;
  let month;
  let day;
  let match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:\D.*)?$/);

  if (match) {
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  } else {
    match = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})(?:\D.*)?$/);
    if (!match) {
      return null;
    }

    const first = Number(match[1]);
    const second = Number(match[2]);
    year = Number(match[3]);

    // The portal's canonical date convention is M/d/yyyy. When the first
    // component cannot be a month, safely accept an unambiguous d/M/yyyy value.
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    } else {
      month = first;
      day = second;
    }
  }

  if (!isValidCalendarDate_(year, month, day)) {
    return null;
  }

  const key = [
    String(year).padStart(4, '0'),
    String(month).padStart(2, '0'),
    String(day).padStart(2, '0')
  ].join('-');

  return {
    key: key,
    display: month + '/' + day + '/' + year
  };
}

function isValidCalendarDate_(year, month, day) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    year < 1900 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {
    return false;
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;
}

function formatDateKey_(key) {
  const parts = key.split('-');
  return Number(parts[1]) + '/' + Number(parts[2]) + '/' + parts[0];
}

function parseTime_(rawValue, displayValue) {
  if (isValidDate_(rawValue)) {
    const twentyFourHourTime = Utilities.formatDate(
      rawValue,
      CONFIG.timeZone,
      'HH:mm:ss'
    );
    return timePartsToResult_(twentyFourHourTime);
  }

  if (typeof rawValue === 'number' && Number.isFinite(rawValue)) {
    let fraction = rawValue % 1;
    if (fraction < 0) {
      fraction += 1;
    }
    const seconds = Math.round(fraction * 86400) % 86400;
    return {
      seconds: seconds,
      display: formatTimeSeconds_(seconds)
    };
  }

  const candidates = [rawValue, displayValue];
  for (let i = 0; i < candidates.length; i += 1) {
    const parsed = parseTimeString_(candidates[i]);
    if (parsed) {
      return parsed;
    }
  }

  return null;
}

function parseTimeString_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    return null;
  }

  const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/i);
  if (!match) {
    return null;
  }

  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = match[3] == null ? 0 : Number(match[3]);
  const meridiem = match[4] ? match[4].toUpperCase() : '';

  if (minute > 59 || second > 59) {
    return null;
  }

  if (meridiem) {
    if (hour < 1 || hour > 12) {
      return null;
    }
    if (hour === 12) {
      hour = 0;
    }
    if (meridiem === 'PM') {
      hour += 12;
    }
  } else if (hour > 23) {
    return null;
  }

  const seconds = hour * 3600 + minute * 60 + second;
  return {
    seconds: seconds,
    display: formatTimeSeconds_(seconds)
  };
}

function timePartsToResult_(twentyFourHourTime) {
  const parts = twentyFourHourTime.split(':').map(Number);
  const seconds = parts[0] * 3600 + parts[1] * 60 + parts[2];
  return {
    seconds: seconds,
    display: formatTimeSeconds_(seconds)
  };
}

function formatTimeSeconds_(seconds) {
  const hour24 = Math.floor(seconds / 3600);
  const minute = Math.floor((seconds % 3600) / 60);
  const second = seconds % 60;
  const meridiem = hour24 >= 12 ? 'PM' : 'AM';
  let hour12 = hour24 % 12;
  if (hour12 === 0) {
    hour12 = 12;
  }

  return String(hour12).padStart(2, '0') + ':' +
    String(minute).padStart(2, '0') + ':' +
    String(second).padStart(2, '0') + ' ' + meridiem;
}

function isValidDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' &&
    !Number.isNaN(value.getTime());
}

function isPlausibleEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getNextNumericSkId_(candidateSheet) {
  const lastRow = candidateSheet.getLastRow();
  let highestId = 0;

  if (lastRow > CONFIG.headerRow) {
    candidateSheet
      .getRange(CONFIG.headerRow + 1, 1, lastRow - CONFIG.headerRow, 1)
      .getValues()
      .forEach(function(row) {
        const match = String(row[0] || '').trim().match(/^(?:SK\s*)?(\d+)$/i);
        if (match) {
          const numericId = Number(match[1]);
          if (Number.isSafeInteger(numericId) && numericId > highestId) {
            highestId = numericId;
          }
        }
      });
  }

  if (highestId >= Number.MAX_SAFE_INTEGER) {
    throw new Error('The numeric SK Tech registration ID range has been exhausted.');
  }

  // onFormSubmit holds the script lock while scanning and writing, so concurrent
  // submissions cannot choose the same next sequential ID.
  return highestId + 1;
}

function protectSheetText_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function getMailOptions_() {
  const options = {
    bcc: CONFIG.supportEmail,
    name: 'SK Tech',
    replyTo: CONFIG.supportEmail
  };

  try {
    const aliases = GmailApp.getAliases().map(function(alias) {
      return alias.toLowerCase();
    });
    if (aliases.indexOf(CONFIG.supportEmail.toLowerCase()) !== -1) {
      options.from = CONFIG.supportEmail;
      delete options.replyTo;
    }
  } catch (error) {
    console.warn('Could not inspect Gmail aliases; using reply-to instead.');
  }

  return options;
}
