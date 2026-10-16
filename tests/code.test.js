'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const CODE_PATH = path.resolve(__dirname, '..', 'Code.gs');
const CODE_SOURCE = fs.readFileSync(CODE_PATH, 'utf8');
const FIXED_NOW = '2026-10-16T06:30:00.000Z';

const VERIFIED_HEADERS = [
  'Sk Tech Register ID',
  'Full Name',
  'Round',
  'Technologies Required*',
  'Interview Company',
  'Interview Date',
  'Interview Time (From) or If Time Not confirmed plz select 00:00 like Assessment',
  'Interview Time (To) or If Time Not confirmed plz select 00:00 like Assessment',
  'Batch'
];

function createFixedDate(nowIso) {
  const NativeDate = Date;

  return class FixedDate extends NativeDate {
    constructor(...args) {
      super(...(args.length ? args : [nowIso]));
    }

    static now() {
      return new NativeDate(nowIso).getTime();
    }
  };
}

function getZonedParts(date, timeZone) {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23'
  });

  return Object.fromEntries(
    formatter
      .formatToParts(date)
      .filter((part) => part.type !== 'literal')
      .map((part) => [part.type, part.value])
  );
}

function formatAppsScriptDate(date, timeZone, pattern) {
  const parts = getZonedParts(date, timeZone);
  const hour24 = Number(parts.hour) % 24;
  const hour12 = hour24 % 12 || 12;
  const meridiem = hour24 >= 12 ? 'PM' : 'AM';

  switch (pattern) {
    case 'yyyy-MM-dd':
      return `${parts.year}-${parts.month}-${parts.day}`;
    case 'M/d/yyyy':
      return `${Number(parts.month)}/${Number(parts.day)}/${parts.year}`;
    case 'M/d/yyyy hh:mm:ss a':
      return `${Number(parts.month)}/${Number(parts.day)}/${parts.year} ` +
        `${String(hour12).padStart(2, '0')}:${parts.minute}:${parts.second} ${meridiem}`;
    case 'HH:mm:ss':
      return `${String(hour24).padStart(2, '0')}:${parts.minute}:${parts.second}`;
    case 'hh:mm:ss a':
      return `${String(hour12).padStart(2, '0')}:${parts.minute}:${parts.second} ${meridiem}`;
    default:
      throw new Error(`Unsupported Utilities.formatDate pattern in test: ${pattern}`);
  }
}

function displayValue(value) {
  if (value == null) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return '';
  }
  return String(value);
}

function createInterviewSheet({
  headers = VERIFIED_HEADERS,
  rawRows = [],
  displayRows = rawRows.map((row) => row.map(displayValue))
} = {}) {
  return {
    getName() {
      return 'Pannel List';
    },
    getLastRow() {
      return rawRows.length + 1;
    },
    getRange(startRow, startColumn, rowCount, columnCount) {
      const columnOffset = startColumn - 1;

      if (startRow === 1) {
        return {
          getDisplayValues() {
            return [headers.slice(columnOffset, columnOffset + columnCount)];
          }
        };
      }

      const rowOffset = startRow - 2;
      const select = (rows) => rows
        .slice(rowOffset, rowOffset + rowCount)
        .map((row) => row.slice(columnOffset, columnOffset + columnCount));

      return {
        getValues() {
          return select(rawRows);
        },
        getDisplayValues() {
          return select(displayRows);
        }
      };
    }
  };
}

function createSpreadsheet(sheet) {
  return {
    getSheetByName(name) {
      return name === 'Pannel List' ? sheet : null;
    },
    getSheets() {
      return [sheet];
    }
  };
}

function loadCode({ spreadsheet, now = FIXED_NOW } = {}) {
  const context = vm.createContext({
    Date: createFixedDate(now),
    SpreadsheetApp: {
      openById() {
        if (!spreadsheet) {
          throw new Error('No spreadsheet mock supplied for this test.');
        }
        return spreadsheet;
      },
      flush() {}
    },
    Utilities: {
      formatDate: formatAppsScriptDate
    },
    HtmlService: {},
    GmailApp: {},
    LockService: {},
    console: {
      log() {},
      warn() {}
    }
  });

  vm.runInContext(CODE_SOURCE, context, { filename: CODE_PATH });
  return context;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function interviewRow(id, date, start, end) {
  return [
    id,
    `${id} Candidate`,
    'Round 1',
    'Data Engineer',
    'Example Company',
    date,
    start,
    end,
    '2'
  ];
}

test('header validation accepts verified labels with supported whitespace and case variations', () => {
  const context = loadCode();
  const actualHeaders = [
    '  SK   TECH register id  ',
    '\nFull\tName ',
    ' ROUND ',
    'Technologies\u00a0Required  *  ',
    ' Interview   Company ',
    ' interview date ',
    ' Interview Time (From) or If Time Not confirmed plz select 00:00 like Assessment ',
    '\tInterview Time (To) or If Time Not confirmed plz select 00:00 like Assessment\n',
    ' batch '
  ];

  assert.doesNotThrow(() => context.validateInterviewHeaders_(actualHeaders));
});

test('header validation reports a mismatched verified column', () => {
  const context = loadCode();
  const actualHeaders = VERIFIED_HEADERS.slice();
  actualHeaders[3] = 'Technology Stack';

  assert.throws(
    () => context.validateInterviewHeaders_(actualHeaders),
    (error) => {
      assert.match(error.message, /column D/);
      assert.match(error.message, /Technologies Required\*/);
      assert.match(error.message, /Technology Stack/);
      return true;
    }
  );
});

test('getInterviewData filters past rows, sorts by date and start time, and assigns sequential statuses', () => {
  const rawRows = [
    interviewRow('SK-Past', '10/15/2026', '09:00 AM', '10:00 AM'),
    interviewRow('SK-Future-Late', '10/17/2026', '04:15 PM', '05:00 PM'),
    interviewRow('SK-Today-Conflict', '10/16/2026', '04:30 PM', '05:30 PM'),
    interviewRow('SK-Today-First', '10/16/2026', '04:00 PM', '05:00 PM'),
    interviewRow('SK-Future-Early', '10/17/2026', '02:30 PM', '03:30 PM'),
    ['', '', '', '', '', '', '', '', '']
  ];
  const sheet = createInterviewSheet({ rawRows });
  const context = loadCode({ spreadsheet: createSpreadsheet(sheet) });

  const result = plain(context.getInterviewData());

  assert.deepEqual(
    result.rows.map((row) => row.values[0]),
    ['SK-Today-First', 'SK-Today-Conflict', 'SK-Future-Early', 'SK-Future-Late']
  );
  assert.deepEqual(
    result.rows.map((row) => row.status),
    ['today', 'conflict', 'upcoming', 'upcoming']
  );
  assert.equal(result.rows[0].values[5], '10/16/2026');
  assert.equal(result.rows[0].values[6], '04:00:00 PM');
  assert.equal(result.rows[0].values[7], '05:00:00 PM');
  assert.equal(result.timeZone, 'Asia/Kolkata');
});

test('date and time values are normalized to portal display formats', () => {
  const rawRows = [[
    'SK9',
    'Formatted Candidate',
    'Round 2',
    'Full Stack Developer',
    'Example Company',
    new Date('2026-10-17T18:30:00.000Z'),
    0.5,
    '13:05',
    4
  ]];
  const displayRows = [[
    'SK9',
    'Formatted Candidate',
    'Round 2',
    'Full Stack Developer',
    'Example Company',
    '10/18/2026',
    '12:00 PM',
    '1:05 PM',
    '4'
  ]];
  const sheet = createInterviewSheet({ rawRows, displayRows });
  const context = loadCode({ spreadsheet: createSpreadsheet(sheet) });

  const result = plain(context.getInterviewData());

  assert.deepEqual(result.rows[0], {
    values: [
      'SK9',
      'Formatted Candidate',
      'Round 2',
      'Full Stack Developer',
      'Example Company',
      '10/18/2026',
      '12:00:00 PM',
      '01:05:00 PM',
      '4'
    ],
    status: 'upcoming'
  });

  assert.deepEqual(
    plain(context.parseDate_('17/10/2026', '')),
    { key: '2026-10-17', display: '10/17/2026' }
  );
  assert.deepEqual(
    plain(context.parseTime_('12:00 AM', '')),
    { seconds: 0, display: '12:00:00 AM' }
  );
  assert.equal(context.parseDate_('2/30/2026', ''), null);
});

test('getNextNumericSkId_ advances past numeric and SK-prefixed values', () => {
  const candidateIds = [1, 'SK2', ' sk 17 ', 'SK009', 42, 'not-an-id', 'SK-100'];
  const candidateSheet = {
    getLastRow() {
      return candidateIds.length + 1;
    },
    getRange(startRow, startColumn, rowCount, columnCount) {
      assert.equal(startRow, 2);
      assert.equal(startColumn, 1);
      assert.equal(rowCount, candidateIds.length);
      assert.equal(columnCount, 1);
      return {
        getValues() {
          return candidateIds.map((id) => [id]);
        }
      };
    }
  };
  const context = loadCode();

  assert.equal(context.getNextNumericSkId_(candidateSheet), 43);
});

test('getNextNumericSkId_ starts at one when the candidate sheet has only a header', () => {
  const candidateSheet = {
    getLastRow() {
      return 1;
    }
  };
  const context = loadCode();

  assert.equal(context.getNextNumericSkId_(candidateSheet), 1);
});
