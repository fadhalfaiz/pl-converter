/**
 * Pure PL rules and calculations. This file does not call Google services.
 * Posting dates use ISO yyyy-MM-dd; row numbers always refer to the source GL.
 */
var PLCore = (function () {
  'use strict';

  var REQUIRED_HEADERS = [
    'Project', 'GL Account', 'Posting Date',
    'In profit center local currency', 'Text'
  ];
  var ITEM_ROWS = [4, 7, 8, 9];
  for (var itemRow = 11; itemRow <= 32; itemRow++) ITEM_ROWS.push(itemRow);
  var DISPLAY_ROWS = [4, 6, 7, 8, 9, 10];
  for (var displayRow = 11; displayRow <= 34; displayRow++) DISPLAY_ROWS.push(displayRow);
  DISPLAY_ROWS = DISPLAY_ROWS.concat([36, 37, 39, 41, 42, 43, 44, 46, 48]);
  var MONTH_NAMES = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DES'];
  var EXCLUDED_ACCOUNTS = ['DIRECT LABOR', 'SALARY AND COMPENSATION', 'DEPRECIATION & AMORTIZATION'];
  var MONTH_ALIASES = [
    ['january', 'januari', 'jan'], ['february', 'februari', 'feb'],
    ['march', 'maret', 'mar'], ['april', 'apr'], ['mei', 'may'],
    ['june', 'juni', 'jun'], ['july', 'juli', 'jul'],
    ['agustus', 'august', 'agust', 'agst', 'agu', 'ags', 'agt', 'aug'],
    ['september', 'sept', 'sep'], ['october', 'oktober', 'okt', 'oct'],
    ['november', 'nov'], ['december', 'desember', 'des', 'dec']
  ];
  var aliasToMonth = Object.create(null);
  MONTH_ALIASES.forEach(function (aliases, month) {
    aliases.forEach(function (alias) { aliasToMonth[alias] = month + 1; });
  });
  var monthPatternSource = '(^|[^A-Za-z])(' + Object.keys(aliasToMonth).sort(function (a, b) {
    return b.length - a.length;
  }).join('|') + ')(?![A-Za-z])';

  function cleanText(value) {
    return value == null ? '' : String(value).replace(/\t/g, ' ').replace(/["“”]/g, '').trim().replace(/\s+/g, ' ');
  }

  function validateReport(reportYear, ytdMonth) {
    if (!Number.isInteger(reportYear) || reportYear < 1900 || reportYear > 9999) {
      throw new Error('Tahun Laporan tidak valid.');
    }
    if (!Number.isInteger(ytdMonth) || ytdMonth < 1 || ytdMonth > 12) {
      throw new Error('Bulan YTD harus antara Januari dan Desember.');
    }
  }

  function postingParts(postingISO) {
    var match = typeof postingISO === 'string' && /^(\d{4})-(\d{2})-(\d{2})$/.exec(postingISO);
    if (!match) throw new Error('Posting Date harus berupa tanggal Excel yang valid.');
    var year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
    var leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
    var days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (year < 1 || month < 1 || month > 12 || day < 1 || day > days[month - 1]) {
      throw new Error('Posting Date harus berupa tanggal Excel yang valid.');
    }
    return { year: year, month: month, day: day };
  }

  function excluded(account) {
    return EXCLUDED_ACCOUNTS.indexOf(String(account == null ? '' : account).trim().toUpperCase()) !== -1;
  }

  function periodResult(month, year, carry) {
    return {
      periodMonth: month,
      periodYear: year,
      carryOverYear: carry,
      period: carry !== null ? 'CO ' + carry : String(month).padStart(2, '0') + '-' + year
    };
  }

  function fourDigitYear(text) {
    var match = /(?:^|[^\d])(20\d{2})(?!\d)/.exec(text);
    return match ? Number(match[1]) : null;
  }

  function parsePeriod(text, postingISO, account, reportYear, ytdMonth) {
    if (ytdMonth == null) ytdMonth = 12;
    validateReport(reportYear, ytdMonth);
    var posting = postingParts(postingISO);
    var normalized = cleanText(text);
    var matches = [], match;
    var pattern = new RegExp(monthPatternSource, 'gi');
    while ((match = pattern.exec(normalized)) !== null) {
      matches.push({ month: aliasToMonth[match[2].toLowerCase()], end: match.index + match[0].length });
    }
    var last = matches.length ? matches[matches.length - 1] : null;
    var month = last ? last.month : null;
    var year = null, inferredYear = false, notes = [];
    if (last) {
      var suffix = normalized.slice(last.end);
      var short = /^\s*[-/.'’]*\s*(\d{2})(?!\d)/.exec(suffix);
      year = short ? 2000 + Number(short[1]) : fourDigitYear(suffix);
      if (year === null) year = fourDigitYear(normalized);
      if (year === null) {
        inferredYear = true;
        year = month > ytdMonth ? reportYear - 1 : reportYear;
        notes.push('Bulan tanpa tahun -> tahun dianggap ' + year);
      }
      if (matches.length > 1) {
        notes.push(year === reportYear ? 'Rentang bulan tahun laporan -> ikut posting date' : 'Rentang bulan, dipakai bulan akhir');
      }
      if (!inferredYear && year < posting.year && month <= posting.month) {
        notes.push('Cek: tahun ' + year + ' tapi bulan <= bulan posting, kemungkinan salah ketik');
      }
    } else {
      year = fourDigitYear(normalized);
      if (year === null) {
        var trailing = /(?:^|\s)(25|26)\s*[-–]?\s*$/.exec(normalized);
        if (trailing && /[A-Za-z]/.test(normalized.slice(0, trailing.index))) {
          year = 2000 + Number(trailing[1]);
          notes.push('Tahun dibaca dari angka di akhir teks, mohon konfirmasi');
        }
      }
    }

    var result;
    if (excluded(account)) {
      result = periodResult(posting.month, posting.year, null);
      notes = ['GL dikecualikan -> ikut posting date'];
    } else if (year !== null && year < reportYear) {
      result = periodResult(null, null, year);
    } else if (matches.length > 1 && year === reportYear) {
      month = null;
      result = periodResult(posting.month, posting.year, null);
    } else if (month !== null) {
      result = periodResult(month, year, null);
    } else {
      result = periodResult(posting.month, posting.year, null);
    }
    result.textMonth = month;
    result.textYear = year;
    result.notes = notes.join('; ');
    return result;
  }

  function effectivePeriod(postingISO, account, reportYear, textMonth, textYear) {
    validateReport(reportYear, 12);
    var posting = postingParts(postingISO);
    if (excluded(account)) return periodResult(posting.month, posting.year, null);
    if (textYear != null && textYear < reportYear) return periodResult(null, null, textYear);
    if (textYear != null && textMonth != null) return periodResult(textMonth, textYear, null);
    return periodResult(posting.month, posting.year, null);
  }

  function blank(value) { return value == null || value === ''; }

  function parseRows(rows, reportYear, ytdMonth) {
    if (ytdMonth == null) ytdMonth = 12;
    validateReport(reportYear, ytdMonth);
    if (!Array.isArray(rows) || !rows.length || !Array.isArray(rows[0])) {
      throw new Error('File GL tidak memiliki baris data.');
    }
    if (rows[0].length !== 5) {
      throw new Error('File GL harus memiliki tepat lima kolom sumber agar kolom bantu F:I tersedia.');
    }
    var sourceHeaders = Object.create(null);
    rows[0].forEach(function (name, index) {
      if (!blank(name)) sourceHeaders[String(name).trim().toLowerCase()] = index + 1;
    });
    var missing = REQUIRED_HEADERS.filter(function (name) { return !sourceHeaders[name.toLowerCase()]; });
    if (missing.length) throw new Error('Kolom GL tidak ditemukan: ' + missing.join(', '));
    var headers = {};
    REQUIRED_HEADERS.forEach(function (name) { headers[name] = sourceHeaders[name.toLowerCase()]; });
    var projects = Object.create(null), records = [];
    for (var index = 1; index < rows.length; index++) {
      var source = rows[index];
      if (!Array.isArray(source)) throw new Error('Baris GL tidak valid pada baris ' + (index + 1) + '.');
      if (source.length > 5 && source.slice(5).some(function (value) { return !blank(value); })) {
        throw new Error('File GL harus memiliki tepat lima kolom sumber agar kolom bantu F:I tersedia.');
      }
      var vals = {};
      REQUIRED_HEADERS.forEach(function (name) { vals[name] = source[headers[name] - 1]; });
      if (REQUIRED_HEADERS.every(function (name) { return blank(vals[name]); })) continue;
      var row = index + 1;
      var project = String(vals.Project == null ? '' : vals.Project).trim();
      if (!project) throw new Error('Project kosong pada baris ' + row + '.');
      projects[project] = true;
      try { postingParts(vals['Posting Date']); }
      catch (error) { throw new Error('Posting Date bukan tanggal Excel pada baris ' + row + '.'); }
      var amount = vals['In profit center local currency'];
      if (typeof amount !== 'number' || !Number.isFinite(amount)) {
        throw new Error('Nilai GL bukan angka pada baris ' + row + '.');
      }
      var account = String(vals['GL Account'] == null ? '' : vals['GL Account']).trim();
      if (!account) throw new Error('GL Account kosong pada baris ' + row + '.');
      var parsed = parsePeriod(vals.Text, vals['Posting Date'], account, reportYear, ytdMonth);
      records.push({
        row: row, project: project, account: account,
        postingDate: vals['Posting Date'], amount: amount,
        text: vals.Text == null ? '' : String(vals.Text),
        textMonth: parsed.textMonth, textYear: parsed.textYear,
        period: parsed.period, notes: parsed.notes
      });
    }
    if (!records.length) throw new Error('File GL tidak memiliki baris data.');
    var projectNames = Object.keys(projects);
    if (projectNames.length !== 1) {
      throw new Error('File GL memuat lebih dari satu Project. Unggah satu proyek per laporan.');
    }
    return { project: projectNames[0], headers: headers, records: records };
  }

  function editInteger(value, low, high, label, row) {
    if (blank(value) || (typeof value === 'string' && value.trim() === '')) return null;
    if ((typeof value !== 'number' && typeof value !== 'string') || !Number.isInteger(Number(value))) {
      throw new Error(label + ' tidak valid pada baris ' + row + '.');
    }
    value = Number(value);
    if (value < low || value > high) throw new Error(label + ' di luar rentang pada baris ' + row + '.');
    return value;
  }

  function applyEdits(records, edits, reportYear) {
    validateReport(reportYear, 12);
    if (!Array.isArray(records) || !Array.isArray(edits)) throw new Error('Data koreksi periode tidak valid.');
    var result = records.map(function (record) { return Object.assign({}, record); });
    var byRow = Object.create(null), editedRows = Object.create(null);
    result.forEach(function (record) { byRow[record.row] = record; });
    edits.forEach(function (edit) {
      if (!edit || !Number.isInteger(Number(edit.row)) || !byRow[Number(edit.row)]) {
        throw new Error('Baris GL untuk koreksi tidak ditemukan.');
      }
      var row = Number(edit.row);
      if (editedRows[row]) throw new Error('Koreksi berulang pada baris ' + row + '.');
      editedRows[row] = true;
      var month = editInteger(edit.textMonth, 1, 12, 'Bulan (Text)', row);
      var year = editInteger(edit.textYear, 1900, 9999, 'Tahun (Text)', row);
      var record = byRow[row];
      var period = effectivePeriod(record.postingDate, record.account, reportYear, month, year);
      record.textMonth = month;
      record.textYear = year;
      record.period = period.period;
    });
    return result;
  }

  function zeros() { return new Array(16).fill(0); }

  function calculate(records, labels, reportYear, ytdMonth) {
    if (ytdMonth == null) ytdMonth = 12;
    validateReport(reportYear, ytdMonth);
    if (!Array.isArray(records) || !labels || typeof labels !== 'object') throw new Error('Data laporan tidak valid.');
    var columns = ['Carry Over ' + (reportYear - 1)].concat(MONTH_NAMES, ['TOTAL YTD ' + reportYear, 'TOTAL SAP', 'SELISIH']);
    var rowLabels = {}, accountRows = Object.create(null);
    DISPLAY_ROWS.forEach(function (row) { rowLabels[row] = String(labels[row] == null ? '' : labels[row]).trim(); });
    ITEM_ROWS.forEach(function (row) {
      var label = rowLabels[row];
      if (!accountRows[label]) accountRows[label] = [];
      accountRows[label].push(row);
    });
    var unknown = Object.create(null);
    records.forEach(function (record) {
      var account = String(record.account == null ? '' : record.account).trim();
      if (!accountRows[account]) unknown[account] = true;
    });
    var missing = Object.keys(unknown).sort();
    if (missing.length) throw new Error('GL Account tidak ada di template PL: ' + missing.join(', '));
    var data = records.filter(function (record) {
      var posting = postingParts(record.postingDate);
      return posting.year === reportYear && posting.month <= ytdMonth;
    });
    var values = {};
    ITEM_ROWS.forEach(function (row) { values[row] = zeros(); });
    data.forEach(function (record) {
      accountRows[String(record.account).trim()].forEach(function (row) {
        var amount = (row === 4 ? -1 : 1) * record.amount;
        var vals = values[row];
        if (record.period === 'CO ' + (reportYear - 1)) vals[0] += amount;
        for (var month = 1; month <= ytdMonth; month++) {
          if (record.period === String(month).padStart(2, '0') + '-' + reportYear) vals[month] += amount;
        }
        vals[14] += amount;
      });
    });
    ITEM_ROWS.forEach(function (row) {
      values[row][13] = values[row].slice(1, 13).reduce(function (sum, amount) { return sum + amount; }, 0);
      values[row][15] = values[row][13] - values[row][14];
    });
    [6, 10, 42, 43, 33, 34, 36, 37, 39, 41, 44, 46, 48].forEach(function (row) { values[row] = zeros(); });
    for (var col = 0; col < columns.length; col++) {
      for (var overhead = 11; overhead <= 32; overhead++) values[33][col] += values[overhead][col];
      values[34][col] = values[7][col] + values[8][col] + values[9][col] + values[33][col];
      values[36][col] = values[4][col] - values[34][col];
      values[37][col] = col === 15 ? null : values[4][col] ? values[36][col] / values[4][col] : 0;
      values[39][col] = col === 0 ? 0 : col === 15 ? values[39][14] - values[39][13] : 0.01 * values[4][col];
      values[41][col] = col === 0 ? 0 : col <= 12 ? Math.max(0, values[36][col] - values[39][col]) :
        col < 15 ? values[36][col] - values[39][col] : values[41][13] - values[41][14];
      values[44][col] = values[41][col] + values[42][col] + values[43][col];
      values[46][col] = col === 0 ? 0 : col === 15 ? values[46][14] - values[46][13] : Math.max(0, 0.22 * values[41][col]);
      values[48][col] = col === 15 ? values[48][13] - values[48][14] : values[44][col] - values[46][col];
    }
    var carryRecords = data.filter(function (record) { return record.period === 'CO ' + (reportYear - 1); }).map(function (record) {
      return Object.assign({}, record, { plAmount: (String(record.account).trim() === rowLabels[4] ? -1 : 1) * record.amount });
    });
    var carryByAccount = Object.create(null);
    carryRecords.forEach(function (record) {
      var account = String(record.account).trim();
      if (!carryByAccount[account]) carryByAccount[account] = { account: account, amount: 0, count: 0 };
      carryByAccount[account].amount += record.plAmount;
      carryByAccount[account].count++;
    });
    return {
      columns: columns,
      rows: DISPLAY_ROWS.map(function (row) { return { row: row, label: rowLabels[row], values: values[row] }; }),
      evaluationRate: values[37][13],
      carry: Object.keys(carryByAccount).sort().map(function (account) { return carryByAccount[account]; }),
      reviewRecords: data.filter(function (record) { return Boolean(record.notes); }).map(function (record) { return Object.assign({}, record); }),
      carryRecords: carryRecords
    };
  }

  function titleCase(value) {
    return String(value).toLowerCase().replace(/(^|[^a-z])([a-z])/g, function (_, prefix, letter) {
      return prefix + letter.toUpperCase();
    });
  }

  function roundedThousands(value) {
    var integer = Math.floor(value), remainder = value - integer;
    var rounded = remainder > 0.5 || (remainder === 0.5 && integer % 2 !== 0) ? integer + 1 : integer;
    return String(rounded).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  function defaultJustifications(preview, reportYear) {
    validateReport(reportYear, 12);
    if (!preview || !Array.isArray(preview.rows)) throw new Error('Preview laporan tidak valid.');
    var result = [];
    preview.rows.forEach(function (entry) {
      if (ITEM_ROWS.indexOf(entry.row) === -1 || entry.values[0] === 0) return;
      var amount = entry.values[0];
      var name = entry.row === 4 ? 'Revenue' : titleCase(entry.label);
      result.push({
        row: entry.row,
        text: (result.length + 1) + '. ' + name + ', nilai realisasi ' + (amount < 0 ? 'lebih kecil' : 'lebih besar') +
          ' sebesar ' + roundedThousands(Math.abs(amount)) + ' dikarenakan adanya Carry Over ' + (reportYear - 1)
      });
    });
    return result;
  }

  function numericFormat(original) {
    var sections = String(original || 'General').split(';');
    var positive = sections[0], zero = '_-* "-"??_-', negative;
    if (positive === 'General') return '#,##0;\\-* #,##0_-;' + zero;
    if (positive.indexOf('%') !== -1) {
      negative = positive.indexOf('0.0%') !== -1 ? '\\-* 0.0%_-' : '\\-* 0%_-';
    } else if (positive.indexOf('#,##0') !== -1) {
      negative = '\\-* #,##0_-';
    } else {
      return String(original || 'General');
    }
    sections[1] = negative;
    sections[2] = zero;
    return sections.join(';');
  }

  return {
    ITEM_ROWS: ITEM_ROWS,
    DISPLAY_ROWS: DISPLAY_ROWS,
    MONTH_NAMES: MONTH_NAMES,
    REQUIRED_HEADERS: REQUIRED_HEADERS,
    parsePeriod: parsePeriod,
    effectivePeriod: effectivePeriod,
    parseRows: parseRows,
    applyEdits: applyEdits,
    calculate: calculate,
    defaultJustifications: defaultJustifications,
    numericFormat: numericFormat
  };
})();
