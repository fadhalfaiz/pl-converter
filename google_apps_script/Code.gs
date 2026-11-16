/** Google Apps Script backend. The first web visit initializes its Drive workspace. */
var PL_APP = Object.freeze({
  xlsxMime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  sheetsMime: 'application/vnd.google-apps.spreadsheet',
  folderKey: 'PL_FOLDER_ID',
  templateKey: 'PL_TEMPLATE_ID',
  jobPrefix: 'PL_JOB_',
  maxUploadBytes: 5 * 1024 * 1024,
  maxRows: 5000,
  sessionMs: 4 * 60 * 60 * 1000,
  signatures: {
    creatorName: 'Fadhlillah Rifqi H', creatorRole: 'UH Performance BBM',
    reviewerName: 'Rizky Farlyan', reviewerRole: 'Manager TOPM Dept.',
    approverName: 'Rahman Ilahi', approverRole: 'GM Transportation Operation'
  }
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PL REALISASI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Idempotent installer, also usable directly from the Apps Script editor. */
function setupApp_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var props = PropertiesService.getScriptProperties();
    var folderId = props.getProperty(PL_APP.folderKey);
    if (!folderId || !managedFileExists_(folderId)) {
      folderId = Drive.Files.create({
        name: 'PL REALISASI - Data Aplikasi', mimeType: 'application/vnd.google-apps.folder',
        appProperties: managedTags_('folder')
      }, null, {fields: 'id'}).id;
      props.setProperty(PL_APP.folderKey, folderId);
    }
    var templateId = props.getProperty(PL_APP.templateKey);
    if (!templateId || !managedFileExists_(templateId)) {
      var blob = Utilities.newBlob(Utilities.base64Decode(PL_TEMPLATE.base64),
        PL_APP.xlsxMime, 'Tamplate_PL.xlsx');
      var file = createManagedFile_('Template PL REALISASI', PL_APP.sheetsMime, blob, folderId, 'template');
      templateId = file.id;
      try {
        var book = openSpreadsheet_(templateId);
        var pl = book.getSheetByName('PL');
        if (!pl) throw new Error('Template bawaan tidak memiliki sheet PL.');
        book.getSheets().forEach(function(sheet) {
          if (sheet.getSheetId() !== pl.getSheetId()) book.deleteSheet(sheet);
        });
        book.setSpreadsheetTimeZone('Asia/Jakarta');
        book.setSpreadsheetLocale('en_US');
        templateLabels_(book);
        props.setProperty(PL_APP.templateKey, templateId);
      } catch (error) {
        trashManagedFile_(templateId);
        throw error;
      }
    }
    if (!ScriptApp.getProjectTriggers().some(function(trigger) {
      return trigger.getHandlerFunction() === 'cleanupExpiredReports_';
    })) ScriptApp.newTrigger('cleanupExpiredReports_').timeBased().everyHours(1).create();
    console.log('Setup selesai. Folder kerja: https://drive.google.com/drive/folders/' + folderId);
    return {ok: true, message: 'Setup selesai. Aplikasi siap dideploy.'};
  } finally { lock.releaseLock(); }
}

function getBootstrap() {
  configuration_();
  return {
    ready: true,
    message: '',
    defaultYear: 2026, defaultYtdMonth: 12, maxUploadBytes: PL_APP.maxUploadBytes,
    signatures: PL_APP.signatures
  };
}

function uploadGl(request) {
  requireObject_(request);
  var config = configuration_();
  var clientKey = validateClientKey_(request.clientKey);
  var year = validateInteger_(request.reportYear, 1900, 9999, 'Tahun Laporan');
  var month = validateInteger_(request.ytdMonth, 1, 12, 'Bulan YTD');
  var fileName = String(request.fileName || '');
  if (!/\.xlsx$/i.test(fileName)) throw new Error('Unggah file GL dengan format .xlsx.');
  var base64 = String(request.base64 || '');
  if (!base64 || base64.length > Math.ceil(PL_APP.maxUploadBytes / 3) * 4 + 4 ||
      !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)) throw new Error('File tidak valid atau melebihi batas 5 MB.');
  var bytes = Utilities.base64Decode(base64);
  if (bytes.length > PL_APP.maxUploadBytes || bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) {
    throw new Error('File harus berupa workbook Excel .xlsx yang valid, maksimal 5 MB.');
  }
  var sourceId = null;
  var stateId = null;
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  try {
    var blob = Utilities.newBlob(bytes, PL_APP.xlsxMime, fileName);
    sourceId = createManagedFile_('GL - ' + fileName.slice(0, 160), PL_APP.sheetsMime,
      blob, config.folderId, 'source').id;
    var book = openSpreadsheet_(sourceId);
    var sheet = book.getSheets()[0];
    if (sheet.getName() === 'PL') throw new Error('Nama sheet GL tidak boleh PL karena dipakai untuk laporan.');
    if (sheet.getLastRow() > PL_APP.maxRows + 1) throw new Error('Maksimal 5.000 baris GL per laporan.');
    if (sheet.getLastColumn() !== 5) throw new Error('File GL harus memiliki tepat lima kolom sumber.');
    var values = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), 5).getValues();
    var tz = book.getSpreadsheetTimeZone();
    values.forEach(function(row) {
      row.forEach(function(value, index) {
        if (value instanceof Date) row[index] = Utilities.formatDate(value, tz, 'yyyy-MM-dd');
      });
    });
    var parsed = PLCore.parseRows(values, year, month);
    var template = openSpreadsheet_(config.templateId);
    var labels = templateLabels_(template);
    PLCore.calculate(parsed.records, labels, year, month);
    var expires = Date.now() + PL_APP.sessionMs;
    var state = {
      version: 1, token: token, clientHash: sha256_(clientKey), sourceId: sourceId,
      sourceSheetId: sheet.getSheetId(), sheetName: sheet.getName(), sourceLastRow: values.length,
      project: parsed.project, headers: parsed.headers, records: parsed.records,
      labels: labels, reportYear: year, ytdMonth: month, expiresAt: expires
    };
    stateId = createManagedFile_('Sesi PL ' + token.slice(0, 8), 'application/json',
      Utilities.newBlob(JSON.stringify(state), 'application/json'), config.folderId, 'state').id;
    PropertiesService.getScriptProperties().setProperty(PL_APP.jobPrefix + token,
      JSON.stringify({stateId: stateId, sourceId: sourceId, expiresAt: expires}));
    return reportView_(state);
  } catch (error) {
    if (stateId) trashManagedFile_(stateId);
    if (sourceId) trashManagedFile_(sourceId);
    throw new Error('GL gagal diproses: ' + error.message);
  }
}

function updateReport(request) {
  var job = loadJob_(request);
  job.state.records = PLCore.applyEdits(job.state.records, request.edits || [], job.state.reportYear);
  saveJob_(job);
  return reportView_(job.state);
}

function exportReport(request) {
  var job = loadJob_(request);
  var state = job.state;
  state.records = PLCore.applyEdits(state.records, request.edits || [], state.reportYear);
  var preview = PLCore.calculate(state.records, state.labels, state.reportYear, state.ytdMonth);
  var justifications = validatedJustifications_(request.justifications, preview, state.reportYear);
  var signatures = validatedSignatures_(request.signatures);
  saveJob_(job);
  var config = configuration_();
  var outputId = null;
  try {
    outputId = Drive.Files.copy({name: 'PL - ' + state.project, parents: [config.folderId],
      appProperties: managedTags_('export')}, config.templateId, {fields: 'id'}).id;
    var output = openSpreadsheet_(outputId);
    output.setSpreadsheetLocale('en_US');
    output.setSpreadsheetTimeZone('Asia/Jakarta');
    var pl = output.getSheetByName('PL');
    output.getSheets().forEach(function(sheet) {
      if (sheet.getSheetId() !== pl.getSheetId()) output.deleteSheet(sheet);
    });
    var original = openSpreadsheet_(state.sourceId).getSheetById(state.sourceSheetId);
    if (!original) throw new Error('Sheet sumber sesi tidak tersedia. Unggah ulang GL.');
    var data = original.copyTo(output).setName(state.sheetName);
    output.setActiveSheet(data);
    output.moveActiveSheet(1);
    prepareGlSheet_(data, state);
    preparePlSheet_(pl, state, data.getName());
    var printBottom = writeNarrative_(pl, justifications, signatures);
    SpreadsheetApp.flush();
    validateComputedReport_(pl, preview);
    var name = safeFileName_(state.project, state.reportYear, state.ytdMonth);
    var exported = exportSpreadsheet_(outputId);
    exported = patchPrintLayout_(exported, printBottom, name);
    return {fileName: name, mimeType: PL_APP.xlsxMime,
      base64: Utilities.base64Encode(exported.getBytes())};
  } finally { if (outputId) trashManagedFile_(outputId); }
}

function discardReport(request) {
  var job = loadJob_(request);
  removeJob_(job.key, job.meta);
  return {ok: true};
}

/** Editor smoke test: exercises actual Google import, preview, edits, and XLSX export. */
function runSmokeTest_() {
  setupApp_();
  var config = configuration_();
  var sourceId;
  var token;
  var clientKey = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  try {
    sourceId = createManagedFile_('UJI - GL Apps Script', PL_APP.sheetsMime, null, config.folderId, 'export').id;
    var book = openSpreadsheet_(sourceId);
    book.setSpreadsheetTimeZone('Asia/Jakarta');
    var sheet = book.getSheets()[0].setName('SAP');
    var values = [PLCore.REQUIRED_HEADERS,
      ['UJI', 'OPERATING REVENUE', new Date('2026-01-15T00:00:00Z'), -1000, 'Jan 26'],
      ['UJI', 'RENT', new Date('2026-01-15T00:00:00Z'), 200, 'Des 25'],
      ['UJI', 'RENT', new Date('2026-02-15T00:00:00Z'), 100, 'Des 24'],
      ['UJI', 'DIRECT LABOR', new Date('2026-01-15T00:00:00Z'), 100, 'Des 24'],
      ['UJI', 'RENT', new Date('2026-02-15T00:00:00Z'), 50, 'Jan-Mar 26'],
      ['UJI', 'OPERATING REVENUE', new Date('2026-03-15T00:00:00Z'), -999, 'Jan 26'],
      ['UJI', 'SALARY AND COMPENSATION', new Date('2026-01-15T00:00:00Z'), 10, 'Nov 24']];
    sheet.getRange(1, 1, values.length, 5).setValues(values);
    sheet.getRange(2, 3, values.length - 1, 1).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    var input = exportSpreadsheet_(sourceId);
    var view = uploadGl({fileName: 'Uji_GL.xlsx', base64: Utilities.base64Encode(input.getBytes()),
      reportYear: 2026, ytdMonth: 2, clientKey: clientKey});
    token = view.token;
    var auth = {token: token, clientKey: clientKey};
    var state = loadJob_(auth).state;
    if (state.records.find(function(row) { return row.row === 4; }).period !== 'CO 2024' ||
        state.records.find(function(row) { return row.row === 5; }).period !== '01-2026' ||
        Math.abs(view.evaluationRate - 0.84) > 1e-10) throw new Error('Pengujian aturan periode/evaluasi gagal.');
    var changed = updateReport({token: token, clientKey: clientKey,
      edits: [{row: 6, textMonth: 1, textYear: 2026}]});
    if (loadJob_(auth).state.records.find(function(row) { return row.row === 6; }).period !== '01-2026') {
      throw new Error('Pengujian koreksi periode gagal.');
    }
    var lines = changed.justifications.map(function(item) { return {row: item.row, text: 'Uji justifikasi - ' + item.text}; });
    var download = exportReport({token: token, clientKey: clientKey, edits: [], justifications: lines,
      signatures: {creatorName: 'Pembuat Uji', creatorRole: 'Jabatan Pembuat Uji',
        reviewerName: 'Mengetahui Uji', reviewerRole: 'Manager TOPM Dept.',
        approverName: 'Menyetujui Uji', approverRole: 'Jabatan Menyetujui Uji'}});
    var bytes = Utilities.base64Decode(download.base64);
    var blob = Utilities.newBlob(bytes, PL_APP.xlsxMime, 'UJI_' + download.fileName);
    var entries = Utilities.unzip(blob.copyBlob().setContentType('application/zip'));
    var workbook = entries.find(function(entry) { return entry.getName() === 'xl/workbook.xml'; });
    if (!workbook || workbook.getDataAsString('UTF-8').indexOf('_xlnm.Print_Area') < 0) {
      throw new Error('Pengujian area cetak Excel gagal.');
    }
    var output = createManagedFile_('UJI_' + download.fileName, PL_APP.xlsxMime, blob, config.folderId, 'smoke-result');
    var url = 'https://drive.google.com/file/d/' + output.id + '/view';
    console.log('Semua pengujian Google selesai. Periksa Excel: ' + url);
    return {ok: true, message: 'Impor, koreksi, perhitungan, dan ekspor berhasil.', fileUrl: url};
  } finally {
    if (token) {
      try { discardReport({token: token, clientKey: clientKey}); }
      catch (error) { console.warn(error.message); }
    }
    if (sourceId) trashManagedFile_(sourceId);
  }
}

function configuration_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(PL_APP.folderKey);
  var templateId = props.getProperty(PL_APP.templateKey);
  if (!folderId || !templateId) {
    setupApp_();
    folderId = props.getProperty(PL_APP.folderKey);
    templateId = props.getProperty(PL_APP.templateKey);
  }
  if (!folderId || !templateId) throw new Error('Folder kerja atau template belum berhasil disiapkan. Muat ulang halaman untuk mencoba lagi.');
  return {folderId: folderId, templateId: templateId};
}

function reportView_(state) {
  var result = PLCore.calculate(state.records, state.labels, state.reportYear, state.ytdMonth);
  return {
    token: state.token, project: state.project, sheetName: state.sheetName,
    reportYear: state.reportYear, ytdMonth: state.ytdMonth, recordCount: state.records.length,
    columns: result.columns, rows: result.rows, evaluationRate: result.evaluationRate,
    justifications: PLCore.defaultJustifications(result, state.reportYear),
    reviewRecords: result.reviewRecords, carry: result.carry, carryRecords: result.carryRecords,
    expiresAt: new Date(state.expiresAt).toISOString()
  };
}

function loadJob_(request) {
  requireObject_(request);
  var token = String(request.token || '');
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('Sesi laporan tidak valid. Unggah ulang GL.');
  var clientKey = validateClientKey_(request.clientKey);
  var key = PL_APP.jobPrefix + token;
  var raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) throw new Error('Sesi laporan tidak ditemukan. Unggah ulang GL.');
  var meta = JSON.parse(raw);
  if (Date.now() >= meta.expiresAt) {
    removeJob_(key, meta);
    throw new Error('Sesi laporan telah berakhir setelah 4 jam. Unggah ulang GL.');
  }
  var state = JSON.parse(DriveApp.getFileById(meta.stateId).getBlob().getDataAsString('UTF-8'));
  if (state.token !== token || state.clientHash !== sha256_(clientKey)) {
    throw new Error('Sesi laporan tidak cocok dengan browser ini.');
  }
  return {key: key, meta: meta, state: state};
}

function saveJob_(job) {
  DriveApp.getFileById(job.meta.stateId).setContent(JSON.stringify(job.state));
}

function removeJob_(key, meta) {
  var success = trashManagedFile_(meta.sourceId) && trashManagedFile_(meta.stateId);
  if (success) PropertiesService.getScriptProperties().deleteProperty(key);
}

/** Hourly trigger only removes expired files created by this project. */
function cleanupExpiredReports_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = Date.now();
  var activeFiles = {};
  Object.keys(all).forEach(function(key) {
    if (key.indexOf(PL_APP.jobPrefix) !== 0) return;
    try {
      var meta = JSON.parse(all[key]);
      if (now >= meta.expiresAt) removeJob_(key, meta);
      else { activeFiles[meta.sourceId] = true; activeFiles[meta.stateId] = true; }
    } catch (error) { console.warn('Pembersihan sesi gagal: ' + error.message); }
  });
  var config = configuration_();
  var query = "'" + config.folderId + "' in parents and trashed=false and " +
    "(appProperties has { key='plRole' and value='export' } or " +
    "appProperties has { key='plRole' and value='source' } or " +
    "appProperties has { key='plRole' and value='state' })";
  var pageToken;
  do {
    var page = Drive.Files.list({q: query, fields: 'nextPageToken,files(id,createdTime,appProperties)',
      pageSize: 100, pageToken: pageToken});
    (page.files || []).forEach(function(file) {
      if (!activeFiles[file.id] && now - Date.parse(file.createdTime) > PL_APP.sessionMs) trashManagedFile_(file.id);
    });
    pageToken = page.nextPageToken;
  } while (pageToken);
}

function managedTags_(role) {
  return {plConverter: ScriptApp.getScriptId(), plRole: role};
}

function createManagedFile_(name, mimeType, blob, folderId, role) {
  return Drive.Files.create({name: name, mimeType: mimeType, parents: [folderId],
    appProperties: managedTags_(role)}, blob, {fields: 'id'});
}

function managedFileExists_(id) {
  try {
    var file = Drive.Files.get(id, {fields: 'trashed,appProperties'});
    return !file.trashed && file.appProperties && file.appProperties.plConverter === ScriptApp.getScriptId();
  } catch (error) { return false; }
}

function trashManagedFile_(id) {
  if (!id) return true;
  try {
    var file = Drive.Files.get(id, {fields: 'trashed,appProperties'});
    if (file.trashed) return true;
    if (!file.appProperties || file.appProperties.plConverter !== ScriptApp.getScriptId()) return false;
    Drive.Files.update({trashed: true}, id);
    return true;
  } catch (error) { console.warn('File sementara belum dibersihkan: ' + error.message); return false; }
}

function openSpreadsheet_(id) {
  for (var attempt = 0; attempt < 3; attempt++) {
    try { return SpreadsheetApp.openById(id); }
    catch (error) { if (attempt === 2) throw error; Utilities.sleep(600 * (attempt + 1)); }
  }
}

function templateLabels_(book) {
  var pl = book.getSheetByName('PL');
  if (!pl) throw new Error('Template harus memiliki sheet PL.');
  var names = pl.getRange('B1:B48').getDisplayValues();
  var labels = {};
  PLCore.DISPLAY_ROWS.forEach(function(row) { labels[row] = String(names[row - 1][0] || '').trim(); });
  if (labels[4] !== 'OPERATING REVENUE' || labels[37] !== 'GPM') {
    throw new Error('Posisi ITEMS pada template PL tidak sesuai template bawaan.');
  }
  return labels;
}

function prepareGlSheet_(data, state) {
  if (data.getMaxColumns() < 10) data.insertColumnsAfter(data.getMaxColumns(), 10 - data.getMaxColumns());
  var last = state.sourceLastRow;
  var helpers = Array.from({length: last - 1}, function() { return ['', '', '', '']; });
  var normalizedAccounts = Array.from({length: last - 1}, function() { return ['']; });
  var accountCol = columnLetter_(state.headers['GL Account']);
  var postingCol = columnLetter_(state.headers['Posting Date']);
  state.records.forEach(function(record) {
    var row = record.row;
    var a = 'UPPER(TRIM($' + accountCol + row + '))';
    var p = '$' + postingCol + row;
    normalizedAccounts[row - 2] = ['=TRIM($' + accountCol + row + ')'];
    helpers[row - 2] = [record.textMonth === null ? '' : record.textMonth,
      record.textYear === null ? '' : record.textYear,
      '=IF(OR(' + a + '="DIRECT LABOR",' + a + '="SALARY AND COMPENSATION",' +
      a + '="DEPRECIATION & AMORTIZATION"),DATE(YEAR(' + p + '),MONTH(' + p + '),1),' +
      'IF(AND($G' + row + '<>"",$G' + row + '<PL!$C$2),"CO "&$G' + row + ',' +
      'IF(AND($G' + row + '<>"",$F' + row + '<>""),DATE($G' + row + ',$F' + row + ',1),' +
      'DATE(YEAR(' + p + '),MONTH(' + p + '),1))))', record.notes || ''];
  });
  data.getRange(1, 6, 1, 4).setValues([['Bulan (Text)', 'Tahun (Text)', 'Periode PL', 'Catatan']]);
  data.getRange(2, 6, last - 1, 4).setValues(helpers);
  data.getRange('J1').setValue('GL Account (normalisasi)');
  data.getRange(2, 10, last - 1, 1).setValues(normalizedAccounts);
  data.hideColumns(10);
  data.getRange(1, 1, 1, 9).setBackground('#FFF2CC').setFontFamily('Calibri Light').setFontSize(12).setFontWeight('bold');
  data.getRange(2, 6, last - 1, 2).setFontColor('#0000FF').setNumberFormat('0');
  data.getRange(2, 8, last - 1, 1).setNumberFormat('mmm-yy');
  var amount = data.getRange(2, state.headers['In profit center local currency'], last - 1, 1);
  amount.setNumberFormats(amount.getNumberFormats().map(function(row) { return [PLCore.numericFormat(row[0])]; }));
  if (data.getFilter()) data.getFilter().remove();
  data.getRange(1, 1, last, 9).createFilter();
  data.setFrozenRows(1);
  data.setColumnWidths(6, 3, 120);
  data.setColumnWidth(9, 560);
}

function preparePlSheet_(pl, state, sheetName) {
  if (pl.getMaxColumns() < 27) pl.insertColumnsAfter(pl.getMaxColumns(), 27 - pl.getMaxColumns());
  pl.getRange('B1').setValue('PL REALISASI ' + state.project.toUpperCase());
  pl.getRange('B2').setValue('TAHUN');
  pl.getRange('C2').setValue(state.reportYear).setFontColor('#0000FF').setBackground('#FFF2CC');
  pl.getRange('M2:N2').clearContent();
  pl.getRange('AA1:AA2').setValues([['BULAN YTD'], [state.ytdMonth]]);
  pl.hideColumns(27);
  pl.getRange('C3').setFormula('="Carry over "&($C$2-1)');
  pl.getRange('D3:O3').setValues([PLCore.MONTH_NAMES]);
  pl.getRange('P3:R3').setValues([['TOTAL YTD ' + state.reportYear, 'TOTAL SAP', 'SELISIH']]);
  var formulas = reportFormulaMatrix_(state.headers, sheetName);
  var target = pl.getRange('C4:R48');
  target.setValues(formulas);
  target.setNumberFormats(target.getNumberFormats().map(function(row) { return row.map(PLCore.numericFormat); }));
}

/** Matrix matches PLCore/Python preview, including the monthly operating income floor. */
function reportFormulaMatrix_(headers, sheetName) {
  var result = Array.from({length: 45}, function() { return Array(16).fill(''); });
  function put(row, col, formula) { result[row - 4][col - 3] = formula; }
  var quoted = "'" + sheetName.replace(/'/g, "''") + "'";
  var amount = quoted + '!$' + columnLetter_(headers['In profit center local currency']) + ':$' + columnLetter_(headers['In profit center local currency']);
  var account = quoted + '!$J:$J';
  var posting = quoted + '!$' + columnLetter_(headers['Posting Date']) + ':$' + columnLetter_(headers['Posting Date']);
  var period = quoted + '!$H:$H';
  var window = posting + ',">="&DATE($C$2,1,1),' + posting + ',"<"&DATE($C$2,$AA$2+1,1)';
  PLCore.ITEM_ROWS.forEach(function(row) {
    var sign = row === 4 ? '-' : '';
    put(row, 3, '=' + sign + 'SUMIFS(' + amount + ',' + account + ',$B' + row + ',' + period + ',"CO "&($C$2-1),' + window + ')');
    for (var month = 1; month <= 12; month++) {
      put(row, month + 3, '=IF(' + month + '<=$AA$2,' + sign + 'SUMIFS(' + amount + ',' + account + ',$B' + row + ',' +
        period + ',">="&DATE($C$2,' + month + ',1),' + period + ',"<"&DATE($C$2,' + (month + 1) + ',1),' + window + '),0)');
    }
    put(row, 16, '=SUM(D' + row + ':O' + row + ')');
    put(row, 17, '=' + sign + 'SUMIFS(' + amount + ',' + account + ',$B' + row + ',' + window + ')');
    put(row, 18, '=P' + row + '-Q' + row);
  });
  for (var col = 3; col <= 18; col++) {
    var c = columnLetter_(col);
    put(33, col, '=SUM(' + c + '11:' + c + '32)');
    put(34, col, '=SUM(' + c + '7:' + c + '9,' + c + '11:' + c + '32)');
    put(36, col, col === 18 ? '=P36-Q36' : '=' + c + '4-' + c + '34');
    if (col < 18) put(37, col, '=IFERROR(' + c + '36/' + c + '4,0)');
    put(39, col, col === 3 ? '=0' : col === 18 ? '=Q39-P39' : '=1%*' + c + '4');
    put(41, col, col === 3 ? '=0' : col === 18 ? '=P41-Q41' : col <= 15 ? '=MAX(0,' + c + '36-' + c + '39)' : '=' + c + '36-' + c + '39');
    [42, 43].forEach(function(row) { put(row, col, col === 18 ? '=P' + row + '-Q' + row : '=0'); });
    put(44, col, col === 18 ? '=P44-Q44' : '=SUM(' + c + '41:' + c + '43)');
    put(46, col, col === 3 ? '=0' : col === 18 ? '=Q46-P46' : '=MAX(0,22%*' + c + '41)');
    put(48, col, col === 18 ? '=P48-Q48' : '=' + c + '44-' + c + '46');
  }
  return result;
}

function writeNarrative_(pl, lines, signatures) {
  // Imported Excel freeze panes can split B:R. Sheets rejects evaluation
  // merges across that boundary, so unfreeze columns on the export copy.
  pl.setFrozenColumns(0);
  pl.getRange('B54:B60').clearContent();
  var extra = Math.max(0, lines.length + 2 - 7);
  if (extra) pl.insertRowsBefore(61, extra);
  if (lines.length) {
    var range = pl.getRange(54, 2, lines.length, 17);
    range.setFontFamily('Calibri Light').setFontSize(12).setVerticalAlignment('top').setWrap(false);
    lines.forEach(function(line, index) { setLiteral_(pl.getRange(54 + index, 2), line); });
  }
  var evalRow = 54 + lines.length;
  pl.getRange(evalRow, 2, 1, 17).merge().setFontFamily('Calibri Light').setFontSize(12).setFontWeight('bold').setVerticalAlignment('top');
  setLiteral_(pl.getRange(evalRow, 2), 'Hasil Evaluasi :');
  pl.getRange(evalRow + 1, 2, 1, 17).merge().setFontFamily('Calibri Light').setFontSize(12).setFontWeight('normal').setVerticalAlignment('top');
  pl.getRange(evalRow + 1, 2).setFormula('="Nilai NPAT setelah dilakukan cleasing sebesar "&SUBSTITUTE(TEXT(P37,"0.0%"),".",",")');
  pl.setRowHeights(evalRow, 2, 22);
  var headerRow = 61 + extra;
  [[3, 'Menyetujui,', signatures.approverName, signatures.approverRole],
   [6, 'Mengetahui,', signatures.reviewerName, signatures.reviewerRole],
   [9, 'Dibuat oleh,', signatures.creatorName, signatures.creatorRole]].forEach(function(item) {
    if (pl.getRange(headerRow, item[0]).getDisplayValue() !== item[1]) {
      throw new Error('Blok tanda tangan tidak sesuai dengan template bawaan.');
    }
    setLiteral_(pl.getRange(headerRow + 5, item[0]), item[2]);
    setLiteral_(pl.getRange(headerRow + 6, item[0]), item[3]);
  });
  return 68 + extra;
}

function validateComputedReport_(pl, preview) {
  var errors = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|ERROR!)/;
  var values = pl.getRange('C4:R48').getValues();
  preview.rows.forEach(function(row) {
    if (row.row === 6 || row.row === 10) return;
    row.values.forEach(function(expected, index) {
      if (expected === null) return;
      var actual = values[row.row - 4][index];
      if (typeof actual === 'string' && errors.test(actual)) throw new Error('Rumus gagal pada ' + columnLetter_(index + 3) + row.row + ': ' + actual);
      var tolerance = row.row === 37 ? 1e-10 : 0.01;
      if (typeof actual !== 'number' || Math.abs(actual - expected) > tolerance) {
        throw new Error('Hasil rumus berbeda dari preview pada ' + columnLetter_(index + 3) + row.row + '. Coba proses ulang.');
      }
    });
  });
}

function validatedJustifications_(input, preview, year) {
  var defaults = PLCore.defaultJustifications(preview, year);
  if (input === undefined || input === null) return defaults.map(function(item) { return item.text; });
  if (!Array.isArray(input) || input.length > 30) throw new Error('Isian justifikasi tidak valid.');
  var map = {};
  input.forEach(function(item) {
    requireObject_(item);
    var row = validateInteger_(item.row, 4, 32, 'Baris justifikasi');
    if (!defaults.some(function(value) { return value.row === row; }) || Object.prototype.hasOwnProperty.call(map, row)) {
      throw new Error('Baris justifikasi tidak sesuai dengan laporan.');
    }
    map[row] = validateText_(item.text, 5000, 'Justifikasi');
  });
  return defaults.map(function(item) { return Object.prototype.hasOwnProperty.call(map, item.row) ? map[item.row] : item.text; });
}

function validatedSignatures_(input) {
  input = input || {};
  requireObject_(input);
  var result = {};
  Object.keys(PL_APP.signatures).forEach(function(key) {
    result[key] = input[key] === undefined ? PL_APP.signatures[key] : validateText_(input[key], 200, 'Nama/jabatan');
  });
  return result;
}

function setLiteral_(range, text) {
  range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(String(text))
    .setTextStyle(range.getTextStyle()).build());
}

function exportSpreadsheet_(id) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '/export?mimeType=' + encodeURIComponent(PL_APP.xlsxMime);
  for (var attempt = 0; attempt < 3; attempt++) {
    var response = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
    if (response.getResponseCode() === 200) return response.getBlob().setContentType(PL_APP.xlsxMime);
    if ([429, 500, 502, 503, 504].indexOf(response.getResponseCode()) >= 0 && attempt < 2) { Utilities.sleep(600 * (attempt + 1)); continue; }
    throw new Error('Ekspor Excel gagal (HTTP ' + response.getResponseCode() + '). Ukuran ekspor Google Drive dibatasi 10 MB.');
  }
}

/** Restore the template print area/page settings after the Google Sheets round trip. */
function patchPrintLayout_(blob, bottom, fileName) {
  var entries = Utilities.unzip(blob.setContentType('application/zip'));
  var byName = {};
  entries.forEach(function(entry) { byName[entry.getName()] = entry; });
  if (!byName['xl/workbook.xml'] || !byName['xl/_rels/workbook.xml.rels']) throw new Error('Struktur Excel hasil ekspor tidak valid.');
  var workbook = byName['xl/workbook.xml'].getDataAsString('UTF-8');
  var sheets = workbook.match(/<(?:\w+:)?sheet\b[^>]*>/g) || [];
  var index = sheets.findIndex(function(tag) { return xmlAttribute_(tag, 'name') === 'PL'; });
  if (index < 0) throw new Error('Sheet PL tidak ditemukan dalam hasil ekspor.');
  var relation = xmlAttribute_(sheets[index], 'r:id');
  var rels = byName['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8').match(/<(?:\w+:)?Relationship\b[^>]*>/g) || [];
  var target = '';
  rels.forEach(function(tag) { if (xmlAttribute_(tag, 'Id') === relation) target = xmlAttribute_(tag, 'Target'); });
  var path = target.indexOf('/') === 0 ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
  if (!byName[path]) throw new Error('Relasi sheet PL pada hasil ekspor tidak valid.');
  byName['xl/workbook.xml'].setDataFromString(patchWorkbookPrintArea_(workbook, index, bottom), 'UTF-8');
  byName[path].setDataFromString(restoreSheetPrintSettings_(byName[path].getDataAsString('UTF-8')), 'UTF-8');
  if (byName['xl/styles.xml']) {
    byName['xl/styles.xml'].setDataFromString(patchNumericFormats_(byName['xl/styles.xml'].getDataAsString('UTF-8')), 'UTF-8');
  }
  return Utilities.zip(entries, fileName).setContentType(PL_APP.xlsxMime);
}

function patchNumericFormats_(xml) {
  return xml.replace(/<numFmt\b[^>]*>/g, function(tag) {
    var encoded = xmlAttribute_(tag, 'formatCode');
    var format = encoded.replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    var updated = PLCore.numericFormat(format);
    var escaped = updated.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
      .replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return tag.replace(/formatCode="[^"]*"/, function() { return 'formatCode="' + escaped + '"'; });
  });
}

function xmlAttribute_(tag, name) {
  var escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var match = tag.match(new RegExp('(?:^|\\s)' + escaped + '="([^"]*)"'));
  return match ? match[1] : '';
}

function patchWorkbookPrintArea_(xml, sheetIndex, bottom) {
  var prefix = ((xml.match(/<(\w+:)?workbook\b/) || [])[1]) || '';
  var block = new RegExp('<' + prefix + 'definedNames\\b[^>]*>([\\s\\S]*?)</' + prefix + 'definedNames>');
  var current = (xml.match(block) || [])[1] || '';
  current = current.replace(new RegExp('<' + prefix + 'definedName\\b[^>]*>[\\s\\S]*?</' + prefix + 'definedName>', 'g'), function(tag) {
    return xmlAttribute_(tag, 'name') === '_xlnm.Print_Area' && xmlAttribute_(tag, 'localSheetId') === String(sheetIndex) ? '' : tag;
  });
  var name = '<' + prefix + 'definedName name="_xlnm.Print_Area" localSheetId="' + sheetIndex + '">' +
    "'PL'!$A$1:$S$" + bottom + '</' + prefix + 'definedName>';
  var replacement = '<' + prefix + 'definedNames>' + current + name + '</' + prefix + 'definedNames>';
  if (block.test(xml)) return xml.replace(block, function() { return replacement; });
  var empty = new RegExp('<' + prefix + 'definedNames\\b[^>]*/>');
  if (empty.test(xml)) return xml.replace(empty, function() { return replacement; });
  var before = new RegExp('<' + prefix + '(?:calcPr|oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\\b');
  var match = before.exec(xml);
  var position = match ? match.index : xml.lastIndexOf('</' + prefix + 'workbook>');
  return xml.slice(0, position) + replacement + xml.slice(position);
}

function restoreSheetPrintSettings_(xml) {
  var settings = PL_TEMPLATE.printSettings;
  var successors = 'headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst';
  [['printOptions', settings.printOptions, 'pageMargins|pageSetup|' + successors],
   ['pageMargins', settings.pageMargins, 'pageSetup|' + successors],
   ['pageSetup', settings.pageSetup, successors]].forEach(function(item) {
    if (item[1]) xml = replaceSheetElement_(xml, item[0], item[1], item[2]);
  });
  if (settings.pageSetUpPr) {
    var section = /<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/;
    if (section.test(xml)) xml = xml.replace(section, function(tag) {
      return replaceSheetElement_(tag, 'pageSetUpPr', settings.pageSetUpPr, '');
    });
    else if (/<sheetPr\b[^>]*\/>/.test(xml)) xml = xml.replace(/<sheetPr\b[^>]*\/>/, '<sheetPr>' + settings.pageSetUpPr + '</sheetPr>');
    else xml = xml.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr>' + settings.pageSetUpPr + '</sheetPr>');
  }
  return xml;
}

function replaceSheetElement_(xml, name, fragment, successors) {
  var element = new RegExp('<' + name + '\\b[^>]*(?:/>|>[\\s\\S]*?</' + name + '>)');
  if (element.test(xml)) return xml.replace(element, function() { return fragment; });
  var next = successors ? new RegExp('<(?:' + successors + ')\\b').exec(xml) : null;
  var position = next ? next.index : xml.lastIndexOf('</');
  return xml.slice(0, position) + fragment + xml.slice(position);
}

function requireObject_(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Permintaan aplikasi tidak valid.');
}

function validateInteger_(value, low, high, label) {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < low || value > high) throw new Error(label + ' tidak valid.');
  return value;
}

function validateClientKey_(value) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{32,128}$/.test(value)) throw new Error('Identitas browser tidak valid. Muat ulang halaman.');
  return value;
}

function validateText_(value, max, label) {
  if (typeof value !== 'string' || value.length > max) throw new Error(label + ' tidak valid atau terlalu panjang.');
  return value.trim();
}

function sha256_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function(byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function columnLetter_(column) {
  var name = '';
  while (column > 0) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); }
  return name;
}

function safeFileName_(project, year, month) {
  return 'PL_' + project.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') + '_' + year + '_YTD_' + String(month).padStart(2, '0') + '.xlsx';
}
