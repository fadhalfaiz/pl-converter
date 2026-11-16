'use strict';

/* Tests pure Apps Script backend functions. No Google services are mocked or
 * invoked, so these tests do not claim to validate live Drive/Sheets behavior.
 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {spawnSync} = require('node:child_process');

const packageRoot = path.resolve(__dirname, '..');
const projectRoot = path.resolve(packageRoot, '..');
const index = process.argv.indexOf('--python');
const python = index >= 0 ? process.argv[index + 1] : (process.env.PL_TEST_PYTHON || 'python');
function pythonJson(script, input) {
  const output = spawnSync(python, [path.join(__dirname, script)], {
    cwd: projectRoot, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024,
    input: input === undefined ? undefined : JSON.stringify(input),
  });
  if (output.error) throw output.error;
  assert.equal(output.status, 0, output.stderr);
  return JSON.parse(output.stdout);
}
const fixture = pythonJson('export_parity_fixture.py');
const ctx = vm.createContext({console});
for (const name of ['Core.gs', 'Code.gs']) {
  vm.runInContext(fs.readFileSync(path.join(packageRoot, name), 'utf8'), ctx, {filename: name});
}
const templateFile = fs.readdirSync(packageRoot).find(name => name.endsWith('.gs') &&
  /var\s+PL_TEMPLATE\s*=/.test(fs.readFileSync(path.join(packageRoot, name), 'utf8')));
assert(templateFile, 'Generate the embedded PL_TEMPLATE package before running backend tests.');
vm.runInContext(fs.readFileSync(path.join(packageRoot, templateFile), 'utf8'), ctx, {filename: templateFile});
const plain = value => JSON.parse(JSON.stringify(value));
const core = vm.runInContext('PLCore', ctx);
let checks = 0;
let formulas = 0;

const expectedMatrix = fixture.pythonFormulaMatrix;
const matrix = ctx.reportFormulaMatrix_(expectedMatrix.headers, expectedMatrix.sheetName);
assert.equal(matrix.length, 45);
assert(matrix.every(row => row.length === 16));
for (let offset = 0; offset < 45; offset += 1) {
  const row = offset + 4;
  for (let index = 0; index < 16; index += 1) {
    const column = ctx.columnLetter_(index + 3);
    const actual = matrix[offset][index];
    let expected = expectedMatrix.values[offset][index];
    // Apps Script uses hidden J to match Core's trimmed account names while
    // keeping the original five source columns unchanged in the workbook.
    if (typeof expected === 'string' && expected.startsWith('=')) {
      const escapedSheet = "'" + expectedMatrix.sheetName.replace(/'/g, "''") + "'";
      const accountColumn = ctx.columnLetter_(expectedMatrix.headers['GL Account']);
      expected = expected.replaceAll(`${escapedSheet}!$${accountColumn}:$${accountColumn}`, `${escapedSheet}!$J:$J`);
    }
    if (!core.DISPLAY_ROWS.includes(row)) expected = '';
    if (row === 41 && index >= 1 && index <= 12) expected = `=MAX(0,${column}36-${column}39)`;
    if (row === 46 && index >= 1 && index <= 14) expected = `=MAX(0,22%*${column}41)`;
    // SUM of per-account differences equals the difference of the totals.
    if (row === 33 && index === 15) expected = '=SUM(R11:R32)';
    if (row === 34 && index === 15) expected = '=SUM(R7:R9,R11:R32)';
    if (index === 0 && [39, 41, 42, 43, 46].includes(row)) expected = '=0';
    if (index === 0 && row === 44) expected = '=SUM(C41:C43)';
    if (index === 0 && row === 48) expected = '=C44-C46';
    if (expected === null) expected = '';
    assert.equal(actual, expected, `${column}${row}: formula matrix differs from expected output`);
    if (actual.startsWith('=')) formulas += 1;
  }
}
checks += 720;
assert.equal(matrix[37 - 4][16 - 3], '=IFERROR(P36/P4,0)', 'evaluation must use P37/GPM rate');
const basic = ctx.reportFormulaMatrix_({
  'In profit center local currency': 4, 'GL Account': 2, 'Posting Date': 3,
}, "SAP O'Brien");
assert(basic[0][1].includes("'SAP O''Brien'!$D:$D"), 'amount column and sheet name escaping');
assert(basic[0][1].includes("'SAP O''Brien'!$J:$J,$B4"), 'normalized account helper and template label reference');
assert(basic[0][1].includes("'SAP O''Brien'!$C:$C,\">=\"&DATE($C$2,1,1)"), 'posting date lower bound');
assert(basic[0][1].includes("'SAP O''Brien'!$C:$C,\"<\"&DATE($C$2,$AA$2+1,1)"), 'posting date YTD cutoff');
assert(basic[0][1].startsWith('=IF(1<=$AA$2,-SUMIFS('), 'revenue sign inversion');
assert(basic[17 - 4][1].startsWith('=IF(1<=$AA$2,SUMIFS('), 'other account retains source sign');
checks += 7;

const synthetic = fixture.workbooks.find(item => item.name === 'synthetic February YTD');
const parsed = core.parseRows(synthetic.rows, 2026, 2);
const preview = core.calculate(parsed.records, fixture.labels, 2026, 2);
const defaults = plain(core.defaultJustifications(preview, 2026));
assert.deepEqual(plain(ctx.validatedJustifications_(undefined, preview, 2026)), defaults.map(item => item.text));
assert.deepEqual(plain(ctx.validatedJustifications_([], preview, 2026)), defaults.map(item => item.text));
const changed = [...defaults].reverse().map(item => ({row: item.row, text: `  Custom ${item.row}  `}));
assert.deepEqual(plain(ctx.validatedJustifications_(changed, preview, 2026)), defaults.map(item => `Custom ${item.row}`));
const override = [{row: defaults[0].row, text: '=SUM(A1:A2)'}];
assert.equal(ctx.validatedJustifications_(override, preview, 2026)[0], '=SUM(A1:A2)', 'literal user text retained for rich-text write');
assert.deepEqual(plain(ctx.validatedSignatures_()), plain(vm.runInContext('PL_APP.signatures', ctx)));
assert.equal(ctx.validatedSignatures_({reviewerRole: '  Manager TOPM Dept.  '}).reviewerRole, 'Manager TOPM Dept.');
assert.equal(ctx.validatedSignatures_({creatorName: '=1+1'}).creatorName, '=1+1');
assert.throws(() => ctx.validatedSignatures_({creatorName: 'a'.repeat(201)}));
assert.throws(() => ctx.validatedSignatures_({creatorName: 123}));
assert.throws(() => ctx.validatedSignatures_([]));
assert.throws(() => ctx.validatedJustifications_([{row: 5, text: 'Wrong row'}], preview, 2026));
assert.throws(() => ctx.validatedJustifications_([
  {row: defaults[0].row, text: 'a'}, {row: defaults[0].row, text: 'b'},
], preview, 2026));
assert.throws(() => ctx.validatedJustifications_([{row: defaults[0].row, text: 'a'.repeat(5001)}], preview, 2026));
assert.throws(() => ctx.validatedJustifications_('not an array', preview, 2026));
checks += 14;

assert.equal(ctx.xmlAttribute_('<sheet name="PL" r:id="rId2"/>', 'r:id'), 'rId2');
assert.equal(ctx.xmlAttribute_('<sheet notname="Fake" name="PL"/>', 'name'), 'PL');
assert.equal(ctx.xmlAttribute_('<sheet name="PL"/>', 'sheetId'), '');
checks += 3;

const original = fixture.templateXml;
const documents = [];
function workbook(label, xml, sheetIndex = original.plIndex, bottom = 76) {
  const patched = ctx.patchWorkbookPrintArea_(xml, sheetIndex, bottom);
  const twice = ctx.patchWorkbookPrintArea_(patched, sheetIndex, bottom);
  assert(twice === patched, `${label}: print area must be idempotent`);
  documents.push({label, xml: patched, kind: 'workbook', sheetIndex, bottom});
}
workbook('original workbook XML', original.workbook);
workbook('workbook without existing print area', original.workbook.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/, ''));
workbook('workbook self-closing definedNames', original.workbook.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/, '<definedNames/>'));
const extraRange = '<definedName name="KeepMe">\'PL\'!$B$4</definedName><definedName name="_xlnm.Print_Area" localSheetId="987">\'Other\'!$A$1:$B$3</definedName>';
workbook('preserve other named ranges', original.workbook.replace('<definedNames>', '<definedNames>' + extraRange));
workbook('prefixed workbook namespace', original.workbook.replace(/(<\/?)([A-Za-z][\w]*)(?=[\s>\/])/g, '$1x:$2')
  .replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'));

function worksheet(label, xml) {
  const restored = ctx.restoreSheetPrintSettings_(xml);
  assert(ctx.restoreSheetPrintSettings_(restored) === restored, `${label}: print settings must be idempotent`);
  documents.push({label, xml: restored, kind: 'worksheet'});
}
worksheet('original worksheet XML', original.sheet);
const withoutSettings = original.sheet
  .replace(/<printOptions\b[^>]*(?:\/>|>[\s\S]*?<\/printOptions>)/g, '')
  .replace(/<pageMargins\b[^>]*(?:\/>|>[\s\S]*?<\/pageMargins>)/g, '')
  .replace(/<pageSetup\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetup>)/g, '')
  .replace(/<pageSetUpPr\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetUpPr>)/g, '');
worksheet('worksheet print settings missing', withoutSettings);
worksheet('worksheet sheetPr absent', withoutSettings.replace(/<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/, ''));
worksheet('worksheet sheetPr self closing', withoutSettings.replace(/<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/, '<sheetPr/>'));
worksheet('worksheet stale printer relation replaced', original.sheet.replace(/<pageSetup\b[^>]*>/,
  '<pageSetup orientation="portrait" r:id="printerOld"/>'));

documents.push({label: 'original style XML', xml: original.styles, kind: 'styleSheet'});
documents.push({label: 'patched original style XML', xml: ctx.patchNumericFormats_(original.styles), kind: 'styleSheet'});
const syntheticStyle = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="5">' +
  '<numFmt numFmtId="164" formatCode="#,##0;(#,##0);0"/>' +
  '<numFmt numFmtId="165" formatCode="0.0%;(0.0%);0.0%"/>' +
  '<numFmt numFmtId="166" formatCode="mmm-yy"/>' +
  '<numFmt numFmtId="167" formatCode="&quot;A&amp;B&quot; #,##0;(#,##0);0"/>' +
  '<numFmt numFmtId="168" formatCode="&quot;$&amp;&quot; #,##0;(#,##0);0"/>' +
  '</numFmts></styleSheet>';
const patchedStyle = ctx.patchNumericFormats_(syntheticStyle);
assert(ctx.patchNumericFormats_(patchedStyle) === patchedStyle, 'format XML must be idempotent');
documents.push({label: 'synthetic escaped numeric formats', xml: patchedStyle, kind: 'styleSheet'});

const inspected = pythonJson('inspect_xml.py', {documents});
for (const item of inspected) {
  const source = documents.find(document => document.label === item.label);
  if (source.kind === 'workbook') {
    const relevant = item.definedNames.filter(name => name.name === '_xlnm.Print_Area' && name.localSheetId === String(source.sheetIndex));
    assert.equal(relevant.length, 1, `${item.label}: one PL print area required`);
    assert.equal(relevant[0].text, `'PL'!$A$1:$S$${source.bottom}`, `${item.label}: dynamic print bottom`);
    if (item.label === 'preserve other named ranges') {
      assert(item.definedNames.some(name => name.name === 'KeepMe'));
      assert(item.definedNames.some(name => name.localSheetId === '987'));
    }
  } else if (source.kind === 'worksheet') {
    assert.equal(item.pageSetup.orientation, 'landscape', `${item.label}: original template orientation`);
    assert.equal(item.pageSetUpPr.fitToPage, '1', `${item.label}: fit-to-page retained`);
  }
  checks += 1;
}
const oldStyles = inspected.find(item => item.label === 'original style XML').numFmts;
const updatedStyles = inspected.find(item => item.label === 'patched original style XML').numFmts;
assert.deepEqual(updatedStyles.map(item => item.numFmtId), oldStyles.map(item => item.numFmtId), 'retain style format IDs');
oldStyles.forEach((item, index) => {
  assert.equal(updatedStyles[index].formatCode, core.numericFormat(item.formatCode), `XML round trip format ${item.numFmtId}`);
  checks += 1;
});
const escapedFormats = inspected.find(item => item.label === 'synthetic escaped numeric formats').numFmts;
assert.equal(escapedFormats[0].formatCode, '#,##0;\\-* #,##0_-;_-* "-"??_-');
assert.equal(escapedFormats[1].formatCode, '0.0%;\\-* 0.0%_-;_-* "-"??_-');
assert.equal(escapedFormats[2].formatCode, 'mmm-yy');
assert.equal(escapedFormats[3].formatCode.split(';')[0], '"A&B" #,##0');
assert.equal(escapedFormats[4].formatCode.split(';')[0], '"$&" #,##0');
checks += 6;

console.log(`PASS: ${checks} backend assertions, ${formulas} formula cells, ${documents.length} namespace/order-valid XML variants.`);
console.log('Pure Node/Python validation only; real Apps Script services and Google XLSX round trip remain deployment checks.');
