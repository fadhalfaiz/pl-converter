'use strict';

/* Run locally with Node.js and Python, without Google services or credentials.
 * The fixture is generated in memory from the existing app and source XLSX.
 * Example: node google_apps_script/tests/parity.test.js --python <python.exe>
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {spawnSync} = require('node:child_process');

const packageRoot = path.resolve(__dirname, '..');
const projectRoot = path.resolve(packageRoot, '..');
const pythonArg = process.argv.indexOf('--python');
const python = pythonArg >= 0 ? process.argv[pythonArg + 1] : (process.env.PL_TEST_PYTHON || 'python');
assert(python, 'Pass an executable after --python.');
const generated = spawnSync(python, [path.join(__dirname, 'export_parity_fixture.py')], {
  cwd: projectRoot, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024,
});
if (generated.error) throw generated.error;
assert.equal(generated.status, 0, generated.stderr || 'Python fixture export failed.');
const fixture = JSON.parse(generated.stdout);
const context = vm.createContext({console});
vm.runInContext(fs.readFileSync(path.join(packageRoot, 'Core.gs'), 'utf8'), context, {filename: 'Core.gs'});
const core = vm.runInContext('PLCore', context);
const plain = value => JSON.parse(JSON.stringify(value));
let checks = 0;
let cells = 0;
let recordCount = 0;

function near(actual, expected, label) {
  if (expected === null) {
    assert.equal(actual, null, label);
  } else {
    assert.equal(typeof actual, 'number', `${label}: expected numeric cell`);
    assert(Number.isFinite(actual), `${label}: cell is not finite`);
    const tolerance = Math.max(1e-5, Math.abs(expected) * 1e-12);
    assert(Math.abs(actual - expected) <= tolerance,
      `${label}: ${actual} does not equal ${expected} (tolerance ${tolerance})`);
  }
  cells += 1;
}

function sameRecords(actual, expected, name) {
  assert.equal(actual.length, expected.length, `${name}: record count`);
  for (let index = 0; index < expected.length; index += 1) {
    const actualRecord = actual[index];
    const expectedRecord = expected[index];
    for (const key of ['row', 'project', 'account', 'postingDate', 'amount', 'text', 'textMonth', 'textYear', 'period', 'notes']) {
      assert.deepEqual(plain(actualRecord[key]), expectedRecord[key], `${name}: row ${expectedRecord.row} ${key}`);
    }
    recordCount += 1;
  }
}

function samePreview(actual, expected, name, year) {
  assert.deepEqual(plain(actual.columns), expected.columns, `${name}: preview columns`);
  assert.equal(actual.rows.length, expected.rows.length, `${name}: preview row count`);
  for (let index = 0; index < expected.rows.length; index += 1) {
    const row = actual.rows[index];
    const expectedRow = expected.rows[index];
    assert.equal(row.row, expectedRow.row, `${name}: row order`);
    assert.equal(row.label, expectedRow.label, `${name}: row ${expectedRow.row} template label`);
    assert.equal(row.values.length, expectedRow.values.length, `${name}: row ${expectedRow.row} column count`);
    row.values.forEach((value, column) => near(value, expectedRow.values[column],
      `${name}: row ${row.row} ${expected.columns[column]}`));
  }
  near(actual.evaluationRate, expected.evaluationRate, `${name}: P37 evaluation percentage`);
  assert.deepEqual(plain(core.defaultJustifications(actual, year)).map(entry => entry.text),
    expected.justifications, `${name}: automatic carry over justifications`);
}

for (const entry of fixture.parseCases) {
  const result = core.parsePeriod(...entry.args);
  for (const key of Object.keys(entry.expected)) {
    assert.deepEqual(plain(result[key]), entry.expected[key], `${JSON.stringify(entry.args)}: ${key}`);
  }
  checks += 1;
}

for (const workbook of fixture.workbooks) {
  const parsed = core.parseRows(workbook.rows, workbook.reportYear, workbook.ytdMonth);
  assert.equal(parsed.project, workbook.project, `${workbook.name}: project`);
  assert.deepEqual(plain(parsed.headers), workbook.headers, `${workbook.name}: shuffled header mapping`);
  sameRecords(parsed.records, workbook.records, workbook.name);
  samePreview(core.calculate(parsed.records, fixture.labels, workbook.reportYear, workbook.ytdMonth),
    workbook.preview, workbook.name, workbook.reportYear);
  let current = parsed.records;
  for (const [index, transition] of (workbook.transitions || []).entries()) {
    const before = JSON.stringify(current);
    const updated = core.applyEdits(current, transition.edits, workbook.reportYear);
    assert.equal(JSON.stringify(current), before, `${workbook.name}: edit ${index} must not mutate input`);
    sameRecords(updated, transition.records, `${workbook.name} edit ${index}`);
    samePreview(core.calculate(updated, fixture.labels, workbook.reportYear, workbook.ytdMonth),
      transition.preview, `${workbook.name} edit ${index}`, workbook.reportYear);
    current = updated;
    checks += 1;
  }
  checks += 1;
}

const synthetic = fixture.workbooks.find(workbook => workbook.name === 'synthetic February YTD');
const syntheticRecords = core.parseRows(synthetic.rows, 2026, 2).records;
function rejects(label, callback) {
  assert.throws(callback, undefined, label);
  checks += 1;
}
rejects('unknown edit row must reject', () => core.applyEdits(syntheticRecords, [{row: 987654, textMonth: 1, textYear: 2026}], 2026));
for (const month of [0, 13, 1.5, 'bogus']) {
  rejects(`invalid edited month ${month}`, () => core.applyEdits(syntheticRecords, [{row: 14, textMonth: month, textYear: 2026}], 2026));
}
for (const year of [1899, 10000, 2026.5, 'bogus']) {
  rejects(`invalid edited year ${year}`, () => core.applyEdits(syntheticRecords, [{row: 14, textMonth: 1, textYear: year}], 2026));
}
rejects('YTD 13 must reject', () => core.parseRows(synthetic.rows, 2026, 13));
rejects('missing source header must reject', () => core.parseRows([['Project'], ['Test']], 2026, 2));
rejects('six source columns must reject', () => core.parseRows(synthetic.rows.map(row => [...row, 'extra']), 2026, 2));
const missingAccount = plain(syntheticRecords);
missingAccount[0].account = 'ACCOUNT NOT IN TEMPLATE';
rejects('unmapped account must reject', () => core.calculate(missingAccount, fixture.labels, 2026, 2));
const multipleProjects = plain(synthetic.rows);
multipleProjects[2][2] = 'Second project';
rejects('multiple projects must reject', () => core.parseRows(multipleProjects, 2026, 2));

const zeros = '_-* "-"??_-';
for (const [input, positive, negative] of [
  ['General', '#,##0', '\\-* #,##0_-'],
  ['#,##0', '#,##0', '\\-* #,##0_-'],
  ['#,##0;(#,##0);0', '#,##0', '\\-* #,##0_-'],
  ['0%', '0%', '\\-* 0%_-'],
  ['0.0%;(0.0%);0.0%', '0.0%', '\\-* 0.0%_-'],
]) {
  const sections = core.numericFormat(input).split(';');
  assert.equal(sections[0], positive, `positive Excel number format for ${input}`);
  assert.equal(sections[1], negative, `spaced negative sign for ${input}`);
  assert.equal(sections[2], zeros, `zero dash alignment for ${input}`);
  assert(!sections[1].includes('('), `negative values must not use parentheses for ${input}`);
  checks += 1;
}
assert.equal(core.numericFormat('mmm-yy'), 'mmm-yy', 'date display must retain its format');
checks += 1;

// These are syntax and deployment contract checks, not Google API execution.
const backendPath = path.join(packageRoot, 'Code.gs');
if (fs.existsSync(backendPath)) {
  const source = fs.readFileSync(backendPath, 'utf8');
  new vm.Script(source, {filename: 'Code.gs'});
  assert(/function\s+doGet\s*\(/.test(source), 'Apps Script HTML entry point missing');
  checks += 1;
}
const manifestPath = path.join(packageRoot, 'appsscript.json');
if (fs.existsSync(manifestPath)) {
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  assert.equal(manifest.runtimeVersion, 'V8', 'Apps Script runtime must support Core.gs');
  assert(manifest.oauthScopes.includes('https://www.googleapis.com/auth/drive'), 'Drive conversion/export scope missing');
  checks += 1;
}

console.log(`PASS: ${checks} scenarios; ${recordCount} parsed records; ${cells} PL cells / evaluation values matched the Python app.`);
console.log('Validated locally with Node VM; live Apps Script / Drive conversion and XLSX export require Google deployment.');
