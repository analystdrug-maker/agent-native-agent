const test = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('node:zlib');
const Q = require('./quant.js');

// 최소 xlsx(zip) 만들기 — 엑셀이 쓰는 것처럼 deflate 압축, 공유 문자열 사용
function makeXlsx(sheets) {
  const shared = [], si = (s) => { let i = shared.indexOf(s); if (i < 0) { i = shared.length; shared.push(s); } return i; };
  const col = (i) => String.fromCharCode(65 + i);
  const files = {};
  sheets.forEach((rows, n) => {
    files[`xl/worksheets/sheet${n + 1}.xml`] = `<?xml version="1.0"?><worksheet><sheetData>${rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => v === '' ? '' : typeof v === 'number'
      ? `<c r="${col(ci)}${ri + 1}"><v>${v}</v></c>` : `<c r="${col(ci)}${ri + 1}" t="s"><v>${si(v)}</v></c>`).join('')}</row>`).join('')}</sheetData></worksheet>`;
  });
  files['xl/sharedStrings.xml'] = `<sst>${shared.map((s) => `<si><t>${s.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</t></si>`).join('')}</sst>`;
  files['xl/workbook.xml'] = `<workbook><sheets>${sheets.map((_, n) => `<sheet name="시트${n + 1}" sheetId="${n + 1}" r:id="rId${n + 1}"/>`).join('')}</sheets></workbook>`;
  files['xl/_rels/workbook.xml.rels'] = `<Relationships>${sheets.map((_, n) => `<Relationship Id="rId${n + 1}" Target="worksheets/sheet${n + 1}.xml"/>`).join('')}</Relationships>`;
  const locals = [], centrals = []; let off = 0;
  for (const [name, text] of Object.entries(files)) {
    const nm = Buffer.from(name), raw = Buffer.from(text), data = zlib.deflateRawSync(raw);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(8, 8); lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(nm.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(8, 10); ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(raw.length, 24); ch.writeUInt16LE(nm.length, 28); ch.writeUInt32LE(off, 42);
    locals.push(lh, nm, data); centrals.push(ch, nm); off += 30 + nm.length + data.length;
  }
  const cd = Buffer.concat(centrals), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(centrals.length / 2, 8); end.writeUInt16LE(centrals.length / 2, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return new Uint8Array(Buffer.concat([...locals, cd, end]));
}

// 기기 결과(긴 형식) — 주입마다 내부표준 + 성분 이온 1·2
const IS = 'Desipramine-D3';
function longRows() {
  const rows = [['SA', 'Analyte', 'Area', 'RT']];
  const inj = (sa, isA, a1, a2, b1 = 0) => rows.push([sa, IS, isA, 5.6], [sa, 'Drug 1', a1, 4.4], [sa, 'Drug 2', a2, 4.4], [sa, 'Other 1', b1, 0], [sa, 'Other 2', 0, 0]);
  inj('M', 0, 500, 0);
  [[1, 0.05], [2, 0.1], [3, 0.25], [4, 0.5], [5, 1]].forEach(([l, c]) => inj(`STD${l}`, 1e6, c * 2e6, c * 1e6));   // 면적비 = 2·농도
  inj('M', 0, 0, 0);
  inj('100HB1', 1e6, 0.6e6, 0.3e6); inj('100HB2', 2e6, 1.2e6, 0.6e6); inj('100HB3', 1e6, 0.66e6, 0.6e6);   // 0.3, 0.3, 0.33 mg/L; 마지막은 이온비 이상
  inj('100PB1', 1e6, 4e6, 2e6);   // 2 mg/L → ULOQ 초과
  return rows;
}
const CONC = { Drug: { 1: 0.05, 2: 0.1, 3: 0.25, 4: 0.5, 5: 1 } };

test('긴 형식 xlsx: 열·내부표준·주입 분류·반복 묶음', async () => {
  const { sheets } = await Q.readTable('book.xlsx', makeXlsx([longRows()]));
  const rows = sheets[0].rows;
  const g = Q.guessColumns(rows);
  assert.deepEqual(g.map, { sample: 0, analyte: 1, area: 2, rt: 3 });
  const { runs, analytes, isGuess } = Q.buildRuns(rows, g);
  assert.equal(isGuess, IS);
  assert.deepEqual(analytes.map((a) => [a.name, a.quant, a.qual]), [[IS, IS, null], ['Drug', 'Drug 1', 'Drug 2'], ['Other', 'Other 1', 'Other 2']]);
  assert.deepEqual(runs.map((r) => r.label), ['M #1', 'STD1', 'STD2', 'STD3', 'STD4', 'STD5', 'M #2', '100HB1', '100HB2', '100HB3', '100PB1']);
  assert.deepEqual(runs.map((r) => r.kind).join(), 'blank,std,std,std,std,std,blank,sample,sample,sample,sample');
  assert.deepEqual([...new Set(runs.filter((r) => r.kind === 'sample').map((r) => r.group))], ['100HB', '100PB1']);
});

test('내부표준 면적비 검량선 → 반복 평균·CV·이온비·ULOQ·불검출·블랭크 잔류', () => {
  const rows = longRows();
  const r = Q.analyze(rows, { ...Q.guessColumns(rows), conc: CONC });
  assert.equal(r.is, IS);
  assert.equal(r.weighting, '1/x');   // 기본 가중치
  assert.deepEqual(r.levels, [1, 2, 3, 4, 5]);
  const d = r.results.Drug;
  assert.ok(Math.abs(d.fit.slope - 2) < 1e-9 && Math.abs(d.fit.intercept) < 1e-9 && d.fit.r2 > 0.999999);
  assert.ok(Math.abs(d.ionRef - 0.5) < 1e-12);
  const hb = d.groups.find((x) => x.name === '100HB');
  assert.equal(hb.n, 3);
  assert.ok(Math.abs(hb.mean - 0.31) < 1e-9);
  assert.ok(hb.cv > 5 && hb.cv < 6);
  assert.equal(hb.ionFail, true);
  assert.equal(d.samples[1].conc.toFixed(6), '0.300000');   // IS 면적이 두 배여도 면적비로 같은 농도
  assert.equal(d.groups.find((x) => x.name === '100PB1').flag, '>ULOQ');
  assert.equal(d.blanks[0].area, 500);
  assert.ok(Math.abs(d.blanks[0].pct - 0.5) < 1e-9);   // 500 / STD1 100000
  const o = r.results.Other;
  assert.equal(o.detected, false);
  assert.equal(o.fit, null);
  assert.equal(o.groups[0].flag, 'ND');
  assert.ok(!(IS in r.results));
});

test('농도 미입력·표준 제외·가중치·내부표준 끄기', () => {
  const rows = longRows(), g = Q.guessColumns(rows);
  assert.equal(Q.analyze(rows, g).results.Drug.fit, null);
  assert.deepEqual(Q.analyze(rows, g).results.Drug.missingConc, [1, 2, 3, 4, 5]);
  assert.equal(Q.analyze(rows, g).results.Drug.groups[0].flag, '');   // 검출됐지만 농도 미입력 — 불검출 아님
  assert.equal(Q.analyze(rows, g).results.Other.groups[0].flag, 'ND');
  const ex = Q.analyze(rows, { ...g, conc: CONC, exclude: { Drug: [5] }, weighting: '1/x2' });
  assert.equal(ex.results.Drug.fit.n, 4);
  assert.equal(ex.results.Drug.range.hi, 0.5);
  assert.equal(ex.results.Drug.stds[0].limit, 20);
  const noIs = Q.analyze(rows, { ...g, conc: CONC, is: '' });
  assert.equal(noIs.is, '');
  assert.ok(Math.abs(noIs.results.Drug.fit.slope - 2e6) < 1e-3);
  assert.ok(IS in noIs.results);
});

test('결과 CSV: 시료×성분 표 + 검량선 + 주입별', () => {
  const rows = longRows();
  const out = Q.resultCsv(Q.analyze(rows, { ...Q.guessColumns(rows), conc: CONC }));
  assert.ok(out.startsWith('\uFEFF"정량 결과 (mg/L, 반복 평균)"'));
  assert.match(out, /\r\n시료,Drug,Other\r\n100HB,0.31,ND\r\n100PB1,2 >ULOQ,ND\r\n/);
  assert.match(out, /Drug,100HB3,시료,660000,1000000,0.66,,0.33,,81.8,이온비 벗어남/);
});

test('CSV(한글·따옴표·천 단위 쉼표)도 읽는다', async () => {
  const csv = '시료,성분,피크 면적\r\nSTD1,"약물 A",1000\r\nSTD2,"약물 A","2,000"\r\n환자1,"약물 A","1,500"\r\n';
  const { sheets } = await Q.readTable('a.csv', new TextEncoder().encode('﻿' + csv));
  const r = Q.analyze(sheets[0].rows, { ...Q.guessColumns(sheets[0].rows), conc: { '약물 A': { 1: 1, 2: 2 } } });
  assert.ok(Math.abs(r.results['약물 A'].samples[0].conc - 1.5) < 1e-9);
});

test('.xls(구형 바이너리)는 안내 메시지로 거절', async () => {
  await assert.rejects(Q.readTable('a.xls', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])), /xlsx로 다시 저장/);
});
