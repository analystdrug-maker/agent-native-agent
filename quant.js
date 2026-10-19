// 약물 정량(혈중 농도, mg/L) — 기기(LC-MS/MS) 결과 엑셀에서 내부표준 면적비로 검량선을 세우고 시료 농도를 역산한다.
// 의존성 없음: 브라우저(대시보드 › 약물 정량 탭)와 Node(quant.test.cjs)에서 같은 코드가 돈다.
//
// 입력은 "긴 형식": 한 행 = 한 주입(시료)의 한 성분(이온).  예) SA | Analyte | Area | RT
//   28188HB1 | Desipramine-D3 | 1480000 | 5.64     ← 내부표준
//   28188HB1 | Quetiapine 1   | 17600000 | 4.48    ← 정량 이온
//   28188HB1 | Quetiapine 2   | 9630000  | 4.48    ← 확인 이온
// 같은 SA가 연달아 나오는 행 묶음이 한 번의 주입이다(M 블랭크처럼 같은 이름이 여러 번 나와도 따로 센다).
//
//   readTable(name, bytes)    → { sheets: [{ name, rows }] }   .xlsx(내장 zip 해제) / .csv / .tsv
//   guessColumns(rows)        → { header, columns, map }        머리글 줄과 열 역할 추정
//   buildRuns(rows, opts)     → { runs, analytes, isGuess }      주입·성분·이온 정리
//   analyze(rows, opts)       → { runs, analytes, is, levels, results }  성분마다 검량선 + 역산 + 반복 평균
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Quant = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------------- 파일 → 표 ---------------- */
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  // zip 중앙 디렉터리만 읽어 { 경로: () => Promise<Uint8Array> } 를 만든다(필요한 파일만 푼다)
  function readZip(b) {
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (u32(b, i) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('엑셀(.xlsx) 파일이 아니거나 손상되었습니다');
    const n = u16(b, eocd + 10);
    let p = u32(b, eocd + 16);
    const files = {};
    const dec = new TextDecoder();
    for (let k = 0; k < n; k++) {
      if (u32(b, p) !== 0x02014b50) break;
      const method = u16(b, p + 10), csize = u32(b, p + 20), nlen = u16(b, p + 28), elen = u16(b, p + 30), clen = u16(b, p + 32), off = u32(b, p + 42);
      const name = dec.decode(b.subarray(p + 46, p + 46 + nlen));
      files[name] = async () => {
        const start = off + 30 + u16(b, off + 26) + u16(b, off + 28);
        const data = b.subarray(start, start + csize);
        if (method === 0) return data;
        if (method === 8) return inflateRaw(data);
        throw new Error(`지원하지 않는 압축 방식(${method})`);
      };
      p += 46 + nlen + elen + clen;
    }
    return files;
  }

  const unesc = (s) => s.replace(/&(lt|gt|amp|quot|apos|#x[0-9a-f]+|#\d+);/gi, (m, e) => {
    const l = e.toLowerCase();
    if (l === 'lt') return '<'; if (l === 'gt') return '>'; if (l === 'amp') return '&'; if (l === 'quot') return '"'; if (l === 'apos') return "'";
    return String.fromCodePoint(l[1] === 'x' ? parseInt(l.slice(2), 16) : parseInt(l.slice(1), 10));
  });
  const attr = (tag, a) => { const m = new RegExp(`\\s${a}="([^"]*)"`).exec(tag); return m ? unesc(m[1]) : null; };
  const texts = (xml) => { let s = ''; xml.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (_, t) => { s += unesc(t); }); return s; };
  const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/\d+$/, '')) n = n * 26 + (ch.toUpperCase().charCodeAt(0) - 64); return n - 1; };

  async function readXlsx(bytes) {
    const zip = readZip(bytes), dec = new TextDecoder();
    const text = async (p) => (zip[p] ? dec.decode(await zip[p]()) : '');
    const shared = [];
    (await text('xl/sharedStrings.xml')).replace(/<si>([\s\S]*?)<\/si>/g, (_, si) => { shared.push(texts(si)); });
    const rels = {};
    (await text('xl/_rels/workbook.xml.rels')).replace(/<Relationship\b[^>]*>/g, (t) => { rels[attr(t, 'Id')] = attr(t, 'Target'); });
    const book = await text('xl/workbook.xml');
    const list = [];
    book.replace(/<sheet\b[^>]*>/g, (t) => { list.push({ name: attr(t, 'name') || 'Sheet', rid: attr(t, 'r:id') }); });
    const sheets = [];
    for (const s of list) {
      let target = rels[s.rid] || '';
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await text(target);
      if (!xml) continue;
      const rows = [];
      xml.replace(/<row\b([^>]*)>([\s\S]*?)<\/row>/g, (_, ra, body) => {
        const rAttr = attr(ra, 'r'), ri = rAttr ? Number(rAttr) - 1 : rows.length;
        const row = [];
        let ci = 0;
        body.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (__, ca, inner = '') => {
          const ref = attr(ca, 'r'); if (ref) ci = colIndex(ref);
          const t = attr(ca, 't'), v = /<v>([\s\S]*?)<\/v>/.exec(inner);
          let val = '';
          if (t === 's') val = v ? shared[Number(v[1])] ?? '' : '';
          else if (t === 'inlineStr') val = texts(inner);
          else if (t === 'b') val = v ? (v[1] === '1' ? 'TRUE' : 'FALSE') : '';
          else if (v) { val = unesc(v[1]); if (t !== 'str' && t !== 'e' && val !== '' && isFinite(val)) val = Number(val); }
          row[ci++] = val;
        });
        for (let i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = '';
        rows[ri] = row;
      });
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      sheets.push({ name: s.name, rows });
    }
    if (!sheets.length) throw new Error('시트를 찾지 못했습니다');
    return sheets;
  }

  function decodeText(bytes) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, ''); }
    catch { return new TextDecoder('euc-kr').decode(bytes); }   // 한글 엑셀이 저장한 CSV(CP949)
  }
  function parseDelimited(text) {
    const first = text.split(/\r?\n/, 1)[0] || '';
    const sep = (first.match(/\t/g) || []).length > (first.match(/,/g) || []).length ? '\t' : (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ',';
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; continue; }
      if (ch === '"') q = true;
      else if (ch === sep) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  async function readTable(name, bytes) {
    bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) return { sheets: await readXlsx(bytes) };
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('예전 엑셀 형식(.xls)은 읽을 수 없습니다 — 엑셀에서 .xlsx로 다시 저장해 주세요');
    return { sheets: [{ name: String(name || 'CSV').replace(/\.[^.]+$/, ''), rows: parseDelimited(decodeText(bytes)) }] };
  }

  /* ---------------- 열 역할 추정 ---------------- */
  const ROLES = {
    sample: /^sa$|sample|시료|검체|injection|file|data\s*file|^id$|^name$/i,
    analyte: /analyte|compound|component|성분|약물|물질|target/i,
    area: /면적|피크|피이크|area|response/i,
    rt: /^rt$|retention|머무름|r\.t\./i,
  };
  const ROLE_ORDER = ['analyte', 'rt', 'area', 'sample'];

  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v ?? '').trim().replace(/,/g, '');
    if (!s || !/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
    return Number(s);
  }

  function guessColumns(rows, forcedHeader) {
    let header = 0, best = -1;
    if (forcedHeader >= 0) header = forcedHeader;
    else for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const r = rows[i] || [];
      const score = r.filter((c) => typeof c === 'string' && c.trim() && num(c) === null && Object.values(ROLES).some((re) => re.test(c.trim()))).length;
      if (score > best) { best = score; header = i; }
    }
    const head = (rows[header] || []).map((c) => String(c ?? '').trim());
    const map = { sample: -1, analyte: -1, area: -1, rt: -1 };
    const taken = new Set();
    for (const role of ROLE_ORDER) {
      const idx = head.findIndex((h, i) => !taken.has(i) && h && ROLES[role].test(h));
      if (idx >= 0) { map[role] = idx; taken.add(idx); }
    }
    // 시료 열을 못 찾으면 남은 글자 열 중 첫 번째
    if (map.sample < 0) {
      const data = rows.slice(header + 1, header + 40);
      const idx = head.findIndex((_, i) => !taken.has(i) && data.some((r) => r[i] !== '' && r[i] != null && num(r[i]) === null));
      if (idx >= 0) map.sample = idx;
    }
    return { header, columns: head, map };
  }

  /* ---------------- 주입 · 성분 정리 ---------------- */
  const STD_RE = /^(?:std|cal|standard|표준)[\s_-]*(\d+)$/i;
  const BLANK_RE = /^(?:m|mb|blk|blank|solvent|meoh|공시료|블랭크|바탕)[\s_-]*\d*$/i;
  const IS_RE = /[-_\s(]d\d{1,2}\)?$|\bistd\b|internal|내부\s*표준|13c|15n/i;
  // "Quetiapine 1" → { base: 'Quetiapine', ion: 1 }
  const splitIon = (name) => { const m = /^(.*\S)\s+(\d)$/.exec(name); return m ? { base: m[1], ion: Number(m[2]) } : { base: name, ion: 1 }; };
  // 반복 시료 묶음: 28188HB1·28188HB2·28188HB3 → 28188HB
  const groupKey = (id) => { const m = /^(.*[A-Za-z가-힣])[\s_-]?(\d{1,2})$/.exec(id); return m ? m[1] : id; };

  function buildRuns(rows, opts = {}) {
    const { header = 0, map = {} } = opts;
    const runs = [];
    let cur = null;
    rows.slice(header + 1).forEach((r, i) => {
      if (!r) return;
      const id = String(r[map.sample] ?? '').trim(), an = String(r[map.analyte] ?? '').trim();
      if (!id || !an) return;
      if (!cur || cur.id !== id || an in cur.area) { cur = { id, line: header + i + 2, area: {}, rt: {} }; runs.push(cur); }
      cur.area[an] = num(r[map.area]);
      if (map.rt >= 0) cur.rt[an] = num(r[map.rt]);
    });
    const count = {}, seen = {};
    for (const r of runs) count[r.id] = (count[r.id] || 0) + 1;
    for (const r of runs) {
      seen[r.id] = (seen[r.id] || 0) + 1;
      r.label = count[r.id] > 1 ? `${r.id} #${seen[r.id]}` : r.id;
      const s = STD_RE.exec(r.id);
      r.kind = s ? 'std' : BLANK_RE.test(r.id) ? 'blank' : 'sample';
      if (s) r.level = Number(s[1]);
    }
    // 반복 묶음은 같은 이름 틀을 가진 시료가 2개 이상일 때만
    const gCount = {};
    for (const r of runs) if (r.kind === 'sample') gCount[groupKey(r.id)] = (gCount[groupKey(r.id)] || 0) + 1;
    for (const r of runs) if (r.kind === 'sample') r.group = gCount[groupKey(r.id)] > 1 ? groupKey(r.id) : r.id;

    const names = [];
    for (const r of runs) for (const an of Object.keys(r.area)) if (!names.includes(an)) names.push(an);
    const analytes = [];
    for (const n of names) {
      const { base, ion } = splitIon(n);
      let a = analytes.find((x) => x.name === base);
      if (!a) { a = { name: base, ions: [] }; analytes.push(a); }
      a.ions.push({ ion, key: n });
    }
    for (const a of analytes) { a.ions.sort((x, y) => x.ion - y.ion); a.quant = a.ions[0].key; a.qual = a.ions[1] ? a.ions[1].key : null; }
    const isGuess = (analytes.find((a) => IS_RE.test(a.name)) || {}).name || '';
    return { runs, analytes, isGuess };
  }

  /* ---------------- 검량선 · 역산 ---------------- */
  // 가중 최소제곱 y = a·x + b  (w = 1, 1/x, 1/x²)
  function fitLine(pts, weighting) {
    const w = (x) => (weighting === '1/x' ? 1 / x : weighting === '1/x2' ? 1 / (x * x) : 1);
    let Sw = 0, Sx = 0, Sy = 0, Sxx = 0, Sxy = 0;
    for (const p of pts) { const wi = w(p.x); Sw += wi; Sx += wi * p.x; Sy += wi * p.y; Sxx += wi * p.x * p.x; Sxy += wi * p.x * p.y; }
    const d = Sw * Sxx - Sx * Sx;
    if (!(Math.abs(d) > 0)) return null;
    const slope = (Sw * Sxy - Sx * Sy) / d, intercept = (Sy - slope * Sx) / Sw;
    const my = pts.reduce((s, p) => s + p.y, 0) / pts.length;
    let ssr = 0, sst = 0;
    for (const p of pts) { ssr += (p.y - (slope * p.x + intercept)) ** 2; sst += (p.y - my) ** 2; }
    return { slope, intercept, r2: sst > 0 ? 1 - ssr / sst : 1, n: pts.length, weighting: weighting || 'none' };
  }

  const ION_TOL = 20;   // 확인 이온 비 허용 편차(상대 %)
  const mean = (a) => a.reduce((s, x) => s + x, 0) / a.length;
  const sd = (a) => (a.length > 1 ? Math.sqrt(a.reduce((s, x) => s + (x - mean(a)) ** 2, 0) / (a.length - 1)) : null);

  // opts: { header, map, is: '내부표준 성분명', weighting('none'|'1/x'|'1/x2', 기본 1/x), conc: { 성분: { 수준: mg/L } }, exclude: { 성분: [수준] } }
  function analyze(rows, opts = {}) {
    const { runs, analytes, isGuess } = buildRuns(rows, opts);
    const is = opts.is === '' ? '' : opts.is != null && analytes.some((a) => a.name === opts.is) ? opts.is : isGuess;   // '' = 내부표준 쓰지 않음
    const isKey = is ? analytes.find((a) => a.name === is).quant : null;
    const weighting = opts.weighting || '1/x';   // 기본 1/x — 낮은 농도 표준의 정확도를 지킨다(가중치 없음이면 STD1이 ±20 %를 넘기 쉽다)
    const levels = [...new Set(runs.filter((r) => r.kind === 'std').map((r) => r.level))].sort((a, b) => a - b);
    const ratio = (run, key) => {
      const a = run.area[key];
      if (a == null) return null;
      if (!isKey) return a;
      const i = run.area[isKey];
      return i > 0 ? a / i : null;
    };
    const results = {};
    for (const an of analytes) {
      if (an.name === is) continue;
      const conc = (opts.conc || {})[an.name] || {};
      const excl = (opts.exclude || {})[an.name] || [];
      const ionOf = (run) => (an.qual && run.area[an.quant] > 0 && run.area[an.qual] > 0 ? run.area[an.qual] / run.area[an.quant] : null);
      const stds = runs.filter((r) => r.kind === 'std').map((r) => {
        const nominal = num(conc[r.level]);
        return { run: r.label, level: r.level, nominal, area: r.area[an.quant], isArea: isKey ? r.area[isKey] : null, y: ratio(r, an.quant), ion: ionOf(r), excluded: excl.includes(r.level) };
      });
      const used = stds.filter((s) => !s.excluded && s.nominal > 0 && s.y != null && s.area > 0);
      const fit = used.length >= 2 ? fitLine(used.map((s) => ({ x: s.nominal, y: s.y })), weighting) : null;
      const lo = used.length ? Math.min(...used.map((s) => s.nominal)) : null;
      const hi = used.length ? Math.max(...used.map((s) => s.nominal)) : null;
      const back = (y) => (fit && fit.slope ? (y - fit.intercept) / fit.slope : null);
      const ionRefs = used.map((s) => s.ion).filter((x) => x != null);
      const ionRef = ionRefs.length ? mean(ionRefs) : null;
      for (const s of stds) {
        s.calc = s.y != null ? back(s.y) : null;
        s.accuracy = s.calc != null && s.nominal > 0 ? (s.calc / s.nominal) * 100 : null;
        s.limit = s.nominal === lo ? 20 : 15;   // 생체시료 분석법 밸리데이션: LLOQ ±20 %, 그 외 ±15 %
        s.pass = s.accuracy != null && !s.excluded ? Math.abs(s.accuracy - 100) <= s.limit : null;
      }
      const samples = runs.filter((r) => r.kind === 'sample').map((r) => {
        const area = r.area[an.quant], y = ratio(r, an.quant);
        const o = { run: r.label, id: r.id, group: r.group, area, isArea: isKey ? r.area[isKey] : null, y, rt: r.rt[an.quant] ?? null, ion: ionOf(r) };
        if (!(area > 0)) { o.flag = 'ND'; o.conc = null; }
        else if (y == null) { o.flag = 'IS 없음'; o.conc = null; }
        else {
          o.conc = back(y);
          o.flag = o.conc == null ? '' : o.conc < lo ? '<LLOQ' : o.conc > hi ? '>ULOQ' : '';
        }
        o.ionDev = o.ion != null && ionRef ? (o.ion / ionRef - 1) * 100 : null;
        o.ionOk = o.ionDev != null ? Math.abs(o.ionDev) <= ION_TOL : null;
        return o;
      });
      const groups = [];
      for (const s of samples) {
        let g = groups.find((x) => x.name === s.group);
        if (!g) { g = { name: s.group, runs: [] }; groups.push(g); }
        g.runs.push(s);
      }
      for (const g of groups) {
        const vals = g.runs.filter((s) => s.conc != null && s.flag !== 'ND').map((s) => s.conc);
        g.n = vals.length;
        g.mean = vals.length ? mean(vals) : null;
        g.sd = sd(vals);
        g.cv = g.sd != null && g.mean ? (g.sd / g.mean) * 100 : null;
        g.flag = vals.length ? (g.mean < lo ? '<LLOQ' : g.mean > hi ? '>ULOQ' : '')
          : g.runs.every((s) => s.flag === 'ND') ? 'ND' : g.runs.some((s) => s.flag === 'IS 없음') ? 'IS 없음' : '';   // '' = 검출됐지만 검량선 없음
        g.ionFail = g.runs.some((s) => s.ionOk === false);
      }
      // 블랭크 잔류: 블랭크 정량 이온 면적 ÷ 최저 표준 면적
      const lowStd = used.filter((s) => s.nominal === lo);
      const lowArea = lowStd.length ? mean(lowStd.map((s) => s.area)) : null;
      const blanks = runs.filter((r) => r.kind === 'blank').map((r) => ({ run: r.label, area: r.area[an.quant] ?? null, pct: lowArea && r.area[an.quant] != null ? (r.area[an.quant] / lowArea) * 100 : null }));
      const detected = samples.some((s) => s.area > 0) || stds.some((s) => s.area > 0);
      results[an.name] = { name: an.name, quant: an.quant, qual: an.qual, fit, range: { lo, hi }, stds, samples, groups, blanks, ionRef, detected,
        missingConc: stds.filter((s) => !(s.nominal > 0)).map((s) => s.level) };
    }
    return { runs, analytes, is, levels, weighting, results };
  }

  /* ---------------- 내보내기 ---------------- */
  const csvCell = (v) => { const s = v == null ? '' : String(v); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const toCsv = (rows) => '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
  const sig = (v, d = 4) => (v == null || !isFinite(v) ? '' : Number(v.toPrecision(d)).toString());

  function resultCsv(res) {
    const names = Object.keys(res.results);
    const groups = [];
    for (const n of names) for (const g of res.results[n].groups) if (!groups.includes(g.name)) groups.push(g.name);
    const out = [['정량 결과 (mg/L, 반복 평균)', `내부표준: ${res.is || '없음'}`, `가중치: ${res.weighting}`], ['시료', ...names]];
    for (const gname of groups) out.push([gname, ...names.map((n) => {
      const g = res.results[n].groups.find((x) => x.name === gname);
      return !g ? '' : g.flag === 'ND' || g.flag === 'IS 없음' ? g.flag : `${sig(g.mean)}${g.flag ? ' ' + g.flag : ''}`;
    })]);
    out.push([], ['검량선', '기울기', '절편', 'R²', '범위(mg/L)', '사용 표준 수']);
    for (const n of names) { const r = res.results[n]; out.push([n, sig(r.fit?.slope, 6), sig(r.fit?.intercept, 6), r.fit ? r.fit.r2.toFixed(5) : '', r.fit ? `${r.range.lo}–${r.range.hi}` : '', r.fit ? r.fit.n : 0]); }
    out.push([], ['성분', '주입', '구분', '면적', 'IS 면적', '면적비', '이론 농도(mg/L)', '정량값(mg/L)', '정확도(%)', '이온비 편차(%)', '비고']);
    for (const n of names) {
      const r = res.results[n];
      for (const s of r.stds) out.push([n, s.run, '표준', s.area, s.isArea, sig(s.y, 6), s.nominal ?? '', sig(s.calc), s.accuracy != null ? s.accuracy.toFixed(1) : '', '', s.excluded ? '검량선 제외' : s.pass === false ? `±${s.limit}% 벗어남` : '']);
      for (const s of r.samples) out.push([n, s.run, '시료', s.area, s.isArea, sig(s.y, 6), '', sig(s.conc), '', s.ionDev != null ? s.ionDev.toFixed(1) : '', [s.flag, s.ionOk === false ? '이온비 벗어남' : ''].filter(Boolean).join(' · ')]);
    }
    return toCsv(out);
  }

  return { readTable, parseDelimited, guessColumns, buildRuns, fitLine, analyze, resultCsv, num, sig, ION_TOL };
});
