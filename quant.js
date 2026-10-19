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
  // 검체 종류: 시료 묶음 이름 끝의 영문 코드(28188HB → HB). 혈액은 mg/L, 조직(간 등)은 균질액 농도 × 균질화 배수 = mg/kg
  //   균질화 배수 = 균질액 부피(mL) ÷ 조직 무게(g)  예) 간 1 g + 물 4 mL → 5   (균질액 밀도 ≈ 1로 봄)
  const SPECIMENS = {
    HB: { name: '심장혈', unit: 'mg/L' }, PB: { name: '말초혈', unit: 'mg/L' }, FB: { name: '대퇴혈', unit: 'mg/L' },
    LV: { name: '간', unit: 'mg/kg', tissue: true },
  };
  function specimenOf(group, custom = {}) {
    const m = /([A-Za-z]{1,4})[\s_-]?\d{0,2}$/.exec(String(group || '')), code = m ? m[1].toUpperCase() : '';
    const base = SPECIMENS[code] || { name: code || '시료', unit: 'mg/L' };
    const c = custom[code] || {};
    const tissue = c.unit ? c.unit === 'mg/kg' : !!base.tissue;
    const f = num(c.factor);
    return { code, name: c.name || base.name, unit: tissue ? 'mg/kg' : 'mg/L', tissue, factor: tissue ? (f > 0 ? f : null) : 1 };
  }

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

  // opts: { header, map, is: '내부표준 성분명', weighting('none'|'1/x'|'1/x2', 기본 1/x), conc: { 성분: { 수준: mg/L } }, exclude: { 성분: [수준] },
  //         dil: { 시료 묶음 또는 주입 ID: 희석배수 }, specimens: { 검체 코드: { name, unit, factor(균질화 배수) } } }
  //   판정(LLOQ/ULOQ)은 희석 전 측정 농도로, 보고 농도 = 측정 × 희석배수 (조직이면 × 균질화 배수, 단위 mg/kg)
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
    const dilMap = opts.dil || {};
    const dilOf = (r) => { const v = num(dilMap[r.id] ?? dilMap[r.group]); return v > 0 ? v : 1; };
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
        const dil = dilOf(r), sp = specimenOf(r.group, opts.specimens);
        const o = { run: r.label, id: r.id, group: r.group, area, isArea: isKey ? r.area[isKey] : null, y, rt: r.rt[an.quant] ?? null, ion: ionOf(r), dil,
          specimen: sp.code, specName: sp.name, unit: sp.unit, homog: sp.tissue ? sp.factor : null, homogMissing: sp.tissue && !sp.factor };
        if (!(area > 0)) { o.flag = 'ND'; o.conc = o.measured = null; }
        else if (y == null) { o.flag = 'IS 없음'; o.conc = o.measured = null; }
        else {
          o.measured = back(y);   // 측정 용액 농도(희석 전)
          o.conc = o.measured != null ? o.measured * dil * (o.homog || 1) : null;   // 조직은 균질화 배수까지 곱해 mg/kg
          o.flag = o.measured == null ? '' : o.measured < lo ? '<LLOQ' : o.measured > hi ? '>ULOQ' : '';
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
        const ok = g.runs.filter((s) => s.conc != null && s.flag !== 'ND');
        const vals = ok.map((s) => s.conc), meas = ok.length ? mean(ok.map((s) => s.measured)) : null;
        g.dil = g.runs[0].dil;
        Object.assign(g, { specimen: g.runs[0].specimen, specName: g.runs[0].specName, unit: g.runs[0].unit, homog: g.runs[0].homog, homogMissing: g.runs[0].homogMissing });
        g.n = vals.length;
        g.mean = vals.length ? mean(vals) : null;
        g.sd = sd(vals);
        g.cv = g.sd != null && g.mean ? (g.sd / g.mean) * 100 : null;
        g.flag = vals.length ? (meas < lo ? '<LLOQ' : meas > hi ? '>ULOQ' : '')
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

  /* ---------------- 재분석 병합 ---------------- */
  // 원본에서 ULOQ 초과였던 시료 묶음은, 희석 재분석(re)에서 같은 묶음이 정량 범위 안으로 들어왔으면 그 값으로 바꾼다.
  // → { 성분: { 묶음: { mean, sd, cv, n, flag, ionFail, dil, source: 'base'|'re', orig, lo, hi } } }
  function finalTable(base, re) {
    const out = {};
    for (const [n, r] of Object.entries(base.results)) {
      out[n] = {};
      const rr = re && re.results[n];
      for (const g of r.groups) {
        const rg = rr && rr.fit ? rr.groups.find((x) => x.name === g.name) : null;
        const use = !!(rg && g.flag === '>ULOQ' && rg.mean != null && rg.flag === '');
        const src = use ? rg : g, rng = (use ? rr : r).range;
        out[n][g.name] = { mean: src.mean, sd: src.sd, cv: src.cv, n: src.n, flag: src.flag, ionFail: src.ionFail, dil: src.dil,
          specimen: g.specimen, specName: g.specName, unit: g.unit, homog: src.homog, homogMissing: g.homogMissing,
          loRep: rng.lo != null ? rng.lo * (src.dil || 1) * (src.homog || 1) : null,   // 보고 단위로 환산한 LLOQ(희석·균질화 반영)
          source: use ? 're' : 'base', orig: use ? g.mean : null, lo: rng.lo, hi: rng.hi };
      }
    }
    return out;
  }

  /* ---------------- 보고 전 점검(QC) ---------------- */
  const QC_LIMITS = { r2: 0.99, ion: ION_TOL, blank: 20, cv: 15, isLo: 50, isHi: 150 };
  // res: analyze 결과, rows/opts: 가중치 근거(가중치 없음과 비교)를 남기려고 다시 계산할 때 쓴다
  // fin: finalTable(res, re) — 희석 재분석으로 대체된 ULOQ 초과는 해결된 것으로 본다
  function qc(res, rows, opts = {}, fin = null) {
    const items = [], per = {};
    const add = (list, key, label, ok, detail) => list.push({ key, label, ok, detail });
    // 내부표준 면적 변동: 표준·시료 주입의 IS 면적이 표준 평균의 50–150 % 안
    const isA = res.is ? res.analytes.find((a) => a.name === res.is) : null;
    if (isA) {
      const stdIs = res.runs.filter((r) => r.kind === 'std').map((r) => r.area[isA.quant]).filter((v) => v > 0);
      const ref = stdIs.length ? mean(stdIs) : null;
      const pts = res.runs.filter((r) => r.kind !== 'blank').map((r) => ({ run: r.label, pct: ref ? (r.area[isA.quant] / ref) * 100 : null }));
      const bad = pts.filter((p) => p.pct == null || p.pct < QC_LIMITS.isLo || p.pct > QC_LIMITS.isHi);
      const lo = Math.min(...pts.map((p) => p.pct ?? 0)), hi = Math.max(...pts.map((p) => p.pct ?? 0));
      add(items, 'is', `내부표준(${res.is}) 면적 변동`, !bad.length, `표준 평균 대비 ${lo.toFixed(0)}–${hi.toFixed(0)} % (기준 ${QC_LIMITS.isLo}–${QC_LIMITS.isHi} %)${bad.length ? ` · 벗어남: ${bad.map((b) => b.run).join(', ')}` : ''}`);
    } else add(items, 'is', '내부표준', false, '내부표준을 지정하지 않아 면적 그대로 계산함');
    const tissue = [];
    for (const r of Object.values(res.results)) for (const g of r.groups) if (g.unit === 'mg/kg' && !tissue.some((t) => t.name === g.name)) tissue.push(g);
    if (tissue.length) {
      const miss = tissue.filter((g) => g.homogMissing);
      add(items, 'homog', '조직 균질화 배수', !miss.length, miss.length ? `입력 필요: ${miss.map((g) => `${g.name}(${g.specName})`).join(', ')} — 없으면 mg/kg 값이 균질액 농도(mg/L)와 같게 계산됨`
        : tissue.map((g) => `${g.name}(${g.specName}) ×${g.homog} → mg/kg`).join(', '));
    }
    let alt = null;
    if (rows && res.weighting !== 'none') alt = analyze(rows, { ...opts, weighting: 'none' });
    for (const [n, r] of Object.entries(res.results)) {
      const L = per[n] = [];
      if (!r.detected) { add(L, 'nd', '검출', null, '모든 주입에서 불검출 — 점검 대상 아님'); continue; }
      if (!r.fit) { add(L, 'fit', '검량선', false, '표준 농도가 없어 검량선을 세우지 못함'); continue; }
      add(L, 'r2', '검량선 R²', r.fit.r2 >= QC_LIMITS.r2, `R² = ${r.fit.r2.toFixed(5)} (기준 ≥ ${QC_LIMITS.r2})`);
      const used = r.stds.filter((s) => !s.excluded && s.nominal > 0), fails = used.filter((s) => s.pass === false), ex = r.stds.filter((s) => s.excluded);
      add(L, 'std', '표준 정확도', !fails.length && used.length >= 5,
        `${used.length}개 사용${ex.length ? `(제외 ${ex.map((s) => s.run).join(', ')})` : ''} · ${fails.length ? `부적합 ${fails.map((s) => `${s.run} ${s.accuracy.toFixed(1)} %`).join(', ')}` : `모두 ±15 % (최저 ±20 %) 안 — ${used.map((s) => `${s.accuracy.toFixed(1)}`).join(' / ')} %`}${used.length < 5 ? ' · 사용 표준 5개 미만' : ''}`);
      if (alt && alt.results[n] && alt.results[n].fit) {
        const af = alt.results[n].stds.filter((s) => !s.excluded && s.nominal > 0 && s.pass === false);
        add(L, 'weight', `가중치 ${res.weighting} 근거`, null, af.length ? `가중치 없이 계산하면 부적합 ${af.map((s) => `${s.run} ${s.accuracy.toFixed(1)} %`).join(', ')} → ${res.weighting} 사용` : `가중치 없이도 모든 표준 적합(R² ${alt.results[n].fit.r2.toFixed(5)})`);
      }
      if (r.qual) {
        const ib = r.samples.filter((s) => s.ionOk === false), it = r.samples.filter((s) => s.ionDev != null);
        const worst = it.length ? it.reduce((a, b) => (Math.abs(b.ionDev) > Math.abs(a.ionDev) ? b : a)) : null;
        add(L, 'ion', '확인 이온비', !ib.length, it.length ? `기준 이온비 ${sig(r.ionRef, 3)} · ${ib.length ? `벗어남 ${ib.map((s) => `${s.run} ${s.ionDev.toFixed(1)} %`).join(', ')}` : `모두 ±${QC_LIMITS.ion} % 안(최대 ${worst.run} ${worst.ionDev > 0 ? '+' : ''}${worst.ionDev.toFixed(1)} %)`}` : '검출된 시료 없음');
      }
      if (r.blanks.length) {
        const mx = r.blanks.reduce((a, b) => ((b.pct ?? 0) > (a.pct ?? 0) ? b : a));
        add(L, 'blank', '블랭크 잔류', !((mx.pct ?? 0) > QC_LIMITS.blank), (mx.area ?? 0) > 0 ? `최대 ${mx.run} 면적 ${Math.round(mx.area).toLocaleString('en-US')} = 최저 표준의 ${(mx.pct ?? 0).toFixed(2)} % (기준 ≤ ${QC_LIMITS.blank} %)` : `블랭크 ${r.blanks.length}회 모두 불검출`);
      }
      const reps = r.groups.filter((g) => g.n >= 2 && g.cv != null);
      if (reps.length) add(L, 'cv', '반복 정밀도(CV)', reps.every((g) => g.cv <= QC_LIMITS.cv), reps.map((g) => `${g.name} ${g.cv.toFixed(1)} %`).join(', ') + ` (기준 ≤ ${QC_LIMITS.cv} %)`);
      const out = r.groups.filter((g) => g.flag === '>ULOQ' || g.flag === '<LLOQ');
      const fixed = (g) => !!(fin && fin[n] && fin[n][g.name] && fin[n][g.name].source === 're');
      add(L, 'range', '정량 범위', !out.some((g) => g.flag === '>ULOQ' && !fixed(g)), out.length ? out.map((g) => `${g.name} ${g.flag === '>ULOQ'
        ? (fixed(g) ? `ULOQ 초과 → 희석 재분석(×${fin[n][g.name].dil}) ${sig(fin[n][g.name].mean)} mg/L로 대체` : 'ULOQ 초과(희석 재분석 필요)')
        : `LLOQ 미만(< ${sig(r.range.lo * (g.dil || 1) * (g.homog || 1))} ${g.unit || 'mg/L'} 로 보고)`}`).join(', ') : `모두 측정 농도 ${r.range.lo}–${r.range.hi} mg/L 안`);
    }
    const all = [...items, ...Object.values(per).flat()];
    return { items, analytes: per, pass: all.every((x) => x.ok !== false), fails: all.filter((x) => x.ok === false).length, limits: QC_LIMITS };
  }

  /* ---------------- .xlsx 쓰기(의존성 없음: 무압축 zip + SpreadsheetML) ---------------- */
  const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  function zipStore(files) {
    const enc = new TextEncoder(), parts = [], central = [];
    let off = 0;
    const w16 = (a, o, v) => { a[o] = v & 0xff; a[o + 1] = (v >>> 8) & 0xff; };
    const w32 = (a, o, v) => { w16(a, o, v & 0xffff); w16(a, o + 2, (v >>> 16) & 0xffff); };
    for (const f of files) {
      const name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      const lh = new Uint8Array(30 + name.length);
      w32(lh, 0, 0x04034b50); w16(lh, 4, 20); w16(lh, 6, 0x0800); w16(lh, 8, 0); w16(lh, 10, 0); w16(lh, 12, 0x21);
      w32(lh, 14, crc); w32(lh, 18, data.length); w32(lh, 22, data.length); w16(lh, 26, name.length); lh.set(name, 30);
      const ch = new Uint8Array(46 + name.length);
      w32(ch, 0, 0x02014b50); w16(ch, 4, 20); w16(ch, 6, 20); w16(ch, 8, 0x0800); w16(ch, 10, 0); w16(ch, 12, 0); w16(ch, 14, 0x21);
      w32(ch, 16, crc); w32(ch, 20, data.length); w32(ch, 24, data.length); w16(ch, 28, name.length); w32(ch, 42, off); ch.set(name, 46);
      parts.push(lh, data); central.push(ch); off += lh.length + data.length;
    }
    const cdLen = central.reduce((s, c) => s + c.length, 0), end = new Uint8Array(22);
    w32(end, 0, 0x06054b50); w16(end, 8, files.length); w16(end, 10, files.length); w32(end, 12, cdLen); w32(end, 16, off);
    const all = [...parts, ...central, end], out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
    let p = 0; for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }
  const xesc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
  const colName = (i) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  // 셀 스타일 번호(styles.xml cellXfs 순서)
  const S = { plain: 0, head: 1, title: 2, h2: 3, note: 4, text: 5, bold: 6, conc: 7, area: 8, ratio: 9, pct: 10, uloq: 11, ok: 12, bad: 13, center: 14, rt: 15 };
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="165" formatCode="0.0000"/><numFmt numFmtId="166" formatCode="0.0%"/></numFmts>
<fonts count="8"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><b/><sz val="14"/><name val="Arial"/></font><font><b/><sz val="12"/><name val="Arial"/></font><font><sz val="9"/><color rgb="FF666666"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFC65911"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF375623"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFC00000"/><name val="Arial"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="16"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="164" fontId="5" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="7" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  // rows: [[셀…]…], 셀 = 값 | { v, s } | null.  cols: 열 너비 배열, freeze: 고정할 머리글 행 수
  function sheetXml({ rows, cols = [], freeze = 0, grid = false }) {
    const cell = (c, ri, ci) => {
      if (c == null || c === '') return '';
      const o = typeof c === 'object' ? c : { v: c }, v = o.v, s = o.s ?? (typeof v === 'number' ? 0 : 0), ref = colName(ci) + (ri + 1);
      if (v == null || v === '') return o.s != null ? `<c r="${ref}" s="${s}"/>` : '';
      if (typeof v === 'number') return Number.isFinite(v) ? `<c r="${ref}" s="${s}"><v>${v}</v></c>` : `<c r="${ref}" s="${s}"/>`;
      return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
    };
    const body = rows.map((r, ri) => { const cs = (r || []).map((c, ci) => cell(c, ri, ci)).join(''); return cs ? `<row r="${ri + 1}">${cs}</row>` : ''; }).join('');
    const pane = freeze ? `<pane ySplit="${freeze}" topLeftCell="A${freeze + 1}" activePane="bottomLeft" state="frozen"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"${grid ? '' : ' showGridLines="0"'}>${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${
      cols.length ? `<cols>${cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : ''}<sheetData>${body}</sheetData></worksheet>`;
  }
  // sheets: [{ name, rows, cols, freeze, grid }] → Uint8Array(.xlsx)
  function writeXlsx(sheets) {
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const names = sheets.map((s, i) => String(s.name || `Sheet${i + 1}`).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31));
    return zipStore([
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${
        sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` },
      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets>${names.map((n, i) => `<sheet name="${xesc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>` },
      { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="${R}/styles" Target="styles.xml"/></Relationships>` },
      { name: 'xl/styles.xml', data: STYLES },
      ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s) })),
    ]);
  }

  // 결과 보고서 .xlsx — 결과 요약 · QC 점검 · 검량선 · 주입별 정량 (· 재분석) · 원본 데이터
  // opts: { file, rows(원본 시트 전체), re, reFile, reRows, qc, at }
  function resultXlsx(res, opts = {}) {
    const fin = finalTable(res, opts.re);
    const names = Object.keys(res.results);
    const groups = [];
    for (const n of names) for (const g of res.results[n].groups) if (!groups.includes(g.name)) groups.push(g.name);
    const H = (labels) => labels.map((v) => ({ v, s: S.head }));
    const T = (v) => ({ v, s: S.text }), B = (v) => ({ v, s: S.bold }), N = (v, s) => ({ v: v == null || !Number.isFinite(v) ? '' : v, s });
    const flagTxt = (f) => ({ '>ULOQ': 'ULOQ 초과', '<LLOQ': 'LLOQ 미만', ND: '불검출' }[f] ?? f ?? '');
    const levelsTxt = (() => {
      const sets = names.map((n) => res.results[n].stds.map((s) => s.nominal).join('/'));
      const same = sets.every((x) => x === sets[0]);
      return same && sets[0] ? `STD${res.levels[0]}–${res.levels[res.levels.length - 1]} = ${res.results[names[0]].stds.map((s) => s.nominal ?? '?').join(' / ')} mg/L` : '표준 농도는 성분마다 다름(검량선 시트)';
    })();
    const wTxt = { none: '가중치 없음', '1/x': '1/x 가중', '1/x2': '1/x² 가중' }[res.weighting] || res.weighting;

    // 1. 결과 요약
    const sum = [[{ v: '혈중 약물 농도 정량 결과', s: S.title }],
      [{ v: `원본: ${opts.file || ''}${opts.reFile ? ` · 재분석: ${opts.reFile}` : ''} · 내부표준 ${res.is || '없음'} · 검량선 ${wTxt} 1차식 · ${levelsTxt}${opts.at ? ` · 작성 ${opts.at}` : ''}`, s: S.note }], [],
      [{ v: '보고 표 (혈액 mg/L · 조직 mg/kg)', s: S.h2 }], H(['시료', '검체', '단위', ...names])];
    const gInfo = (g) => { for (const n of names) { const c = fin[n][g]; if (c) return c; } return {}; };
    for (const g of groups) sum.push([B(g), T(gInfo(g).specName || ''), { v: gInfo(g).unit || 'mg/L', s: S.center }, ...names.map((n) => {
      const c = fin[n][g]; if (!c) return T('');
      if (c.flag === 'ND') return { v: '불검출', s: S.center };
      if (c.flag === 'IS 없음') return { v: 'IS 없음', s: S.center };
      if (c.mean == null) return { v: '–', s: S.center };
      if (c.flag === '<LLOQ') return { v: `< ${sig(c.loRep)}`, s: S.center };
      return N(c.mean, c.flag === '>ULOQ' ? S.uloq : S.conc);
    })]);
    sum.push([{ v: '주황 칸 = ULOQ(최고 표준) 초과 — 검량 범위 밖 외삽값이므로 희석 후 재분석 권장. "< 값" = LLOQ 미만. 불검출 = 정량 이온 면적 0.', s: S.note }]);
    const repl = [];
    for (const n of names) for (const g of groups) { const c = fin[n][g]; if (c && c.source === 're') repl.push(`${g} ${n}: 원본 ${sig(c.orig)} (ULOQ 초과) → 재분석 ×${c.dil} 희석값 ${sig(c.mean)}`); }
    if (repl.length) sum.push([{ v: `재분석으로 대체: ${repl.join(' · ')}`, s: S.note }]);
    sum.push([], [{ v: '상세 (반복 통계)', s: S.h2 }], H(['성분', '시료', '검체', '단위', 'n', '희석배수', '균질화 배수', '평균', 'SD', 'CV', '판정', '확인 이온비', '출처', 'LLOQ (측정 mg/L)', 'ULOQ (측정 mg/L)']));
    for (const n of names) for (const g of groups) {
      const c = fin[n][g]; if (!c) continue;
      sum.push([T(n), T(g), T(c.specName || ''), { v: c.unit || 'mg/L', s: S.center }, { v: c.n, s: S.center }, N(c.dil, S.text), c.unit === 'mg/kg' ? (c.homog ? N(c.homog, S.text) : { v: '입력 필요', s: S.bad }) : T(''), N(c.mean, S.conc), N(c.sd, S.ratio), N(c.cv != null ? c.cv / 100 : null, S.pct),
        { v: c.flag ? flagTxt(c.flag) : c.mean != null ? '정량 범위 내' : '', s: c.flag === '>ULOQ' || c.flag === '<LLOQ' ? S.bad : S.center },
        { v: c.n ? (c.ionFail ? '벗어남 있음' : '모두 적합') : '', s: c.ionFail ? S.bad : S.center }, T(c.source === 're' ? '재분석' : '원본'), N(c.lo, S.conc), N(c.hi, S.conc)]);
    }
    sum.push([], [{ v: '계산 방법 및 가정', s: S.h2 }],
      [{ v: `· 면적비 = 정량 이온 면적 ÷ 같은 주입의 내부표준(${res.is || '없음'}) 면적. 성분명 끝 "1" = 정량 이온, "2" = 확인 이온(이온비 허용 ±${ION_TOL} %).`, s: S.note }],
      [{ v: `· 검량선: ${wTxt} 최소제곱 1차식. 표준 정확도 기준 ±15 %, 최저 농도 ±20 %. 판정(LLOQ/ULOQ)은 희석 전 측정 농도 기준, 보고값은 × 희석배수.`, s: S.note }],
      [{ v: '· 이름 끝 숫자만 다른 시료(예: 28188HB1–3)는 반복으로 묶어 평균·SD·CV. STD n = 표준, M = 블랭크.', s: S.note }],
      [{ v: '· 검체: 이름 끝 코드 — HB 심장혈 · PB 말초혈(mg/L), LV 간(mg/kg). 조직 mg/kg = 균질액 농도(mg/L) × 균질화 배수(균질액 mL ÷ 조직 g, 밀도 ≈ 1).', s: S.note }]);

    // 2. QC 점검
    const q = opts.qc || qc(res, opts.rows, {}, fin);
    const qcRows = [[{ v: `보고 전 점검 — ${q.pass ? '모두 통과' : `확인 필요 ${q.fails}건`}`, s: S.title }], [], H(['성분', '항목', '결과', '내용'])];
    const qrow = (n, it) => qcRows.push([T(n), T(it.label), { v: it.ok === true ? '통과' : it.ok === false ? '확인 필요' : '참고', s: it.ok === true ? S.ok : it.ok === false ? S.bad : S.center }, T(it.detail)]);
    for (const it of q.items) qrow('전체', it);
    for (const n of names) for (const it of q.analytes[n] || []) qrow(n, it);

    // 3. 검량선
    const cal = [[{ v: `검량선 — 면적비 = 기울기 × 농도 + 절편 (${wTxt})`, s: S.title }], [], H(['성분', '기울기', '절편', 'R²', 'LLOQ (mg/L)', 'ULOQ (mg/L)', '사용 표준 수', '기준 이온비', '표준 판정'])];
    for (const n of names) {
      const r = res.results[n], f = r.fit;
      const used = r.stds.filter((s) => !s.excluded && s.nominal > 0);
      cal.push([B(n), N(f?.slope, S.ratio), N(f?.intercept, S.ratio), N(f?.r2, S.ratio), N(f ? r.range.lo : null, S.conc), N(f ? r.range.hi : null, S.conc), N(f ? f.n : null, S.center), N(r.ionRef, S.ratio),
        { v: !r.detected ? '불검출' : !f ? '검량선 없음' : used.some((s) => s.pass === false) ? '부적합 있음' : '모두 적합', s: !r.detected || !f ? S.center : used.some((s) => s.pass === false) ? S.bad : S.ok }]);
    }
    cal.push([], H(['성분', '표준', '이론 농도', '정량 이온 면적', 'IS 면적', '면적비', '사용', '역산 농도', '정확도', '판정']));
    for (const n of names) for (const s of res.results[n].stds) cal.push([T(n), T(s.run), N(s.nominal, S.conc), N(s.area, S.area), N(s.isArea, S.area), N(s.y, S.ratio), { v: s.excluded ? '제외' : '사용', s: S.center },
      N(s.calc, S.conc), N(s.accuracy != null ? s.accuracy / 100 : null, S.pct), { v: s.excluded ? '제외' : s.pass == null ? '' : s.pass ? '적합' : `±${s.limit} % 초과`, s: s.pass === false ? S.bad : s.pass ? S.ok : S.center }]);
    const bl = names.flatMap((n) => res.results[n].blanks.map((b) => [T(n), T(b.run), N(b.area, S.area), N(b.pct != null ? b.pct / 100 : null, S.pct)]));
    if (bl.length) cal.push([], [{ v: '블랭크 잔류 (정량 이온 면적 ÷ 최저 표준 면적, 기준 ≤ 20 %)', s: S.h2 }], H(['성분', '블랭크', '면적', '최저 표준 대비']), ...bl);

    // 4. 주입별
    const runSheet = (r0, title) => {
      const rows = [[{ v: title, s: S.title }], [], H(['성분', '시료', '주입', 'RT (min)', '정량 이온 면적', 'IS 면적', '면적비', '측정 농도 (mg/L)', '희석배수', '균질화 배수', '정량값', '단위', '판정', '이온비 편차', '이온비 판정'])];
      for (const n of Object.keys(r0.results)) for (const s of r0.results[n].samples) rows.push([T(n), T(s.group), T(s.run), N(s.rt, S.rt), N(s.area, S.area), N(s.isArea, S.area), N(s.y, S.ratio), N(s.measured, S.conc), N(s.dil, S.text), s.unit === 'mg/kg' ? (s.homog ? N(s.homog, S.text) : { v: '입력 필요', s: S.bad }) : T(''), N(s.conc, S.conc), { v: s.unit, s: S.center },
        { v: flagTxt(s.flag), s: s.flag === '>ULOQ' || s.flag === '<LLOQ' ? S.bad : S.center }, N(s.ionDev != null ? s.ionDev / 100 : null, S.pct), { v: s.ionOk == null ? '' : s.ionOk ? '적합' : '벗어남', s: s.ionOk === false ? S.bad : S.center }]);
      return rows;
    };
    const RUNCOLS = [24, 12, 12, 9, 15, 13, 10, 15, 9, 11, 14, 8, 12, 11, 11];
    const rawSheet = (rows) => (rows || []).map((r) => (r || []).map((v) => (v === '' || v == null ? null : v)));
    const sheets = [
      { name: '결과 요약', rows: sum, cols: [24, 14, 9, 16, 20, 16, 16, 16, 14, 10, 14, 14, 10, 14, 14] },
      { name: 'QC 점검', rows: qcRows, cols: [22, 24, 12, 100] },
      { name: '검량선', rows: cal, cols: [24, 13, 13, 15, 13, 13, 12, 13, 13, 12] },
      { name: '주입별 정량', rows: runSheet(res, '주입별 정량 — 원본'), cols: RUNCOLS, freeze: 3 },
    ];
    if (opts.re) sheets.push({ name: '재분석 주입별', rows: runSheet(opts.re, `주입별 정량 — 재분석 (${opts.reFile || ''})`), cols: RUNCOLS, freeze: 3 });
    sheets.push({ name: '원본 데이터', rows: rawSheet(opts.rows), cols: [14, 26, 14, 8], freeze: 1, grid: true });
    if (opts.re && opts.reRows) sheets.push({ name: '재분석 원본', rows: rawSheet(opts.reRows), cols: [14, 26, 14, 8], freeze: 1, grid: true });
    return writeXlsx(sheets);
  }

  /* ---------------- 내보내기 ---------------- */
  const csvCell = (v) => { const s = v == null ? '' : String(v); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const toCsv = (rows) => '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
  const sig = (v, d = 4) => (v == null || !isFinite(v) ? '' : Number(v.toPrecision(d)).toString());

  function resultCsv(res) {
    const names = Object.keys(res.results);
    const groups = [];
    for (const n of names) for (const g of res.results[n].groups) if (!groups.includes(g.name)) groups.push(g.name);
    const out = [['정량 결과 (반복 평균, 혈액 mg/L · 조직 mg/kg)', `내부표준: ${res.is || '없음'}`, `가중치: ${res.weighting}`], ['시료', '검체', '단위', ...names]];
    const gOf = (gname) => { for (const n of names) { const g = res.results[n].groups.find((x) => x.name === gname); if (g) return g; } return {}; };
    for (const gname of groups) out.push([gname, gOf(gname).specName || '', gOf(gname).unit || 'mg/L', ...names.map((n) => {
      const g = res.results[n].groups.find((x) => x.name === gname);
      return !g ? '' : g.flag === 'ND' || g.flag === 'IS 없음' ? g.flag : `${sig(g.mean)}${g.flag ? ' ' + g.flag : ''}`;
    })]);
    out.push([], ['검량선', '기울기', '절편', 'R²', '범위(mg/L)', '사용 표준 수']);
    for (const n of names) { const r = res.results[n]; out.push([n, sig(r.fit?.slope, 6), sig(r.fit?.intercept, 6), r.fit ? r.fit.r2.toFixed(5) : '', r.fit ? `${r.range.lo}–${r.range.hi}` : '', r.fit ? r.fit.n : 0]); }
    out.push([], ['성분', '주입', '구분', '면적', 'IS 면적', '면적비', '이론 농도(mg/L)', '정량값', '단위', '정확도(%)', '이온비 편차(%)', '비고']);
    for (const n of names) {
      const r = res.results[n];
      for (const s of r.stds) out.push([n, s.run, '표준', s.area, s.isArea, sig(s.y, 6), s.nominal ?? '', sig(s.calc), 'mg/L', s.accuracy != null ? s.accuracy.toFixed(1) : '', '', s.excluded ? '검량선 제외' : s.pass === false ? `±${s.limit}% 벗어남` : '']);
      for (const s of r.samples) out.push([n, s.run, '시료', s.area, s.isArea, sig(s.y, 6), '', sig(s.conc), s.unit, '', s.ionDev != null ? s.ionDev.toFixed(1) : '', [s.flag, s.ionOk === false ? '이온비 벗어남' : ''].filter(Boolean).join(' · ')]);
    }
    return toCsv(out);
  }

  return { readTable, parseDelimited, guessColumns, buildRuns, fitLine, specimenOf, SPECIMENS, analyze, finalTable, qc, writeXlsx, resultXlsx, resultCsv, num, sig, ION_TOL, QC_LIMITS };
});
