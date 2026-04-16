#!/usr/bin/env node
/**
 * Rebuild BUILTIN_PER_PERSON and BUILTIN_DATA in index.html from an EF tracker .xlsx.
 * Context (ATTENDEES, NAME_ALIASES, REGION_MAP) is read from index.html.
 *
 * Usage:
 *   node tools/aem-summit-meetings/build-builtin.mjs <file.xlsx> [--sheet NAME] [--write]
 *   [--banner-date "14 Apr 2026"] [--banner-file "… · sheet 4_14"]
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const out = { positional: [], flags: {} };
  for (let i = 2; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--write') out.flags.write = true;
    else if (a === '--sheet' && argv[i + 1]) out.flags.sheet = argv[++i];
    else if (a === '--html' && argv[i + 1]) out.flags.html = argv[++i];
    else if (a === '--banner-date' && argv[i + 1]) out.flags.bannerDate = argv[++i];
    else if (a === '--banner-file' && argv[i + 1]) out.flags.bannerFile = argv[++i];
    else if (a.startsWith('-')) throw new Error(`Unknown flag: ${a}`);
    else out.positional.push(a);
  }
  return out;
}

function skipBalancedJsValue(html, start) {
  const open = html[start];
  if (open !== '{' && open !== '[') throw new Error(`Expected {{ or [ at ${start}`);
  const stack = [open === '{' ? '}' : ']'];
  let i = start + 1;
  let inStr = false;
  let q = '';
  let esc = false;
  while (i < html.length && stack.length) {
    const c = html[i];
    if (inStr) {
      if (esc) { esc = false; i += 1; continue; }
      if (c === '\\') { esc = true; i += 1; continue; }
      if (c === q) inStr = false;
      i += 1;
      continue;
    }
    if (c === '"' || c === "'") {
      inStr = true;
      q = c;
      i += 1;
      continue;
    }
    if (c === '{') { stack.push('}'); i += 1; continue; }
    if (c === '[') { stack.push(']'); i += 1; continue; }
    if (c === '}' || c === ']') {
      const want = stack.pop();
      if (c !== want) throw new Error(`Bracket mismatch at ${i}`);
      i += 1;
      continue;
    }
    i += 1;
  }
  if (stack.length) throw new Error('Unclosed bracket/brace');
  return i - 1;
}

function extractTopLevelValue(html, decl) {
  const re = new RegExp(`(?:const|let)\\s+${decl}\\s*=\\s*`);
  const m = html.match(re);
  if (!m) throw new Error(`Could not find declaration: ${decl}`);
  const valueStart = m.index + m[0].length;
  const valueEnd = skipBalancedJsValue(html, valueStart);
  const valueSrc = html.slice(valueStart, valueEnd + 1);
  const semi = html.indexOf(';', valueEnd + 1);
  if (semi === -1) throw new Error(`No semicolon after ${decl}`);
  return { valueSrc, endSemi: semi, declStart: m.index };
}

function evalValue(src) {
  return new Function(`return (${src})`)();
}

function loadContextFromHtml(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const attendees = evalValue(extractTopLevelValue(html, 'ATTENDEES').valueSrc);
  const nameAliases = evalValue(extractTopLevelValue(html, 'NAME_ALIASES').valueSrc);
  const regionMap = evalValue(extractTopLevelValue(html, 'REGION_MAP').valueSrc);
  return { html, attendees, nameAliases, regionMap };
}

function findBestSheet(wb) {
  const today = new Date();
  const dated = wb.SheetNames.map((name) => {
    const mm = name.match(/(\d{1,2})_(\d{1,2})/);
    if (!mm) return null;
    const month = parseInt(mm[1], 10);
    const day = parseInt(mm[2], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    const d = new Date(today.getFullYear(), month - 1, day);
    if (d > today) d.setFullYear(today.getFullYear() - 1);
    return { name, date: d, diff: Math.abs(today - d) };
  }).filter(Boolean);
  if (!dated.length) return null;
  dated.sort((a, b) => a.diff - b.diff);
  return dated[0].name;
}

function pickSheet(wb, explicitSheet) {
  if (explicitSheet) {
    if (!wb.SheetNames.includes(explicitSheet)) {
      throw new Error(`Sheet not found: "${explicitSheet}". Available: ${wb.SheetNames.join(', ')}`);
    }
    return explicitSheet;
  }
  const best = findBestSheet(wb);
  if (best) return best;
  for (const sn of wb.SheetNames) {
    const header = (XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, range: 0 })[0] || []);
    if (header.includes('NAME-ROLE-COMPANY')) return sn;
  }
  return null;
}

function processRows(rows, ATTENDEES, NAME_ALIASES, REGION_MAP) {
  const attendeeSet = new Set(ATTENDEES.map((a) => a.name));
  function normName(raw) {
    if (!raw) return '';
    let n = String(raw).split(' - ')[0];
    n = n.replace(/[\u00a0\u2000-\u200b\u202f\u205f\u3000]/g, ' ');
    n = n.replace(/ +/g, ' ').trim();
    n = NAME_ALIASES[n] || n;
    n = n.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    return n;
  }
  const codeSummitAttendees = {};
  rows.forEach((row) => {
    const name = normName(row['NAME-ROLE-COMPANY']);
    if (!attendeeSet.has(name)) return;
    const code = row.CODE;
    if (!code) return;
    if (!codeSummitAttendees[code]) codeSummitAttendees[code] = new Set();
    codeSummitAttendees[code].add(name);
  });
  const multiCodes = new Set(Object.entries(codeSummitAttendees)
    .filter(([, s]) => s.size > 1)
    .map(([c]) => c));

  const meetingsByCode = {};
  rows.forEach((row) => {
    const code = String(row.CODE || '');
    if (!code) return;
    if (!meetingsByCode[code]) {
      meetingsByCode[code] = {
        code,
        title: String(row.TITLE || '').trim(),
        account: String(row['AS26 ACCOUNT NAME'] || '').trim(),
        all_attendees: new Set(),
        summit_attendees: new Set(),
      };
    }
    const name = normName(row['NAME-ROLE-COMPANY']);
    meetingsByCode[code].all_attendees.add(name);
    if (attendeeSet.has(name)) meetingsByCode[code].summit_attendees.add(name);
  });
  Object.values(meetingsByCode).forEach((m) => {
    m.region = REGION_MAP[m.account] || 'Unknown';
    m.all_attendees = [...m.all_attendees].sort();
    m.summit_attendees = [...m.summit_attendees].sort();
  });

  const newPerPerson = {};
  ATTENDEES.forEach((a) => { newPerPerson[a.name] = []; });
  rows.forEach((row) => {
    const name = normName(row['NAME-ROLE-COMPANY']);
    if (!attendeeSet.has(name)) return;
    const code = String(row.CODE || '');
    if (meetingsByCode[code] && !newPerPerson[name].some((m) => m.code === code)) {
      newPerPerson[name].push({
        ...meetingsByCode[code],
        all_attendees: [...meetingsByCode[code].all_attendees],
        summit_attendees: [...meetingsByCode[code].summit_attendees],
      });
    }
  });
  Object.keys(newPerPerson).forEach((name) => {
    newPerPerson[name].sort((a, b) => a.account.localeCompare(b.account));
  });

  const agg = {};
  rows.forEach((row) => {
    const name = normName(row['NAME-ROLE-COMPANY']);
    if (!attendeeSet.has(name)) return;
    const code = row.CODE;
    const acct = String(row['AS26 ACCOUNT NAME'] || '').trim();
    const region = REGION_MAP[acct] || 'Unknown';
    if (region === 'Unknown') return;
    const isMulti = multiCodes.has(code);
    if (!agg[name]) agg[name] = { codes: new Set(), multi_codes: new Set(), byRegion: {} };
    if (code) {
      agg[name].codes.add(code);
      if (isMulti) agg[name].multi_codes.add(code);
    }
    if (region) {
      if (!agg[name].byRegion[region]) {
        agg[name].byRegion[region] = {
          codes: new Set(), multiCodes: new Set(), accounts: new Set(), multiAccounts: new Set(),
        };
      }
      if (code) {
        agg[name].byRegion[region].codes.add(code);
        if (isMulti) agg[name].byRegion[region].multiCodes.add(code);
      }
      if (acct) {
        agg[name].byRegion[region].accounts.add(acct);
        if (isMulti) agg[name].byRegion[region].multiAccounts.add(acct);
      }
    }
  });
  const meetingMap = {};
  Object.entries(agg).forEach(([name, d]) => {
    const by_region = {};
    const multi_by_region = {};
    Object.entries(d.byRegion).forEach(([r, rd]) => {
      by_region[r] = {
        count: rd.codes.size,
        multi_count: rd.multiCodes.size,
        accounts: [...rd.accounts].sort(),
      };
      if (rd.multiCodes.size > 0) {
        multi_by_region[r] = { count: rd.multiCodes.size, accounts: [...rd.multiAccounts].sort() };
      }
    });
    meetingMap[name] = {
      total: d.codes.size,
      multi_total: d.multi_codes.size,
      by_region,
      multi_by_region,
    };
  });

  return { perPersonMeetings: newPerPerson, builtinData: meetingMap };
}

function replaceConst(html, constName, newRhs) {
  const { declStart, endSemi } = extractTopLevelValue(html, constName);
  const head = html.slice(0, declStart);
  const tail = html.slice(endSemi + 1);
  return `${head}const ${constName} = ${newRhs};${tail}`;
}

function maybePatchBanner(html, bannerDate, bannerFile, sourceLine) {
  let h = html;
  if (bannerDate) {
    h = h.replace(
      /(<strong id="dataBannerDate">)([^<]*)(<\/strong>)/,
      `$1${bannerDate}$3`,
    );
  }
  if (bannerFile) {
    h = h.replace(
      /(<span class="admin-data-file" id="dataBannerFile">)([^<]*)(<\/span>)/,
      `$1${bannerFile}$3`,
    );
  }
  if (sourceLine) {
    h = h.replace(
      /(<span id="dataSourceText">)([^<]*)(<\/span>)/,
      `$1${sourceLine}$3`,
    );
  }
  return h;
}

function main() {
  const { positional, flags } = parseArgs(process.argv);
  if (!positional.length) {
    console.error('Usage: node build-builtin.mjs <file.xlsx> [--sheet NAME] [--write] ...');
    process.exit(1);
  }
  const xlsxPath = path.resolve(positional[0]);
  if (!fs.existsSync(xlsxPath)) {
    console.error(`File not found: ${xlsxPath}`);
    process.exit(1);
  }
  const htmlPath = path.resolve(flags.html || path.join(__dirname, 'index.html'));
  const { html, attendees, nameAliases, regionMap } = loadContextFromHtml(htmlPath);

  const wb = XLSX.read(fs.readFileSync(xlsxPath), { type: 'buffer' });
  const sheetName = pickSheet(wb, flags.sheet);
  if (!sheetName) throw new Error('Could not pick a sheet.');
  const sheet = wb.Sheets[sheetName];
  const header = (XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0 })[0] || []);
  if (!header.includes('NAME-ROLE-COMPANY')) {
    throw new Error(`Sheet "${sheetName}" has no NAME-ROLE-COMPANY column.`);
  }
  const rows = XLSX.utils.sheet_to_json(sheet);
  const { perPersonMeetings, builtinData } = processRows(rows, attendees, nameAliases, regionMap);

  const perPersonLit = JSON.stringify(perPersonMeetings);
  const dataLit = JSON.stringify(builtinData);
  const baseFile = path.basename(xlsxPath);
  const bannerDate = flags.bannerDate || '14 Apr 2026';
  const bannerFile = flags.bannerFile || `${baseFile} · sheet ${sheetName}`;
  const sourceLine = `Showing built-in data (${baseFile}, sheet ${sheetName}) — upload a new file above to refresh`;

  console.error(`Sheet: ${sheetName} (${rows.length} rows)`);
  console.error(`Sizes: BUILTIN_PER_PERSON ${perPersonLit.length} chars, BUILTIN_DATA ${dataLit.length} chars`);

  if (!flags.write) {
    console.error('Dry run (no --write).');
    process.exit(0);
  }

  let out = replaceConst(html, 'BUILTIN_PER_PERSON', perPersonLit);
  out = replaceConst(out, 'BUILTIN_DATA', dataLit);
  out = maybePatchBanner(out, bannerDate, bannerFile, sourceLine);
  fs.writeFileSync(htmlPath, out, 'utf8');
  console.error(`Wrote ${htmlPath}`);
}

main();
