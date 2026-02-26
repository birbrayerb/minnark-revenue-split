import * as XLSX from 'xlsx';

/**
 * MinnARK Revenue Split Pipeline v3
 * Reads Invoice tabs directly, looks up POs in Program_Info, classifies DO from Domestic_Payments.
 */

// ── Helpers ──

function parseDate(val) {
  if (val == null) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s) return null;
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
    // Try "DD Mon YYYY"
    const m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
    if (m) {
      const d2 = new Date(`${m[2]} ${m[1]}, ${m[3]}`);
      if (!isNaN(d2.getTime())) return d2;
    }
  }
  return null;
}

function normalizePO(po) {
  if (po == null) return '';
  let s = String(po).trim();
  // Remove .0 from numeric strings
  if (s.endsWith('.0')) s = s.slice(0, -2);
  const stripped = s.replace(/^0+/, '');
  return stripped || s;
}

function safeFloat(v) {
  if (v == null) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function getMonthKey(date) {
  return date.getFullYear() * 100 + (date.getMonth() + 1);
}

function getMonthName(ym) {
  const months = ['', 'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[ym % 100]} ${Math.floor(ym / 100)}`;
}

// ── File reading ──

export function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        // Cap oversized sheets to 50 columns
        for (const name of workbook.SheetNames) {
          const ws = workbook.Sheets[name];
          if (ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            if (range.e.c > 50) {
              range.e.c = 50;
              ws['!ref'] = XLSX.utils.encode_range(range);
            }
          }
        }
        resolve(workbook);
      } catch (err) {
        reject(new Error(`Failed to parse ${file.name}: ${err.message}`));
      }
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Detect file type from sheet names.
 */
export function detectFileType(workbook) {
  const names = workbook.SheetNames.map(n => n.trim());
  if (names.some(n => n === 'TOTAL DI' || n === 'SUMMARY')) return 'jpm';
  if (names.some(n => n.includes('Payment Details'))) return 'domestic';
  if (names.some(n => n === 'New Program Info')) return 'program_blackfin';
  if (names.some(n => n.includes('Mizar PO'))) return 'program_mizar';
  if (names.some(n => n === 'PO Info')) return 'domestic_pos';
  return null;
}

export function fileTypeLabel(type) {
  const labels = {
    jpm: '📊 JPM Monthly',
    domestic: '📋 Domestic Payments',
    program_blackfin: '📘 Program Info (Blackfin)',
    program_mizar: '📗 Program Info (Mizar)',
    domestic_pos: '📦 Domestic POs',
  };
  return labels[type] || '❓ Unknown';
}

// ── Program Info Loading ──

function loadBlackfinPO(workbook) {
  const ws = workbook.Sheets['New Program Info'];
  if (!ws) return {};
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const lookup = {}; // normPO -> [{team, program, item, cost_total}]
  // Row 1 = headers, row 2+ = data
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row[12] == null) continue;
    const po = normalizePO(row[12]);
    if (!po) continue;
    const team = row[5] ? String(row[5]).trim() : 'Blackfin';
    const program = row[6] ? String(row[6]).trim() : '';
    const item = row[4] ? String(row[4]).trim() : '';
    const costTotal = safeFloat(row[9]);
    if (!lookup[po]) lookup[po] = [];
    lookup[po].push({ team, program, item, cost_total: costTotal });
  }
  return lookup;
}

function loadMizarPO(workbook) {
  // Sheet "Mizar PO 2025", row 4 = headers, row 5+ = data
  const sheetName = workbook.SheetNames.find(n => n.includes('Mizar PO'));
  if (!sheetName) return {};
  const ws = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const lookup = {};
  for (let i = 4; i < data.length; i++) {
    const row = data[i];
    if (!row || row[12] == null) continue;
    const poRaw = String(row[12]).trim();
    if (!poRaw || poRaw.toLowerCase() === 'total' || poRaw.toLowerCase() === 'subtotal' || poRaw.startsWith('=')) continue;
    const po = normalizePO(row[12]);
    if (!po) continue;
    const team = row[5] ? String(row[5]).trim() : 'Mizar';
    const program = row[6] ? String(row[6]).trim() : '';
    const item = row[4] ? String(row[4]).trim() : '';
    const costTotal = safeFloat(row[9]);
    if (!lookup[po]) lookup[po] = [];
    lookup[po].push({ team, program, item, cost_total: costTotal });
  }
  return lookup;
}

function loadDomesticPOs(workbook) {
  const ws = workbook.Sheets['PO Info'];
  if (!ws) return new Set();
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const poSet = new Set();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row[1] == null) continue;
    poSet.add(normalizePO(row[1]));
  }
  return poSet;
}

// ── Invoice Tab Reading ──

function readInvoiceTabs(workbook, filename) {
  const tabs = workbook.SheetNames.filter(n => n.startsWith('Invoice'));
  const rows = [];
  for (const tab of tabs) {
    const ws = workbook.Sheets[tab];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    for (let i = 1; i < data.length; i++) {
      const rv = data[i];
      if (!rv) continue;
      const invoiceRef = rv[0] != null ? String(rv[0]).trim() : '';
      const date = parseDate(rv[2]);
      const netAmount = safeFloat(rv[7]);
      const po = rv[14] != null ? String(rv[14]).trim() : '';
      if (!invoiceRef && !netAmount) continue;
      if (netAmount === 0) continue;
      rows.push({ invoiceRef, date, netAmount, po, sourceFile: filename, tab });
    }
  }
  return rows;
}

// ── Domestic Payments Reading ──

function readDomesticPayments(workbook) {
  const sheetName = workbook.SheetNames.find(n => n.includes('2025_Payment Details'))
    || workbook.SheetNames.find(n => n.includes('Payment Details'));
  if (!sheetName) return [];

  const ws = workbook.Sheets[sheetName];
  // Cap columns
  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    range.e.c = Math.min(range.e.c, 50);
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });

  const rows = [];
  let lastDate = null;
  for (let i = 1; i < data.length; i++) {
    const rv = data[i];
    if (!rv) continue;
    const pd = parseDate(rv[2]);
    if (pd) lastDate = pd;
    const team = rv[27] ? String(rv[27]).trim() : null;
    const program = rv[33] ? String(rv[33]).trim() : null;
    const paid = rv[32];
    if (!lastDate || !team || paid == null) continue;
    const amount = Number(paid);
    if (!isNaN(amount) && amount !== 0) {
      rows.push({ date: lastDate, team, program: program || 'UNKNOWN', amount });
    }
  }
  return rows;
}

// ── Proportional Split ──

function proportionalSplit(amount, items, splitType) {
  const totalCost = items.reduce((s, i) => s + i.cost_total, 0);
  if (totalCost === 0 || items.length === 0) {
    return items.length > 0
      ? items.map(i => ({ type: splitType, team: i.team, program: i.program || i.item || 'Unknown', amount: amount / items.length }))
      : [{ type: splitType, team: 'Unknown', program: 'Unknown', amount }];
  }
  return items.map(i => ({
    type: splitType,
    team: i.team,
    program: i.program || i.item || 'Unknown',
    amount: amount * (i.cost_total / totalCost),
  }));
}

// ── Main Processing ──

/**
 * Process all uploaded files.
 * @returns {{ monthlyResults: Array, unmatchedRecords: Array }}
 */
export function processFiles(jpmFiles, domesticFile, blackfinFile, mizarFile, domesticPOsFile) {
  // Load PO lookups
  const diLookup = {};
  if (blackfinFile) {
    const bf = loadBlackfinPO(blackfinFile.workbook);
    Object.assign(diLookup, bf);
  }
  if (mizarFile) {
    const mz = loadMizarPO(mizarFile.workbook);
    // Merge — don't overwrite, append
    for (const [po, items] of Object.entries(mz)) {
      if (diLookup[po]) diLookup[po].push(...items);
      else diLookup[po] = items;
    }
  }

  // Load domestic PO set
  const domesticPOSet = domesticPOsFile ? loadDomesticPOs(domesticPOsFile.workbook) : new Set();

  // Load DO rows
  const allDoRows = domesticFile ? readDomesticPayments(domesticFile.workbook) : [];
  const doByMonth = {};
  for (const r of allDoRows) {
    const ym = getMonthKey(r.date);
    if (!doByMonth[ym]) doByMonth[ym] = [];
    doByMonth[ym].push(r);
  }

  // Read all invoice rows from all JPM files
  const allInvoiceRows = [];
  for (const { workbook, filename } of jpmFiles) {
    const rows = readInvoiceTabs(workbook, filename);
    allInvoiceRows.push(...rows);
  }

  // Classify each invoice row
  const classified = []; // {monthKey, type, team, program, amount}
  const unmatched = [];

  for (const row of allInvoiceRows) {
    const normPO = normalizePO(row.po);
    const monthKey = row.date ? getMonthKey(row.date) : 0;

    // Check DI lookup
    if (normPO && diLookup[normPO]) {
      const splits = proportionalSplit(row.netAmount, diLookup[normPO], 'DI');
      for (const s of splits) {
        classified.push({ monthKey, ...s });
      }
      continue;
    }

    // Check if DO: strip last 2 chars from invoice ref → domestic PO
    const invoiceRef = row.invoiceRef;
    if (invoiceRef.length >= 3) {
      const domesticPO = normalizePO(invoiceRef.slice(0, -2));
      if (domesticPO && domesticPOSet.has(domesticPO)) {
        // It's DO — handled by Domestic_Payments aggregation
        continue;
      }
    }

    // Check if starts with "700" — likely fuel passthrough
    // Still add as unmatched so user can exclude
    unmatched.push({
      id: unmatched.length,
      invoiceRef: row.invoiceRef,
      po: row.po,
      date: row.date,
      netAmount: row.netAmount,
      sourceFile: row.sourceFile,
      tab: row.tab,
      monthKey,
      // Default assignment
      assignedTeam: '',
      assignedProgram: '',
      classification: normPO.startsWith('700') ? 'Exclude' : '',
    });
  }

  // Build monthly results
  // Collect unique months from classified + DO
  const allMonths = new Set();
  for (const c of classified) allMonths.add(c.monthKey);
  for (const ym of Object.keys(doByMonth)) allMonths.add(Number(ym));

  const monthlyResults = [];
  for (const ym of [...allMonths].sort()) {
    if (ym === 0) continue;
    const diRows = classified.filter(c => c.monthKey === ym && c.type === 'DI');
    const doRows = doByMonth[ym] || [];

    const diByTeamProg = {};
    for (const r of diRows) {
      const key = `${r.team}|||${r.program}`;
      diByTeamProg[key] = (diByTeamProg[key] || 0) + r.amount;
    }

    const doByTeamProg = {};
    for (const r of doRows) {
      const key = `${r.team}|||${r.program}`;
      doByTeamProg[key] = (doByTeamProg[key] || 0) + r.amount;
    }

    const diBlackfin = {}, diMizar = {};
    for (const [key, amt] of Object.entries(diByTeamProg)) {
      const [team, prog] = key.split('|||');
      const t = team.trim().toLowerCase();
      if (t.includes('blackfin') || t === 'bf') diBlackfin[prog] = (diBlackfin[prog] || 0) + amt;
      else if (t.includes('mizar') || t === 'mz') diMizar[prog] = (diMizar[prog] || 0) + amt;
      else diBlackfin[prog] = (diBlackfin[prog] || 0) + amt; // default
    }

    const doBlackfin = {}, doMizar = {};
    for (const [key, amt] of Object.entries(doByTeamProg)) {
      const [team, prog] = key.split('|||');
      const t = team.trim().toLowerCase();
      if (t.includes('blackfin') || t === 'bf') doBlackfin[prog] = (doBlackfin[prog] || 0) + amt;
      else if (t.includes('mizar') || t === 'mz') doMizar[prog] = (doMizar[prog] || 0) + amt;
      else doBlackfin[prog] = (doBlackfin[prog] || 0) + amt;
    }

    const sum = obj => Object.values(obj).reduce((a, b) => a + b, 0);
    const round2 = obj => {
      const r = {};
      for (const [k, v] of Object.entries(obj)) r[k] = Math.round(v * 100) / 100;
      return r;
    };

    const actual = {
      di_blackfin: sum(diBlackfin),
      di_mizar: sum(diMizar),
      do_blackfin: sum(doBlackfin),
      do_mizar: sum(doMizar),
    };
    actual.di_total = actual.di_blackfin + actual.di_mizar;
    actual.do_total = actual.do_blackfin + actual.do_mizar;
    actual.grand_total = actual.di_total + actual.do_total;

    monthlyResults.push({
      month: getMonthName(ym),
      monthKey: ym,
      actual,
      di_blackfin_programs: round2(diBlackfin),
      di_mizar_programs: round2(diMizar),
      do_blackfin_programs: round2(doBlackfin),
      do_mizar_programs: round2(doMizar),
      di_row_count: diRows.length,
      do_row_count: doRows.length,
    });
  }

  return { monthlyResults, unmatchedRecords: unmatched };
}

/**
 * Apply manual unmatched assignments to monthly results.
 * Returns updated monthlyResults with manual assignments folded in.
 */
export function applyManualAssignments(baseResults, unmatchedRecords) {
  // Clone base results
  const results = baseResults.map(r => ({
    ...r,
    actual: { ...r.actual },
    di_blackfin_programs: { ...r.di_blackfin_programs },
    di_mizar_programs: { ...r.di_mizar_programs },
    do_blackfin_programs: { ...r.do_blackfin_programs },
    do_mizar_programs: { ...r.do_mizar_programs },
  }));

  // Create month lookup
  const monthMap = {};
  for (const r of results) monthMap[r.monthKey] = r;

  for (const rec of unmatchedRecords) {
    if (!rec.assignedTeam || !rec.classification || rec.classification === 'Exclude') continue;
    const ym = rec.monthKey;
    if (!monthMap[ym]) {
      // Create month entry
      const entry = {
        month: getMonthName(ym),
        monthKey: ym,
        actual: { di_blackfin: 0, di_mizar: 0, do_blackfin: 0, do_mizar: 0, di_total: 0, do_total: 0, grand_total: 0 },
        di_blackfin_programs: {}, di_mizar_programs: {},
        do_blackfin_programs: {}, do_mizar_programs: {},
        di_row_count: 0, do_row_count: 0,
      };
      results.push(entry);
      monthMap[ym] = entry;
    }
    const mr = monthMap[ym];
    const team = rec.assignedTeam.toLowerCase();
    const cat = rec.classification; // DI or DO
    const prog = rec.assignedProgram || 'Manual';
    const amt = rec.netAmount;

    const isBF = team.includes('blackfin');
    const progKey = cat === 'DI'
      ? (isBF ? 'di_blackfin_programs' : 'di_mizar_programs')
      : (isBF ? 'do_blackfin_programs' : 'do_mizar_programs');
    const actKey = cat === 'DI'
      ? (isBF ? 'di_blackfin' : 'di_mizar')
      : (isBF ? 'do_blackfin' : 'do_mizar');

    mr[progKey][prog] = (mr[progKey][prog] || 0) + amt;
    mr.actual[actKey] += amt;
    mr.actual[cat === 'DI' ? 'di_total' : 'do_total'] += amt;
    mr.actual.grand_total += amt;
  }

  results.sort((a, b) => a.monthKey - b.monthKey);
  return results;
}

// ── Export ──

export function exportCSV(results) {
  const lines = ['Month,Category,Team,Program,Amount'];
  for (const r of results) {
    const add = (cat, team, progs) => {
      for (const [prog, amt] of Object.entries(progs)) {
        lines.push(`${r.month},${cat},${team},"${prog}",${amt.toFixed(2)}`);
      }
    };
    add('DI', 'Blackfin', r.di_blackfin_programs);
    add('DI', 'Mizar', r.di_mizar_programs);
    add('DO', 'Blackfin', r.do_blackfin_programs);
    add('DO', 'Mizar', r.do_mizar_programs);
  }
  return lines.join('\n');
}

export function exportExcel(results) {
  const wb = XLSX.utils.book_new();
  const summaryData = [['Month', '', 'Blackfin DI', 'Mizar DI', 'Blackfin DO', 'Mizar DO', 'Grand Total']];
  for (const r of results) {
    summaryData.push([r.month, '', r.actual.di_blackfin, r.actual.di_mizar, r.actual.do_blackfin, r.actual.do_mizar, r.actual.grand_total]);
  }
  if (results.length > 1) {
    const totals = ['TOTAL', ''];
    for (const key of ['di_blackfin', 'di_mizar', 'do_blackfin', 'do_mizar', 'grand_total']) {
      totals.push(results.reduce((s, r) => s + r.actual[key], 0));
    }
    summaryData.push(totals);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), 'Summary');

  const detailData = [['Month', 'Category', 'Team', 'Program', 'Amount']];
  for (const r of results) {
    const add = (cat, team, progs) => {
      for (const [prog, amt] of Object.entries(progs)) detailData.push([r.month, cat, team, prog, amt]);
    };
    add('DI', 'Blackfin', r.di_blackfin_programs);
    add('DI', 'Mizar', r.di_mizar_programs);
    add('DO', 'Blackfin', r.do_blackfin_programs);
    add('DO', 'Mizar', r.do_mizar_programs);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailData), 'Detail');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}
