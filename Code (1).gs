/**
 * MULTI-ENTITY CLOSE ASSISTANT
 * Google Sheets + Apps Script + Claude API
 *
 * What it does, in one click ("Close Tools > Run full close"):
 *   A. Matches the bank feed to what is already recorded in the books (QBO register)
 *   B. Categorizes unbooked bank activity: rules first, then Claude for the rest
 *   C. Runs exception checks: duplicates, transpositions, unmatched intercompany,
 *      large items, low-confidence AI picks
 *   D. Builds a bank reconciliation per entity that must tie to the penny
 *   E. Produces a QuickBooks Online journal entry import file for clean items only
 *   F. Writes a summary + review queue, and logs every run (audit trail)
 *
 * Approved AI picks can be promoted into Rules, so next month they cost nothing.
 */

// ===================== CONFIG =====================
const S = {
  SUMMARY: 'Summary', BANK: 'Bank_Txns', BOOK: 'Book_Register', REC: 'Bank_Rec',
  EXPORT: 'QBO_JE_Import', ENT: 'Entities', COA: 'COA', RULES: 'Rules',
  SETTINGS: 'Settings', LOG: 'Run_Log'
};
const IC_TOKEN = 'AUTO-INTERCOMPANY';
const DUE_FROM = '1800 Due From Affiliates';
const DUE_TO = '2800 Due To Affiliates';
const PERIOD_LABEL = 'September 2026';

const BANK_HEADERS = ['Txn ID', 'Date', 'Entity', 'Description', 'Amount', 'Status', 'Account',
  'Method', 'Confidence', 'Notes / AI Reason', 'Flags', 'Approve', 'Matched Book ID'];
const B = { ID: 0, DATE: 1, ENT: 2, DESC: 3, AMT: 4, STATUS: 5, ACCT: 6, METHOD: 7, CONF: 8,
  NOTES: 9, FLAGS: 10, APPROVE: 11, MATCH: 12 };
const BOOK_HEADERS = ['Book ID', 'Date', 'Entity', 'Description', 'Amount', 'Matched Bank ID'];
const K = { ID: 0, DATE: 1, ENT: 2, DESC: 3, AMT: 4, MATCH: 5 };
const ENT_HEADERS = ['Code', 'Legal Name', 'Type', 'Intercompany Keyword', 'QBO Cash Account',
  'Opening Balance (9/1)', 'Bank Statement Ending Balance (9/30)'];
const RULE_HEADERS = ['Description contains', 'Account', 'Note', 'Hold for review (Y/N)'];

// Flags that stop an item from going into the import file until someone reviews it
const HOLD_FLAGS = ['DUP', 'TRANSPOSITION', 'IC-UNMATCHED', 'NO-CATEGORY', 'AI-LOW', 'RULE-HOLD'];

const MONEY_FMT = '$#,##0.00;[Red]-$#,##0.00';

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Close Tools')
    .addItem('1. Set up / reset demo data', 'setupDemo')
    .addItem('2. Save Claude API key', 'setApiKey')
    .addSeparator()
    .addItem('▶ Run full close', 'runFullClose')
    .addSeparator()
    .addItem('Save approved AI picks as rules', 'promoteApprovedToRules')
    .addItem('Clear AI cache', 'clearAiCache')
    .addSeparator().addSubMenu(qboMenu_())
    .addToUi();
}

// ===================== SMALL HELPERS =====================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }
function getOrCreate_(name) { return sheet_(name) || ss_().insertSheet(name); }
function cents_(x) { return Math.round(Number(x || 0) * 100); }
function fromCents_(c) { return c / 100; }
function daysApart_(a, b) { return Math.abs(new Date(a).getTime() - new Date(b).getTime()) / 86400000; }
function norm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ' ').trim(); }
function money_(x) {
  const n = Number(x) || 0;
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-$' : '$') + s;
}
function fmtDate_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MM/dd/yyyy');
}
function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function safeAlert_(title, msg) {
  try { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(title + ': ' + msg); }
}
function toast_(msg) { try { ss_().toast(msg, 'Close Tools', 5); } catch (e) { /* ignore */ } }

function readRows_(name, ncols) {
  const sh = sheet_(name);
  if (!sh) throw new Error('Missing tab "' + name + '". Run "Set up / reset demo data" first.');
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ncols).getValues().filter(r => r[0] !== '' && r[0] !== null);
}
function writeTable_(name, headers, rows) {
  const sh = getOrCreate_(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
  return sh;
}
function addFlag_(row, code, msg) {
  const f = code + ': ' + msg;
  row[B.FLAGS] = row[B.FLAGS] ? row[B.FLAGS] + ' | ' + f : f;
}
function hasHoldFlag_(row) {
  const f = String(row[B.FLAGS] || '');
  return HOLD_FLAGS.some(code => f.indexOf(code + ':') !== -1);
}
function log_(action, details) {
  const sh = getOrCreate_(S.LOG);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Action', 'Details']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([now_(), action, details]);
}
function readSettings_() {
  const out = { threshold: 25000, windowDays: 3, model: 'claude-sonnet-5-5', minConf: 0.8, batch: 20 };
  const sh = sheet_(S.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, v]) => {
    const key = norm_(k);
    if (key.indexOf('LARGE ITEM') === 0) out.threshold = Number(v);
    else if (key.indexOf('MATCH WINDOW') === 0) out.windowDays = Number(v);
    else if (key.indexOf('AI MODEL') === 0) out.model = String(v).trim();
    else if (key.indexOf('AI CONFIDENCE') === 0) out.minConf = Number(v);
    else if (key.indexOf('AI BATCH') === 0) out.batch = Number(v);
  });
  return out;
}
function readEntities_() {
  return readRows_(S.ENT, ENT_HEADERS.length).map(r => ({
    code: String(r[0]), name: String(r[1]), type: String(r[2]), keyword: norm_(r[3]),
    cash: String(r[4]), opening: Number(r[5]), stmtEnd: Number(r[6])
  }));
}
function readCoa_() {
  return readRows_(S.COA, 3).map(r => String(r[0]) + ' ' + String(r[1]));
}

// ===================== API KEY & CACHE =====================
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Claude API key',
    'Paste your key from console.anthropic.com. It is stored privately in your Google account (Script Properties), never in the sheet.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim()) {
    PropertiesService.getUserProperties().setProperty('CLAUDE_API_KEY', r.getResponseText().trim());
    ui.alert('Saved. You can now run the close.');
  }
}
function clearAiCache() {
  const props = PropertiesService.getDocumentProperties();
  Object.keys(props.getProperties()).forEach(k => { if (k.indexOf('AI|') === 0) props.deleteProperty(k); });
  toast_('AI cache cleared.');
}

// ===================== MAIN =====================
function runFullClose() {
  const cfg = readSettings_();
  const entities = readEntities_();
  const coa = readCoa_();
  const bank = readRows_(S.BANK, BANK_HEADERS.length);
  const book = readRows_(S.BOOK, BOOK_HEADERS.length);
  const rules = readRows_(S.RULES, RULE_HEADERS.length);

  // Reset working columns. Rows you approved keep their account (they are "locked").
  bank.forEach(r => {
    const locked = r[B.APPROVE] === true && r[B.ACCT];
    r[B.STATUS] = ''; r[B.FLAGS] = ''; r[B.MATCH] = '';
    if (!locked) { r[B.ACCT] = ''; r[B.METHOD] = ''; r[B.CONF] = ''; r[B.NOTES] = ''; }
    else if (String(r[B.METHOD]).indexOf('approved') === -1) r[B.METHOD] = r[B.METHOD] + ' (approved)';
  });
  book.forEach(k => { k[K.MATCH] = ''; });

  toast_('Step A: matching bank to books…');
  stepMatch_(bank, book, cfg);
  toast_('Step B: categorizing…');
  const aiStats = stepCategorize_(bank, rules, entities, coa, cfg);
  toast_('Step C: exception checks…');
  stepFlags_(bank, book, entities, cfg);
  toast_('Step D: bank reconciliation…');
  const recs = stepReconcile_(bank, book, entities);
  toast_('Step E: building QBO import…');
  const exported = stepExport_(bank, entities, coa);

  // Write results back
  writeTable_(S.BANK, BANK_HEADERS, bank);
  writeTable_(S.BOOK, BOOK_HEADERS, book);
  formatBank_(bank.length);
  formatBook_(book.length);

  stepSummary_(bank, entities, recs, aiStats);
  log_('Run full close', bank.length + ' bank lines; ' + exported + ' ready to import; ' +
    bank.filter(r => r[B.STATUS] === 'Exception').length + ' held; AI calls: ' + aiStats.calls +
    (aiStats.error ? '; AI error: ' + aiStats.error : ''));
  try { ss_().setActiveSheet(sheet_(S.SUMMARY)); } catch (e) { /* ignore */ }
  if (aiStats.error) safeAlert_('Finished, but the AI step had a problem', aiStats.error +
    '\n\nRule-based steps still ran. Check your API key (Close Tools > Save Claude API key) and credits.');
  else toast_('Close complete. See Summary.');
}

// ===================== STEP A: MATCH =====================
function stepMatch_(bank, book, cfg) {
  const used = {};
  bank.forEach(r => {
    for (let i = 0; i < book.length; i++) {
      const k = book[i];
      if (used[k[K.ID]]) continue;
      if (k[K.ENT] !== r[B.ENT]) continue;
      if (cents_(k[K.AMT]) !== cents_(r[B.AMT])) continue;
      if (daysApart_(k[K.DATE], r[B.DATE]) > cfg.windowDays) continue;
      used[k[K.ID]] = true;
      k[K.MATCH] = r[B.ID];
      r[B.MATCH] = k[K.ID];
      r[B.STATUS] = 'Booked';
      if (!(r[B.APPROVE] === true && r[B.ACCT])) { r[B.METHOD] = 'Matched to books'; r[B.NOTES] = 'Already in QBO as ' + k[K.ID]; }
      break;
    }
    if (!r[B.MATCH]) r[B.STATUS] = 'To Book';
  });
}

// ===================== STEP B: CATEGORIZE =====================
function stepCategorize_(bank, rules, entities, coa, cfg) {
  const stats = { rule: 0, ai: 0, calls: 0, error: '' };
  const todo = bank.filter(r => r[B.STATUS] === 'To Book' && !(r[B.APPROVE] === true && r[B.ACCT]));

  // Rules first
  const needAi = [];
  todo.forEach(r => {
    const desc = norm_(r[B.DESC]);
    const rule = rules.find(x => x[0] && desc.indexOf(norm_(x[0])) !== -1);
    if (!rule) { needAi.push(r); return; }
    let acct = String(rule[1]);
    if (acct === IC_TOKEN) acct = Number(r[B.AMT]) < 0 ? DUE_FROM : DUE_TO;
    r[B.ACCT] = acct;
    r[B.METHOD] = 'Rule';
    r[B.CONF] = 1;
    r[B.NOTES] = 'Rule: "' + rule[0] + '"' + (rule[2] ? ' – ' + rule[2] : '');
    if (norm_(rule[3]) === 'Y') addFlag_(r, 'RULE-HOLD', rule[2] || 'Rule requires review');
    stats.rule++;
  });
  if (!needAi.length) return stats;

  // Then Claude, with a cache so reruns are consistent and free
  const props = PropertiesService.getDocumentProperties();
  const key = PropertiesService.getUserProperties().getProperty('CLAUDE_API_KEY');
  const cacheKey = r => 'AI|' + r[B.ENT] + '|' + norm_(r[B.DESC]) + '|' + cents_(r[B.AMT]);
  const uncached = [];
  needAi.forEach(r => {
    const hit = props.getProperty(cacheKey(r));
    if (hit) applyAi_(r, JSON.parse(hit), coa, cfg, true);
    else uncached.push(r);
  });

  if (uncached.length && !key) {
    stats.error = 'No Claude API key saved.';
    uncached.forEach(r => { r[B.METHOD] = 'Needs review'; addFlag_(r, 'NO-CATEGORY', 'No rule matched and AI is not set up'); });
    return finishAiStats_(needAi, stats);
  }

  const entMap = {};
  entities.forEach(e => { entMap[e.code] = e; });
  for (let i = 0; i < uncached.length; i += cfg.batch) {
    const chunk = uncached.slice(i, i + cfg.batch);
    try {
      const results = askClaude_(chunk, entMap, coa, cfg, key);
      stats.calls++;
      chunk.forEach(r => {
        const res = results[String(r[B.ID])];
        if (!res) { r[B.METHOD] = 'Needs review'; addFlag_(r, 'NO-CATEGORY', 'AI returned no answer for this line'); return; }
        props.setProperty(cacheKey(r), JSON.stringify(res));
        applyAi_(r, res, coa, cfg, false);
      });
    } catch (e) {
      stats.error = String(e.message || e);
      chunk.forEach(r => { r[B.METHOD] = 'Needs review'; addFlag_(r, 'NO-CATEGORY', 'AI call failed'); });
    }
  }
  return finishAiStats_(needAi, stats);
}
function finishAiStats_(needAi, stats) {
  stats.ai = needAi.filter(r => String(r[B.METHOD]).indexOf('AI') === 0).length;
  return stats;
}
function applyAi_(r, res, coa, cfg, fromCache) {
  const acct = String(res.account || '').trim();
  const conf = Number(res.confidence) || 0;
  r[B.METHOD] = 'AI';
  r[B.CONF] = Math.round(conf * 100) / 100;
  r[B.NOTES] = (res.reason || '') + (fromCache ? ' (cached)' : '');
  if (acct === IC_TOKEN || acct === DUE_FROM || acct === DUE_TO) {
    r[B.ACCT] = Number(r[B.AMT]) < 0 ? DUE_FROM : DUE_TO;
  } else if (coa.indexOf(acct) !== -1) {
    r[B.ACCT] = acct;
  } else {
    r[B.ACCT] = '';
    addFlag_(r, 'NO-CATEGORY', 'AI was unsure – needs a human decision');
    return;
  }
  if (conf < cfg.minConf) addFlag_(r, 'AI-LOW', 'Confidence ' + Math.round(conf * 100) + '% – review, then tick Approve');
  if (res.flag) addFlag_(r, 'AI-NOTE', String(res.flag));
}
function askClaude_(rows, entMap, coa, cfg, key) {
  const system = [
    'You are a meticulous staff accountant at a single-family office.',
    'Categorize each bank transaction to exactly one account from the chart of accounts provided.',
    'Rules:',
    '- Use an account EXACTLY as written in the list, or the word UNSURE.',
    '- Consider the entity type (holding LLC, trust, real-estate LLC).',
    '- Separate capital items from expenses (e.g. a roof replacement is capitalizable).',
    '- For private fund distributions, consider return of capital vs. income; lower confidence if it cannot be told from the description.',
    '- Note possible personal expenses of the principal, or anything a CFO should see.',
    '- confidence is 0.0 to 1.0: how sure a reviewer would be without seeing more documents.',
    'Respond with ONLY a JSON array, no other text:',
    '[{"id":"<txn id>","account":"<account or UNSURE>","confidence":0.0,"reason":"<20 words max>","flag":"<short review note, or empty string>"}]'
  ].join('\n');
  const ents = Object.keys(entMap).map(c => c + ' = ' + entMap[c].name + ' (' + entMap[c].type + ')').join('\n');
  const txns = rows.map(r => ({
    id: String(r[B.ID]), entity: r[B.ENT], date: fmtDate_(r[B.DATE]),
    description: String(r[B.DESC]), amount: Number(r[B.AMT])
  }));
  const user = 'Entities:\n' + ents + '\n\nChart of accounts:\n' + coa.join('\n') +
    '\n\nTransactions (positive = money in, negative = money out):\n' + JSON.stringify(txns, null, 1);

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: cfg.model, max_tokens: 4000, system: system,
      messages: [{ role: 'user', content: user }] })
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) throw new Error('Claude API error ' + code + ': ' + body.slice(0, 300));
  const data = JSON.parse(body);
  const text = (data.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  const m = text.match(/\[[\s\S]*\]/);
  if (!m) throw new Error('Could not read the AI response: ' + text.slice(0, 200));
  const out = {};
  JSON.parse(m[0]).forEach(x => { out[String(x.id)] = x; });
  return out;
}

// ===================== STEP C: EXCEPTION CHECKS =====================
function stepFlags_(bank, book, entities, cfg) {
  // 1. Possible duplicates: same entity, date, amount and description
  const seen = {};
  bank.forEach(r => {
    const k = [r[B.ENT], fmtDate_(r[B.DATE]), cents_(r[B.AMT]), norm_(r[B.DESC])].join('|');
    if (seen[k]) addFlag_(r, 'DUP', 'Possible duplicate of ' + seen[k]);
    else seen[k] = r[B.ID];
  });

  // 2. Transposition: unbooked bank line vs. unmatched book line, difference divisible by 9
  const openBook = book.filter(k => !k[K.MATCH]);
  bank.filter(r => r[B.STATUS] === 'To Book').forEach(r => {
    openBook.forEach(k => {
      if (k[K.ENT] !== r[B.ENT]) return;
      if (daysApart_(k[K.DATE], r[B.DATE]) > cfg.windowDays) return;
      const a = cents_(r[B.AMT]), b = cents_(k[K.AMT]);
      const diff = Math.abs(a - b);
      if (diff === 0 || diff % 9 !== 0 || (a < 0) !== (b < 0)) return;
      if (diff / Math.abs(a) > 0.05) return;
      addFlag_(r, 'TRANSPOSITION', 'Books show ' + money_(k[K.AMT]) + ' as ' + k[K.ID] +
        ' (diff ' + money_(fromCents_(diff)) + ', divisible by 9). Fix with a correcting entry, do not re-book.');
    });
  });

  // 3. Intercompany: every transfer must have the mirror entry in the other entity's bank
  bank.forEach(r => {
    if (r[B.ACCT] !== DUE_FROM && r[B.ACCT] !== DUE_TO) return;
    const desc = norm_(r[B.DESC]);
    const cp = entities.find(e => e.code !== r[B.ENT] && e.keyword && desc.indexOf(e.keyword) !== -1);
    if (!cp) { addFlag_(r, 'IC-UNMATCHED', 'Cannot tell which affiliate this transfer is with'); return; }
    r[B.NOTES] = (r[B.NOTES] ? r[B.NOTES] + ' | ' : '') + 'Counterparty: ' + cp.code;
    const mirror = bank.find(x => x[B.ENT] === cp.code && cents_(x[B.AMT]) === -cents_(r[B.AMT]) &&
      daysApart_(x[B.DATE], r[B.DATE]) <= cfg.windowDays);
    if (!mirror) addFlag_(r, 'IC-UNMATCHED', 'No matching ' + money_(-Number(r[B.AMT])) + ' in ' + cp.code + "'s bank within " + cfg.windowDays + ' days');
  });

  // 4. Large items (informational)
  bank.filter(r => r[B.STATUS] === 'To Book').forEach(r => {
    if (Math.abs(Number(r[B.AMT])) >= cfg.threshold) addFlag_(r, 'LARGE', 'At or above ' + money_(cfg.threshold) + ' – CFO visibility');
  });

  // Anything with a hold flag is an exception
  bank.forEach(r => { if (r[B.STATUS] === 'To Book' && hasHoldFlag_(r)) r[B.STATUS] = 'Exception'; });
}

// ===================== STEP D: BANK RECONCILIATION =====================
function stepReconcile_(bank, book, entities) {
  const rows = [];
  const bold = [];
  const results = {};
  entities.forEach(e => {
    const bookOnly = book.filter(k => k[K.ENT] === e.code && !k[K.MATCH]);
    const bankOnly = bank.filter(r => r[B.ENT] === e.code && !r[B.MATCH]);
    const dit = bookOnly.filter(k => Number(k[K.AMT]) > 0);
    const oc = bookOnly.filter(k => Number(k[K.AMT]) < 0);
    const sum = arr => arr.reduce((t, x) => t + cents_(x[4]), 0); // col 4 = amount in both tabs
    const stmt = cents_(e.stmtEnd);
    const adjBank = stmt + sum(dit) + sum(oc);
    const bookBal = cents_(e.opening) + book.filter(k => k[K.ENT] === e.code).reduce((t, k) => t + cents_(k[K.AMT]), 0);
    const adjBook = bookBal + sum(bankOnly);
    const diff = adjBank - adjBook;
    results[e.code] = diff;

    bold.push(rows.length + 2);
    rows.push([e.code + ' – ' + e.name + '   (' + e.cash + ')', '', '']);
    rows.push(['Bank statement ending balance', fromCents_(stmt), '']);
    rows.push(['Add: deposits in transit', fromCents_(sum(dit)), dit.length + ' item(s)']);
    dit.forEach(k => rows.push(['      ' + k[K.ID] + '  ' + fmtDate_(k[K.DATE]) + '  ' + k[K.DESC], fromCents_(cents_(k[K.AMT])), 'In books, not yet cleared the bank']));
    rows.push(['Less: outstanding checks / payments', fromCents_(sum(oc)), oc.length + ' item(s)']);
    oc.forEach(k => rows.push(['      ' + k[K.ID] + '  ' + fmtDate_(k[K.DATE]) + '  ' + k[K.DESC], fromCents_(cents_(k[K.AMT])), 'In books, not yet cleared the bank']));
    bold.push(rows.length + 2);
    rows.push(['Adjusted bank balance', fromCents_(adjBank), '']);
    rows.push(['Book balance (opening + QBO register)', fromCents_(bookBal), '']);
    rows.push(['Add/(less): bank activity not yet booked', fromCents_(sum(bankOnly)), bankOnly.length + ' line(s) – see Bank_Txns']);
    bold.push(rows.length + 2);
    rows.push(['Adjusted book balance', fromCents_(adjBook), '']);
    bold.push(rows.length + 2);
    rows.push(['DIFFERENCE', fromCents_(diff), diff === 0 ? '✔ Ties to the penny' : '✖ Does not tie – investigate']);
    rows.push(['', '', '']);
  });
  const sh = writeTable_(S.REC, ['Bank Reconciliation – ' + PERIOD_LABEL, 'Amount', 'Detail'], rows);
  if (rows.length) sh.getRange(2, 2, rows.length, 1).setNumberFormat(MONEY_FMT);
  bold.forEach(r => sh.getRange(r, 1, 1, 3).setFontWeight('bold'));
  sh.setColumnWidth(1, 460); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 300);
  return results;
}

// ===================== STEP E: QBO JOURNAL ENTRY IMPORT =====================
function stepExport_(bank, entities, coa) {
  const entMap = {};
  entities.forEach(e => { entMap[e.code] = e; });
  const valid = {};
  coa.forEach(a => { valid[a] = true; });
  const out = [];
  let n = 0;
  bank.forEach(r => {
    if (r[B.STATUS] !== 'To Book' || !valid[r[B.ACCT]]) return;
    const e = entMap[r[B.ENT]];
    const amt = Math.abs(Number(r[B.AMT]));
    const jn = 'BF-' + r[B.ID];
    const date = fmtDate_(r[B.DATE]);
    const cp = String(r[B.NOTES]).match(/Counterparty: (\w+)/);
    const name = cp ? (entMap[cp[1]] ? entMap[cp[1]].name : cp[1]) : '';
    const moneyIn = Number(r[B.AMT]) > 0;
    out.push([r[B.ENT], jn, date, e.cash, moneyIn ? amt : '', moneyIn ? '' : amt, r[B.DESC], '', r[B.ID]]);
    out.push([r[B.ENT], jn, date, r[B.ACCT], moneyIn ? '' : amt, moneyIn ? amt : '', r[B.DESC], name, r[B.ID]]);
    r[B.STATUS] = 'Ready to import';
    n++;
  });
  const sh = writeTable_(S.EXPORT, ['Entity', 'JournalNo', 'JournalDate', 'Account', 'Debits', 'Credits',
    'Description', 'Name', 'Source Txn'], out);
  if (out.length) sh.getRange(2, 5, out.length, 2).setNumberFormat('#,##0.00');
  sh.autoResizeColumns(1, 9);
  return n;
}

// ===================== STEP F: SUMMARY =====================
function stepSummary_(bank, entities, recs, aiStats) {
  const sh = getOrCreate_(S.SUMMARY);
  sh.clear();
  sh.getRange(1, 1).setValue('Multi-Entity Close Assistant – ' + PERIOD_LABEL).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue('Last run: ' + now_() + '   |   Built with Google Apps Script + Claude');

  const headers = ['Entity', 'Bank lines', 'Matched to books', 'Auto: rules', 'Auto: AI',
    'Ready to import', 'Held for review', 'Bank rec difference', 'Rec status'];
  const rows = entities.map(e => {
    const rs = bank.filter(r => r[B.ENT] === e.code);
    const c = f => rs.filter(f).length;
    const diff = recs[e.code] || 0;
    return [e.code + ' – ' + e.name, rs.length, c(r => r[B.STATUS] === 'Booked'),
      c(r => r[B.METHOD] === 'Rule'), c(r => String(r[B.METHOD]).indexOf('AI') === 0),
      c(r => r[B.STATUS] === 'Ready to import'), c(r => r[B.STATUS] === 'Exception'),
      fromCents_(diff), diff === 0 ? '✔ Ties' : '✖ Investigate'];
  });
  const tot = ['TOTAL'];
  for (let i = 1; i <= 7; i++) tot.push(rows.reduce((t, r) => t + Number(r[i]), 0));
  tot.push(rows.every(r => r[8] === '✔ Ties') ? '✔ All entities tie' : '✖ Investigate');
  rows.push(tot);

  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#1f3a5f').setFontColor('#ffffff');
  sh.getRange(5, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(5 + rows.length - 1, 1, 1, headers.length).setFontWeight('bold');
  sh.getRange(5, 8, rows.length, 1).setNumberFormat(MONEY_FMT);

  const total = bank.length;
  const touchless = bank.filter(r => r[B.STATUS] === 'Booked' || r[B.STATUS] === 'Ready to import').length;
  let line = 7 + rows.length;
  sh.getRange(line, 1).setValue('Touchless rate: ' + Math.round(touchless / Math.max(total, 1) * 100) + '% of bank lines (' +
    touchless + ' of ' + total + ') were matched or categorized and ready to import with no manual entry.').setFontWeight('bold');
  sh.getRange(line + 1, 1).setValue('Each entity is a separate QBO company: filter QBO_JE_Import by Entity and import via Settings > Import data > Journal entries.');

  const queue = bank.filter(r => r[B.FLAGS]).map(r => [r[B.ID], r[B.ENT], r[B.DESC], Number(r[B.AMT]), r[B.STATUS], r[B.FLAGS]]);
  line += 3;
  sh.getRange(line, 1).setValue('Review queue (' + queue.length + ' flagged lines)').setFontSize(13).setFontWeight('bold');
  if (queue.length) {
    sh.getRange(line + 1, 1, 1, 6).setValues([['Txn', 'Entity', 'Description', 'Amount', 'Status', 'Flags']])
      .setFontWeight('bold').setBackground('#e8eef6');
    sh.getRange(line + 2, 1, queue.length, 6).setValues(queue);
    sh.getRange(line + 2, 4, queue.length, 1).setNumberFormat(MONEY_FMT);
  }
  sh.setColumnWidth(1, 300);
  for (let c = 2; c <= 9; c++) sh.setColumnWidth(c, 125);
  sh.setColumnWidth(3, 260); sh.setColumnWidth(6, 520);
}

// ===================== FORMATTING =====================
function formatBank_(n) {
  const sh = sheet_(S.BANK);
  if (!n) return;
  sh.getRange(2, 2, n, 1).setNumberFormat('MM/dd/yyyy');
  sh.getRange(2, 5, n, 1).setNumberFormat(MONEY_FMT);
  sh.getRange(2, 9, n, 1).setNumberFormat('0%');
  sh.getRange(2, B.APPROVE + 1, n, 1).insertCheckboxes();
  const vals = sh.getRange(2, 1, n, BANK_HEADERS.length).getValues();
  const colors = vals.map(r => {
    const c = r[B.STATUS] === 'Exception' ? '#fde2e1' : r[B.STATUS] === 'Ready to import' ? '#e3f4e1'
      : r[B.STATUS] === 'Booked' ? '#f1f3f4' : '#ffffff';
    return BANK_HEADERS.map(() => c);
  });
  sh.getRange(2, 1, n, BANK_HEADERS.length).setBackgrounds(colors);
  sh.setColumnWidth(4, 320); sh.setColumnWidth(7, 240); sh.setColumnWidth(10, 320); sh.setColumnWidth(11, 420);
}
function formatBook_(n) {
  const sh = sheet_(S.BOOK);
  if (!n) return;
  sh.getRange(2, 2, n, 1).setNumberFormat('MM/dd/yyyy');
  sh.getRange(2, 5, n, 1).setNumberFormat(MONEY_FMT);
  sh.autoResizeColumns(1, BOOK_HEADERS.length);
}

// ===================== LEARNING: AI PICKS -> RULES =====================
function promoteApprovedToRules() {
  const bank = readRows_(S.BANK, BANK_HEADERS.length);
  const rules = readRows_(S.RULES, RULE_HEADERS.length);
  const existing = {};
  rules.forEach(r => { existing[norm_(r[0])] = true; });
  const added = [];
  bank.forEach(r => {
    if (r[B.APPROVE] !== true || !r[B.ACCT] || String(r[B.METHOD]).indexOf('AI') !== 0) return;
    const p = ruleKeyword_(r[B.DESC]);
    if (!p || existing[p]) return;
    const acct = (r[B.ACCT] === DUE_FROM || r[B.ACCT] === DUE_TO) ? IC_TOKEN : r[B.ACCT];
    rules.push([p, acct, 'Learned from approved AI pick on ' + r[B.ID] + ' (' + now_().slice(0, 10) + ')', 'N']);
    existing[p] = true;
    added.push(p + '  →  ' + acct);
  });
  writeTable_(S.RULES, RULE_HEADERS, rules);
  sheet_(S.RULES).autoResizeColumns(1, RULE_HEADERS.length);
  log_('Promote AI picks to rules', added.length ? added.join('; ') : 'Nothing new to add');
  safeAlert_('Rules updated', added.length
    ? 'Added ' + added.length + ' rule(s):\n\n' + added.join('\n') + '\n\nNext run, these are handled by rules, with no AI call.'
    : 'No new approved AI picks to save. Tick "Approve" on AI rows you agree with, then try again.');
}
// Description up to the first token containing a digit or '#', e.g. "HOME DEPOT #4471" -> "HOME DEPOT"
function ruleKeyword_(desc) {
  const toks = norm_(desc).split(' ');
  const keep = [];
  for (let i = 0; i < toks.length; i++) { if (/[0-9#]/.test(toks[i])) break; keep.push(toks[i]); }
  while (keep.length && keep[keep.length - 1] === '-') keep.pop();
  return keep.join(' ');
}

// ===================== DEMO DATA =====================
function setupDemo() {
  try {
    if (sheet_(S.BANK)) {
      const ui = SpreadsheetApp.getUi();
      if (ui.alert('Reset demo data?', 'This rebuilds every tab with fresh sample data.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    }
  } catch (e) { /* no UI (run from editor): just proceed */ }

  [S.SUMMARY, S.BANK, S.BOOK, S.REC, S.EXPORT, S.ENT, S.COA, S.RULES, S.SETTINGS, S.LOG].forEach(getOrCreate_);
  ss_().getSheets().forEach(sh => {
    if (/^(Sheet|Hoja|Feuille|Tabelle)\s?1$/.test(sh.getName()) && sh.getLastRow() === 0) ss_().deleteSheet(sh);
  });
  clearAiCache();

  const d = day => new Date(2026, 8, day); // September 2026

  // ---- Entities ----
  const ents = [
    ['HOLD', 'Oakline Holdings LLC', 'Holding LLC – investments & office', 'OAKLINE HOLDINGS', '1010 Cash – Operating', 1250000, 0],
    ['TRUST', 'Oakline 2019 Family Trust', 'Irrevocable family trust', 'OAKLINE FAMILY TRUST', '1010 Cash – Operating', 480000, 0],
    ['PROP', 'Oakline Properties LLC', 'Real estate LLC – rental property', 'OAKLINE PROPERTIES', '1010 Cash – Operating', 95000, 0]
  ];

  // ---- Bank feed (what the bank says) ----
  const bankRaw = [
    ['HOLD', 2, 'WIRE OUT CAPITAL CALL - SUMMIT GROWTH FUND III', -250000],
    ['HOLD', 5, 'SCHWAB TRANSFER TO BROKERAGE', -100000],
    ['HOLD', 8, 'INTERCO TRANSFER TO OAKLINE PROPERTIES', -40000],
    ['HOLD', 10, 'DISTRIBUTION - HARBORVIEW RE FUND II', 62500],
    ['HOLD', 12, 'AMEX PAYMENT', -18432.17],
    ['HOLD', 15, 'PAYROLL - GUSTO', -24800],
    ['HOLD', 16, 'KLEIN & MARSH CPA INV 2026-114', -8750],
    ['HOLD', 19, 'AERIS JET CHARTER', -38900],
    ['HOLD', 20, 'INTEREST PAID', 1243.88],
    ['HOLD', 25, 'GOOGLE WORKSPACE', -86.40],
    ['HOLD', 25, 'GOOGLE WORKSPACE', -86.40],
    ['HOLD', 28, 'WIRE IN FROM OAKLINE FAMILY TRUST', 15000],
    ['HOLD', 29, 'COINBASE TRANSFER', -50000],
    ['HOLD', 30, 'BANK SERVICE FEE', -45],
    ['TRUST', 3, 'FIDELITY MONEYLINE TRANSFER', 20000],
    ['TRUST', 9, 'TRUSTEE FEE - NORTHGATE TRUST CO', -3125],
    ['TRUST', 15, 'DISTRIBUTION TO BENEFICIARY - J. OAKLINE', -12000],
    ['TRUST', 18, 'CHARITABLE GIFT - COASTAL ARTS FOUNDATION', -25000],
    ['TRUST', 27, 'TRANSFER TO OAKLINE HOLDINGS', -15000],
    ['TRUST', 30, 'INTEREST PAID', 312.40],
    ['TRUST', 30, 'IRS USATAXPYMT ESTIMATED TAX', -45000],
    ['PROP', 1, 'RENT DEPOSIT - UNIT 2B', 4200],
    ['PROP', 2, 'RENT DEPOSIT - UNIT 3A', 3850],
    ['PROP', 5, 'MORTGAGE PMT - FIRST COASTAL BANK', -9812.44],
    ['PROP', 9, 'INTERCO TRANSFER FROM OAKLINE HOLDINGS', 40000],
    ['PROP', 12, 'HOME DEPOT #4471', -1284.19],
    ['PROP', 14, 'DUKE ENERGY', -612.30],
    ['PROP', 16, 'PINELLAS COUNTY TAX COLLECTOR', -14906],
    ['PROP', 20, 'BRIGHT PROPERTY MGMT FEE', -1215],
    ['PROP', 24, 'GULF ROOFING - ROOF REPLACEMENT DEPOSIT', -18500],
    ['PROP', 26, 'INTERCO TRANSFER TO OAKLINE HOLDINGS', -10000],
    ['PROP', 30, 'BANK SERVICE FEE', -25]
  ];
  const bank = bankRaw.map((x, i) => {
    const row = BANK_HEADERS.map(() => '');
    row[B.ID] = 'B-' + String(i + 1).padStart(3, '0');
    row[B.DATE] = d(x[1]); row[B.ENT] = x[0]; row[B.DESC] = x[2]; row[B.AMT] = x[3];
    row[B.APPROVE] = false;
    return row;
  });

  // ---- Book register (what the outside accountant already recorded in QBO) ----
  const book = [
    ['K-001', d(2), 'HOLD', 'Capital call – Summit Growth Fund III', -250000, ''],
    ['K-002', d(5), 'HOLD', 'Transfer to Schwab brokerage', -100000, ''],
    ['K-003', d(11), 'HOLD', 'AmEx payment – August statement', -18423.17, ''],
    ['K-004', d(15), 'HOLD', 'Payroll – September', -24800, ''],
    ['K-005', d(29), 'HOLD', 'Check 1044 – October office rent', -6500, ''],
    ['K-006', d(15), 'TRUST', 'Beneficiary distribution – J. Oakline', -12000, ''],
    ['K-007', d(30), 'TRUST', 'Deposit – custodian cash sweep', 5000, ''],
    ['K-008', d(1), 'PROP', 'Rent – Unit 2B', 4200, ''],
    ['K-009', d(5), 'PROP', 'Mortgage payment – First Coastal', -9812.44, ''],
    ['K-010', d(28), 'PROP', 'Check 2210 – Gulf Plumbing', -740, '']
  ];

  // Statement ending balance = opening + bank activity (as if typed from the bank statement)
  ents.forEach(e => {
    const c = cents_(e[5]) + bank.filter(r => r[B.ENT] === e[0]).reduce((t, r) => t + cents_(r[B.AMT]), 0);
    e[6] = fromCents_(c);
  });

  // ---- Standardized chart of accounts (same across every entity) ----
  const coa = [
    ['1010', 'Cash – Operating', 'Asset'], ['1500', 'Investments – Brokerage', 'Asset'],
    ['1510', 'Investments – Private Funds', 'Asset'], ['1550', 'Digital Assets', 'Asset'],
    ['1600', 'Building & Improvements', 'Asset'], ['1700', 'Prepaid / Estimated Taxes', 'Asset'],
    ['1800', 'Due From Affiliates', 'Asset'], ['2100', 'Credit Card Payable', 'Liability'],
    ['2600', 'Mortgage Payable', 'Liability'], ['2800', 'Due To Affiliates', 'Liability'],
    ['3100', 'Member / Grantor Contributions', 'Equity'], ['3500', 'Distributions to Members / Beneficiaries', 'Equity'],
    ['4100', 'Interest Income', 'Income'], ['4200', 'Dividend Income', 'Income'],
    ['4300', 'Rental Income', 'Income'], ['4400', 'Fund Distributions – Income', 'Income'],
    ['6100', 'Salaries & Wages', 'Expense'], ['6300', 'Professional Fees – Tax & Accounting', 'Expense'],
    ['6310', 'Professional Fees – Legal', 'Expense'], ['6350', 'Trustee & Fiduciary Fees', 'Expense'],
    ['6400', 'Software & Subscriptions', 'Expense'], ['6600', 'Travel', 'Expense'],
    ['6700', 'Repairs & Maintenance', 'Expense'], ['6710', 'Property Management Fees', 'Expense'],
    ['6720', 'Utilities', 'Expense'], ['6750', 'Property Taxes', 'Expense'],
    ['6800', 'Charitable Contributions', 'Expense'], ['6900', 'Bank Fees', 'Expense']
  ];

  // ---- Rules (checked top to bottom; first match wins) ----
  const rules = [
    ['INTERCO', IC_TOKEN, 'Intercompany – Due From if money out, Due To if money in', 'N'],
    ['TRANSFER TO OAKLINE', IC_TOKEN, 'Intercompany', 'N'],
    ['WIRE IN FROM OAKLINE', IC_TOKEN, 'Intercompany', 'N'],
    ['CAPITAL CALL', '1510 Investments – Private Funds', 'Increases fund basis; update commitment tracker', 'N'],
    ['SCHWAB TRANSFER', '1500 Investments – Brokerage', '', 'N'],
    ['COINBASE', '1550 Digital Assets', 'Tie to exchange statement / cost basis', 'N'],
    ['AMEX PAYMENT', '2100 Credit Card Payable', '', 'N'],
    ['PAYROLL', '6100 Salaries & Wages', '', 'N'],
    ['INTEREST PAID', '4100 Interest Income', '', 'N'],
    ['BANK SERVICE FEE', '6900 Bank Fees', '', 'N'],
    ['GOOGLE WORKSPACE', '6400 Software & Subscriptions', '', 'N'],
    ['RENT DEPOSIT', '4300 Rental Income', '', 'N'],
    ['MORTGAGE PMT', '2600 Mortgage Payable', 'Split principal / interest per lender statement', 'Y'],
    ['IRS USATAXPYMT', '1700 Prepaid / Estimated Taxes', 'Match to tax calendar estimate', 'N']
  ];

  const settings = [
    ['Large item threshold ($)', 25000],
    ['Match window (days)', 3],
    ['AI model', 'claude-sonnet-5-5'],
    ['AI confidence threshold (0-1)', 0.8],
    ['AI batch size', 20]
  ];

  writeTable_(S.ENT, ENT_HEADERS, ents);
  sheet_(S.ENT).getRange(2, 6, ents.length, 2).setNumberFormat(MONEY_FMT);
  sheet_(S.ENT).autoResizeColumns(1, ENT_HEADERS.length);
  writeTable_(S.BANK, BANK_HEADERS, bank);
  formatBank_(bank.length);
  writeTable_(S.BOOK, BOOK_HEADERS, book);
  formatBook_(book.length);
  writeTable_(S.COA, ['Account #', 'Account Name', 'Type'], coa);
  sheet_(S.COA).autoResizeColumns(1, 3);
  writeTable_(S.RULES, RULE_HEADERS, rules);
  sheet_(S.RULES).autoResizeColumns(1, RULE_HEADERS.length);
  writeTable_(S.SETTINGS, ['Setting', 'Value'], settings);
  sheet_(S.SETTINGS).autoResizeColumns(1, 2);
  writeTable_(S.REC, ['Bank Reconciliation – ' + PERIOD_LABEL, 'Amount', 'Detail'], []);
  writeTable_(S.EXPORT, ['Entity', 'JournalNo', 'JournalDate', 'Account', 'Debits', 'Credits', 'Description', 'Name', 'Source Txn'], []);
  const sum = getOrCreate_(S.SUMMARY);
  sum.clear();
  sum.getRange(1, 1).setValue('Multi-Entity Close Assistant – ' + PERIOD_LABEL).setFontSize(16).setFontWeight('bold');
  sum.getRange(2, 1).setValue('Demo data loaded. Next: Close Tools > Save Claude API key, then Close Tools > ▶ Run full close.');
  log_('Setup demo data', bank.length + ' bank lines, ' + book.length + ' book entries, 3 entities');
  try { ss_().setActiveSheet(sum); } catch (e) { /* ignore */ }
  toast_('Demo data ready.');
}
