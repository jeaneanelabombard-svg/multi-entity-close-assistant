/**
 * QUICKBOOKS ONLINE CONNECTION  (add-on file for the Multi-Entity Close Assistant)
 *
 * What it adds, under Close Tools > QuickBooks:
 *   1. Save QuickBooks app keys      – Intuit Client ID / Secret, stored privately (never in the sheet)
 *   2. Connect an entity             – sign in to ONE QuickBooks company per entity (accounting scope only)
 *   3. Test connection               – safety check: company name in QuickBooks must match the entity's legal name
 *   4. Create missing accounts       – builds the standardized chart of accounts (matched by account number)
 *   5. Load demo books               – SANDBOX ONLY: posts the opening balance + the demo book entries
 *   6. Pull books from QuickBooks    – replaces that entity's Book_Register rows with what is actually in QuickBooks
 *   7. Post approved lines           – posts ONLY lines marked "Ready to import" AND ticked Approve,
 *                                      after a confirmation, skipping anything already in QuickBooks,
 *                                      and logs every entry to the QBO_Posted tab (audit trail)
 *
 * Demo loop: 5 → 6 → Run full close → tick Approve on a few lines → 7 → 6 → Run full close
 *            (the posted lines now come back from QuickBooks and match as "Booked").
 *
 * Requires the OAuth2 for Apps Script library:
 *   Libraries (+) → Script ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF → Look up → Add (identifier: OAuth2)
 *
 * Known limitation (by design for the demo): "Pull books" reads journal entries that hit the cash account.
 * In production, cash also moves through checks, deposits, transfers and bank-feed categorizations,
 * so the pull would use the General Ledger report for the cash account instead.
 */

const QBO = {
  ENV: 'sandbox',
  API_BASE: 'https://sandbox-quickbooks.api.intuit.com/v3/company/',
  AUTH_URL: 'https://appcenter.intuit.com/connect/oauth2',
  TOKEN_URL: 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer',
  SCOPE: 'com.intuit.quickbooks.accounting',   // least access: no payments scope requested
  MINOR: '75',
  PERIOD_START: '2026-09-01',
  PERIOD_END: '2026-09-30',
  POSTED_TAB: 'QBO_Posted',
  ACCOUNTS_TAB: 'QBO_Accounts'
};

// QuickBooks account types for the standardized chart (anything not listed falls back to the COA "Type" column)
const QBO_TYPES = {
  '1010': 'Bank', '1500': 'Other Current Asset', '1510': 'Other Asset', '1550': 'Other Current Asset',
  '1600': 'Fixed Asset', '1700': 'Other Current Asset', '1710': 'Other Current Asset',
  '1800': 'Other Current Asset', '2100': 'Credit Card', '2600': 'Long Term Liability',
  '2800': 'Other Current Liability'
};
// Detail types QuickBooks needs spelled out (otherwise equity defaults to "Opening Balance Equity", which can exist only once)
const QBO_SUBTYPES = {
  '1600': 'Buildings', '3100': 'PartnerContributions', '3500': 'PartnerDistributions'
};
const QBO_TYPE_FALLBACK = {
  ASSET: 'Other Current Asset', LIABILITY: 'Other Current Liability', EQUITY: 'Equity',
  INCOME: 'Income', EXPENSE: 'Expense'
};
// Accounts QuickBooks needs for the demo books that are not on the COA tab
const QBO_EXTRA_ACCOUNTS = [['1710', 'Prepaid Expenses', 'Asset']];

// Demo only: the other side of each book entry the outside accountant "already recorded"
const SEED_OFFSETS = {
  'K-001': '1510', 'K-002': '1500', 'K-003': '2100', 'K-004': '6100', 'K-005': '1710',
  'K-006': '3500', 'K-007': '1500', 'K-008': '4300', 'K-009': '2600', 'K-010': '6700'
};

// ===================== MENU (called from onOpen in Code.gs) =====================
function qboMenu_() {
  return SpreadsheetApp.getUi().createMenu('QuickBooks')
    .addItem('1. Save QuickBooks app keys', 'qboSaveKeys')
    .addItem('2. Connect an entity', 'qboConnect')
    .addItem('3. Test connection', 'qboTestConnection')
    .addItem('4. Create missing accounts', 'qboCreateAccounts')
    .addItem('5. Load demo books (sandbox only)', 'qboLoadDemoBooks')
    .addItem('6. Pull books from QuickBooks', 'qboPullBooks')
    .addItem('7. Post approved lines', 'qboPostApproved')
    .addSeparator()
    .addItem('Pull chart of accounts', 'qboPullAccounts')
    .addItem('Disconnect an entity', 'qboDisconnect');
}

// ===================== KEYS, SIGN-IN, CONNECTION =====================
function qboSaveKeys() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt('QuickBooks Client ID',
    'Paste the Development Client ID from your Intuit app (Keys and credentials).', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK || !a.getResponseText().trim()) return;
  const b = ui.prompt('QuickBooks Client Secret',
    'Paste the Development Client Secret. It is stored privately in your Google account, never in the sheet.',
    ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK || !b.getResponseText().trim()) return;
  const p = PropertiesService.getUserProperties();
  p.setProperty('QBO_CLIENT_ID', a.getResponseText().trim());
  p.setProperty('QBO_CLIENT_SECRET', b.getResponseText().trim());
  ui.alert('Saved. Next: Close Tools > QuickBooks > 2. Connect an entity.');
}

function qboService_(code) {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('QBO_CLIENT_ID');
  const secret = props.getProperty('QBO_CLIENT_SECRET');
  if (!id || !secret) throw new Error('QuickBooks app keys are not saved. Use Close Tools > QuickBooks > 1. Save QuickBooks app keys.');
  return OAuth2.createService('QBO_' + code)
    .setAuthorizationBaseUrl(QBO.AUTH_URL)
    .setTokenUrl(QBO.TOKEN_URL)
    .setClientId(id)
    .setClientSecret(secret)
    .setScope(QBO.SCOPE)
    .setCallbackFunction('qboAuthCallback')
    .setPropertyStore(props)
    .setCache(CacheService.getUserCache())
    .setLock(LockService.getUserLock());
}

function qboRealm_(code) {
  return PropertiesService.getUserProperties().getProperty('QBO_REALM_' + code);
}

function qboConnect() {
  const e = qboPickEntity_();
  if (!e) return;
  const url = qboService_(e.code).getAuthorizationUrl({ entity: e.code });
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px">' +
    '<p>Sign in to the QuickBooks company for <b>' + esc_(e.name) + '</b>.</p>' +
    '<p><a href="' + url + '" target="_blank">Open QuickBooks sign-in</a></p>' +
    '<p style="color:#555">When the new tab says "Connected", close it and this box, then run 3. Test connection.</p>' +
    '</div>').setWidth(440).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Connect ' + e.code + ' to QuickBooks');
}

function qboAuthCallback(request) {
  const code = String(request.parameter.entity || '');
  const ok = qboService_(code).handleCallback(request);
  const realm = String(request.parameter.realmId || '');
  if (ok && realm) {
    PropertiesService.getUserProperties().setProperty('QBO_REALM_' + code, realm);
    return HtmlService.createHtmlOutput('<p style="font-family:Arial">Connected ' + esc_(code) +
      ' to QuickBooks company ID ' + esc_(realm) + '. You can close this tab.</p>');
  }
  return HtmlService.createHtmlOutput('<p style="font-family:Arial">Connection failed or was cancelled. Close this tab and try again.</p>');
}

function qboDisconnect() {
  const e = qboPickEntity_();
  if (!e) return;
  try { qboService_(e.code).reset(); } catch (err) { /* keys may be missing; nothing to reset */ }
  PropertiesService.getUserProperties().deleteProperty('QBO_REALM_' + e.code);
  log_('QBO disconnect', e.code);
  safeAlert_('Disconnected', e.code + ' is no longer connected to QuickBooks.');
}

// Safety stop: never read or post unless the connected company is the right entity
function qboCheckCompany_(e) {
  const realm = qboRealm_(e.code);
  const info = qboApi_(e.code, 'get', 'companyinfo/' + realm).CompanyInfo || {};
  const name = String(info.CompanyName || '');
  if (norm_(name) !== norm_(e.name)) {
    throw new Error('Safety stop: ' + e.code + ' is connected to QuickBooks company "' + name + '" (ID ' + realm +
      '), but the entity is "' + e.name + '". Rename the company in QuickBooks, or reconnect the right company.');
  }
  return { name: name, realm: realm };
}

function qboTestConnection() {
  const e = qboPickEntity_();
  if (!e) return;
  try {
    const c = qboCheckCompany_(e);
    log_('QBO test connection', e.code + ' → ' + c.name + ' (company ID ' + c.realm + ')');
    safeAlert_('Connected', e.code + ' → "' + c.name + '" (company ID ' + c.realm + ').\nThe names match, so reading and posting are allowed.');
  } catch (err) {
    safeAlert_('QuickBooks check failed', String(err.message || err));
  }
}

// ===================== API HELPERS =====================
function qboApi_(code, method, path, body) {
  const svc = qboService_(code);
  const realm = qboRealm_(code);
  if (!svc.hasAccess() || !realm) {
    throw new Error(code + ' is not connected to QuickBooks. Use Close Tools > QuickBooks > 2. Connect an entity.');
  }
  const url = QBO.API_BASE + realm + '/' + path + (path.indexOf('?') === -1 ? '?' : '&') + 'minorversion=' + QBO.MINOR;
  const opts = {
    method: method, muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + svc.getAccessToken(), Accept: 'application/json' }
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  const resp = UrlFetchApp.fetch(url, opts);
  const status = resp.getResponseCode();
  const text = resp.getContentText();
  if (status < 200 || status >= 300) throw new Error('QuickBooks error ' + status + ': ' + text.slice(0, 400));
  return JSON.parse(text);
}

function qboQuery_(code, table, where) {
  const out = [];
  let start = 1;
  while (true) {
    const q = 'select * from ' + table + (where ? ' where ' + where : '') + ' startposition ' + start + ' maxresults 1000';
    const res = qboApi_(code, 'get', 'query?query=' + encodeURIComponent(q));
    const page = (res.QueryResponse && res.QueryResponse[table]) || [];
    page.forEach(x => out.push(x));
    if (page.length < 1000) break;
    start += 1000;
  }
  return out;
}

function qboAccounts_(e) {
  const all = qboQuery_(e.code, 'Account', null);
  const byNum = {};
  const names = {};
  all.forEach(a => {
    if (a.AcctNum) byNum[String(a.AcctNum)] = a;
    names[norm_(a.Name)] = true;
  });
  return { all: all, byNum: byNum, names: names };
}

function qboJournalEntries_(e) {
  return qboQuery_(e.code, 'JournalEntry', "TxnDate <= '" + QBO.PERIOD_END + "'");
}

function qboJe_(doc, isoDate, desc, cashId, otherId, signedAmt) {
  const amt = Math.round(Math.abs(Number(signedAmt)) * 100) / 100;
  const inflow = Number(signedAmt) > 0;
  const line = (id, type) => ({
    Amount: amt, Description: String(desc).slice(0, 4000), DetailType: 'JournalEntryLineDetail',
    JournalEntryLineDetail: { PostingType: type, AccountRef: { value: String(id) } }
  });
  return {
    DocNumber: doc, TxnDate: isoDate, PrivateNote: String(desc).slice(0, 4000),
    Line: [line(cashId, inflow ? 'Debit' : 'Credit'), line(otherId, inflow ? 'Credit' : 'Debit')]
  };
}

// ===================== CHART OF ACCOUNTS =====================
function qboCreateAccounts() {
  const e = qboPickEntity_();
  if (!e) return;
  try {
    qboCheckCompany_(e);
    const have = qboAccounts_(e);
    const wanted = readRows_(S.COA, 3).map(r => [String(r[0]), String(r[1]), String(r[2])]).concat(QBO_EXTRA_ACCOUNTS);
    const made = [];
    const failed = [];
    wanted.forEach(([num, name, type]) => {
      if (have.byNum[num]) return;
      // The sandbox's sample company already uses names like "Travel" and "Utilities"; QuickBooks names must be unique
      const qboName = have.names[norm_(name)] ? num + ' ' + name : name;
      try {
        const body = {
          Name: qboName, AcctNum: num,
          AccountType: QBO_TYPES[num] || QBO_TYPE_FALLBACK[norm_(type)] || 'Expense'
        };
        if (QBO_SUBTYPES[num]) body.AccountSubType = QBO_SUBTYPES[num];
        qboApi_(e.code, 'post', 'account', body);
        made.push(num + ' ' + qboName);
      } catch (err) {
        failed.push(num + ' ' + name + ': ' + String(err.message || err).slice(0, 160));
      }
    });
    qboWriteAccountsTab_(e);
    log_('QBO create accounts', e.code + ': created ' + made.length + ', failed ' + failed.length);
    safeAlert_('Chart of accounts', 'Created ' + made.length + ' account(s) in QuickBooks.' +
      (failed.length ? '\n\nCould not create:\n' + failed.join('\n') : '\n\nEvery account on the COA tab now exists in QuickBooks.') +
      '\n\nSee the ' + QBO.ACCOUNTS_TAB + ' tab.');
  } catch (err) {
    safeAlert_('Stopped', String(err.message || err));
  }
}

function qboPullAccounts() {
  const e = qboPickEntity_();
  if (!e) return;
  try {
    qboCheckCompany_(e);
    const n = qboWriteAccountsTab_(e);
    log_('QBO pull accounts', e.code + ': ' + n + ' accounts');
    safeAlert_('Chart of accounts pulled', n + ' QuickBooks accounts for ' + e.code + ' are on the ' + QBO.ACCOUNTS_TAB + ' tab.');
  } catch (err) {
    safeAlert_('Stopped', String(err.message || err));
  }
}

function qboWriteAccountsTab_(e) {
  const headers = ['Entity', 'QBO Id', 'Number', 'Name', 'Type', 'Balance', 'On COA tab?'];
  const coaNums = {};
  readRows_(S.COA, 3).forEach(r => { coaNums[String(r[0])] = true; });
  const sh = sheet_(QBO.ACCOUNTS_TAB);
  const others = (sh && sh.getLastRow() > 1)
    ? sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues().filter(r => r[0] && r[0] !== e.code)
    : [];
  const mine = qboAccounts_(e).all
    .map(a => [e.code, String(a.Id), String(a.AcctNum || ''), a.Name, a.AccountType, Number(a.CurrentBalance || 0),
      a.AcctNum && coaNums[String(a.AcctNum)] ? 'Yes' : ''])
    .sort((x, y) => (x[2] || 'zzzz').localeCompare(y[2] || 'zzzz'));
  const all = others.concat(mine);
  const out = writeTable_(QBO.ACCOUNTS_TAB, headers, all);
  if (all.length) out.getRange(2, 6, all.length, 1).setNumberFormat(MONEY_FMT);
  out.autoResizeColumns(1, headers.length);
  return mine.length;
}

// ===================== DEMO BOOKS (SANDBOX ONLY) =====================
function qboLoadDemoBooks() {
  if (QBO.ENV !== 'sandbox') { safeAlert_('Not allowed', 'Demo books can only be loaded into a sandbox company.'); return; }
  const e = qboPickEntity_();
  if (!e) return;
  try {
    qboCheckCompany_(e);
    const accts = qboAccounts_(e).byNum;
    const cash = accts[acctNum_(e.cash)];
    const equity = accts['3100'];
    if (!cash || !equity) throw new Error('Run "4. Create missing accounts" first: ' + e.cash + ' and 3100 must exist in QuickBooks.');

    const existing = {};
    qboJournalEntries_(e).forEach(j => { if (j.DocNumber) existing[j.DocNumber] = true; });

    const items = [{ doc: 'OPEN-0831', date: '2026-08-31', desc: 'Opening cash balance', amt: e.opening, other: equity }];
    readRows_(S.BOOK, BOOK_HEADERS.length)
      .filter(k => k[K.ENT] === e.code && /^K-/.test(String(k[K.ID])))
      .forEach(k => items.push({
        doc: String(k[K.ID]), date: qboDate_(k[K.DATE]), desc: String(k[K.DESC]),
        amt: Number(k[K.AMT]), other: accts[SEED_OFFSETS[String(k[K.ID])]]
      }));

    const posted = [];
    const skipped = [];
    items.forEach(it => {
      if (existing[it.doc]) { skipped.push(it.doc + ' (already in QuickBooks)'); return; }
      if (!it.other) { skipped.push(it.doc + ' (offset account missing)'); return; }
      try {
        qboApi_(e.code, 'post', 'journalentry', qboJe_(it.doc, it.date, it.desc, cash.Id, it.other.Id, it.amt));
        posted.push(it.doc);
      } catch (err) {
        skipped.push(it.doc + ': ' + String(err.message || err).slice(0, 160));
      }
    });
    log_('QBO load demo books', e.code + ': posted ' + posted.length + ', skipped ' + skipped.length);
    safeAlert_('Demo books loaded', 'Posted ' + posted.length + ' entr(ies) to QuickBooks: ' + (posted.join(', ') || 'none') +
      (skipped.length ? '\n\nSkipped:\n' + skipped.join('\n') : '') + '\n\nNext: 6. Pull books from QuickBooks.');
  } catch (err) {
    safeAlert_('Stopped', String(err.message || err));
  }
}

// ===================== PULL BOOKS =====================
function qboPullBooks() {
  const e = qboPickEntity_();
  if (!e) return;
  try {
    qboCheckCompany_(e);
    const cash = qboAccounts_(e).byNum[acctNum_(e.cash)];
    if (!cash) throw new Error('Cash account ' + e.cash + ' is not in QuickBooks. Run "4. Create missing accounts" first.');

    let openingC = 0;
    const rows = [];
    qboJournalEntries_(e).forEach(je => {
      let c = 0;
      let desc = '';
      (je.Line || []).forEach(l => {
        const d = l.JournalEntryLineDetail;
        if (!d || !d.AccountRef || String(d.AccountRef.value) !== String(cash.Id)) return;
        c += (d.PostingType === 'Debit' ? 1 : -1) * cents_(l.Amount);
        if (!desc) desc = l.Description || '';
      });
      if (!c) return;
      if (je.TxnDate < QBO.PERIOD_START) { openingC += c; return; }
      const p = String(je.TxnDate).split('-').map(Number);
      rows.push([je.DocNumber || ('QBO-' + je.Id), new Date(p[0], p[1] - 1, p[2]), e.code,
        desc || je.PrivateNote || '', fromCents_(c), '']);
    });
    rows.sort((a, b) => a[1] - b[1]);

    const others = readRows_(S.BOOK, BOOK_HEADERS.length)
      .filter(k => k[K.ENT] !== e.code)
      .map(k => { k[K.MATCH] = ''; return k; });
    const all = others.concat(rows);
    writeTable_(S.BOOK, BOOK_HEADERS, all);
    formatBook_(all.length);

    const diff = openingC - cents_(e.opening);
    const tie = diff === 0
      ? 'Opening cash in QuickBooks (' + money_(fromCents_(openingC)) + ') ties to the Entities tab.'
      : 'WARNING: opening cash in QuickBooks is ' + money_(fromCents_(openingC)) + ' but the Entities tab says ' +
        money_(e.opening) + ' (difference ' + money_(fromCents_(diff)) + '). Fix this before trusting the bank rec.';
    log_('QBO pull books', e.code + ': ' + rows.length + ' September cash entries; ' + tie);
    safeAlert_('Books pulled from QuickBooks', e.code + ': ' + rows.length +
      ' September entries hitting ' + e.cash + ' now on Book_Register.\n\n' + tie + '\n\nNext: Close Tools > ▶ Run full close.');
  } catch (err) {
    safeAlert_('Stopped', String(err.message || err));
  }
}

// ===================== POST APPROVED LINES =====================
function qboPostApproved() {
  const e = qboPickEntity_();
  if (!e) return;
  const ui = SpreadsheetApp.getUi();
  try {
    const co = qboCheckCompany_(e);
    const picks = readRows_(S.BANK, BANK_HEADERS.length)
      .filter(r => r[B.ENT] === e.code && r[B.STATUS] === 'Ready to import' && r[B.APPROVE] === true);
    if (!picks.length) {
      ui.alert('Nothing to post', 'Only lines that are "Ready to import" AND ticked Approve are ever sent.\n\n' +
        'Tick Approve on the ' + e.code + ' lines you have reviewed, then try again.', ui.ButtonSet.OK);
      return;
    }
    const accts = qboAccounts_(e).byNum;
    const cash = accts[acctNum_(e.cash)];
    if (!cash) throw new Error('Cash account ' + e.cash + ' is not in QuickBooks.');
    const existing = {};
    qboJournalEntries_(e).forEach(j => { if (j.DocNumber) existing[j.DocNumber] = j.Id; });

    const gross = picks.reduce((t, r) => t + Math.abs(Number(r[B.AMT])), 0);
    if (ui.alert('Post to QuickBooks?', 'Post ' + picks.length + ' journal entr(ies), ' + money_(gross) +
      ' gross, to "' + co.name + '" (company ID ' + co.realm + ')?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    const who = Session.getActiveUser().getEmail() || 'unknown user';
    const posted = [];
    const skipped = [];
    picks.forEach(r => {
      const doc = 'BF-' + r[B.ID];
      if (existing[doc]) { skipped.push(doc + ': already in QuickBooks (ID ' + existing[doc] + ') – not posted twice'); return; }
      const acct = accts[acctNum_(r[B.ACCT])];
      if (!acct) { skipped.push(doc + ': account "' + r[B.ACCT] + '" is not in QuickBooks'); return; }
      const cp = String(r[B.NOTES]).match(/Counterparty: (\w+)/);
      const desc = String(r[B.DESC]) + (cp ? ' | Counterparty: ' + cp[1] : '') + ' | Bank line ' + r[B.ID];
      try {
        const res = qboApi_(e.code, 'post', 'journalentry',
          qboJe_(doc, qboDate_(r[B.DATE]), desc, cash.Id, acct.Id, Number(r[B.AMT])));
        posted.push([now_(), e.code, r[B.ID], doc, String(res.JournalEntry.Id), Number(r[B.AMT]), String(r[B.ACCT]), who]);
      } catch (err) {
        skipped.push(doc + ': ' + String(err.message || err).slice(0, 160));
      }
    });

    const sh = getOrCreate_(QBO.POSTED_TAB);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 8).setValues([['Posted at', 'Entity', 'Bank line', 'Journal no.', 'QBO Id', 'Amount', 'Account', 'Approved / posted by']])
        .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    if (posted.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, posted.length, 8).setValues(posted);
      sh.getRange(start, 6, posted.length, 1).setNumberFormat(MONEY_FMT);
      sh.autoResizeColumns(1, 8);
    }
    log_('QBO post approved', e.code + ': posted ' + posted.length + ', skipped ' + skipped.length);
    ui.alert('QuickBooks posting', 'Posted ' + posted.length + ' entr(ies). Each one is on the ' + QBO.POSTED_TAB + ' tab with its QuickBooks ID.' +
      (skipped.length ? '\n\nNot posted:\n' + skipped.join('\n') : '') +
      '\n\nNext: 6. Pull books from QuickBooks, then ▶ Run full close. Posted lines should come back as "Booked".', ui.ButtonSet.OK);
  } catch (err) {
    safeAlert_('QuickBooks posting stopped', String(err.message || err));
  }
}

// ===================== SMALL HELPERS =====================
function qboPickEntity_() {
  const ents = readEntities_();
  const codes = ents.map(x => x.code);
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Which entity?', 'Type an entity code: ' + codes.join(', ') +
    '\n(leave blank for ' + codes[0] + ')', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const code = norm_(r.getResponseText()) || codes[0];
  const e = ents.find(x => x.code === code);
  if (!e) { ui.alert('Unknown entity code: ' + code); return null; }
  return e;
}
function acctNum_(s) { return String(s || '').trim().split(' ')[0]; }
function qboDate_(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
