/*****************************************************************
 * RESTAURANT BILLING & FEEDBACK AUTOMATION  -  Code.gs
 * Container-bound script: open the Google Sheet > Extensions > Apps Script.
 * Sections: CONFIG | SETUP | SHEET FORMATTING | FORM BUILDER |
 *           EMAIL TEMPLATES | CORE LOGIC | TRIGGER HANDLERS |
 *           TRIGGER INSTALLER | UTILITIES
 *****************************************************************/

// ===================================================================
// CONFIG  (edit everything here; nothing else needs touching)
// ===================================================================
const CONFIG = {
  RESTAURANT_NAME: '[PLACEHOLDER: restaurant name]',
  OWNER_EMAIL:     '[PLACEHOLDER: owner email]',
  SENDER_NAME:     '',                       // optional; blank = uses RESTAURANT_NAME as the inbox sender name
  // Customers' replies go here. For query detection (column J) this must be
  // this Gmail account, or leave as placeholder so replies reach the sender.
  REPLY_TO:        '[PLACEHOLDER: reply-to email]',
  LOGO_URL:        '',                       // optional, public https image URL
  ADDRESS:         '[PLACEHOLDER: restaurant address]',
  PHONE:           '[PLACEHOLDER: restaurant phone]',
  CURRENCY_SYMBOL: '₹',
  AMOUNT_GROUPING: 'INTL',                   // 'INTL' = 1,234,567.00 | 'INDIAN' = 12,34,567.00 (emails only)
  TIMEZONE:        'Asia/Kolkata',
  REMINDER_INTERVAL_MINUTES: 10,
  MAX_REMINDERS:   2,                        // dropdown labels exist for 1-2 only
  LOW_SCORE_THRESHOLD: 3.0,
  ESCALATION_TEXT: '[PLACEHOLDER: to be provided]',
  ENABLE_OWNER_ALERT: true,
  SHEET_NAME:      'Billing',                // [DEFAULT - not specified in brief; change if you like]
  RESPONSES_SHEET_NAME: 'Responses',         // form's linked destination sheet
  SHEET_ROWS:      1000,                     // rows formatted / validated
  FORM_LIMIT_ONE_RESPONSE: false,            // keep FALSE: true blocks repeat customers (limit is per Google account, not per invoice). Duplicates are already blocked per invoice in code.
  FORM_DESCRIPTION: 'Thank you for dining with us. Please take a moment to rate your visit. It helps us serve you better.',
  FORM_CONFIRMATION: 'Thank you for your feedback. We truly appreciate your time.',
  FORM_ID:         ''                        // auto-saved to Script Properties after Create Form
};

// ---- Fixed structure (do not edit unless you change the sheet layout) ----
const COL = { NAME:1, EMAIL:2, INV:3, AMOUNT:4, STATUS:5, FOLLOW:6, REPLY:7, SENT:8,
              FORM:9, QUERY:10, SCORE:11, REMCOUNT:12, LASTREM:13, ERR:14 };
const HEADERS = ['Name','Email','Invoice Number','Invoice Amount','Status','Follow-up',
  'Need to Reply','Time Sent','Time Received (Form)','Time Received (Query)',
  'Score','Reminder Count','Last Reminder Time','Error Log'];
const STATUS_LIST = ['Pending','Ready','Sent','Feedback Sent','Stop'];
const FOLLOW_LIST = ['Not Started','Awaiting Response','Reminder 1 Sent','Reminder 2 Sent',
  'Responded','Reminders Exhausted','Stopped'];
const REPLY_LIST  = ['Yes','No','Resolved'];
const RATING_TITLES = ['Food Taste and Quality','Service and Staff','Ambience','Cleanliness','Value for Money'];
const INVOICE_FIELD_TITLE = 'Invoice Number';
const PROTECT_DESC = 'Script-managed column';
const TRIGGER_HANDLERS = ['onEditHandler','onFormSubmitHandler','safetySweep','scanQueries','dailySummary'];
const MAX_RUN_MS = 4.5 * 60 * 1000;          // stay under the 6-minute execution limit

const BRAND = { BLACK:'#0B0B0B', GOLD:'#C9A24B', CREAM:'#F5EFE0', INK:'#1C1A16', DEEPGOLD:'#8A6D1F' };
const FONT  = "Georgia, 'Times New Roman', Times, serif";

// ===================================================================
// SETUP  (menu + one-click setup)
// ===================================================================

/** Simple trigger: builds the custom menu when the sheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Billing Automation')
    .addItem('Run Full Setup (all steps)', 'runFullSetup')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addItem('Create Form', 'createFeedbackForm')
    .addItem('Install Triggers', 'installTriggers')
    .addItem('Send Ready Now', 'sendReadyNow')
    .addItem('Remove Triggers', 'removeTriggers')
    .addToUi();
}

/** One click: sheet, then form, then triggers. Safe to re-run. */
function runFullSetup() {
  setupSheet();
  createFeedbackForm();
  installTriggers();
  notify_('Setup complete', 'Sheet, form and triggers are ready. Now follow the manual checklist for the form theme.');
}

// ===================================================================
// SHEET FORMATTING
// ===================================================================

/** Creates/formats the main sheet: headers, colours, validation, conditional rules, protection. Idempotent. */
function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss);
  const rows = CONFIG.SHEET_ROWS, nCols = HEADERS.length, bodyRows = rows - 1;

  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());

  // Header row
  sh.getRange(1, 1, 1, nCols).setValues([HEADERS])
    .setBackground(BRAND.BLACK).setFontColor(BRAND.GOLD).setFontWeight('bold')
    .setFontFamily('Georgia').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(1, COL.SCORE, 1, 4).setBackground('#3A3226');   // helper columns K-N: distinct header shade
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);

  // Body: banded cream rows; H-N slightly muted
  const body = sh.getRange(2, 1, bodyRows, nCols);
  const bgs = [];
  for (let r = 0; r < bodyRows; r++) {
    const even = r % 2 === 0, line = [];
    for (let c = 0; c < nCols; c++) {
      const muted = c >= COL.SENT - 1;
      line.push(muted ? (even ? '#EDE6D2' : '#F2EDDD') : (even ? '#F5EFE0' : '#FBF7EC'));
    }
    bgs.push(line);
  }
  body.setBackgrounds(bgs).setFontFamily('Georgia').setFontSize(10)
      .setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setFontColor(BRAND.INK)
      .setBorder(true, true, true, true, true, true, BRAND.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2, COL.SENT, bodyRows, 7).setFontColor('#6E6757');

  // Column widths
  [170, 230, 130, 130, 120, 160, 120, 160, 160, 160, 80, 100, 160, 320]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  // Number formats and alignment
  const col = function (c) { return sh.getRange(2, c, bodyRows, 1); };
  col(COL.INV).setNumberFormat('@');                                   // keep invoice numbers as typed text
  col(COL.AMOUNT).setNumberFormat('"' + CONFIG.CURRENCY_SYMBOL + '" #,##0.00').setHorizontalAlignment('right');
  [COL.SENT, COL.FORM, COL.QUERY, COL.LASTREM].forEach(function (c) {
    col(c).setNumberFormat('dd-mmm-yyyy hh:mm').setHorizontalAlignment('center');
  });
  col(COL.SCORE).setNumberFormat('0.00').setHorizontalAlignment('center');
  col(COL.REMCOUNT).setNumberFormat('0').setHorizontalAlignment('center');
  [COL.STATUS, COL.FOLLOW, COL.REPLY].forEach(function (c) { col(c).setHorizontalAlignment('center'); });

  applyValidation_(sh, rows);
  applyConditionalFormats_(sh, rows);

  // Group helper columns K-N (only once)
  try { if (!sh.getColumnGroup(COL.SCORE, 1)) sh.getRange(1, COL.SCORE, 1, 4).shiftColumnGroupDepth(1); }
  catch (e) { try { sh.getRange(1, COL.SCORE, 1, 4).shiftColumnGroupDepth(1); } catch (e2) { Logger.log('Column group: ' + e2); } }

  // Warning-only protection on script-managed columns (F, H-N)
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === PROTECT_DESC) p.remove();
  });
  [[COL.FOLLOW, 1], [COL.SENT, 7]].forEach(function (a) {
    sh.getRange(2, a[0], bodyRows, a[1]).protect().setDescription(PROTECT_DESC).setWarningOnly(true);
  });

  ss.toast('Sheet "' + CONFIG.SHEET_NAME + '" is formatted.', 'Billing Automation', 5);
}

/** Returns the main sheet, creating (or renaming a blank default sheet) if missing. */
function getOrCreateSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sh) return sh;
  const all = ss.getSheets();
  if (all.length === 1 && all[0].getLastRow() === 0 && all[0].getLastColumn() === 0) {
    all[0].setName(CONFIG.SHEET_NAME);
    return all[0];
  }
  return ss.insertSheet(CONFIG.SHEET_NAME, 0);
}

/** Applies dropdowns and input validation (invalid input rejected). */
function applyValidation_(sh, rows) {
  const n = rows - 1;
  const dv = function () { return SpreadsheetApp.newDataValidation(); };
  const list = function (c, values) {
    sh.getRange(2, c, n, 1).setDataValidation(
      dv().requireValueInList(values, true).setAllowInvalid(false).setHelpText('Choose a value from the list.').build());
  };
  list(COL.STATUS, STATUS_LIST);
  list(COL.FOLLOW, FOLLOW_LIST);
  list(COL.REPLY, REPLY_LIST);

  sh.getRange(2, COL.NAME, n, 1).setDataValidation(
    dv().requireFormulaSatisfied('=LEN(TRIM(A2))>0').setAllowInvalid(false).setHelpText('Name cannot be blank.').build());
  sh.getRange(2, COL.EMAIL, n, 1).setDataValidation(
    dv().requireTextIsEmail().setAllowInvalid(false).setHelpText('Enter a valid email address.').build());
  sh.getRange(2, COL.INV, n, 1).setDataValidation(
    dv().requireFormulaSatisfied('=AND(LEN(TRIM(C2))>0,SUMPRODUCT(--($C$2:$C$' + rows + '=C2))=1)')
      .setAllowInvalid(false).setHelpText('Invoice Number must be unique and not blank.').build());
  sh.getRange(2, COL.AMOUNT, n, 1).setDataValidation(
    dv().requireNumberGreaterThan(0).setAllowInvalid(false).setHelpText('Enter a number greater than 0.').build());
}

/** Applies conditional formatting for Status, Need to Reply and Score. Replaces all rules on the sheet. */
function applyConditionalFormats_(sh, rows) {
  const n = rows - 1;
  const rng = function (c) { return sh.getRange(2, c, n, 1); };
  const txt = function (c, text, bg, fg, bold) {
    let b = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg);
    if (bold) b = b.setBold(true);
    return b.setRanges([rng(c)]).build();
  };
  const rules = [
    txt(COL.REPLY, 'Yes',      '#C0392B', '#FFFFFF', true),
    txt(COL.REPLY, 'No',       '#CFE8CF', '#1E4620', false),
    txt(COL.REPLY, 'Resolved', '#D9D9D9', '#555555', false),
    txt(COL.STATUS, 'Pending',       '#D9D9D9', '#444444', false),
    txt(COL.STATUS, 'Ready',         '#BBD6F5', '#123A66', false),
    txt(COL.STATUS, 'Sent',          '#CFE8CF', '#1E4620', false),
    txt(COL.STATUS, 'Feedback Sent', '#B7DDB7', '#1E4620', false),
    txt(COL.STATUS, 'Stop',          '#7B1F1F', '#FFFFFF', true),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($K2),$K2<' + Number(CONFIG.LOW_SCORE_THRESHOLD) + ')')
      .setFontColor('#C0392B').setBold(true).setRanges([rng(COL.SCORE)]).build()
  ];
  sh.setConditionalFormatRules(rules);
}

// ===================================================================
// FORM BUILDER
// ===================================================================

/** Creates the feedback form once, links it to the "Responses" sheet, saves FORM_ID. Idempotent. */
function createFeedbackForm() {
  const ss = SpreadsheetApp.getActive();
  const existing = getFormId_();
  if (existing) {
    try {
      const f = FormApp.openById(existing);
      f.setLimitOneResponsePerUser(CONFIG.FORM_LIMIT_ONE_RESPONSE);   // re-apply setting on the existing form
      notify_('Form already exists', 'Not creating another one.\nEdit URL: ' + f.getEditUrl());
      return f;
    } catch (e) { Logger.log('Stored FORM_ID not accessible; creating a new form. ' + e); }
  }

  const form = FormApp.create(CONFIG.RESTAURANT_NAME + ' Dining Feedback');
  form.setDescription(CONFIG.FORM_DESCRIPTION)
      .setConfirmationMessage(CONFIG.FORM_CONFIRMATION)
      .setCollectEmail(false)
      .setLimitOneResponsePerUser(CONFIG.FORM_LIMIT_ONE_RESPONSE)
      .setShowLinkToRespondAgain(false)
      .setAcceptingResponses(true);
  if (typeof form.setPublished === 'function') { try { form.setPublished(true); } catch (e) { Logger.log(e); } }

  form.addTextItem().setTitle(INVOICE_FIELD_TITLE).setRequired(true).setHelpText('Please do not change this');
  RATING_TITLES.forEach(function (t) {
    form.addScaleItem().setTitle(t).setBounds(1, 5).setLabels('Poor', 'Excellent').setRequired(true);
  });

  PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());

  // Link responses and rename the auto-created sheet to "Responses"
  const before = ss.getSheets().map(function (s) { return s.getSheetId(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  const created = ss.getSheets().filter(function (s) { return before.indexOf(s.getSheetId()) === -1; })[0];
  if (created && !ss.getSheetByName(CONFIG.RESPONSES_SHEET_NAME)) {
    created.setName(CONFIG.RESPONSES_SHEET_NAME);
    try {
      created.getRange(1, 1, 1, Math.max(created.getLastColumn(), 1))
        .setBackground(BRAND.BLACK).setFontColor(BRAND.GOLD).setFontWeight('bold').setFontFamily('Georgia');
      created.setFrozenRows(1);
    } catch (e) { Logger.log('Responses styling: ' + e); }
  } else if (created) {
    Logger.log('A sheet named "' + CONFIG.RESPONSES_SHEET_NAME + '" already exists; left linked sheet as "' + created.getName() + '".');
  }

  notify_('Form created', 'Edit URL: ' + form.getEditUrl() + '\n\nNext: apply the theme using the manual checklist.');
  return form;
}

/** Builds the per-customer prefilled form link (Invoice Number prefilled). */
function prefilledUrl_(inv) {
  const form = getForm_();
  const item = form.getItems(FormApp.ItemType.TEXT)
    .filter(function (i) { return i.getTitle() === INVOICE_FIELD_TITLE; })[0];
  if (!item) throw new Error('Form is missing the "' + INVOICE_FIELD_TITLE + '" field.');
  return form.createResponse().withItemResponse(item.asTextItem().createResponse(inv)).toPrefilledUrl();
}

// ===================================================================
// EMAIL TEMPLATES  (HTML + plain-text fallback; every dynamic value escaped)
// ===================================================================

/** Wraps inner HTML in the shared fine-dining shell (black / gold / cream). */
function shell_(preheader, inner) {
  const c = CONFIG;
  const brand = c.LOGO_URL
    ? `<img src="${esc_(c.LOGO_URL)}" alt="${esc_(c.RESTAURANT_NAME)}" width="140" style="display:block;margin:0 auto;max-width:140px;height:auto;border:0;">`
    : `<div style="font-family:${FONT};font-size:22px;letter-spacing:6px;color:${BRAND.GOLD};text-transform:uppercase;">${esc_(c.RESTAURANT_NAME)}</div>`;
  const goldLine = `<tr><td height="2" bgcolor="${BRAND.GOLD}" style="height:2px;line-height:2px;font-size:2px;background-color:${BRAND.GOLD};">&nbsp;</td></tr>`;
  return `<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
<title>${esc_(c.RESTAURANT_NAME)}</title>
<style>
@media only screen and (max-width:620px){
  .px{padding-left:22px !important;padding-right:22px !important;}
  .total{font-size:36px !important;}
  .stack{display:block !important;width:100% !important;text-align:left !important;}
}
@media (prefers-color-scheme: dark){
  .card{background-color:#17140E !important;}
  .callout{background-color:#2A2416 !important;}
  .ink{color:#F5EFE0 !important;}
  .lbl{color:${BRAND.GOLD} !important;}
}
</style></head>
<body style="margin:0;padding:0;background-color:${BRAND.BLACK};">
<div style="display:none;font-size:1px;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;">${esc_(preheader)}</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="${BRAND.BLACK}" style="background-color:${BRAND.BLACK};"><tr><td align="center" style="padding:24px 12px;">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;">
<tr><td align="center" bgcolor="${BRAND.BLACK}" style="padding:30px 20px 24px 20px;background-color:${BRAND.BLACK};">${brand}</td></tr>
${goldLine}
<tr><td class="px card" bgcolor="${BRAND.CREAM}" style="background-color:${BRAND.CREAM};padding:44px 48px;">${inner}</td></tr>
${goldLine}
<tr><td align="center" bgcolor="${BRAND.BLACK}" style="padding:26px 20px;background-color:${BRAND.BLACK};font-family:${FONT};font-size:12px;line-height:1.8;letter-spacing:1px;color:#CFC6AE;">${esc_(c.ADDRESS)}<br>${esc_(c.PHONE)}</td></tr>
</table></td></tr></table></body></html>`;
}

/** Paragraph helper. */
function p_(html, align) {
  return `<p class="ink" style="margin:0 0 18px 0;font-family:${FONT};font-size:16px;line-height:1.75;color:${BRAND.INK};${align ? 'text-align:' + align + ';' : ''}">${html}</p>`;
}
/** Thin gold divider. */
function rule_() {
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:26px 0;"><tr><td height="1" bgcolor="${BRAND.GOLD}" style="height:1px;line-height:1px;font-size:1px;background-color:${BRAND.GOLD};">&nbsp;</td></tr></table>`;
}
/** Small-caps label + value pair. */
function kv_(label, valueHtml) {
  return `<div class="lbl" style="font-family:${FONT};font-variant:small-caps;letter-spacing:2px;font-size:13px;color:${BRAND.DEEPGOLD};">${label}</div><div class="ink" style="font-family:${FONT};font-size:17px;line-height:1.6;color:${BRAND.INK};margin-top:4px;">${valueHtml}</div>`;
}
/** Gold call-to-action button (table-based, works in Gmail/Outlook). */
function button_(url, label) {
  return `<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="margin:30px auto 14px auto;"><tr><td align="center" bgcolor="${BRAND.GOLD}" style="background-color:${BRAND.GOLD};border-radius:2px;"><a href="${esc_(url)}" target="_blank" style="display:inline-block;padding:15px 34px;font-family:${FONT};font-size:15px;font-weight:bold;letter-spacing:2px;color:${BRAND.BLACK};text-decoration:none;text-transform:uppercase;">${esc_(label)}</a></td></tr></table>`;
}
/** Distinct callout block (used for the low-score escalation text). */
function callout_(title, textHtml) {
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 26px 0;"><tr><td width="4" bgcolor="${BRAND.GOLD}" style="width:4px;background-color:${BRAND.GOLD};font-size:1px;">&nbsp;</td><td class="callout" bgcolor="#EADFC0" style="background-color:#EADFC0;padding:18px 20px;"><div class="lbl" style="font-family:${FONT};font-variant:small-caps;letter-spacing:2px;font-size:13px;color:${BRAND.DEEPGOLD};">${title}</div><div class="ink" style="font-family:${FONT};font-size:15px;line-height:1.7;color:${BRAND.INK};margin-top:6px;">${textHtml}</div></td></tr></table>`;
}
/** Sign-off block. */
function signoff_() {
  return p_(`With warm regards,<br>${esc_(senderName_())}<br>${esc_(CONFIG.RESTAURANT_NAME)}`);
}
/** Muted link fallback shown under the button. */
function linkFallback_(url) {
  return `<p class="ink" style="margin:0 0 6px 0;font-family:${FONT};font-size:12px;line-height:1.6;color:#6E6757;text-align:center;">Button not working? Copy this link into your browser:<br><span style="word-break:break-all;">${esc_(url)}</span></p>`;
}

/** Invoice email. rec = {name,email,inv,amount}. */
function invoiceTpl_(rec) {
  const date = fmtDate_(new Date(), 'dd MMM yyyy');
  const amount = formatAmount_(rec.amount);
  const inner = `
<div class="lbl" style="font-family:${FONT};font-size:13px;letter-spacing:7px;color:${BRAND.DEEPGOLD};text-align:center;">INVOICE</div>
${rule_()}
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
  <td class="stack" width="50%" valign="top" style="padding-bottom:8px;">${kv_('Invoice No.', esc_(rec.inv))}</td>
  <td class="stack" width="50%" valign="top" align="right" style="padding-bottom:8px;">${kv_('Date', esc_(date))}</td>
</tr></table>
${rule_()}
${kv_('Billed To', esc_(rec.name) + '<br><span style="font-size:14px;color:#6E6757;">' + esc_(rec.email) + '</span>')}
${rule_()}
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="${BRAND.BLACK}" style="background-color:${BRAND.BLACK};margin:6px 0 30px 0;"><tr><td align="center" style="padding:30px 16px;">
  <div style="font-family:${FONT};font-variant:small-caps;letter-spacing:4px;font-size:13px;color:${BRAND.CREAM};">Total</div>
  <div class="total" style="font-family:${FONT};font-size:46px;line-height:1.2;color:${BRAND.GOLD};margin-top:8px;">${esc_(amount)}</div>
</td></tr></table>
${p_('Thank you for dining with us. It was a pleasure to host you, and we hope to welcome you again soon.', 'center')}`;
  const text = `INVOICE - ${CONFIG.RESTAURANT_NAME}\n\nInvoice No.: ${rec.inv}\nDate: ${date}\nBilled to: ${rec.name} (${rec.email})\n\nTotal: ${amount}\n\nThank you for dining with us. It was a pleasure to host you, and we hope to welcome you again soon.\n\n${CONFIG.ADDRESS}\n${CONFIG.PHONE}`;
  return { subject: `Your invoice ${rec.inv} from ${CONFIG.RESTAURANT_NAME}`, html: shell_(`Invoice ${rec.inv}: ${amount}`, inner), text: text };
}

/** Feedback request email (sent right after the invoice). */
function feedbackTpl_(rec, url) {
  const inner = `
${p_(`Dear ${esc_(rec.name)},`)}
${p_(`Thank you for choosing ${esc_(CONFIG.RESTAURANT_NAME)}. We would be grateful to hear how your visit was. Your honest words help us look after every guest better.`)}
${button_(url, 'Share Your Experience')}
${p_('It takes under a minute.', 'center')}
${linkFallback_(url)}
${rule_()}
${signoff_()}`;
  const text = `Dear ${rec.name},\n\nThank you for choosing ${CONFIG.RESTAURANT_NAME}. We would be grateful to hear how your visit was.\n\nShare your experience (takes under a minute):\n${url}\n\nWith warm regards,\n${senderName_()}\n${CONFIG.RESTAURANT_NAME}\n${CONFIG.ADDRESS} | ${CONFIG.PHONE}`;
  return { subject: `How was your visit? Feedback for invoice ${rec.inv}`, html: shell_('A short note on your recent visit. It takes under a minute.', inner), text: text };
}

/** Reminder email; n = 1 (gentle) or 2+ (final, still gentle). */
function reminderTpl_(rec, url, n) {
  const first = n <= 1;
  const lead = first
    ? 'In case it slipped through the cracks, we would still love to hear about your recent visit. A minute of your time means a great deal to us.'
    : 'This is our last note on this. If you have a moment, your feedback would mean a lot to our team. If now is not a good time, no worries at all.';
  const inner = `
${p_(`Dear ${esc_(rec.name)},`)}
${p_(esc_(lead))}
${button_(url, 'Share Your Experience')}
${p_('It takes under a minute.', 'center')}
${linkFallback_(url)}
${rule_()}
${signoff_()}`;
  const text = `Dear ${rec.name},\n\n${lead}\n\nShare your experience (takes under a minute):\n${url}\n\nWith warm regards,\n${senderName_()}\n${CONFIG.RESTAURANT_NAME}\n${CONFIG.ADDRESS} | ${CONFIG.PHONE}`;
  return {
    subject: first ? `A gentle reminder: your feedback for invoice ${rec.inv}` : `A last note: your feedback for invoice ${rec.inv}`,
    html: shell_(first ? 'A gentle reminder about your feedback.' : 'One last note about your feedback.', inner), text: text
  };
}

/** Thank-you follow-up; if escalation=true adds the ESCALATION_TEXT callout. */
function thankYouTpl_(rec, escalation) {
  const hasEsc = escalation && !isPlaceholder_(CONFIG.ESCALATION_TEXT);
  const escHtml = hasEsc ? callout_('We would like to make this right', esc_(CONFIG.ESCALATION_TEXT).replace(/\r?\n/g, '<br>')) : '';
  const lead = escalation
    ? 'Thank you for your honest feedback. We take it seriously, and we are sorry your visit did not meet your expectations.'
    : 'Thank you for taking the time to share your thoughts. Feedback like yours is what keeps us improving, and we are delighted to have hosted you.';
  const inner = `
${p_(`Dear ${esc_(rec.name)},`)}
${p_(esc_(lead))}
${escHtml}
${rule_()}
${signoff_()}`;
  const text = `Dear ${rec.name},\n\n${lead}\n\n${hasEsc ? CONFIG.ESCALATION_TEXT + '\n\n' : ''}With warm regards,\n${senderName_()}\n${CONFIG.RESTAURANT_NAME}\n${CONFIG.ADDRESS} | ${CONFIG.PHONE}`;
  return { subject: `Thank you for your feedback: invoice ${rec.inv}`, html: shell_('Thank you for your feedback.', inner), text: text, escalationOmitted: escalation && !hasEsc };
}

/** Plain, scannable owner alert for a low score. */
function ownerAlertTpl_(rec, score, link) {
  const row = function (k, v) { return `<tr><td style="padding:6px 14px 6px 0;color:#555;font-weight:bold;">${k}</td><td style="padding:6px 0;">${v}</td></tr>`; };
  const html = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;max-width:560px;">
<h3 style="margin:0 0 12px 0;color:#C0392B;">Low feedback score: reply needed</h3>
<table role="presentation" cellpadding="0" cellspacing="0" border="0">
${row('Customer', esc_(rec.name))}${row('Email', esc_(rec.email))}${row('Invoice', esc_(rec.inv))}${row('Score', esc_(String(score)) + ' / 5 (threshold ' + esc_(String(CONFIG.LOW_SCORE_THRESHOLD)) + ')')}
</table>
<p style="margin:16px 0 0 0;"><a href="${esc_(link)}" target="_blank">Open the sheet row</a></p></div>`;
  const text = `Low feedback score: reply needed\nCustomer: ${rec.name}\nEmail: ${rec.email}\nInvoice: ${rec.inv}\nScore: ${score} / 5 (threshold ${CONFIG.LOW_SCORE_THRESHOLD})\nSheet row: ${link}`;
  return { subject: `[Low score ${score}] ${rec.name}: invoice ${rec.inv}`, html: html, text: text };
}

/** Sends one prepared email through MailApp. */
function sendMail_(to, tpl) {
  const opts = { to: to, subject: String(tpl.subject).replace(/[\r\n]+/g, ' '), htmlBody: tpl.html, body: tpl.text, name: senderName_() };
  if (isValidEmail_(CONFIG.REPLY_TO)) opts.replyTo = CONFIG.REPLY_TO;
  MailApp.sendEmail(opts);
}

// ===================================================================
// CORE LOGIC  (callers must hold the script lock)
// ===================================================================

/** Send flow for one row: invoice (Ready) then feedback email (Sent). Retries the feedback stage if it failed earlier. */
function processSendRow_(sh, row) {
  const v = sh.getRange(row, 1, 1, COL.ERR).getValues()[0];
  const status = String(v[COL.STATUS - 1]).trim();
  const follow = String(v[COL.FOLLOW - 1]).trim();
  const needsInvoice = status === 'Ready';
  const needsFeedback = status === 'Sent' && v[COL.SENT - 1] instanceof Date && (follow === '' || follow === 'Not Started');
  if (!needsInvoice && !needsFeedback) return false;

  const rec = { name: String(v[0]).trim(), email: String(v[1]).trim(), inv: String(v[2]).trim(), amount: v[3] };
  try {
    assertConfig_();
    if (needsInvoice) {
      validateRow_(sh, rec);
      if (MailApp.getRemainingDailyQuota() < 2) throw new Error('Daily email quota below 2. Will retry on a later sweep.');
      sendMail_(rec.email, invoiceTpl_(rec));
      sh.getRange(row, COL.STATUS).setValue('Sent');
      sh.getRange(row, COL.SENT).setValue(new Date());
      SpreadsheetApp.flush();
    } else {
      if (!isValidEmail_(rec.email)) throw new Error('Invalid email address.');
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Daily email quota exhausted. Will retry later.');
    }
    // Stop could have been set while the invoice was sending
    if (String(sh.getRange(row, COL.STATUS).getValue()).trim() === 'Stop') return true;
    sendMail_(rec.email, feedbackTpl_(rec, prefilledUrl_(rec.inv)));
    sh.getRange(row, COL.STATUS).setValue('Feedback Sent');
    sh.getRange(row, COL.FOLLOW).setValue('Awaiting Response');
    sh.getRange(row, COL.ERR).clearContent();
    return true;
  } catch (err) {
    setError_(sh, row, err.message || String(err));
    return false;
  }
}

/** Validates a Ready row (name, email, invoice number, amount, uniqueness). Throws on problems. */
function validateRow_(sh, rec) {
  if (!rec.name) throw new Error('Name is blank.');
  if (!isValidEmail_(rec.email)) throw new Error('Invalid email address.');
  if (!rec.inv) throw new Error('Invoice Number is blank.');
  if (typeof rec.amount !== 'number' || !isFinite(rec.amount) || rec.amount <= 0) throw new Error('Invoice Amount must be a number greater than 0.');
  if (findInvoiceRows_(sh, rec.inv).length > 1) throw new Error('Duplicate Invoice Number: "' + rec.inv + '".');
}

/** Processes one form response: score, write K/I, follow-up emails, owner alert. Idempotent. */
function processFormResponse_(sh, response) {
  const parsed = parseResponse_(response);
  if (!parsed.inv) return;
  const found = findInvoiceRows_(sh, parsed.inv);
  if (found.length !== 1) { Logger.log('Response for invoice "' + parsed.inv + '": ' + found.length + ' matching rows. Ignored.'); return; }
  const row = found[0];
  const cur = sh.getRange(row, 1, 1, COL.ERR).getValues()[0];
  const status = String(cur[COL.STATUS - 1]).trim();

  if (status === 'Stop') return;                                   // Stop: ignore
  if (status !== 'Feedback Sent' && status !== 'Sent') return;     // feedback never requested
  if (cur[COL.FORM - 1] instanceof Date || String(cur[COL.FOLLOW - 1]).trim() === 'Responded' || cur[COL.SCORE - 1] !== '') return; // duplicate
  if (parsed.score === null) { setError_(sh, row, 'Form response has missing or invalid ratings.'); return; }

  // Claim the stage first so a retry can never send the follow-up twice
  const score = parsed.score;
  sh.getRange(row, COL.SCORE).setValue(score);
  sh.getRange(row, COL.FORM).setValue(response.getTimestamp());
  sh.getRange(row, COL.FOLLOW).setValue('Responded');
  const low = score < CONFIG.LOW_SCORE_THRESHOLD;
  const g = String(cur[COL.REPLY - 1]).trim();
  if (low) sh.getRange(row, COL.REPLY).setValue('Yes');
  else if (g === '') sh.getRange(row, COL.REPLY).setValue('No');   // never overwrite a pending customer query ("Yes")
  SpreadsheetApp.flush();

  const rec = { name: String(cur[0]).trim(), email: String(cur[1]).trim(), inv: String(cur[2]).trim() };
  try {
    assertConfig_();
    const tpl = thankYouTpl_(rec, low);
    sendMail_(rec.email, tpl);
    if (tpl.escalationOmitted) setError_(sh, row, 'ESCALATION_TEXT is still a placeholder; callout omitted from the low-score email.');
  } catch (err) { setError_(sh, row, 'Follow-up email failed: ' + (err.message || err)); }

  if (low && CONFIG.ENABLE_OWNER_ALERT) {
    try {
      if (!isValidEmail_(CONFIG.OWNER_EMAIL)) throw new Error('OWNER_EMAIL is not set.');
      const link = SpreadsheetApp.getActive().getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;
      sendMail_(CONFIG.OWNER_EMAIL, ownerAlertTpl_(rec, score, link));
    } catch (err) { setError_(sh, row, 'Owner alert failed: ' + (err.message || err)); }
  }
}

/** Extracts the invoice number and average score from a FormResponse. */
function parseResponse_(response) {
  let inv = '';
  const ratings = {};
  response.getItemResponses().forEach(function (ir) {
    const t = ir.getItem().getTitle();
    if (t === INVOICE_FIELD_TITLE) inv = String(ir.getResponse()).trim();
    else if (RATING_TITLES.indexOf(t) > -1) ratings[t] = Number(ir.getResponse());
  });
  let sum = 0, ok = true;
  RATING_TITLES.forEach(function (t) {
    const n = ratings[t];
    if (typeof n !== 'number' || !isFinite(n) || n < 1 || n > 5) ok = false; else sum += n;
  });
  return { inv: inv, score: ok ? Math.round((sum / RATING_TITLES.length) * 100) / 100 : null };
}

/** Sends due reminders and marks exhausted rows. Re-checks Stop right before each send. */
function runReminders_(sh, data, t0) {
  let quota = MailApp.getRemainingDailyQuota();
  const now = new Date();
  const done = ['Stopped', 'Responded', 'Reminders Exhausted'];
  data.forEach(function (r) {
    if (Date.now() - t0 > MAX_RUN_MS) return;
    if (r.status !== 'Feedback Sent' || done.indexOf(r.follow) > -1 || r.formTime instanceof Date) return;
    const base = (r.lastRem instanceof Date) ? r.lastRem : r.sent;
    if (!(base instanceof Date)) return;
    if ((now.getTime() - base.getTime()) / 60000 < CONFIG.REMINDER_INTERVAL_MINUTES) return;
    try {
      const fresh = sh.getRange(r.row, COL.STATUS, 1, 2).getValues()[0];
      if (String(fresh[0]).trim() !== 'Feedback Sent' || done.indexOf(String(fresh[1]).trim()) > -1) return;
      if (r.remCount >= CONFIG.MAX_REMINDERS) { sh.getRange(r.row, COL.FOLLOW).setValue('Reminders Exhausted'); return; }
      if (quota < 1) { setError_(sh, r.row, 'Daily email quota exhausted. Reminder postponed.'); return; }
      assertConfig_();
      const n = r.remCount + 1;
      sendMail_(r.email, reminderTpl_(r, prefilledUrl_(r.inv), n));
      quota--;
      sh.getRange(r.row, COL.REMCOUNT).setValue(n);
      sh.getRange(r.row, COL.LASTREM).setValue(new Date());
      sh.getRange(r.row, COL.FOLLOW).setValue('Reminder ' + Math.min(n, 2) + ' Sent');
      sh.getRange(r.row, COL.ERR).clearContent();
    } catch (err) { setError_(sh, r.row, 'Reminder failed: ' + (err.message || err)); }
  });
}

/** Safety net: picks up responses whose form trigger was missed (only rows still awaiting). */
function reconcileResponses_(sh, data) {
  const active = data.filter(function (r) {
    return r.status === 'Feedback Sent' && !(r.formTime instanceof Date) && r.sent instanceof Date &&
      ['Awaiting Response', 'Reminder 1 Sent', 'Reminder 2 Sent'].indexOf(r.follow) > -1;
  });
  if (!active.length || !getFormId_()) return;
  const since = new Date(Math.min.apply(null, active.map(function (r) { return r.sent.getTime(); })) - 60000);
  getForm_().getResponses(since).forEach(function (resp) { processFormResponse_(sh, resp); });
}

/** Finds sheet rows whose Invoice Number matches (trimmed, case-insensitive). */
function findInvoiceRows_(sh, inv) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const key = norm_(inv), out = [];
  sh.getRange(2, COL.INV, last - 1, 1).getValues().forEach(function (x, i) { if (norm_(x[0]) === key) out.push(i + 2); });
  return out;
}

// ===================================================================
// TRIGGER HANDLERS
// ===================================================================

/** Installable onEdit: Ready -> send flow; Stop -> Follow-up = Stopped. Handles pasted multi-row edits. */
function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SHEET_NAME) return;
    if (COL.STATUS < e.range.getColumn() || COL.STATUS > e.range.getLastColumn()) return;
    const r1 = Math.max(e.range.getRow(), 2), r2 = e.range.getLastRow();
    if (r2 < r1) return;
    const t0 = Date.now();
    withLock_(function () {
      for (let r = r1; r <= r2; r++) {
        if (Date.now() - t0 > MAX_RUN_MS) break;                  // sweep will finish the rest
        const st = String(sh.getRange(r, COL.STATUS).getValue()).trim();
        if (st === 'Stop') sh.getRange(r, COL.FOLLOW).setValue('Stopped');
        else if (st === 'Ready') processSendRow_(sh, r);
      }
    }, 25000);
  } catch (err) { Logger.log('onEditHandler error: ' + err); }
}

/** Installable form-submit handler: scores the response and sends follow-ups. */
function onFormSubmitHandler(e) {
  try {
    if (!e || !e.response) return;
    withLock_(function () { processFormResponse_(getSheet_(), e.response); }, 60000);
  } catch (err) { Logger.log('onFormSubmitHandler error: ' + err); }
}

/** Every minute: Stop sync, missed Ready rows, missed responses, then reminder timing. */
function safetySweep() {
  try {
    withLock_(function () {
      const sh = getSheet_(), t0 = Date.now();
      let data = readRows_(sh);
      data.forEach(function (r) {
        if (r.status === 'Stop' && r.follow !== 'Stopped') sh.getRange(r.row, COL.FOLLOW).setValue('Stopped');
      });
      data.forEach(function (r) {
        if (Date.now() - t0 > MAX_RUN_MS) return;
        if (r.status === 'Ready' || r.status === 'Sent') processSendRow_(sh, r.row);
      });
      reconcileResponses_(sh, readRows_(sh));
      runReminders_(sh, readRows_(sh), t0);
    }, 50000);
  } catch (err) { Logger.log('safetySweep error: ' + err); }
}

/** Menu action: send every Ready row now (also retries a failed feedback stage). */
function sendReadyNow() {
  let count = 0;
  withLock_(function () {
    const sh = getSheet_(), t0 = Date.now();
    readRows_(sh).forEach(function (r) {
      if (Date.now() - t0 > MAX_RUN_MS) return;
      if ((r.status === 'Ready' || r.status === 'Sent') && processSendRow_(sh, r.row)) count++;
    });
  }, 30000);
  notify_('Send Ready Now', count + ' row(s) processed. Check the Error Log column for any problems.');
}

/** Every 5 minutes: finds customer replies in Gmail; writes J and sets Need to Reply = Yes (unless Resolved). */
function scanQueries() {
  try {
    withLock_(function () {
      const sh = getSheet_(), t0 = Date.now(), cutoff = Date.now() - 30 * 86400000;
      readRows_(sh).forEach(function (r) {
        if (Date.now() - t0 > MAX_RUN_MS) return;
        if (r.status !== 'Sent' && r.status !== 'Feedback Sent') return;   // Stop / Pending rows skipped
        if (!(r.sent instanceof Date) || r.sent.getTime() < cutoff || r.queryTime instanceof Date) return;
        if (!isValidEmail_(r.email) || !r.inv) return;
        try {
          const email = r.email.toLowerCase();
          const day = fmtDate_(new Date(r.sent.getTime() - 86400000), 'yyyy/MM/dd');
          let found = null;
          GmailApp.search('from:"' + email + '" after:' + day, 0, 10).forEach(function (th) {
            if (found || norm_(th.getFirstMessageSubject()).indexOf(norm_(r.inv)) < 0) return;
            th.getMessages().forEach(function (m) {
              if (!found && m.getFrom().toLowerCase().indexOf(email) > -1 && m.getDate().getTime() > r.sent.getTime()) found = m.getDate();
            });
          });
          if (found) {
            sh.getRange(r.row, COL.QUERY).setValue(found);
            if (String(sh.getRange(r.row, COL.REPLY).getValue()).trim() !== 'Resolved') sh.getRange(r.row, COL.REPLY).setValue('Yes');
          }
        } catch (err) { setError_(sh, r.row, 'Query scan failed: ' + (err.message || err)); }
      });
    }, 50000);
  } catch (err) { Logger.log('scanQueries error: ' + err); }
}

/** Daily ~9 AM: emails the owner a summary. */
function dailySummary() {
  try {
    if (!isValidEmail_(CONFIG.OWNER_EMAIL)) { Logger.log('dailySummary skipped: OWNER_EMAIL not set.'); return; }
    const sh = getSheet_(), data = readRows_(sh);
    const today = fmtDate_(new Date(), 'yyyy-MM-dd');
    const sentToday = data.filter(function (r) { return r.sent instanceof Date && fmtDate_(r.sent, 'yyyy-MM-dd') === today; }).length;
    const pending = data.filter(function (r) { return r.status === 'Feedback Sent' && !(r.formTime instanceof Date) && r.follow !== 'Stopped'; }).length;
    const needReply = data.filter(function (r) { return r.reply === 'Yes'; });
    const errors = data.filter(function (r) { return r.err !== ''; });
    const quota = MailApp.getRemainingDailyQuota();
    const li = function (arr, f) { return arr.slice(0, 15).map(function (r) { return '<li>' + f(r) + '</li>'; }).join('') || '<li>None</li>'; };
    const html = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;max-width:560px;">
<h3 style="margin:0 0 12px 0;">Daily summary: ${esc_(CONFIG.RESTAURANT_NAME)}</h3>
<p>Invoices sent today: <b>${sentToday}</b><br>Awaiting feedback: <b>${pending}</b><br>Rows needing a reply: <b>${needReply.length}</b><br>Rows with errors: <b>${errors.length}</b><br>Email quota remaining: <b>${quota}</b></p>
<p><b>Needs a reply</b></p><ul>${li(needReply, function (r) { return esc_(r.name) + ' (' + esc_(r.inv) + ')'; })}</ul>
<p><b>Errors</b></p><ul>${li(errors, function (r) { return esc_(r.name) + ' (' + esc_(r.inv) + '): ' + esc_(r.err); })}</ul>
<p><a href="${esc_(SpreadsheetApp.getActive().getUrl())}">Open the sheet</a></p></div>`;
    const text = `Daily summary: ${CONFIG.RESTAURANT_NAME}\nInvoices sent today: ${sentToday}\nAwaiting feedback: ${pending}\nRows needing a reply: ${needReply.length}\nRows with errors: ${errors.length}\nEmail quota remaining: ${quota}\n${SpreadsheetApp.getActive().getUrl()}`;
    sendMail_(CONFIG.OWNER_EMAIL, { subject: `Daily summary: ${CONFIG.RESTAURANT_NAME}`, html: html, text: text });
  } catch (err) { Logger.log('dailySummary error: ' + err); }
}

// ===================================================================
// TRIGGER INSTALLER
// ===================================================================

/** Deletes this project's old triggers, then installs the five installable triggers. (onOpen is a simple trigger.) */
function installTriggers() {
  const ss = SpreadsheetApp.getActive();
  getSheet_();                                              // throws if sheet missing
  const formId = getFormId_();
  if (!formId) throw new Error('No form found. Run "Create Form" first.');
  removeTriggers_();
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(formId).onFormSubmit().create();
  ScriptApp.newTrigger('safetySweep').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('scanQueries').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(9).nearMinute(0).inTimezone(CONFIG.TIMEZONE).create();
  ss.toast('5 triggers installed.', 'Billing Automation', 5);
}

/** Menu action: removes this project's triggers. */
function removeTriggers() {
  const n = removeTriggers_();
  notify_('Triggers removed', n + ' trigger(s) deleted. Nothing will send automatically until you run Install Triggers.');
}

/** Deletes triggers whose handler is one of ours. Returns the count. */
function removeTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) > -1) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

// ===================================================================
// UTILITIES
// ===================================================================

/** Runs fn while holding the script lock; returns false if the lock could not be obtained. */
function withLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs || 20000)) { Logger.log('Could not obtain lock; skipped this run.'); return false; }
  try { fn(); return true; } finally { lock.releaseLock(); }
}

/** Main sheet (throws if missing). */
function getSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" not found. Run Setup Sheet.');
  return sh;
}

/** Reads all data rows into objects. */
function readRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, COL.ERR).getValues().map(function (v, i) {
    return { row: i + 2, name: String(v[0]).trim(), email: String(v[1]).trim(), inv: String(v[2]).trim(), amount: v[3],
      status: String(v[4]).trim(), follow: String(v[5]).trim(), reply: String(v[6]).trim(),
      sent: v[7], formTime: v[8], queryTime: v[9], score: v[10], remCount: Number(v[11]) || 0, lastRem: v[12], err: String(v[13]) };
  });
}

/** FORM_ID from Script Properties (falls back to CONFIG.FORM_ID). */
function getFormId_() { return PropertiesService.getScriptProperties().getProperty('FORM_ID') || CONFIG.FORM_ID || ''; }
/** Opens the feedback form. */
function getForm_() {
  const id = getFormId_();
  if (!id) throw new Error('No feedback form yet. Run "Create Form".');
  return FormApp.openById(id);
}

/** Writes a timestamped message to the row's Error Log (column N). */
function setError_(sh, row, msg) {
  sh.getRange(row, COL.ERR).setValue(fmtDate_(new Date(), 'dd-MMM-yyyy HH:mm') + ' | ' + String(msg).substring(0, 300));
}

/** Sender display name: SENDER_NAME if set, otherwise the restaurant name. */
function senderName_() { return isPlaceholder_(CONFIG.SENDER_NAME) ? CONFIG.RESTAURANT_NAME : CONFIG.SENDER_NAME; }

/** Blocks sending while required CONFIG values are still placeholders. */
function assertConfig_() {
  ['RESTAURANT_NAME', 'ADDRESS', 'PHONE'].forEach(function (k) {
    if (isPlaceholder_(CONFIG[k])) throw new Error('CONFIG.' + k + ' is still a placeholder. Edit the CONFIG block.');
  });
}

/** Shows an alert, or a toast / log when no UI is available. */
function notify_(title, msg) {
  Logger.log(title + ': ' + msg);                      // full text (incl. form URL) is always in the execution log
  try { SpreadsheetApp.getActive().toast(String(msg).split('\n')[0], title, 10); } catch (e) {}   // non-blocking
}

/** Escapes text for safe use in HTML (also attributes). */
function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
/** Trim + lowercase for comparisons. */
function norm_(s) { return String(s === null || s === undefined ? '' : s).trim().toLowerCase(); }
/** Basic email format check. */
function isValidEmail_(s) { return /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(String(s || '').trim()); }
/** True for empty values or "[PLACEHOLDER...]" strings. */
function isPlaceholder_(s) { const t = String(s || '').trim(); return t === '' || /^\[PLACEHOLDER/i.test(t); }
/** Formats a date in CONFIG.TIMEZONE. */
function fmtDate_(d, pattern) { return Utilities.formatDate(d, CONFIG.TIMEZONE, pattern); }
/** Formats an amount with currency symbol and grouping for emails. */
function formatAmount_(n) {
  const parts = Math.abs(Number(n)).toFixed(2).split('.');
  let int = parts[0];
  if (CONFIG.AMOUNT_GROUPING === 'INDIAN') {
    const last3 = int.slice(-3), rest = int.slice(0, -3);
    int = rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3 : last3;
  } else {
    int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  return CONFIG.CURRENCY_SYMBOL + ' ' + int + '.' + parts[1];
}
