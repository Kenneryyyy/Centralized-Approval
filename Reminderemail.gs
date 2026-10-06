/**
 * ============================================================================
 *  SCHEDULED REMINDER EMAIL — paste into Code.gs (or add as a new file
 *  ReminderEmail.gs; Apps Script shares one global scope across .gs files).
 *
 *  Reuses from your existing Code.gs:
 *    TABLES, appSheetFind_(), normalizeApprovalRow_(), getManagersCached_()
 *  Sends via GmailApp (needs the gmail.send scope — see oauthScopes note in chat).
 *
 *  Entry points
 *    sendScheduledReminderEmail(recipientEmail)  -> one approver's pending list
 *    runScheduledReminders()                     -> trigger target: loops every
 *                                                   approver who has pending items
 *    createDailyReminderTrigger() / createWeeklyReminderTrigger()
 *    deleteReminderTriggers()
 * ============================================================================
 */

const REMINDER_CONFIG = {
  SUBJECT_PREFIX: '[Reminder] Pending Requests Summary',
  SENDER_NAME: 'Approval Dashboard',
  SENDER_EMAIL: 'Approval-noreply@ravago.com.ph',  // must be a verified "Send mail as" alias of the account that owns the trigger
  TIMEZONE: 'Asia/Manila',
  MAX_ROWS_IN_EMAIL: 50,          // keeps the email light; remainder is summarized
  TRIGGER_HOUR: 8                 // 8 AM in TIMEZONE
};

// ---------------------------------------------------------------------------
// DATA
// ---------------------------------------------------------------------------

/** All Pending rows (server-side Selector first, Apps Script filter as safety net). */
function fetchAllPendingRows_() {
  let rows;
  try {
    rows = appSheetFind_(TABLES.APPROVALS,
      'Filter(' + TABLES.APPROVALS + ', LOWER([Status])="pending")');
  } catch (err) {
    Logger.log('Pending selector failed, falling back to unfiltered Find: ' + err.message);
    rows = appSheetFind_(TABLES.APPROVALS);
  }
  return rows
    .map(normalizeApprovalRow_)
    .filter(function (r) { return (r.Status || '').toString().trim().toLowerCase() === 'pending'; })
    .sort(function (a, b) {
      // oldest first — the longest-waiting request is the most urgent
      return new Date(a.CreatedAt).getTime() - new Date(b.CreatedAt).getTime();
    });
}

function filterRowsForApprover_(rows, email) {
  const target = (email || '').toString().trim().toLowerCase();
  return rows.filter(function (r) {
    return (r.AprEmail || '').toString().trim().toLowerCase() === target;
  });
}

/** True if `address` is this account's own address or a verified Gmail "Send mail as" alias. */
function canSendAs_(address) {
  const target = (address || '').toLowerCase();
  if (!target) return false;
  try {
    if (Session.getEffectiveUser().getEmail().toLowerCase() === target) return true;
    return GmailApp.getAliases().some(function (a) { return a.toLowerCase() === target; });
  } catch (err) {
    Logger.log('Could not read Gmail aliases: ' + err.message);
    return false;
  }
}

function lookupApproverName_(email) {
  try {
    const target = email.trim().toLowerCase();
    const match = getManagersCached_().find(function (m) {
      return (m.Email || m.email || '').toString().trim().toLowerCase() === target;
    });
    return match ? (match.Name || match.name || '') : '';
  } catch (err) {
    return '';
  }
}

// ---------------------------------------------------------------------------
// PUBLIC FUNCTIONS
// ---------------------------------------------------------------------------

/**
 * Sends ONE reminder email to ONE recipient listing their pending requests.
 * Does nothing (and sends no email) if they have none.
 *
 * @param {string} recipientEmail
 * @param {Object[]} [preFetchedPending] Optional: the full pending list, so a
 *        batch run doesn't re-query AppSheet for every approver.
 * @return {{sent:boolean, count:number, reason?:string}}
 */
function sendScheduledReminderEmail(recipientEmail, preFetchedPending) {
  if (!recipientEmail) throw new Error('recipientEmail is required.');

  const allPending = preFetchedPending || fetchAllPendingRows_();
  const rows = filterRowsForApprover_(allPending, recipientEmail);

  if (!rows.length) {
    Logger.log('No pending requests for %s — reminder skipped.', recipientEmail);
    return { sent: false, count: 0, reason: 'no pending requests' };
  }

  const recipientName = lookupApproverName_(recipientEmail);
  const html = buildReminderHtml_(rows, recipientName);
  const plain = buildReminderPlainText_(rows, recipientName);

  const mailOptions = {
    htmlBody: html,
    name: REMINDER_CONFIG.SENDER_NAME
  };
  if (canSendAs_(REMINDER_CONFIG.SENDER_EMAIL)) {
    mailOptions.from = REMINDER_CONFIG.SENDER_EMAIL;
  } else {
    Logger.log('WARNING: "%s" is not a Send-mail-as alias of this account — sending from the default address instead.', REMINDER_CONFIG.SENDER_EMAIL);
  }

  GmailApp.sendEmail(
    recipientEmail,
    REMINDER_CONFIG.SUBJECT_PREFIX + ' (' + rows.length + ')',
    plain,              // plain-text fallback for clients that block HTML
    mailOptions
  );

  return { sent: true, count: rows.length };
}

/**
 * Trigger target. Fetches pending requests ONCE, groups them by approver,
 * and sends each approver their own summary. One failing recipient never
 * stops the others.
 */
function runScheduledReminders() {
  const allPending = fetchAllPendingRows_();

  const approvers = {};
  allPending.forEach(function (r) {
    const e = (r.AprEmail || '').toString().trim().toLowerCase();
    if (e) approvers[e] = true;
  });

  const emails = Object.keys(approvers);
  Logger.log('Scheduled reminder run: %s pending request(s) across %s approver(s).', allPending.length, emails.length);

  emails.forEach(function (email) {
    try {
      const res = sendScheduledReminderEmail(email, allPending);
      Logger.log('Reminder -> %s: %s', email, JSON.stringify(res));
    } catch (err) {
      Logger.log('Reminder FAILED for %s (continuing): %s', email, err.message);
    }
  });
}

// ---------------------------------------------------------------------------
// DETAILSJSON — safe parsing + rendering
// ---------------------------------------------------------------------------

/**
 * Never throws. Returns an array of {k, v} pairs.
 *   '' / null / '{}'            -> []
 *   '{"Item":"Laptop"}'         -> [{k:'Item', v:'Laptop'}]
 *   '["a","b"]'                 -> [{k:'Item 1', v:'a'}, {k:'Item 2', v:'b'}]
 *   malformed / plain text      -> [{k:'Details', v:<raw text>}]
 *   nested object values        -> compact JSON string
 */
function parseDetailsSafe_(raw) {
  if (raw === null || raw === undefined) return [];
  const text = String(raw).trim();
  if (!text) return [];

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    return [{ k: 'Details', v: text }];   // not valid JSON — show as-is
  }

  const stringify = function (val) {
    if (val === null || val === undefined || val === '') return '—';
    if (typeof val === 'object') {
      try { return JSON.stringify(val); } catch (e) { return String(val); }
    }
    return String(val);
  };

  if (Array.isArray(parsed)) {
    return parsed.map(function (item, i) { return { k: 'Item ' + (i + 1), v: stringify(item) }; });
  }
  if (parsed && typeof parsed === 'object') {
    return Object.keys(parsed).map(function (key) { return { k: key, v: stringify(parsed[key]) }; });
  }
  return [{ k: 'Details', v: stringify(parsed) }];   // bare string/number/boolean
}

/**
 * Returns the key/value rows as a full-width, stacked block (no right-hand column).
 * Returns '' when there is nothing to show so the card simply omits the section.
 */
function renderDetailsHtml_(raw) {
  let pairs;
  try {
    pairs = parseDetailsSafe_(raw);
  } catch (err) {
    pairs = [];
  }
  if (!pairs.length) return '';

  const rows = pairs.map(function (p) {
    return '<tr>' +
      '<td class="stack" valign="top" width="38%" style="padding:6px 12px 6px 0;font-size:15px;line-height:1.5;color:#6b7079;font-weight:600;">' +
        reminderEscape_(p.k) + '</td>' +
      '<td class="stack" valign="top" style="padding:6px 0;font-size:16px;line-height:1.5;color:#14161a;word-break:break-word;">' +
        reminderEscape_(p.v) + '</td>' +
    '</tr>';
  }).join('');

  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">' + rows + '</table>';
}

// ---------------------------------------------------------------------------
// FORMATTING HELPERS
// ---------------------------------------------------------------------------

function formatCreatedAtForEmail_(value) {
  if (!value) return '&mdash;';
  const d = new Date(value);
  if (isNaN(d.getTime())) return reminderEscape_(String(value));  // unparseable — show raw
  return Utilities.formatDate(d, REMINDER_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
}

function reminderEscape_(str) {
  return String(str === null || str === undefined ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ---------------------------------------------------------------------------
// EMAIL BODY
// ---------------------------------------------------------------------------

function buildRequestCardHtml_(r) {
  const FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
  const label = 'font-size:14px;line-height:1.4;font-weight:700;color:#6b7079;text-transform:uppercase;letter-spacing:0.03em;margin:0 0 4px;';
  const value = 'font-size:16px;line-height:1.5;color:#14161a;margin:0;word-break:break-word;';
  const dash = '<span style="color:#9a9ea6;">&mdash;</span>';

  const detailsHtml = renderDetailsHtml_(r.DetailsJSON);

  return '' +
  '<tr><td class="px" style="padding:0 24px 16px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #e6e7eb;border-radius:10px;border-collapse:separate;background:#ffffff;">' +

      // Header: Request ID + App Source
      '<tr><td colspan="2" bgcolor="#f4f5f7" style="padding:14px 16px;background:#f4f5f7;border-bottom:1px solid #e6e7eb;border-radius:10px 10px 0 0;">' +
        '<span style="font-family:Consolas,Menlo,monospace;font-size:18px;font-weight:700;color:#14161a;">' + reminderEscape_(r.RequestID) + '</span>' +
        '<span style="display:inline-block;margin-left:10px;background:#e0e7ff;color:#2447c4;font-size:14px;font-weight:700;padding:3px 11px;border-radius:999px;">' + reminderEscape_(r.AppSource) + '</span>' +
      '</td></tr>' +

      // Requester + Created At (side by side on desktop, stacked on mobile)
      '<tr>' +
        '<td class="stack" width="50%" valign="top" style="padding:16px 16px 4px;font-family:' + FONT + ';">' +
          '<p style="' + label + '">Requester Name</p><p style="' + value + '">' + reminderEscape_(r.RequesterName || '') + '</p>' +
        '</td>' +
        '<td class="stack" width="50%" valign="top" style="padding:16px 16px 4px;font-family:' + FONT + ';">' +
          '<p style="' + label + '">Created At</p><p style="' + value + '">' + formatCreatedAtForEmail_(r.CreatedAt) + '</p>' +
        '</td>' +
      '</tr>' +

      // Reason — full width
      '<tr><td colspan="2" style="padding:12px 16px 4px;font-family:' + FONT + ';">' +
        '<p style="' + label + '">Reason</p><p style="' + value + '">' + (r.Description ? reminderEscape_(r.Description) : dash) + '</p>' +
      '</td></tr>' +

      // Details — full width, directly underneath this request
      (detailsHtml
        ? '<tr><td colspan="2" style="padding:12px 16px 16px;font-family:' + FONT + ';">' +
            '<p style="' + label + '">Details</p>' +
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb;border:1px solid #e6e7eb;border-left:4px solid #2f5cf0;border-radius:6px;border-collapse:separate;">' +
              '<tr><td style="padding:8px 14px;">' + detailsHtml + '</td></tr>' +
            '</table>' +
          '</td></tr>'
        : '<tr><td colspan="2" style="padding:0 0 4px;font-size:1px;line-height:1px;">&nbsp;</td></tr>') +

    '</table>' +
  '</td></tr>';
}

function buildReminderHtml_(rows, recipientName) {
  const shown = rows.slice(0, REMINDER_CONFIG.MAX_ROWS_IN_EMAIL);
  const hidden = rows.length - shown.length;

  let dashboardUrl = '';
  try { dashboardUrl = ScriptApp.getService().getUrl(); } catch (err) { /* not deployed yet */ }

  const cards = shown.map(buildRequestCardHtml_).join('');
  const greeting = recipientName ? 'Hi ' + reminderEscape_(recipientName) + ',' : 'Hello,';
  const timestamp = Utilities.formatDate(new Date(), REMINDER_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');

  return '' +
  '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
  '<style>' +
    '@media only screen and (max-width:600px){' +
      '.px{padding-left:12px !important;padding-right:12px !important;}' +
      '.stack{display:block !important;width:100% !important;box-sizing:border-box !important;}' +
      '.outer{padding:12px 0 !important;}' +
    '}' +
  '</style></head>' +
  '<body style="margin:0;padding:0;background:#f4f5f7;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f5f7;"><tr><td class="outer" align="center" style="padding:24px 12px;">' +

    '<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:640px;background:#ffffff;border:1px solid #e6e7eb;border-radius:10px;border-collapse:separate;">' +

      // ---- Banner ----
      '<tr><td class="px" bgcolor="#2f5cf0" style="background:#2f5cf0;padding:22px 24px;border-radius:10px 10px 0 0;">' +
        '<div style="font-size:14px;font-weight:700;letter-spacing:0.06em;color:#dbe4ff;text-transform:uppercase;">&#9200; Scheduled Reminder &middot; Automated</div>' +
        '<div style="font-size:24px;line-height:1.3;font-weight:700;color:#ffffff;margin-top:6px;">Pending Requests Summary</div>' +
      '</td></tr>' +

      // ---- Intro ----
      '<tr><td class="px" style="padding:24px 24px 16px;font-size:16px;line-height:1.6;color:#14161a;">' +
        '<p style="margin:0 0 10px;">' + greeting + '</p>' +
        '<p style="margin:0;">This is an <b>automated scheduled reminder</b>. You currently have ' +
          '<b style="color:#d97706;">' + rows.length + ' pending request' + (rows.length === 1 ? '' : 's') + '</b> awaiting your decision (oldest first).</p>' +
      '</td></tr>' +

      // ---- Request cards ----
      cards +

      (hidden > 0
        ? '<tr><td class="px" style="padding:0 24px 16px;font-size:15px;line-height:1.5;color:#6b7079;">+ ' + hidden + ' more pending request' + (hidden === 1 ? '' : 's') + ' not shown. Open the dashboard to see them all.</td></tr>'
        : '') +

      // ---- CTA ----
      (dashboardUrl
        ? '<tr><td class="px" align="left" style="padding:8px 24px 24px;">' +
            '<a href="' + reminderEscape_(dashboardUrl) + '" style="display:inline-block;background:#2f5cf0;color:#ffffff;text-decoration:none;font-size:16px;font-weight:700;padding:14px 26px;border-radius:8px;">Open Approval Dashboard</a>' +
          '</td></tr>'
        : '') +

      // ---- Footer ----
      '<tr><td class="px" style="padding:18px 24px 24px;font-size:14px;line-height:1.5;color:#6b7079;border-top:1px solid #e6e7eb;">' +
        'Sent automatically on ' + timestamp + ' (' + reminderEscape_(REMINDER_CONFIG.TIMEZONE) + '). This is a scheduled reminder &mdash; please do not reply to this message.' +
      '</td></tr>' +

    '</table>' +
  '</td></tr></table></body></html>';
}

/** Plain-text twin of the email (shown by clients that don't render HTML). */
function buildReminderPlainText_(rows, recipientName) {
  const lines = [
    'SCHEDULED REMINDER (automated) - Pending Requests Summary',
    '',
    (recipientName ? 'Hi ' + recipientName + ',' : 'Hello,'),
    'You have ' + rows.length + ' pending request(s) awaiting your decision:',
    ''
  ];
  rows.slice(0, REMINDER_CONFIG.MAX_ROWS_IN_EMAIL).forEach(function (r) {
    const details = parseDetailsSafe_(r.DetailsJSON)
      .map(function (p) { return p.k + ': ' + p.v; }).join('; ');
    lines.push('- ' + r.RequestID + ' | ' + r.AppSource + ' | ' + r.RequesterName + ' | ' +
      formatCreatedAtForEmail_(r.CreatedAt).replace(/&mdash;/g, '-') +
      (r.Description ? ' | Reason: ' + r.Description : '') +
      (details ? ' | ' + details : ''));
  });
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// TRIGGER SETUP — run ONE of these manually, ONCE, from the editor (▶)
// ---------------------------------------------------------------------------

function deleteReminderTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runScheduledReminders') ScriptApp.deleteTrigger(t);
  });
}

/** Every day at TRIGGER_HOUR (default 8 AM Manila). */
function createDailyReminderTrigger() {
  deleteReminderTriggers();   // prevents duplicate triggers => duplicate emails
  ScriptApp.newTrigger('runScheduledReminders')
    .timeBased()
    .everyDays(1)
    .atHour(REMINDER_CONFIG.TRIGGER_HOUR)
    .inTimezone(REMINDER_CONFIG.TIMEZONE)
    .create();
  Logger.log('Daily reminder trigger created.');
}

/** Every Monday at TRIGGER_HOUR. */
function createWeeklyReminderTrigger() {
  deleteReminderTriggers();
  ScriptApp.newTrigger('runScheduledReminders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(REMINDER_CONFIG.TRIGGER_HOUR)
    .inTimezone(REMINDER_CONFIG.TIMEZONE)
    .create();
  Logger.log('Weekly reminder trigger created.');
}

// ---------------------------------------------------------------------------
// TEST HELPERS
// ---------------------------------------------------------------------------

/** Replace with your own address, then ▶ run. Sends only if you have pending items. */
function testSendScheduledReminderEmail() {
  Logger.log(JSON.stringify(sendScheduledReminderEmail('kennery.villacaol@ravago.com.ph')));
}

/** Verifies the JSON parser against good, empty, and malformed input — no email sent. */
function testParseDetailsSafe() {
  [
    '{"Item":"Laptop","Urgency":"High"}',
    '', null, '{}', '["a","b"]',
    '{"broken": ', 'just plain text',
    '{"Nested":{"a":1},"Empty":""}'
  ].forEach(function (s) {
    Logger.log('%s  ->  %s', s, JSON.stringify(parseDetailsSafe_(s)));
  });
}
