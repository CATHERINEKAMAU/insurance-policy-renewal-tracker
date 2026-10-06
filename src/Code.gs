// ══════════════════════════════════════════════════════════════════
//  INSURANCE POLICY RENEWAL TRACKER PRO — Apps Script v4.1
//  Your Brokerage Name
//  Paste into Extensions → Apps Script → Save → Run setupDailyTrigger()
//
//  v4.1 fixes vs v4.0:
//   • Dates could render a day off if the Apps Script PROJECT timezone
//     (Extensions → Apps Script → Project Settings) didn't match the
//     spreadsheet's timezone. canonicalDate()/getTodayCanonical() now
//     build dates through Utilities.parseDate() against the SHEET's
//     own timezone explicitly, so they can never drift regardless of
//     what the project setting is.
//   • Fonts could fall back to Calibri/Arial after a paste or autofill,
//     because nothing re-asserted "Inter" on edited cells. onEditHandler
//     now calls enforceFont() on every edit, and a new one-time
//     fixAllFormatting() sweeps the whole register to bring any
//     existing drift back into line.
// ══════════════════════════════════════════════════════════════════

const SHEET_NAME   = 'Policy Register';
const LOG_SHEET     = 'Activity Log';
const CAL_NAME      = 'Insurance Renewal Tracker';
const FIRST_ROW      = 5;
const LAST_ROW       = 304;

// Where the Daily Digest and Weekly Summary land. Leave blank to use
// whichever Google account runs the script.
const TEAM_INBOX = '';

// House font — kept identical across every data cell in the register.
const DATA_FONT = 'Inter';
const DATA_FONT_SIZE = 10;

// ── TRIGGER TIMES — everything defaults to 8am. Change any of these
//    (e.g. to a near-future hour while testing) then re-run
//    setupDailyTrigger() to apply. WEEKLY_SUMMARY also runs Mondays.
const TRIGGER_TIMES = {
  CALENDAR_SYNC_HOUR:   8,
  DAILY_DIGEST_HOUR:    8,
  WEEKLY_SUMMARY_HOUR:  8,
  WEEKLY_SUMMARY_MINUTE: 0,
  MILESTONE_CHECK_HOUR: 8
};

// Column map (Policy Register)
const COL_ID         = 2;   // B
const COL_CLIENT     = 3;   // C
const COL_CLASS      = 4;   // D
const COL_DESC       = 5;   // E
const COL_DUE        = 6;   // F
const COL_STATUS     = 7;   // G
const COL_OVERDUE    = 8;   // H
const COL_NOTES      = 9;   // I
const COL_EMAIL      = 10;  // J
const COL_MILESTONE  = 16;  // P — hidden "Last Milestone Sent" tracker

// Client-facing renewal nudges fire once, on the day a policy crosses
// each of these thresholds (days remaining). 0 = due today.
const MILESTONES = [60, 30, 14, 7, 3, 1, 0];

// Brand colours
const BRAND = {
  burgundy: '#6F1D1B', gold: '#BB9457', chocolate: '#432818',
  orange: '#99582A', cream: '#FFE6A7', paper: '#FBF3E4', line: '#EFE3D0',
  muted: '#8A7A6D'
};

// ── SETUP: run this once (and again any time you change TRIGGER_TIMES) ─
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger('syncAllPoliciesToCalendar').timeBased().atHour(TRIGGER_TIMES.CALENDAR_SYNC_HOUR).everyDays(1).create();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(TRIGGER_TIMES.DAILY_DIGEST_HOUR).everyDays(1).create();
  ScriptApp.newTrigger('sendWeeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(TRIGGER_TIMES.WEEKLY_SUMMARY_HOUR).nearMinute(TRIGGER_TIMES.WEEKLY_SUMMARY_MINUTE).create();
  ScriptApp.newTrigger('checkMilestones').timeBased().atHour(TRIGGER_TIMES.MILESTONE_CHECK_HOUR).everyDays(1).create();

  // Always resolve the SAME calendar instead of risking a duplicate
  getOrCreateCalendar();

  // Push everything currently on the sheet straight away, so events
  // show up on the calendar immediately instead of waiting for the trigger.
  const result = syncAllPoliciesToCalendarSilent();

  // Bring dates and fonts into line the moment setup runs.
  normalizeAllDueDatesSilent();
  fixAllFormattingSilent();

  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    '✓ Calendar sync — daily ' + fmtHour(TRIGGER_TIMES.CALENDAR_SYNC_HOUR) + ', plus instantly as you edit (' + result.synced + ' polic' + (result.synced === 1 ? 'y' : 'ies') + ' pushed just now)\n' +
    '✓ Daily Digest — daily ' + fmtHour(TRIGGER_TIMES.DAILY_DIGEST_HOUR) + ', to your team inbox\n' +
    '✓ Weekly Summary — Monday ' + fmtHour(TRIGGER_TIMES.WEEKLY_SUMMARY_HOUR) + ', to your team inbox\n' +
    '✓ Milestone Alert — checked daily ' + fmtHour(TRIGGER_TIMES.MILESTONE_CHECK_HOUR) + ', sent to clients only the day a policy crosses 60/30/14/7/3/1/0 days out\n' +
    '✓ Dates re-normalized and fonts reset to ' + DATA_FONT + ' across the whole register\n\n' +
    'All four default to 8am — change TRIGGER_TIMES at the top of the script any time (handy for testing) and re-run this function to apply.\n\n' +
    'IMPORTANT: also check Extensions → Apps Script → Project Settings → Time zone matches this sheet\'s timezone (' + getTZ() + '). They should always match.'
  );
}

function fmtHour(h) {
  const period = h >= 12 ? 'pm' : 'am';
  const h12 = ((h % 12) === 0) ? 12 : (h % 12);
  return h12 + ':00' + period;
}

// ── MANUAL TEST HELPERS — run any of these directly from the Apps
//    Script editor (▶ button) to preview an email without waiting
//    for the trigger to fire. ───────────────────────────────────
function testDailyDigest() { sendDailyDigest(); }
function testWeeklySummary() { sendWeeklySummary(); }
function testMilestoneCheck() { checkMilestones(); }

// ── RUN THIS TO VERIFY THE FIX / FORCE-FIX EVERY EXISTING ROW ─────
//    Select this function in the dropdown next to ▶ and run it.
//    It re-normalizes every due date already on the sheet (handy if
//    some were entered before this version was installed) and pops
//    up a confirmation so you can see it actually ran.
function normalizeAllDueDates() {
  const fixed = normalizeAllDueDatesSilent();
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert(
      '✓ Checked ' + fixed + ' polic' + (fixed === 1 ? 'y' : 'ies') + '.\n\n' +
      'Every due date has been re-read through this sheet\'s timezone (' + getTZ() + ') and rewritten as d mmm yyyy.\n\n' +
      'If a date still looks wrong after this, it is not a timezone issue — reply with the exact client, the date you typed, and what it shows now.'
    );
  }
}

function normalizeAllDueDatesSilent() {
  let fixed = 0;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
    const client = sheet.getRange(row, COL_CLIENT).getValue();
    if (!client || client.toString().trim() === '') continue;
    normalizeDueDate(row);
    fixed++;
  }
  return fixed;
}

// ── RUN THIS TO RESET EVERY DATA CELL BACK TO THE HOUSE FONT ─────
//    Fixes drift caused by pasting from elsewhere, autofill, or
//    Sheets applying its own default on a newly-touched cell.
function fixAllFormatting() {
  const fixed = fixAllFormattingSilent();
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert('✓ Reset formatting on ' + fixed + ' row(s). Every cell B:J is now ' + DATA_FONT + ' ' + DATA_FONT_SIZE + 'pt.');
  }
}

function fixAllFormattingSilent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  let fixed = 0;
  for (let row = FIRST_ROW; row <= LAST_ROW; row++) {
    const client = sheet.getRange(row, COL_CLIENT).getValue();
    if (!client || client.toString().trim() === '') continue;
    sheet.getRange(row, COL_ID, 1, COL_EMAIL - COL_ID + 1)
      .setFontFamily(DATA_FONT)
      .setFontSize(DATA_FONT_SIZE);
    fixed++;
  }
  return fixed;
}

// ── ON EDIT: keep calendar + date formatting + fonts in sync ─────
function onEditHandler(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < FIRST_ROW) return;

  // Re-assert the house font on whatever was just typed/pasted, so
  // formatting never has a chance to drift.
  if (col >= COL_ID && col <= COL_EMAIL) {
    enforceFont(row, col);
  }

  if (col === COL_DUE) {
    normalizeDueDate(row);
  }

  // Any column that changes what the calendar event shows should re-sync.
  const watchedCols = [COL_CLIENT, COL_CLASS, COL_DESC, COL_DUE, COL_STATUS, COL_NOTES];
  if (watchedCols.includes(col)) {
    Utilities.sleep(400);
    syncPolicyToCalendar(row);
  }
}

function enforceFont(row, col) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.getRange(row, col).setFontFamily(DATA_FONT).setFontSize(DATA_FONT_SIZE);
}

// ── DATE HARMONIZER — whatever a person types (numbers, words,
//    dd/mm/yyyy, "12 Aug 2026", "August 12, 2026", "8th August 2026"...)
//    becomes the same real date, in the same display format — and
//    never drifts by a day, no matter what timezone the SCRIPT PROJECT
//    happens to be set to.
//
//    Apps Script's underlying clock is always UTC, and the Apps Script
//    project has its own default timezone separate from the sheet's
//    (File → Settings on the sheet). Building a Date with the plain
//    `new Date(y, m-1, d)` constructor uses the PROJECT's timezone —
//    if that ever doesn't match the sheet's timezone, the date can
//    land a day off once Sheets re-renders it. The fix: every date we
//    build gets constructed with Utilities.parseDate() against the
//    SHEET's timezone explicitly, so it's immune to the project
//    setting entirely.
function getTZ() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

function extractCalendarDate(raw) {
  const tz = getTZ();

  if (raw instanceof Date && !isNaN(raw.getTime())) {
    const parts = Utilities.formatDate(raw, tz, 'yyyy-MM-dd').split('-');
    return { y: parseInt(parts[0], 10), m: parseInt(parts[1], 10), d: parseInt(parts[2], 10) };
  }

  if (typeof raw === 'string' && raw.trim() !== '') {
    const cleaned = raw.trim();

    // dd/mm/yyyy or dd-mm-yyyy first — avoids JS misreading it as US mm/dd
    const slash = cleaned.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (slash) {
      const d = parseInt(slash[1], 10);
      const m = parseInt(slash[2], 10);
      const y = slash[3].length === 2 ? 2000 + parseInt(slash[3], 10) : parseInt(slash[3], 10);
      return { y: y, m: m, d: d };
    }

    // Word dates ("8th August 2026", "August 8, 2026") — strip ordinal
    // suffixes so "8th" doesn't confuse the parser, let JS read it, then
    // re-extract the calendar day through the spreadsheet's timezone
    // instead of trusting the parsed Date's raw UTC hours.
    const noOrdinals = cleaned.replace(/(\d+)(st|nd|rd|th)\b/gi, '$1');
    const parsedDate = new Date(noOrdinals);
    if (!isNaN(parsedDate.getTime())) {
      const parts = Utilities.formatDate(parsedDate, tz, 'yyyy-MM-dd').split('-');
      return { y: parseInt(parts[0], 10), m: parseInt(parts[1], 10), d: parseInt(parts[2], 10) };
    }
  }

  return null;
}

// A clean Date object for a given calendar day, built explicitly against
// the SHEET's timezone — never the script project's default timezone.
// This is the core fix: it makes "day drift" structurally impossible.
function canonicalDate(y, m, d) {
  const tz = getTZ();
  const pad = n => String(n).padStart(2, '0');
  const iso = y + '-' + pad(m) + '-' + pad(d) + 'T00:00:00';
  return Utilities.parseDate(iso, tz, "yyyy-MM-dd'T'HH:mm:ss");
}

// "Today" as a clean calendar-day Date, read through the spreadsheet's
// own timezone (safer than new Date() + setHours near a midnight boundary).
function getTodayCanonical() {
  const ymd = extractCalendarDate(new Date());
  return canonicalDate(ymd.y, ymd.m, ymd.d);
}

function normalizeDueDate(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const cell  = sheet.getRange(row, COL_DUE);
  const raw   = cell.getValue();
  const ymd   = extractCalendarDate(raw);

  if (ymd) {
    cell.setValue(canonicalDate(ymd.y, ymd.m, ymd.d));
  }
  cell.setNumberFormat('d mmm yyyy');
  cell.setFontFamily(DATA_FONT).setFontSize(DATA_FONT_SIZE);
}

// ── SYNC A SINGLE POLICY ROW TO CALENDAR ─────────────────────────
function syncPolicyToCalendar(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const policyId = data[COL_ID - 1];
  const client    = data[COL_CLIENT - 1];
  const cls       = data[COL_CLASS - 1];
  const desc      = data[COL_DESC - 1];
  const due       = data[COL_DUE - 1];
  const status    = data[COL_STATUS - 1];
  const notes     = data[COL_NOTES - 1];

  if (!client || client.toString().trim() === '') return;

  const cal = getOrCreateCalendar();
  const title = 'Policy Renewal – ' + client;

  if (!due) return;
  const dueYmd = extractCalendarDate(due);
  if (!dueYmd) return;
  const dueDate = canonicalDate(dueYmd.y, dueYmd.m, dueYmd.d);
  dueDate.setHours(9, 0, 0, 0);
  const existing = cal.getEventsForDay(dueDate).filter(ev => ev.getTitle() === title);

  if (status === 'Renewed') {
    existing.forEach(ev => ev.deleteEvent());
    logActivity(policyId, 'Calendar event removed', 'Policy marked Renewed');
    return;
  }

  const description = [
    'Client: ' + client,
    'Insurance Class: ' + cls,
    'Description: ' + (desc || '-'),
    'Due Date: ' + Utilities.formatDate(dueDate, getTZ(), 'd MMMM yyyy'),
    'Notes: ' + (notes || '-'),
    '',
    'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].join('\n');

  const today = getTodayCanonical();
  const daysLeft = Math.round((dueDate - today) / 86400000);
  const eventColor = daysLeft < 0 ? CalendarApp.EventColor.RED
                    : (daysLeft <= 7 ? CalendarApp.EventColor.ORANGE
                    : CalendarApp.EventColor.YELLOW);

  if (existing.length === 0) {
    const end = new Date(dueDate); end.setHours(10, 0, 0, 0);
    const ev = cal.createEvent(title, dueDate, end, { description: description });
    ev.addPopupReminder(60 * 24);
    ev.addPopupReminder(60);
    ev.setColor(eventColor);
    logActivity(policyId, 'Calendar event created', 'Due ' + Utilities.formatDate(dueDate, getTZ(), 'd MMM yyyy'));
  } else {
    existing[0].setDescription(description);
    existing[0].setColor(eventColor);
    if (existing[0].getStartTime().getTime() !== dueDate.getTime()) {
      const end = new Date(dueDate); end.setHours(10, 0, 0, 0);
      existing[0].setTime(dueDate, end);
    }
    logActivity(policyId, 'Calendar event updated', 'Due date or details changed');
  }
}

function getOrCreateCalendar() {
  const calList = CalendarApp.getCalendarsByName(CAL_NAME);
  return calList.length > 0 ? calList[0] : CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.CHERRY_BLOSSOM });
}

// ── BULK SYNC: run manually any time to reconcile everything ─────
function syncAllPoliciesToCalendar() {
  const result = syncAllPoliciesToCalendarSilent();
  const ui = SpreadsheetApp.getUi();
  if (ui) ui.alert('✓ Sync complete. ' + result.synced + ' polic' + (result.synced === 1 ? 'y' : 'ies') + ' checked against the calendar.');
}

function syncAllPoliciesToCalendarSilent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getRange(FIRST_ROW, 1, LAST_ROW - FIRST_ROW + 1, 10).getValues();
  let synced = 0;
  for (let i = 0; i < data.length; i++) {
    const client = data[i][COL_CLIENT - 1];
    if (!client || client.toString().trim() === '') continue;
    syncPolicyToCalendar(FIRST_ROW + i);
    synced++;
  }
  return { synced: synced };
}

// ── Read every open (not renewed) policy once, shared by digest,
//    weekly summary, and milestone check. ────────────────────────
function getOpenPolicies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getRange(FIRST_ROW, 1, LAST_ROW - FIRST_ROW + 1, 16).getValues();
  const today = getTodayCanonical();
  const rows = [];

  data.forEach((r, i) => {
    const client = r[COL_CLIENT - 1];
    if (!client || client.toString().trim() === '') return;
    if (r[COL_STATUS - 1] === 'Renewed') return;
    const due = r[COL_DUE - 1];
    if (!due) return;
    const dueYmd = extractCalendarDate(due);
    if (!dueYmd) return;
    const dueDate = canonicalDate(dueYmd.y, dueYmd.m, dueYmd.d);
    const daysLeft = Math.round((dueDate - today) / 86400000);
    rows.push({
      row: FIRST_ROW + i,
      id: r[COL_ID - 1],
      client: client,
      cls: r[COL_CLASS - 1],
      desc: r[COL_DESC - 1],
      due: dueDate,
      daysLeft: daysLeft,
      email: r[COL_EMAIL - 1],
      lastMilestone: r[COL_MILESTONE - 1]
    });
  });
  return rows;
}

function resolveTeamInbox() {
  return (TEAM_INBOX && TEAM_INBOX.trim() !== '') ? TEAM_INBOX.trim() : Session.getActiveUser().getEmail();
}

// ── DAILY DIGEST — short glance, team inbox ───────────────────────
function sendDailyDigest() {
  const policies = getOpenPolicies();
  const dueToday = policies.filter(p => p.daysLeft === 0);
  const overdue  = policies.filter(p => p.daysLeft < 0);
  const upcoming = policies.filter(p => p.daysLeft > 0).sort((a, b) => a.daysLeft - b.daysLeft).slice(0, 5);

  const to = resolveTeamInbox();
  let body = statChips(dueToday.length, overdue.length, upcoming.length);
  if (dueToday.length) body += sectionCard('Due today', dueToday);
  if (upcoming.length) body += sectionCard('Next 5 coming up', upcoming);
  if (!dueToday.length && !upcoming.length) body += emptyState('Nothing due today or in the days just ahead. 🎉');

  MailApp.sendEmail({
    to: to,
    subject: '☀️ Daily Digest — ' + dueToday.length + ' due today, ' + overdue.length + ' overdue',
    htmlBody: emailShell({
      topRight: dueToday.length + ' due today · ' + overdue.length + ' overdue',
      emoji: '☀️', heading: 'Daily Digest',
      intro: 'Here is this morning\'s snapshot.',
      bodyHtml: body, recipientEmail: to,
      gradientFrom: BRAND.chocolate, gradientTo: BRAND.burgundy,
      ctaLabel: 'Open Policy Register →', showCta: true
    })
  });
  logActivity('', 'Daily digest sent', dueToday.length + ' due today, ' + overdue.length + ' overdue, ' + upcoming.length + ' upcoming shown');
}

// ── WEEKLY SUMMARY — fuller list, team inbox ──────────────────────
function sendWeeklySummary() {
  const policies = getOpenPolicies();
  const next30 = policies.filter(p => p.daysLeft >= 0 && p.daysLeft <= 30).sort((a, b) => a.daysLeft - b.daysLeft);
  const overdue = policies.filter(p => p.daysLeft < 0).sort((a, b) => a.daysLeft - b.daysLeft);

  const to = resolveTeamInbox();
  let body = statChips(policies.filter(p => p.daysLeft === 0).length, overdue.length, next30.length);
  if (next30.length) body += sectionCard('Due in the next 30 days', next30);
  if (overdue.length) body += sectionCard('Full overdue list', overdue);
  if (!next30.length && !overdue.length) body += emptyState('Nothing due in the next 30 days, and nothing overdue. 🎉');

  MailApp.sendEmail({
    to: to,
    subject: '🗓️ Weekly Summary — ' + next30.length + ' due within 30 days, ' + overdue.length + ' overdue',
    htmlBody: emailShell({
      topRight: next30.length + ' upcoming · ' + overdue.length + ' overdue',
      emoji: '🗓️', heading: 'Weekly Summary',
      intro: 'Your fuller view for the week ahead.',
      bodyHtml: body, recipientEmail: to,
      gradientFrom: BRAND.burgundy, gradientTo: BRAND.orange,
      ctaLabel: 'Open Policy Register →', showCta: true
    })
  });
  logActivity('', 'Weekly summary sent', next30.length + ' due within 30 days, ' + overdue.length + ' overdue');
}

// ── MILESTONE ALERT — client-facing, once per policy per threshold
//    (60/30/14/7/3/1/0 days out) ──────────────────────────────────
function checkMilestones() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const policies = getOpenPolicies();

  policies.forEach(p => {
    if (!MILESTONES.includes(p.daysLeft)) return;
    if (p.lastMilestone === p.daysLeft) return; // already sent for this threshold
    if (!p.email) return;

    const badgeText = p.daysLeft === 0 ? 'due today' : p.daysLeft + ' day(s) away';
    const intro = 'A quick reminder that your policy below is ' + badgeText + '.';
    const body = sectionCard('', [p]);

    MailApp.sendEmail({
      to: p.email,
      subject: 'Renewal reminder — ' + p.client + ' (' + badgeText + ')',
      htmlBody: emailShell({
        topRight: 'Policy ' + (p.id || ''),
        emoji: '📌', heading: 'Renewal Reminder',
        intro: intro, bodyHtml: body, recipientEmail: p.email,
        gradientFrom: BRAND.gold, gradientTo: BRAND.orange,
        showCta: false // never hand a client a link to the internal sheet
      })
    });

    sheet.getRange(p.row, COL_MILESTONE).setValue(p.daysLeft);
    logActivity(p.id, 'Milestone alert sent', badgeText + ' (' + p.daysLeft + ' day threshold)');
  });
}

// ── NAME HELPER — turns "jane.doe@yourbrokerage.com" into "Jane" ─────
function nameFromEmail(email) {
  if (!email) return 'there';
  const local = email.split('@')[0];
  const first = local.split(/[._+-]/)[0];
  if (!first) return 'there';
  return first.charAt(0).toUpperCase() + first.slice(1).toLowerCase();
}

// ══════════════════════════════════════════════════════════════════
//  EMAIL DESIGN SYSTEM — dark gradient header + pill CTA + card
//  sections with label:value rows (inspired by modern lead-notification
//  templates). Built entirely with tables so it survives Gmail/Outlook,
//  stays under 600px, and never forces horizontal scroll on a phone.
// ══════════════════════════════════════════════════════════════════

function emailShell(opts) {
  const greetingName = nameFromEmail(opts.recipientEmail);
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const dateStr = Utilities.formatDate(new Date(), getTZ(), 'd MMMM yyyy');

  const ctaBlock = opts.showCta
    ? '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:18px auto 0 auto;"><tr><td style="border-radius:999px;background-color:' + BRAND.gold + ';background-image:linear-gradient(135deg,' + BRAND.gold + ',' + BRAND.orange + ');">' +
        '<a href="' + sheetUrl + '" style="display:inline-block;padding:13px 30px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;letter-spacing:0.5px;color:#FFFFFF;text-decoration:none;text-transform:uppercase;">' + opts.ctaLabel + '</a>' +
      '</td></tr></table>'
    : '';

  const replyNote = !opts.showCta
    ? '<p style="font-family:Arial,sans-serif;font-size:12px;color:' + BRAND.muted + ';text-align:center;margin:16px 0 0 0;">Questions about this policy? Just reply to this email.</p>'
    : '';

  return '<div style="background:#F4EFE6;padding:16px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;margin:0 auto;background:#FFFFFF;border-radius:14px;overflow:hidden;border:1px solid ' + BRAND.line + ';">' +

      // ── slim top bar ──
      '<tr><td style="background-color:' + BRAND.chocolate + ';padding:10px 20px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
          '<td style="font-family:Arial,sans-serif;font-size:10px;letter-spacing:1px;color:' + BRAND.cream + ';text-transform:uppercase;">Your Brokerage Name</td>' +
          '<td style="text-align:right;font-family:Arial,sans-serif;font-size:10px;color:' + BRAND.cream + ';white-space:nowrap;">' + escapeHtml(opts.topRight || dateStr) + '</td>' +
        '</tr></table>' +
      '</td></tr>' +

      // ── gradient hero ──
      '<tr><td style="background-color:' + opts.gradientFrom + ';background-image:linear-gradient(135deg,' + opts.gradientFrom + ',' + opts.gradientTo + ');padding:34px 24px;text-align:center;">' +
        '<div style="width:46px;height:46px;line-height:46px;border-radius:50%;background:rgba(255,255,255,0.18);color:#FFFFFF;font-family:Arial,sans-serif;font-weight:bold;font-size:16px;text-align:center;margin:0 auto 14px auto;">BC</div>' +
        '<div style="font-family:Arial,sans-serif;font-size:22px;font-weight:bold;color:#FFFFFF;">' + opts.emoji + ' ' + escapeHtml(opts.heading) + '</div>' +
        '<div style="font-family:Arial,sans-serif;font-size:13px;color:' + BRAND.cream + ';margin-top:8px;max-width:420px;margin-left:auto;margin-right:auto;">Hi ' + escapeHtml(greetingName) + ' — ' + opts.intro + '</div>' +
        ctaBlock +
      '</td></tr>' +

      // ── body ──
      '<tr><td style="padding:20px;">' +
        opts.bodyHtml +
        replyNote +
      '</td></tr>' +

      // ── footer ──
      '<tr><td style="background:' + BRAND.paper + ';padding:16px 20px;border-top:1px solid ' + BRAND.line + ';">' +
        '<div style="font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:' + BRAND.chocolate + ';">Insurance Policy Renewal Tracker PRO</div>' +
        '<div style="font-family:Arial,sans-serif;font-size:11px;color:' + BRAND.muted + ';">Automated Renewal Notification · Powered by Google Workspace</div>' +
      '</td></tr>' +
    '</table>' +
  '</div>';
}

// Three little stat pills across the top of team emails.
function statChips(dueTodayCount, overdueCount, otherCount) {
  const chip = (label, value, color) =>
    '<td style="width:33.3%;padding:4px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + BRAND.paper + ';border:1px solid ' + BRAND.line + ';border-radius:10px;"><tr><td style="padding:12px 6px;text-align:center;">' +
        '<div style="font-family:Arial,sans-serif;font-size:22px;font-weight:bold;color:' + color + ';">' + value + '</div>' +
        '<div style="font-family:Arial,sans-serif;font-size:9px;color:' + BRAND.muted + ';text-transform:uppercase;letter-spacing:0.5px;">' + label + '</div>' +
      '</td></tr></table>' +
    '</td>';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;margin-bottom:6px;"><tr>' +
    chip('Due today', dueTodayCount, BRAND.chocolate) +
    chip('Overdue', overdueCount, BRAND.burgundy) +
    chip('Coming up', otherCount, BRAND.gold) +
    '</tr></table>';
}

function emptyState(message) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;margin-top:12px;"><tr><td style="padding:18px;text-align:center;background:' + BRAND.paper + ';border:1px dashed ' + BRAND.line + ';border-radius:10px;font-family:Arial,sans-serif;font-size:13px;color:' + BRAND.muted + ';">' + escapeHtml(message) + '</td></tr></table>';
}

// A titled card containing one or more policy rows, styled like a
// lead-notification "information card": dark header strip, clean
// label:value rows underneath, alternating row shading.
function sectionCard(title, policies) {
  const titleHtml = title
    ? '<p style="font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:' + BRAND.chocolate + ';text-transform:uppercase;letter-spacing:0.5px;margin:18px 0 8px 0;">' + escapeHtml(title) + '</p>'
    : '';

  let cards = '';
  policies.forEach(p => {
    cards += policyCard(p);
  });

  return titleHtml + cards;
}

function policyCard(p) {
  const dueStr = Utilities.formatDate(p.due, getTZ(), 'd MMM yyyy');
  const badgeColor = p.daysLeft < 0 ? BRAND.burgundy : (p.daysLeft <= 7 ? BRAND.orange : BRAND.gold);
  const badgeText = p.daysLeft < 0 ? Math.abs(p.daysLeft) + ' days overdue'
                    : (p.daysLeft === 0 ? 'Due today' : p.daysLeft + ' day(s) left');

  const fieldRow = (label, value, isLast) =>
    '<tr>' +
      '<td style="padding:9px 14px;width:38%;font-family:Arial,sans-serif;font-size:11px;color:' + BRAND.muted + ';text-transform:uppercase;letter-spacing:0.3px;vertical-align:top;' + (isLast ? '' : 'border-bottom:1px solid ' + BRAND.line + ';') + '">' + label + '</td>' +
      '<td style="padding:9px 14px;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:' + BRAND.chocolate + ';word-break:break-word;overflow-wrap:break-word;vertical-align:top;' + (isLast ? '' : 'border-bottom:1px solid ' + BRAND.line + ';') + '">' + value + '</td>' +
    '</tr>';

  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;max-width:100%;margin-bottom:12px;background:#FFFFFF;border:1px solid ' + BRAND.line + ';border-radius:10px;overflow:hidden;">' +
    // header strip
    '<tr><td style="background:' + BRAND.chocolate + ';padding:10px 14px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
        '<td style="font-family:Arial,sans-serif;font-size:14px;font-weight:bold;color:#FFFFFF;word-break:break-word;">' + escapeHtml(p.client) + '</td>' +
        '<td style="text-align:right;white-space:nowrap;">' +
          '<span style="background:' + badgeColor + ';color:#FFFFFF;font-family:Arial,sans-serif;font-size:10px;font-weight:bold;padding:3px 9px;border-radius:12px;display:inline-block;">' + badgeText + '</span>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>' +
    // body rows
    '<tr><td>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;">' +
        fieldRow('Class', escapeHtml(p.cls || '-'), false) +
        fieldRow('Description', escapeHtml(p.desc || '-'), false) +
        fieldRow('Due date', dueStr, true) +
      '</table>' +
    '</td></tr>' +
  '</table>';
}

function escapeHtml(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── ACTIVITY LOG ───────────────────────────────────────────────────
function logActivity(policyId, action, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return;
  const nextRow = sheet.getLastRow() + 1;
  const row = Math.max(nextRow, 5);
  sheet.getRange(row, 2, 1, 4).setValues([[new Date(), policyId || '', action, notes || '']]);
}
