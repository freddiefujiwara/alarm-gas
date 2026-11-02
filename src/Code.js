// alarm-gas
//
// A small alarm scheduler backed by a Google Sheets sheet.
// A time-driven trigger runs myFunction() every minute. When a row's
// datetime is close to now, the row's commands are sent to the home
// command API one by one, and the row is removed. Rows whose
// datetime has already passed are removed without firing,
// so the sheet cleans itself up.
//
// The API base URL and the default speech command prefix live in
// script properties, not in this public source code:
//   API_BASE_URL        Base URL of the home command API.
//   DEFAULT_SPEECH_PATH Path prefix of the speech command used for
//                       legacy text alarms, ending with "/-s/".
//
// Sheet layout (the sheet is machine-managed; there is no header row):
//   Column A:     datetime to fire (a Date value)
//   Columns B..K: up to 10 commands. A command is a path on the home
//                 command API (for example "/hue/lights/off")
//                 or a full URL.
//
// Web API (doGet, legacy and anonymous, used by ha.ze.gs):
//   time     "HH:MM" or "HH:MM:SS" (today), or "YYYYMMDD-HH:MM".
//   cmd      A command. Repeat the parameter to register up to 10.
//   text     Legacy speech text. When no cmd is given, one default
//            speech command for the Nest speaker is registered
//            instead, like the old version did. The text is optional:
//            with no text at all, the alarm just announces the time
//            (for example "7時30分です。"), also like the old version.
//   callback Legacy JSONP callback name.
//
// Web API (doPost, JSON body, used by the alarm frontend):
//   Actions require a Google ID token in "idToken"; the token
//   is verified and its email must be in the AVAILABLE_GMAILS script
//   property (the same scheme as the other private sites). The OAuth
//   client ID used for verification is GOOGLE_OAUTH_CLIENT_ID.
//   { action: "list" }   -> the registered alarms.
//   { action: "set", time, commands } or { ..., text }
//                        -> register (or overwrite) an alarm.
//   { action: "delete", time } -> remove the alarm for that minute.
//   { action: "list_rules" }   -> the recurring rules.
//   { action: "set_rule", time, weekdays | exclude_weekdays,
//     until?, commands | text }
//                        -> add a recurring rule (at most MAX_RULES).
//   { action: "delete_rule", id } -> remove one recurring rule.
//   { action: "issue_share" }  -> issue a temporary access link.
//   { action: "revoke_share" } -> revoke the active access links.
//   As an alternative to "idToken", the list/set/delete actions also
//   accept a "share_token" from an issued access link. A share token
//   has the same rights as the owner (it can register and delete
//   alarms); share-authenticated responses carry "_share.expires_at".
//
// Recurring rules live in a separate sheet named "rules" (created on
// demand, with a header row): id, time of day ("HH:MM"), weekdays
// (comma-separated active days, 0 = Sunday; empty means every day),
// until ("YYYYMMDD", empty means no end), last_fired ("YYYYMMDD"),
// then up to 10 commands. The every-minute trigger fires a rule when
// today is an active day within the fire window, marks last_fired so
// it cannot fire twice, and deletes rules whose until date has passed.
//
// A request for a datetime that already exists overwrites that row.
// A request for a past datetime is rejected with an error.
// Note: Apps Script web apps cannot set an HTTP status code, so
// errors are reported in the JSON body: { "ok": false, "error": "..." }.

const MAX_COMMANDS = 10;
const FIRE_WINDOW_MS = 60 * 1000;

const getProperty_ = (name) => {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error(`script property ${name} is not set`);
  }
  return value;
};

const apiBaseUrl_ = () => getProperty_('API_BASE_URL');
const defaultSpeechPath_ = () => getProperty_('DEFAULT_SPEECH_PATH');

const pad2_ = (n) => String(n).padStart(2, '0');

// "HH:MM", "HH:MM:SS" -> today at that time.
// "YYYYMMDD-HH:MM" (seconds optional) -> that date and time.
// Anything else -> null.
export const parseTarget_ = (value, now) => {
  if (typeof value !== 'string') return null;
  const v = value.trim();
  let m = v.match(/^(\d{4})(\d{2})(\d{2})-(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const year = +m[1];
    const month = +m[2];
    const day = +m[3];
    const hour = +m[4];
    const minute = +m[5];
    const second = +(m[6] || 0);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    if (hour > 23 || minute > 59 || second > 59) return null;
    const d = new Date(year, month - 1, day, hour, minute, second);
    if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
      return null;
    }
    return d;
  }
  m = v.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const hour = +m[1];
    const minute = +m[2];
    const second = +(m[3] || 0);
    if (hour > 23 || minute > 59 || second > 59) return null;
    return new Date(now.getFullYear(), now.getMonth(), now.getDate(), hour, minute, second);
  }
  return null;
};

export const formatTarget_ = (d) =>
  `${d.getFullYear()}${pad2_(d.getMonth() + 1)}${pad2_(d.getDate())}-${pad2_(d.getHours())}:${pad2_(d.getMinutes())}`;

// Same wall-clock minute (seconds are ignored for identity).
export const sameMinute_ = (a, b) =>
  a.getFullYear() === b.getFullYear() &&
  a.getMonth() === b.getMonth() &&
  a.getDate() === b.getDate() &&
  a.getHours() === b.getHours() &&
  a.getMinutes() === b.getMinutes();

// What a text-less alarm says after the time: a nudge to get moving.
export const EMPTY_TEXT_SUFFIX_ = '時間ですよ。時間ですよ。そろそろ時間ですよ。';

// The speech text the old version built: "7時30分です。<text>".
// With no text, the suffix above follows the time instead.
export const buildSpeechText_ = (target, text) =>
  `${target.getHours() % 12 || 12}時${target.getMinutes() ? target.getMinutes() + '分' : 'ちょうど'}です。${text ? text : EMPTY_TEXT_SUFFIX_}`;

// A speech command (path on the home command API) for legacy text.
export const defaultSpeechCommand_ = (target, text) =>
  defaultSpeechPath_() + encodeURIComponent(buildSpeechText_(target, text));

// Commands from the request. When no cmd is given, this falls back
// to one default speech command: with the legacy "text" when present,
// or the time announcement (with its nudge suffix) when it is not.
export const collectRequestCommands_ = (e, target) => {
  const params = (e && e.parameters) || {};
  let cmds = [];
  if (Array.isArray(params.cmd)) {
    cmds = params.cmd.slice();
  } else if (e && e.parameter && e.parameter.cmd) {
    cmds = [e.parameter.cmd];
  }
  cmds = cmds.map((c) => String(c).trim()).filter((c) => c.length > 0);
  if (cmds.length === 0) {
    const raw = e && e.parameter ? e.parameter.text : null;
    const text = raw ? String(raw).trim() : '';
    cmds = [defaultSpeechCommand_(target, text)];
  }
  return cmds;
};

const respond_ = (e, payload) => {
  const output = ContentService.createTextOutput();
  const callback = e && e.parameter ? e.parameter.callback : undefined;
  if (typeof callback !== 'undefined') {
    // Legacy JSONP shape, kept for the old frontend: cb&&cb({...});
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    output.setContent(callback + '&&' + callback + '(' + JSON.stringify(payload) + ');');
  } else {
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(JSON.stringify(payload));
  }
  return output;
};

export const doGet = (e) => {
  const params = (e && e.parameter) || {};
  const now = new Date();
  const target = parseTarget_(params.time || '', now);
  if (!target) {
    return respond_(e, {
      ok: false,
      error: 'invalid time format: use HH:MM or YYYYMMDD-HH:MM',
    });
  }
  if (target.getTime() <= now.getTime()) {
    return respond_(e, { ok: false, error: 'time is in the past' });
  }
  const commands = collectRequestCommands_(e, target);
  if (commands.length > MAX_COMMANDS) {
    return respond_(e, {
      ok: false,
      error: `too many commands: max is ${MAX_COMMANDS}`,
    });
  }
  upsertAlarm_(target, commands);
  return respond_(e, { ok: true, time: formatTarget_(target), commands: commands.length });
};

// Write one alarm row, overwriting the row for the same minute when
// one already exists.
export const upsertAlarm_ = (target, commands) => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    let rowIndex = -1;
    values.forEach((row, i) => {
      const d = normalizeCellDate_(row[0], now);
      if (d && sameMinute_(d, target)) rowIndex = i;
    });
    const rowValues = [target].concat(commands);
    if (rowIndex >= 0) {
      sheet.getRange(rowIndex + 1, 1, 1, MAX_COMMANDS + 1).clearContent();
      sheet.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  } finally {
    lock.releaseLock();
  }
};

// The registered alarms as [{ target, commands }], oldest first.
// Rows whose first cell is not a readable date are skipped.
export const readAlarms_ = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const alarms = [];
  values.forEach((row) => {
    const target = normalizeCellDate_(row[0], now);
    if (!target) return;
    const commands = row
      .slice(1)
      .map((c) => String(c == null ? '' : c).trim())
      .filter((c) => c.length > 0);
    alarms.push({ target, commands });
  });
  alarms.sort((a, b) => a.target.getTime() - b.target.getTime());
  return alarms;
};

// Remove the alarm row for the target minute. Returns true when a
// row was removed.
export const deleteAlarm_ = (target) => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    for (let i = 0; i < values.length; i++) {
      const d = normalizeCellDate_(values[i][0], now);
      if (d && sameMinute_(d, target)) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
};

// --- Authentication for doPost (same scheme as the other private
// sites: verify the Google ID token, then check the allowlist) ---

const allowedEmails_ = () =>
  (getProperty_('AVAILABLE_GMAILS') || '')
    .split(',')
    .map((v) => v.trim().toLowerCase())
    .filter((v) => v.length > 0);

// Throws an Error when the token is missing, invalid, or not allowed.
export const verifyIdToken_ = (idToken) => {
  if (!idToken) throw new Error('missing id token');
  const clientId = getProperty_('GOOGLE_OAUTH_CLIENT_ID');
  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) throw new Error('token verification failed');
  const payload = JSON.parse(res.getContentText());
  if (payload.iss !== 'accounts.google.com' && payload.iss !== 'https://accounts.google.com') {
    throw new Error('invalid iss');
  }
  if (payload.aud !== clientId) throw new Error('invalid aud');
  const exp = Number(payload.exp || 0);
  if (!exp || exp * 1000 <= Date.now()) throw new Error('token expired');
  if (payload.email_verified !== true && payload.email_verified !== 'true') {
    throw new Error('email not verified');
  }
  const email = String(payload.email || '').toLowerCase();
  if (!email) throw new Error('missing email');
  if (!allowedEmails_().includes(email)) throw new Error('forbidden email');
  return email;
};

// --- Temporary share sessions (temporary access links) ---
// A logged-in owner can issue a URL token that grants the same access
// as the owner for SHARE_TTL_SECONDS without a Google account (the
// link is handed to a trusted agent to operate the alarms). Sessions
// are stored in a dedicated sheet in this spreadsheet (token hashes
// only, never raw tokens), which doubles as an audit log. The alarm
// code only ever reads the first sheet, so the share sheet is never
// exposed as alarm data.

const SHARE_SHEET_NAME = 'share_sessions';
// One month, so the owner can hand the link over well in advance.
const SHARE_TTL_SECONDS = 30 * 24 * 60 * 60;
const SHARE_HEADERS = [
  'token_hash',
  'issued_at',
  'expires_at',
  'revoked_at',
  'access_count',
  'last_accessed_at',
];

export const getShareSheet_ = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHARE_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHARE_SHEET_NAME);
    sheet.appendRow(SHARE_HEADERS);
  }
  return sheet;
};

export const hashShareToken_ = (token) =>
  Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))
    .map((byte) => ((byte % 256) + 256) % 256)
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');

export const readShareRows_ = () => {
  const sheet = getShareSheet_();
  const values = sheet.getDataRange().getValues();
  return { sheet, values };
};

export const shareCellIso_ = (value) => {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
};

export const revokeActiveShareSessions_ = (nowIso) => {
  const { sheet, values } = readShareRows_();
  values.forEach((row, index) => {
    if (index === 0) {
      return;
    }
    const revokedAt = shareCellIso_(row[3]);
    const expiresAt = shareCellIso_(row[2]);
    if (!revokedAt && expiresAt && expiresAt > nowIso) {
      sheet.getRange(index + 1, 4).setValue(nowIso);
    }
  });
};

export const issueShareSession_ = () => {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  revokeActiveShareSessions_(nowIso);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getShareSheet_().appendRow([hashShareToken_(token), nowIso, expiresIso, '', 0, '']);
  return {
    share_token: token,
    issued_at: nowIso,
    expires_at: expiresIso,
    ttl_seconds: SHARE_TTL_SECONDS,
  };
};

export const revokeShareSession_ = () => {
  revokeActiveShareSessions_(new Date().toISOString());
  return { revoked: true };
};

export const validateShareToken_ = (token) => {
  if (!token) {
    return null;
  }
  const hash = hashShareToken_(token);
  const { sheet, values } = readShareRows_();
  const nowIso = new Date().toISOString();
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    if (row[0] !== hash) {
      continue;
    }
    if (shareCellIso_(row[3])) {
      return null;
    }
    const expiresAt = shareCellIso_(row[2]);
    if (!expiresAt || expiresAt <= nowIso) {
      return null;
    }
    sheet.getRange(index + 1, 5).setValue(Number(row[4] || 0) + 1);
    sheet.getRange(index + 1, 6).setValue(nowIso);
    return { expires_at: expiresAt };
  }
  return null;
};

// --- Machine sync token ---
// A single long-lived token for machine-to-machine rule sync: the
// time-table script mirrors its sheet into recurring rules here, so
// firing lives in one place while the sheet stays private to that
// script (this project never reads it). The token lives only in the
// SYNC_TOKEN script property, never in code, and is scoped to the
// rule actions below — it cannot touch one-shot alarms or share
// sessions, and it does not expire like a share session.

const SYNC_ACTIONS_ = ['list_rules', 'set_rule', 'delete_rule'];

export const validateSyncToken_ = (token) => {
  if (!token) {
    return false;
  }
  const expected = getProperty_('SYNC_TOKEN');
  return !!expected && String(token) === expected;
};

// Commands for a doPost "set" body: an explicit commands array wins;
// otherwise fall back to one speech command (text may be empty).
export const collectBodyCommands_ = (body, target) => {
  let cmds = [];
  if (Array.isArray(body.commands)) {
    cmds = body.commands.map((c) => String(c).trim()).filter((c) => c.length > 0);
  }
  if (cmds.length === 0) {
    const text = body.text ? String(body.text).trim() : '';
    cmds = [defaultSpeechCommand_(target, text)];
  }
  return cmds;
};

const jsonOut_ = (payload) => {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify(payload));
  return output;
};

const authError_ = (err) => {
  const message = (err && err.message) || 'unauthorized';
  return jsonOut_({ ok: false, status: message === 'forbidden email' ? 403 : 401, error: message });
};

export const doPost = (e) => {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON body' });
  }

  // Issuing and revoking access links always requires the owner's
  // Google ID token; a share token cannot mint more share tokens.
  if (body.action === 'issue_share' || body.action === 'revoke_share') {
    try {
      verifyIdToken_(body.idToken);
    } catch (err) {
      return authError_(err);
    }
    const result = body.action === 'issue_share' ? issueShareSession_() : revokeShareSession_();
    return jsonOut_({ ok: true, ...result });
  }

  // The remaining actions accept the owner's ID token, a share token
  // from an issued access link (owner-equivalent), or — for the rule
  // actions only — the machine sync token.
  let shareSession = null;
  let syncAuth = false;
  if (body.sync_token) {
    if (!validateSyncToken_(body.sync_token)) {
      return jsonOut_({ ok: false, status: 401, error: 'sync token invalid' });
    }
    if (!SYNC_ACTIONS_.includes(body.action)) {
      return jsonOut_({ ok: false, status: 403, error: 'sync token may only manage rules' });
    }
    syncAuth = true;
  } else if (body.share_token) {
    shareSession = validateShareToken_(body.share_token);
    if (!shareSession) {
      return jsonOut_({ ok: false, status: 401, error: 'share session expired or invalid' });
    }
  } else {
    try {
      verifyIdToken_(body.idToken);
    } catch (err) {
      return authError_(err);
    }
  }
  const out = (payload) =>
    shareSession
      ? { ...payload, _share: { expires_at: shareSession.expires_at } }
      : syncAuth
        ? { ...payload, _sync: true }
        : payload;

  if (body.action === 'list') {
    return jsonOut_(
      out({
        ok: true,
        alarms: readAlarms_().map((a) => ({ time: formatTarget_(a.target), commands: a.commands })),
      })
    );
  }

  if (body.action === 'delete' || body.action === 'set') {
    const now = new Date();
    const target = parseTarget_(body.time || '', now);
    if (!target) {
      return jsonOut_({ ok: false, error: 'invalid time format: use HH:MM or YYYYMMDD-HH:MM' });
    }
    if (body.action === 'delete') {
      return jsonOut_(out({ ok: true, deleted: deleteAlarm_(target) ? 1 : 0 }));
    }
    if (target.getTime() <= now.getTime()) {
      return jsonOut_({ ok: false, error: 'time is in the past' });
    }
    const commands = collectBodyCommands_(body, target);
    if (commands.length > MAX_COMMANDS) {
      return jsonOut_({ ok: false, error: `too many commands: max is ${MAX_COMMANDS}` });
    }
    upsertAlarm_(target, commands);
    return jsonOut_(out({ ok: true, time: formatTarget_(target), commands: commands.length }));
  }

  if (body.action === 'list_rules') {
    return jsonOut_(out({ ok: true, rules: readRules_().map(serializeRule_) }));
  }

  if (body.action === 'delete_rule') {
    const id = body.id ? String(body.id).trim() : '';
    if (!id) {
      return jsonOut_({ ok: false, error: 'missing rule id' });
    }
    return jsonOut_(out({ ok: true, deleted: deleteRule_(id) ? 1 : 0 }));
  }

  if (body.action === 'set_rule') {
    const time = parseTimeOfDay_(body.time || '');
    if (!time) {
      return jsonOut_({ ok: false, error: 'invalid time format: use HH:MM' });
    }
    let weekdays;
    if (body.exclude_weekdays != null) {
      const excluded = parseWeekdays_(body.exclude_weekdays);
      if (!excluded) {
        return jsonOut_({ ok: false, error: 'invalid exclude_weekdays: use day numbers 0-6' });
      }
      weekdays = ALL_WEEKDAYS_.filter((d) => !excluded.includes(d));
      if (weekdays.length === 0) {
        return jsonOut_({ ok: false, error: 'no active weekdays left' });
      }
    } else {
      weekdays = parseWeekdays_(body.weekdays);
      if (!weekdays) {
        return jsonOut_({ ok: false, error: 'invalid weekdays: use day numbers 0-6' });
      }
    }
    const until = parseDateKey_(body.until);
    if (until === null) {
      return jsonOut_({ ok: false, error: 'invalid until date: use YYYYMMDD' });
    }
    if (until && until < formatDateKey_(new Date())) {
      return jsonOut_({ ok: false, error: 'until date is in the past' });
    }
    const target = new Date();
    target.setHours(time.hour, time.minute, 0, 0);
    const commands = collectBodyCommands_(body, target);
    if (commands.length > MAX_COMMANDS) {
      return jsonOut_({ ok: false, error: `too many commands: max is ${MAX_COMMANDS}` });
    }
    const rule = {
      id: Utilities.getUuid(),
      hour: time.hour,
      minute: time.minute,
      weekdays,
      until,
      lastFired: '',
      commands,
    };
    try {
      addRule_(rule);
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
    return jsonOut_(out({ ok: true, rule: serializeRule_(rule) }));
  }

  return jsonOut_({ ok: false, error: 'unknown action' });
};

// Fire one row's command cells: convert each cell to a URL, gather
// any legacy speech texts into one default speech command in front,
// and call the home command API once per command.
export const fireCells_ = (target, cells) => {
  const urls = [];
  const legacyTexts = [];
  cells.forEach((c) => {
    const url = toUrl_(c);
    if (url) urls.push(url);
    else legacyTexts.push(c);
  });
  if (legacyTexts.length > 0) {
    // Old-format row ([time, text]): speak it like the old version.
    urls.unshift(apiBaseUrl_() + defaultSpeechCommand_(target, legacyTexts.join(' ')));
  }
  urls.slice(0, MAX_COMMANDS).forEach((url) => {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      method: 'get',
    });
    Logger.log(url + ' -> ' + res.getResponseCode());
  });
};

// --- Recurring rules ---
// A rule fires at a fixed time of day on its active weekdays, with an
// optional end date. Rules are stored in their own sheet so the
// one-shot sheet keeps its self-cleaning shape.

const RULES_SHEET_NAME = 'rules';
const MAX_RULES = 20;
const ALL_WEEKDAYS_ = [0, 1, 2, 3, 4, 5, 6];
const RULES_HEADERS = ['id', 'time', 'weekdays', 'until', 'last_fired', 'commands'];

export const getRulesSheet_ = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(RULES_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RULES_SHEET_NAME);
    sheet.appendRow(RULES_HEADERS);
  }
  return sheet;
};

// "HH:MM" (seconds optional) -> { hour, minute } or null.
export const parseTimeOfDay_ = (value) => {
  if (typeof value !== 'string') return null;
  const m = value.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  const hour = +m[1];
  const minute = +m[2];
  const second = +(m[3] || 0);
  if (hour > 23 || minute > 59 || second > 59) return null;
  return { hour, minute };
};

export const formatTimeOfDay_ = (hour, minute) => `${pad2_(hour)}:${pad2_(minute)}`;

export const formatDateKey_ = (d) =>
  `${d.getFullYear()}${pad2_(d.getMonth() + 1)}${pad2_(d.getDate())}`;

// Weekdays input (array of day numbers, or a comma-separated string)
// -> a sorted unique array, or null when a value is not a day number.
// An empty input means every day.
export const parseWeekdays_ = (value) => {
  if (value == null || value === '') return ALL_WEEKDAYS_.slice();
  const items = Array.isArray(value) ? value : String(value).split(',');
  const days = [];
  for (const item of items) {
    if (typeof item === 'string' && item.trim() === '') continue;
    const n = Number(item);
    if (!Number.isInteger(n) || n < 0 || n > 6) return null;
    if (!days.includes(n)) days.push(n);
  }
  if (days.length === 0) return ALL_WEEKDAYS_.slice();
  return days.sort((a, b) => a - b);
};

// An "until" cell or input -> "YYYYMMDD", or '' when empty. Returns
// null for a value that is not a real date. Accepts "YYYYMMDD",
// "YYYY-MM-DD", a Date, or an 8-digit number (Sheets may store the
// bare digits as a number).
export const parseDateKey_ = (value) => {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : formatDateKey_(value);
  }
  const digits = String(value).replace(/\D/g, '');
  if (!/^\d{8}$/.test(digits)) return null;
  const year = +digits.slice(0, 4);
  const month = +digits.slice(4, 6);
  const day = +digits.slice(6, 8);
  const d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
    return null;
  }
  return digits;
};

// A rule row's time cell: a "HH:MM" string, or a Date (Sheets turns
// "7:30" into a time value on a 1899 date; only the clock matters).
export const parseRuleTimeCell_ = (value) => {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return { hour: value.getHours(), minute: value.getMinutes() };
  }
  return parseTimeOfDay_(value == null ? '' : String(value));
};

// One rules-sheet row -> a rule object, or null when the row is not a
// readable rule (blank rows are handled by the caller).
export const rowToRule_ = (row) => {
  const id = String(row[0] == null ? '' : row[0]).trim();
  if (!id) return null;
  const time = parseRuleTimeCell_(row[1]);
  if (!time) return null;
  const weekdays = parseWeekdays_(row[2] == null ? '' : String(row[2]));
  if (!weekdays) return null;
  const until = parseDateKey_(row[3]);
  if (until === null) return null;
  const lastFired = parseDateKey_(row[4]) || '';
  const commands = row
    .slice(5)
    .map((c) => String(c == null ? '' : c).trim())
    .filter((c) => c.length > 0);
  return { id, hour: time.hour, minute: time.minute, weekdays, until, lastFired, commands };
};

export const serializeRule_ = (rule) => ({
  id: rule.id,
  time: formatTimeOfDay_(rule.hour, rule.minute),
  weekdays: rule.weekdays,
  until: rule.until || null,
  commands: rule.commands,
});

// The stored rules, in sheet order. Malformed rows are skipped.
export const readRules_ = () => {
  const sheet = getRulesSheet_();
  const values = sheet.getDataRange().getValues();
  const rules = [];
  values.forEach((row, i) => {
    if (i === 0) return; // header
    const rule = rowToRule_(row);
    if (rule) rules.push(rule);
  });
  return rules;
};

// Append a rule. Throws when the rule limit has been reached.
export const addRule_ = (rule) => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (readRules_().length >= MAX_RULES) {
      throw new Error(`too many rules: max is ${MAX_RULES}`);
    }
    getRulesSheet_().appendRow([
      rule.id,
      formatTimeOfDay_(rule.hour, rule.minute),
      rule.weekdays.join(','),
      rule.until || '',
      '',
      ...rule.commands,
    ]);
  } finally {
    lock.releaseLock();
  }
};

// Remove the rule with this id. Returns true when one was removed.
export const deleteRule_ = (id) => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getRulesSheet_();
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] == null ? '' : values[i][0]).trim() === id) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
};

// Fire the rules that are due today. Called from the every-minute
// trigger: a rule fires when today is one of its weekdays, the clock
// is inside the fire window, and it has not fired today yet
// (last_fired guards against the two trigger runs inside the
// window). Rules past their until date are deleted.
export const processRules_ = (now) => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(RULES_SHEET_NAME);
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const todayKey = formatDateKey_(now);
  const deleteRows = [];

  values.forEach((row, i) => {
    if (i === 0) return; // header
    const rule = rowToRule_(row);
    if (!rule) {
      const hasContent = row.some((c) => String(c == null ? '' : c).trim().length > 0);
      if (hasContent) deleteRows.push(i + 1); // unreadable row: clean up
      return;
    }
    if (rule.until && todayKey > rule.until) {
      deleteRows.push(i + 1);
      return;
    }
    if (!rule.weekdays.includes(now.getDay())) return;
    if (rule.lastFired === todayKey) return;
    const target = new Date(now.getFullYear(), now.getMonth(), now.getDate(), rule.hour, rule.minute, 0);
    if (classify_(target, now) !== 'fire') return;
    fireCells_(target, rule.commands);
    sheet.getRange(i + 1, 5).setValue(todayKey);
  });

  deleteRows
    .sort((a, b) => b - a)
    .forEach((rowNumber) => sheet.deleteRow(rowNumber));
};

// A sheet cell -> Date (or null). Time-only values left by the old
// version come back with a year of 1899; those mean "today at that
// time", like the old code treated them.
export const cellToDate_ = (value) => {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string' && value.trim().length > 0) {
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
};

export const normalizeCellDate_ = (value, now) => {
  const d = cellToDate_(value);
  if (!d) return null;
  if (d.getFullYear() < 2000) {
    return new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      d.getHours(),
      d.getMinutes(),
      d.getSeconds()
    );
  }
  return d;
};

// A command cell -> full URL, or null when the cell is legacy speech
// text. Commands come in two forms:
// - path form:  "/switchbot-command/-d/X/-c/turnOff"
// - space form: "switchbot-command -d X -c turnOff" (what the alarm
//   page sends now; the first token is the ASCII command name, so
//   Japanese speech text is never mistaken for a command)
export const toUrl_ = (command) => {
  if (command.startsWith('http://') || command.startsWith('https://')) {
    return command;
  }
  if (command.startsWith('/')) {
    return apiBaseUrl_() + command;
  }
  const tokens = command.trim().split(/ +/).filter((t) => t.length > 0);
  if (tokens.length >= 2 && /^[A-Za-z0-9][A-Za-z0-9-]*$/.test(tokens[0])) {
    return apiBaseUrl_() + '/' + tokens.join('/');
  }
  return null;
};

// 'fire' inside the window, 'cleanup' after it, 'keep' before it.
export const classify_ = (target, now) => {
  const diff = target.getTime() - now.getTime();
  if (Math.abs(diff) <= FIRE_WINDOW_MS) return 'fire';
  return diff < 0 ? 'cleanup' : 'keep';
};

// Runs every minute from the time-driven trigger. The function name is
// kept from the old version so the existing trigger keeps working.
export const myFunction = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const deleteRows = [];

  values.forEach((row, i) => {
    const cells = row
      .slice(1)
      .map((c) => String(c == null ? '' : c).trim())
      .filter((c) => c.length > 0);
    const target = normalizeCellDate_(row[0], now);
    if (!target && cells.length === 0) {
      return; // blank row: leave it alone
    }
    if (!target) {
      deleteRows.push(i + 1); // a row we cannot read: clean it up
      return;
    }
    const kind = classify_(target, now);
    if (kind === 'keep') return;

    if (kind === 'fire') {
      fireCells_(target, cells);
    }
    deleteRows.push(i + 1);
  });

  // Delete from the bottom up so row numbers do not shift.
  deleteRows
    .sort((a, b) => b - a)
    .forEach((rowNumber) => sheet.deleteRow(rowNumber));

  processRules_(now);
};
