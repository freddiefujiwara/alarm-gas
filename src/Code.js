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
//   Every action requires a Google ID token in "idToken"; the token
//   is verified and its email must be in the AVAILABLE_GMAILS script
//   property (the same scheme as the other private sites). The OAuth
//   client ID used for verification is GOOGLE_OAUTH_CLIENT_ID.
//   { action: "list" }   -> the registered alarms.
//   { action: "set", time, commands } or { ..., text }
//                        -> register (or overwrite) an alarm.
//   { action: "delete", time } -> remove the alarm for that minute.
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

export const doPost = (e) => {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON body' });
  }
  try {
    verifyIdToken_(body.idToken);
  } catch (err) {
    const message = (err && err.message) || 'unauthorized';
    return jsonOut_({ ok: false, status: message === 'forbidden email' ? 403 : 401, error: message });
  }

  if (body.action === 'list') {
    return jsonOut_({
      ok: true,
      alarms: readAlarms_().map((a) => ({ time: formatTarget_(a.target), commands: a.commands })),
    });
  }

  if (body.action === 'delete' || body.action === 'set') {
    const now = new Date();
    const target = parseTarget_(body.time || '', now);
    if (!target) {
      return jsonOut_({ ok: false, error: 'invalid time format: use HH:MM or YYYYMMDD-HH:MM' });
    }
    if (body.action === 'delete') {
      return jsonOut_({ ok: true, deleted: deleteAlarm_(target) ? 1 : 0 });
    }
    if (target.getTime() <= now.getTime()) {
      return jsonOut_({ ok: false, error: 'time is in the past' });
    }
    const commands = collectBodyCommands_(body, target);
    if (commands.length > MAX_COMMANDS) {
      return jsonOut_({ ok: false, error: `too many commands: max is ${MAX_COMMANDS}` });
    }
    upsertAlarm_(target, commands);
    return jsonOut_({ ok: true, time: formatTarget_(target), commands: commands.length });
  }

  return jsonOut_({ ok: false, error: 'unknown action' });
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
    }
    deleteRows.push(i + 1);
  });

  // Delete from the bottom up so row numbers do not shift.
  deleteRows
    .sort((a, b) => b - a)
    .forEach((rowNumber) => sheet.deleteRow(rowNumber));
};
