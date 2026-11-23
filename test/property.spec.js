import { beforeEach, describe, expect, it, vi } from "vitest";
import * as fc from "fast-check";
import {
  buildSpeechText_,
  collectBodyCommands_,
  defaultSpeechCommand_,
  doPost,
  formatTarget_,
  formatTimeOfDay_,
  parseDateKey_,
  parseLastFired_,
  parseTarget_,
  parseTimeOfDay_,
  parseWeekdays_,
  rowToRule_,
  serializeRule_,
} from "../src/Code.js";

// Property-based tests (fast-check) over the same outside-input
// boundary as test/abnormal.spec.js: doPost JSON bodies and the
// rules-sheet rows built from them.
//
// Generator domains, and why:
// - Body/cell values are drawn from the JSON domain (strings, finite
//   numbers, booleans, null, arrays, plain objects) plus one explicit
//   unstringifiable object ({ toString: [] }), which JSON can produce
//   and which must never make a reader throw. Sheet cells themselves
//   are scalars, so this domain is a superset of the real one.
// - Speech text is built from whole code points plus one explicit
//   lone-surrogate unit: JSON string escapes ("\ud800") can carry a
//   lone surrogate into a request body, so it is in-domain here and
//   the encoder must survive it. Other strings use whole code points
//   only, since well-formed text has no other surrogate shapes.
// - Dates come from fc.date with noInvalidDate: true, so the harness
//   itself never feeds an Invalid Date into the code under test.
// - The -0/+0 distinction (Object.is) is avoided by comparing day
//   sets and counts, never raw signed zeros.

// --- GAS mocks (mirroring test/Code.spec.js) -------------------------
let sheetValues;
let rulesValues;
let shareValues;
let ownerValues;
let outputContent;

const dataSheet = (get, onAppend) => ({
  getDataRange: vi.fn(() => ({ getValues: vi.fn(get) })),
  appendRow: vi.fn((row) => onAppend(row)),
  deleteRow: vi.fn((n) => get().splice(n - 1, 1)),
  insertColumnBefore: vi.fn((col) => {
    get().forEach((row) => row.splice(col - 1, 0, ""));
  }),
  getRange: vi.fn((row, col) => ({
    clearContent: vi.fn(),
    setValues: vi.fn(),
    setValue: vi.fn((value) => {
      get()[row - 1][col - 1] = value;
    }),
  })),
});
const mockAlarmSheet = dataSheet(
  () => sheetValues,
  (row) => sheetValues.push([...row])
);
const mockRulesSheet = dataSheet(
  () => rulesValues,
  (row) => rulesValues.push([...row])
);
const mockShareSheet = dataSheet(
  () => shareValues,
  (row) => shareValues.push([...row])
);
const mockOwnerSheet = dataSheet(
  () => ownerValues,
  (row) => ownerValues.push([...row])
);
const mockSpreadsheet = {
  getActiveSheet: vi.fn(() => mockAlarmSheet),
  getSheetByName: vi.fn((name) => {
    if (name === "rules") return mockRulesSheet;
    if (name === "share_sessions") return mockShareSheet;
    if (name === "owner_sessions") return mockOwnerSheet;
    return null;
  }),
  insertSheet: vi.fn((name) => {
    if (name === "rules") return mockRulesSheet;
    if (name === "owner_sessions") return mockOwnerSheet;
    return mockShareSheet;
  }),
};
global.SpreadsheetApp = {
  getActiveSheet: vi.fn(() => mockAlarmSheet),
  getActiveSpreadsheet: vi.fn(() => mockSpreadsheet),
  openById: vi.fn(() => mockSpreadsheet),
};
global.ContentService = {
  createTextOutput: vi.fn(() => ({
    setMimeType: vi.fn(),
    setContent: vi.fn((c) => {
      outputContent = c;
    }),
  })),
  MimeType: { JSON: "JSON", JAVASCRIPT: "JAVASCRIPT" },
};
global.LockService = {
  getScriptLock: vi.fn(() => ({ waitLock: vi.fn(), releaseLock: vi.fn() })),
};
let uuidCounter = 0;
global.Utilities = {
  DigestAlgorithm: { SHA_256: "SHA_256" },
  computeDigest: vi.fn((algo, str) =>
    Array.from({ length: 32 }, (_, i) => String(str).charCodeAt(i % String(str).length) % 128)
  ),
  getUuid: vi.fn(() => `00000000-0000-4000-8000-${String(++uuidCounter).padStart(12, "0")}`),
};
global.Logger = { log: vi.fn() };
global.UrlFetchApp = {
  fetch: vi.fn((url) => {
    if (String(url).startsWith("https://oauth2.googleapis.com/tokeninfo")) {
      return {
        getResponseCode: () => 200,
        getContentText: () =>
          JSON.stringify({
            iss: "https://accounts.google.com",
            aud: "dummy-client-id",
            exp: String(Math.floor(Date.now() / 1000) + 3600),
            email_verified: "true",
            email: "owner@example.com",
          }),
      };
    }
    return { getResponseCode: () => 200, getContentText: () => "OK" };
  }),
};
global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn(
      (name) =>
        ({
          API_BASE_URL: "http://home-api.example",
          DEFAULT_SPEECH_PATH: "/speak/-h/192.0.2.1/-v/60/-s/",
          AVAILABLE_GMAILS: "owner@example.com",
          GOOGLE_OAUTH_CLIENT_ID: "dummy-client-id",
          HOME_API_TOKEN: "dummy-home-api-token",
          SYNC_TOKEN: "test-sync-token",
        })[name] || null
    ),
  })),
};

beforeEach(() => {
  sheetValues = [];
  rulesValues = [["id", "time", "weekdays", "until", "last_fired", "exceptions", "commands"]];
  shareValues = [
    ["token_hash", "issued_at", "expires_at", "revoked_at", "access_count", "last_accessed_at"],
  ];
  ownerValues = [
    ["token_hash", "issued_at", "expires_at", "revoked_at", "access_count", "last_accessed_at"],
  ];
  outputContent = null;
  uuidCounter = 0;
  vi.clearAllMocks();
});

const unstringifiable = { toString: [] };
const jsonValue = fc.oneof(
  fc.string(),
  fc.double({ noNaN: true }),
  fc.boolean(),
  fc.constant(null),
  fc.array(fc.oneof(fc.string(), fc.double({ noNaN: true }), fc.boolean(), fc.constant(null)), {
    maxLength: 4,
  }),
  fc.record({ a: fc.string() }),
  fc.constant(unstringifiable)
);
const isValidDaySet = (days) =>
  Array.isArray(days) &&
  days.every((d) => Number.isInteger(d) && d >= 0 && d <= 6) &&
  days.every((d, i) => i === 0 || days[i - 1] < d);
const isRealDateKey = (key) => {
  if (!/^\d{8}$/.test(key)) return false;
  const y = +key.slice(0, 4);
  const m = +key.slice(4, 6);
  const d = +key.slice(6, 8);
  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d;
};
const hasLoneSurrogate = (s) =>
  /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/.test(s);

describe("parseWeekdays_ properties", () => {
  it("never throws and only yields null or a valid day set", () => {
    fc.assert(
      fc.property(fc.array(jsonValue, { maxLength: 6 }), (items) => {
        const days = parseWeekdays_(items);
        return days === null || isValidDaySet(days);
      }),
      { numRuns: 100 }
    );
  });

  it("conserves well-formed day inputs in both shapes", () => {
    fc.assert(
      fc.property(
        fc.uniqueArray(fc.integer({ min: 0, max: 6 }), { minLength: 1, maxLength: 7 }),
        (days) => {
          const sorted = [...days].sort((a, b) => a - b);
          expect(parseWeekdays_(days)).toEqual(sorted);
          expect(parseWeekdays_(days.join(","))).toEqual(sorted);
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe("time and date key round-trips", () => {
  it("formatTimeOfDay_ output always parses back", () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: 23 }), fc.integer({ min: 0, max: 59 }), (h, m) => {
        expect(parseTimeOfDay_(formatTimeOfDay_(h, m))).toEqual({ hour: h, minute: m });
      }),
      { numRuns: 100 }
    );
  });

  it("formatTarget_ output parses back to the same minute", () => {
    fc.assert(
      fc.property(
        fc.date({
          min: new Date(2000, 0, 1),
          max: new Date(2099, 11, 31),
          noInvalidDate: true,
        }),
        (d) => {
          const atMinute = new Date(
            d.getFullYear(),
            d.getMonth(),
            d.getDate(),
            d.getHours(),
            d.getMinutes()
          );
          expect(parseTarget_(formatTarget_(atMinute), new Date())).toEqual(atMinute);
        }
      ),
      { numRuns: 100 }
    );
  });

  it("parseDateKey_ only yields '', null, or a real calendar date", () => {
    fc.assert(
      fc.property(jsonValue, (value) => {
        const key = parseDateKey_(value);
        return key === "" || key === null || isRealDateKey(key);
      }),
      { numRuns: 100 }
    );
  });

  it("parseLastFired_ only yields '' or a valid datetime key", () => {
    fc.assert(
      fc.property(jsonValue, (value) => {
        const key = parseLastFired_(value);
        if (key === "") return true;
        if (!/^\d{8}(\d{4})?$/.test(key)) return false;
        if (!isRealDateKey(key.slice(0, 8))) return false;
        if (key.length === 12) {
          const hh = +key.slice(8, 10);
          const mm = +key.slice(10, 12);
          if (hh > 23 || mm > 59) return false;
        }
        return true;
      }),
      { numRuns: 100 }
    );
  });
});

describe("collectBodyCommands_ properties", () => {
  it("never throws and yields only non-empty string commands", () => {
    fc.assert(
      fc.property(
        fc.record({
          commands: fc.oneof(
            fc.constant(undefined),
            fc.array(jsonValue, { maxLength: 6 })
          ),
          text: fc.oneof(fc.constant(undefined), jsonValue),
        }),
        (body) => {
          const cmds = collectBodyCommands_(body, new Date(2999, 0, 1, 7, 30));
          return (
            Array.isArray(cmds) &&
            cmds.length >= 1 &&
            cmds.every((c) => typeof c === "string" && c.length > 0) &&
            cmds.every((c) => c !== "[object Object]")
          );
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe("speech text encoding properties", () => {
  const speechText = fc
    .array(fc.constantFrom("a", "あ", "🎉", " ", "\ud800", "\udc00", "時"), { maxLength: 20 })
    .map((units) => units.join(""));

  it("never throws and the encoded text carries no lone surrogate", () => {
    fc.assert(
      fc.property(speechText, (text) => {
        const cmd = defaultSpeechCommand_(new Date(2026, 10, 2, 7, 30), text);
        const spoken = decodeURIComponent(cmd.split("/-s/")[1]);
        return !hasLoneSurrogate(spoken);
      }),
      { numRuns: 100 }
    );
  });

  it("round-trips well-formed text unchanged", () => {
    fc.assert(
      fc.property(fc.string(), (text) => {
        fc.pre(!hasLoneSurrogate(text));
        const target = new Date(2026, 10, 2, 7, 30);
        const cmd = defaultSpeechCommand_(target, text);
        return decodeURIComponent(cmd.split("/-s/")[1]) === buildSpeechText_(target, text);
      }),
      { numRuns: 100 }
    );
  });
});

describe("rule row round-trip", () => {
  it("rowToRule_ reads back a serialized well-formed rule", () => {
    fc.assert(
      fc.property(
        fc.record({
          id: fc.string({ minLength: 1, maxLength: 12 }).filter((s) => s.trim() === s),
          hour: fc.integer({ min: 0, max: 23 }),
          minute: fc.integer({ min: 0, max: 59 }),
          weekdays: fc.uniqueArray(fc.integer({ min: 0, max: 6 }), {
            minLength: 1,
            maxLength: 7,
          }),
          until: fc.oneof(
            fc.constant(""),
            fc
              .date({ min: new Date(2000, 0, 1), max: new Date(2099, 11, 31), noInvalidDate: true })
              .map(
                (d) =>
                  `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(
                    d.getDate()
                  ).padStart(2, "0")}`
              )
          ),
          exceptions: fc.uniqueArray(
            fc
              .date({ min: new Date(2000, 0, 1), max: new Date(2099, 11, 31), noInvalidDate: true })
              .map(
                (d) =>
                  `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(
                    d.getDate()
                  ).padStart(2, "0")}`
              ),
            { maxLength: 3 }
          ),
          commands: fc.array(
            fc.string({ minLength: 1, maxLength: 10 }).filter((s) => s.trim() === s),
            { maxLength: 3 }
          ),
        }),
        (rule) => {
          const sorted = [...rule.weekdays].sort((a, b) => a - b);
          const sortedExceptions = [...rule.exceptions].sort();
          const row = [
            rule.id,
            formatTimeOfDay_(rule.hour, rule.minute),
            sorted.join(","),
            rule.until,
            "",
            sortedExceptions.join(","),
            ...rule.commands,
          ];
          expect(rowToRule_(row)).toEqual({
            id: rule.id,
            hour: rule.hour,
            minute: rule.minute,
            weekdays: sorted,
            until: rule.until,
            lastFired: "",
            exceptions: sortedExceptions,
            commands: rule.commands,
          });
          expect(serializeRule_(rowToRule_(row)).time).toBe(
            formatTimeOfDay_(rule.hour, rule.minute)
          );
        }
      ),
      { numRuns: 100 }
    );
  });

  it("rowToRule_ never throws on deformed rows and shapes its output", () => {
    fc.assert(
      fc.property(fc.array(jsonValue, { maxLength: 8 }), (row) => {
        const rule = rowToRule_(row);
        if (rule === null) return true;
        return (
          Number.isInteger(rule.hour) &&
          rule.hour >= 0 &&
          rule.hour <= 23 &&
          Number.isInteger(rule.minute) &&
          rule.minute >= 0 &&
          rule.minute <= 59 &&
          isValidDaySet(rule.weekdays) &&
          rule.commands.every((c) => typeof c === "string" && c.length > 0)
        );
      }),
      { numRuns: 100 }
    );
  });
});

describe("doPost totality", () => {
  it("always answers with parseable JSON carrying a boolean ok", () => {
    fc.assert(
      fc.property(
        fc.record(
          {
            action: fc.constantFrom(
              "list",
              "set",
              "delete",
              "list_rules",
              "set_rule",
              "delete_rule",
              "issue_share",
              "revoke_share",
              "exchange_session",
              "revoke_session",
              "bogus"
            ),
            sync_token: fc.oneof(fc.constant("test-sync-token"), jsonValue),
            time: fc.oneof(fc.constant("29990101-07:30"), fc.constant("07:30"), jsonValue),
            weekdays: fc.oneof(fc.constant(undefined), fc.array(jsonValue, { maxLength: 4 })),
            until: fc.oneof(fc.constant(undefined), jsonValue),
            id: fc.oneof(fc.constant(undefined), jsonValue),
            commands: fc.oneof(fc.constant(undefined), fc.array(jsonValue, { maxLength: 4 })),
            text: fc.oneof(fc.constant(undefined), jsonValue),
          },
          { requiredKeys: ["action"] }
        ),
        (body) => {
          doPost({ postData: { contents: JSON.stringify(body) } });
          const parsed = JSON.parse(outputContent);
          return typeof parsed.ok === "boolean";
        }
      ),
      { numRuns: 100 }
    );
  });
});
