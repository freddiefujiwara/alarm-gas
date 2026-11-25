import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  collectBodyCommands_,
  defaultSpeechCommand_,
  doPost,
  parseDateKey_,
  parseLastFired_,
  parseWeekdays_,
  rowToRule_,
} from "../src/Code.js";

// Abnormal-case tests: the "input from outside" for alarm-gas is
// (a) the JSON request bodies the alarm frontend and the time-table
// sync send to doPost, and (b) the rules-sheet rows those bodies turn
// into (Sheets hands cells back as scalars, but the shaping functions
// are probed over the full JSON domain as a superset, matching the
// fleet-wide hardening pattern). A deformed value must never invent a
// weekday, a date, or a command, must never delete the wrong rule,
// and must never throw out of doPost — the API answers { ok: false }
// instead. Well-formed behaviour is pinned unchanged.

// --- GAS mocks (mirroring test/Code.spec.js) -------------------------
let sheetValues;
let appended;
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
  getLastRow: vi.fn(() => get().length),
  getRange: vi.fn((row, col) => ({
    clearContent: vi.fn(),
    setValues: vi.fn(),
    setNumberFormat: vi.fn(),
    setValue: vi.fn((value) => {
      get()[row - 1][col - 1] = value;
    }),
  })),
});

const mockAlarmSheet = dataSheet(
  () => sheetValues,
  (row) => appended.push(row)
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
  appended = [];
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

// An object JSON can produce whose toString is not callable: String()
// and implicit conversion throw TypeError on it.
const unstringifiable = { toString: [] };

const post = (body) => doPost({ postData: { contents: JSON.stringify(body) } });
const postRaw = (contents) => doPost({ postData: { contents } });
const resultJson = () => JSON.parse(outputContent);
const sync = { sync_token: "test-sync-token" };
const auth = { idToken: "dummy-token" };

describe("parseWeekdays_ with deformed entries", () => {
  it("does not invent a weekday from booleans, null, or empty arrays", () => {
    // Bare Number() coercion used to read true as day 1, false and
    // null and [] as day 0, so a deformed rule fired on a wrong day.
    expect(parseWeekdays_([true])).toBeNull();
    expect(parseWeekdays_([false])).toBeNull();
    expect(parseWeekdays_([null])).toBeNull();
    expect(parseWeekdays_([[]])).toBeNull();
  });

  it("rejects an array-wrapped day instead of unwrapping it", () => {
    expect(parseWeekdays_([[7]])).toBeNull();
    expect(parseWeekdays_([[1]])).toBeNull();
  });

  it("never throws on unstringifiable values", () => {
    expect(parseWeekdays_([unstringifiable])).toBeNull();
    expect(parseWeekdays_(unstringifiable)).toBeNull();
    expect(parseWeekdays_([{}])).toBeNull();
  });

  it("keeps well-formed shapes unchanged", () => {
    expect(parseWeekdays_([" 3 ", "0"])).toEqual([0, 3]);
    expect(parseWeekdays_([3.5])).toBeNull();
    expect(parseWeekdays_(3)).toEqual([3]);
  });
});

describe("parseDateKey_ / parseLastFired_ with deformed values", () => {
  it("never throws on an unstringifiable until", () => {
    expect(parseDateKey_(unstringifiable)).toBeNull();
    expect(parseDateKey_([20261231])).toBeNull();
    expect(parseDateKey_(true)).toBeNull();
  });

  it("reads an unstringifiable last_fired cell as never fired", () => {
    expect(parseLastFired_(unstringifiable)).toBe("");
    expect(parseLastFired_([202611020730])).toBe("");
  });

  it("keeps well-formed shapes unchanged", () => {
    expect(parseDateKey_("2026-02-28")).toBe("20260228");
    expect(parseLastFired_("202611020730")).toBe("202611020730");
    expect(parseLastFired_("202602300730")).toBe("");
  });
});

describe("collectBodyCommands_ with deformed bodies", () => {
  const target = new Date(2999, 0, 1, 7, 30);

  it("drops unstringifiable command entries instead of throwing", () => {
    const cmds = collectBodyCommands_({ commands: [unstringifiable, "/a"] }, target);
    expect(cmds).toEqual(["/a"]);
  });

  it("does not store a plain object as an '[object Object]' command", () => {
    const cmds = collectBodyCommands_({ commands: [{}] }, target);
    expect(cmds).toHaveLength(1);
    expect(cmds[0]).not.toContain("[object Object]");
    expect(decodeURIComponent(cmds[0].split("/-s/")[1])).toContain("時間ですよ");
  });

  it("reads an unstringifiable text as no text", () => {
    const cmds = collectBodyCommands_({ text: unstringifiable }, target);
    expect(cmds).toHaveLength(1);
    expect(decodeURIComponent(cmds[0].split("/-s/")[1])).toContain("時間ですよ");
  });

  it("keeps scalar coercions unchanged", () => {
    expect(collectBodyCommands_({ commands: [42, true, " /b "] }, target)).toEqual([
      "42",
      "true",
      "/b",
    ]);
  });
});

describe("speech command encoding", () => {
  const target = new Date(2026, 10, 2, 7, 30);

  it("does not throw on a lone surrogate in the text (JSON can carry one)", () => {
    const cmd = defaultSpeechCommand_(target, "a\ud800b");
    const spoken = decodeURIComponent(cmd.split("/-s/")[1]);
    expect(spoken).toBe("7時30分です。a\ufffdb");
  });

  it("keeps a well-formed surrogate pair intact", () => {
    const cmd = defaultSpeechCommand_(target, "🎉");
    expect(decodeURIComponent(cmd.split("/-s/")[1])).toBe("7時30分です。🎉");
  });
});

describe("rowToRule_ with deformed rows", () => {
  it("drops an unstringifiable command cell instead of throwing", () => {
    const rule = rowToRule_(["r1", "07:30", "1", "", "", "", unstringifiable, "/a"]);
    expect(rule).not.toBeNull();
    expect(rule.commands).toEqual(["/a"]);
  });

  it("reads a row with an unstringifiable id or weekdays as unreadable", () => {
    expect(rowToRule_([unstringifiable, "07:30", "1", "", "", "", "/a"])).toBeNull();
    expect(rowToRule_(["r1", "07:30", unstringifiable, "", "", "", "/a"])).toBeNull();
  });

  it("keeps a well-formed row byte-identical in shape", () => {
    expect(rowToRule_(["r1", "07:30", "1,5", "20261231", "202611020730", "20261225", "/a"])).toEqual({
      id: "r1",
      hour: 7,
      minute: 30,
      weekdays: [1, 5],
      until: "20261231",
      lastFired: "202611020730",
      exceptions: ["20261225"],
      commands: ["/a"],
    });
  });
});

describe("doPost totality over deformed bodies", () => {
  it("answers a JSON null body instead of throwing", () => {
    expect(() => postRaw("null")).not.toThrow();
    expect(resultJson().ok).toBe(false);
  });

  it("rejects deformed tokens with 401 instead of throwing", () => {
    post({ action: "list", share_token: unstringifiable });
    expect(resultJson()).toMatchObject({ ok: false, status: 401 });
    post({ action: "list", session_token: unstringifiable });
    expect(resultJson()).toMatchObject({ ok: false, status: 401 });
    post({ action: "list_rules", sync_token: unstringifiable });
    expect(resultJson()).toMatchObject({ ok: false, status: 401 });
  });

  it("revokes nothing but still answers for an unstringifiable session token", () => {
    post({ action: "revoke_session", session_token: unstringifiable });
    expect(resultJson()).toEqual({ ok: true, revoked: true });
  });
});

describe("doPost rules with deformed bodies", () => {
  it("does not delete a rule when the id is an array wrapping its id", () => {
    // String([123]) used to become "123" and delete the rule whose
    // stored id is the number 123 (the bucket-list-gas wrong-delete).
    rulesValues.push([123, "07:00", "1", "", "", "", "/a"]);
    post({ ...sync, action: "delete_rule", id: [123] });
    expect(resultJson()).toMatchObject({ ok: false });
    expect(rulesValues).toHaveLength(2);
  });

  it("still deletes by a numeric id (Sheets stores bare ids as numbers)", () => {
    rulesValues.push([123, "07:00", "1", "", "", "", "/a"]);
    post({ ...sync, action: "delete_rule", id: 123 });
    expect(resultJson()).toMatchObject({ ok: true, deleted: 1 });
    expect(rulesValues).toHaveLength(1);
  });

  it("answers 'missing rule id' for an unstringifiable id instead of throwing", () => {
    post({ ...sync, action: "delete_rule", id: unstringifiable });
    expect(resultJson()).toEqual({ ok: false, error: "missing rule id" });
  });

  it("rejects invented weekdays from a set_rule body", () => {
    post({ ...sync, action: "set_rule", time: "07:00", weekdays: [true], commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    expect(rulesValues).toHaveLength(1);
  });

  it("stores a speech fallback instead of throwing on unstringifiable commands", () => {
    post({ ...sync, action: "set_rule", time: "07:00", commands: [unstringifiable] });
    expect(resultJson().ok).toBe(true);
    expect(resultJson().rule.commands).toHaveLength(1);
    expect(resultJson().rule.commands[0]).not.toContain("[object Object]");
  });

  it("rejects an unstringifiable until instead of throwing", () => {
    post({ ...sync, action: "set_rule", time: "07:00", until: unstringifiable, commands: ["/a"] });
    expect(resultJson()).toEqual({ ok: false, error: "invalid until date: use YYYYMMDD" });
    expect(rulesValues).toHaveLength(1);
  });

  it("accepts a rule whose text carries a lone surrogate", () => {
    postRaw(
      '{"sync_token":"test-sync-token","action":"set_rule","time":"07:00","text":"\\ud800"}'
    );
    expect(resultJson().ok).toBe(true);
  });

  it("lists the good rules when one stored row is deformed", () => {
    rulesValues.push(["bad", "07:30", "1", "", "", "", unstringifiable]);
    rulesValues.push(["good", "08:00", "1", "", "", "", "/a"]);
    post({ ...auth, action: "list_rules" });
    expect(resultJson().ok).toBe(true);
    expect(resultJson().rules.map((r) => r.id)).toEqual(["bad", "good"]);
    expect(resultJson().rules[0].commands).toEqual([]);
    expect(resultJson().rules[1].commands).toEqual(["/a"]);
  });
});
