import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildSpeechText_,
  classify_,
  collectRequestCommands_,
  defaultSpeechCommand_,
  doGet,
  doPost,
  myFunction,
  parseDateKey_,
  parseTarget_,
  parseTimeOfDay_,
  parseWeekdays_,
  processRules_,
  toUrl_,
} from "../src/Code.js";

// ---- Mocks for Google Apps Script globals ----
let sheetValues;
let appended;
let deletedRows;
let rangeSets;
let rangeClears;
let fetched;
let outputContent;
let outputMime;

const mockSheet = {
  getDataRange: vi.fn(() => ({ getValues: vi.fn(() => sheetValues) })),
  appendRow: vi.fn((row) => appended.push(row)),
  deleteRow: vi.fn((n) => deletedRows.push(n)),
  getRange: vi.fn(() => ({
    clearContent: vi.fn(() => rangeClears.push(1)),
    setValues: vi.fn((v) => rangeSets.push(v)),
  })),
};

global.SpreadsheetApp = {
  getActiveSheet: vi.fn(() => mockSheet),
  getActiveSpreadsheet: vi.fn(() => mockSpreadsheet),
};

// ---- Share session mocks (mirror the health-gas test stubs) ----
let uuidCounter = 0;
let shareValues;
let shareSheetExists;
let shareInserted;

const mockShareSheet = {
  getDataRange: vi.fn(() => ({ getValues: vi.fn(() => shareValues) })),
  appendRow: vi.fn((row) => shareValues.push([...row])),
  getRange: vi.fn((row, col) => ({
    setValue: vi.fn((value) => {
      shareValues[row - 1][col - 1] = value;
    }),
  })),
};

let rulesValues;
let rulesSheetExists;

const mockRulesSheet = {
  getDataRange: vi.fn(() => ({ getValues: vi.fn(() => rulesValues) })),
  appendRow: vi.fn((row) => rulesValues.push([...row])),
  deleteRow: vi.fn((n) => rulesValues.splice(n - 1, 1)),
  getRange: vi.fn((row, col) => ({
    setValue: vi.fn((value) => {
      rulesValues[row - 1][col - 1] = value;
    }),
  })),
};

const mockSpreadsheet = {
  getSheetByName: vi.fn((name) => {
    if (name === "share_sessions") return shareSheetExists ? mockShareSheet : null;
    if (name === "rules") return rulesSheetExists ? mockRulesSheet : null;
    return null;
  }),
  insertSheet: vi.fn((name) => {
    if (name === "rules") {
      rulesSheetExists = true;
      return mockRulesSheet;
    }
    shareSheetExists = true;
    shareInserted += 1;
    return mockShareSheet;
  }),
};

global.Utilities = {
  DigestAlgorithm: { SHA_256: "SHA_256" },
  computeDigest: vi.fn((algo, str) =>
    Array.from({ length: 32 }, (_, i) => String(str).charCodeAt(i % String(str).length) % 128)
  ),
  getUuid: vi.fn(
    () => `00000000-0000-4000-8000-${String(++uuidCounter).padStart(12, "0")}`
  ),
};

const shareHashOf = (token) =>
  Array.from({ length: 32 }, (_, i) => token.charCodeAt(i % token.length) % 128)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
global.ContentService = {
  createTextOutput: vi.fn(() => ({
    setMimeType: vi.fn((m) => {
      outputMime = m;
    }),
    setContent: vi.fn((c) => {
      outputContent = c;
    }),
  })),
  MimeType: { JSON: "JSON", JAVASCRIPT: "JAVASCRIPT" },
};
global.LockService = {
  getScriptLock: vi.fn(() => ({ waitLock: vi.fn(), releaseLock: vi.fn() })),
};
let tokenInfo;
global.UrlFetchApp = {
  fetch: vi.fn((url) => {
    if (String(url).startsWith("https://oauth2.googleapis.com/tokeninfo")) {
      return {
        getResponseCode: () => tokenInfo.code,
        getContentText: () => JSON.stringify(tokenInfo.payload),
      };
    }
    fetched.push(url);
    return { getResponseCode: () => 200, getContentText: () => "OK" };
  }),
};
global.Logger = { log: vi.fn() };
global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn(
      (name) =>
        ({
          API_BASE_URL: "http://home-api.example",
          DEFAULT_SPEECH_PATH: "/speak/-h/192.0.2.1/-v/60/-s/",
          AVAILABLE_GMAILS: "owner@example.com",
          GOOGLE_OAUTH_CLIENT_ID: "dummy-client-id",
        })[name] || null
    ),
  })),
};

beforeEach(() => {
  sheetValues = [];
  appended = [];
  deletedRows = [];
  rangeSets = [];
  rangeClears = [];
  fetched = [];
  outputContent = null;
  outputMime = null;
  uuidCounter = 0;
  shareSheetExists = true;
  shareInserted = 0;
  shareValues = [
    ["token_hash", "issued_at", "expires_at", "revoked_at", "access_count", "last_accessed_at"],
  ];
  rulesSheetExists = true;
  rulesValues = [["id", "time", "weekdays", "until", "last_fired", "commands"]];
  tokenInfo = {
    code: 200,
    payload: {
      iss: "https://accounts.google.com",
      aud: "dummy-client-id",
      exp: String(Math.floor(Date.now() / 1000) + 3600),
      email_verified: "true",
      email: "owner@example.com",
    },
  };
  vi.clearAllMocks();
});

const ev = (parameter, parameters) => ({ parameter, parameters: parameters || {} });
const resultJson = () => JSON.parse(outputContent);

describe("parseTarget_", () => {
  const now = new Date(2026, 10, 2, 1, 0, 0);
  it("reads HH:MM as today", () => {
    expect(parseTarget_("04:50", now)).toEqual(new Date(2026, 10, 2, 4, 50, 0));
  });
  it("reads H:MM:SS as today with seconds", () => {
    expect(parseTarget_("4:50:30", now)).toEqual(new Date(2026, 10, 2, 4, 50, 30));
  });
  it("reads YYYYMMDD-HH:MM", () => {
    expect(parseTarget_("20270109-10:25", now)).toEqual(new Date(2027, 0, 9, 10, 25, 0));
  });
  it("rejects impossible dates and times", () => {
    expect(parseTarget_("20270230-10:00", now)).toBeNull();
    expect(parseTarget_("25:00", now)).toBeNull();
    expect(parseTarget_("04:60", now)).toBeNull();
    expect(parseTarget_("tomorrow", now)).toBeNull();
    expect(parseTarget_("", now)).toBeNull();
  });
});

describe("speech helpers", () => {
  it("builds the legacy speech text", () => {
    expect(buildSpeechText_(new Date(2026, 10, 2, 7, 30), "おはよう")).toBe(
      "7時30分です。おはよう"
    );
    expect(buildSpeechText_(new Date(2026, 10, 2, 13, 0), "昼です")).toBe(
      "1時ちょうどです。昼です"
    );
  });
  it("nudges instead when the text is empty", () => {
    expect(buildSpeechText_(new Date(2026, 10, 2, 7, 30), "")).toBe(
      "7時30分です。時間ですよ。時間ですよ。そろそろ時間ですよ。"
    );
  });
  it("builds the default speech command", () => {
    const cmd = defaultSpeechCommand_(new Date(2026, 10, 2, 7, 30), "おはよう");
    expect(cmd.startsWith("/speak/-h/192.0.2.1/-v/60/-s/")).toBe(true);
    expect(decodeURIComponent(cmd.split("/-s/")[1])).toBe("7時30分です。おはよう");
  });
});

describe("collectRequestCommands_", () => {
  const target = new Date(2026, 10, 2, 7, 30);
  it("collects repeated cmd parameters", () => {
    const e = ev({}, { cmd: ["/hue/lights/off", " /x ", ""] });
    expect(collectRequestCommands_(e, target)).toEqual(["/hue/lights/off", "/x"]);
  });
  it("falls back to a default speech command for legacy text", () => {
    const e = ev({ text: "おはよう" });
    const cmds = collectRequestCommands_(e, target);
    expect(cmds).toHaveLength(1);
    expect(cmds[0]).toContain("/speak/-h/192.0.2.1/");
  });
  it("falls back to the time announcement when nothing is given", () => {
    const cmds = collectRequestCommands_(ev({}), target);
    expect(cmds).toHaveLength(1);
    expect(decodeURIComponent(cmds[0].split("/-s/")[1])).toBe(
      "7時30分です。時間ですよ。時間ですよ。そろそろ時間ですよ。"
    );
  });
});

describe("doGet", () => {
  it("registers a new alarm", () => {
    const e = ev({ time: "29990101-04:50", cmd: "/hue/lights/off" }, { cmd: ["/hue/lights/off"] });
    doGet(e);
    expect(resultJson()).toEqual({ ok: true, time: "29990101-04:50", commands: 1 });
    expect(appended).toHaveLength(1);
    expect(appended[0][1]).toBe("/hue/lights/off");
    expect(appended[0][0]).toEqual(new Date(2999, 0, 1, 4, 50, 0));
  });
  it("rejects a past time", () => {
    doGet(ev({ time: "20200101-04:50", cmd: "/x" }, { cmd: ["/x"] }));
    expect(resultJson().ok).toBe(false);
    expect(resultJson().error).toContain("past");
    expect(appended).toHaveLength(0);
  });
  it("rejects more than 10 commands", () => {
    const cmds = Array.from({ length: 11 }, (_, i) => `/c${i}`);
    doGet(ev({ time: "29990101-04:50" }, { cmd: cmds }));
    expect(resultJson().ok).toBe(false);
    expect(resultJson().error).toContain("10");
    expect(appended).toHaveLength(0);
  });
  it("rejects a bad format", () => {
    doGet(ev({ time: "soon", cmd: "/x" }, { cmd: ["/x"] }));
    expect(resultJson().ok).toBe(false);
  });
  it("registers a time announcement alarm when neither cmd nor text is given", () => {
    doGet(ev({ time: "29990101-04:50" }));
    expect(resultJson()).toEqual({ ok: true, time: "29990101-04:50", commands: 1 });
    expect(appended).toHaveLength(1);
    expect(appended[0][1]).toContain("/speak/-h/192.0.2.1/");
    expect(decodeURIComponent(appended[0][1].split("/-s/")[1])).toBe(
      "4時50分です。時間ですよ。時間ですよ。そろそろ時間ですよ。"
    );
  });
  it("overwrites an existing row for the same datetime", () => {
    const existing = new Date(2999, 0, 1, 4, 50, 0);
    sheetValues = [[existing, "/old"]];
    doGet(ev({ time: "29990101-04:50", cmd: "/new" }, { cmd: ["/new"] }));
    expect(resultJson().ok).toBe(true);
    expect(appended).toHaveLength(0);
    expect(rangeClears).toHaveLength(1);
    expect(rangeSets[0][0][1]).toBe("/new");
  });
  it("answers JSONP when a callback is given", () => {
    doGet(ev({ time: "29990101-04:50", cmd: "/x", callback: "_cb" }, { cmd: ["/x"] }));
    expect(outputMime).toBe("JAVASCRIPT");
    expect(outputContent.startsWith("_cb&&_cb(")).toBe(true);
  });
});

describe("classify_ and toUrl_", () => {
  const now = new Date(2026, 10, 2, 4, 50, 0);
  it("classifies by the 1 minute window", () => {
    expect(classify_(new Date(2026, 10, 2, 4, 50, 30), now)).toBe("fire");
    expect(classify_(new Date(2026, 10, 2, 4, 49, 30), now)).toBe("fire");
    expect(classify_(new Date(2026, 10, 2, 4, 48, 0), now)).toBe("cleanup");
    expect(classify_(new Date(2026, 10, 2, 4, 52, 0), now)).toBe("keep");
  });
  it("maps commands to URLs", () => {
    expect(toUrl_("/hue/lights/off")).toBe("http://home-api.example/hue/lights/off");
    expect(toUrl_("https://example.com/x")).toBe("https://example.com/x");
    expect(toUrl_("おはよう")).toBeNull();
  });
  it("maps the space form to the same URL as the path form", () => {
    expect(toUrl_("switchbot-command -d X -c turnOff")).toBe(
      "http://home-api.example/switchbot-command/-d/X/-c/turnOff"
    );
    expect(toUrl_("hue lights off")).toBe("http://home-api.example/hue/lights/off");
  });
  it("does not mistake speech text for a space-form command", () => {
    expect(toUrl_("おはよう ございます")).toBeNull();
    expect(toUrl_("test")).toBeNull();
  });
});

describe("myFunction", () => {
  it("fires a due row's commands in order and deletes the row", () => {
    sheetValues = [[new Date(), "/hue/lights/off", "/switchbot-command/-d/X/-c/turnOff"]];
    myFunction();
    expect(fetched).toEqual([
      "http://home-api.example/hue/lights/off",
      "http://home-api.example/switchbot-command/-d/X/-c/turnOff",
    ]);
    expect(deletedRows).toEqual([1]);
  });
  it("deletes an overdue row without firing", () => {
    sheetValues = [[new Date(Date.now() - 10 * 60 * 1000), "/hue/lights/off"]];
    myFunction();
    expect(fetched).toEqual([]);
    expect(deletedRows).toEqual([1]);
  });
  it("keeps a future row", () => {
    sheetValues = [[new Date(Date.now() + 10 * 60 * 1000), "/hue/lights/off"]];
    myFunction();
    expect(fetched).toEqual([]);
    expect(deletedRows).toEqual([]);
  });
  it("speaks a legacy [time, text] row with the default speech command", () => {
    const timeOnly = new Date();
    timeOnly.setFullYear(1899); // Sheets returns time-only cells like this
    sheetValues = [[timeOnly, "おはよう"]];
    myFunction();
    expect(fetched).toHaveLength(1);
    expect(fetched[0]).toContain("http://home-api.example/speak/-h/192.0.2.1/");
    expect(deletedRows).toEqual([1]);
  });
  it("handles mixed rows and deletes from the bottom up", () => {
    const due = new Date();
    const overdue = new Date(Date.now() - 30 * 60 * 1000);
    const future = new Date(Date.now() + 30 * 60 * 1000);
    sheetValues = [
      [overdue, "/a"],
      [future, "/b"],
      [due, "/c"],
    ];
    myFunction();
    expect(fetched).toEqual(["http://home-api.example/c"]);
    expect(deletedRows).toEqual([3, 1]);
  });
});

describe("doPost", () => {
  const post = (body) => doPost({ postData: { contents: JSON.stringify(body) } });
  const auth = { idToken: "dummy-token" };

  it("rejects a request without a valid token", () => {
    post({ action: "list" });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
  });

  it("rejects an email that is not allowed", () => {
    tokenInfo.payload.email = "someone-else@example.com";
    post({ ...auth, action: "list" });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(403);
  });

  it("rejects an expired token", () => {
    tokenInfo.payload.exp = String(Math.floor(Date.now() / 1000) - 60);
    post({ ...auth, action: "list" });
    expect(resultJson().ok).toBe(false);
  });

  it("lists the registered alarms, oldest first", () => {
    sheetValues = [
      [new Date(2999, 0, 1, 7, 30, 0), "/hue/lights/off"],
      [new Date(2999, 0, 1, 4, 50, 0), "/speak/x"],
      ["", ""],
    ];
    post({ ...auth, action: "list" });
    expect(resultJson()).toEqual({
      ok: true,
      alarms: [
        { time: "29990101-04:50", commands: ["/speak/x"] },
        { time: "29990101-07:30", commands: ["/hue/lights/off"] },
      ],
    });
  });

  it("registers an alarm with explicit commands", () => {
    post({ ...auth, action: "set", time: "29990101-04:50", commands: ["/a", " /b "] });
    expect(resultJson()).toEqual({ ok: true, time: "29990101-04:50", commands: 2 });
    expect(appended).toHaveLength(1);
    expect(appended[0].slice(1)).toEqual(["/a", "/b"]);
  });

  it("registers a speech alarm from text, or a time announcement without", () => {
    post({ ...auth, action: "set", time: "29990101-04:50", text: "おはよう" });
    expect(resultJson().ok).toBe(true);
    expect(decodeURIComponent(appended[0][1].split("/-s/")[1])).toBe("4時50分です。おはよう");
    post({ ...auth, action: "set", time: "29990101-05:00" });
    expect(resultJson().ok).toBe(true);
    expect(decodeURIComponent(appended[1][1].split("/-s/")[1])).toBe(
      "5時ちょうどです。時間ですよ。時間ですよ。そろそろ時間ですよ。"
    );
  });

  it("rejects a past time, too many commands, and unknown actions", () => {
    post({ ...auth, action: "set", time: "20200101-04:50", commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    const cmds = Array.from({ length: 11 }, (_, i) => `/c${i}`);
    post({ ...auth, action: "set", time: "29990101-04:50", commands: cmds });
    expect(resultJson().ok).toBe(false);
    post({ ...auth, action: "bogus" });
    expect(resultJson().ok).toBe(false);
    expect(appended).toHaveLength(0);
  });

  it("deletes the alarm for the given minute", () => {
    sheetValues = [
      [new Date(2999, 0, 1, 4, 50, 0), "/a"],
      [new Date(2999, 0, 1, 7, 30, 0), "/b"],
    ];
    post({ ...auth, action: "delete", time: "29990101-07:30" });
    expect(resultJson()).toEqual({ ok: true, deleted: 1 });
    expect(deletedRows).toEqual([2]);
    post({ ...auth, action: "delete", time: "29990101-09:00" });
    expect(resultJson()).toEqual({ ok: true, deleted: 0 });
  });

  it("rejects invalid JSON", () => {
    doPost({ postData: { contents: "not json" } });
    expect(resultJson().ok).toBe(false);
  });
});

describe("doPost share sessions", () => {
  const post = (body) => doPost({ postData: { contents: JSON.stringify(body) } });
  const auth = { idToken: "dummy-token" };
  const issueToken = () => {
    post({ ...auth, action: "issue_share" });
    return resultJson().share_token;
  };

  it("requires an ID token to issue a share session", () => {
    post({ action: "issue_share" });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
    expect(shareValues).toHaveLength(1);
  });

  it("issues a 30-day share session and stores only its hash", () => {
    const before = Date.now();
    post({ ...auth, action: "issue_share" });
    const result = resultJson();
    expect(result.ok).toBe(true);
    expect(result.ttl_seconds).toBe(30 * 24 * 60 * 60);
    expect(result.share_token).toMatch(/^[0-9a-f]{64}$/);
    expect(Date.parse(result.expires_at) - before).toBeGreaterThan(29 * 24 * 60 * 60 * 1000);
    expect(Date.parse(result.expires_at) - before).toBeLessThanOrEqual(
      30 * 24 * 60 * 60 * 1000 + 5000
    );
    expect(shareValues).toHaveLength(2);
    expect(shareValues[1][0]).toBe(shareHashOf(result.share_token));
    expect(shareValues[1][0]).not.toBe(result.share_token);
    expect(shareValues[1][4]).toBe(0);
  });

  it("creates the share sheet with headers when it does not exist yet", () => {
    shareValues = [];
    shareSheetExists = false;
    post({ ...auth, action: "issue_share" });
    expect(resultJson().ok).toBe(true);
    expect(shareInserted).toBe(1);
    expect(shareValues[0]).toEqual([
      "token_hash",
      "issued_at",
      "expires_at",
      "revoked_at",
      "access_count",
      "last_accessed_at",
    ]);
    expect(shareValues).toHaveLength(2);
  });

  it("revokes the previous session when a new one is issued", () => {
    const first = issueToken();
    const second = issueToken();
    expect(first).not.toBe(second);
    expect(shareValues[1][3]).not.toBe("");
    expect(shareValues[2][3]).toBe("");
    post({ action: "list", share_token: first });
    expect(resultJson()).toEqual({
      ok: false,
      status: 401,
      error: "share session expired or invalid",
    });
  });

  it("lists alarms with a valid share token and counts the access", () => {
    const token = issueToken();
    sheetValues = [[new Date(2999, 0, 1, 4, 50, 0), "/speak/x"]];
    post({ action: "list", share_token: token });
    const result = resultJson();
    expect(result.ok).toBe(true);
    expect(result.alarms).toEqual([{ time: "29990101-04:50", commands: ["/speak/x"] }]);
    expect(result._share.expires_at).toBe(shareValues[1][2]);
    expect(shareValues[1][4]).toBe(1);
    expect(shareValues[1][5]).not.toBe("");
  });

  it("lets a share token register and delete alarms (owner-equivalent)", () => {
    const token = issueToken();
    post({ action: "set", share_token: token, time: "29990101-04:50", commands: ["/a"] });
    expect(resultJson().ok).toBe(true);
    expect(appended).toHaveLength(1);
    expect(appended[0].slice(1)).toEqual(["/a"]);
    sheetValues = [[new Date(2999, 0, 1, 4, 50, 0), "/a"]];
    post({ action: "delete", share_token: token, time: "29990101-04:50" });
    expect(resultJson()).toEqual({
      ok: true,
      deleted: 1,
      _share: { expires_at: shareValues[1][2] },
    });
    expect(deletedRows).toEqual([1]);
  });

  it("rejects an expired share token", () => {
    const token = "expired-token";
    shareValues.push([
      shareHashOf(token),
      new Date(Date.now() - 3600_000).toISOString(),
      new Date(Date.now() - 1800_000).toISOString(),
      "",
      0,
      "",
    ]);
    post({ action: "list", share_token: token });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
  });

  it("rejects a revoked share token", () => {
    const token = "revoked-token";
    shareValues.push([
      shareHashOf(token),
      new Date().toISOString(),
      new Date(Date.now() + 1800_000).toISOString(),
      new Date().toISOString(),
      0,
      "",
    ]);
    post({ action: "list", share_token: token });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
  });

  it("rejects an unknown share token", () => {
    post({ action: "list", share_token: "no-such-token" });
    expect(resultJson()).toEqual({
      ok: false,
      status: 401,
      error: "share session expired or invalid",
    });
  });

  it("revokes the active session via the revoke action", () => {
    const token = issueToken();
    post({ ...auth, action: "revoke_share" });
    expect(resultJson()).toEqual({ ok: true, revoked: true });
    post({ action: "list", share_token: token });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
  });

  it("does not let a share token issue more share sessions", () => {
    const token = issueToken();
    post({ action: "issue_share", share_token: token });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().status).toBe(401);
    expect(shareValues).toHaveLength(2);
  });
});

describe("rule parsing helpers", () => {
  it("parses a time of day", () => {
    expect(parseTimeOfDay_("7:30")).toEqual({ hour: 7, minute: 30 });
    expect(parseTimeOfDay_("07:05:00")).toEqual({ hour: 7, minute: 5 });
    expect(parseTimeOfDay_("24:00")).toBeNull();
    expect(parseTimeOfDay_("7:60")).toBeNull();
    expect(parseTimeOfDay_("abc")).toBeNull();
    expect(parseTimeOfDay_("20261102-07:30")).toBeNull();
  });

  it("parses weekdays, defaulting to every day", () => {
    expect(parseWeekdays_(undefined)).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(parseWeekdays_("")).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(parseWeekdays_([5, 1, 1])).toEqual([1, 5]);
    expect(parseWeekdays_("6,0")).toEqual([0, 6]);
    expect(parseWeekdays_([7])).toBeNull();
    expect(parseWeekdays_(["x"])).toBeNull();
  });

  it("parses until dates in several shapes", () => {
    expect(parseDateKey_(undefined)).toBe("");
    expect(parseDateKey_("20261231")).toBe("20261231");
    expect(parseDateKey_("2026-12-31")).toBe("20261231");
    expect(parseDateKey_(20261231)).toBe("20261231");
    expect(parseDateKey_(new Date(2026, 11, 31))).toBe("20261231");
    expect(parseDateKey_("20260230")).toBeNull();
    expect(parseDateKey_("tomorrow")).toBeNull();
  });
});

describe("doPost rules", () => {
  const post = (body) => doPost({ postData: { contents: JSON.stringify(body) } });
  const auth = { idToken: "dummy-token" };

  it("adds, lists, and deletes a rule", () => {
    post({ ...auth, action: "set_rule", time: "07:30", weekdays: [1, 2, 3, 4, 5], commands: ["/a"] });
    const created = resultJson();
    expect(created.ok).toBe(true);
    expect(created.rule.time).toBe("07:30");
    expect(created.rule.weekdays).toEqual([1, 2, 3, 4, 5]);
    expect(created.rule.until).toBeNull();
    expect(rulesValues).toHaveLength(2);

    post({ ...auth, action: "list_rules" });
    expect(resultJson().rules).toEqual([created.rule]);

    post({ ...auth, action: "delete_rule", id: created.rule.id });
    expect(resultJson()).toEqual({ ok: true, deleted: 1 });
    expect(rulesValues).toHaveLength(1);
    post({ ...auth, action: "delete_rule", id: created.rule.id });
    expect(resultJson()).toEqual({ ok: true, deleted: 0 });
  });

  it("turns exclude_weekdays into the active set", () => {
    post({ ...auth, action: "set_rule", time: "06:00", exclude_weekdays: [0, 6], text: "おはよう" });
    expect(resultJson().rule.weekdays).toEqual([1, 2, 3, 4, 5]);
    post({ ...auth, action: "set_rule", time: "06:00", exclude_weekdays: [0, 1, 2, 3, 4, 5, 6] });
    expect(resultJson().ok).toBe(false);
  });

  it("rejects an invalid time, weekdays, and a past until", () => {
    post({ ...auth, action: "set_rule", time: "25:00", commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    post({ ...auth, action: "set_rule", time: "07:00", weekdays: [9], commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    post({ ...auth, action: "set_rule", time: "07:00", until: "20200101", commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    expect(rulesValues).toHaveLength(1);
  });

  it("accepts an until date and reports it", () => {
    post({ ...auth, action: "set_rule", time: "07:00", until: "29991231", commands: ["/a"] });
    expect(resultJson().rule.until).toBe("29991231");
  });

  it("enforces the rule limit", () => {
    for (let i = 0; i < 20; i += 1) {
      rulesValues.push([`id-${i}`, "07:00", "0,1,2,3,4,5,6", "", "", "/a"]);
    }
    post({ ...auth, action: "set_rule", time: "07:00", commands: ["/a"] });
    expect(resultJson().ok).toBe(false);
    expect(resultJson().error).toContain("too many rules");
  });
});

describe("processRules_", () => {
  const monday = new Date(2026, 10, 2, 7, 30, 10); // 2026-11-02 is a Monday

  it("fires a due rule once and marks last_fired", () => {
    rulesValues.push(["r1", "07:30", "0,1,2,3,4,5,6", "", "", "/hue/lights/off"]);
    processRules_(monday);
    expect(fetched).toEqual(["http://home-api.example/hue/lights/off"]);
    expect(rulesValues[1][4]).toBe("20261102");
    processRules_(new Date(2026, 10, 2, 7, 30, 40));
    expect(fetched).toHaveLength(1);
  });

  it("does not fire on an inactive weekday or outside the window", () => {
    rulesValues.push(["r1", "07:30", "2,3,4,5,6", "", "", "/a"]); // Mondays off
    processRules_(monday);
    expect(fetched).toHaveLength(0);
    rulesValues[1][2] = "0,1,2,3,4,5,6";
    processRules_(new Date(2026, 10, 2, 9, 0, 0));
    expect(fetched).toHaveLength(0);
  });

  it("deletes rules past their until date", () => {
    rulesValues.push(["r1", "07:30", "0,1,2,3,4,5,6", "20261101", "", "/a"]);
    processRules_(monday);
    expect(rulesValues).toHaveLength(1);
    expect(fetched).toHaveLength(0);
  });

  it("does nothing when the rules sheet does not exist", () => {
    rulesSheetExists = false;
    expect(() => processRules_(monday)).not.toThrow();
  });
});
