import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildSpeechText_,
  classify_,
  collectRequestCommands_,
  defaultSpeechCommand_,
  doGet,
  doPost,
  myFunction,
  parseTarget_,
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

global.SpreadsheetApp = { getActiveSheet: vi.fn(() => mockSheet) };
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
  it("falls back to a time-only announcement when nothing is given", () => {
    const cmds = collectRequestCommands_(ev({}), target);
    expect(cmds).toHaveLength(1);
    expect(decodeURIComponent(cmds[0].split("/-s/")[1])).toBe("7時30分です。");
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
  it("registers a time-only alarm when neither cmd nor text is given", () => {
    doGet(ev({ time: "29990101-04:50" }));
    expect(resultJson()).toEqual({ ok: true, time: "29990101-04:50", commands: 1 });
    expect(appended).toHaveLength(1);
    expect(appended[0][1]).toContain("/speak/-h/192.0.2.1/");
    expect(decodeURIComponent(appended[0][1].split("/-s/")[1])).toBe("4時50分です。");
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

  it("registers a speech alarm from text, or a time-only one without", () => {
    post({ ...auth, action: "set", time: "29990101-04:50", text: "おはよう" });
    expect(resultJson().ok).toBe(true);
    expect(decodeURIComponent(appended[0][1].split("/-s/")[1])).toBe("4時50分です。おはよう");
    post({ ...auth, action: "set", time: "29990101-05:00" });
    expect(resultJson().ok).toBe(true);
    expect(decodeURIComponent(appended[1][1].split("/-s/")[1])).toBe("5時ちょうどです。");
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
