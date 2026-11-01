import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildSpeechText_,
  classify_,
  collectRequestCommands_,
  defaultSpeechCommand_,
  doGet,
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
global.UrlFetchApp = {
  fetch: vi.fn((url) => {
    fetched.push(url);
    return { getResponseCode: () => 200, getContentText: () => "OK" };
  }),
};
global.Logger = { log: vi.fn() };

beforeEach(() => {
  sheetValues = [];
  appended = [];
  deletedRows = [];
  rangeSets = [];
  rangeClears = [];
  fetched = [];
  outputContent = null;
  outputMime = null;
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
  it("builds the default Nest speech command", () => {
    const cmd = defaultSpeechCommand_(new Date(2026, 10, 2, 7, 30), "おはよう");
    expect(cmd.startsWith("/google-home-speaker-wrapper/-h/192.168.1.22/-v/60/-s/")).toBe(true);
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
    expect(cmds[0]).toContain("/google-home-speaker-wrapper/");
  });
  it("returns empty when nothing is given", () => {
    expect(collectRequestCommands_(ev({}), target)).toEqual([]);
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
  it("rejects a bad format and a missing command", () => {
    doGet(ev({ time: "soon", cmd: "/x" }, { cmd: ["/x"] }));
    expect(resultJson().ok).toBe(false);
    doGet(ev({ time: "29990101-04:50" }));
    expect(resultJson().ok).toBe(false);
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
    expect(toUrl_("/hue/lights/off")).toBe("http://a.ze.gs/hue/lights/off");
    expect(toUrl_("https://example.com/x")).toBe("https://example.com/x");
    expect(toUrl_("おはよう")).toBeNull();
  });
});

describe("myFunction", () => {
  it("fires a due row's commands in order and deletes the row", () => {
    sheetValues = [[new Date(), "/hue/lights/off", "/switchbot-command/-d/X/-c/turnOff"]];
    myFunction();
    expect(fetched).toEqual([
      "http://a.ze.gs/hue/lights/off",
      "http://a.ze.gs/switchbot-command/-d/X/-c/turnOff",
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
  it("speaks a legacy [time, text] row on the Nest speaker", () => {
    const timeOnly = new Date();
    timeOnly.setFullYear(1899); // Sheets returns time-only cells like this
    sheetValues = [[timeOnly, "おはよう"]];
    myFunction();
    expect(fetched).toHaveLength(1);
    expect(fetched[0]).toContain("http://a.ze.gs/google-home-speaker-wrapper/-h/192.168.1.22/");
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
    expect(fetched).toEqual(["http://a.ze.gs/c"]);
    expect(deletedRows).toEqual([3, 1]);
  });
});
