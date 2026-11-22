import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  HOMECOMING_MESSAGE_,
  homecomingAnnounceMs_,
  jstParts_,
  processPresence_,
} from "../src/Code.js";

// ---- Mocks for Google Apps Script globals (same shape as
// test/presence.spec.js, plus a calendar-gas route) ----
let props;
let fired;
let presenceResponse;
let calendarResponse;
let calendarStatus;

const BASE_PROPS = {
  API_BASE_URL: "http://home-api.example",
  HOME_API_TOKEN: "dummy-home-token",
  DEFAULT_SPEECH_PATH: "/speak/-h/192.0.2.1/-v/60/-s/",
  PRESENCE_GAS_URL: "https://presence.example/exec",
  PRESENCE_API_TOKEN: "dummy-presence-token",
  PRESENCE_SPEECH_PATH: "/google-home-speaker-wrapper/-h/192.0.2.99/-s/",
  CALENDAR_GAS_URL: "https://calendar.example/exec",
  CALENDAR_API_TOKEN: "dummy-calendar-token",
  HOMECOMING_SPEECH_PATH: "/google-home-speaker-wrapper/-h/192.0.2.50/-s/",
  HOMECOMING_DRY_RUN: "1",
};

global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn((name) => props[name] || null),
    setProperty: vi.fn((name, value) => {
      props[name] = value;
    }),
  })),
};

global.UrlFetchApp = {
  fetch: vi.fn((url) => {
    const u = String(url);
    if (u.startsWith("https://presence.example/")) {
      return {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify(presenceResponse),
      };
    }
    if (u.startsWith("https://calendar.example/")) {
      return {
        getResponseCode: () => calendarStatus,
        getContentText: () => JSON.stringify(calendarResponse),
      };
    }
    fired.push(u);
    return { getResponseCode: () => 200, getContentText: () => "OK" };
  }),
};

global.Logger = { log: vi.fn() };
global.LockService = {
  getScriptLock: vi.fn(() => ({ waitLock: vi.fn(), releaseLock: vi.fn() })),
};
global.ContentService = {
  createTextOutput: vi.fn(() => ({
    setMimeType: vi.fn(),
    setContent: vi.fn(),
  })),
  MimeType: { JSON: "JSON", JAVASCRIPT: "JAVASCRIPT" },
};
global.Utilities = {
  DigestAlgorithm: { SHA_256: "SHA_256" },
  computeDigest: vi.fn(() => []),
  getUuid: vi.fn(() => "00000000-0000-4000-8000-000000000001"),
};
global.SpreadsheetApp = {
  getActiveSheet: vi.fn(() => ({
    getDataRange: vi.fn(() => ({ getValues: vi.fn(() => []) })),
  })),
  getActiveSpreadsheet: vi.fn(() => ({
    getActiveSheet: vi.fn(() => ({
      getDataRange: vi.fn(() => ({ getValues: vi.fn(() => []) })),
    })),
    getSheetByName: vi.fn(() => null),
    insertSheet: vi.fn(() => ({
      appendRow: vi.fn(),
      getDataRange: vi.fn(() => ({ getValues: vi.fn(() => []) })),
    })),
  })),
  openById: vi.fn(),
};

// ---- Helpers ----
// A JST wall-clock time on 2026-11-24 as a Date.
const jst = (h, m = 0) => new Date(Date.UTC(2026, 10, 24, h - 9, m, 0, 0));
const iso = (d) => d.toISOString();

const atWork = (since) => ({
  person: "freddiefujiwara",
  present: true,
  place: "work",
  action: "arrive",
  since: iso(since),
  ageSec: 60,
});

const leftWork = (since, ageSec) => ({
  person: "freddiefujiwara",
  present: false,
  place: "work",
  action: "leave",
  since: iso(since),
  ageSec,
});

const arrivedHome = (since, ageSec = 60) => ({
  person: "freddiefujiwara",
  present: true,
  place: "home",
  action: "arrive",
  since: iso(since),
  ageSec,
});

const step = (now, presence) => {
  presenceResponse = presence;
  return processPresence_(now);
};

const history = () => {
  const raw = props.PRESENCE_STATE;
  if (!raw) return [];
  const state = JSON.parse(raw);
  return (state.homecoming && state.homecoming.history) || [];
};
const types = () => history().map((e) => e.type);

const noPlan = { ok: true, date: "2026-11-24", hasNightPlan: false, nightPlanEnd: null, lastWorkEventEnd: null };

beforeEach(() => {
  props = { ...BASE_PROPS };
  fired = [];
  calendarStatus = 200;
  calendarResponse = { ...noPlan };
  presenceResponse = atWork(jst(9, 0));
});

describe("jstParts_", () => {
  it("converts a UTC instant to JST date and minutes", () => {
    expect(jstParts_(new Date("2026-11-24T09:30:00.000Z"))).toEqual({
      date: "2026-11-24",
      minutes: 18 * 60 + 30,
    });
    expect(jstParts_(new Date("2026-11-24T15:30:00.000Z"))).toEqual({
      date: "2026-11-25",
      minutes: 30,
    });
  });
});

describe("homecomingAnnounceMs_", () => {
  const leave = jst(18, 0).getTime();
  it("is leave + 20 minutes with no plans", () => {
    expect(homecomingAnnounceMs_(leave, "2026-11-24", noPlan)).toBe(jst(18, 20).getTime());
  });
  it("holds until daytime commitments end", () => {
    const digest = { ...noPlan, lastWorkEventEnd: iso(jst(18, 45)) };
    expect(homecomingAnnounceMs_(leave, "2026-11-24", digest)).toBe(jst(18, 45).getTime());
  });
  it("shifts to night plan end + 30 minutes", () => {
    const digest = { ...noPlan, hasNightPlan: true, nightPlanEnd: iso(jst(21, 0)) };
    expect(homecomingAnnounceMs_(leave, "2026-11-24", digest)).toBe(jst(21, 30).getTime());
  });
  it("is null when the shifted time reaches the 22:30 cutoff", () => {
    const digest = { ...noPlan, hasNightPlan: true, nightPlanEnd: iso(jst(22, 0)) };
    expect(homecomingAnnounceMs_(leave, "2026-11-24", digest)).toBe(null);
    const justIn = { ...noPlan, hasNightPlan: true, nightPlanEnd: iso(jst(21, 59)) };
    expect(homecomingAnnounceMs_(leave, "2026-11-24", justIn)).toBe(jst(22, 29).getTime());
  });
});

describe("processHomecoming_ via processPresence_", () => {
  it("ignores a daytime walk", () => {
    step(jst(9, 1), atWork(jst(9, 0)));
    step(jst(12, 1), leftWork(jst(12, 0), 60));
    step(jst(12, 21), leftWork(jst(12, 0), 21 * 60));
    expect(types()).not.toContain("candidate");
    expect(types()).not.toContain("announce");
  });

  it("cancels the candidate when the person returns to work (evening walk)", () => {
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(17, 11), leftWork(jst(17, 10), 60));
    expect(types()).toContain("candidate");
    step(jst(17, 26), atWork(jst(17, 25)));
    const cancelled = history().find((e) => e.type === "cancelled");
    expect(cancelled.reason).toBe("work/arrive");
    step(jst(17, 35), atWork(jst(17, 25)));
    expect(types()).not.toContain("announce");
  });

  it("records a would-be announcement in DRY_RUN, once per day", () => {
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 11), leftWork(jst(18, 10), 60));
    step(jst(18, 29), leftWork(jst(18, 10), 19 * 60));
    expect(types()).not.toContain("announce");
    const res = step(jst(18, 30), leftWork(jst(18, 10), 20 * 60));
    expect(res.homecoming.announced).toBe(true);
    expect(res.homecoming.dryRun).toBe(true);
    expect(fired).toEqual([]); // DRY_RUN never touches a device
    const announce = history().find((e) => e.type === "announce");
    expect(announce.dryRun).toBe(true);
    expect(announce.hasNightPlan).toBe(false);
    step(jst(18, 35), leftWork(jst(18, 10), 25 * 60));
    expect(history().filter((e) => e.type === "announce")).toHaveLength(1);
  });

  it("shifts the announcement past a night plan", () => {
    calendarResponse = { ...noPlan, hasNightPlan: true, nightPlanEnd: iso(jst(21, 0)) };
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 1), leftWork(jst(18, 0), 60));
    step(jst(18, 20), leftWork(jst(18, 0), 20 * 60));
    expect(types()).not.toContain("announce");
    step(jst(21, 29), leftWork(jst(18, 0), 3 * 60 * 60 + 29 * 60));
    expect(types()).not.toContain("announce");
    step(jst(21, 30), leftWork(jst(18, 0), 3 * 60 * 60 + 30 * 60));
    expect(types()).toContain("announce");
  });

  it("never announces when the night plan ends too late", () => {
    calendarResponse = { ...noPlan, hasNightPlan: true, nightPlanEnd: iso(jst(22, 30)) };
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 1), leftWork(jst(18, 0), 60));
    step(jst(18, 20), leftWork(jst(18, 0), 20 * 60));
    const digest = history().find((e) => e.type === "digest");
    expect(digest.announceMs).toBe(null);
    step(jst(21, 0), leftWork(jst(18, 0), 3 * 60 * 60));
    expect(types()).not.toContain("announce");
  });

  it("holds the announcement until daytime commitments end", () => {
    calendarResponse = { ...noPlan, lastWorkEventEnd: iso(jst(18, 30)) };
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(17, 31), leftWork(jst(17, 30), 60));
    step(jst(17, 50), leftWork(jst(17, 30), 20 * 60));
    expect(types()).not.toContain("announce");
    step(jst(18, 30), leftWork(jst(17, 30), 60 * 60));
    expect(types()).toContain("announce");
  });

  it("does not start a candidate without a work arrival that day", () => {
    step(jst(18, 11), leftWork(jst(18, 10), 60));
    step(jst(18, 35), leftWork(jst(18, 10), 25 * 60));
    expect(types()).not.toContain("candidate");
    expect(types()).toContain("leave-observed");
    expect(types()).not.toContain("announce");
  });

  it("waits out an unreadable calendar, then announces", () => {
    calendarResponse = { ok: false, error: "boom" };
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 11), leftWork(jst(18, 10), 60));
    const waiting = step(jst(18, 30), leftWork(jst(18, 10), 20 * 60));
    expect(waiting.homecoming.waiting).toBe("calendar unknown");
    expect(types()).not.toContain("announce");
    calendarResponse = { ...noPlan };
    step(jst(18, 31), leftWork(jst(18, 10), 21 * 60));
    expect(types()).toContain("announce");
  });

  it("cancels the candidate when the person arrives home", () => {
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 11), leftWork(jst(18, 10), 60));
    step(jst(18, 41), arrivedHome(jst(18, 40)));
    const cancelled = history().find((e) => e.type === "cancelled");
    expect(cancelled.reason).toBe("home/arrive");
    expect(types()).not.toContain("announce");
  });

  it("fires the Nest speech command when not in DRY_RUN", () => {
    delete props.HOMECOMING_DRY_RUN;
    step(jst(8, 56), atWork(jst(8, 55)));
    step(jst(18, 11), leftWork(jst(18, 10), 60));
    step(jst(18, 30), leftWork(jst(18, 10), 20 * 60));
    expect(fired).toHaveLength(1);
    expect(fired[0]).toContain("192.0.2.50");
    expect(fired[0]).toContain(encodeURIComponent(HOMECOMING_MESSAGE_));
    const announce = history().find((e) => e.type === "announce");
    expect(announce.dryRun).toBe(false);
  });
});
