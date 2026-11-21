import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  PRESENCE_ARRIVE_MESSAGE_,
  absenceSecondsFromEvents_,
  absenceSecondsFromState_,
  decideTatamiAcMode_,
  parseOffCommands_,
  parseTemperatureFeed_,
  processPresence_,
  shouldHandleLeave_,
  shouldNotifyArrive_,
} from "../src/Code.js";

// ---- Mocks for Google Apps Script globals ----
let props;
let fired;
let presenceResponse;
let eventsResponse;
let presenceStatus;
let tempResponse;
let tempStatus;

const BASE_PROPS = {
  API_BASE_URL: "http://home-api.example",
  HOME_API_TOKEN: "test-home-token",
  DEFAULT_SPEECH_PATH: "/speak/-h/192.0.2.1/-v/60/-s/",
  PRESENCE_GAS_URL: "https://presence.example/exec",
  PRESENCE_API_TOKEN: "test-presence-token",
  PRESENCE_SPEECH_PATH: "/google-home-speaker-wrapper/-h/192.0.2.99/-s/",
  PRESENCE_TATAMI_OFF_COMMANDS: JSON.stringify([
    "/dirigera/tatami/off",
    "/switchbot-ac/off",
    "/switchbot-circulator/off",
  ]),
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
  fetch: vi.fn((url, options) => {
    const u = String(url);
    if (u.startsWith("https://presence.example/")) {
      const isEvents = u.includes("events=1");
      const payload = isEvents ? eventsResponse : presenceResponse;
      return {
        getResponseCode: () => presenceStatus,
        getContentText: () => JSON.stringify(payload),
      };
    }
    if (u.startsWith("https://temp.example/")) {
      return {
        getResponseCode: () => tempStatus,
        getContentText: () => tempResponse,
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

const T_LEAVE = "2026-11-22T00:00:00.000Z";
const T_ARRIVE = "2026-11-22T00:40:00.000Z"; // 40 minutes later

const arrivePresence = (ageSec, since = T_ARRIVE) => ({
  person: "freddiefujiwara",
  present: true,
  place: "home",
  action: "arrive",
  since,
  ageSec,
});

const leavePresence = (ageSec, since = T_LEAVE) => ({
  person: "freddiefujiwara",
  present: false,
  place: "home",
  action: "leave",
  since,
  ageSec,
});

const tempFeed = (temperature, date = "11/22/2026 12:40:00 AM") =>
  `cb&&cb(${JSON.stringify({
    conditions: [{ Date: date, Temperature: temperature, Humid: 40 }],
    status: "off",
    setting: {},
  })});`;

beforeEach(() => {
  props = { ...BASE_PROPS };
  fired = [];
  tempStatus = 200;
  tempResponse = tempFeed(24);
  presenceStatus = 200;
  presenceResponse = arrivePresence(200);
  eventsResponse = {
    events: [
      { timestamp: T_LEAVE, person: "freddiefujiwara", place: "home", action: "leave" },
      { timestamp: T_ARRIVE, person: "freddiefujiwara", place: "home", action: "arrive" },
    ],
  };
  vi.clearAllMocks();
});

describe("parseOffCommands_", () => {
  it("parses a JSON array", () => {
    expect(parseOffCommands_('["/a/off", " /b/off "]')).toEqual(["/a/off", "/b/off"]);
  });
  it("parses one command per line and pipe-separated text", () => {
    expect(parseOffCommands_("/a/off\n/b/off")).toEqual(["/a/off", "/b/off"]);
    expect(parseOffCommands_("/a/off | /b/off")).toEqual(["/a/off", "/b/off"]);
  });
  it("reads empty or non-string input as no commands", () => {
    expect(parseOffCommands_("")).toEqual([]);
    expect(parseOffCommands_(null)).toEqual([]);
    expect(parseOffCommands_(undefined)).toEqual([]);
  });
});

describe("absenceSecondsFromEvents_", () => {
  it("measures the leave -> arrive gap", () => {
    expect(absenceSecondsFromEvents_(eventsResponse.events, T_ARRIVE)).toBe(2400);
  });
  it("is null when the previous event is not a leave", () => {
    const events = [
      { timestamp: T_LEAVE, person: "freddiefujiwara", place: "work", action: "arrive" },
      { timestamp: T_ARRIVE, person: "freddiefujiwara", place: "home", action: "arrive" },
    ];
    expect(absenceSecondsFromEvents_(events, T_ARRIVE)).toBeNull();
  });
  it("is null when there is no previous event or the arrival differs", () => {
    expect(
      absenceSecondsFromEvents_(
        [{ timestamp: T_ARRIVE, person: "freddiefujiwara", place: "home", action: "arrive" }],
        T_ARRIVE
      )
    ).toBeNull();
    expect(absenceSecondsFromEvents_(eventsResponse.events, T_LEAVE)).toBeNull();
    expect(absenceSecondsFromEvents_(null, T_ARRIVE)).toBeNull();
  });
});

describe("absenceSecondsFromState_", () => {
  it("measures the gap from the tracked previous leave", () => {
    const state = { prevEvent: { action: "leave", place: "home", since: T_LEAVE } };
    expect(absenceSecondsFromState_(state, arrivePresence(200))).toBe(2400);
  });
  it("is null without a tracked leave", () => {
    expect(absenceSecondsFromState_({}, arrivePresence(200))).toBeNull();
    expect(
      absenceSecondsFromState_(
        { prevEvent: { action: "arrive", place: "home", since: T_LEAVE } },
        arrivePresence(200)
      )
    ).toBeNull();
  });
});

describe("shouldNotifyArrive_ / shouldHandleLeave_", () => {
  it("applies the delay, absence and freshness windows to arrivals", () => {
    expect(shouldNotifyArrive_(arrivePresence(200), 2400, null)).toBe(true);
    expect(shouldNotifyArrive_(arrivePresence(179), 2400, null)).toBe(false);
    expect(shouldNotifyArrive_(arrivePresence(200), 1799, null)).toBe(false);
    expect(shouldNotifyArrive_(arrivePresence(200), null, null)).toBe(false);
    expect(shouldNotifyArrive_(arrivePresence(3601), 2400, null)).toBe(false);
    expect(shouldNotifyArrive_(arrivePresence(200), 2400, T_ARRIVE)).toBe(false);
    expect(shouldNotifyArrive_({ ...arrivePresence(200), place: "work" }, 2400, null)).toBe(false);
    expect(shouldNotifyArrive_(null, 2400, null)).toBe(false);
  });
  it("applies the delay and freshness windows to leaves", () => {
    expect(shouldHandleLeave_(leavePresence(300), null)).toBe(true);
    expect(shouldHandleLeave_(leavePresence(299), null)).toBe(false);
    expect(shouldHandleLeave_(leavePresence(7201), null)).toBe(false);
    expect(shouldHandleLeave_(leavePresence(400), T_LEAVE)).toBe(false);
    expect(shouldHandleLeave_(arrivePresence(400), null)).toBe(false);
  });
});

describe("processPresence_", () => {
  const now = new Date("2026-11-22T00:43:20.000Z");

  it("does nothing when presence is not configured", () => {
    delete props.PRESENCE_GAS_URL;
    expect(processPresence_(now)).toEqual({ ok: false, reason: "unknown" });
    expect(fired).toEqual([]);
  });

  it("does nothing when the presence read fails", () => {
    presenceStatus = 500;
    expect(processPresence_(now).ok).toBe(false);
    expect(fired).toEqual([]);
  });

  it("speaks once on the tatami speaker after a long absence", () => {
    const result = processPresence_(now);
    expect(result.fired).toEqual(["arrive-speech"]);
    expect(fired).toHaveLength(1);
    expect(fired[0]).toBe(
      "http://home-api.example/google-home-speaker-wrapper/-h/192.0.2.99/-s/" +
        encodeURIComponent(PRESENCE_ARRIVE_MESSAGE_)
    );
    // The raw Japanese text must not appear unencoded in the URL.
    expect(fired[0]).not.toContain("社用携帯");
    // A second check for the same arrival does not speak again.
    expect(processPresence_(now).fired).toEqual([]);
    expect(fired).toHaveLength(1);
  });

  it("stays silent for a short absence and for a too-fresh arrival", () => {
    eventsResponse = {
      events: [
        { timestamp: "2026-11-22T00:30:00.000Z", person: "freddiefujiwara", place: "home", action: "leave" },
        { timestamp: T_ARRIVE, person: "freddiefujiwara", place: "home", action: "arrive" },
      ],
    };
    expect(processPresence_(now).fired).toEqual([]);
    presenceResponse = arrivePresence(60);
    expect(processPresence_(now).fired).toEqual([]);
    expect(fired).toEqual([]);
  });

  it("falls back to the tracked leave when the events list is unavailable", () => {
    eventsResponse = { persons: {} }; // no events array: list not served
    presenceResponse = leavePresence(60);
    processPresence_(now); // observes and tracks the leave
    presenceResponse = arrivePresence(200);
    const result = processPresence_(now);
    expect(result.absenceSec).toBe(2400);
    expect(result.fired).toEqual(["arrive-speech"]);
    expect(fired).toHaveLength(1);
  });

  it("runs the tatami off commands once, five minutes after a leave", () => {
    presenceResponse = leavePresence(400);
    const result = processPresence_(now);
    expect(result.fired).toEqual(["tatami-off"]);
    expect(fired).toEqual([
      "http://home-api.example/dirigera/tatami/off",
      "http://home-api.example/switchbot-ac/off",
      "http://home-api.example/switchbot-circulator/off",
    ]);
    expect(processPresence_(now).fired).toEqual([]);
    expect(fired).toHaveLength(3);
  });

  it("does not run the off commands for a too-fresh or stale leave", () => {
    presenceResponse = leavePresence(120);
    expect(processPresence_(now).fired).toEqual([]);
    presenceResponse = leavePresence(60 * 60 * 5);
    expect(processPresence_(now).fired).toEqual([]);
    expect(fired).toEqual([]);
  });

  it("leaves the event unhandled while the off commands are not configured", () => {
    delete props.PRESENCE_TATAMI_OFF_COMMANDS;
    presenceResponse = leavePresence(400);
    const result = processPresence_(now);
    expect(result.fired).toEqual([]);
    expect(result.skipped).toBe("off commands not configured");
    const state = JSON.parse(props.PRESENCE_STATE);
    expect(state.handledLeaveSince).toBeUndefined();
  });
});

describe("parseTemperatureFeed_", () => {
  it("reads the latest entry of the JSONP feed", () => {
    const parsed = parseTemperatureFeed_(tempFeed(25.3));
    expect(parsed.temperature).toBe(25.3);
    expect(parsed.timeMs).toBe(new Date("11/22/2026 12:40:00 AM").getTime());
  });
  it("reads plain JSON too", () => {
    const parsed = parseTemperatureFeed_(
      JSON.stringify({ conditions: [{ Date: "11/22/2026 12:40:00 AM", Temperature: 19 }] })
    );
    expect(parsed.temperature).toBe(19);
  });
  it("is null for unusable bodies", () => {
    expect(parseTemperatureFeed_("not json")).toBeNull();
    expect(parseTemperatureFeed_(JSON.stringify({ conditions: [] }))).toBeNull();
    expect(
      parseTemperatureFeed_(
        JSON.stringify({ conditions: [{ Date: "11/22/2026 12:40:00 AM", Temperature: "hot" }] })
      )
    ).toBeNull();
    expect(parseTemperatureFeed_(null)).toBeNull();
  });
});

describe("decideTatamiAcMode_", () => {
  it("heats below 20, cools above 27, and rests in between by default", () => {
    expect(decideTatamiAcMode_(19.9)).toBe("heat");
    expect(decideTatamiAcMode_(20)).toBeNull();
    expect(decideTatamiAcMode_(24)).toBeNull();
    expect(decideTatamiAcMode_(27)).toBeNull();
    expect(decideTatamiAcMode_(27.1)).toBe("cool");
    expect(decideTatamiAcMode_(NaN)).toBeNull();
  });
  it("honours threshold properties over the defaults", () => {
    props.PRESENCE_AC_HEAT_BELOW = "22";
    props.PRESENCE_AC_COOL_ABOVE = "26";
    expect(decideTatamiAcMode_(21)).toBe("heat");
    expect(decideTatamiAcMode_(26.5)).toBe("cool");
  });
});

describe("processPresence_ tatami air conditioner (case 2)", () => {
  const now = new Date("2026-11-22T00:43:20.000Z");
  const enableAc = () => {
    props.PRESENCE_TEMP_URL = "https://temp.example/exec?callback=cb";
    props.PRESENCE_TATAMI_AC_HEAT_COMMAND = "switchbot-custom -d TESTDEVICE -c Hot";
    props.PRESENCE_TATAMI_AC_COOL_COMMAND = "switchbot-custom -d TESTDEVICE -c Cool";
    // Keep the speech out of these cases: a short absence stays silent.
    eventsResponse = {
      events: [
        { timestamp: "2026-11-22T00:30:00.000Z", person: "freddiefujiwara", place: "home", action: "leave" },
        { timestamp: T_ARRIVE, person: "freddiefujiwara", place: "home", action: "arrive" },
      ],
    };
  };

  it("starts heating once when the room is cold on arrival", () => {
    enableAc();
    tempResponse = tempFeed(18);
    const result = processPresence_(now);
    expect(result.fired).toEqual(["tatami-ac-heat"]);
    expect(fired).toEqual(["http://home-api.example/switchbot-custom/-d/TESTDEVICE/-c/Hot"]);
    expect(processPresence_(now).fired).toEqual([]);
    expect(fired).toHaveLength(1);
  });

  it("starts cooling when the room is hot on arrival", () => {
    enableAc();
    tempResponse = tempFeed(29.5);
    expect(processPresence_(now).fired).toEqual(["tatami-ac-cool"]);
    expect(fired).toEqual(["http://home-api.example/switchbot-custom/-d/TESTDEVICE/-c/Cool"]);
  });

  it("decides to do nothing at a comfortable temperature", () => {
    enableAc();
    tempResponse = tempFeed(24);
    const result = processPresence_(now);
    expect(result.fired).toEqual([]);
    expect(result.acMode).toBeNull();
    const state = JSON.parse(props.PRESENCE_STATE);
    expect(state.acArriveSince).toBe(T_ARRIVE);
  });

  it("does nothing for a stale temperature and stays undecided", () => {
    enableAc();
    tempResponse = tempFeed(18, "11/21/2026 10:40:00 PM");
    const result = processPresence_(now);
    expect(result.fired).toEqual([]);
    expect(result.ac).toBe("temperature unknown");
    const state = JSON.parse(props.PRESENCE_STATE);
    expect(state.acArriveSince).toBeUndefined();
  });

  it("does nothing when the temperature feed is not configured", () => {
    tempResponse = tempFeed(18);
    const result = processPresence_(now);
    // Only the welcome-home speech fires (the default events show a
    // long absence); no air-conditioner command runs.
    expect(result.fired).toEqual(["arrive-speech"]);
    expect(result.ac).toBe("temperature unknown");
    expect(fired).toHaveLength(1);
  });
});
