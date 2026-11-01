# alarm-gas

The back side of the Alarm app. It keeps alarms in a Google Sheets
sheet and fires them from a clock trigger.

## How it works

- One sheet is the alarm list. There is no header row.
  - Column A: the date and time to fire.
  - Columns B to K: up to 10 commands to run at that time.
  - A command is a path on the home command API (for example
    `/hue/lights/off`) or a full URL.
- `myFunction` runs every minute from a time-driven trigger.
  - If the time is within 1 minute of now, it sends the commands
    one by one, then deletes the row.
  - If the time has already passed, it deletes the row without
    sending anything. The sheet cleans itself up.

## Web API (doGet)

Call the web app URL with these parameters.

| Name | Meaning |
| --- | --- |
| `time` | `HH:MM` or `HH:MM:SS` means today. `YYYYMMDD-HH:MM` means that date and time. |
| `cmd` | One command. Repeat this parameter for more commands (max 10). |
| `text` | Old style speech text. If `cmd` is empty and `text` is set, one default speech command is saved instead. |
| `callback` | Old style JSONP callback name. |

Rules:

- If a row with the same date and time exists, the new request
  overwrites that row.
- A past date or time is an error.
- More than 10 commands is an error.
- The answer is JSON: `{"ok":true,...}` or
  `{"ok":false,"error":"..."}`. Apps Script web apps cannot set an
  HTTP status code, so errors are only in the JSON body.

Example:

```
.../exec?time=20270109-07:30&cmd=/hue/lights/on&cmd=/switchbot-command/-d/XXX/-c/turnOn
```

## Setup

1. Open the Apps Script project bound to the "Alarm" sheet.
2. In the project settings, set these script properties. Their
   values are private, so they are not in this repository.
   - `API_BASE_URL`: base URL of the home command API.
   - `DEFAULT_SPEECH_PATH`: path prefix of the speech command used
     for old style text alarms, ending with `/-s/`.
3. Set `.clasp.json` to that project's script ID.
4. Run `npm install`, `npm run build`, and `npm run deploy`
   (clasp push) to send `dist/` to the project.
5. Deploy as a web app (execute as you, access: anyone, anonymous).
6. Add a time-driven trigger: `myFunction`, every minute.

GitHub Actions (Node CI) can build and deploy on push when the
repository secrets `CLASPRC_JSON` and `DEPLOYMENT_ID` are set.

## Test

```
npm test
```

Unit tests mock the Apps Script services (SpreadsheetApp,
UrlFetchApp, ContentService, LockService) and check the time
parsing, the API rules (past time, overwrite, the 10 command
limit, the legacy text fallback, JSONP), and the trigger logic
(fire, clean up, keep, legacy rows, delete order).
