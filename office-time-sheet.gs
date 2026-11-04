const SHEET_NAME = "Office Time Data";
const HEADERS = [
  "Date",
  "Sign in",
  "Sign out",
  "Break",
  "Productive Work",
  "Total Work",
  "Vs required",
  "Tasks Details",
  "Break Details",
  "Summary"
];

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Missing request data.");
  }

  const payload = JSON.parse(e.postData.contents);
  if (!payload || !Array.isArray(payload.rows) || payload.rows.length > 5000) {
    throw new Error("Invalid or oversized rows data.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("This script must be bound to the target spreadsheet.");
    }

    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
    sheet.clear();

    const values = payload.rows.map(row => [
      sheetDate(row.date, spreadsheet.getSpreadsheetTimeZone()),
      timeValue(row.in),
      timeValue(row.out),
      durationValue(row.breakMinutes),
      durationValue((row.tasks || []).reduce((sum, task) => sum + (validMinutes(task.durationMinutes) || 0), 0)),
      totalWorkValue(row),
      signedDifference(row.differenceMinutes),
      taskDetails(row.tasks || []),
      breakDetails(row.breaks || []),
      summaryDetails(row)
    ]);

    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    if (values.length) {
      sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
      sheet.getRange(2, 1, values.length, 1).setNumberFormat("mmm dd, yy : ddd");
      sheet.getRange(2, 2, values.length, 2).setNumberFormat("h:mm AM/PM");
      sheet.getRange(2, 4, values.length, 3).setNumberFormat("[h]:mm");
      sheet.getRange(2, 8, values.length, 3).setWrap(true).setVerticalAlignment("middle");
      sheet.autoResizeRows(2, values.length);
    }

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setBackground("#d9e8fb");
    sheet.setRowHeight(1, 32);
    [145, 90, 90, 90, 125, 100, 105, 400, 260, 380].forEach((width, index) =>
      sheet.setColumnWidth(index + 1, width)
    );
    if (values.length) {
      sheet.getRange(2, 1, values.length, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: values.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function sheetDate(value, timezone) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return "";
  return Utilities.parseDate(value, timezone, "yyyy-MM-dd");
}

function timeValue(value) {
  const match = String(value || "").match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) return "";
  return (Number(match[1]) * 60 + Number(match[2])) / 1440;
}

function durationValue(value) {
  if (value == null || !Number.isFinite(Number(value))) return "";
  return Math.max(0, Number(value)) / 1440;
}

function totalWorkValue(row) {
  const suppliedMinutes = validMinutes(row.totalWorkMinutes);
  if (suppliedMinutes !== null) return durationValue(suppliedMinutes);

  const signIn = timeMinutes(row.in);
  const signOut = timeMinutes(row.out);
  if (signIn === null || signOut === null) return "";
  const elapsedMinutes = (signOut - signIn + 1440) % 1440;
  return durationValue(elapsedMinutes);
}

function timeMinutes(value) {
  const match = String(value || "").match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : null;
}

function signedDifference(value) {
  const minutes = validMinutes(value);
  if (minutes === null) return "";
  const sign = minutes > 0 ? "+" : minutes < 0 ? "-" : "";
  return `${sign}${formatMinutes(Math.abs(minutes))}`;
}

function taskDetails(tasks) {
  const groups = [];
  tasks.forEach(task => {
    const name = String(task.name || "").trim();
    const key = name.toLocaleLowerCase();
    let group = name ? groups.find(item => item.key === key) : null;
    if (!group) {
      group = { key: name ? key : `untitled-${groups.length}`, name: name || "Untitled task", sessions: [] };
      groups.push(group);
    }
    group.sessions.push(task);
    group.status = task.status || "Not started";
  });

  return safeCell(groups.map(group => {
    const recorded = group.sessions.map(task => validMinutes(task.durationMinutes)).filter(value => value !== null);
    const total = recorded.reduce((sum, duration) => sum + duration, 0);
    const lines = [`${group.name}${recorded.length ? ` — ${formatMinutes(total)}` : ""}`];
    group.sessions.forEach(task => {
      lines.push(
        `Time: ${formatTime(task.start)} - ${formatTime(task.end)}`,
        `Duration: ${formatMinutes(task.durationMinutes)}`
      );
    });
    lines.push(`Status: ${group.status}`);
    return lines.join("\n");
  }).join("\n\n"));
}

function breakDetails(breaks) {
  return safeCell(breaks.map(item => [
    `Time: ${formatTime(item.start)} - ${formatTime(item.end)}`,
    `Duration: ${formatMinutes(item.durationMinutes)}`
  ].join("\n")).join("\n\n"));
}

function summaryDetails(row) {
  const dayType = row.dayType === "offday" ? "Off day" : row.dayType === "holiday" ? "Holiday" : "Workday";
  if (dayType !== "Workday") {
    const lines = ["Summary", `Day Type: ${dayType}`];
    if (row.note) lines.push(`Day Note: ${row.note}`);
    return safeCell(lines.join("\n"));
  }

  const taskTotal = (row.tasks || []).reduce((sum, task) =>
    sum + (validMinutes(task.durationMinutes) || 0), 0);
  const lines = [
    "Summary",
    `Arrival Time: ${formatTime(row.in)}`,
    `Total Break Time: ${formatMinutes(row.breakMinutes)}`,
    `Leaving Time: ${formatTime(row.out)}`,
    `Total Productive Hours: ${formatMinutes(taskTotal)}`
  ];
  if (row.note) lines.push(`Day Note: ${row.note}`);
  return safeCell(lines.join("\n"));
}

function formatTime(value) {
  const match = String(value || "").match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) return "Not set";
  const hour = Number(match[1]);
  return `${hour % 12 || 12}:${match[2]} ${hour < 12 ? "AM" : "PM"}`;
}

function validMinutes(value) {
  const number = Number(value);
  return value == null || value === "" || !Number.isFinite(number) ? null : number;
}

function formatMinutes(value) {
  const minutes = validMinutes(value);
  if (minutes === null) return "Not recorded";
  return `${Math.floor(minutes / 60)}h ${String(minutes % 60).padStart(2, "0")}m`;
}

function safeCell(value) {
  const text = value == null ? "" : String(value);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}
