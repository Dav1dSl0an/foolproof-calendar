/**
 * add-to-calendar.js
 * Modern ESM calendar widget using a styled dropdown button with service icons.
 */

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

import { addDays, addHours } from "https://esm.sh/date-fns@3.6.0";
import { toZonedTime } from "https://esm.sh/date-fns-tz@3.0.0";

function validateEvent(event) {
  if (!event || typeof event !== "object") {
    throw new TypeError("createCalendar: `data` must be an object.");
  }
  if (!(event.start instanceof Date) || isNaN(event.start.getTime())) {
    throw new TypeError("createCalendar: `data.start` must be a valid Date.");
  }
  if (event.end !== undefined) {
    if (!(event.end instanceof Date) || isNaN(event.end.getTime())) {
      throw new TypeError("createCalendar: `data.end` must be a valid Date.");
    }
    if (event.end <= event.start) {
      throw new RangeError(
        "createCalendar: `data.end` must be after `data.start`.",
      );
    }
  } else if (event.duration !== undefined) {
    if (typeof event.duration !== "number" || event.duration <= 0) {
      throw new TypeError(
        "createCalendar: `data.duration` must be a positive number (minutes).",
      );
    }
  } else {
    throw new Error(
      "createCalendar: either `data.end` or `data.duration` is required.",
    );
  }
}

// ---------------------------------------------------------------------------
// Date helpers
// ---------------------------------------------------------------------------

const MS_PER_MINUTE = 60 * 1000;

function toUtcString(date) {
  return date.toISOString().replace(/[-:]|\.\d+/g, "");
}

function resolveEndDate(event) {
  return (
    event.end ??
    new Date(event.start.getTime() + event.duration * MS_PER_MINUTE)
  );
}

function durationMinutes(start, end) {
  return Math.round((end.getTime() - start.getTime()) / MS_PER_MINUTE);
}

function toYahooDuration(minutes) {
  const hh = String(Math.floor(minutes / 60)).padStart(2, "0");
  const mm = String(minutes % 60).padStart(2, "0");
  return hh + mm;
}

// ---------------------------------------------------------------------------
// URL builders
// ---------------------------------------------------------------------------

/**
 * Formats a Date as a compact LOCAL datetime string (no trailing Z):
 * 20260510T140000 — used with Google's `ctz` param so Google interprets
 * the time in the supplied timezone rather than as UTC/GMT.
 * @param {Date} date
 * @returns {string}
 */
function toLocalString(date) {
  // Build YYYYMMDDTHHmmss from the date's UTC fields.
  // Callers are responsible for passing a Date whose UTC values already
  // reflect the desired wall-clock time in the target timezone, OR for
  // relying on the `ctz` param to let Google do the conversion.
  //
  // The simplest correct approach: strip the Z from toISOString() so Google
  // treats the value as a "floating" local time in the `ctz` timezone.
  return date.toISOString().replace(/[-:]|\.\d+|Z$/g, "");
}

function googleUrl(event) {
  const end = resolveEndDate(event);

  // When a timezone is supplied, pass floating local times + ctz so Google
  // displays the event in the correct timezone instead of defaulting to GMT.
  // When no timezone is supplied, fall back to UTC strings (trailing Z).
  let startStr, endStr;
  if (event.timezone) {
    startStr = toLocalString(event.start);
    endStr = toLocalString(end);
  } else {
    startStr = toUtcString(event.start);
    endStr = toUtcString(end);
  }

  const url = new URL("https://www.google.com/calendar/render");
  url.searchParams.set("action", "TEMPLATE");
  url.searchParams.set("text", event.title ?? "");
  url.searchParams.set("dates", `${startStr}/${endStr}`);
  url.searchParams.set("details", event.description ?? "");
  url.searchParams.set("location", event.address ?? "");
  if (event.timezone) {
    url.searchParams.set("ctz", event.timezone);
  }
  return url.href;
}

function yahooUrl(event) {
  const end = resolveEndDate(event);
  const duration = toYahooDuration(durationMinutes(event.start, end));
  const url = new URL("https://calendar.yahoo.com/");
  url.searchParams.set("v", "60");
  url.searchParams.set("view", "d");
  url.searchParams.set("type", "20");
  url.searchParams.set("title", event.title ?? "");
  url.searchParams.set("st", toUtcString(event.start));
  url.searchParams.set("dur", duration);
  url.searchParams.set("desc", event.description ?? "");
  url.searchParams.set("in_loc", event.address ?? "");
  return url.href;
}

function icsDataUrl(event) {
  const end = resolveEndDate(event);
  const dtStart = toUtcString(event.start);
  const dtEnd = toUtcString(end);
  const tzid = event.timezone;
  const startLine = tzid
    ? `DTSTART;TZID=${tzid}:${dtStart}`
    : `DTSTART:${dtStart}`;
  const endLine = tzid ? `DTEND;TZID=${tzid}:${dtEnd}` : `DTEND:${dtEnd}`;

  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//add-to-calendar//EN",
    "BEGIN:VEVENT",
    `URL:${document.URL}`,
    startLine,
    endLine,
    `SUMMARY:${event.title ?? ""}`,
    `DESCRIPTION:${event.description ?? ""}`,
    `LOCATION:${event.address ?? ""}`,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");

  return `data:text/calendar;charset=utf8,${encodeURIComponent(ics)}`;
}

// ---------------------------------------------------------------------------

function outlookComUrl(event) {
  const end = resolveEndDate(event);
  const fmt = (d) => d.toISOString().replace(/[.]\d+Z$/, "");
  const url = new URL("https://outlook.live.com/calendar/0/action/compose");
  url.searchParams.set("rru", "addevent");
  url.searchParams.set("subject", event.title ?? "");
  url.searchParams.set("startdt", fmt(event.start));
  url.searchParams.set("enddt", fmt(end));
  url.searchParams.set("body", event.description ?? "");
  url.searchParams.set("location", event.address ?? "");
  if (event.timezone) {
    url.searchParams.set("tz", event.timezone);
  }
  return url.href;
}

// SVG icons for each service
// ---------------------------------------------------------------------------

export const ICONS = {
  google: `<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
    <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
    <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
    <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
    <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.31-8.16 2.31-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
  </svg>`,

  yahoo: `<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
    <rect width="48" height="48" rx="8" fill="#6001D2"/>
    <text x="50%" y="54%" dominant-baseline="middle" text-anchor="middle"
      font-family="Georgia, serif" font-size="22" font-weight="bold" fill="#fff">Y!</text>
  </svg>`,

  ical: `<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
    <rect width="48" height="48" rx="8" fill="#fff" stroke="#ddd" stroke-width="1.5"/>
    <rect y="0" width="48" height="15" rx="4" fill="#F05138"/>
    <rect y="11" width="48" height="4" fill="#F05138"/>
    <text x="50%" y="70%" dominant-baseline="middle" text-anchor="middle"
      font-family="Georgia, serif" font-size="15" font-weight="bold" fill="#1a1a1a">iCal</text>
  </svg>`,

  outlookcom: `<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
    <rect width="48" height="48" rx="8" fill="#0078D4"/>
    <rect x="6" y="14" width="36" height="22" rx="3" fill="#fff"/>
    <polyline points="6,14 24,28 42,14" fill="none" stroke="#0078D4" stroke-width="2.5" stroke-linejoin="round"/>
  </svg>`,

  outlook: `<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
    <rect width="48" height="48" rx="8" fill="#0078D4"/>
    <rect x="5" y="10" width="24" height="28" rx="3" fill="#fff"/>
    <rect x="19" y="10" width="24" height="28" rx="3" fill="#50A0E0"/>
    <ellipse cx="17" cy="24" rx="7" ry="8.5" fill="#0078D4"/>
    <ellipse cx="17" cy="24" rx="5" ry="6.5" fill="#fff"/>
  </svg>`,
};

// ---------------------------------------------------------------------------
// DOM builder
// ---------------------------------------------------------------------------

function buildWidget(event, calendarId, options) {
  // Derive a safe filename from the event title e.g. "team-meeting.ics"
  const safeTitle =
    (event.title ?? "event")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "event";
  const icsFilename = `${safeTitle}.ics`;

  const services = [
    {
      key: "google",
      label: "Google Calendar",
      href: googleUrl(event),
      download: null,
    },
    {
      key: "yahoo",
      label: "Yahoo! Calendar",
      href: yahooUrl(event),
      download: null,
    },
    {
      key: "ical",
      label: "Apple iCal",
      href: icsDataUrl(event),
      download: icsFilename,
    },
    {
      key: "outlookcom",
      label: "Outlook.com",
      href: outlookComUrl(event),
      download: null,
    },
    {
      key: "outlook",
      label: "Outlook (desktop)",
      href: icsDataUrl(event),
      download: icsFilename,
    },
  ];

  return services;
}

export function createCalendar({ data: event, options = {} } = {}) {
  validateEvent(event);
  const calendarId =
    options.id ?? `foolproof-${Math.floor(Math.random() * 1_000_000)}`;

  return buildWidget(event, calendarId, options);
}

export function getNextMeeting(
  TIMEZONE,
  MEETING_DAY,
  MEETING_HOUR,
  DURATION_HOURS,
) {
  const now = toZonedTime(new Date(), TIMEZONE);

  let start = new Date(now);

  let daysUntil = MEETING_DAY - start.getDay();
  if (daysUntil < 0) daysUntil += 7;

  // If today is Thursday and we're at/after noon → next week
  if (
    daysUntil === 0 &&
    (start.getHours() > MEETING_HOUR ||
      (start.getHours() === MEETING_HOUR && start.getMinutes() > 0))
  ) {
    daysUntil = 7;
  }

  start = addDays(start, daysUntil);
  start.setHours(MEETING_HOUR, 0, 0, 0);

  // Calculate end time
  const end = addHours(start, DURATION_HOURS);

  return {
    start,
    end,
  };
}

export function convertString(dateStr) {
  const date = new Date(dateStr);

  // Shift by the offset to treat local time as UTC
  const offsetMs = date.getTimezoneOffset() * 60 * 1000; // browser's local offset
  const localMs = date.getTime() + date.getTimezoneOffset() * 60 * 1000;

  // Extract components from the original string's local time
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");

  const result = `${year}-${month}-${day}T${hours}:${minutes}:${seconds}Z`;
  return result;
}
