// Dates on Excel's serial numbers: days since 1899-12-30 (so 1 is
// 1899-12-31, which Excel calls 1900-01-01; the two agree from 1900-03-01
// on, Excel counting a 29 February 1900 that never was), a time the
// fraction of a day. Pure date arithmetic, for the library's date functions
// and numfmt.ts. Or, in the 1904 date system (a workbook's setting, as
// old Mac Excel's), days since 1904-01-01, 0 that day: setDateSystem
// moves the epoch every function here counts from.
let EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;

export type DateSystem = 1900 | 1904;
export function setDateSystem(system: DateSystem): void {
  EPOCH = system === 1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
}

// Excel's calendar has a 29 February 1900 that never was (day 60) and a day 0, 0 January 1900: in the
// 1900 date system a serial under 61 is a day later in the calendar than the real one it
// numbers (serial 1 is 1 January 1900, the real 31 December 1899), so those are worked out by hand.
const phantom = (): boolean => EPOCH === Date.UTC(1899, 11, 30);
export function serial(y: number, m: number, d: number): number {
  // January and February 1900 count from 0 January: 31 days then the days of February (which has 29).
  if (phantom() && y === 1900 && (m === 1 || m === 2)) return (m === 1 ? 0 : 31) + d;
  return (Date.UTC(y, m - 1, d) - EPOCH) / DAY;
}

// A serial's year, month (1-12), day and weekday (0 Sunday).
export function parts(s: number): { y: number; m: number; d: number; weekday: number } {
  const f = Math.floor(s);
  if (phantom() && f >= 0 && f < 61) {
    const real = new Date(EPOCH + f * DAY);
    return f === 0 ? { y: 1900, m: 1, d: 0, weekday: real.getUTCDay() } : f < 32 ? { y: 1900, m: 1, d: f, weekday: real.getUTCDay() } : { y: 1900, m: 2, d: f - 31, weekday: real.getUTCDay() };
  }
  const t = new Date(EPOCH + f * DAY);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate(), weekday: t.getUTCDay() };
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const LONG_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
// A month's name, in full or its first three letters (and "Sept"), 1 to 12; 0 if not one.
const monthNamed = (name: string): number => {
  const n = name.toLowerCase(), i = LONG_MONTHS.indexOf(n);
  return i >= 0 ? i + 1 : n === "sept" ? 9 : n.length === 3 ? MONTHS.indexOf(n) + 1 : 0;
};
// A year written with two digits is 2000 to 2029, or 1930 to 1999.
const yearOf = (digits: string): number => (digits.length === 2 ? (Number(digits) < 30 ? 2000 : 1900) + Number(digits) : Number(digits));

// Text as a date, as en-US Excel reads it: 2024-01-15, month first with / or -
// (1/15/2024, 1-15-24), a month named (15-Jan-2024, 15 January 2024, Jan 15, 2024),
// with spaces allowed round the separators ("2026 - 01 - 01"); null if not one.
export function parseDate(text: string): number | null {
  const t = text.trim().replace(/\s*([/-])\s*/g, "$1");
  // A month of 1 to 12 and a day that month has: 30/2/2010 is no date.
  // Before 1900 is no date (and Excel's 29 February 1900, which never was, is: serial() and parts() count it).
  const real = (y: number, mo: number, d: number): number | null =>
    y < 1900 ? null : mo >= 1 && mo <= 12 && d >= 1 && d <= parts(serial(y, mo + 1, 0)).d ? serial(y, mo, d) : null;
  let m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(t);
  if (m) return real(Number(m[1]), Number(m[2]), Number(m[3]));
  m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$/.exec(t);
  if (m) return real(yearOf(m[3]), Number(m[1]), Number(m[2]));
  m = /^(\d{1,2})[- ]([A-Za-z]+)[- ,]+(\d{2}|\d{4})$/.exec(t);
  if (m && monthNamed(m[2])) return real(yearOf(m[3]), monthNamed(m[2]), Number(m[1]));
  m = /^([A-Za-z]+)[- ](\d{1,2}),?[- ](\d{2}|\d{4})$/.exec(t);
  if (m && monthNamed(m[1])) return real(yearOf(m[3]), monthNamed(m[1]), Number(m[2]));
  return null;
}

// A time written in text, in seconds from midnight (24:00 and the like come to 86,400 or
// more): "6:45 PM", "22:48:00", "4 pm", "4 : 35" (spaces round the colons are
// allowed), "12:00 am". Minutes and seconds are under 60 (a seconds of 60 is
// accepted where it makes exactly 24:00:00, as 23:60 does); an hour is 0 to 23 (24
// with any minutes and seconds), or 1 to 12 with am or pm; `hours` says how far
// an hour may go (a date's time, "2026-01-01 25:00:00", carries whole days).
function secondsOf(text: string, hours: number): number | null {
  const t = text.trim().replace(/\s*:\s*/g, ":");
  let m = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2})(?:\.\d+)?)?\s*(?:([AaPp])[Mm]?)?$/.exec(t);
  if (!m) {
    m = /^(\d{1,2})()()\s*([AaPp])[Mm]$/.exec(t);
    if (!m) return null;
  }
  let h = Number(m[1]);
  const min = Number(m[2] || 0), sec = Number(m[3] || 0), half = m[4]?.toUpperCase();
  if (half) {
    if (h < 1 || h > 12 || min > 59 || sec > 59) return null;
    if (half === "P" && h < 12) h += 12;
    if (half === "A" && h === 12) h = 0;
  } else {
    const total = h * 3600 + min * 60 + sec;
    const rolls = (h === 24 && min <= 59 && sec <= 59) || (h === 23 && min === 60 && sec <= 59) || (h <= 23 && min <= 59 && sec === 60 && total === 86400);
    if (!rolls && (h > hours || min > 59 || sec > 59)) return null;
  }
  return h * 3600 + min * 60 + sec;
}

// A date's text and the time after it ("2026-01-01 24:00:00", "7/7/2026 4:35 PM",
// "2026-01-01T09:30"), or only one of them; null when it is neither. `hours`: how far
// the time's hour may go (see secondsOf).
function stamp(text: string, hours: number): { day: number | null; seconds: number | null } | null {
  const t = text.trim();
  const clock = /(?:^|[ T])(\d{1,4}\s*:\s*\d{1,2}(?:\s*:\s*\d{1,2}(?:\.\d+)?)?(?:\s*[AaPp][Mm]?)?|\d{1,2}\s*[AaPp][Mm])\s*$/.exec(t);
  if (!clock) { const d = parseDate(t); return d === null ? null : { day: d, seconds: null }; }
  const seconds = secondsOf(clock[1], hours), rest = t.slice(0, clock.index).trim();
  if (seconds === null) return null;
  if (rest === "") return { day: null, seconds };
  const day = parseDate(rest);
  return day === null ? null : { day, seconds };
}

// Text as a time of day, the fraction of a day: "6:45 PM", "22:48:00",
// "12:00 am", a date in front ignored; null if not one (24:00 is none).
export function parseTime(text: string): number | null {
  const s = stamp(text, 23);
  return s && s.seconds !== null && s.seconds < 86400 ? s.seconds / 86400 : null;
}

// TIMEVALUE's text: a time as parseTime reads it, where 24:00:00 and the like (23:60,
// 23:59:60) come round to 0.
export function timeValue(text: string): number | null {
  const s = stamp(text, 23);
  return s && s.seconds !== null ? (s.seconds % 86400) / 86400 : null;
}

// Text as a date with its time: the serial and the day's fraction (a time of 24 hours
// or more carries into the date); a date alone; null if it is not one (a time alone
// is parseTime's).
export function parseDateTime(text: string): number | null {
  const s = stamp(text, 9999);
  return s && s.day !== null ? s.day + (s.seconds ?? 0) / 86400 : null;
}

// Months on, the day kept (or the month's last, when it has fewer).
export function addMonths(s: number, months: number): number {
  const { y, m, d } = parts(s);
  const last = parts(serial(y, m + months + 1, 0)).d;
  return serial(y, m + months, Math.min(d, last));
}

export const endOfMonth = (s: number, months: number): number => {
  const { y, m } = parts(s);
  return serial(y, m + months + 1, 0);
};

// Which days are the weekend, Monday first (NETWORKDAYS.INTL's string:
// "0000011" is Saturday and Sunday).
export type Weekend = readonly boolean[];
export const SATURDAY_SUNDAY: Weekend = [false, false, false, false, false, true, true];

// A weekend as the .INTL functions take it: a number 1 to 7 for two days
// (1 Saturday and Sunday, 2 Sunday and Monday ... 7 Friday and Saturday),
// 11 to 17 for one (11 Sunday, 12 Monday ... 17 Saturday), or seven 1s
// and 0s from Monday; null when it is none of those (or every day).
export function weekendOf(spec: number | string): Weekend | null {
  if (typeof spec === "string") {
    if (!/^[01]{7}$/.test(spec)) return null;
    return [...spec].map((c) => c === "1");
  }
  const days: boolean[] = [false, false, false, false, false, false, false];
  if (spec >= 1 && spec <= 7) {
    // Pair n is day n + 4 and the next, Monday 0: 1 is Saturday (5) and Sunday (6).
    days[(spec + 4) % 7] = true;
    days[(spec + 5) % 7] = true;
  } else if (spec >= 11 && spec <= 17) {
    days[(spec - 11 + 6) % 7] = true;
  } else return null;
  return days;
}

const weekend = (s: number, which: Weekend): boolean => which[(parts(s).weekday + 6) % 7];

// Weekdays from start to end, both counted, holidays not (either order: negative).
export function workdaysBetween(start: number, end: number, holidays: ReadonlySet<number>, which: Weekend = SATURDAY_SUNDAY): number {
  const [from, to] = [Math.floor(start), Math.floor(end)];
  const [a, b] = from <= to ? [from, to] : [to, from];
  let n = 0;
  for (let s = a; s <= b; s++) if (!weekend(s, which) && !holidays.has(s)) n++;
  return from <= to ? n : -n;
}

// The workday `days` workdays from start (back, when negative). Whole weeks are jumped, so a
// count of a billion is no longer than a count of ten; a holiday on the way (a workday that is
// one) is one more workday to go. A weekend of all seven days has no workday: null.
export function workdayFrom(start: number, days: number, holidays: ReadonlySet<number>, which: Weekend = SATURDAY_SUNDAY): number | null {
  const per = which.filter((w) => !w).length;
  if (per === 0) return null;
  const step = Math.sign(days);
  let s = Math.floor(start), left = Math.abs(days);
  const sorted = [...holidays].filter((h) => !weekend(h, which)).sort((x, y) => x - y);
  while (left > 0) {
    const from = s;
    const weeks = Math.floor((left - 1) / per);
    s += step * 7 * weeks;
    left -= weeks * per;
    while (left > 0) {
      s += step;
      if (!weekend(s, which)) left--;
    }
    // The holidays this stretch passed over (from, s] (or [s, from)) are workdays not walked.
    left = sorted.filter((h) => (step > 0 ? h > from && h <= s : h < from && h >= s)).length;
  }
  return s;
}

// The week of the year, counted from 1 at the week January 1 is in, weeks
// starting on `start` (0 Sunday, as parts' weekday).
export function weekOfYear(s: number, start: number): number {
  const { y } = parts(s), jan1 = serial(y, 1, 1);
  const offset = (parts(jan1).weekday - start + 7) % 7;
  return Math.floor(Math.floor(s) - jan1 + offset) / 7 + 1 | 0;
}

// ISO 8601's week: weeks start on Monday, week 1 holds the year's first
// Thursday.
export function isoWeek(s: number): number {
  const day = Math.floor(s), monday = (parts(day).weekday + 6) % 7;
  const thursday = day - monday + 3;
  return Math.floor((thursday - serial(parts(thursday).y, 1, 1)) / 7) + 1;
}

const lastOfFebruary = (p: { y: number; m: number; d: number }): boolean => p.m === 2 && p.d === parts(serial(p.y, 3, 0)).d;

// Days between on a 360-day year of 30-day months: the US (NASD) way, or
// the European. US: a start on the 31st, or the last of February, is the
// 30th (and then an end on the last of February too); an end on the
// 31st is the 30th when the start is the 30th or later, else the 1st of
// the next month. European: a 31st is the 30th, either end.
export function days360(start: number, end: number, european: boolean): number {
  // A serial within 1E-5 of the next day is that day (45658.9999999 is 45659, as Excel reads it here).
  const near = (s: number): number => Math.floor(Number(s.toFixed(5)));
  const a = parts(near(start)), b = parts(near(end));
  let sd = a.d, ed = b.d, em = b.m, ey = b.y;
  if (european) {
    if (sd === 31) sd = 30;
    if (ed === 31) ed = 30;
  } else {
    if (sd === 31 || lastOfFebruary(a)) sd = 30;
    if (ed === 31) {
      if (sd >= 30) ed = 30;
      else { ed = 1; em++; if (em > 12) { em = 1; ey++; } }
    }
  }
  return (ey - a.y) * 360 + (em - a.m) * 30 + (ed - sd);
}

const isLeap = (y: number): boolean => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

// YEARFRAC's bases: 0 US 30/360 (its own February rule: an end on the
// 31st is the 30th when the start is the 30th or 31st), 1 actual over
// the actual year (the average length of the years spanned), 2 actual
// over 360, 3 actual over 365, 4 European 30/360.
export function yearFraction(start: number, end: number, basis: number): number | null {
  const [a, b] = start <= end ? [start, end] : [end, start];
  const days = Math.floor(b) - Math.floor(a);
  switch (basis) {
    case 0: {
      const p = parts(a), q = parts(b);
      let sd = p.d, ed = q.d;
      if (lastOfFebruary(p) && lastOfFebruary(q)) ed = 30;
      if (sd === 31 || lastOfFebruary(p)) sd = 30;
      if (ed === 31 && sd >= 30) ed = 30;
      return ((q.y - p.y) * 360 + (q.m - p.m) * 30 + (ed - sd)) / 360;
    }
    case 1: {
      const p = parts(a), q = parts(b);
      if (p.y === q.y) return days / (isLeap(p.y) ? 366 : 365);
      // Within a year (the end no later in the next year than the start's
      // month and day): a year of 366 days if a 29 February falls between
      // them, else 365, as Excel counts; longer spans average the years.
      if (q.y === p.y + 1 && (q.m < p.m || (q.m === p.m && q.d <= p.d))) {
        const leapDay = (isLeap(p.y) && (p.m < 3)) || (isLeap(q.y) && (q.m > 2 || (q.m === 2 && q.d === 29)));
        return days / (leapDay ? 366 : 365);
      }
      const span = serial(q.y + 1, 1, 1) - serial(p.y, 1, 1);
      return days / (span / (q.y - p.y + 1));
    }
    case 2: return days / 360;
    case 3: return days / 365;
    case 4: return days360(a, b, true) / 360;
    default: return null;
  }
}

// DATEDIF's units: whole years, months or days between, or the months
// past the years (YM), the days past the months (MD), the days past the
// years (YD).
export function dateDif(start: number, end: number, unit: string): number | null {
  const a = parts(start), b = parts(end);
  let months = (b.y - a.y) * 12 + (b.m - a.m);
  if (b.d < a.d) months--;
  switch (unit.toUpperCase()) {
    case "Y": return Math.floor(months / 12);
    case "M": return months;
    case "D": return Math.floor(end) - Math.floor(start);
    case "YM": return months % 12;
    case "MD": return b.d >= a.d ? b.d - a.d : Math.floor(end) - serial(b.y, b.m - 1, a.d);
    case "YD": {
      const from = serial(b.m > a.m || (b.m === a.m && b.d >= a.d) ? b.y : b.y - 1, a.m, a.d);
      return Math.floor(end) - from;
    }
    default: return null;
  }
}

// An injected clock (workbook.clock): a fixed serial instead of the host's
// clock, so a test, or a demonstration, can pin TODAY() and NOW(); null is the
// host's. Not saved with a file.
let fixedNow: number | null = null;
export const setClock = (serial: number | null): void => { fixedNow = serial; };
export const clockIsFixed = (): boolean => fixedNow !== null;

// Now as a serial, in the page's local time (the coarse clock in the browser),
// or the injected clock's.
export const now = (): number => {
  if (fixedNow !== null) return fixedNow;
  const t = new Date();
  return (t.getTime() - t.getTimezoneOffset() * 60_000 - EPOCH) / DAY;
};

// A clock's text as a serial: "2026-10-21", "2026-10-21 09:30" (or with
// seconds, or a T between), or a serial number; null if it is none of these.
export function parseInstant(text: string): number | null {
  const t = text.trim();
  if (/^\d+(\.\d+)?$/.test(t)) return Number(t);
  const m = /^(\d{4}-\d{1,2}-\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (!m) return null;
  const day = parseDate(m[1]);
  if (day === null) return null;
  const h = Number(m[2] ?? 0), min = Number(m[3] ?? 0), sec = Number(m[4] ?? 0);
  if (h > 23 || min > 59 || sec > 59) return null;
  return day + (h * 3600 + min * 60 + sec) / 86400;
}

// A serial as "2026-10-21 09:30:00".
export function instantText(s: number): string {
  const { y, m, d } = parts(s), secs = Math.round((s - Math.floor(s)) * 86400), p = (n: number, w = 2) => String(n).padStart(w, "0");
  return `${p(y, 4)}-${p(m)}-${p(d)} ${p(Math.floor(secs / 3600))}:${p(Math.floor(secs / 60) % 60)}:${p(secs % 60)}`;
}
