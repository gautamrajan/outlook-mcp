/**
 * Agent-friendly date formatting utilities.
 * Converts raw ISO timestamps and Windows timezone names into
 * human-readable strings with day names and relative markers.
 */
const config = require('../config');

// Comprehensive Windows timezone name → IANA timezone identifier mapping.
// Source: Unicode CLDR windowsZones.xml (territory "001" defaults).
const WINDOWS_TO_IANA = {
  'Dateline Standard Time': 'Etc/GMT+12',
  'UTC-11': 'Pacific/Pago_Pago',
  'Aleutian Standard Time': 'America/Adak',
  'Hawaiian Standard Time': 'Pacific/Honolulu',
  'Marquesas Standard Time': 'Pacific/Marquesas',
  'Alaskan Standard Time': 'America/Anchorage',
  'UTC-09': 'Pacific/Gambier',
  'Pacific Standard Time (Mexico)': 'America/Tijuana',
  'UTC-08': 'Etc/GMT+8',
  'Pacific Standard Time': 'America/Los_Angeles',
  'US Mountain Standard Time': 'America/Phoenix',
  'Mountain Standard Time (Mexico)': 'America/Mazatlan',
  'Mountain Standard Time': 'America/Denver',
  'Yukon Standard Time': 'America/Whitehorse',
  'Central America Standard Time': 'America/Guatemala',
  'Central Standard Time': 'America/Chicago',
  'Easter Island Standard Time': 'Pacific/Easter',
  'Central Standard Time (Mexico)': 'America/Mexico_City',
  'Canada Central Standard Time': 'America/Regina',
  'SA Pacific Standard Time': 'America/Bogota',
  'Eastern Standard Time (Mexico)': 'America/Cancun',
  'Eastern Standard Time': 'America/New_York',
  'Haiti Standard Time': 'America/Port-au-Prince',
  'Cuba Standard Time': 'America/Havana',
  'US Eastern Standard Time': 'America/Indianapolis',
  'Turks And Caicos Standard Time': 'America/Grand_Turk',
  'Paraguay Standard Time': 'America/Asuncion',
  'Atlantic Standard Time': 'America/Halifax',
  'Venezuela Standard Time': 'America/Caracas',
  'Central Brazilian Standard Time': 'America/Cuiaba',
  'SA Western Standard Time': 'America/La_Paz',
  'Pacific SA Standard Time': 'America/Santiago',
  'Newfoundland Standard Time': 'America/St_Johns',
  'Tocantins Standard Time': 'America/Araguaina',
  'E. South America Standard Time': 'America/Sao_Paulo',
  'SA Eastern Standard Time': 'America/Cayenne',
  'Argentina Standard Time': 'America/Buenos_Aires',
  'Greenland Standard Time': 'America/Godthab',
  'Montevideo Standard Time': 'America/Montevideo',
  'Magallanes Standard Time': 'America/Punta_Arenas',
  'Saint Pierre Standard Time': 'America/Miquelon',
  'Bahia Standard Time': 'America/Bahia',
  'UTC-02': 'Etc/GMT+2',
  'Mid-Atlantic Standard Time': 'Etc/GMT+2',
  'Azores Standard Time': 'Atlantic/Azores',
  'Cape Verde Standard Time': 'Atlantic/Cape_Verde',
  'UTC': 'Etc/UTC',
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'Sao Tome Standard Time': 'Africa/Sao_Tome',
  'Morocco Standard Time': 'Africa/Casablanca',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw',
  'W. Central Africa Standard Time': 'Africa/Lagos',
  'Jordan Standard Time': 'Asia/Amman',
  'GTB Standard Time': 'Europe/Bucharest',
  'Middle East Standard Time': 'Asia/Beirut',
  'Egypt Standard Time': 'Africa/Cairo',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'Syria Standard Time': 'Asia/Damascus',
  'West Bank Standard Time': 'Asia/Hebron',
  'South Africa Standard Time': 'Africa/Johannesburg',
  'FLE Standard Time': 'Europe/Kiev',
  'Israel Standard Time': 'Asia/Jerusalem',
  'South Sudan Standard Time': 'Africa/Juba',
  'Kaliningrad Standard Time': 'Europe/Kaliningrad',
  'Sudan Standard Time': 'Africa/Khartoum',
  'Libya Standard Time': 'Africa/Tripoli',
  'Namibia Standard Time': 'Africa/Windhoek',
  'Arabic Standard Time': 'Asia/Baghdad',
  'Turkey Standard Time': 'Europe/Istanbul',
  'Arab Standard Time': 'Asia/Riyadh',
  'Belarus Standard Time': 'Europe/Minsk',
  'Russian Standard Time': 'Europe/Moscow',
  'E. Africa Standard Time': 'Africa/Nairobi',
  'Volgograd Standard Time': 'Europe/Volgograd',
  'Iran Standard Time': 'Asia/Tehran',
  'Arabian Standard Time': 'Asia/Dubai',
  'Astrakhan Standard Time': 'Europe/Astrakhan',
  'Azerbaijan Standard Time': 'Asia/Baku',
  'Russia Time Zone 3': 'Europe/Samara',
  'Mauritius Standard Time': 'Indian/Mauritius',
  'Saratov Standard Time': 'Europe/Saratov',
  'Georgian Standard Time': 'Asia/Tbilisi',
  'Caucasus Standard Time': 'Asia/Yerevan',
  'Afghanistan Standard Time': 'Asia/Kabul',
  'West Asia Standard Time': 'Asia/Tashkent',
  'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',
  'Pakistan Standard Time': 'Asia/Karachi',
  'Qyzylorda Standard Time': 'Asia/Qyzylorda',
  'India Standard Time': 'Asia/Kolkata',
  'Sri Lanka Standard Time': 'Asia/Colombo',
  'Nepal Standard Time': 'Asia/Kathmandu',
  'Central Asia Standard Time': 'Asia/Almaty',
  'Bangladesh Standard Time': 'Asia/Dhaka',
  'Omsk Standard Time': 'Asia/Omsk',
  'Myanmar Standard Time': 'Asia/Rangoon',
  'SE Asia Standard Time': 'Asia/Bangkok',
  'Altai Standard Time': 'Asia/Barnaul',
  'W. Mongolia Standard Time': 'Asia/Hovd',
  'North Asia Standard Time': 'Asia/Krasnoyarsk',
  'N. Central Asia Standard Time': 'Asia/Novosibirsk',
  'Tomsk Standard Time': 'Asia/Tomsk',
  'China Standard Time': 'Asia/Shanghai',
  'North Asia East Standard Time': 'Asia/Irkutsk',
  'Singapore Standard Time': 'Asia/Singapore',
  'W. Australia Standard Time': 'Australia/Perth',
  'Taipei Standard Time': 'Asia/Taipei',
  'Ulaanbaatar Standard Time': 'Asia/Ulaanbaatar',
  'Aus Central W. Standard Time': 'Australia/Eucla',
  'Transbaikal Standard Time': 'Asia/Chita',
  'Tokyo Standard Time': 'Asia/Tokyo',
  'North Korea Standard Time': 'Asia/Pyongyang',
  'Korea Standard Time': 'Asia/Seoul',
  'Yakutsk Standard Time': 'Asia/Yakutsk',
  'Cen. Australia Standard Time': 'Australia/Adelaide',
  'AUS Central Standard Time': 'Australia/Darwin',
  'E. Australia Standard Time': 'Australia/Brisbane',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  'West Pacific Standard Time': 'Pacific/Port_Moresby',
  'Tasmania Standard Time': 'Australia/Hobart',
  'Vladivostok Standard Time': 'Asia/Vladivostok',
  'Lord Howe Standard Time': 'Australia/Lord_Howe',
  'Bougainville Standard Time': 'Pacific/Bougainville',
  'Russia Time Zone 10': 'Asia/Srednekolymsk',
  'Magadan Standard Time': 'Asia/Magadan',
  'Norfolk Standard Time': 'Pacific/Norfolk',
  'Sakhalin Standard Time': 'Asia/Sakhalin',
  'Central Pacific Standard Time': 'Pacific/Guadalcanal',
  'Russia Time Zone 11': 'Asia/Kamchatka',
  'New Zealand Standard Time': 'Pacific/Auckland',
  'UTC+12': 'Etc/GMT-12',
  'Fiji Standard Time': 'Pacific/Fiji',
  'Chatham Islands Standard Time': 'Pacific/Chatham',
  'UTC+13': 'Etc/GMT-13',
  'Tonga Standard Time': 'Pacific/Tongatapu',
  'Samoa Standard Time': 'Pacific/Apia',
  'Line Islands Standard Time': 'Pacific/Kiritimati',
};

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Resolve a Windows timezone name to an IANA timezone identifier.
 * Falls back to America/Los_Angeles with a warning if unmapped.
 * @param {string} windowsTzName
 * @returns {string} IANA timezone
 */
function resolveIanaTimezone(windowsTzName) {
  if (!windowsTzName) windowsTzName = config.DEFAULT_TIMEZONE;
  const iana = WINDOWS_TO_IANA[windowsTzName];
  if (iana) return iana;

  // Maybe it's already an IANA name — try it with Intl
  try {
    Intl.DateTimeFormat('en-US', { timeZone: windowsTzName });
    return windowsTzName;
  } catch {
    // Not a valid IANA name either
  }

  console.warn(`[date-helpers] Unmapped timezone "${windowsTzName}", falling back to America/Los_Angeles. Add it to WINDOWS_TO_IANA.`);
  return 'America/Los_Angeles';
}

/**
 * Get the current wall-clock date/time in a given IANA timezone.
 * @param {string} ianaTz - IANA timezone identifier
 * @param {Date} [now] - Optional Date object for testability
 * @returns {{ year: number, month: number, day: number, hour: number, minute: number }}
 */
function getNowInTimezone(ianaTz, now) {
  if (!now) now = new Date();
  const parts = {};
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: ianaTz,
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false,
  });
  for (const { type, value } of formatter.formatToParts(now)) {
    if (type === 'year') parts.year = parseInt(value, 10);
    else if (type === 'month') parts.month = parseInt(value, 10);
    else if (type === 'day') parts.day = parseInt(value, 10);
    else if (type === 'hour') parts.hour = parseInt(value, 10);
    else if (type === 'minute') parts.minute = parseInt(value, 10);
  }
  return parts;
}

/**
 * Format a 12-hour time string from hour and minute.
 * @param {number} h - Hour (0-23)
 * @param {number} m - Minute
 * @returns {string} e.g. "2:30 PM"
 */
function fmt12h(h, m) {
  const period = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 || 12;
  return `${h12}:${String(m).padStart(2, '0')} ${period}`;
}

/**
 * Format a Graph API event datetime into agent-friendly text.
 * The datetime from Graph is already in the user's timezone (via Prefer header).
 *
 * @param {string} isoDateTime - e.g. "2026-03-11T10:00:00.0000000"
 * @param {string} ianaTz - IANA timezone (used only for today/tomorrow comparison)
 * @param {{ year: number, month: number, day: number }} todayParts - Today's date parts
 * @returns {string} e.g. "Wed, Mar 11 · 10:00 AM" or "Mon, Mar 9 · 2:00 PM (Today)"
 */
function formatEventTime(isoDateTime, ianaTz, todayParts) {
  // Parse date and time from the ISO string (already in user's TZ)
  const [datePart, timePart] = isoDateTime.split('T');
  const [yStr, mStr, dStr] = datePart.split('-');
  const y = parseInt(yStr, 10);
  const mo = parseInt(mStr, 10);
  const d = parseInt(dStr, 10);

  // Parse time
  const [hStr, minStr] = timePart.split(':');
  const h = parseInt(hStr, 10);
  const min = parseInt(minStr, 10);

  // Day of week — construct a UTC date with the same y/m/d to get correct day name
  const dow = DAY_NAMES[new Date(Date.UTC(y, mo - 1, d)).getUTCDay()];
  const monthName = MONTH_NAMES[mo - 1];

  // Relative marker
  let marker = '';
  if (y === todayParts.year && mo === todayParts.month && d === todayParts.day) {
    marker = ' (Today)';
  } else {
    // Check tomorrow
    const todayDate = new Date(Date.UTC(todayParts.year, todayParts.month - 1, todayParts.day));
    const tomorrowDate = new Date(todayDate.getTime() + 24 * 60 * 60 * 1000);
    if (y === tomorrowDate.getUTCFullYear() &&
        mo === tomorrowDate.getUTCMonth() + 1 &&
        d === tomorrowDate.getUTCDate()) {
      marker = ' (Tomorrow)';
    }
  }

  // Omit year if current year
  const yearSuffix = y !== todayParts.year ? `, ${y}` : '';

  return {
    formatted: `${dow}, ${monthName} ${d}${yearSuffix} · ${fmt12h(h, min)}${marker}`,
    dateKey: datePart, // for comparing start/end dates
    year: y, month: mo, day: d,
  };
}

/**
 * Format a combined start–end time string for an event.
 * Combines onto one line when same day, separate lines when spanning days.
 *
 * @param {string} startIso - Start datetime ISO string
 * @param {string} endIso - End datetime ISO string
 * @param {string} ianaTz - IANA timezone
 * @param {{ year: number, month: number, day: number }} todayParts
 * @returns {string}
 */
function formatEventRange(startIso, endIso, ianaTz, todayParts) {
  const start = formatEventTime(startIso, ianaTz, todayParts);
  const end = formatEventTime(endIso, ianaTz, todayParts);

  if (start.dateKey === end.dateKey) {
    // Same day — extract just the time portion from end
    const [, timePart] = endIso.split('T');
    const [hStr, minStr] = timePart.split(':');
    const h = parseInt(hStr, 10);
    const min = parseInt(minStr, 10);
    return `${start.formatted} – ${fmt12h(h, min)}`;
  }
  // Multi-day
  return `${start.formatted} –\n   ${end.formatted}`;
}

/**
 * Format an all-day event date range.
 * Graph returns all-day events as midnight-to-next-midnight, e.g.
 * start: 2026-03-09T00:00:00, end: 2026-03-10T00:00:00 for a single day.
 *
 * @param {string} startIso - Start datetime ISO string
 * @param {string} endIso - End datetime ISO string
 * @param {{ year: number, month: number, day: number }} todayParts
 * @returns {string} e.g. "Mon, Mar 9 (All day)" or "Mon, Mar 9 – Wed, Mar 11 (All day)"
 */
function formatAllDayRange(startIso, endIso, todayParts) {
  const [startDate] = startIso.split('T');
  const [sY, sM, sD] = startDate.split('-').map(Number);

  // End date from Graph is exclusive (day after last day), so subtract 1
  const [endDateStr] = endIso.split('T');
  const [endY, endM, endD] = endDateStr.split('-').map(Number);
  const endExclusive = new Date(Date.UTC(endY, endM - 1, endD));
  endExclusive.setUTCDate(endExclusive.getUTCDate() - 1);
  const eY = endExclusive.getUTCFullYear();
  const eM = endExclusive.getUTCMonth() + 1;
  const eD = endExclusive.getUTCDate();

  const startDow = DAY_NAMES[new Date(Date.UTC(sY, sM - 1, sD)).getUTCDay()];
  const startMonth = MONTH_NAMES[sM - 1];
  const startYearSuffix = sY !== todayParts.year ? `, ${sY}` : '';

  // Relative marker for start
  let marker = '';
  if (sY === todayParts.year && sM === todayParts.month && sD === todayParts.day) {
    marker = ' (Today)';
  } else {
    const todayDate = new Date(Date.UTC(todayParts.year, todayParts.month - 1, todayParts.day));
    const tomorrowDate = new Date(todayDate.getTime() + 24 * 60 * 60 * 1000);
    if (sY === tomorrowDate.getUTCFullYear() &&
        sM === tomorrowDate.getUTCMonth() + 1 &&
        sD === tomorrowDate.getUTCDate()) {
      marker = ' (Tomorrow)';
    }
  }

  // Single-day all-day event
  if (sY === eY && sM === eM && sD === eD) {
    return `${startDow}, ${startMonth} ${sD}${startYearSuffix} (All day)${marker}`;
  }

  // Multi-day all-day event
  const endDow = DAY_NAMES[new Date(Date.UTC(eY, eM - 1, eD)).getUTCDay()];
  const endMonth = MONTH_NAMES[eM - 1];
  const endYearSuffix = eY !== todayParts.year ? `, ${eY}` : '';
  return `${startDow}, ${startMonth} ${sD}${startYearSuffix} – ${endDow}, ${endMonth} ${eD}${endYearSuffix} (All day)${marker}`;
}

/**
 * Generate a reference timestamp header for calendar responses.
 * @param {string} ianaTz - IANA timezone
 * @param {Date} [now] - Optional Date for testability
 * @returns {string} e.g. "Current time: Mon, Mar 9, 2026 11:58 AM PDT (UTC-7)"
 */
function formatReferenceTimestamp(ianaTz, now) {
  if (!now) now = new Date();

  // Get abbreviated timezone name (e.g. "PDT")
  const abbrFormatter = new Intl.DateTimeFormat('en-US', {
    timeZone: ianaTz,
    timeZoneName: 'short',
  });
  let tzAbbr = '';
  for (const { type, value } of abbrFormatter.formatToParts(now)) {
    if (type === 'timeZoneName') tzAbbr = value;
  }

  // Get offset (e.g. "GMT-7")
  const offsetFormatter = new Intl.DateTimeFormat('en-US', {
    timeZone: ianaTz,
    timeZoneName: 'shortOffset',
  });
  let offsetStr = '';
  for (const { type, value } of offsetFormatter.formatToParts(now)) {
    if (type === 'timeZoneName') offsetStr = value;
  }
  // Convert "GMT-7" → "UTC-7"
  offsetStr = offsetStr.replace('GMT', 'UTC');

  const parts = getNowInTimezone(ianaTz, now);
  const dow = DAY_NAMES[new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay()];
  const monthName = MONTH_NAMES[parts.month - 1];

  return `Current time: ${dow}, ${monthName} ${parts.day}, ${parts.year} ${fmt12h(parts.hour, parts.minute)} ${tzAbbr} (${offsetStr})`;
}

/**
 * Format an email receivedDateTime (UTC ISO string) into the user's timezone.
 * @param {string} isoDateString - e.g. "2026-03-09T21:30:00Z"
 * @param {string} ianaTz - IANA timezone
 * @returns {string} e.g. "Mon, Mar 9 · 2:30 PM" or "Mon, Mar 9, 2024 · 2:30 PM"
 */
function formatEmailDate(isoDateString, ianaTz) {
  const date = new Date(isoDateString);
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: ianaTz,
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false,
    weekday: 'short',
  });

  const parts = {};
  for (const { type, value } of formatter.formatToParts(date)) {
    parts[type] = value;
  }

  const monthName = MONTH_NAMES[parseInt(parts.month, 10) - 1];
  // Use the configured timezone to determine current year, not the server's local clock
  const currentYear = getNowInTimezone(ianaTz).year;
  const year = parseInt(parts.year, 10);
  const yearSuffix = year !== currentYear ? `, ${year}` : '';
  const h = parseInt(parts.hour, 10);
  const m = parseInt(parts.minute, 10);

  return `${parts.weekday}, ${monthName} ${parseInt(parts.day, 10)}${yearSuffix} · ${fmt12h(h, m)}`;
}

module.exports = {
  WINDOWS_TO_IANA,
  resolveIanaTimezone,
  getNowInTimezone,
  formatEventTime,
  formatEventRange,
  formatAllDayRange,
  formatReferenceTimestamp,
  formatEmailDate,
};
