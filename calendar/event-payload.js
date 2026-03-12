/**
 * Shared Graph payload builders for calendar event mutations.
 */
const { DEFAULT_TIMEZONE } = require('../config');

function hasOwn(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

function normalizeDateTime(value) {
  return {
    dateTime: value.dateTime || value,
    timeZone: value.timeZone || DEFAULT_TIMEZONE
  };
}

function mapAttendees(attendees) {
  return attendees.map(email => ({
    emailAddress: { address: email },
    type: 'required'
  }));
}

function buildCreateEventPayload(args) {
  return {
    subject: args.subject,
    start: normalizeDateTime(args.start),
    end: normalizeDateTime(args.end),
    attendees: args.attendees ? mapAttendees(args.attendees) : undefined,
    body: { contentType: 'HTML', content: args.body || '' }
  };
}

function buildUpdateEventPayload(args) {
  const payload = {};

  if (hasOwn(args, 'subject')) {
    payload.subject = args.subject;
  }

  if (hasOwn(args, 'start')) {
    payload.start = normalizeDateTime(args.start);
  }

  if (hasOwn(args, 'end')) {
    payload.end = normalizeDateTime(args.end);
  }

  if (hasOwn(args, 'attendees')) {
    payload.attendees = mapAttendees(args.attendees);
  }

  if (hasOwn(args, 'body')) {
    payload.body = { contentType: 'HTML', content: args.body };
  }

  if (hasOwn(args, 'location')) {
    payload.location = { displayName: args.location };
  }

  if (hasOwn(args, 'isAllDay')) {
    payload.isAllDay = args.isAllDay;
  }

  return payload;
}

module.exports = {
  buildCreateEventPayload,
  buildUpdateEventPayload
};
