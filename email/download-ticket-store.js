const crypto = require('crypto');
const config = require('../config');

const tickets = new Map();

function cleanupExpiredTickets() {
  const now = Date.now();

  for (const [token, ticket] of tickets.entries()) {
    if (ticket.expiresAt <= now) {
      tickets.delete(token);
    }
  }
}

function issueDownloadTicket(ticketData) {
  cleanupExpiredTickets();

  const token = crypto.randomBytes(24).toString('base64url');
  const ticket = {
    token,
    expiresAt: Date.now() + config.ATTACHMENT_DOWNLOAD_TTL_MS,
    consumed: false,
    ...ticketData,
  };

  tickets.set(token, ticket);
  return ticket;
}

function lookupDownloadTicket(token) {
  if (!token || !tickets.has(token)) {
    cleanupExpiredTickets();
    return { status: 'invalid', ticket: null };
  }

  const ticket = tickets.get(token);
  if (ticket.expiresAt <= Date.now()) {
    tickets.delete(token);
    return { status: 'expired', ticket: null };
  }

  if (ticket.consumed === true) {
    tickets.delete(token);
    return { status: 'consumed', ticket: null };
  }

  cleanupExpiredTickets();
  return { status: 'valid', ticket };
}

function markDownloadTicketConsumed(token) {
  const result = lookupDownloadTicket(token);
  if (result.status !== 'valid') {
    return false;
  }

  const ticket = tickets.get(token);
  ticket.consumed = true;
  return true;
}

function clearDownloadTickets() {
  tickets.clear();
}

module.exports = {
  issueDownloadTicket,
  lookupDownloadTicket,
  markDownloadTicketConsumed,
  clearDownloadTickets,
};
