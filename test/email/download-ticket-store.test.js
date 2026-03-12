const config = require('../../config');
const {
  issueDownloadTicket,
  lookupDownloadTicket,
  markDownloadTicketConsumed,
  clearDownloadTickets,
} = require('../../email/download-ticket-store');

describe('download ticket store', () => {
  beforeEach(() => {
    clearDownloadTickets();
    jest.restoreAllMocks();
  });

  test('issues and looks up a valid ticket', () => {
    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 123,
    });

    const result = lookupDownloadTicket(ticket.token);

    expect(result.status).toBe('valid');
    expect(result.ticket.emailId).toBe('email-1');
    expect(result.ticket.attachmentId).toBe('att-1');
  });

  test('lazy cleanup removes expired tickets', () => {
    const nowSpy = jest.spyOn(Date, 'now');
    nowSpy.mockReturnValue(1000);

    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 123,
    });

    nowSpy.mockReturnValue(1000 + config.ATTACHMENT_DOWNLOAD_TTL_MS + 1);

    const result = lookupDownloadTicket(ticket.token);
    expect(result.status).toBe('expired');
  });

  test('consumed tickets cannot be reused', () => {
    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 123,
    });

    expect(markDownloadTicketConsumed(ticket.token)).toBe(true);
    expect(lookupDownloadTicket(ticket.token).status).toBe('consumed');
  });

  test('tickets remain valid until explicitly consumed', () => {
    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 123,
    });

    expect(lookupDownloadTicket(ticket.token).status).toBe('valid');
    expect(lookupDownloadTicket(ticket.token).status).toBe('valid');
  });
});
