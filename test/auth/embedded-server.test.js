const { Readable } = require('stream');

jest.mock('../../utils/graph-api', () => ({
  streamGraphAPI: jest.fn(),
}));

const { streamGraphAPI } = require('../../utils/graph-api');
const { issueDownloadTicket, clearDownloadTickets } = require('../../email/download-ticket-store');
const { startEmbeddedServer, stopAuthServer } = require('../../auth/embedded-server');

describe('embedded server attachment downloads', () => {
  beforeEach(() => {
    clearDownloadTickets();
    jest.clearAllMocks();
  });

  afterEach(async () => {
    await stopAuthServer();
  });

  test('returns 404 for an invalid download token', async () => {
    const baseUrl = await startEmbeddedServer(0);
    const response = await fetch(`${baseUrl}/attachments/download/invalid-token`);

    expect(response.status).toBe(404);
    expect(await response.text()).toBe('Attachment download URL is invalid.');
  });

  test('returns 410 for an expired or consumed token', async () => {
    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 4,
    });
    const baseUrl = await startEmbeddedServer(0);

    streamGraphAPI.mockResolvedValue({
      statusCode: 200,
      headers: {
        'content-type': 'text/plain',
        'content-length': '4',
      },
      stream: Readable.from(Buffer.from('test')),
    });

    const first = await fetch(`${baseUrl}/attachments/download/${ticket.token}`);
    expect(first.status).toBe(200);

    const second = await fetch(`${baseUrl}/attachments/download/${ticket.token}`);
    expect(second.status).toBe(410);
    expect(await second.text()).toBe('Download URL expired. Request a new attachment download URL.');
  });

  test('streams a successful local attachment download with expected headers', async () => {
    const ticket = issueDownloadTicket({
      accessToken: 'token',
      authContext: null,
      emailId: 'email-1',
      attachmentId: 'att-1',
      name: 'report.txt',
      contentType: 'text/plain',
      size: 4,
    });
    const baseUrl = await startEmbeddedServer(0);

    streamGraphAPI.mockResolvedValue({
      statusCode: 200,
      headers: {
        'content-type': 'text/plain',
        'content-length': '4',
      },
      stream: Readable.from(Buffer.from('test')),
    });

    const response = await fetch(`${baseUrl}/attachments/download/${ticket.token}`);

    expect(response.status).toBe(200);
    expect(response.headers.get('cache-control')).toBe('no-store');
    expect(response.headers.get('content-disposition')).toContain('report.txt');
    expect(await response.text()).toBe('test');
  });
});
