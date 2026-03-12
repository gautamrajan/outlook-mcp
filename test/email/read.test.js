jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const handleReadEmail = require('../../email/read');
const { callGraphAPI, callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

describe('handleReadEmail', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('returns email body unchanged when there are no attachments', async () => {
    callGraphAPI.mockResolvedValue({
      id: 'email-1',
      subject: 'No Attachments',
      from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
      toRecipients: [],
      ccRecipients: [],
      bccRecipients: [],
      receivedDateTime: '2024-06-01T10:00:00Z',
      body: { contentType: 'text', content: 'Hello world' },
      hasAttachments: false,
      importance: 'normal',
    });

    const result = await handleReadEmail({ id: 'email-1' });

    expect(result.content[0].text).toContain('Subject: No Attachments');
    expect(result.content[0].text).not.toMatch(/\n\nAttachments:\n/);
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('appends attachment summary when attachments exist', async () => {
    callGraphAPI.mockResolvedValueOnce({
      id: 'email-1',
      subject: 'With Attachments',
      from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
      toRecipients: [],
      ccRecipients: [],
      bccRecipients: [],
      receivedDateTime: '2024-06-01T10:00:00Z',
      body: { contentType: 'text', content: 'Hello world' },
      hasAttachments: true,
      importance: 'normal',
    });
    callGraphAPIPaginated.mockResolvedValueOnce({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-file',
          name: 'report.txt',
          contentType: 'text/plain',
          size: 128,
          isInline: false,
        }
      ]
    });

    const result = await handleReadEmail({ id: 'email-1' });

    expect(result.content[0].text).toContain('Attachments:');
    expect(result.content[0].text).toContain('report.txt');
    expect(result.content[0].text).toContain('ID: att-file');
    expect(result.content[0].text).toContain('Download Supported: Yes');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(1);
  });

  test('appends warning when attachment lookup fails', async () => {
    callGraphAPI.mockResolvedValueOnce({
      id: 'email-1',
      subject: 'With Attachments',
      from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
      toRecipients: [],
      ccRecipients: [],
      bccRecipients: [],
      receivedDateTime: '2024-06-01T10:00:00Z',
      body: { contentType: 'text', content: 'Hello world' },
      hasAttachments: true,
      importance: 'normal',
    });
    callGraphAPIPaginated.mockRejectedValueOnce(new Error('attachment lookup failed'));

    const result = await handleReadEmail({ id: 'email-1' });

    expect(result.content[0].text).toContain(
      'Attachments: unable to load attachment metadata (attachment lookup failed)'
    );
  });
});
