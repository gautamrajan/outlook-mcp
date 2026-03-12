jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../auth/request-context', () => ({
  isHostedMode: jest.fn(),
  getUserContext: jest.fn(),
}));
jest.mock('../../auth/embedded-server', () => ({
  startEmbeddedServer: jest.fn(),
}));
jest.mock('../../auth/hosted-config', () => ({
  getConfiguredServerBaseUrl: jest.fn(),
}));

const handleListAttachments = require('../../email/list-attachments');
const handleGetAttachmentDownloadUrl = require('../../email/get-attachment-download-url');
const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { isHostedMode, getUserContext } = require('../../auth/request-context');
const { startEmbeddedServer } = require('../../auth/embedded-server');
const { getConfiguredServerBaseUrl } = require('../../auth/hosted-config');
const { clearDownloadTickets } = require('../../email/download-ticket-store');

describe('attachment tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    clearDownloadTickets();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    isHostedMode.mockReturnValue(false);
    getUserContext.mockReturnValue(null);
    startEmbeddedServer.mockResolvedValue('http://localhost:3333');
    getConfiguredServerBaseUrl.mockReturnValue('https://outlook.example.com');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('list-attachments requires emailId', async () => {
    const result = await handleListAttachments({});
    expect(result.content[0].text).toBe('Email ID is required to list attachments.');
  });

  test('list-attachments formats supported and unsupported attachments', async () => {
    callGraphAPIPaginated.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-file',
          name: 'report.txt',
          contentType: 'text/plain',
          size: 128,
          isInline: false,
        },
        {
          '@odata.type': '#microsoft.graph.itemAttachment',
          id: 'att-item',
          name: 'forwarded.eml',
          contentType: 'message/rfc822',
          size: 512,
          isInline: false,
        },
        {
          '@odata.type': '#microsoft.graph.referenceAttachment',
          id: 'att-ref',
          name: 'shared.url',
          contentType: 'application/octet-stream',
          size: 64,
          isInline: false,
        }
      ]
    });

    const result = await handleListAttachments({ emailId: 'email-1' });

    expect(result.content[0].text).toContain('Found 3 attachments for email email-1');
    expect(result.content[0].text).toContain('ID: att-file');
    expect(result.content[0].text).toContain('Download Supported: Yes');
    expect(result.content[0].text).toContain('ID: att-item');
    expect(result.content[0].text).toContain('item attachments are not supported in v1');
    expect(result.content[0].text).toContain('ID: att-ref');
    expect(result.content[0].text).toContain('reference attachments are not supported in v1');
  });

  test('list-attachments handles empty results', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [] });

    const result = await handleListAttachments({ emailId: 'email-1' });

    expect(result.content[0].text).toBe('No attachments found for email email-1.');
  });

  test('list-attachments handles authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListAttachments({ emailId: 'email-1' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('list-attachments handles Graph errors', async () => {
    callGraphAPIPaginated.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleListAttachments({ emailId: 'email-1' });

    expect(result.content[0].text).toBe('Error listing attachments: Graph API Error');
  });

  test('get-attachment-download-url requires emailId', async () => {
    const result = await handleGetAttachmentDownloadUrl({ attachmentId: 'att-file' });
    expect(result.content[0].text).toBe(
      'Email ID is required to create an attachment download URL.'
    );
  });

  test('get-attachment-download-url requires attachmentId', async () => {
    const result = await handleGetAttachmentDownloadUrl({ emailId: 'email-1' });
    expect(result.content[0].text).toBe(
      'Attachment ID is required to create an attachment download URL.'
    );
  });

  test('get-attachment-download-url returns a localhost URL in local mode', async () => {
    callGraphAPIPaginated.mockResolvedValue({
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

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-file'
    });

    expect(startEmbeddedServer).toHaveBeenCalledTimes(1);
    expect(result.content[0].text).toContain('http://localhost:3333/attachments/download/');
    expect(result.content[0].text).toContain('curl -L "http://localhost:3333/attachments/download/');
  });

  test('get-attachment-download-url returns a hosted URL in hosted mode', async () => {
    isHostedMode.mockReturnValue(true);
    getUserContext.mockReturnValue({
      userId: 'user-oid',
      authMethod: 'connector',
      entraToken: 'connector-jwt',
    });
    callGraphAPIPaginated.mockResolvedValue({
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

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-file'
    });

    expect(getConfiguredServerBaseUrl).toHaveBeenCalled();
    expect(result.content[0].text).toContain('https://outlook.example.com/attachments/download/');
  });

  test('get-attachment-download-url rejects unsupported types', async () => {
    callGraphAPIPaginated.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.itemAttachment',
          id: 'att-item',
          name: 'forwarded.eml',
          contentType: 'message/rfc822',
          size: 512,
          isInline: false,
        }
      ]
    });

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-item'
    });

    expect(result.content[0].text).toBe(
      'Attachment forwarded.eml cannot be downloaded: item attachments are not supported in v1'
    );
  });

  test('get-attachment-download-url rejects oversize files', async () => {
    callGraphAPIPaginated.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-big',
          name: 'large.bin',
          contentType: 'application/octet-stream',
          size: (25 * 1024 * 1024) + 1,
          isInline: false,
        }
      ]
    });

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-big'
    });

    expect(result.content[0].text).toBe(
      'Attachment large.bin cannot be downloaded: attachment exceeds the maximum supported size of 25 MB'
    );
  });

  test('get-attachment-download-url rejects hosted mode without a public base URL', async () => {
    isHostedMode.mockReturnValue(true);
    getConfiguredServerBaseUrl.mockReturnValue(null);
    callGraphAPIPaginated.mockResolvedValue({
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

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-file'
    });

    expect(result.content[0].text).toBe(
      'Attachment downloads in hosted mode require PUBLIC_BASE_URL or HOSTED_REDIRECT_URI to be configured.'
    );
  });

  test('get-attachment-download-url falls back to request-derived hosted base URL', async () => {
    isHostedMode.mockReturnValue(true);
    getConfiguredServerBaseUrl.mockReturnValue(null);
    getUserContext.mockReturnValue({
      userId: 'user-oid',
      authMethod: 'connector',
      entraToken: 'connector-jwt',
      serverBaseUrl: 'https://derived.example.com',
    });
    callGraphAPIPaginated.mockResolvedValue({
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

    const result = await handleGetAttachmentDownloadUrl({
      emailId: 'email-1',
      attachmentId: 'att-file'
    });

    expect(result.content[0].text).toContain('https://derived.example.com/attachments/download/');
  });
});
