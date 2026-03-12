const {
  emailTools,
  ENABLE_SEND_EMAIL_TOOL,
  handleSendEmail,
  handleListAttachments,
  handleGetAttachmentDownloadUrl,
} = require('../../email');

describe('email tool registration', () => {
  test('send-email is deregistered when the feature flag is disabled', () => {
    expect(ENABLE_SEND_EMAIL_TOOL).toBe(false);
    expect(emailTools.find((tool) => tool.name === 'send-email')).toBeUndefined();
  });

  test('send-email handler remains exported for easy re-enable later', () => {
    expect(typeof handleSendEmail).toBe('function');
  });

  test('attachment tools are registered', () => {
    expect(emailTools.find((tool) => tool.name === 'list-attachments')).toBeDefined();
    expect(emailTools.find((tool) => tool.name === 'get-attachment-download-url')).toBeDefined();
    expect(typeof handleListAttachments).toBe('function');
    expect(typeof handleGetAttachmentDownloadUrl).toBe('function');
  });
});
