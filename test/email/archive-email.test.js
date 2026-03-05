const handleArchiveEmail = require('../../email/archive-email');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleArchiveEmail', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should require email id', async () => {
    const result = await handleArchiveEmail({});

    expect(result.content[0].text).toBe('Email ID is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('should archive email successfully', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ id: 'archived-message-id-1' });

    const result = await handleArchiveEmail({ id: 'email-123' });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      'me/messages/email-123/move',
      { destinationId: 'archive' }
    );
    expect(result.content[0].text).toContain('Email successfully archived.');
    expect(result.content[0].text).toContain('Archived Message ID: archived-message-id-1');
  });

  test('should encode message id in endpoint path', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ id: 'archived-message-id-2' });
    const rawId = 'AAMkAGI0L2ZvbytiYXI9PQ==';

    await handleArchiveEmail({ id: rawId });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'POST',
      `me/messages/${encodeURIComponent(rawId)}/move`,
      { destinationId: 'archive' }
    );
  });

  test('should return auth required when authentication is missing', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleArchiveEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('should return auth required when auth error code is AUTH_REQUIRED', async () => {
    const authError = new Error('Token refresh failed');
    authError.code = 'AUTH_REQUIRED';
    ensureAuthenticated.mockRejectedValue(authError);

    const result = await handleArchiveEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('should handle mailbox ownership mismatch errors', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(
      new Error("The specified object doesn't belong to the targeted mailbox.")
    );

    const result = await handleArchiveEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID."
    );
  });

  test('should surface generic API errors', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleArchiveEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe('Failed to archive email: Graph API Error');
  });
});
