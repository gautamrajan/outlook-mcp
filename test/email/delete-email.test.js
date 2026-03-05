const handleDeleteEmail = require('../../email/delete-email');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleDeleteEmail', () => {
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
    const result = await handleDeleteEmail({});

    expect(result.content[0].text).toBe('Email ID is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('should delete email successfully', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({});

    const result = await handleDeleteEmail({ id: 'email-123' });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      'me/messages/email-123'
    );
    expect(result.content[0].text).toBe('Email successfully deleted.');
  });

  test('should encode message id in endpoint path', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({});
    const rawId = 'AAMkAGI0L2ZvbytiYXI9PQ==';

    await handleDeleteEmail({ id: rawId });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'DELETE',
      `me/messages/${encodeURIComponent(rawId)}`
    );
  });

  test('should return auth required when authentication is missing', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleDeleteEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('should return auth required when auth error code is AUTH_REQUIRED', async () => {
    const authError = new Error('Token refresh failed');
    authError.code = 'AUTH_REQUIRED';
    ensureAuthenticated.mockRejectedValue(authError);

    const result = await handleDeleteEmail({ id: 'email-123' });

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

    const result = await handleDeleteEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID."
    );
  });

  test('should surface generic API errors', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleDeleteEmail({ id: 'email-123' });

    expect(result.content[0].text).toBe('Failed to delete email: Graph API Error');
  });
});
