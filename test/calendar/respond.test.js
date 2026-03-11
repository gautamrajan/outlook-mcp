const handleAcceptEvent = require('../../calendar/accept');
const handleDeclineEvent = require('../../calendar/decline');
const handleTentativelyAcceptEvent = require('../../calendar/tentatively-accept');
const { calendarTools } = require('../../calendar');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('calendar RSVP handlers', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
  });

  describe('handleAcceptEvent', () => {
    test('returns validation error when eventId is missing', async () => {
      const result = await handleAcceptEvent({});

      expect(result.content[0].text).toBe('Event ID is required to accept an event.');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('accepts without sending a response and ignores comment input', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockResolvedValue({});

      const result = await handleAcceptEvent({ eventId: 'event-1', comment: 'Please attend' });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        'dummy_access_token',
        'POST',
        'me/events/event-1/accept',
        { sendResponse: false }
      );
      expect(result.content[0].text).toBe(
        'Event with ID event-1 has been successfully accepted without sending a response.'
      );
    });

    test('returns authentication required message on auth failure', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleAcceptEvent({ eventId: 'event-1' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('returns Graph API error message on failure', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleAcceptEvent({ eventId: 'event-1' });

      expect(result.content[0].text).toBe('Error accepting event: Graph API Error');
    });
  });

  describe('handleDeclineEvent', () => {
    test('returns validation error when eventId is missing', async () => {
      const result = await handleDeclineEvent({});

      expect(result.content[0].text).toBe('Event ID is required to decline an event.');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('declines without sending a response and ignores comment input', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockResolvedValue({});

      const result = await handleDeclineEvent({ eventId: 'event-2', comment: 'Cannot make it' });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        'dummy_access_token',
        'POST',
        'me/events/event-2/decline',
        { sendResponse: false }
      );
      expect(result.content[0].text).toBe(
        'Event with ID event-2 has been successfully declined without sending a response.'
      );
    });

    test('returns authentication required message on auth failure', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleDeclineEvent({ eventId: 'event-2' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('returns Graph API error message on failure', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleDeclineEvent({ eventId: 'event-2' });

      expect(result.content[0].text).toBe('Error declining event: Graph API Error');
    });
  });

  describe('handleTentativelyAcceptEvent', () => {
    test('returns validation error when eventId is missing', async () => {
      const result = await handleTentativelyAcceptEvent({});

      expect(result.content[0].text).toBe('Event ID is required to tentatively accept an event.');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('tentatively accepts without sending a response and ignores comment input', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockResolvedValue({});

      const result = await handleTentativelyAcceptEvent({
        eventId: 'event-3',
        comment: 'Maybe'
      });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        'dummy_access_token',
        'POST',
        'me/events/event-3/tentativelyAccept',
        { sendResponse: false }
      );
      expect(result.content[0].text).toBe(
        'Event with ID event-3 has been tentatively accepted without sending a response.'
      );
    });

    test('returns authentication required message on auth failure', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleTentativelyAcceptEvent({ eventId: 'event-3' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('returns Graph API error message on failure', async () => {
      ensureAuthenticated.mockResolvedValue('dummy_access_token');
      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleTentativelyAcceptEvent({ eventId: 'event-3' });

      expect(result.content[0].text).toBe('Error tentatively accepting event: Graph API Error');
    });
  });
});

describe('calendar RSVP tool schemas', () => {
  const getTool = (name) => calendarTools.find(tool => tool.name === name);

  test('accept-event does not expose comment and still requires eventId', () => {
    const tool = getTool('accept-event');

    expect(tool.inputSchema.properties.comment).toBeUndefined();
    expect(tool.inputSchema.properties.eventId).toEqual({
      type: 'string',
      description: 'The ID of the event to accept'
    });
    expect(tool.inputSchema.required).toEqual(['eventId']);
  });

  test('decline-event does not expose comment and still requires eventId', () => {
    const tool = getTool('decline-event');

    expect(tool.inputSchema.properties.comment).toBeUndefined();
    expect(tool.inputSchema.properties.eventId).toEqual({
      type: 'string',
      description: 'The ID of the event to decline'
    });
    expect(tool.inputSchema.required).toEqual(['eventId']);
  });

  test('tentatively-accept-event does not expose comment and still requires eventId', () => {
    const tool = getTool('tentatively-accept-event');

    expect(tool.inputSchema.properties.comment).toBeUndefined();
    expect(tool.inputSchema.properties.eventId).toEqual({
      type: 'string',
      description: 'The ID of the event to tentatively accept'
    });
    expect(tool.inputSchema.required).toEqual(['eventId']);
  });
});
