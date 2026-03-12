const handleUpdateEvent = require('../../calendar/update');
const { calendarTools } = require('../../calendar');
const { DEFAULT_TIMEZONE } = require('../../config');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const expectedDateTimeSchema = {
  oneOf: [
    {
      type: 'string',
      description: 'An ISO 8601 date-time string'
    },
    {
      type: 'object',
      properties: {
        dateTime: {
          type: 'string',
          description: 'The date-time value in ISO 8601 format'
        },
        timeZone: {
          type: 'string',
          description: 'Optional Windows timezone name'
        }
      },
      required: ['dateTime']
    }
  ],
  description: 'The event time as an ISO 8601 string or a dateTime/timeZone object'
};

describe('handleUpdateEvent', () => {
  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
  });

  test('returns validation error when eventId is missing', async () => {
    const result = await handleUpdateEvent({ subject: 'Updated subject' });

    expect(result.content[0].text).toBe('Event ID is required to update an event.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('returns validation error when no mutable fields are provided', async () => {
    const result = await handleUpdateEvent({ eventId: 'event-1' });

    expect(result.content[0].text).toBe(
      'At least one event field must be provided to update an event.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('returns validation error when only start is provided', async () => {
    const result = await handleUpdateEvent({
      eventId: 'event-1',
      start: '2024-03-10T10:00:00'
    });

    expect(result.content[0].text).toBe(
      'Start and end times must be provided together when updating an event.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('returns validation error when only end is provided', async () => {
    const result = await handleUpdateEvent({
      eventId: 'event-1',
      end: '2024-03-10T11:00:00'
    });

    expect(result.content[0].text).toBe(
      'Start and end times must be provided together when updating an event.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('patches only the provided subject field', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    const result = await handleUpdateEvent({
      eventId: 'event-1',
      subject: 'Updated subject'
    });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      'dummy_access_token',
      'PATCH',
      'me/events/event-1',
      { subject: 'Updated subject' }
    );
    expect(result.content[0].text).toBe('Event with ID event-1 has been successfully updated.');
  });

  test('uses default timezone for string start and end inputs', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      start: '2024-03-10T10:00:00',
      end: '2024-03-10T11:00:00'
    });

    const payload = callGraphAPI.mock.calls[0][3];
    expect(payload).toEqual({
      start: {
        dateTime: '2024-03-10T10:00:00',
        timeZone: DEFAULT_TIMEZONE
      },
      end: {
        dateTime: '2024-03-10T11:00:00',
        timeZone: DEFAULT_TIMEZONE
      }
    });
  });

  test('preserves explicit timezone values for object start and end inputs', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      start: {
        dateTime: '2024-03-10T10:00:00',
        timeZone: 'Pacific Standard Time'
      },
      end: {
        dateTime: '2024-03-10T11:00:00'
      }
    });

    const payload = callGraphAPI.mock.calls[0][3];
    expect(payload).toEqual({
      start: {
        dateTime: '2024-03-10T10:00:00',
        timeZone: 'Pacific Standard Time'
      },
      end: {
        dateTime: '2024-03-10T11:00:00',
        timeZone: DEFAULT_TIMEZONE
      }
    });
  });

  test('does not send omitted optional fields', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      isAllDay: true
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({ isAllDay: true });
  });

  test('maps attendees correctly when provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      attendees: ['a@example.com', 'b@example.com']
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({
      attendees: [
        {
          emailAddress: { address: 'a@example.com' },
          type: 'required'
        },
        {
          emailAddress: { address: 'b@example.com' },
          type: 'required'
        }
      ]
    });
  });

  test('sends an empty attendee array to clear attendees', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      attendees: []
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({
      attendees: []
    });
  });

  test('sends an empty body string when provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      body: ''
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({
      body: {
        contentType: 'HTML',
        content: ''
      }
    });
  });

  test('sends an empty location display name when provided', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-1',
      location: ''
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({
      location: {
        displayName: ''
      }
    });
  });

  test('returns authentication required message on auth failure', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleUpdateEvent({
      eventId: 'event-1',
      subject: 'Updated subject'
    });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('returns Graph API error message on failure', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleUpdateEvent({
      eventId: 'event-1',
      subject: 'Updated subject'
    });

    expect(result.content[0].text).toBe('Error updating event: Graph API Error');
  });
});

describe('calendar event tool schemas', () => {
  const getTool = (name) => calendarTools.find(tool => tool.name === name);

  test('update-event exposes the expected mutable fields and requires only eventId', () => {
    const tool = getTool('update-event');

    expect(tool).toBeDefined();
    expect(Object.keys(tool.inputSchema.properties)).toEqual([
      'eventId',
      'subject',
      'start',
      'end',
      'attendees',
      'body',
      'location',
      'isAllDay'
    ]);
    expect(tool.inputSchema.properties.eventId).toEqual({
      type: 'string',
      description: 'The ID of the event to update'
    });
    expect(tool.inputSchema.properties.start).toEqual(expectedDateTimeSchema);
    expect(tool.inputSchema.properties.end).toEqual(expectedDateTimeSchema);
    expect(tool.inputSchema.required).toEqual(['eventId']);
  });

  test('create-event start and end schemas match the shared date-time schema', () => {
    const tool = getTool('create-event');

    expect(tool.inputSchema.properties.start).toEqual(expectedDateTimeSchema);
    expect(tool.inputSchema.properties.end).toEqual(expectedDateTimeSchema);
  });
});
