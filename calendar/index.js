/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleUpdateEvent = require('./update');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
const handleAcceptEvent = require('./accept');
const handleTentativelyAcceptEvent = require('./tentatively-accept');

const dateTimeInputSchema = {
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

// Calendar tool definitions
const calendarTools = [
  {
    name: "list-events",
    description: "Lists upcoming events from your calendar",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of events to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "decline-event",
    description: "Declines a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to decline"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "create-event",
    description: "Creates a new calendar event",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "The subject of the event"
        },
        start: dateTimeInputSchema,
        end: dateTimeInputSchema,
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "List of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Optional body content for the event"
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "update-event",
    description: "Updates an existing calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to update"
        },
        subject: {
          type: "string",
          description: "The updated subject of the event"
        },
        start: dateTimeInputSchema,
        end: dateTimeInputSchema,
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "Replacement list of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Updated body content for the event"
        },
        location: {
          type: "string",
          description: "Updated display name for the event location"
        },
        isAllDay: {
          type: "boolean",
          description: "Whether the event should be marked as all-day"
        }
      },
      required: ["eventId"]
    },
    handler: handleUpdateEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to cancel"
        },
        comment: {
          type: "string",
          description: "Optional comment for cancelling the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to delete"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  },
  {
    name: "accept-event",
    description: "Accepts a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to accept"
        }
      },
      required: ["eventId"]
    },
    handler: handleAcceptEvent
  },
  {
    name: "tentatively-accept-event",
    description: "Tentatively accepts a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to tentatively accept"
        }
      },
      required: ["eventId"]
    },
    handler: handleTentativelyAcceptEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleDeclineEvent,
  handleCreateEvent,
  handleUpdateEvent,
  handleCancelEvent,
  handleDeleteEvent,
  handleAcceptEvent,
  handleTentativelyAcceptEvent
};
