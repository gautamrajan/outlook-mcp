/**
 * Update event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { buildUpdateEventPayload } = require('./event-payload');

const MUTABLE_FIELDS = ['subject', 'start', 'end', 'attendees', 'body', 'location', 'isAllDay'];

function hasOwn(args, key) {
  return Object.prototype.hasOwnProperty.call(args, key);
}

/**
 * Update event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateEvent(args) {
  const { eventId } = args;

  if (!eventId) {
    return {
      content: [{
        type: 'text',
        text: 'Event ID is required to update an event.'
      }]
    };
  }

  const hasUpdateField = MUTABLE_FIELDS.some(field => hasOwn(args, field));
  if (!hasUpdateField) {
    return {
      content: [{
        type: 'text',
        text: 'At least one event field must be provided to update an event.'
      }]
    };
  }

  if (hasOwn(args, 'start') !== hasOwn(args, 'end')) {
    return {
      content: [{
        type: 'text',
        text: 'Start and end times must be provided together when updating an event.'
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/events/${encodeURIComponent(eventId)}`;
    const payload = buildUpdateEventPayload(args);

    await callGraphAPI(accessToken, 'PATCH', endpoint, payload);

    return {
      content: [{
        type: 'text',
        text: `Event with ID ${eventId} has been successfully updated.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: 'text',
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: 'text',
        text: `Error updating event: ${error.message}`
      }]
    };
  }
}

module.exports = handleUpdateEvent;
