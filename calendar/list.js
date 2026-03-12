/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveIanaTimezone, getNowInTimezone, formatEventRange, formatAllDayRange, formatReferenceTimestamp } = require('../utils/date-helpers');

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Build API endpoint — calendarView expands recurring events
    // and requires startDateTime/endDateTime
    let endpoint = 'me/calendarView';

    const now = new Date();
    const thirtyDaysOut = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    // Add query parameters
    const queryParams = {
      startDateTime: now.toISOString(),
      endDateTime: thirtyDaysOut.toISOString(),
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };
    
    // Request times in the user's configured timezone
    const preferHeaders = {
      'Prefer': `outlook.timezone="${config.DEFAULT_TIMEZONE}"`
    };

    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams, preferHeaders);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No calendar events found."
        }]
      };
    }
    
    // Resolve timezone and compute reference info
    const ianaTz = resolveIanaTimezone(config.DEFAULT_TIMEZONE);
    const todayParts = getNowInTimezone(ianaTz);
    const refTimestamp = formatReferenceTimestamp(ianaTz);

    // Format results
    const eventList = response.value.map((event, index) => {
      const cancelPrefix = event.isCancelled ? '[CANCELED] ' : '';
      const timeRange = event.isAllDay
        ? formatAllDayRange(event.start.dateTime, event.end.dateTime, todayParts)
        : formatEventRange(event.start.dateTime, event.end.dateTime, ianaTz, todayParts);
      const location = event.location?.displayName;
      const locationLine = location ? `\n   Location: ${location}` : '';
      const summaryLine = event.bodyPreview ? `\n   Summary: ${event.bodyPreview}` : '';

      return `${index + 1}. ${cancelPrefix}${event.subject}\n   ${timeRange}${locationLine}${summaryLine}\n   ID: ${event.id}\n`;
    }).join("\n");

    return {
      content: [{
        type: "text",
        text: `${refTimestamp}\n\nFound ${response.value.length} events:\n\n${eventList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error listing events: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEvents;
