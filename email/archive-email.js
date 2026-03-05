/**
 * Archive email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Archive email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleArchiveEmail(args) {
  const emailId = args.id;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/messages/${encodeURIComponent(emailId)}/move`;
    const moveData = { destinationId: 'archive' };

    try {
      const result = await callGraphAPI(accessToken, 'POST', endpoint, moveData);
      const movedMessageId = result?.id ? `\nArchived Message ID: ${result.id}` : '';

      return {
        content: [{
          type: "text",
          text: `Email successfully archived.${movedMessageId}`
        }]
      };
    } catch (error) {
      console.error(`Error archiving email: ${error.message}`);

      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [{
            type: "text",
            text: "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID."
          }]
        };
      }

      if (error.message.includes("UNAUTHORIZED")) {
        return {
          content: [{
            type: "text",
            text: "Authentication failed. Please re-authenticate and try again."
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Failed to archive email: ${error.message}`
        }]
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required' || error.code === 'AUTH_REQUIRED') {
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
        text: `Error accessing email: ${error.message}`
      }]
    };
  }
}

module.exports = handleArchiveEmail;
