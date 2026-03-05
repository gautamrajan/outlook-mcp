/**
 * Delete email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Delete email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteEmail(args) {
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
    const endpoint = `me/messages/${encodeURIComponent(emailId)}`;

    try {
      await callGraphAPI(accessToken, 'DELETE', endpoint);

      return {
        content: [{
          type: "text",
          text: "Email successfully deleted."
        }]
      };
    } catch (error) {
      console.error(`Error deleting email: ${error.message}`);

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
          text: `Failed to delete email: ${error.message}`
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

module.exports = handleDeleteEmail;
