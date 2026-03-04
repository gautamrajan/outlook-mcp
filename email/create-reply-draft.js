/**
 * Create reply draft functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create reply draft handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateReplyDraft(args) {
  const { id, body, replyAll = true, includeChain = true } = args;

  if (!id) {
    return {
      content: [{
        type: "text",
        text: "Message ID is required."
      }]
    };
  }

  if (!body) {
    return {
      content: [{
        type: "text",
        text: "Reply body content is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // When includeChain is true (default): use `comment` field
    // Graph API prepends the comment above the auto-generated quoted chain
    // When false: use `message.body` which replaces the chain entirely
    const requestBody = includeChain
      ? { comment: body }
      : {
          message: {
            body: {
              contentType: body.includes('<html') ? 'html' : 'text',
              content: body
            }
          }
        };

    const endpoint = replyAll
      ? `me/messages/${id}/createReplyAll`
      : `me/messages/${id}/createReply`;

    const result = await callGraphAPI(accessToken, 'POST', endpoint, requestBody);

    const recipientCount = (result.toRecipients?.length || 0)
      + (result.ccRecipients?.length || 0);

    return {
      content: [{
        type: "text",
        text: `Reply draft created successfully!\n\nDraft ID: ${result.id}\nSubject: ${result.subject}\nRecipients: ${recipientCount}\nReply All: ${replyAll}\nMessage Length: ${body.length} characters\n\nThe reply draft is now in your Outlook Drafts folder, threaded with the original conversation.`
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
        text: `Error creating reply draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateReplyDraft;
