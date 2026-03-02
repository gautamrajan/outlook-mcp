/**
 * Create email draft functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create email draft handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateDraft(args) {
  const { to, cc, bcc, subject, body, importance = 'normal' } = args;

  if (!subject) {
    return {
      content: [{
        type: "text",
        text: "Subject is required."
      }]
    };
  }

  if (!body) {
    return {
      content: [{
        type: "text",
        text: "Body content is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Format recipients
    const toRecipients = to ? to.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    const ccRecipients = cc ? cc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    const bccRecipients = bcc ? bcc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    // Build message object (sent directly, not wrapped like sendMail)
    const messageObject = {
      subject,
      body: {
        contentType: body.includes('<html') ? 'html' : 'text',
        content: body
      },
      importance
    };

    if (toRecipients.length > 0) messageObject.toRecipients = toRecipients;
    if (ccRecipients.length > 0) messageObject.ccRecipients = ccRecipients;
    if (bccRecipients.length > 0) messageObject.bccRecipients = bccRecipients;

    // POST /me/messages creates a draft in the Drafts folder
    const result = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);

    return {
      content: [{
        type: "text",
        text: `Draft created successfully!\n\nSubject: ${subject}\nDraft ID: ${result.id}\nRecipients: ${toRecipients.length}${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ''}${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ''}\nMessage Length: ${body.length} characters\n\nThe draft is now in your Outlook Drafts folder.`
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
        text: `Error creating draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateDraft;
