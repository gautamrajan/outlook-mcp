const { ensureAuthenticated } = require('../auth');
const {
  listMessageAttachments,
  formatAttachmentListText,
} = require('./attachments');

async function handleListAttachments(args) {
  const emailId = args.emailId;

  if (!emailId) {
    return {
      content: [{
        type: 'text',
        text: 'Email ID is required to list attachments.'
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const attachments = await listMessageAttachments(accessToken, emailId);

    return {
      content: [{
        type: 'text',
        text: formatAttachmentListText(emailId, attachments),
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
        text: `Error listing attachments: ${error.message}`
      }]
    };
  }
}

module.exports = handleListAttachments;
