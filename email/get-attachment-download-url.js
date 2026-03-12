const config = require('../config');
const { ensureAuthenticated } = require('../auth');
const { isHostedMode, getUserContext } = require('../auth/request-context');
const { getConfiguredServerBaseUrl } = require('../auth/hosted-config');
const { startEmbeddedServer } = require('../auth/embedded-server');
const {
  listMessageAttachments,
  sanitizeAttachmentFilename,
} = require('./attachments');
const { issueDownloadTicket } = require('./download-ticket-store');

async function handleGetAttachmentDownloadUrl(args) {
  const emailId = args.emailId;
  const attachmentId = args.attachmentId;

  if (!emailId) {
    return {
      content: [{
        type: 'text',
        text: 'Email ID is required to create an attachment download URL.'
      }]
    };
  }

  if (!attachmentId) {
    return {
      content: [{
        type: 'text',
        text: 'Attachment ID is required to create an attachment download URL.'
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const attachments = await listMessageAttachments(accessToken, emailId);
    const attachment = attachments.find((item) => item.id === attachmentId);

    if (!attachment) {
      return {
        content: [{
          type: 'text',
          text: `Attachment ${attachmentId} was not found on email ${emailId}.`
        }]
      };
    }

    if (!attachment.downloadSupported) {
      return {
        content: [{
          type: 'text',
          text: `Attachment ${attachment.name} cannot be downloaded: ${attachment.downloadReason}`
        }]
      };
    }

    if (attachment.size > config.MAX_ATTACHMENT_DOWNLOAD_BYTES) {
      return {
        content: [{
          type: 'text',
          text: 'Attachment exceeds the maximum supported size of 25 MB.'
        }]
      };
    }

    let baseUrl;
    if (isHostedMode()) {
      baseUrl = getConfiguredServerBaseUrl(config) || getUserContext()?.serverBaseUrl;
      if (!baseUrl) {
        return {
          content: [{
            type: 'text',
            text: 'Attachment downloads in hosted mode require PUBLIC_BASE_URL or HOSTED_REDIRECT_URI to be configured.'
          }]
        };
      }
    } else {
      baseUrl = await startEmbeddedServer();
    }

    const authContext = getUserContext();
    const ticket = issueDownloadTicket({
      accessToken,
      authContext,
      emailId,
      attachmentId,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
    });

    const safeName = sanitizeAttachmentFilename(attachment.name);
    const url = `${baseUrl}/attachments/download/${ticket.token}`;

    return {
      content: [{
        type: 'text',
        text: `Attachment download URL created for ${attachment.name}.

URL: ${url}
Content Type: ${attachment.contentType}
Size: ${attachment.size} bytes
Expiry: One-time use, valid for 5 minutes
Suggested Filename: ${safeName}

Example:
curl -L "${url}" -o "${safeName}"`
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
        text: `Error creating attachment download URL: ${error.message}`
      }]
    };
  }
}

module.exports = handleGetAttachmentDownloadUrl;
