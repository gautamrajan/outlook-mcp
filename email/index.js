/**
 * Email module for Outlook MCP server
 */
const ENABLE_SEND_EMAIL_TOOL = false;

const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleListAttachments = require('./list-attachments');
const handleGetAttachmentDownloadUrl = require('./get-attachment-download-url');
const handleCreateDraft = require('./create-draft');
const handleMarkAsRead = require('./mark-as-read');
const handleArchiveEmail = require('./archive-email');
const handleCreateReplyDraft = require('./create-reply-draft');

const sendEmailTool = {
  name: "send-email",
  description: "Composes and sends a new email",
  inputSchema: {
    type: "object",
    properties: {
      to: {
        type: "string",
        description: "Comma-separated list of recipient email addresses"
      },
      cc: {
        type: "string",
        description: "Comma-separated list of CC recipient email addresses"
      },
      bcc: {
        type: "string",
        description: "Comma-separated list of BCC recipient email addresses"
      },
      subject: {
        type: "string",
        description: "Email subject"
      },
      body: {
        type: "string",
        description: "Email body content (can be plain text or HTML)"
      },
      importance: {
        type: "string",
        description: "Email importance (normal, high, low)",
        enum: ["normal", "high", "low"]
      },
      saveToSentItems: {
        type: "boolean",
        description: "Whether to save the email to sent items"
      }
    },
    required: ["to", "subject", "body"]
  },
  handler: handleSendEmail
};

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails. Defaults to fuzzy token matching."
        },
        queryExactPhrase: {
          type: "boolean",
          description: "If true, treat query as an exact phrase. Quote characters in the input are still sanitized; phrase behavior is controlled by this flag."
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name. Defaults to fuzzy token matching within the sender field."
        },
        fromExactPhrase: {
          type: "boolean",
          description: "If true, treat from as an exact phrase within the sender field. Quote characters in the input are still sanitized; phrase behavior is controlled by this flag."
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name. Defaults to fuzzy token matching within the recipient field."
        },
        toExactPhrase: {
          type: "boolean",
          description: "If true, treat to as an exact phrase within the recipient field. Quote characters in the input are still sanitized; phrase behavior is controlled by this flag."
        },
        subject: {
          type: "string",
          description: "Filter by email subject. Defaults to fuzzy token matching within the subject field."
        },
        subjectExactPhrase: {
          type: "boolean",
          description: "If true, treat subject as an exact phrase within the subject field. Quote characters in the input are still sanitized; phrase behavior is controlled by this flag."
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "list-attachments",
    description: "Lists attachments for a specific email",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "ID of the email whose attachments should be listed"
        }
      },
      required: ["emailId"]
    },
    handler: handleListAttachments
  },
  {
    name: "get-attachment-download-url",
    description: "Creates a one-time download URL for an email attachment",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "ID of the email containing the attachment"
        },
        attachmentId: {
          type: "string",
          description: "ID of the attachment to download"
        }
      },
      required: ["emailId", "attachmentId"]
    },
    handler: handleGetAttachmentDownloadUrl
  },
  ...(ENABLE_SEND_EMAIL_TOOL ? [sendEmailTool] : []),
  {
    name: "create-draft",
    description: "Creates an email draft in the Drafts folder without sending it",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: ["subject", "body"]
    },
    handler: handleCreateDraft
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "archive-email",
    description: "Moves an email to the Archive folder",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to archive"
        }
      },
      required: ["id"]
    },
    handler: handleArchiveEmail
  },
  {
    name: "create-reply-draft",
    description: "Creates a reply draft to an existing email, preserving the conversation thread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "The ID of the email to reply to"
        },
        body: {
          type: "string",
          description: "Reply body content (can be plain text or HTML)"
        },
        replyAll: {
          type: "boolean",
          description: "Reply to all recipients (default: true). Set to false to reply only to the sender."
        },
        includeChain: {
          type: "boolean",
          description: "Include the quoted conversation chain below the reply (default: true). Set to false to replace the email body entirely."
        }
      },
      required: ["id", "body"]
    },
    handler: handleCreateReplyDraft
  }
];

module.exports = {
  emailTools,
  ENABLE_SEND_EMAIL_TOOL,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleListAttachments,
  handleGetAttachmentDownloadUrl,
  handleCreateDraft,
  handleMarkAsRead,
  handleArchiveEmail,
  handleCreateReplyDraft
};
