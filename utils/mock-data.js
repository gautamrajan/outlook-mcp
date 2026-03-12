/**
 * Mock data functions for test mode
 */
const { Readable } = require('stream');

/**
 * Simulates Microsoft Graph API responses for testing
 * @param {string} method - HTTP method
 * @param {string} path - API path
 * @param {object} data - Request data
 * @param {object} queryParams - Query parameters
 * @returns {object} - Simulated API response
 */
function simulateGraphAPIResponse(method, path, data, queryParams) {
  console.error(`Simulating response for: ${method} ${path}`);
  
  if (method === 'GET') {
    if (path.includes('/attachments') && !path.endsWith('/$value')) {
      return {
        value: [
          {
            '@odata.type': '#microsoft.graph.fileAttachment',
            id: 'simulated-attachment-file',
            name: 'report.txt',
            contentType: 'text/plain',
            size: 128,
            isInline: false,
          },
          {
            '@odata.type': '#microsoft.graph.referenceAttachment',
            id: 'simulated-attachment-reference',
            name: 'shared-link.url',
            contentType: 'application/octet-stream',
            size: 64,
            isInline: false,
          }
        ]
      };
    }

    if (path.includes('messages') && !path.includes('sendMail')) {
      // Simulate a successful email list/search response
      if (path.includes('/messages/')) {
        // Single email response
        return {
          id: "simulated-email-id",
          subject: "Simulated Email Subject",
          from: {
            emailAddress: {
              name: "Simulated Sender",
              address: "sender@example.com"
            }
          },
          toRecipients: [{
            emailAddress: {
              name: "Recipient Name",
              address: "recipient@example.com"
            }
          }],
          ccRecipients: [],
          bccRecipients: [],
          receivedDateTime: new Date().toISOString(),
          bodyPreview: "This is a simulated email preview...",
          body: {
            contentType: "text",
            content: "This is the full content of the simulated email. Since we can't connect to the real Microsoft Graph API, we're returning this placeholder content instead."
          },
          hasAttachments: false,
          importance: "normal",
          isRead: false,
          internetMessageHeaders: []
        };
      } else {
        // Email list response
        return {
          value: [
            {
              id: "simulated-email-1",
              subject: "Important Meeting Tomorrow",
              from: {
                emailAddress: {
                  name: "John Doe",
                  address: "john@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date().toISOString(),
              bodyPreview: "Let's discuss the project status...",
              hasAttachments: false,
              importance: "high",
              isRead: false
            },
            {
              id: "simulated-email-2",
              subject: "Weekly Report",
              from: {
                emailAddress: {
                  name: "Jane Smith",
                  address: "jane@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 86400000).toISOString(), // Yesterday
              bodyPreview: "Please find attached the weekly report...",
              hasAttachments: true,
              importance: "normal",
              isRead: true
            },
            {
              id: "simulated-email-3",
              subject: "Question about the project",
              from: {
                emailAddress: {
                  name: "Bob Johnson",
                  address: "bob@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 172800000).toISOString(), // 2 days ago
              bodyPreview: "I had a question about the timeline...",
              hasAttachments: false,
              importance: "normal",
              isRead: false
            }
          ]
        };
      }
    } else if (path.includes('mailFolders')) {
      // Simulate a mail folders response
      return {
        value: [
          { id: "inbox", displayName: "Inbox" },
          { id: "drafts", displayName: "Drafts" },
          { id: "sentItems", displayName: "Sent Items" },
          { id: "deleteditems", displayName: "Deleted Items" }
        ]
      };
    }
  } else if (method === 'POST' && path.includes('sendMail')) {
    // Simulate a successful email send
    return {};
  }
  
  // If we get here, we don't have a simulation for this endpoint
  console.error(`No simulation available for: ${method} ${path}`);
  return {};
}

/**
 * Simulates streaming Microsoft Graph API responses for testing.
 * @param {string} method - HTTP method
 * @param {string} path - API path
 * @returns {{statusCode: number, headers: object, stream: import('stream').Readable}}
 */
function simulateGraphAPIStreamResponse(method, path) {
  console.error(`Simulating stream response for: ${method} ${path}`);

  if (method === 'GET' && path.includes('/attachments/') && path.endsWith('/$value')) {
    const payload = Buffer.from('Simulated attachment content\n', 'utf8');
    return {
      statusCode: 200,
      headers: {
        'content-type': 'text/plain',
        'content-length': String(payload.length),
      },
      stream: Readable.from(payload),
    };
  }

  return {
    statusCode: 404,
    headers: { 'content-type': 'text/plain' },
    stream: Readable.from(Buffer.from('Not found', 'utf8')),
  };
}

module.exports = {
  simulateGraphAPIResponse
  ,
  simulateGraphAPIStreamResponse
};
