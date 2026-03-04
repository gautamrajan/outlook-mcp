/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';
const tenantId = process.env.OUTLOOK_TENANT_ID || process.env.MS_TENANT_ID || 'common';

module.exports = {
  // Server information
  SERVER_NAME: "outlook-assistant",
  SERVER_VERSION: "1.0.0",

  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',

  // Transport
  MCP_TRANSPORT: process.env.MCP_TRANSPORT || 'stdio',  // 'stdio' or 'http'
  PORT: parseInt(process.env.PORT, 10) || 3000,

  // Hosted mode configuration
  HOSTED: {
    enabled: (process.env.MCP_TRANSPORT || 'stdio').toLowerCase() === 'http',
    tokenEncryptionKey: process.env.TOKEN_ENCRYPTION_KEY || '',
    tokenStorePath: process.env.TOKEN_STORE_PATH || path.join(homeDir, '.outlook-mcp-hosted-tokens.json'),
    sessionStorePath: process.env.SESSION_STORE_PATH || path.join(homeDir, '.outlook-mcp-sessions.json'),
    hostedRedirectUri: process.env.HOSTED_REDIRECT_URI || '',  // e.g. https://myserver.com/auth/callback
  },

  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || process.env.MS_CLIENT_SECRET || '',
    tenantId,
    tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    redirectUri: 'http://localhost:3333/auth/callback',
    scopes: ['offline_access', 'Mail.Read', 'Mail.ReadWrite', 'User.Read', 'Calendars.Read', 'Calendars.ReadWrite', 'Contacts.Read'],
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3333',
    hostedRedirectUri: process.env.HOSTED_REDIRECT_URI || '',
    hostedTokenStorePath: process.env.TOKEN_STORE_PATH || path.join(homeDir, '.outlook-mcp-hosted-tokens.json'),
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",
};
