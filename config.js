/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

// Single source of truth for OAuth scopes; must cover every Graph call the tools make.
const DEFAULT_SCOPES = [
  'offline_access',
  'User.Read',
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Files.Read',
  'Files.ReadWrite',
  'Contacts.Read'
];

const tenantId = process.env.MS_TENANT_ID || 'common';
const authorityHost = (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');
const authPort = parseInt(process.env.MS_AUTH_PORT, 10) || 3333;

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",

  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',

  // Authentication flow: 'device_code' or 'authorization_code' (default)
  AUTH_FLOW: process.env.AUTH_FLOW || 'authorization_code',

  // Authentication configuration. Client credentials accept both env-var
  // namespaces: MS_* (.env / auth server) and OUTLOOK_* (MCP client config).
  AUTH_CONFIG: {
    clientId: process.env.MS_CLIENT_ID || process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.MS_CLIENT_SECRET || process.env.OUTLOOK_CLIENT_SECRET || '',
    tenantId,
    authorityHost,
    authPort,
    redirectUri: process.env.MS_REDIRECT_URI || `http://localhost:${authPort}/auth/callback`,
    authEndpoint: process.env.MS_AUTH_ENDPOINT || `${authorityHost}/${tenantId}/oauth2/v2.0/authorize`,
    tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
    scopes: process.env.MS_SCOPES ? process.env.MS_SCOPES.split(' ') : DEFAULT_SCOPES,
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json')
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  // Windows timezone for creating events via Graph API
  DEFAULT_TIMEZONE: "Central European Standard Time",
  // IANA timezone for display (used in Prefer: outlook.timezone header)
  // Set OUTLOOK_TIMEZONE env var (e.g. "Europe/Oslo") to send the
  // Prefer: outlook.timezone header with calendar requests.
  // When empty, times are returned in UTC.
  DISPLAY_TIMEZONE: process.env.OUTLOOK_TIMEZONE || "",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // Power Automate / Flow constants
  FLOW_API_ENDPOINT: 'https://api.flow.microsoft.com',
  FLOW_SCOPE: 'https://service.flow.microsoft.com/.default',
};
