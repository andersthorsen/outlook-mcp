/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const tokenStorage = require('./storage');
const { AuthFlow } = require('./auth-flow');

const authFlow = new AuthFlow(tokenStorage);

function textResponse(text) {
  return { content: [{ type: "text", text }] };
}

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return textResponse(
    `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nAuth flow: ${config.AUTH_FLOW}`
  );
}

/**
 * Starts the OAuth 2.0 device code flow: returns a verification URL and user
 * code, then polls for completion in the background. No callback port needed.
 * @returns {Promise<object>} - MCP response
 */
async function startDeviceCodeFlow() {
  const deviceCode = await tokenStorage.initiateDeviceCodeFlow();

  // Start polling in the background — it will save tokens when complete
  const pollPromise = tokenStorage.pollForDeviceCodeToken(
    deviceCode.device_code,
    deviceCode.interval || 5,
    deviceCode.expires_in || 900
  );

  // Store the poll promise so check-auth-status can report on it
  tokenStorage._pendingDeviceCodePoll = pollPromise;
  pollPromise
    .then(() => { tokenStorage._pendingDeviceCodePoll = null; })
    .catch(() => { tokenStorage._pendingDeviceCodePoll = null; });

  return textResponse(
    `To authenticate, open your browser and go to:\n\n  ${deviceCode.verification_uri}\n\nEnter the code: ${deviceCode.user_code}\n\nThis code expires in ${Math.floor((deviceCode.expires_in || 900) / 60)} minutes.\n\nOnce you've completed the sign-in, use 'check-auth-status' to verify.`
  );
}

/**
 * Starts the OAuth 2.0 authorization code flow: spins up an in-process
 * callback server and returns the Microsoft sign-in URL.
 * @returns {Promise<object>} - MCP response
 */
async function startAuthCodeFlow() {
  const flow = await authFlow.start();
  return textResponse(
    `To authenticate with Microsoft, open this link in your browser:\n\n${flow.authUrl}\n\nThe link is valid for ${flow.expiresInMinutes} minutes. After signing in you will be redirected to ${flow.redirectUri}, which this server is temporarily listening on — if the server runs in a container, that port must be published to the host. Once done, use \`check-auth-status\` to confirm.`
  );
}

/**
 * Authentication tool handler. Runs the flow selected by AUTH_FLOW:
 * device code (no callback port needed) or authorization code with an
 * in-process callback server.
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  const useDeviceCode = config.AUTH_FLOW === 'device_code';

  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    tokenManager.createTestTokens();
    return textResponse('Successfully authenticated with Microsoft Graph API (test mode)');
  }

  if (!force) {
    if (tokenStorage._pendingDeviceCodePoll) {
      return textResponse('Authentication is already in progress. Complete the sign-in in your browser, or use check-auth-status to verify.');
    }
    const accessToken = await tokenStorage.getValidAccessToken();
    if (accessToken) {
      const expiresAt = new Date(tokenStorage.getExpiryTime()).toISOString();
      return textResponse(
        `Already authenticated. The access token is valid until ${expiresAt} and refreshes automatically.\n\nUse \`authenticate\` with \`force: true\` to sign in again (e.g. to switch account or grant new scopes).`
      );
    }
  } else {
    await tokenStorage.clearTokens();
  }

  try {
    return useDeviceCode ? await startDeviceCodeFlow() : await startAuthCodeFlow();
  } catch (error) {
    return textResponse(`Could not start the authentication flow: ${error.message}`);
  }
}

/**
 * Check authentication status tool handler. Reports the real token state,
 * attempting an automatic refresh when the access token has expired.
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  if (config.USE_TEST_MODE) {
    const testTokens = tokenManager.loadTokenCache();
    return textResponse(testTokens ? 'Authenticated and ready (test mode)' : 'Not authenticated (test mode)');
  }

  const tokens = await tokenStorage.getTokens();

  if (!tokens || !tokens.access_token) {
    if (tokenStorage._pendingDeviceCodePoll) {
      return textResponse('Waiting for you to complete the device code sign-in...');
    }
    return textResponse('Not authenticated. Use the `authenticate` tool to sign in.');
  }

  const scopeInfo = tokens.scope ? `\nGranted scopes: ${tokens.scope}` : '';

  if (!tokenStorage.isTokenExpired()) {
    const expiresAt = new Date(tokenStorage.getExpiryTime()).toISOString();
    return textResponse(`Authenticated and ready. Access token expires at ${expiresAt}.${scopeInfo}`);
  }

  // Expired (or about to expire): verify that automatic refresh actually works.
  const refreshedToken = await tokenStorage.getValidAccessToken();
  if (refreshedToken) {
    const expiresAt = new Date(tokenStorage.getExpiryTime()).toISOString();
    return textResponse(`Authenticated and ready. The access token was refreshed automatically and expires at ${expiresAt}.${scopeInfo}`);
  }

  if (tokenStorage._pendingDeviceCodePoll) {
    return textResponse('Waiting for you to complete the device code sign-in...');
  }

  return textResponse('Authentication has expired and could not be refreshed. Use the `authenticate` tool to sign in again.');
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data. Uses the device code flow (AUTH_FLOW=device_code) or returns a Microsoft sign-in URL with the OAuth callback handled by this server; no separate auth server is needed.",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API, including token expiry and granted scopes",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  authFlow,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
