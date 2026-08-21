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
    `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
  );
}

/**
 * Authentication tool handler. Starts an in-process OAuth flow and returns
 * the Microsoft sign-in URL; the callback is handled by this server.
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    tokenManager.createTestTokens();
    return textResponse('Successfully authenticated with Microsoft Graph API (test mode)');
  }

  if (!force) {
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

  let flow;
  try {
    flow = await authFlow.start();
  } catch (error) {
    return textResponse(`Could not start the authentication flow: ${error.message}`);
  }

  return textResponse(
    `To authenticate with Microsoft, open this link in your browser:\n\n${flow.authUrl}\n\nThe link is valid for ${flow.expiresInMinutes} minutes. After signing in you will be redirected to ${flow.redirectUri}, which this server is temporarily listening on — if the server runs in a container, that port must be published to the host. Once done, use \`check-auth-status\` to confirm.`
  );
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
    description: "Authenticate with Microsoft Graph API to access Outlook data. Returns a Microsoft sign-in URL; the OAuth callback is handled by this server, no separate auth server is needed.",
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
